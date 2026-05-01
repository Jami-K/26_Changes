import glob as glob_module
import hashlib
import io
import os
import threading
import time

from flask import (Flask, abort, jsonify, redirect, render_template,
                   request, send_file, session, url_for)

from database import (add_recipient, delete_recipient,
                      get_all_changes, get_change_by_id, get_recipients,
                      get_setting, init_db, set_pdf_path, set_setting)
from excel_reader import sync_excel
from pptx_gen import generate_pptx

app = Flask(__name__)
app.secret_key = 'mat-change-internal-2026'

IMAGE_EXTS  = ('png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp')
SERVER_URL  = 'http://192.168.60.160:5500'
EXCEL_PATH  = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'index.xlsx')


def _build_email_body(change: dict) -> str:
    notes = (change.get('notes') or '').strip()
    if ',' in notes:
        items = [n.strip() for n in notes.split(',') if n.strip()]
        notes_fmt = '\n   '.join(f'- {item}' for item in items)
    else:
        notes_fmt = notes

    return (
        f"안녕하십니까, 품질관리팀입니다.\n"
        f"{change.get('product_name', '')}에 대한 변경점 발생하여 공유드립니다.\n\n"
        f"1. 제품명 : {change.get('product_name', '')}\n"
        f"2. 대상 : {change.get('changed_material', '')}\n"
        f"3. 혼용여부 : {change.get('mixable', '')}\n"
        f"4. 변경내용 : {change.get('change_content', '')}\n"
        f"5. 예상적용일 : {change.get('apply_date', '')}\n"
        f"6. 특이사항\n"
        f"   : {notes_fmt}\n\n"
        f"자세한 내용은 첨부파일 혹은 자재변경점관리 페이지({SERVER_URL})을 확인하시기 바랍니다.\n"
        f"감사합니다."
    )


# ── 헬퍼 ──────────────────────────────────────────────────────────────

def resolve_folder(path: str) -> str:
    """상대 경로를 앱 기준 절대 경로로 변환."""
    if not path:
        return ''
    if not os.path.isabs(path):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), path)
    return path


def get_images_in_folder(folder: str) -> list[str]:
    """폴더 내 이미지 파일 목록(절대 경로, 정렬)을 반환."""
    folder = resolve_folder(folder)
    if not folder or not os.path.isdir(folder):
        return []
    files = []
    for ext in IMAGE_EXTS:
        files.extend(glob_module.glob(os.path.join(folder, f'*.{ext}')))
        files.extend(glob_module.glob(os.path.join(folder, f'*.{ext.upper()}')))
    return sorted(set(files))


def check_pptx_auth() -> bool:
    pw_hash = get_setting('pptx_pw_hash')
    if not pw_hash:
        return True  # 비밀번호 미설정 → 모두 허용
    return session.get('pptx_auth') is True


def is_admin() -> bool:
    return not bool(get_setting('pptx_pw_hash')) or session.get('pptx_auth') is True


# ── 백그라운드 동기화 ─────────────────────────────────────────────────

def _do_sync():
    if os.path.isfile(EXCEL_PATH):
        try:
            sync_excel(EXCEL_PATH)
            set_setting('last_sync', time.strftime('%Y-%m-%d %H:%M:%S'))
        except Exception as e:
            print(f'[sync] {e}')


def _bg_sync():
    while True:
        time.sleep(300)
        _do_sync()


threading.Thread(target=_bg_sync, daemon=True).start()


# ── 라우트 ────────────────────────────────────────────────────────────

@app.route('/')
def index():
    q = request.args.get('q', '').strip()
    changes = get_all_changes(q)
    last_sync = get_setting('last_sync', '-')
    return render_template('list.html', changes=changes, q=q, last_sync=last_sync)


@app.route('/detail/<int:cid>')
def detail(cid):
    change = get_change_by_id(cid)
    if not change:
        abort(404)
    folder = (change.get('design_file') or '').strip()
    images = get_images_in_folder(folder)
    image_list = [{'idx': i, 'name': os.path.basename(p)} for i, p in enumerate(images)]
    return render_template('detail.html', change=change, image_list=image_list,
                           is_admin=is_admin())


@app.route('/image/<int:cid>/<int:idx>')
def serve_image(cid, idx):
    change = get_change_by_id(cid)
    if not change:
        abort(404)
    images = get_images_in_folder((change.get('design_file') or '').strip())
    if idx < 0 or idx >= len(images):
        abort(404)
    return send_file(images[idx])


@app.route('/mail/<int:cid>')
def mail(cid):
    if not is_admin():
        abort(403)
    change = get_change_by_id(cid)
    if not change:
        abort(404)
    body       = _build_email_body(change)
    recipients = get_recipients()
    subject    = f"[변경점 공유] {change.get('product_name', '')} - {change.get('changed_material', '')} 변경"
    return render_template('mail.html', change=change, body=body,
                           recipients=recipients, subject=subject)


@app.route('/pdf/<int:cid>')
def serve_pdf(cid):
    change = get_change_by_id(cid)
    if not change:
        abort(404)
    path = (change.get('pdf_path') or '').strip()
    if not path or not os.path.isfile(path):
        abort(404)
    return send_file(path, mimetype='application/pdf',
                     as_attachment=False,
                     download_name=os.path.basename(path))


@app.route('/set_pdf/<int:cid>', methods=['POST'])
def set_pdf(cid):
    if not is_admin():
        abort(403)
    if not get_change_by_id(cid):
        abort(404)
    path = request.form.get('pdf_path', '').strip()
    set_pdf_path(cid, path)
    return redirect(url_for('detail', cid=cid))


# ── PPTX 탭 (비밀번호 보호) ───────────────────────────────────────────

@app.route('/pptx_view/<int:cid>', methods=['GET', 'POST'])
def pptx_view(cid):
    change = get_change_by_id(cid)
    if not change:
        abort(404)

    if request.method == 'POST':
        pw = request.form.get('password', '')
        stored = get_setting('pptx_pw_hash')
        if stored and hashlib.sha256(pw.encode()).hexdigest() == stored:
            session['pptx_auth'] = True
            return redirect(url_for('pptx_view', cid=cid))
        return render_template('pptx_tab.html', change=change,
                               need_auth=True, error='비밀번호가 올바르지 않습니다.')

    pw_set = bool(get_setting('pptx_pw_hash'))
    if pw_set and not session.get('pptx_auth'):
        return render_template('pptx_tab.html', change=change,
                               need_auth=True, error=None)

    return render_template('pptx_tab.html', change=change, need_auth=False)


@app.route('/pptx_download/<int:cid>')
def pptx_download(cid):
    if not check_pptx_auth():
        abort(403)
    change = get_change_by_id(cid)
    if not change:
        abort(404)
    data = generate_pptx(change)
    fname = f"변경점_{change['product_code']}_{change['notice_date']}.pptx"
    return send_file(
        io.BytesIO(data),
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        as_attachment=True,
        download_name=fname,
    )


# ── 설정 ──────────────────────────────────────────────────────────────

@app.route('/settings', methods=['GET', 'POST'])
def settings():
    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'settings_auth':
            pw = request.form.get('password', '')
            stored = get_setting('pptx_pw_hash')
            if stored and hashlib.sha256(pw.encode()).hexdigest() == stored:
                session['pptx_auth'] = True
            else:
                session['pptx_auth'] = False
            return redirect(url_for('settings'))

        if action == 'sync':
            if not check_pptx_auth():
                return jsonify(success=False, error='인증이 필요합니다')
            if not os.path.isfile(EXCEL_PATH):
                return jsonify(success=False, error='index.xlsx 파일을 찾을 수 없습니다')
            try:
                n = sync_excel(EXCEL_PATH)
                set_setting('last_sync', time.strftime('%Y-%m-%d %H:%M:%S'))
                return jsonify(success=True, inserted=n)
            except Exception as e:
                return jsonify(success=False, error=str(e))

        if action == 'set_pptx_pw':
            pw = request.form.get('pptx_pw', '').strip()
            if pw:
                set_setting('pptx_pw_hash', hashlib.sha256(pw.encode()).hexdigest())
            else:
                set_setting('pptx_pw_hash', '')
            return redirect(url_for('settings'))

        if action == 'add_recipient':
            name  = request.form.get('name', '').strip()
            email = request.form.get('email', '').strip()
            if name and email:
                add_recipient(name, email)
            return redirect(url_for('settings'))

        if action == 'delete_recipient':
            delete_recipient(request.form.get('rid'))
            return redirect(url_for('settings'))

    recipients   = get_recipients()
    last_sync    = get_setting('last_sync', '-')
    pptx_pw_set  = bool(get_setting('pptx_pw_hash'))
    is_auth      = bool(session.get('pptx_auth')) or not pptx_pw_set
    return render_template('settings.html',
                           recipients=recipients,
                           last_sync=last_sync,
                           pptx_pw_set=pptx_pw_set,
                           is_auth=is_auth)


if __name__ == '__main__':
    init_db()
    _do_sync()
    app.run(host='0.0.0.0', port=5000, debug=False)
