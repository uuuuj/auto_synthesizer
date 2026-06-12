"""
send_mail.py — 본인 Gmail(SMTP)로 파일을 첨부해 실제 발송하는 헬퍼.

[1회 설정]
 1) 구글 계정에서 2단계 인증(2-Step Verification) 켜기
 2) https://myaccount.google.com/apppasswords 에서 '앱 비밀번호'(16자) 발급
 3) 이 폴더의 .env 파일에 아래 2줄 추가 (공백 있어도 자동 제거됨):
        GMAIL_USER=yourname@gmail.com
        GMAIL_APP_PW=abcd efgh ijkl mnop

[사용법]
    python send_mail.py <받는주소> <파일> [파일2 ...] \
           [--subject "제목"] [--body "본문"] [--cc 주소] [--no-zip]

[기본 동작]
 - .py 등 메일서버가 흔히 차단하는 확장자는 자동으로 zip 으로 묶어 첨부
   (--no-zip 으로 끌 수 있음)
 - 첨부 총용량 25MB 권장 상한

[예]
    python send_mail.py yjay1793.lee@samsung.com auto_synthesize_gui_dpv2.py
    python send_mail.py a@b.com app.py make_sample_data.py --subject "소스 전달"
"""

import os
import sys
import ssl
import smtplib
import zipfile
import argparse
import mimetypes
import tempfile
from email.message import EmailMessage

# 메일서버가 자주 차단하는 확장자 — 자동 zip 대상
RISKY_EXT = {'.py', '.pyw', '.js', '.exe', '.bat', '.cmd', '.sh',
             '.jar', '.ps1', '.vbs', '.scr', '.com'}


def load_env(path='.env'):
    """.env 파일을 읽어 환경변수로 로드 (python-dotenv 없이도 동작)."""
    if not os.path.exists(path):
        return
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            # 따옴표 제거
            v = v.strip().strip('"').strip("'")
            os.environ.setdefault(k.strip(), v)


def main():
    load_env()

    ap = argparse.ArgumentParser(description="Gmail로 파일 첨부 메일 발송")
    ap.add_argument('to', help="받는 사람 이메일")
    ap.add_argument('files', nargs='+', help="첨부할 파일 1개 이상")
    ap.add_argument('--subject', default=None, help="제목")
    ap.add_argument('--body', default=None, help="본문")
    ap.add_argument('--cc', default=None, help="참조(CC) 주소")
    ap.add_argument('--no-zip', action='store_true',
                    help="코드 파일 자동 zip 압축 비활성화")
    a = ap.parse_args()

    user = os.environ.get('GMAIL_USER')
    pw = os.environ.get('GMAIL_APP_PW')
    if not user or not pw:
        sys.exit("❌ .env 에 GMAIL_USER / GMAIL_APP_PW 를 설정하세요.\n"
                 "   (앱 비밀번호: https://myaccount.google.com/apppasswords)")
    pw = pw.replace(' ', '')   # 앱 비밀번호의 공백 제거

    # 존재 확인
    for f in a.files:
        if not os.path.exists(f):
            sys.exit(f"❌ 파일 없음: {f}")

    # 첨부 준비 — 위험 확장자는 하나의 zip 으로 묶음
    risky = [f for f in a.files
             if os.path.splitext(f)[1].lower() in RISKY_EXT]
    safe = [f for f in a.files if f not in risky]
    attach_paths = []
    tmpzip = None
    zipped = False
    if risky and not a.no_zip:
        tmpzip = os.path.join(tempfile.gettempdir(), 'mail_attach.zip')
        with zipfile.ZipFile(tmpzip, 'w', zipfile.ZIP_DEFLATED) as z:
            for f in risky:
                z.write(f, os.path.basename(f))
        attach_paths.append(tmpzip)
        attach_paths.extend(safe)
        zipped = True
    else:
        attach_paths = list(a.files)

    names = [os.path.basename(f) for f in a.files]
    subject = a.subject or f"[자동발송] {', '.join(names)}"
    body = a.body or (
        "자동 발송된 파일입니다.\n\n"
        f"포함 파일: {', '.join(names)}"
        + ("\n(코드 파일은 메일서버 차단 방지를 위해 zip으로 묶어 첨부했습니다.)"
           if zipped else ""))

    # 메시지 작성
    msg = EmailMessage()
    msg['From'] = user
    msg['To'] = a.to
    if a.cc:
        msg['Cc'] = a.cc
    msg['Subject'] = subject
    msg.set_content(body)

    total = 0
    for p in attach_paths:
        ctype, _ = mimetypes.guess_type(p)
        maintype, subtype = (ctype.split('/', 1) if ctype
                             else ('application', 'octet-stream'))
        with open(p, 'rb') as fp:
            data = fp.read()
        total += len(data)
        msg.add_attachment(data, maintype=maintype, subtype=subtype,
                           filename=os.path.basename(p))
    if total > 25 * 1024 * 1024:
        sys.exit(f"❌ 첨부 총용량 {total/1024/1024:.1f}MB — 25MB 초과. "
                 "Drive 링크 등을 사용하세요.")

    recipients = [a.to] + ([a.cc] if a.cc else [])
    print(f"발송 중... {user} → {a.to}"
          f"{(' (CC ' + a.cc + ')') if a.cc else ''}  "
          f"| 첨부 {len(attach_paths)}개, {total/1024:.0f}KB"
          f"{' [zip]' if zipped else ''}")

    ctx = ssl.create_default_context()
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=ctx,
                              timeout=30) as s:
            s.login(user, pw)
            s.send_message(msg, from_addr=user, to_addrs=recipients)
    except smtplib.SMTPAuthenticationError:
        sys.exit("❌ 인증 실패 — 앱 비밀번호(16자)를 다시 확인하세요. "
                 "(일반 비밀번호가 아니라 앱 비밀번호여야 함)")
    except (smtplib.SMTPException, OSError) as e:
        sys.exit(f"❌ 발송 실패: {e}\n"
                 "   (내부망에서 smtp.gmail.com:465 아웃바운드가 막혀 있을 수 있음)")
    finally:
        if tmpzip and os.path.exists(tmpzip):
            try:
                os.remove(tmpzip)
            except OSError:
                pass

    print(f"✅ 발송 완료: {subject}")


if __name__ == '__main__':
    main()
