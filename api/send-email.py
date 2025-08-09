from http.server import BaseHTTPRequestHandler
import json
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import BytesIO

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None


class handler(BaseHTTPRequestHandler):
    """Vercel Python Function: SMTP 이메일 발송 (/api/send-email)"""

    def _set_cors_headers(self) -> None:
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')

    def do_OPTIONS(self) -> None:
        self.send_response(200)
        self._set_cors_headers()
        self.end_headers()

    def do_POST(self) -> None:
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            if content_length <= 0:
                return self._send_json(400, {'success': False, 'error': '요청 본문이 없습니다.'})

            payload = json.loads(self.rfile.read(content_length).decode('utf-8'))

            # 환경변수에서 기본 발신자/앱 비밀번호 읽기
            sender_email = os.environ.get('SMTP_SENDER_EMAIL', '').strip()
            sender_app_password = os.environ.get('SMTP_APP_PASSWORD', '').strip()
            if not sender_email or not sender_app_password:
                return self._send_json(500, {
                    'success': False,
                    'error': '서버 설정 오류',
                    'message': 'SMTP_SENDER_EMAIL / SMTP_APP_PASSWORD 환경변수가 필요합니다.'
                })

            to_emails = payload.get('to_emails', [])
            if not to_emails:
                return self._send_json(400, {'success': False, 'error': '수신자 이메일이 비어 있습니다.'})

            subject = payload.get('subject', '[Dashboard] 리포트')
            message = payload.get('message', '')
            from_email = (payload.get('from_email') or '').strip() or sender_email
            from_name = (payload.get('from_name') or '').strip() or 'Dashboard'

            # 메일 본문 구성
            html_content = f"""
            <html>
            <head><meta charset="UTF-8"></head>
            <body>
                <h2>📊 리포트</h2>
                <p>{message}</p>
                <hr style="margin:20px 0;border:none;border-top:1px solid #eee" />
                <div style="color:#666;font-size:12px;line-height:1.6">
                  이 이메일은 Samsung AI Dashboard에서 자동으로 발송되었습니다.<br/>
                  © 2025 Samsung AI Experience Group
                </div>
            </body>
            </html>
            """
            text_content = (
                "리포트\n\n"
                f"{message}\n\n"
                "---\n"
                "이 이메일은 Samsung AI Dashboard에서 자동으로 발송되었습니다.\n"
                "© 2025 Samsung AI Experience Group"
            )

            # 공통 첨부물 (XLSX) 생성: attachmentsSheets 형식 재사용
            attachments = []
            sheets_payload = payload.get('attachmentsSheets')
            if sheets_payload and Workbook is not None:
                wb = Workbook()
                default_ws = wb.active
                wb.remove(default_ws)
                for sheet in sheets_payload.get('sheets', []):
                    name = (sheet.get('name', 'Sheet') or 'Sheet')[:31]
                    ws = wb.create_sheet(title=name or 'Sheet')
                    headers = sheet.get('headers', [])
                    rows = sheet.get('rows', [])
                    if headers:
                        ws.append(headers)
                    for row in rows:
                        ws.append(row)
                bio = BytesIO()
                wb.save(bio)
                xlsx_bytes = bio.getvalue()
                part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                part.set_payload(xlsx_bytes)
                encoders.encode_base64(part)
                filename = sheets_payload.get('filename', 'attachments.xlsx')
                part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                attachments.append(part)

            # Gmail SMTP 발송
            smtp_server = 'smtp.gmail.com'
            smtp_port = 587
            sent = 0
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_app_password)

                for recipient in to_emails:
                    msg = MIMEMultipart('mixed')
                    msg['From'] = f"{from_name} <{from_email}>"
                    msg['To'] = recipient
                    msg['Subject'] = subject

                    alt = MIMEMultipart('alternative')
                    alt.attach(MIMEText(text_content, 'plain', 'utf-8'))
                    alt.attach(MIMEText(html_content, 'html', 'utf-8'))
                    msg.attach(alt)

                    for att in attachments:
                        msg.attach(att)

                    server.send_message(msg)
                    sent += 1

            return self._send_json(200, {
                'success': True,
                'message': f'이메일 발송 완료: {sent}명'
            })

        except smtplib.SMTPAuthenticationError:
            return self._send_json(500, {'success': False, 'error': 'Gmail 인증 실패. 앱 비밀번호를 확인해주세요.'})
        except Exception as e:
            return self._send_json(500, {'success': False, 'error': str(e)})

    def _send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self._set_cors_headers()
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

