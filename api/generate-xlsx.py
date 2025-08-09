from http.server import BaseHTTPRequestHandler
import json
from io import BytesIO

try:
    from openpyxl import Workbook
except Exception:
    Workbook = None


class handler(BaseHTTPRequestHandler):
    """Vercel Python Function: XLSX 생성 (/api/generate-xlsx)
    프런트에서 전달한 시트 정의로 XLSX 파일을 만들어 바이너리로 응답합니다.
    """

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
            if Workbook is None:
                return self._send_json(500, {'error': 'openpyxl 미설치'})

            content_length = int(self.headers.get('Content-Length', 0))
            if content_length <= 0:
                return self._send_json(400, {'error': '요청 데이터가 없습니다.'})

            payload = json.loads(self.rfile.read(content_length).decode('utf-8'))
            sheets = payload.get('sheets', [])
            filename = payload.get('filename', 'attachments.xlsx')

            wb = Workbook()
            default_ws = wb.active
            wb.remove(default_ws)
            for sheet in sheets:
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
            data = bio.getvalue()

            self.send_response(200)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{filename}"')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        except Exception as e:
            self._send_json(500, {'error': str(e)})

    def _send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self._set_cors_headers()
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

