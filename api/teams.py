from http.server import BaseHTTPRequestHandler
import urllib.request
import urllib.parse
from urllib.error import HTTPError
import base64
import json
import os


CURSOR_BASE_URL = "https://api.cursor.com"


class handler(BaseHTTPRequestHandler):
    """Vercel Python Function: Cursor Admin API 프록시 (/api/teams/*)

    - 프론트는 `/api/teams/<subpath>`로 호출
    - `vercel.json`에서 `/api/teams/(.*)`를 `/api/teams?subpath=$1`로 rewrite
    - 서버에서 Basic 인증을 붙여 Cursor API로 프록시
    """

    def _set_cors_headers(self) -> None:
        # CORS 허용 헤더 설정
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')

    def do_OPTIONS(self) -> None:
        # CORS preflight 응답
        self.send_response(200)
        self._set_cors_headers()
        self.end_headers()

    def do_GET(self) -> None:
        self._proxy('GET')

    def do_POST(self) -> None:
        self._proxy('POST')

    def _proxy(self, method: str) -> None:
        try:
            # 요청 경로에서 subpath 추출 (rewrite로 전달됨)
            parsed = urllib.parse.urlparse(self.path)
            params = urllib.parse.parse_qs(parsed.query)
            subpath = (params.get('subpath', [''])[0]).lstrip('/')

            if not subpath:
                self._send_json(400, {
                    'error': '잘못된 요청',
                    'message': '유효한 teams 하위 경로가 필요합니다. 예: /api/teams/members'
                })
                return

            target_url = f"{CURSOR_BASE_URL}/teams/{subpath}"

            # 요청 본문 읽기
            content_length = int(self.headers.get('Content-Length', 0))
            body_bytes = None
            if content_length > 0:
                body_bytes = self.rfile.read(content_length)

            # 서버 환경변수에서 API 키 로드 (Vercel 환경변수 사용)
            api_key = os.environ.get('CURSOR_API_KEY', '').strip()
            if not api_key:
                self._send_json(500, {
                    'error': '서버 설정 오류',
                    'message': '환경변수 CURSOR_API_KEY가 설정되지 않았습니다.'
                })
                return

            credentials = f"{api_key}:"
            encoded_credentials = base64.b64encode(credentials.encode()).decode()

            headers = {
                'Content-Type': 'application/json',
                'Authorization': f'Basic {encoded_credentials}'
            }

            req = urllib.request.Request(
                target_url,
                data=body_bytes if method == 'POST' else None,
                headers=headers,
                method=method
            )

            with urllib.request.urlopen(req) as response:
                resp_data = response.read()
                status = response.status
                content_type = response.headers.get('Content-Type', 'application/json')

                self.send_response(status)
                self._set_cors_headers()
                self.send_header('Content-Type', content_type)
                self.end_headers()
                self.wfile.write(resp_data)

        except HTTPError as e:
            try:
                err_body = e.read().decode('utf-8')
                payload = json.loads(err_body)
            except Exception:
                payload = {'error': 'HTTP 오류', 'status': e.code, 'message': e.reason}
            self._send_json(500, payload)

        except Exception as e:
            self._send_json(500, {
                'error': '프록시 오류',
                'message': str(e)
            })

    def _send_json(self, status_code: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status_code)
        self._set_cors_headers()
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

