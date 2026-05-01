#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
採購智能分析平台 V6.1 — 一體化本機伺服器
  ① 靜態檔案服務（開啟儀表板 HTML）
  ② 智慧 CORS Proxy（轉發至任意 AI API，目標 URL 由 ?url= query param 決定）

使用方式（只需一個指令）：
  python procurement_cors_proxy.py            # 預設 Port 8080
  python procurement_cors_proxy.py 9000       # 指定 Port

啟動後：
  1. 瀏覽器開啟 http://localhost:8080
  2. 點選 Procurement_Dashboard_V6-1.html
  3. AI 面板 → ⚙️ AI 設定 → 進階設定 → CORS 代理：選「本機代理 (localhost:8080)」
  4. 端點 URL 照填原本的 https://api.xxx.com/v1，系統自動轉發
"""

import sys, ssl, json, os, mimetypes
import urllib.request, urllib.error
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, unquote, parse_qs

PORT      = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
SERVE_DIR = os.path.dirname(os.path.abspath(__file__))
PROXY_PATH = '/proxy'


class ProcurementServerHandler(BaseHTTPRequestHandler):

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers',
                         'Authorization, Content-Type, x-api-key, '
                         'anthropic-version, anthropic-dangerous-direct-browser-access, '
                         'HTTP-Referer, X-Title')

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.send_header('Content-Length', '0')
        self.end_headers()

    def do_GET(self):
        if self.path.startswith(PROXY_PATH):
            self._proxy('GET')
        else:
            self._serve_static()

    def do_POST(self):
        if self.path.startswith(PROXY_PATH):
            self._proxy('POST')
        else:
            self.send_error(405)

    # ── 靜態檔案 ─────────────────────────────────────────────────────
    def _serve_static(self):
        path = unquote(self.path.split('?')[0])
        local = os.path.normpath(SERVE_DIR + path)
        if not local.startswith(SERVE_DIR):
            self.send_error(403); return
        if os.path.isdir(local):
            self._dir_listing(local, path); return
        if not os.path.isfile(local):
            self.send_error(404); return
        mime, _ = mimetypes.guess_type(local)
        mime = mime or 'application/octet-stream'
        with open(local, 'rb') as f:
            data = f.read()
        self.send_response(200)
        self.send_header('Content-Type', mime + ('; charset=utf-8' if 'text' in mime else ''))
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _dir_listing(self, local_dir, url_path):
        files = sorted(os.listdir(local_dir))
        rows = []
        for f in files:
            full = os.path.join(local_dir, f)
            href = (url_path.rstrip('/') + '/' + f).replace('//', '/')
            icon = '📁 ' if os.path.isdir(full) else '📄 '
            size = f'{os.path.getsize(full):,} B' if os.path.isfile(full) else '—'
            rows.append(
                f'<tr><td><a href="{href}">{icon}{f}</a></td>'
                f'<td style="color:#888">{size}</td></tr>'
            )
        rows = ''.join(rows)
        html = f'''<!DOCTYPE html><html><head><meta charset="utf-8"><title>採購智能分析平台 V6.1</title>
<style>body{{font-family:sans-serif;padding:24px;background:#0d1117;color:#ddd}}
h2{{color:#6366f1}}a{{color:#818cf8;text-decoration:none}}a:hover{{color:#6366f1}}
table{{border-collapse:collapse;width:100%}}td{{padding:7px 14px;border-bottom:1px solid #1e293b}}
</style></head><body>
<h2>🏭 採購智能分析平台 V6.1 — 目錄</h2>
<table>{rows}</table>
<hr style="border-color:#1e293b;margin-top:24px">
<small style="color:#555">Procurement CORS Proxy · Port {PORT} · AI 面板開啟「CORS 代理」即可</small>
</body></html>'''
        data = html.encode()
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    # ── CORS Proxy ────────────────────────────────────────────────────
    def _proxy(self, method):
        parsed_req = urlparse(self.path)
        qs = parse_qs(parsed_req.query)
        target_url_list = qs.get('url', [])
        if not target_url_list:
            err = json.dumps({'error': '缺少 ?url= 參數，無法轉發'}).encode()
            self.send_response(400); self._cors()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(err)))
            self.end_headers(); self.wfile.write(err); return

        target_url = unquote(target_url_list[0])
        parsed = urlparse(target_url)

        length = int(self.headers.get('Content-Length', 0))
        body   = self.rfile.read(length) if length else None

        SKIP = {'host', 'content-length', 'connection',
                'transfer-encoding', 'te', 'trailer'}
        fwd_headers = {}
        for k, v in self.headers.items():
            if k.lower() not in SKIP:
                fwd_headers[k] = v
        fwd_headers['Host'] = parsed.netloc

        auth_val = fwd_headers.get('Authorization') or fwd_headers.get('authorization', '')
        key_val  = fwd_headers.get('x-api-key') or fwd_headers.get('X-Api-Key', '')
        print(f'  [PROXY] → {target_url}')
        if auth_val:
            print(f'  [AUTH ] ✅ Bearer …{auth_val[-6:]}')
        elif key_val:
            print(f'  [AUTH ] ✅ x-api-key …{key_val[-6:]}')
        else:
            print(f'  [AUTH ] ❌ 缺少 Authorization / x-api-key header')

        ctx = ssl.create_default_context()
        req = urllib.request.Request(target_url, data=body, method=method)
        for k, v in fwd_headers.items():
            req.add_unredirected_header(k, v)
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=120) as r:
                self.send_response(r.status)
                self._cors()
                for k, v in r.headers.items():
                    if k.lower() in ('access-control-allow-origin',
                                     'access-control-allow-headers',
                                     'access-control-allow-methods',
                                     'transfer-encoding', 'connection'):
                        continue
                    self.send_header(k, v)
                self.end_headers()
                while True:
                    chunk = r.read(4096)
                    if not chunk: break
                    try: self.wfile.write(chunk); self.wfile.flush()
                    except BrokenPipeError: break

        except urllib.error.HTTPError as e:
            body_err = e.read()
            self.send_response(e.code); self._cors()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(body_err)))
            self.end_headers(); self.wfile.write(body_err)

        except Exception as ex:
            msg = json.dumps({'error': str(ex)}).encode()
            self.send_response(502); self._cors()
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(msg)))
            self.end_headers(); self.wfile.write(msg)

    def log_message(self, fmt, *args):
        try:
            status = args[1]
            path   = args[0].split('"')[1] if '"' in args[0] else args[0]
            print(f'  [{status}] {path}')
        except Exception:
            pass


if __name__ == '__main__':
    print('=' * 60)
    print(f'  採購智能分析平台 V6.1 — 本機伺服器')
    print(f'  服務目錄：{SERVE_DIR}')
    print()
    print(f'  瀏覽器開啟：http://localhost:{PORT}')
    print(f'  點選 Procurement_Dashboard_V6-1.html')
    print()
    print(f'  AI 面板 → 進階設定 → CORS 代理：選「本機代理 (localhost:{PORT})」')
    print(f'  端點 URL 照填原本的 https://api.xxx.com/v1，系統自動轉發')
    print('=' * 60)
    server = HTTPServer(('', PORT), ProcurementServerHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n伺服器已停止。')
