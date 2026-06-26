#!/usr/bin/env python3
from http.server import HTTPServer, BaseHTTPRequestHandler
import json, base64, traceback, os
from io import BytesIO
from datetime import datetime
import openpyxl
from openpyxl.styles import PatternFill

PROHIBITED_STEMS = [
    'мазь','таблет','лекарств','бижутер','украшен','золот','серебр',
    'цеп','ожерель','колье','чокер','подвес','кулон','серьг','сережк',
    'кафф','клипс','пусет','браслет','кольц','оружи','колец',
]
RED    = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
YELLOW = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

def process_excel(usd_rate, file_bytes):
    wb = openpyxl.load_workbook(BytesIO(file_bytes))
    ws = wb.active
    violators, prohibited_items = [], []
    for row in ws.iter_rows(min_row=2):
        try:
            product   = str(row[9].value  or '')
            amount    = float(row[15].value or 0)
            full_name = f"{row[18].value or ''} {row[19].value or ''}".strip()
            amount_usd = round(amount / usd_rate, 2) if usd_rate > 0 else 0
            lower = product.lower()
            is_prohibited = any(s in lower for s in PROHIBITED_STEMS)
            fill = None
            if is_prohibited:
                fill = RED
                prohibited_items.append({'name': full_name, 'product': product, 'amountUSD': amount_usd})
            elif amount_usd > 200:
                fill = YELLOW
                violators.append({'name': full_name, 'product': product, 'amountUSD': amount_usd})
            if fill:
                for c in range(26): row[c].fill = fill
        except Exception:
            continue
    buf = BytesIO()
    wb.save(buf)
    return {
        'success': True,
        'violators': violators,
        'prohibited': prohibited_items,
        'totalProcessed': ws.max_row - 1,
        'fileBase64': base64.b64encode(buf.getvalue()).decode(),
    }

class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {format % args}", flush=True)
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'{"status":"ok"}')
    def do_POST(self):
        try:
            length = int(self.headers.get('Content-Length', 0))
            data   = json.loads(self.rfile.read(length))
            result = process_excel(float(data['usdRate']), base64.b64decode(data['fileBase64']))
            self._respond(200, result)
        except Exception as e:
            self._respond(500, {'success': False, 'error': str(e), 'trace': traceback.format_exc()})
    def _respond(self, code, payload):
        body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5679))
    server = HTTPServer(('0.0.0.0', port), Handler)
    print(f"✅ Excel processor running on port {port}", flush=True)
    server.serve_forever()
