from flask import Flask, request, jsonify
import base64, traceback, os
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill

app = Flask(__name__)

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

@app.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/process', methods=['POST'])
def process():
    try:
        if request.files.get('file'):
            usd_rate   = float(request.form.get('usdRate', 10))
            file_bytes = request.files['file'].read()
        else:
            data       = request.get_json()
            usd_rate   = float(data['usdRate'])
            file_bytes = base64.b64decode(data['fileBase64'])
        return jsonify(process_excel(usd_rate, file_bytes))
    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5679))
    app.run(host='0.0.0.0', port=port)
