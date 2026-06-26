from flask import Flask, request, jsonify
import openpyxl
from openpyxl.styles import PatternFill
import urllib.request
import json, io, base64, os, traceback

app = Flask(__name__)

RED_FILL    = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")


def get_usd_tjs_rate():
    url = 'https://open.er-api.com/v6/latest/USD'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=10) as response:
        data = json.loads(response.read())
    return data['rates']['TJS']


@app.route('/')
@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/process', methods=['POST'])
def process():
    try:
        # Курс: берём переданный или запрашиваем сами
        usd_rate_param = request.form.get('usdRate')
        rate = float(usd_rate_param) if usd_rate_param else get_usd_tjs_rate()
        limit = 200 * rate

        file = request.files['file']
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        # Найти нужные колонки по заголовку
        header_row = [cell.value for cell in ws[1]]
        color_col = sum_col = name_col = phone_col = None
        for i, h in enumerate(header_row):
            if h == '_rowColor':               color_col = i + 1
            if h == 'Сумма':                   sum_col   = i + 1
            if h == 'Имя получателя физ. лица': name_col = i + 1
            if h == 'Контактный номер':         phone_col = i + 1

        # Сгруппировать суммы по человеку (имя + телефон)
        groups = {}
        rows_data = []
        for row in ws.iter_rows(min_row=2):
            name  = row[name_col  - 1].value if name_col  else ''
            phone = row[phone_col - 1].value if phone_col else ''
            key   = str(name) + str(phone)
            amount = row[sum_col - 1].value if sum_col else 0
            color  = row[color_col - 1].value if color_col else 'none'
            groups[key] = groups.get(key, 0) + (amount or 0)
            rows_data.append({'row': row, 'key': key, 'color': color, 'name': str(name)})

        violators, prohibited_items = [], []
        seen_violators = set()

        for item in rows_data:
            row   = item['row']
            color = item['color']
            key   = item['key']
            name  = item['name']

            if color == 'red':
                for cell in row:
                    if color_col and cell.column != color_col:
                        cell.fill = RED_FILL
                prohibited_items.append({
                    'name': name,
                    'amountUSD': round(groups[key] / rate, 2)
                })
            elif groups[key] > limit:
                for cell in row:
                    if color_col and cell.column != color_col:
                        cell.fill = YELLOW_FILL
                if name not in seen_violators:
                    seen_violators.add(name)
                    violators.append({
                        'name': name,
                        'amountUSD': round(groups[key] / rate, 2)
                    })

        # Удалить служебную колонку _rowColor
        if color_col:
            ws.delete_cols(color_col)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return jsonify({
            'success': True,
            'violators': violators,
            'prohibited': prohibited_items,
            'totalProcessed': ws.max_row - 1,
            'usdRate': rate,
            'fileBase64': base64.b64encode(output.read()).decode(),
        })

    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'trace': traceback.format_exc()}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
