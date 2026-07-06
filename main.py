from flask import Flask, request, jsonify
import openpyxl
from openpyxl.styles import PatternFill
import urllib.request
import json, io, base64, os, traceback

app = Flask(__name__)

# Хранилище состояния в памяти
registry_state = {}

RED_FILL    = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")

PROHIBITED_STEMS = [
    'мазь','таблет','лекарств','бижутер','украшен','золот','серебр',
    'цеп','ожерель','колье','чокер','подвес','кулон','серьг','сережк',
    'кафф','клипс','пусет','браслет','кольц','оружи','колец',
]

def get_usd_tjs_rate():
    url = 'https://open.er-api.com/v6/latest/USD'
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=10) as response:
        data = json.loads(response.read())
    return data['rates']['TJS']


def find_col(headers, *names):
    for name in names:
        for i, h in enumerate(headers):
            if h and str(h).strip().upper() == name.strip().upper():
                return i
    for name in names:
        for i, h in enumerate(headers):
            if h and name.lower() in str(h).lower():
                return i
    return None


@app.route('/')
@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/state', methods=['POST'])
def save_state():
    global registry_state
    registry_state = request.get_json()
    return jsonify({'ok': True})


@app.route('/state', methods=['GET'])
def get_state():
    return jsonify(registry_state)


@app.route('/process', methods=['POST'])
def process():
    try:
        usd_rate_param = request.form.get('usdRate')
        rate = float(usd_rate_param) if usd_rate_param else get_usd_tjs_rate()
        limit = 200 * rate

        file = request.files['file']
        wb = openpyxl.load_workbook(file)
        ws = wb.active

        headers = [cell.value for cell in ws[1]]

        sum_col     = find_col(headers, 'Сумма')
        name_col    = find_col(headers, 'Имя получателя')
        surname_col = find_col(headers, 'Фамилия получателя')
        phone_col   = find_col(headers, 'Контактный номер', 'Телефон')
        product_col = find_col(headers, 'Наименование товара')

        print(f"Cols: sum={sum_col} name={name_col} surname={surname_col} phone={phone_col} product={product_col}", flush=True)

        # Максимальная колонка с данными — красить только до неё
        max_data_col = max(
            (i + 1 for i, h in enumerate(headers) if h is not None),
            default=1
        )

        # Сгруппировать суммы по человеку
        groups = {}
        rows_data = []

        for row in ws.iter_rows(min_row=2):
            name    = str(row[name_col].value    or '') if name_col    is not None else ''
            surname = str(row[surname_col].value  or '') if surname_col is not None else ''
            phone   = str(row[phone_col].value    or '') if phone_col   is not None else ''
            product = str(row[product_col].value  or '') if product_col is not None else ''
            full_name = f"{name} {surname}".strip()
            key = full_name + phone

            try:
                amount = float(row[sum_col].value or 0) if sum_col is not None else 0.0
            except (ValueError, TypeError):
                amount = 0.0

            groups[key] = groups.get(key, 0) + amount
            rows_data.append({'row': row, 'key': key, 'full_name': full_name, 'product': product})

        violators, prohibited_items = [], []
        seen_violators, seen_prohibited = set(), set()

        for item in rows_data:
            row     = item['row']
            key     = item['key']
            name    = item['full_name']
            product = item['product']
            lower   = product.lower()

            is_prohibited = any(s in lower for s in PROHIBITED_STEMS)
            is_violator = groups[key] > limit

            # Считаем нарушителей со ВСЕХ строк (включая уже покрашенные)
            if is_violator:
                if name not in seen_violators:
                    seen_violators.add(name)
                    violators.append({'name': name, 'amountUSD': round(groups[key] / rate, 2)})
            elif is_prohibited:
                if name + product not in seen_prohibited:
                    seen_prohibited.add(name + product)
                    prohibited_items.append({'name': name, 'product': product, 'amountUSD': round(groups[key] / rate, 2)})

            # Не перекрашивать уже покрашенные строки
            try:
                existing = row[0].fill
                if existing and existing.fill_type not in (None, 'none') and \
                   existing.fgColor.rgb not in ('00000000', 'FF000000'):
                    continue
            except Exception:
                pass

            # Красить только до последней колонки с данными
            if is_violator:
                for cell in row[:max_data_col]:
                    cell.fill = YELLOW_FILL
            elif is_prohibited:
                for cell in row[:max_data_col]:
                    cell.fill = RED_FILL

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
