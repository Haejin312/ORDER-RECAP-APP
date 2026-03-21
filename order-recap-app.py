import os, io, json, base64, datetime, re, time
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
import fitz

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Carhartt order recap 양식.xlsx')

ROW_FILE_NO   = 1
ROW_SEASON    = 2
ROW_VENDOR    = 3
ROW_SEWING    = 4
ROW_C = {'C1': 11, 'C2': 12, 'C3': 13, 'C4': 14}
ROW_PO       = 17
ROW_UNIT     = 18
ROW_EX       = 19
ROW_DLV      = 20
ROW_STYLE    = 21
ROW_COLOR_CD = 22
ROW_COLOR    = 23
ROW_FABRIC   = 24
ROW_SHIP     = 25
ROW_MODE     = 26
SIZE_ROWS = {
    'XS': 30, 'S': 31, 'M': 32, 'L': 33, 'XL': 34, '2XL': 35,
    '3XL': 36, '4XL': 37, '5XL': 38, '6XL': 39,
    'LTLL': 40, 'XLTLL': 41, '2XLTLL': 42, '3XLTLL': 43, '4XLTLL': 44, '5XLTLL': 45
}
ROW_REV_START = 9
COL_REV_DATE  = 29
COL_REV_NOTES = 30

def pcs_col(i): return 3 + i * 2
def fob_col(i): return 4 + i * 2
def style_header_col(i): return 16 + i * 3

def get_template_ws():
    wb = load_workbook(TEMPLATE_PATH)
    return wb, wb.worksheets[0]

def w(ws, row, col, val, fmt=None):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= row <= mr.max_row and mr.min_col <= col <= mr.max_col:
                cell = ws.cell(row=mr.min_row, column=mr.min_col)
                break
        else:
            return
    cell.value = val
    if fmt:
        cell.number_format = fmt

def parse_date(s):
    if not s: return None
    try:
        if '/' in str(s):
            return datetime.datetime.strptime(str(s), '%m/%d/%Y').date()
        return datetime.date.fromisoformat(str(s))
    except:
        return None

def extract_pdf_text(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    text = ''.join(page.get_text() for page in doc)
    doc.close()
    return text[:6000]


@app.route('/parse-po', methods=['POST'])
def parse_po():
    api_key = request.headers.get('X-Api-Key')
    if not api_key:
        return jsonify({'error': 'API key required'}), 401

    files = request.files.getlist('files')
    results = []
    client = anthropic.Anthropic(api_key=api_key)

    for f in files:
        pdf_bytes = f.read()
        pdf_text = extract_pdf_text(pdf_bytes)

        msg = client.messages.create(
            model='claude-sonnet-4-5',
            max_tokens=2000,
            messages=[{'role': 'user', 'content': [
                {'type': 'text', 'text': f'''Extract ALL data from this Carhartt PO text. Return ONLY valid JSON:

{{
  "po_number": "3510042802",
  "season": "F24",
  "ex_factory_date": "04/28/2024",
  "delivery_date": "05/23/2024",
  "style": "K126",
  "color_code": "HH5",
  "ship_to": "Carhartt Inc",
  "ship_mode": "S-Ocean",
  "hts": "6105.10.0010",
  "fabric_desc": "MENS SHIRT 60% COTTON 40% POLYESTER KNIT HENLEY",
  "sizes": {{
    "XS": {{"pcs": 4, "fob": 4.87}},
    "S": {{"pcs": 32, "fob": 4.87}},
    "M": {{"pcs": 277, "fob": 4.87}},
    "L": {{"pcs": 373, "fob": 4.87}},
    "XL": {{"pcs": 292, "fob": 4.87}},
    "2XL": {{"pcs": 170, "fob": 4.87}},
    "3XL": {{"pcs": 101, "fob": 5.75}},
    "4XL": {{"pcs": 19, "fob": 5.75}},
    "5XL": {{"pcs": 9, "fob": 5.75}},
    "6XL": {{"pcs": 0, "fob": 0}},
    "LTLL": {{"pcs": 14, "fob": 5.48}},
    "XLTLL": {{"pcs": 15, "fob": 5.48}},
    "2XLTLL": {{"pcs": 14, "fob": 5.48}},
    "3XLTLL": {{"pcs": 6, "fob": 5.48}},
    "4XLTLL": {{"pcs": 1, "fob": 5.48}},
    "5XLTLL": {{"pcs": 0, "fob": 0}}
  }},
  "total_pcs": 1327
}}

Rules:
- style: style number only (e.g. "K126", "K128", "K231")
- color_code: color code only after last hyphen in SKU (K128-BLK->"BLK", K126-HH5->"HH5")
- hts: HTS code from "HTS:XXXX" format (e.g. "6105.10.0010")
- fabric_desc: Extract EXACTLY what comes after the pipe "|" in the HTS line.
  Example: "HTS:6105.10.0010|MENS SHIRT 60% COTTON 40% POLYESTER KNIT HENLEY"
  -> fabric_desc = "MENS SHIRT 60% COTTON 40% POLYESTER KNIT HENLEY"
  This is critical - different fiber contents must produce different fabric_desc values.
- XSREG->XS, MREG->M, 2XLREG->2XL (strip REG suffix)
- LTLL->LTLL, 2XLTLL->2XLTLL (keep TLL suffix)
- pcs=0, fob=0 for sizes not in PO
- ship_to: first line only ("Carhartt DC5" or "Carhartt Inc")
- ship_mode: "S-Ocean" for ocean, "Air" for air
- Return ONLY the JSON

PO TEXT:
{pdf_text}'''}
            ]}]
        )

        text = msg.content[0].text.strip()
        clean = re.sub(r'```json|```', '', text).strip()
        po = json.loads(clean)
        time.sleep(3)
        results.append(po)

    # HTS + fabric_desc 조합으로 고유 Combo 슬롯 생성
    slot_key_map = {}
    fabric_slots = []
    combo_idx = 1

    for po in results:
        hts = po.get('hts', '')
        fabric_desc = po.get('fabric_desc', '')
        key = (hts, fabric_desc)
        if key not in slot_key_map:
            combo = f'C{combo_idx}'
            slot_key_map[key] = combo
            fabric_slots.append({
                'combo': combo,
                'hts': hts,
                'body_desc': fabric_desc,
                'body_code': '',
                'trim_code': '',
                'trim_desc': ''
            })
            combo_idx += 1
        po['fabric_combo'] = slot_key_map[key]

    return jsonify({'pos': results, 'fabric_slots': fabric_slots})


@app.route('/build-excel', methods=['POST'])
def build_excel():
    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({'error': 'Template not found'}), 500

    try:
        pos     = json.loads(request.form.get('pos', '[]'))
        extras  = json.loads(request.form.get('extras', '{}'))
        file_no = request.form.get('fileNo', '')
        vendor  = request.form.get('vendorNo', '907697')
        sewing  = request.form.get('sewing', 'Apparel Links')
        today   = request.form.get('today', datetime.date.today().isoformat())

        color_names = extras.get('color_names', {})
        fabrics     = extras.get('fabrics', [])
        sketches    = extras.get('sketches', {})

        for key in request.files:
            if key.startswith('sketch_'):
                style = key.replace('sketch_', '')
                img_bytes = request.files[key].read()
                sketches[style] = base64.b64encode(img_bytes).decode()

        wb, ws = get_template_ws()
        pos.sort(key=lambda p: parse_date(p.get('ex_factory_date', '')) or datetime.date.max)

        w(ws, ROW_FILE_NO, 3, file_no)
        w(ws, ROW_VENDOR,  3, vendor)
        w(ws, ROW_SEWING,  3, sewing)
        if pos:
            w(ws, ROW_SEASON, 3, pos[0].get('season', ''))

        fabrics.sort(key=lambda x: x.get('combo', ''))
        for fab in fabrics:
            row = ROW_C.get(fab.get('combo', ''))
            if not row: continue
            w(ws, row, 3,  fab.get('body_code', ''))
            w(ws, row, 4,  fab.get('body_desc', ''))
            w(ws, row, 7,  fab.get('trim_code', ''))
            w(ws, row, 8,  fab.get('trim_desc', ''))
            w(ws, row, 11, fab.get('hts', ''))

        TARGET_W_PX = round(5 / 2.54 * 96)
        style_idx_map = {}

        for po in pos:
            style = po.get('style', '')
            if style not in style_idx_map:
                idx = len(style_idx_map)
                style_idx_map[style] = idx
                hcol = style_header_col(idx)
                w(ws, 3, hcol, style)

                sketch_b64 = sketches.get(style)
                if sketch_b64:
                    try:
                        img_bytes = base64.b64decode(sketch_b64)
                        pil = PILImage.open(io.BytesIO(img_bytes))
                        ow, oh = pil.size
                        th = round(TARGET_W_PX * oh / ow)
                        pil = pil.resize((TARGET_W_PX, th), PILImage.LANCZOS)
                        buf = io.BytesIO()
                        pil.save(buf, format='PNG')
                        buf.seek(0)
                        xl_img = XLImage(buf)
                        xl_img.anchor = f'{get_column_letter(hcol)}4'
                        xl_img.width  = TARGET_W_PX
                        xl_img.height = th
                        ws.add_image(xl_img)
                    except:
                        pass

        for slot_idx, po in enumerate(pos):
            if slot_idx >= 16: break
            cp = pcs_col(slot_idx)
            cf = fob_col(slot_idx)

            style      = po.get('style', '')
            color      = po.get('color_code', '')
            color_name = color_names.get(color, '')

            w(ws, ROW_PO,       cp, po.get('po_number', ''))
            w(ws, ROW_UNIT,     cp, po.get('total_pcs', 0))
            w(ws, ROW_STYLE,    cp, style)
            w(ws, ROW_COLOR_CD, cp, color)
            w(ws, ROW_COLOR,    cp, color_name)
            w(ws, ROW_SHIP,     cp, po.get('ship_to', 'Carhartt Inc'))
            w(ws, ROW_MODE,     cp, po.get('ship_mode', 'S-Ocean'))

            ex = parse_date(po.get('ex_factory_date'))
            if ex: w(ws, ROW_EX, cp, ex, 'MM/DD/YYYY')

            dlv = parse_date(po.get('delivery_date'))
            if dlv: w(ws, ROW_DLV, cp, dlv, 'MM/DD/YYYY')

            w(ws, ROW_FABRIC, cp, po.get('fabric_combo', 'C1'))

            sizes = po.get('sizes', {})
            for sz, row in SIZE_ROWS.items():
                sd = sizes.get(sz, {})
                pcs = sd.get('pcs', 0)
                fob = sd.get('fob', 0)
                w(ws, row, cp, pcs if pcs else None)
                w(ws, row, cf, fob if fob else None)

        td = parse_date(today) or datetime.date.today()
        w(ws, ROW_REV_START, COL_REV_DATE,  td, 'YYYY-MM-DD')
        w(ws, ROW_REV_START, COL_REV_NOTES, 'PO Issued')

        # 파일명: {file_no}_Order Recap_Team#4_{today}.xlsx
        filename = f'{file_no}_Order Recap_Team#4_{today}.xlsx'

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        xlsx_b64 = base64.b64encode(buf.read()).decode()

        return jsonify({'xlsx_b64': xlsx_b64, 'filename': filename})

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
