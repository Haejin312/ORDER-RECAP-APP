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

def pcs_col(slot_idx): return 3 + slot_idx * 2
def fob_col(slot_idx): return 4 + slot_idx * 2
def style_header_col(style_idx): return 16 + style_idx * 3

ROW_REV_START = 9
COL_REV_DATE  = 29
COL_REV_NOTES = 30

def get_template_ws():
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.worksheets[0]
    return wb, ws

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
    if not s:
        return None
    try:
        if '/' in str(s):
            return datetime.datetime.strptime(str(s), '%m/%d/%Y').date()
        return datetime.date.fromisoformat(str(s))
    except Exception:
        return None


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
        b64 = base64.b64encode(pdf_bytes).decode()

        msg = client.messages.create(
            model='claude-sonnet-4-5',
            max_tokens=2000,
            messages=[{'role': 'user', 'content': [
                {'type': 'document', 'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': b64}},
                {'type': 'text', 'text': '''Extract ALL data from this Carhartt PO PDF. Return ONLY valid JSON, no markdown:

{
  "po_number": "3510042802",
  "season": "F24",
  "ex_factory_date": "04/28/2024",
  "delivery_date": "05/23/2024",
  "style": "K126",
  "color_code": "HH5",
  "ship_to": "Carhartt Inc",
  "ship_mode": "S-Ocean",
  "hts": "6110.20.2069",
  "sizes": {
    "XS":     {"pcs": 4,   "fob": 4.87},
    "S":      {"pcs": 32,  "fob": 4.87},
    "M":      {"pcs": 277, "fob": 4.87},
    "L":      {"pcs": 373, "fob": 4.87},
    "XL":     {"pcs": 292, "fob": 4.87},
    "2XL":    {"pcs": 170, "fob": 4.87},
    "3XL":    {"pcs": 101, "fob": 5.75},
    "4XL":    {"pcs": 19,  "fob": 5.75},
    "5XL":    {"pcs": 9,   "fob": 5.75},
    "6XL":    {"pcs": 0,   "fob": 0},
    "LTLL":   {"pcs": 14,  "fob": 5.48},
    "XLTLL":  {"pcs": 15,  "fob": 5.48},
    "2XLTLL": {"pcs": 14,  "fob": 5.48},
    "3XLTLL": {"pcs": 6,   "fob": 5.48},
    "4XLTLL": {"pcs": 1,   "fob": 5.48},
    "5XLTLL": {"pcs": 0,   "fob": 0}
  },
  "total_pcs": 1327
}

Rules:
- style: extract only style number (e.g. "K126", "K128", "K231")
- color_code: extract only color code after hyphen (K231-PRT -> "PRT", K128-BLK -> "BLK", K126-HH5 -> "HH5")
- XSREG->XS, MREG->M, 2XLREG->2XL (strip REG suffix)
- LTLL->LTLL, 2XLTLL->2XLTLL (keep TLL suffix)
- pcs=0, fob=0 for sizes not in PO
- ship_to: first line only ("Carhartt DC5" or "Carhartt Inc")
- ship_mode: use exactly "S-Ocean" for ocean, "Air" for air
- Return ONLY the JSON'''}
            ]}]
        )

        text = msg.content[0].text.strip()
        clean = re.sub(r'```json|```', '', text).strip()
        po = json.loads(clean)
        time.sleep(10)
        results.append(po)

    return jsonify({'pos': results})


@app.route('/parse-bom', methods=['POST'])
def parse_bom():
    api_key = request.headers.get('X-Api-Key')
    if not api_key:
        return jsonify({'error': 'API key required'}), 401

    files = request.files.getlist('files')
    bom_map = {}
    client = anthropic.Anthropic(api_key=api_key)

    for f in files:
        pdf_bytes = f.read()
        b64_pdf = base64.b64encode(pdf_bytes).decode()

        msg = client.messages.create(
            model='claude-sonnet-4-5',
            max_tokens=2000,
            messages=[{'role': 'user', 'content': [
                {'type': 'document', 'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': b64_pdf}},
                {'type': 'text', 'text': '''Extract from this Carhartt BOM PDF. Return ONLY valid JSON:

{
  "style": "K128",
  "active_colorways": {
    "BLK": "Black",
    "NVY": "Navy",
    "CRH": "Carbon Heather",
    "HGY": "Heather Grey"
  },
  "fabrics": [
    {
      "combo": "C1",
      "body_code": "JE002",
      "body_desc": "6.75oz 100% Cotton Jersey Solid Only",
      "trim_code": "RB002",
      "trim_desc": "8.9oz 1x1 Rib Knit Solid Only",
      "hs_code": "6105.10.0010"
    }
  ]
}

Rules:
- style: style number only (e.g. "K128")
- active_colorways: find "Active Colorways" section, parse CODE-Name format. CODE is 2-4 uppercase letters only.
- fabrics: match C1-C4 combos from Composition section
- Return ONLY the JSON'''}
            ]}]
        )

        text = msg.content[0].text.strip()
        clean = re.sub(r'```json|```', '', text).strip()
        bom = json.loads(clean)

        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        page = doc[0]
        images = page.get_images(full=True)
        sketch_b64 = None
        if images:
            try:
                xref = images[0][0]
                base_img = doc.extract_image(xref)
                sketch_b64 = base64.b64encode(base_img['image']).decode()
            except Exception:
                pass
        doc.close()

        bom['sketch_b64'] = sketch_b64
        time.sleep(10)
        bom_map[bom['style']] = bom

    return jsonify(bom_map)


@app.route('/build-excel', methods=['POST'])
def build_excel():
    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({'error': f'Template not found: {TEMPLATE_PATH}'}), 500

    try:
        data    = request.json
        file_no = data.get('fileNo', '')
        pos     = data.get('pos', [])
        bom_map = data.get('bom', {})
        vendor  = data.get('vendorNo', '907697')
        sewing  = data.get('sewing', 'Apparel Links')
        today   = data.get('today', datetime.date.today().isoformat())

        wb, ws = get_template_ws()

        pos.sort(key=lambda p: parse_date(p.get('ex_factory_date', '')) or datetime.date.max)

        w(ws, ROW_FILE_NO, 3, file_no)
        w(ws, ROW_VENDOR,  3, vendor)
        w(ws, ROW_SEWING,  3, sewing)
        if pos:
            w(ws, ROW_SEASON, 3, pos[0].get('season', ''))

        all_fabrics = []
        seen = set()
        for bom in bom_map.values():
            for fab in bom.get('fabrics', []):
                key = fab.get('combo', '')
                if key not in seen:
                    all_fabrics.append(fab)
                    seen.add(key)
        all_fabrics.sort(key=lambda x: x.get('combo', ''))

        for fab in all_fabrics:
            row = ROW_C.get(fab.get('combo', ''))
            if not row:
                continue
            w(ws, row, 3,  fab.get('body_code', ''))
            w(ws, row, 4,  fab.get('body_desc', ''))
            w(ws, row, 7,  fab.get('trim_code', ''))
            w(ws, row, 8,  fab.get('trim_desc', ''))
            w(ws, row, 11, fab.get('hs_code', ''))

        TARGET_W_PX = round(5 / 2.54 * 96)
        style_idx_map = {}

        for po in pos:
            style = po.get('style', '')
            if style not in style_idx_map:
                idx = len(style_idx_map)
                style_idx_map[style] = idx
                hcol = style_header_col(idx)
                w(ws, 3, hcol, style)

                bom = bom_map.get(style, {})
                sketch_b64 = bom.get('sketch_b64')
                if sketch_b64:
                    try:
                        img_bytes = base64.b64decode(sketch_b64)
                        pil = PILImage.open(io.BytesIO(img_bytes))
                        orig_w, orig_h = pil.size
                        target_h = round(TARGET_W_PX * orig_h / orig_w)
                        pil = pil.resize((TARGET_W_PX, target_h), PILImage.LANCZOS)
                        buf = io.BytesIO()
                        pil.save(buf, format='PNG')
                        buf.seek(0)
                        xl_img = XLImage(buf)
                        xl_img.anchor = f'{get_column_letter(hcol)}4'
                        xl_img.width  = TARGET_W_PX
                        xl_img.height = target_h
                        ws.add_image(xl_img)
                    except Exception:
                        pass

        for slot_idx, po in enumerate(pos):
            if slot_idx >= 16:
                break

            cp = pcs_col(slot_idx)
            cf = fob_col(slot_idx)

            style      = po.get('style', '')
            color      = po.get('color_code', '')
            bom        = bom_map.get(style, {})
            color_name = bom.get('active_colorways', {}).get(color, '')

            w(ws, ROW_PO,       cp, po.get('po_number', ''))
            w(ws, ROW_UNIT,     cp, po.get('total_pcs', 0))
            w(ws, ROW_STYLE,    cp, style)
            w(ws, ROW_COLOR_CD, cp, color)
            w(ws, ROW_COLOR,    cp, color_name)
            w(ws, ROW_SHIP,     cp, po.get('ship_to', 'Carhartt Inc'))
            w(ws, ROW_MODE,     cp, po.get('ship_mode', 'S-Ocean'))

            ex = parse_date(po.get('ex_factory_date'))
            if ex:
                w(ws, ROW_EX, cp, ex, 'MM/DD/YYYY')

            dlv = parse_date(po.get('delivery_date'))
            if dlv:
                w(ws, ROW_DLV, cp, dlv, 'MM/DD/YYYY')

            hts = po.get('hts', '')
            fabric_combo = 'C1'
            for fab in all_fabrics:
                if fab.get('hs_code') == hts:
                    fabric_combo = fab.get('combo', 'C1')
                    break
            w(ws, ROW_FABRIC, cp, fabric_combo)

            sizes = po.get('sizes', {})
            for sz, row in SIZE_ROWS.items():
                sz_data = sizes.get(sz, {})
                pcs = sz_data.get('pcs', 0)
                fob = sz_data.get('fob', 0)
                w(ws, row, cp, pcs if pcs else None)
                w(ws, row, cf, fob if fob else None)

        today_date = parse_date(today) or datetime.date.today()
        w(ws, ROW_REV_START, COL_REV_DATE,  today_date, 'YYYY-MM-DD')
        w(ws, ROW_REV_START, COL_REV_NOTES, 'PO Issued')

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        xlsx_b64 = base64.b64encode(buf.read()).decode()

        return jsonify({
            'xlsx_b64': xlsx_b64,
            'filename': f'OrderRecap_{file_no}_{today}.xlsx'
        })

    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'traceback': traceback.format_exc()}), 500


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
