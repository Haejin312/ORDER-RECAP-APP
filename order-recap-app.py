import os, io, json, base64, datetime, re
from flask import Flask, request, jsonify
from flask_cors import CORS
import anthropic
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
import fitz  # PyMuPDF

app = Flask(__name__)
CORS(app)

# ── 빈 템플릿 (base64) ─────────────────────────────────────
# 서버 시작 시 템플릿 파일 로드
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Carhartt order recap 양식.xlsx')

def get_template_ws():
    wb = load_workbook(TEMPLATE_PATH)
    return wb, wb['Sheet1']

def w(ws, row, col, val, fmt=None):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell): return
    cell.value = val
    if fmt: cell.number_format = fmt

# ── 사이즈 행 번호 동적 탐색 ───────────────────────────────
def get_size_rows(ws):
    size_row = {}
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=2).value
        if val and str(val).strip() in [
            'XS','S','M','L','XL','2XL','3XL','4XL','5XL','6XL',
            'LTLL','XLTLL','2XLTLL','3XLTLL','4XLTLL','5XLTLL'
        ]:
            size_row[str(val).strip()] = row
    return size_row

# ── 헤더 행 탐색 ───────────────────────────────────────────
def get_header_rows(ws):
    headers = {}
    for row in range(1, 30):
        val = ws.cell(row=row, column=1).value
        if not val: continue
        v = str(val).strip().upper()
        if 'PO#' in v or v == 'PO#':         headers['po'] = row
        elif 'ORDER UNIT' in v:               headers['unit'] = row
        elif 'EX FACTORY' in v:               headers['ex'] = row
        elif 'DELIVERY' in v:                 headers['dlv'] = row
        elif v == 'STYLE#':                   headers['style'] = row
        elif 'COLOR CODE' in v:               headers['color_code'] = row
        elif v.startswith('COLOR'):           headers['color'] = row
        elif 'FABRIC' in v and len(v) < 10:   headers['fabric'] = row
        elif 'SHIP TO' in v:                  headers['ship'] = row
        elif 'SHIP MODE' in v:                headers['mode'] = row
    return headers

# ══════════════════════════════════════════════════════════
# API: /parse-po
# ══════════════════════════════════════════════════════════
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
            model='claude-sonnet-4-6',
            max_tokens=4000,
            messages=[{
                'role': 'user',
                'content': [
                    {
                        'type': 'document',
                        'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': b64}
                    },
                    {
                        'type': 'text',
                        'text': '''Extract ALL data from this Carhartt PO PDF. Return ONLY valid JSON, no markdown:

{
  "po_number": "3510042802",
  "po_date": "01/29/2024",
  "season": "F24",
  "ex_factory_date": "04/28/2024",
  "delivery_date": "05/23/2024",
  "style": "K126",
  "color_code": "HH5",
  "ship_to": "Carhartt Inc",
  "ship_mode": "VESSEL",
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
  "total_pcs": 1327,
  "total_amount": 6606.51
}

Size parsing rules:
- XSREG → XS, MREG → M, 2XLREG → 2XL (strip REG suffix)
- LTLL → LTLL, 2XLTLL → 2XLTLL (keep TLL suffix)
- Set pcs=0 fob=0 for sizes not in PO
- Return ONLY the JSON'''
                    }
                ]
            }]
        )

        text = msg.content[0].text.strip()
        clean = re.sub(r'```json|```', '', text).strip()
        po = json.loads(clean)
        results.append(po)

    return jsonify({'pos': results})


# ══════════════════════════════════════════════════════════
# API: /parse-bom
# ══════════════════════════════════════════════════════════
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

        # 1. Claude로 텍스트 데이터 파싱
        msg = client.messages.create(
            model='claude-sonnet-4-6',
            max_tokens=2000,
            messages=[{
                'role': 'user',
                'content': [
                    {
                        'type': 'document',
                        'source': {'type': 'base64', 'media_type': 'application/pdf', 'data': b64_pdf}
                    },
                    {
                        'type': 'text',
                        'text': '''Extract from this Carhartt BOM PDF. Return ONLY valid JSON:

{
  "style": "K126",
  "active_colorways": {
    "HH5": "Heather Stone",
    "BLK": "Black",
    "NVY": "Navy",
    "CRH": "Carbon Heather"
  },
  "fabrics": [
    {
      "combo": "C1",
      "body_code": "JE002",
      "body_desc": "6.75oz 100% Cotton Jersey Solid",
      "trim_code": "RB002",
      "trim_desc": "8.9oz 1x1 Rib 100% Cotton Solid",
      "hs_code": "6110.20.2069"
    },
    {
      "combo": "C2",
      "body_code": "600106",
      "body_desc": "6.75oz 60/40 Cotton/Poly Heather Jersey",
      "trim_code": "600107",
      "trim_desc": "8.9oz 1x1 Rib 60/40 Heather",
      "hs_code": "6105.10.0010"
    }
  ]
}

Rules:
- active_colorways: find "Active Colorways" section, parse "CODE-Name" format
- fabrics: match combos C1-C4 to jersey/rib combinations from Composition section
- Return ONLY the JSON'''
                    }
                ]
            }]
        )

        text = msg.content[0].text.strip()
        clean = re.sub(r'```json|```', '', text).strip()
        bom = json.loads(clean)

        # 2. PyMuPDF로 스케치 이미지 추출 (page 0, image index 0 = front sketch)
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        page = doc[0]
        images = page.get_images(full=True)
        sketch_b64 = None
        if images:
            xref = images[0][0]  # front sketch
            base_img = doc.extract_image(xref)
            sketch_b64 = base64.b64encode(base_img['image']).decode()
        doc.close()

        bom['sketch_b64'] = sketch_b64
        bom_map[bom['style']] = bom

    return jsonify(bom_map)


# ══════════════════════════════════════════════════════════
# API: /build-excel
# ══════════════════════════════════════════════════════════
@app.route('/build-excel', methods=['POST'])
def build_excel():
    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({'error': 'Template not found'}), 500

    data = request.json
    file_no  = data.get('fileNo', '')
    pos      = data.get('pos', [])
    bom_map  = data.get('bom', {})
    vendor   = data.get('vendorNo', '907697')
    sewing   = data.get('sewing', 'Apparel Links')
    today    = data.get('today', datetime.date.today().isoformat())

    wb, ws = get_template_ws()
    size_row = get_size_rows(ws)
    hdr      = get_header_rows(ws)

    # ── 헤더 정보 ──────────────────────────────────────────
    # File#
    for row in range(1, 5):
        val = ws.cell(row=row, column=1).value
        if val and 'FILE' in str(val).upper():
            w(ws, row, 3, file_no)
            break

    # Vendor#
    for row in range(1, 15):
        val = ws.cell(row=row, column=1).value
        if val and 'VENDOR' in str(val).upper():
            w(ws, row, 3, int(vendor))
            break

    # Sewing
    for row in range(1, 15):
        val = ws.cell(row=row, column=1).value
        if val and 'SEWING' in str(val).upper():
            w(ws, row, 3, sewing)
            break

    # ── Fabric Combination ────────────────────────────────
    # BOM에서 모든 스타일의 fabric 수집 (중복 제거)
    all_fabrics = []
    seen_combos = set()
    for bom in bom_map.values():
        for fab in bom.get('fabrics', []):
            key = fab.get('combo', '')
            if key not in seen_combos:
                all_fabrics.append(fab)
                seen_combos.add(key)
    all_fabrics.sort(key=lambda x: x.get('combo', ''))

    # C1~C4 행 찾기
    combo_rows = {}
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val and re.match(r'^C\s*[1-4]$', str(val).strip()):
            combo_rows[str(val).strip().replace(' ', '')] = row

    for fab in all_fabrics:
        combo = fab.get('combo', '')
        row = combo_rows.get(combo)
        if not row: continue
        # Body: col C(3)=Code, D(4)=Desc  /  Trim: G(7)=Code, H(8)=Desc  /  HS: K(11)
        w(ws, row, 3, fab.get('body_code', ''))
        w(ws, row, 4, fab.get('body_desc', ''))
        w(ws, row, 7, fab.get('trim_code', ''))
        w(ws, row, 8, fab.get('trim_desc', ''))
        w(ws, row, 11, fab.get('hs_code', ''))

    # ── Style# 헤더 + 스케치 ─────────────────────────────
    # Style# 헤더 행 (row 3 기준 P/S/V/Y = col 16/19/22/25)
    # Style별 슬롯 컬럼 계산
    style_cols = {}  # style → first PCS col
    TARGET_W_PX = round(5 / 2.54 * 96)  # 189px

    for po in pos:
        style = po.get('style', '')
        if style not in style_cols:
            # style이 몇 번째 슬롯에 처음 등장하는지
            idx = len(style_cols)
            # Style# 헤더 위치: P3=col16, S3=col19, V3=col22, Y3=col25
            header_col = 16 + idx * 3
            w(ws, 3, header_col, style)
            style_cols[style] = header_col

            # 스케치 삽입
            bom = bom_map.get(style, {})
            sketch_b64 = bom.get('sketch_b64')
            if sketch_b64:
                img_bytes = base64.b64decode(sketch_b64)
                pil = PILImage.open(io.BytesIO(img_bytes))
                orig_w, orig_h = pil.size
                target_h = round(TARGET_W_PX * orig_h / orig_w)
                pil = pil.resize((TARGET_W_PX, target_h), PILImage.LANCZOS)
                buf = io.BytesIO()
                pil.save(buf, format='PNG')
                buf.seek(0)
                xl_img = XLImage(buf)
                # anchor: P4→col16, S4→col19, V4→col22
                anchor_col = get_column_letter(header_col)
                xl_img.anchor = f'{anchor_col}4'
                xl_img.width  = TARGET_W_PX
                xl_img.height = target_h
                ws.add_image(xl_img)

    # ── PO 데이터 입력 ────────────────────────────────────
    # PO 슬롯: slot1=col3(PCS)/col4(FOB), slot2=col5/6, ...
    slot_cols = [3 + i * 2 for i in range(16)]  # C,E,G,...

    for slot_idx, po in enumerate(pos):
        if slot_idx >= 16: break
        cp = slot_cols[slot_idx]  # PCS col
        cf = cp + 1               # FOB col

        style  = po.get('style', '')
        color  = po.get('color_code', '')
        bom    = bom_map.get(style, {})
        colorways = bom.get('active_colorways', {})
        color_name = colorways.get(color, '')

        # 헤더 데이터
        if hdr.get('po'):   w(ws, hdr['po'],   cp, int(po.get('po_number', 0)))
        if hdr.get('unit'): w(ws, hdr['unit'],  cp, po.get('total_pcs', 0))
        if hdr.get('ex'):
            d = po.get('ex_factory_date', '')
            w(ws, hdr['ex'], cp, d, 'MM/DD/YYYY')
        if hdr.get('dlv'):
            d = po.get('delivery_date', '')
            w(ws, hdr['dlv'], cp, d, 'MM/DD/YYYY')
        if hdr.get('style'):      w(ws, hdr['style'],      cp, style)
        if hdr.get('color_code'): w(ws, hdr['color_code'], cp, color)
        if hdr.get('color'):      w(ws, hdr['color'],      cp, color_name)
        if hdr.get('ship'):       w(ws, hdr['ship'],       cp, po.get('ship_to', 'Carhartt Inc'))
        if hdr.get('mode'):       w(ws, hdr['mode'],       cp, po.get('ship_mode', 'VESSEL'))

        # Fabric combo 자동 배정 (HTS 코드 기준)
        hts = po.get('hts', '')
        fabric_combo = 'C1'
        for fab in all_fabrics:
            if fab.get('hs_code') == hts:
                fabric_combo = fab.get('combo', 'C1')
                break
        if hdr.get('fabric'): w(ws, hdr['fabric'], cp, fabric_combo)

        # 사이즈 데이터
        sizes = po.get('sizes', {})
        for sz, row in size_row.items():
            sz_data = sizes.get(sz, {})
            pcs = sz_data.get('pcs', 0)
            fob = sz_data.get('fob', 0)
            w(ws, row, cp, pcs if pcs else None)
            w(ws, row, cf, fob if fob else None)

    # ── Revision History ──────────────────────────────────
    # AC=col29, AD=col30, 첫 데이터 row 찾기
    rev_row = None
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=29).value
        if val and 'REVISION' in str(val).upper():
            rev_row = row + 2  # 헤더 2행 아래
            break
    if not rev_row:
        rev_row = 9  # 기본값

    w(ws, rev_row, 29, datetime.date.fromisoformat(today), 'YYYY-MM-DD')
    w(ws, rev_row, 30, 'PO Issued')

    # ── Excel 저장 → base64 반환 ─────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    xlsx_b64 = base64.b64encode(buf.read()).decode()

    return jsonify({
        'xlsx_b64': xlsx_b64,
        'filename': f'OrderRecap_{file_no}_{today}.xlsx'
    })


# ══════════════════════════════════════════════════════════
# 헬스체크
# ══════════════════════════════════════════════════════════
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
