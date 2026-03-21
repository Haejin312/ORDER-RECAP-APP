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

# ── 템플릿 파일 경로 ───────────────────────────────────────
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'Carhartt order recap 양식.xlsx')

def get_template_ws():
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.worksheets[0]  # 시트 이름 무관하게 첫 번째 시트 사용
    return wb, ws

def w(ws, row, col, val, fmt=None):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        # MergedCell이면 병합 범위의 top-left 셀에 직접 접근
        for merged_range in ws.merged_cells.ranges:
            if (merged_range.min_row <= row <= merged_range.max_row and
                merged_range.min_col <= col <= merged_range.max_col):
                cell = ws.cell(row=merged_range.min_row, column=merged_range.min_col)
                break
        else:
            return
    cell.value = val
    if fmt:
        cell.number_format = fmt

# ── 사이즈 행 번호 동적 탐색 (B열=col2) ───────────────────
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

# ── 헤더 행 탐색 (B열=col2 검색으로 수정) ─────────────────
def get_header_rows(ws):
    headers = {}
    # A열(col1)과 B열(col2) 모두 검색
    for col in [1, 2]:
        for row in range(1, 50):
            val = ws.cell(row=row, column=col).value
            if not val:
                continue
            v = str(val).strip().upper()
            if ('PO#' in v or v == 'PO#') and 'po' not in headers:
                headers['po'] = row
            elif 'ORDER UNIT' in v and 'unit' not in headers:
                headers['unit'] = row
            elif 'EX FACTORY' in v and 'ex' not in headers:
                headers['ex'] = row
            elif 'DELIVERY' in v and 'dlv' not in headers:
                headers['dlv'] = row
            elif v == 'STYLE#' and 'style' not in headers:
                headers['style'] = row
            elif 'COLOR CODE' in v and 'color_code' not in headers:
                headers['color_code'] = row
            elif v == 'COLOR' and 'color' not in headers:
                headers['color'] = row
            elif 'FABRIC' in v and len(v) < 10 and 'fabric' not in headers:
                headers['fabric'] = row
            elif 'SHIP TO' in v and 'ship' not in headers:
                headers['ship'] = row
            elif 'SHIP MODE' in v and 'mode' not in headers:
                headers['mode'] = row
    return headers

# ── 날짜 문자열 파싱 헬퍼 ─────────────────────────────────
def parse_date(date_str):
    """MM/DD/YYYY 또는 YYYY-MM-DD 형식 문자열을 date 객체로 변환"""
    if not date_str:
        return None
    try:
        # MM/DD/YYYY
        if '/' in str(date_str):
            return datetime.datetime.strptime(str(date_str), '%m/%d/%Y').date()
        # YYYY-MM-DD
        return datetime.date.fromisoformat(str(date_str))
    except Exception:
        return None

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
            model='claude-sonnet-4-5',
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
  "total_pcs": 1327,
  "total_amount": 6606.51
}

Size parsing rules:
- XSREG → XS, MREG → M, 2XLREG → 2XL (strip REG suffix)
- LTLL → LTLL, 2XLTLL → 2XLTLL (keep TLL suffix)
- Set pcs=0 fob=0 for sizes not in PO
- ship_to: use only first line (e.g. "Carhartt DC5" or "Carhartt Inc")
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

        msg = client.messages.create(
            model='claude-sonnet-4-5',
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

        # PyMuPDF로 스케치 이미지 추출
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        page = doc[0]
        images = page.get_images(full=True)
        sketch_b64 = None
        if images:
            xref = images[0][0]
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
        return jsonify({'error': f'Template not found: {TEMPLATE_PATH}'}), 500

    try:
        data = request.json
        file_no = data.get('fileNo', '')
        pos     = data.get('pos', [])
        bom_map = data.get('bom', {})
        vendor  = data.get('vendorNo', '907697')
        sewing  = data.get('sewing', 'Apparel Links')
        today   = data.get('today', datetime.date.today().isoformat())

        wb, ws = get_template_ws()
        size_row = get_size_rows(ws)
        hdr      = get_header_rows(ws)

        # ── 헤더 정보 ────────────────────────────────────────
        for row in range(1, 10):
            val = ws.cell(row=row, column=1).value
            if val and 'FILE' in str(val).upper():
                w(ws, row, 3, file_no)
                break

        for row in range(1, 15):
            val = ws.cell(row=row, column=1).value
            if val and 'VENDOR' in str(val).upper():
                w(ws, row, 3, vendor)
                break

        for row in range(1, 15):
            val = ws.cell(row=row, column=1).value
            if val and 'SEWING' in str(val).upper():
                w(ws, row, 3, sewing)
                break

        # ── Fabric Combination ──────────────────────────────
        all_fabrics = []
        seen_combos = set()
        for bom in bom_map.values():
            for fab in bom.get('fabrics', []):
                key = fab.get('combo', '')
                if key not in seen_combos:
                    all_fabrics.append(fab)
                    seen_combos.add(key)
        all_fabrics.sort(key=lambda x: x.get('combo', ''))

        # C1~C4 행 찾기 (A열 또는 B열)
        combo_rows = {}
        for row in range(1, ws.max_row + 1):
            for col in [1, 2]:
                val = ws.cell(row=row, column=col).value
                if val and re.match(r'^C\s*[1-4]$', str(val).strip()):
                    key = str(val).strip().replace(' ', '')
                    if key not in combo_rows:
                        combo_rows[key] = row

        for fab in all_fabrics:
            combo = fab.get('combo', '')
            row = combo_rows.get(combo)
            if not row:
                continue
            w(ws, row, 3, fab.get('body_code', ''))
            w(ws, row, 4, fab.get('body_desc', ''))
            w(ws, row, 7, fab.get('trim_code', ''))
            w(ws, row, 8, fab.get('trim_desc', ''))
            w(ws, row, 11, fab.get('hs_code', ''))

        # ── Style# 헤더 + 스케치 ────────────────────────────
        TARGET_W_PX = round(5 / 2.54 * 96)  # 189px
        style_cols = {}

        for po in pos:
            style = po.get('style', '')
            if style not in style_cols:
                idx = len(style_cols)
                header_col = 16 + idx * 3
                w(ws, 3, header_col, style)
                style_cols[style] = header_col

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
                        anchor_col = get_column_letter(header_col)
                        xl_img.anchor = f'{anchor_col}4'
                        xl_img.width  = TARGET_W_PX
                        xl_img.height = target_h
                        ws.add_image(xl_img)
                    except Exception:
                        pass  # 스케치 삽입 실패해도 계속 진행

        # ── PO 데이터 입력 ────────────────────────────────────
        slot_cols = [3 + i * 2 for i in range(16)]

        for slot_idx, po in enumerate(pos):
            if slot_idx >= 16:
                break
            cp = slot_cols[slot_idx]
            cf = cp + 1

            style     = po.get('style', '')
            color     = po.get('color_code', '')
            bom       = bom_map.get(style, {})
            colorways = bom.get('active_colorways', {})
            color_name = colorways.get(color, '')

            if hdr.get('po'):
                w(ws, hdr['po'], cp, po.get('po_number', ''))
            if hdr.get('unit'):
                w(ws, hdr['unit'], cp, po.get('total_pcs', 0))
            if hdr.get('ex'):
                d = parse_date(po.get('ex_factory_date', ''))
                if d:
                    w(ws, hdr['ex'], cp, d, 'MM/DD/YYYY')
            if hdr.get('dlv'):
                d = parse_date(po.get('delivery_date', ''))
                if d:
                    w(ws, hdr['dlv'], cp, d, 'MM/DD/YYYY')
            if hdr.get('style'):
                w(ws, hdr['style'], cp, style)
            if hdr.get('color_code'):
                w(ws, hdr['color_code'], cp, color)
            if hdr.get('color'):
                w(ws, hdr['color'], cp, color_name)
            if hdr.get('ship'):
                w(ws, hdr['ship'], cp, po.get('ship_to', 'Carhartt Inc'))
            if hdr.get('mode'):
                w(ws, hdr['mode'], cp, po.get('ship_mode', 'S-Ocean'))

            # Fabric combo 자동 배정 (HTS 기준)
            hts = po.get('hts', '')
            fabric_combo = 'C1'
            for fab in all_fabrics:
                if fab.get('hs_code') == hts:
                    fabric_combo = fab.get('combo', 'C1')
                    break
            if hdr.get('fabric'):
                w(ws, hdr['fabric'], cp, fabric_combo)

            # 사이즈 데이터
            sizes = po.get('sizes', {})
            for sz, row in size_row.items():
                sz_data = sizes.get(sz, {})
                pcs = sz_data.get('pcs', 0)
                fob = sz_data.get('fob', 0)
                w(ws, row, cp, pcs if pcs else None)
                w(ws, row, cf, fob if fob else None)

        # ── Revision History ──────────────────────────────────
        rev_row = None
        for row in range(1, ws.max_row + 1):
            for col in [1, 29]:
                val = ws.cell(row=row, column=col).value
                if val and 'REVISION' in str(val).upper():
                    rev_row = row + 2
                    break
            if rev_row:
                break
        if not rev_row:
            rev_row = 9

        today_date = parse_date(today) or datetime.date.today()
        w(ws, rev_row, 29, today_date, 'YYYY-MM-DD')
        w(ws, rev_row, 30, 'PO Issued')

        # ── Excel 저장 ────────────────────────────────────────
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


# ══════════════════════════════════════════════════════════
# 헬스체크
# ══════════════════════════════════════════════════════════
@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
