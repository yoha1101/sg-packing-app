import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import OrderedDict, defaultdict
import io
from datetime import datetime

st.set_page_config(page_title="SG Export Document Generator", layout="centered")
st.title("📦 SPECIALGUEST Export Document Generator")

CBM_PER_BOX = 0.088

# ── 품목 분류 및 우선순위 ──────────────────────────────────
PRIORITY = [
    'Insulated Snowboard Jacket', 'Snowboard Jacket', 'Snowboard Pants',
    'Long-sleeve Tee', 'Sweatshirt, Hooded Sweatshirt', 'Sweatpants', 'T-shirt',
    'Cap', 'Beanie', 'Beanie(Angora)', 'HOOD WARMER N', 'BALACLAVA', 'Snowboard Stomppad'
]

SHEET_ORDER = ['Snowboard Pants','Snowboard Jacket','Insulated Snowboard Jacket',
    'Sweatpants','Sweatshirt, Hooded Sweatshirt','Long-sleeve Tee','T-shirt',
    'Beanie','Beanie(Angora)','Cap','HOOD WARMER N','BALACLAVA','Snowboard Stomppad']

MAIN_COLORS = {
    'Snowboard Pants':               'E2F0D9',
    'Snowboard Jacket':              'FFF2CC',
    'Insulated Snowboard Jacket':    'FCE4D6',
    'Sweatpants':                    'EAD1DC',
    'Sweatshirt, Hooded Sweatshirt': 'D9E1F2',
    'Long-sleeve Tee':               'DDEBF7',
    'T-shirt':                       'F2F2F2',
    'Beanie':                        'E2EFDA',
    'Beanie(Angora)':                'FFF2CC',
    'Cap':                           'DAEEF3',
    'HOOD WARMER N':                 'F4CCCC',
    'BALACLAVA':                     'D9EAD3',
    'Snowboard Stomppad':            'D9D9D9',
}

SEC_COLORS = {
    'Snowboard Pants':               'C6EFCE',
    'Snowboard Jacket':              'FFEB9C',
    'Insulated Snowboard Jacket':    'F8CBAD',
    'Sweatpants':                    'D5A6BD',
    'Sweatshirt, Hooded Sweatshirt': 'B4C7E7',
    'Long-sleeve Tee':               'BDD7EE',
    'T-shirt':                       'D9D9D9',
    'Beanie':                        'C6D9B8',
    'Beanie(Angora)':                'FFE699',
    'Cap':                           'B8D9E8',
    'HOOD WARMER N':                 'EA9999',
    'BALACLAVA':                     'A9C9A4',
    'Snowboard Stomppad':            'BFBFBF',
}

def get_prio(d): 
    return PRIORITY.index(d) if d in PRIORITY else 99

def standardize_category(style_raw):
    style = str(style_raw).strip().lower()
    if 'snowboard pants' in style or 'cargo pants' in style:
        return 'Snowboard Pants'
    elif 'snowboard jacket' in style and 'insulated' not in style:
        return 'Snowboard Jacket'
    elif 'insulated' in style and 'jacket' in style:
        return 'Insulated Snowboard Jacket'
    elif 'sweatpants' in style:
        return 'Sweatpants'
    elif 'sweatshirt' in style or 'hooded' in style or 'hoodie' in style:
        return 'Sweatshirt, Hooded Sweatshirt'
    elif 'long-sleeve' in style or 'long sleeve' in style:
        return 'Long-sleeve Tee'
    elif 't-shirt' in style or 'tee' in style:
        return 'T-shirt'
    elif 'cap' in style:
        return 'Cap'
    elif 'angora' in style:
        return 'Beanie(Angora)'
    elif 'beanie' in style:
        return 'Beanie'
    elif 'hood warmer' in style:
        return 'HOOD WARMER N'
    elif 'balaclava' in style:
        return 'BALACLAVA'
    elif 'stomppad' in style or 'stomp pad' in style:
        return 'Snowboard Stomppad'
    return style_raw

# ── Excel 공통 스타일 ─────────────────────────────────────
thin = Side(style='thin')
def tb(): return Border(left=thin, right=thin, top=thin, bottom=thin)
center  = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_a  = Alignment(horizontal='left',   vertical='center', wrap_text=True)
hdr_fill = PatternFill('solid', fgColor='006FC0')
col_fill = PatternFill('solid', fgColor='D9D9D9')
alt_fill = PatternFill('solid', fgColor='F7F6D6')
no_fill  = PatternFill()

def sc(ws, row, col, val, bold=False, size=9, color='000000', fill=None, align=None, border=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name='Arial', bold=bold, size=size, color=color)
    if fill:   c.fill = fill
    if align:  c.alignment = align
    if border: c.border = border
    return c

# ── CSV 파싱 ───────────────────────────────────────────────
def parse_export_csv(csv_file):
    df = pd.read_csv(csv_file)
    df['박스번호'] = df['박스번호'].ffill()
    df['무게(kg)'] = df['무게(kg)'].ffill()
    df['Category'] = df['Style'].apply(standardize_category)
    boxes = {}
    for box_no, group in df.groupby('박스번호'):
        boxes[box_no] = {
            'box_no': box_no,
            'weight': group['무게(kg)'].iloc[0],
            'items': []
        }
        for _, row in group.iterrows():
            boxes[box_no]['items'].append({
                'sku': row['SKU'],
                'style': row['Style'],
                'category': row['Category'],
                'item_name': row['품목명'],
                'hs_code': row['HS Code'],
                'color': row['Color'],
                'size': row['Size'],
                'qty': row['수량'],
                'price': row['단가(KRW)'],
                'material': row['Material']
            })
    return df, boxes

# ── 1. 패킹리스트 ─────────────────────────────────────────
def make_packing_list(boxes, destination):
    wb = Workbook()
    ws = wb.active
    ws.title = destination
    ws.column_dimensions['A'].width = 9
    ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 20
    for col in ['D','E','F','G','H','I']:
        ws.column_dimensions[col].width = 6
    ws.column_dimensions['J'].width = 8
    ws.column_dimensions['K'].width = 9

    headers = ['CTN NO.','STYLE','COLOR','S','M','L','XL','2XL','3XL','TOTAL','G.W(kgs)']
    for i, h in enumerate(headers, 1):
        sc(ws, 1, i, h, bold=True, fill=col_fill, align=center, border=tb())

    cur_row = 2
    for box_no in sorted(boxes.keys()):
        box = boxes[box_no]
        style_data = defaultdict(lambda: {
            'color': '', 'sizes': {'S':0,'M':0,'L':0,'XL':0,'2XL':0,'3XL':0}
        })
        for item in box['items']:
            key = item['item_name']
            style_data[key]['color'] = item['color']
            size = item['size']
            if size in style_data[key]['sizes']:
                style_data[key]['sizes'][size] += item['qty']

        first_in_box = True
        box_first_row = cur_row
        for style, data in style_data.items():
            if first_in_box:
                sc(ws, cur_row, 1, box_no, bold=True, align=center, border=tb())
                first_in_box = False
            else:
                sc(ws, cur_row, 1, '', border=tb())
            sc(ws, cur_row, 2, style, align=left_a, border=tb())
            sc(ws, cur_row, 3, data['color'], align=left_a, border=tb())
            total = 0
            for i, size in enumerate(['S','M','L','XL','2XL','3XL'], 4):
                qty = data['sizes'][size]
                sc(ws, cur_row, i, qty if qty > 0 else '', align=center, border=tb())
                total += qty
            sc(ws, cur_row, 10, total, align=center, border=tb())
            if cur_row == box_first_row:
                sc(ws, cur_row, 11, box['weight'], align=center, border=tb())
            else:
                sc(ws, cur_row, 11, '', border=tb())
            cur_row += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── 2. 인보이스 ───────────────────────────────────────────
def make_invoice(df, messrs, destination, date_str):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Invoice'
    for col, width in [('A',5),('B',10),('C',45),('D',12),('E',50),('F',12),('G',12)]:
        ws.column_dimensions[col].width = width
    ws.merge_cells('A1:G1')
    sc(ws, 1, 1, 'COMMERCIAL INVOICE', bold=True, size=14, align=center, fill=hdr_fill, color='FFFFFF')
    ws.row_dimensions[1].height = 25
    sc(ws, 3, 1, 'SHIPPER:', bold=True)
    ws.merge_cells('A4:D4'); sc(ws, 4, 1, 'SPECIALGUEST®', bold=True)
    ws.merge_cells('A5:D5'); sc(ws, 5, 1, '22, Misagangbyeonhangang-ro 346beon-gil, Hanam-si, Gyeonggi-do, 12923, Republic of Korea')
    ws.merge_cells('A6:D6'); sc(ws, 6, 1, 'Tel: +82 7077643333  Email: specialguest.co.kr@gmail.com')
    sc(ws, 8, 1, 'CONSIGNEE:', bold=True)
    ws.merge_cells('A9:D9'); sc(ws, 9, 1, messrs, bold=True)
    sc(ws, 11, 1, 'DATE:', bold=True); sc(ws, 11, 2, date_str)
    sc(ws, 12, 1, 'DESTINATION:', bold=True); sc(ws, 12, 2, destination)
    headers = ['NO.','HS CODE','Description of goods','QTY','Fabric ratio','Unit Price(KRW)','Amount(KRW)']
    for i, h in enumerate(headers, 1):
        sc(ws, 27, i, h, bold=True, fill=col_fill, align=center, border=tb())
    grouped = df.groupby(['품목명','HS Code','Material','단가(KRW)']).agg({'수량':'sum'}).reset_index()
    cur_row = 28
    total_qty = 0
    total_amount = 0
    for idx, (_, row) in enumerate(grouped.iterrows(), 1):
        sc(ws, cur_row, 1, idx, align=center, border=tb())
        sc(ws, cur_row, 2, row['HS Code'], align=center, border=tb())
        sc(ws, cur_row, 3, row['품목명'], align=left_a, border=tb())
        sc(ws, cur_row, 4, row['수량'], align=center, border=tb())
        sc(ws, cur_row, 5, row['Material'], align=left_a, border=tb())
        sc(ws, cur_row, 6, row['단가(KRW)'], align=center, border=tb())
        amount = row['수량'] * row['단가(KRW)']
        sc(ws, cur_row, 7, amount, align=center, border=tb())
        total_qty += row['수량']
        total_amount += amount
        cur_row += 1
    cur_row += 1
    ws.merge_cells(f'A{cur_row}:C{cur_row}')
    sc(ws, cur_row, 1, 'TOTAL', bold=True, align=center, fill=col_fill, border=tb())
    sc(ws, cur_row, 4, total_qty, bold=True, align=center, fill=col_fill, border=tb())
    sc(ws, cur_row, 7, total_amount, bold=True, align=center, fill=col_fill, border=tb())
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── 3. Actual Packing List ────────────────────────────────
def make_actual_packing_list(boxes, messrs, destination, date_str):
    wb = Workbook()
    ws = wb.active
    ws.title = 'ACTUAL PACKING LIST'
    widths = {'A':16,'B':34,'C':30,'D':9,'E':52,'F':14,'G':11,'H':9}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.merge_cells('A1:H1')
    sc(ws, 1, 1, 'ACTUAL PACKING LIST', bold=True, size=12, color='FFFFFF', fill=hdr_fill, align=center)
    ws.row_dimensions[1].height = 20
    sc(ws, 2, 1, 'SPECIALGUEST®', bold=True, size=10)
    ws.merge_cells('A3:H3')
    sc(ws, 3, 1, 'postal code: 12923 / 22, Misagangbyeonhangang-ro 346beon-gil, Hanam-si, Gyeonggi-do, Republic of Korea', align=left_a)
    ws.merge_cells('A4:H4')
    sc(ws, 4, 1, 'Tel: +82 7077643333  Email: specialguest.co.kr@gmail.com', align=left_a)
    ws.merge_cells('A6:F6'); sc(ws, 6, 1, f'MESSRS: {messrs}', bold=True)
    ws.merge_cells('G6:H6'); sc(ws, 6, 7, f'DATE: {date_str}', size=8)
    ws.merge_cells('A7:F7'); sc(ws, 7, 1, 'SHIPMENT FROM: Republic of Korea', bold=True)
    ws.merge_cells('A8:H8'); sc(ws, 8, 1, f'FINAL DESTINATION: {destination}', bold=True)
    sc(ws, 9, 8, CBM_PER_BOX, bold=True, align=center)
    headers = ['HS code','Description of goods','Fabric ratio','Quantity','** Comments','Dimension (cm)','Weight(kg)','CBM']
    for i, h in enumerate(headers, 1):
        sc(ws, 10, i, h, bold=True, fill=col_fill, align=center, border=tb())
    ws.row_dimensions[10].height = 14
    category_data = defaultdict(lambda: {'hs_code':'','material':'','qty':0,'boxes':[],'weight':0})
    for box_no, box in boxes.items():
        for item in box['items']:
            cat = item['category']
            category_data[cat]['hs_code'] = item['hs_code']
            category_data[cat]['material'] = item['material']
            category_data[cat]['qty'] += item['qty']
            if box_no not in category_data[cat]['boxes']:
                category_data[cat]['boxes'].append(box_no)
                category_data[cat]['weight'] += box['weight']
    sorted_cats = sorted(category_data.keys(), key=get_prio)
    cur_row = 11
    use_alt = False
    total_qty = 0
    total_boxes = 0
    total_weight = 0
    for cat in sorted_cats:
        data = category_data[cat]
        fill = alt_fill if use_alt else no_fill
        use_alt = not use_alt
        num_boxes = len(data['boxes'])
        sc(ws, cur_row, 1, data['hs_code'], bold=True, fill=fill, align=center, border=tb())
        sc(ws, cur_row, 2, cat, bold=True, fill=fill, align=left_a, border=tb())
        sc(ws, cur_row, 3, data['material'], bold=True, fill=fill, align=left_a, border=tb())
        sc(ws, cur_row, 4, data['qty'], bold=True, fill=fill, align=center, border=tb())
        sc(ws, cur_row, 5, num_boxes, bold=True, fill=fill, align=center, border=tb())
        sc(ws, cur_row, 6, '55 x 40 x 40', bold=True, fill=fill, align=center, border=tb())
        sc(ws, cur_row, 7, data['weight'], bold=True, fill=fill, align=center, border=tb())
        cbm_cell = ws.cell(row=cur_row, column=8, value=f'=E{cur_row}*$H$9')
        cbm_cell.font = Font(name='Arial', bold=True, size=9)
        cbm_cell.fill = fill
        cbm_cell.alignment = center
        cbm_cell.border = tb()
        total_qty += data['qty']
        total_boxes += num_boxes
        total_weight += data['weight']
        cur_row += 1
    cur_row += 1
    sc(ws, cur_row, 1, 'Please check the ** comments.', bold=True)
    cur_row += 1
    sc(ws, cur_row, 1, 'N = Nylon / P = Polyester / PU = Polyurethane / C = Cotton / R = Rayon / A = Acrylic', bold=True)
    cur_row += 2
    for i, h in enumerate(['Quantity','CTNS','Dimension (cm)','Weight(kg)','CBM'], start=4):
        sc(ws, cur_row, i, h, bold=True, align=center)
    cur_row += 1
    ws.merge_cells(f'A{cur_row}:C{cur_row}')
    sc(ws, cur_row, 1, 'Total', bold=True)
    sc(ws, cur_row, 4, total_qty, bold=True, align=center)
    sc(ws, cur_row, 5, total_boxes, bold=True, align=center)
    sc(ws, cur_row, 6, '55 x 40 x 40', bold=True, align=center)
    sc(ws, cur_row, 7, total_weight, bold=True, align=center)
    sc(ws, cur_row, 8, round(total_boxes * CBM_PER_BOX, 3), bold=True, align=center)
    cur_row += 2
    ws.merge_cells(f'D{cur_row}:H{cur_row}')
    sc(ws, cur_row, 4, 'MADE IN KOREA', bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── 4. 카테고리별 패킹리스트 ──────────────────────────────
def make_category_packing_list(boxes):
    wb = Workbook()
    wb.remove(wb.active)
    box_main_category = {}
    for box_no, box in boxes.items():
        cat_qty = defaultdict(int)
        for item in box['items']:
            cat_qty[item['category']] += item['qty']
        main_cat = max(cat_qty.items(), key=lambda x: (x[1], -get_prio(x[0])))[0]
        box_main_category[box_no] = main_cat
    category_boxes = defaultdict(list)
    for box_no, main_cat in box_main_category.items():
        category_boxes[main_cat].append(box_no)
    for cat in SHEET_ORDER:
        if cat not in category_boxes:
            continue
        ws = wb.create_sheet(cat[:31])
        main_fill = PatternFill('solid', fgColor=MAIN_COLORS.get(cat, 'FFFFFF'))
        ws.column_dimensions['A'].width = 9
        ws.column_dimensions['B'].width = 38
        ws.column_dimensions['C'].width = 20
        for c in ['D','E','F','G','H','I']:
            ws.column_dimensions[c].width = 6
        ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 9
        headers = ['CTN NO.','STYLE','COLOR','S','M','L','XL','2XL','3XL','TOTAL','G.W(kgs)']
        for i, h in enumerate(headers, 1):
            sc(ws, 1, i, h, bold=True, fill=col_fill, align=center, border=tb())
        cur_row = 2
        total_qty = 0
        for box_no in sorted(category_boxes[cat]):
            box = boxes[box_no]
            style_data = defaultdict(lambda: {
                'color': '', 'category': '', 'sizes': {'S':0,'M':0,'L':0,'XL':0,'2XL':0,'3XL':0}
            })
            for item in box['items']:
                key = item['item_name']
                style_data[key]['color'] = item['color']
                style_data[key]['category'] = item['category']
                size = item['size']
                if size in style_data[key]['sizes']:
                    style_data[key]['sizes'][size] += item['qty']
            first_in_box = True
            for style, data in style_data.items():
                if data['category'] == cat:
                    row_fill = main_fill
                    total_qty += sum(data['sizes'].values())
                else:
                    row_fill = PatternFill('solid', fgColor=SEC_COLORS.get(data['category'], 'EEEEEE'))
                if first_in_box:
                    sc(ws, cur_row, 1, box_no, bold=True, fill=row_fill, align=center, border=tb())
                    first_in_box = False
                else:
                    sc(ws, cur_row, 1, '', fill=row_fill, border=tb())
                sc(ws, cur_row, 2, style, fill=row_fill, align=left_a, border=tb())
                sc(ws, cur_row, 3, data['color'], fill=row_fill, align=left_a, border=tb())
                row_total = 0
                for i, size in enumerate(['S','M','L','XL','2XL','3XL'], 4):
                    qty = data['sizes'][size]
                    sc(ws, cur_row, i, qty if qty > 0 else '', fill=row_fill, align=center, border=tb())
                    row_total += qty
                sc(ws, cur_row, 10, row_total, fill=row_fill, align=center, border=tb())
                sc(ws, cur_row, 11, box['weight'] if first_in_box else '', fill=row_fill, align=center, border=tb())
                cur_row += 1
        cur_row += 1
        ws.merge_cells(f'A{cur_row}:B{cur_row}')
        sc(ws, cur_row, 1, 'TOTAL', bold=True, fill=col_fill)
        sc(ws, cur_row, 10, total_qty, bold=True)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ════════════════════════════════════════════════════════
# ── 5 & 6. 시트 변환기
# ════════════════════════════════════════════════════════

# 실제 제품리스트 컬럼: SKU / 품목명 / Color / 사이즈 / 현재고
# 실제 제품시트 구조:
#   행1: 타이틀
#   행2: CATEGORY / PRODUCT NAME / STYLE NO. / COLOR / S / M / L / XL / 2XL / 3XL / C-TOT / S-TOT / BASE PRICE / ...
#   데이터: 품목명·카테고리는 첫 컬러 행에만, 나머지 컬러는 STYLE NO.+COLOR만 채워짐

SIZE_COLS_SHEET = ['S','M','L','XL','2XL','3XL']   # 시트의 사이즈 컬럼 순서
SIZE_ORDER_LIST = ['XS','S','M','L','XL','2XL','3XL','4XL','FREE']  # 리스트 정렬용

# ── 공통: SKU → (STYLE_NO, SIZE) 파싱 ──────────────────
def parse_sku(sku: str):
    """SKU 끝의 사이즈 suffix 분리. 긴 것 먼저 매칭."""
    for sz in ['3XL','2XL','XL','FREE','XS','S','M','L']:
        if str(sku).upper().endswith(sz):
            return str(sku)[:-len(sz)], sz
    return str(sku), ''


# ── SKU에서 시즌 추출 ────────────────────────────────────
def get_season_from_sku(sku: str) -> str:
    import re
    m = re.search(r'\d{3}[WwSs](\d{2})', str(sku))
    if m:
        y = int(m.group(1))
        return f'{y-1}{y}'
    return '기타'


# ── 5. 리스트 → 시트 (시즌별 시트 분리) ───────────────────
def list_to_sheet(df_raw: pd.DataFrame) -> io.BytesIO:
    """
    제품리스트 CSV (SKU / 품목명 / Color / 사이즈 / 현재고)
    → 제품시트 Excel — 시즌별 시트로 분리 (2526, 2425, 2324 ...)
    현재고가 모두 NaN이면 빈 칸 시트(주문서용)로 생성
    """
    from openpyxl.utils import get_column_letter
    df = df_raw.copy()

    # ── 컬럼 정규화
    rename = {}
    for c in df.columns:
        cl = c.strip().lower()
        if cl in ('품목명','style','item name','product name','description'):
            rename[c] = '품목명'
        elif cl in ('color','컬러','colour','색상'):
            rename[c] = 'Color'
        elif cl in ('사이즈','size','sz'):
            rename[c] = '사이즈'
        elif cl in ('현재고','qty','수량','quantity','current qty','currentqty'):
            rename[c] = '현재고'
        elif cl == 'sku':
            rename[c] = 'SKU'
    df = df.rename(columns=rename)

    required = {'품목명', 'Color', '사이즈'}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"필수 컬럼 없음: {missing}\n현재 컬럼: {list(df.columns)}")

    if '현재고' not in df.columns:
        df['현재고'] = None

    df['현재고'] = pd.to_numeric(df['현재고'], errors='coerce')
    df['품목명'] = df['품목명'].astype(str).str.strip()
    df['Color']  = df['Color'].astype(str).str.strip()
    df['사이즈'] = df['사이즈'].astype(str).str.strip()

    # ── 시즌 컬럼: SKU 있으면 SKU에서 추출, 없으면 전체 하나의 시트
    if 'SKU' in df.columns:
        df['_시즌'] = df['SKU'].apply(get_season_from_sku)
    else:
        df['_시즌'] = '전체'

    # 시즌 순서: 최신 → 오래된 순, 기타는 맨 뒤
    season_order = sorted(
        [s for s in df['_시즌'].unique() if s != '기타'],
        reverse=True
    )
    if '기타' in df['_시즌'].unique():
        season_order.append('기타')

    has_qty = df['현재고'].notna().any()

    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거

    for season in season_order:
        df_s = df[df['_시즌'] == season].copy()
        if df_s.empty:
            continue

        # 이 시즌에 있는 사이즈 — FREE는 헤더에서 제외 (병합 표시)
        all_sizes_in_season = df_s['사이즈'].unique().tolist()
        sizes = [s for s in SIZE_COLS_SHEET if s in all_sizes_in_season] + \
                [s for s in all_sizes_in_season if s not in SIZE_COLS_SHEET and s != 'FREE']
        # sizes = 헤더에 실제로 들어갈 사이즈 목록 (FREE 제외)

        # 컬럼 인덱스 정의
        # A=1(품목명) B=2(Color) C~(sizes) TOTAL S-TOT
        total_col = 3 + len(sizes)       # TOTAL 열
        stot_col  = 3 + len(sizes) + 1   # S-TOT 열

        size_start_col = 3
        size_end_col   = 2 + len(sizes)  # 마지막 사이즈 열 (inclusive)

        # 품목명 순서 보존
        style_order = list(dict.fromkeys(df_s['품목명'].tolist()))

        ws = wb.create_sheet(title=str(season)[:31])

        # 컬럼 너비
        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['B'].width = 24
        for i in range(len(sizes)):
            ws.column_dimensions[get_column_letter(3 + i)].width = 7
        ws.column_dimensions[get_column_letter(total_col)].width = 9   # TOTAL
        ws.column_dimensions[get_column_letter(stot_col)].width  = 9   # S-TOT

        # ── 헤더
        sc(ws, 1, 1, '품목명', bold=True, fill=col_fill, align=center, border=tb())
        sc(ws, 1, 2, 'Color',  bold=True, fill=col_fill, align=center, border=tb())
        for j, sz in enumerate(sizes, 3):
            sc(ws, 1, j, sz, bold=True, fill=col_fill, align=center, border=tb())
        sc(ws, 1, total_col, 'TOTAL', bold=True, fill=col_fill, align=center, border=tb())
        sc(ws, 1, stot_col,  'S-TOT', bold=True, fill=PatternFill('solid', fgColor='BDD7EE'),
           align=center, border=tb())

        cur_row = 2
        grand_total = 0

        # ── 그룹 키워드 → 색상 & 구분선 규칙 ──────────────────
        import re as _re
        GROUP_FILL = {
            'V2':    ('E2EFDA', 'C6E0B4'),   # 파스텔 초록 (메인, S-TOT)
            'EASY':  ('DDEBF7', 'BDD7EE'),   # 파스텔 파랑
            'ORBAN': ('FFEB9C', 'FFD966'),   # 파스텔 노랑
            'COLLAB':('FCE4D6', 'F4B183'),   # 파스텔 핑크 (콜라보)
            'OTHER': ('F2F2F2', 'D9D9D9'),   # 회색
        }

        def get_group(name: str) -> str:
            n = _re.sub(r'^\d+\.\s*', '', str(name)).strip()
            tok = n.split()[0] if n.split() else ''
            if tok == 'V2':     return 'V2'
            if tok == 'EASY':   return 'EASY'
            if tok == 'ORBAN':  return 'ORBAN'
            if tok.startswith('SG') and 'x' in tok.lower(): return 'COLLAB'
            if 'SPECIALGUEST' in n.upper() and 'X' in n.upper(): return 'COLLAB'
            return 'OTHER'

        prev_group    = None
        prev_subgroup = None

        def get_subgroup(name: str) -> str:
            n = _re.sub(r'^\d+\.\s*', '', str(name)).strip()
            tok = n.split()[0] if n.split() else ''
            return tok.upper() if tok.upper().startswith('SG') else ''

        for si, style in enumerate(style_order):
            cur_group    = get_group(style)
            cur_subgroup = get_subgroup(style) if cur_group == 'COLLAB' else ''

            # ── 빈 행 삽입
            need_gap = 0
            if prev_group is not None:
                if cur_group != prev_group:
                    need_gap = 2 if (cur_group in ('V2','EASY','ORBAN') or
                                     prev_group in ('V2','EASY','ORBAN')) else 1
                elif cur_group == 'COLLAB' and cur_subgroup != prev_subgroup:
                    need_gap = 1
            for _ in range(need_gap):
                for col in range(1, stot_col + 1):
                    ws.cell(row=cur_row, column=col).value = None
                cur_row += 1
            prev_group    = cur_group
            prev_subgroup = cur_subgroup

            sdf = df_s[df_s['품목명'] == style]
            color_order = list(dict.fromkeys(sdf['Color'].tolist()))
            style_first_row = cur_row
            style_total = 0

            # ── 행 배경색: 그룹색 기반 교대
            main_hex, stot_hex = GROUP_FILL[cur_group]
            # 같은 그룹 내 품목 인덱스로 교대
            grp_styles = [s for s in style_order if get_group(s) == cur_group]
            grp_idx = grp_styles.index(style)
            if grp_idx % 2 == 0:
                row_fill  = PatternFill('solid', fgColor=main_hex)
                stot_fill = PatternFill('solid', fgColor=stot_hex)
            else:
                # 교대: 살짝 밝게
                lighter = 'F7FCF5' if cur_group=='V2' else ('EEF5FB' if cur_group=='EASY' else ('FFFDE7' if cur_group=='ORBAN' else ('FDF2EC' if cur_group=='COLLAB' else 'FAFAFA')))
                row_fill  = PatternFill('solid', fgColor=lighter)
                stot_fill = PatternFill('solid', fgColor=main_hex)

            # 품목 유형 판별
            item_sizes = set(sdf['사이즈'].unique())
            is_free    = item_sizes <= {'FREE'}
            is_mittens = bool(item_sizes & {'SM', 'LXL'})  # SM/LXL 장갑류

            # 장갑 병합 범위 계산 (시즌에 S가 있으면 S+M+L / 없으면 M+L)
            has_s_in_season = 'S' in sizes
            if is_mittens:
                # SM 그룹: S있으면 S~L(3칸), 없으면 M~L(2칸)
                sm_start = sizes.index('S') + 3 if has_s_in_season else sizes.index('M') + 3
                sm_end   = sizes.index('L') + 3
                # LXL 그룹: XL~3XL (있는 것까지)
                lxl_start = sizes.index('XL') + 3
                lxl_end   = sizes.index('3XL') + 3 if '3XL' in sizes else (sizes.index('2XL') + 3 if '2XL' in sizes else lxl_start)

            for ci, color in enumerate(color_order):
                cdf = sdf[sdf['Color'] == color]
                pivot = cdf.groupby('사이즈')['현재고'].sum()

                row_total_raw = pivot.sum()
                row_total = int(row_total_raw) if pd.notna(row_total_raw) and row_total_raw > 0 else 0
                style_total += row_total

                name_val = style if ci == 0 else ''
                sc(ws, cur_row, 1, name_val, fill=row_fill, align=left_a, border=tb())
                sc(ws, cur_row, 2, color,    fill=row_fill, align=left_a, border=tb())

                if is_free:
                    # 사이즈 열 전체 병합 → FREE 수량 표시
                    free_raw = pivot.get('FREE', None)
                    free_val = int(free_raw) if has_qty and pd.notna(free_raw) and free_raw > 0 else ''
                    for j in range(size_start_col, size_end_col + 1):
                        sc(ws, cur_row, j, '', fill=row_fill, border=tb())
                    ws.merge_cells(
                        start_row=cur_row, start_column=size_start_col,
                        end_row=cur_row,   end_column=size_end_col
                    )
                    mc = ws.cell(row=cur_row, column=size_start_col, value=free_val)
                    mc.font      = Font(name='Arial', bold=True, size=9)
                    mc.fill      = row_fill
                    mc.alignment = Alignment(horizontal='center', vertical='center')
                    mc.border    = Border(left=thin, right=thin, top=thin, bottom=thin)

                elif is_mittens:
                    # SM 그룹 병합
                    sm_raw = pivot.get('SM', None)
                    sm_val = int(sm_raw) if has_qty and pd.notna(sm_raw) and sm_raw > 0 else ''
                    for j in range(size_start_col, size_end_col + 1):
                        sc(ws, cur_row, j, '', fill=row_fill, border=tb())
                    ws.merge_cells(start_row=cur_row, start_column=sm_start,
                                   end_row=cur_row,   end_column=sm_end)
                    mc = ws.cell(row=cur_row, column=sm_start, value=sm_val)
                    mc.font = Font(name='Arial', bold=True, size=9)
                    mc.fill = row_fill
                    mc.alignment = Alignment(horizontal='center', vertical='center')
                    mc.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                    # LXL 그룹 병합
                    lxl_raw = pivot.get('LXL', None)
                    lxl_val = int(lxl_raw) if has_qty and pd.notna(lxl_raw) and lxl_raw > 0 else ''
                    ws.merge_cells(start_row=cur_row, start_column=lxl_start,
                                   end_row=cur_row,   end_column=lxl_end)
                    mc2 = ws.cell(row=cur_row, column=lxl_start, value=lxl_val)
                    mc2.font = Font(name='Arial', bold=True, size=9)
                    mc2.fill = row_fill
                    mc2.alignment = Alignment(horizontal='center', vertical='center')
                    mc2.border = Border(left=thin, right=thin, top=thin, bottom=thin)

                else:
                    for j, sz in enumerate(sizes, 3):
                        raw = pivot.get(sz, None)
                        val = int(raw) if has_qty and pd.notna(raw) and raw > 0 else ''
                        sc(ws, cur_row, j, val, fill=row_fill, align=center, border=tb())

                # TOTAL 열
                sc(ws, cur_row, total_col, row_total if has_qty else '',
                   bold=True, fill=row_fill, align=center, border=tb())
                # S-TOT 열: 빈 칸으로 채워두고 나중에 병합
                sc(ws, cur_row, stot_col, '', fill=stot_fill, border=tb())
                cur_row += 1

            grand_total += style_total

            # 품목명 A열 병합 (여러 컬러)
            if len(color_order) > 1:
                ws.merge_cells(f'A{style_first_row}:A{cur_row - 1}')
                ws.cell(row=style_first_row, column=1).alignment = Alignment(
                    horizontal='left', vertical='center', wrap_text=True
                )

            # S-TOT 열 병합 + 품목 합계 표시
            stot_val = style_total if has_qty else ''
            if len(color_order) > 1:
                ws.merge_cells(
                    start_row=style_first_row, start_column=stot_col,
                    end_row=cur_row - 1,       end_column=stot_col
                )
            mc = ws.cell(row=style_first_row, column=stot_col, value=stot_val)
            mc.font      = Font(name='Arial', bold=True, size=9)
            mc.fill      = stot_fill
            mc.alignment = Alignment(horizontal='center', vertical='center')
            mc.border    = Border(left=thin, right=thin, top=thin, bottom=thin)

        # ── GRAND TOTAL 행
        cur_row += 1
        ws.merge_cells(f'A{cur_row}:B{cur_row}')
        sc(ws, cur_row, 1, 'GRAND TOTAL', bold=True, fill=col_fill, align=center)
        for j, sz in enumerate(sizes, 3):
            if has_qty:
                sz_total = int(df_s[df_s['사이즈'] != 'FREE'].groupby('사이즈')['현재고'].sum().get(sz, 0))
                sc(ws, cur_row, j, sz_total if sz_total > 0 else '', bold=True, fill=col_fill, align=center)
            else:
                sc(ws, cur_row, j, '', bold=True, fill=col_fill, align=center)
        total_val = grand_total if has_qty else ''
        sc(ws, cur_row, total_col, total_val, bold=True, fill=col_fill, align=center)
        sc(ws, cur_row, stot_col,  total_val, bold=True, fill=PatternFill('solid', fgColor='BDD7EE'), align=center)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── 6. 시트 → 리스트 ─────────────────────────────────────
def sheet_to_list(xlsx_file, base_csv: pd.DataFrame = None) -> pd.DataFrame:
    """
    제품시트 Excel (품목명×Color 피벗) → 제품리스트 CSV
    base_csv가 있으면 기존 리스트의 SKU 매핑을 유지하고 현재고만 채워서 반환
    base_csv가 없으면 SKU 없이 품목명/Color/사이즈/현재고만 반환
    """
    wb = load_workbook(xlsx_file, data_only=True)
    ws = wb.active  # 첫 번째 시트 사용
    rows = list(ws.iter_rows(values_only=True))

    # ── 헤더 행 찾기: '품목명' 또는 'STYLE' 또는 'PRODUCT NAME' 포함 행
    header_idx = None
    header = []
    for i, row in enumerate(rows):
        row_upper = [str(c).strip().upper() if c is not None else '' for c in row]
        if any(k in row_upper for k in ('품목명', 'STYLE', 'PRODUCT NAME', 'COLOR')):
            header_idx = i
            header = row_upper
            break

    if header_idx is None:
        raise ValueError("헤더 행을 찾을 수 없습니다. '품목명' 또는 'Color' 컬럼이 있어야 합니다.")

    # ── 컬럼 인덱스 파악
    style_col = next((i for i, h in enumerate(header)
                      if h in ('품목명', 'STYLE', 'PRODUCT NAME', 'DESCRIPTION')), None)
    color_col  = next((i for i, h in enumerate(header)
                      if h in ('COLOR', 'COLOUR', '컬러', '색상')), None)
    if style_col is None or color_col is None:
        raise ValueError(f"품목명/Color 컬럼을 찾을 수 없습니다. 헤더: {header}")

    # 사이즈 컬럼: 헤더에서 S/M/L/XL/2XL/3XL 등 찾기
    size_col_map = {}  # {col_idx: size_name}
    skip_keywords = {'TOTAL','C-TOT','S-TOT','GRAND','BASE PRICE','PRICE','RETAIL',
                     'SUPPLY','CHINA','CATEGORY','PRODUCT','STYLE','COLOR','COLOUR','품목명','컬러',''}
    for i, h in enumerate(header):
        if h in skip_keywords:
            continue
        # S/M/L/XL/2XL/3XL 등 사이즈
        if h in [s.upper() for s in SIZE_ORDER_LIST]:
            size_col_map[i] = h
        # 헤더 상 숫자가 아닌 짧은 문자열 (추가 사이즈)
        elif len(h) <= 5 and i > color_col and 'TOT' not in h and 'PRICE' not in h:
            size_col_map[i] = h

    if not size_col_map:
        raise ValueError(f"사이즈 컬럼을 찾을 수 없습니다. 헤더: {header}")

    # ── 데이터 파싱 (품목명 forward-fill)
    records = []
    last_style = ''

    for row in rows[header_idx + 1:]:
        if all(c is None or str(c).strip() == '' for c in row):
            continue

        style_val = row[style_col] if style_col < len(row) else None
        color_val  = row[color_col]  if color_col  < len(row) else None

        # 품목명 forward-fill (첫 컬러에만 있고 나머지는 비어있음)
        if style_val and str(style_val).strip():
            s = str(style_val).strip()
            if s.upper() not in ('TOTAL','GRAND TOTAL',''):
                last_style = s
        style_val = last_style

        if not style_val or not color_val or str(color_val).strip() == '':
            continue
        if str(color_val).strip().upper() in ('TOTAL','GRAND TOTAL','COLOR'):
            continue

        color_val = str(color_val).strip()

        for col_idx, size_name in size_col_map.items():
            if col_idx >= len(row):
                continue
            cell_val = row[col_idx]
            if cell_val is None or str(cell_val).strip() == '':
                qty = None
            else:
                try:
                    qty = int(float(str(cell_val)))
                except (ValueError, TypeError):
                    continue

            records.append({
                '품목명': style_val,
                'Color':  color_val,
                '사이즈': size_name,
                '현재고': qty,
            })

    if not records:
        raise ValueError("변환된 데이터가 없습니다. 시트 형식을 확인해주세요.")

    df_sheet = pd.DataFrame(records)

    # ── base_csv가 있으면 SKU 매핑 후 현재고 채워서 반환
    if base_csv is not None:
        df_base = base_csv.copy()
        # 컬럼 정규화
        col_rename = {}
        for c in df_base.columns:
            cl = c.strip().lower()
            if cl in ('품목명','style','item name'): col_rename[c] = '품목명'
            elif cl in ('color','컬러','colour'):    col_rename[c] = 'Color'
            elif cl in ('사이즈','size','sz'):        col_rename[c] = '사이즈'
            elif cl in ('현재고','qty','수량','current qty'): col_rename[c] = '현재고'
            elif cl == 'sku': col_rename[c] = 'SKU'
        df_base = df_base.rename(columns=col_rename)

        # 시트 데이터를 lookup dict으로
        sheet_lookup = {}
        for _, r in df_sheet.iterrows():
            key = (str(r['품목명']).strip(), str(r['Color']).strip(), str(r['사이즈']).strip())
            sheet_lookup[key] = r['현재고']

        # 기존 리스트에 현재고 매핑
        def fill_qty(row):
            key = (str(row['품목명']).strip(), str(row['Color']).strip(), str(row['사이즈']).strip())
            return sheet_lookup.get(key, None)

        df_base['현재고'] = df_base.apply(fill_qty, axis=1)

        # 원본 컬럼 순서 복원
        out_cols = [c for c in ['SKU','품목명','Color','사이즈','현재고'] if c in df_base.columns]
        return df_base[out_cols]

    # ── base_csv 없으면 시트 파싱 결과 그대로 반환
    return df_sheet[['품목명','Color','사이즈','현재고']]


# ════════════════════════════════════════════════════════
# MAIN UI — 탭 분리
# ════════════════════════════════════════════════════════

tab1, tab2 = st.tabs(["📦 패킹 문서 생성기", "🔄 시트 변환기"])


# ────────────────────────────────────────────────────────
# TAB 1 : 기존 패킹 문서 생성기
# ────────────────────────────────────────────────────────
with tab1:
    st.caption("해외 출고리스트 업로드 → 패킹리스트 + 인보이스 + Actual Packing List + 카테고리별 시트 자동 생성")
    st.divider()

    col1, col2 = st.columns([1, 1])
    with col1:
        st.subheader("📁 파일 업로드")
        csv_file = st.file_uploader("해외 출고리스트 (.csv)", type=['csv'], key="tab1_upload")
    with col2:
        st.subheader("⚙️ 설정")
        messrs      = st.text_input("거래처명 (MESSRS)", placeholder="예: 傲刃堂（成都）...", key="messrs")
        destination = st.text_input("Final Destination", value="China", key="dest")
        date_str    = st.text_input("DATE", value=datetime.now().strftime("%Y/%m/%d"), key="date")

    st.divider()

    if csv_file:
        with st.spinner("처리 중..."):
            try:
                df, boxes = parse_export_csv(csv_file)
                total_qty    = df['수량'].sum()
                total_boxes  = len(boxes)
                total_weight = df.groupby('박스번호')['무게(kg)'].first().sum()
                total_cbm    = round(total_boxes * CBM_PER_BOX, 3)

                st.success(f"✅ 파싱 완료 — **{total_boxes} CTNs / {total_qty} pcs / {total_weight}kg / CBM {total_cbm}**")

                with st.expander("📊 카테고리별 요약"):
                    cat_summary = df.groupby('Category')['수량'].sum().sort_index()
                    for cat, qty in cat_summary.items():
                        st.write(f"**{cat}**: {qty}pcs")

                st.divider()
                st.subheader("📥 다운로드")
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.download_button("⬇️ 패킹리스트", data=make_packing_list(boxes, destination),
                        file_name=f"PackingList_{date_str.replace('/','')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                with c2:
                    st.download_button("⬇️ 인보이스", data=make_invoice(df, messrs, destination, date_str),
                        file_name=f"Invoice_{date_str.replace('/','')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                with c3:
                    st.download_button("⬇️ Actual Packing List", data=make_actual_packing_list(boxes, messrs, destination, date_str),
                        file_name=f"ActualPackingList_{date_str.replace('/','')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                with c4:
                    st.download_button("⬇️ 카테고리별", data=make_category_packing_list(boxes),
                        file_name=f"CategoryPackingList_{date_str.replace('/','')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
            except Exception as e:
                st.error(f"❌ 오류: {e}")
                import traceback
                st.code(traceback.format_exc())
    else:
        st.info("👆 해외 출고리스트 CSV 파일을 업로드해주세요.")


# ────────────────────────────────────────────────────────
# TAB 2 : 시트 변환기
# ────────────────────────────────────────────────────────
with tab2:
    st.caption("제품 리스트 ↔ 제품 시트 양방향 변환")

    sub_tab1, sub_tab2 = st.tabs(["📋 리스트 → 시트", "📊 시트 → 리스트"])

    # ── 2-1. 리스트 → 시트 ───────────────────────────────
    with sub_tab1:
        st.markdown(
            "**제품리스트 CSV** → **제품 시트 Excel**  \n"
            "품목명 × Color 행, S/M/L/XL/2XL/3XL 열 피벗으로 변환합니다.  \n"
            "현재고가 있으면 수량 채워서 생성, 없으면 빈 칸 시트(주문서용)로 생성합니다."
        )
        st.divider()

        l2s_file = st.file_uploader(
            "제품리스트 업로드 (.csv 또는 .xlsx)",
            type=['csv','xlsx'], key="l2s_upload"
        )

        if l2s_file:
            try:
                if l2s_file.name.endswith('.csv'):
                    df_l2s = pd.read_csv(l2s_file, encoding='utf-8-sig')
                else:
                    df_l2s = pd.read_excel(l2s_file)

                total_rows = len(df_l2s)
                has_qty = '현재고' in df_l2s.columns and df_l2s['현재고'].notna().any()
                styles_count = df_l2s.iloc[:,1].nunique() if len(df_l2s.columns) > 1 else 0

                c1, c2, c3 = st.columns(3)
                c1.metric("총 행수", f"{total_rows:,}")
                c2.metric("품목 수", f"{styles_count}")
                c3.metric("현재고", "있음 ✅" if has_qty else "없음 (빈 칸 시트)")

                with st.expander("미리보기 (상위 10행)"):
                    st.dataframe(df_l2s.head(10), use_container_width=True)

                if st.button("▶️ 시트로 변환", key="l2s_btn", use_container_width=True, type="primary"):
                    with st.spinner("변환 중..."):
                        result_buf = list_to_sheet(df_l2s)
                        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
                        st.success(f"✅ 변환 완료! — {styles_count}개 품목")
                        st.download_button(
                            "⬇️ 제품 시트 Excel 다운로드",
                            data=result_buf,
                            file_name=f"제품시트_{ts}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"❌ 오류: {e}")
                import traceback
                st.code(traceback.format_exc())
        else:
            st.info("👆 제품리스트 CSV 또는 Excel 파일을 업로드해주세요.")

    # ── 2-2. 시트 → 리스트 ───────────────────────────────
    with sub_tab2:
        st.markdown(
            "**제품 시트 Excel** → **제품리스트 CSV**  \n"
            "시트에서 현재고를 읽어, 기존 제품리스트의 `현재고` 컬럼을 채워서 반환합니다.  \n"
            "기존 리스트를 함께 업로드하면 SKU가 유지된 채로 현재고만 업데이트됩니다."
        )
        st.divider()

        col_a, col_b = st.columns(2)
        with col_a:
            s2l_sheet = st.file_uploader(
                "① 제품 시트 Excel (.xlsx)  ← 현재고가 채워진 시트",
                type=['xlsx'], key="s2l_sheet"
            )
        with col_b:
            s2l_base = st.file_uploader(
                "② 기존 제품리스트 CSV (선택) ← SKU 유지용",
                type=['csv','xlsx'], key="s2l_base"
            )

        if s2l_sheet:
            try:
                # 기존 리스트 로드 (선택)
                df_base = None
                if s2l_base:
                    if s2l_base.name.endswith('.csv'):
                        df_base = pd.read_csv(s2l_base, encoding='utf-8-sig')
                    else:
                        df_base = pd.read_excel(s2l_base)
                    st.info(f"📎 기존 리스트 로드됨 — {len(df_base):,}행 (SKU 매핑 적용)")

                if st.button("▶️ 리스트로 변환", key="s2l_btn", use_container_width=True, type="primary"):
                    with st.spinner("변환 중..."):
                        df_result = sheet_to_list(s2l_sheet, base_csv=df_base)

                        total_filled = df_result['현재고'].notna().sum()
                        total_rows   = len(df_result)
                        total_qty    = int(df_result['현재고'].fillna(0).sum())

                        st.success(f"✅ 변환 완료 — {total_rows:,}행 / 현재고 매핑 {total_filled:,}건 / 총 {total_qty:,}pcs")

                        with st.expander("미리보기 (상위 20행)"):
                            st.dataframe(df_result.head(20), use_container_width=True)

                        # 요약: 현재고 있는 것만
                        filled = df_result[df_result['현재고'].notna() & (df_result['현재고'] > 0)]
                        if not filled.empty:
                            with st.expander("📊 품목별 수량 요약"):
                                summary = filled.groupby('품목명')['현재고'].sum().reset_index()
                                summary.columns = ['품목명','수량합계']
                                st.dataframe(summary, use_container_width=True)

                        ts = datetime.now().strftime('%Y%m%d_%H%M%S')

                        # CSV 다운로드
                        csv_out = df_result.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            "⬇️ 제품리스트 CSV 다운로드",
                            data=csv_out.encode('utf-8-sig'),
                            file_name=f"제품리스트_{ts}.csv",
                            mime="text/csv",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"❌ 오류: {e}")
                import traceback
                st.code(traceback.format_exc())
        else:
            st.info("👆 현재고가 채워진 제품 시트 Excel을 업로드해주세요.")
