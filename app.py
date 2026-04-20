import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import OrderedDict, defaultdict
import io
from datetime import datetime

st.set_page_config(page_title="SG Export Document Generator", layout="centered")
st.title("📦 SPECIALGUEST Export Document Generator")
st.caption("해외 출고리스트 업로드 → 패킹리스트 + 인보이스 + Actual Packing List + 카테고리별 시트 자동 생성")

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

# ── 품목명 표준화 (Style 기반) ──────────────────────────────
def standardize_category(style_raw):
    """Style 값을 표준 카테고리로 매핑"""
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
    
    return style_raw  # 매칭 안되면 원본 반환

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

# ── CSV 데이터 파싱 ───────────────────────────────────────
def parse_export_csv(csv_file):
    """해외 출고리스트 CSV 파싱"""
    df = pd.read_csv(csv_file)
    
    # 박스번호 forward fill (수정됨)
    df['박스번호'] = df['박스번호'].ffill()
    df['무게(kg)'] = df['무게(kg)'].ffill()
    
    # 카테고리 표준화
    df['Category'] = df['Style'].apply(standardize_category)
    
    # 박스별 그룹화
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

# ── 1. 패킹리스트 생성 ────────────────────────────────────
def make_packing_list(boxes, destination):
    """패킹리스트 생성 (원본 양식)"""
    wb = Workbook()
    ws = wb.active
    ws.title = destination
    
    # 컬럼 너비
    ws.column_dimensions['A'].width = 9   # CTN NO.
    ws.column_dimensions['B'].width = 38  # STYLE
    ws.column_dimensions['C'].width = 20  # COLOR
    for col in ['D','E','F','G','H','I']:  # S~3XL
        ws.column_dimensions[col].width = 6
    ws.column_dimensions['J'].width = 8   # TOTAL
    ws.column_dimensions['K'].width = 9   # G.W
    
    # 헤더
    headers = ['CTN NO.','STYLE','COLOR','S','M','L','XL','2XL','3XL','TOTAL','G.W(kgs)']
    for i, h in enumerate(headers, 1):
        sc(ws, 1, i, h, bold=True, fill=col_fill, align=center, border=tb())
    
    # 데이터
    cur_row = 2
    for box_no in sorted(boxes.keys()):
        box = boxes[box_no]
        
        # 스타일별 사이즈 집계
        style_data = defaultdict(lambda: {
            'color': '', 'sizes': {'S':0,'M':0,'L':0,'XL':0,'2XL':0,'3XL':0}
        })
        
        for item in box['items']:
            key = item['item_name']
            style_data[key]['color'] = item['color']
            size = item['size']
            if size in style_data[key]['sizes']:
                style_data[key]['sizes'][size] += item['qty']
        
        # 박스번호는 첫 행에만
        first_in_box = True
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
            sc(ws, cur_row, 11, box['weight'] if cur_row == (cur_row - (cur_row - cur_row)) + 2 else '', 
               align=center, border=tb())
            cur_row += 1
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── 2. 인보이스 생성 ──────────────────────────────────────
def make_invoice(df, messrs, destination, date_str):
    """인보이스 생성"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Invoice'
    
    # 컬럼 너비
    for col, width in [('A',5),('B',10),('C',45),('D',12),('E',50),('F',12),('G',10),('H',10)]:
        ws.column_dimensions[col].width = width
    
    # 헤더 섹션
    ws.merge_cells('A1:H1')
    sc(ws, 1, 1, 'COMMERCIAL INVOICE', bold=True, size=14, align=center, fill=hdr_fill, 
       color='FFFFFF')
    ws.row_dimensions[1].height = 25
    
    sc(ws, 3, 1, 'SHIPPER:', bold=True)
    ws.merge_cells('A4:D4')
    sc(ws, 4, 1, 'SPECIALGUEST®', bold=True)
    ws.merge_cells('A5:D5')
    sc(ws, 5, 1, '22, Misagangbyeonhangang-ro 346beon-gil, Hanam-si, Gyeonggi-do, 12923, Republic of Korea')
    ws.merge_cells('A6:D6')
    sc(ws, 6, 1, 'Tel: +82 7077643333  Email: specialguest.co.kr@gmail.com')
    
    sc(ws, 8, 1, 'CONSIGNEE:', bold=True)
    ws.merge_cells('A9:D9')
    sc(ws, 9, 1, messrs, bold=True)
    
    sc(ws, 11, 1, 'DATE:', bold=True)
    sc(ws, 11, 2, date_str)
    sc(ws, 12, 1, 'DESTINATION:', bold=True)
    sc(ws, 12, 2, destination)
    
    # 테이블 헤더 (27행부터)
    headers = ['NO.','HS CODE','Description of goods','QTY','Fabric ratio','Unit Price(KRW)','Amount(KRW)','Remark']
    for i, h in enumerate(headers, 1):
        sc(ws, 27, i, h, bold=True, fill=col_fill, align=center, border=tb())
    
    # 데이터
    grouped = df.groupby(['품목명','HS Code','Material','단가(KRW)']).agg({'수량':'sum'}).reset_index()
    
    cur_row = 28
    total_qty = 0
    total_amount = 0
    
    for idx, row in enumerate(grouped.itertuples(), 1):
        sc(ws, cur_row, 1, idx, align=center, border=tb())
        sc(ws, cur_row, 2, row._2, align=center, border=tb())  # HS Code
        sc(ws, cur_row, 3, row._1, align=left_a, border=tb())   # 품목명
        sc(ws, cur_row, 4, row._5, align=center, border=tb())   # 수량
        sc(ws, cur_row, 5, row._3, align=left_a, border=tb())   # Material
        sc(ws, cur_row, 6, row._4, align=center, border=tb())   # 단가
        amount = row._5 * row._4
        sc(ws, cur_row, 7, amount, align=center, border=tb())
        sc(ws, cur_row, 8, '', border=tb())
        
        total_qty += row._5
        total_amount += amount
        cur_row += 1
    
    # 합계
    cur_row += 1
    ws.merge_cells(f'A{cur_row}:C{cur_row}')
    sc(ws, cur_row, 1, 'TOTAL', bold=True, align=center, fill=col_fill, border=tb())
    sc(ws, cur_row, 4, total_qty, bold=True, align=center, fill=col_fill, border=tb())
    sc(ws, cur_row, 7, total_amount, bold=True, align=center, fill=col_fill, border=tb())
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── 3. Actual Packing List 생성 ───────────────────────────
def make_actual_packing_list(boxes, messrs, destination, date_str):
    """Actual Packing List 생성"""
    wb = Workbook()
    ws = wb.active
    ws.title = 'ACTUAL PACKING LIST'
    
    # 컬럼 너비
    widths = {'A':16,'B':34,'C':30,'D':9,'E':52,'F':14,'G':11,'H':9}
    for col, w in widths.items(): 
        ws.column_dimensions[col].width = w
    
    # 헤더
    ws.merge_cells('A1:H1')
    sc(ws, 1, 1, 'ACTUAL PACKING LIST', bold=True, size=12, color='FFFFFF',
       fill=hdr_fill, align=center)
    ws.row_dimensions[1].height = 20
    
    sc(ws, 2, 1, 'SPECIALGUEST®', bold=True, size=10)
    ws.merge_cells('A3:H3')
    sc(ws, 3, 1, 'postal code: 12923 / 22, Misagangbyeonhangang-ro 346beon-gil, Hanam-si, Gyeonggi-do, Republic of Korea', 
       align=left_a)
    ws.merge_cells('A4:H4')
    sc(ws, 4, 1, 'Tel: +82 7077643333  Email: specialguest.co.kr@gmail.com', align=left_a)
    
    ws.merge_cells('A6:F6')
    sc(ws, 6, 1, f'MESSRS: {messrs}', bold=True)
    ws.merge_cells('G6:H6')
    sc(ws, 6, 7, f'DATE: {date_str}', size=8)
    
    ws.merge_cells('A7:F7')
    sc(ws, 7, 1, 'SHIPMENT FROM: Republic of Korea', bold=True)
    ws.merge_cells('A8:H8')
    sc(ws, 8, 1, f'FINAL DESTINATION: {destination}', bold=True)
    sc(ws, 9, 8, CBM_PER_BOX, bold=True, align=center)
    
    # 테이블 헤더
    headers = ['HS code','Description of goods','Fabric ratio','Quantity','** Comments',
               'Dimension (cm)','Weight(kg)','CBM']
    for i, h in enumerate(headers, 1):
        sc(ws, 10, i, h, bold=True, fill=col_fill, align=center, border=tb())
    ws.row_dimensions[10].height = 14
    
    # 카테고리별 집계
    category_data = defaultdict(lambda: {
        'hs_code': '', 'material': '', 'qty': 0, 'boxes': [], 'weight': 0
    })
    
    for box_no, box in boxes.items():
        for item in box['items']:
            cat = item['category']
            category_data[cat]['hs_code'] = item['hs_code']
            category_data[cat]['material'] = item['material']
            category_data[cat]['qty'] += item['qty']
            if box_no not in category_data[cat]['boxes']:
                category_data[cat]['boxes'].append(box_no)
                category_data[cat]['weight'] += box['weight']
    
    # 우선순위 정렬
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
        
        # CBM 수식
        cbm_cell = ws.cell(row=cur_row, column=8, value=f'=E{cur_row}*$H$9')
        cbm_cell.font = Font(name='Arial', bold=True, size=9)
        cbm_cell.fill = fill
        cbm_cell.alignment = center
        cbm_cell.border = tb()
        
        total_qty += data['qty']
        total_boxes += num_boxes
        total_weight += data['weight']
        cur_row += 1
    
    # 하단
    cur_row += 1
    sc(ws, cur_row, 1, 'Please check the ** comments.', bold=True)
    cur_row += 1
    sc(ws, cur_row, 1, 'N = Nylon / P = Polyester / PU = Polyurethane / C = Cotton / R = Rayon / A = Acrylic', 
       bold=True)
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

# ── 4. 카테고리별 패킹리스트 생성 ──────────────────────────
def make_category_packing_list(boxes):
    """카테고리별 패킹리스트 생성"""
    wb = Workbook()
    wb.remove(wb.active)  # 기본 시트 제거
    
    # 박스별 주 카테고리 결정
    box_main_category = {}
    for box_no, box in boxes.items():
        cat_qty = defaultdict(int)
        for item in box['items']:
            cat_qty[item['category']] += item['qty']
        main_cat = max(cat_qty.items(), key=lambda x: (x[1], -get_prio(x[0])))[0]
        box_main_category[box_no] = main_cat
    
    # 카테고리별 박스 그룹화
    category_boxes = defaultdict(list)
    for box_no, main_cat in box_main_category.items():
        category_boxes[main_cat].append(box_no)
    
    # 시트 생성
    for cat in SHEET_ORDER:
        if cat not in category_boxes:
            continue
        
        ws = wb.create_sheet(cat[:31])
        main_fill = PatternFill('solid', fgColor=MAIN_COLORS.get(cat, 'FFFFFF'))
        
        # 컬럼 너비
        ws.column_dimensions['A'].width = 9
        ws.column_dimensions['B'].width = 38
        ws.column_dimensions['C'].width = 20
        for c in ['D','E','F','G','H','I']:
            ws.column_dimensions[c].width = 6
        ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 9
        
        # 헤더
        headers = ['CTN NO.','STYLE','COLOR','S','M','L','XL','2XL','3XL','TOTAL','G.W(kgs)']
        for i, h in enumerate(headers, 1):
            sc(ws, 1, i, h, bold=True, fill=col_fill, align=center, border=tb())
        
        # 데이터
        cur_row = 2
        total_qty = 0
        
        for box_no in sorted(category_boxes[cat]):
            box = boxes[box_no]
            
            # 스타일별 사이즈 집계
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
                # 색상 결정
                if data['category'] == cat:
                    row_fill = main_fill
                    if data['sizes'].values():
                        total_qty += sum(data['sizes'].values())
                else:
                    row_fill = PatternFill('solid', fgColor=SEC_COLORS.get(data['category'], 'EEEEEE'))
                
                # 박스번호
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
                sc(ws, cur_row, 11, box['weight'] if first_in_box else '', fill=row_fill, 
                   align=center, border=tb())
                cur_row += 1
        
        # TOTAL 행
        cur_row += 1
        ws.merge_cells(f'A{cur_row}:B{cur_row}')
        sc(ws, cur_row, 1, 'TOTAL', bold=True, fill=col_fill)
        sc(ws, cur_row, 10, total_qty, bold=True)
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ════════════════════════════════════════════════════════
# UI
# ════════════════════════════════════════════════════════
st.divider()

col1, col2 = st.columns([1, 1])
with col1:
    st.subheader("📁 파일 업로드")
    csv_file = st.file_uploader("해외 출고리스트 (.csv)", type=['csv'])

with col2:
    st.subheader("⚙️ 설정")
    messrs = st.text_input("거래처명 (MESSRS)", placeholder="예: 傲刃堂（成都）...")
    destination = st.text_input("Final Destination", value="China")
    date_str = st.text_input("DATE", value=datetime.now().strftime("%Y/%m/%d"))

st.divider()

if csv_file:
    with st.spinner("처리 중..."):
        try:
            df, boxes = parse_export_csv(csv_file)
            
            # 통계
            total_qty = df['수량'].sum()
            total_boxes = len(boxes)
            total_weight = df.groupby('박스번호')['무게(kg)'].first().sum()
            total_cbm = round(total_boxes * CBM_PER_BOX, 3)
            
            st.success(f"✅ 파싱 완료 — **{total_boxes} CTNs / {total_qty} pcs / {total_weight}kg / CBM {total_cbm}**")
            
            # 카테고리별 요약
            with st.expander("📊 카테고리별 요약"):
                cat_summary = df.groupby('Category')['수량'].sum().sort_index()
                for cat, qty in cat_summary.items():
                    st.write(f"**{cat}**: {qty}pcs")
            
            st.divider()
            st.subheader("📥 다운로드")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                pack_buf = make_packing_list(boxes, destination)
                st.download_button(
                    "⬇️ 패킹리스트",
                    data=pack_buf,
                    file_name=f"PackingList_{date_str.replace('/', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col2:
                inv_buf = make_invoice(df, messrs, destination, date_str)
                st.download_button(
                    "⬇️ 인보이스",
                    data=inv_buf,
                    file_name=f"Invoice_{date_str.replace('/', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col3:
                actual_buf = make_actual_packing_list(boxes, messrs, destination, date_str)
                st.download_button(
                    "⬇️ Actual Packing List",
                    data=actual_buf,
                    file_name=f"ActualPackingList_{date_str.replace('/', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col4:
                cat_buf = make_category_packing_list(boxes)
                st.download_button(
                    "⬇️ 카테고리별",
                    data=cat_buf,
                    file_name=f"CategoryPackingList_{date_str.replace('/', '')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        except Exception as e:
            st.error(f"❌ 오류: {e}")
            import traceback
            st.code(traceback.format_exc())
else:
    st.info("👆 해외 출고리스트 CSV 파일을 업로드해주세요.")
