import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import OrderedDict
import re, io

st.set_page_config(page_title="SG Packing Generator", layout="centered")
st.title("📦 SPECIALGUEST Packing List Generator")
st.caption("패킹리스트 + 인보이스 업로드 → Actual Packing List + 카테고리별 시트 자동 생성")

CBM_PER_BOX = 0.088

# ── Description 분류 ──────────────────────────────────────
def get_desc_info(hs, fabric):
    hs_s = str(hs).strip()
    fab  = str(fabric).strip() if fabric else ''
    fabu = fab.upper()
    fabn = fab.replace(' ', '')

    if hs_s.startswith('6201'):
        if 'FILLING' in fabu or 'DOWN' in fabu:
            return 'Insulated Snowboard Jacket', 'Shell N100 / Lining P100 / Filling P100', '6201.40.9011'
        return 'Snowboard Jacket', 'Shell N100 / Lining P100', '6201.40.9011'
    elif hs_s.startswith('6203'):
        if 'C 80' in fabu or 'C80' in fabu:
            return 'Sweatpants', 'Shell C80 / P20', '6203.43-0000'
        return 'Snowboard Pants', 'Shell N100 / Lining P100', '6203.43-0000'
    elif hs_s.startswith('6205'):
        return 'Long-sleeve Tee', 'P63 / R31 / PU6 , Rib-P100', '6205.30-0000'
    elif hs_s.startswith('6110'):
        return 'Sweatshirt, Hooded Sweatshirt', 'C80 / P20', '6110.20-0000'
    elif hs_s.startswith('6109'):
        return 'T-shirt', 'C100', '6109.10-0000'
    elif hs_s.startswith('6505'):
        if '88' in fabn and 'S12' in fabn.replace(' ', ''):
            return 'BALACLAVA', 'Shell C 88 / S 12', '6505.00-0000'
        if fabn == 'P100':
            return 'HOOD WARMER N', 'P100', '6505.00-0000'
        if 'ANGORA' in fabu:
            return 'Beanie(Angora)', 'Shell A 70 + Angora 30', '6505.00-0000'
        if 'LINING' in fabu or ('SHELL' in fabu and 'C' in fabu and '100' in fabu):
            if 'N100' in fabn:
                return 'Cap', 'Shell N100 / Lining P100', '6505.00-0000'
            return 'Cap', 'Shell C100 / Lining P100', '6505.00-0000'
        return 'Beanie', 'A100', '6505.00-0000'
    elif hs_s.startswith('3926'):
        return 'Snowboard Stomppad', 'Silicone processing', '3926.90-9000'
    return None, fab, hs_s

PRIORITY = [
    'Insulated Snowboard Jacket', 'Snowboard Jacket', 'Snowboard Pants',
    'Sweatpants', 'Sweatshirt, Hooded Sweatshirt', 'Long-sleeve Tee', 'T-shirt',
    'Beanie', 'Beanie(Angora)', 'Cap', 'HOOD WARMER N', 'BALACLAVA', 'Snowboard Stomppad'
]
def get_prio(d): return PRIORITY.index(d) if d in PRIORITY else 99

def get_main(items):
    totals = {}
    for it in items: totals[it['desc']] = totals.get(it['desc'], 0) + it['qty']
    mx = max(totals.values())
    cands = [d for d, q in totals.items() if q == mx]
    return sorted(cands, key=get_prio)[0]

# ── 인보이스 파싱 ─────────────────────────────────────────
def load_invoice_map(inv_bytes):
    wb = load_workbook(io.BytesIO(inv_bytes), data_only=True)
    ws = wb.active
    entries = []
    for row in ws.iter_rows(min_row=28, max_row=ws.max_row):
        hs = row[1].value; style = row[2].value; fabric = row[4].value
        if not hs or not style: continue
        s = str(style).strip()
        if s in ['Description of goods', '']: continue
        desc, fab_out, hs_out = get_desc_info(hs, fabric)
        if not desc: continue
        sm = re.match(r'^(\d{4})\s', s)
        nm = re.search(r'(\d+)\.', s)
        entries.append({
            'key': s, 'desc': desc, 'fabric': fab_out, 'hs': hs_out,
            'season': sm.group(1) if sm else None,
            'num': nm.group(1).lstrip('0') if nm else None
        })
    return entries

def find_info(pack_style, entries):
    s = pack_style.strip()
    for e in entries:
        if e['key'] == s: return e
    pm = re.match(r'^(\d{4}) (\d+)\.', s)
    if not pm: return None
    p_season, p_num = pm.group(1), pm.group(2).lstrip('0')
    for e in entries:
        if e['season'] == p_season and e['num'] == p_num: return e
    for e in entries:
        if e['season'] is None and e['num'] == p_num: return e
    for e in entries:
        if e.get('num') == p_num: return e
    return None

# ── 패킹리스트 파싱 ───────────────────────────────────────
def parse_packing(pack_bytes, entries):
    try:
        wb = load_workbook(io.BytesIO(pack_bytes), data_only=True, keep_vba=False)
    except:
        wb = load_workbook(io.BytesIO(pack_bytes), data_only=True)
    # 패킹리스트 시트 자동 감지
    PACK_SHEET_NAMES = ['중국', 'Sheet1', '한국', '일본', '미국', '중국시트']
    ws = None
    for sh in PACK_SHEET_NAMES:
        if sh in wb.sheetnames:
            ws = wb[sh]; break
    if not ws: ws = wb.active
    ctns = []; current_ctn = None; unmapped = set()

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ctn_no = str(row[0].value).strip() if row[0].value else ''
        style  = str(row[1].value).strip() if row[1].value else ''
        if ctn_no and ctn_no not in ['', 'CTN NO.']:
            gw = row[10].value if len(row) > 10 else None
            try: gw = float(gw) if gw else None
            except: gw = None
            current_ctn = {'ctn_no': ctn_no, 'gw': gw, 'items': [], 'raw_rows': []}
            ctns.append(current_ctn)
        if not current_ctn or not style: continue
        qty = 0
        for idx in range(3, 9):
            v = row[idx].value
            try: qty += int(v) if v else 0
            except: pass
        if qty == 0: continue
        current_ctn['raw_rows'].append([cell.value for cell in row])
        info = find_info(style, entries)
        if not info:
            unmapped.add(style); continue
        found = False
        for item in current_ctn['items']:
            if item['desc'] == info['desc']:
                item['qty'] += qty; found = True; break
        if not found:
            current_ctn['items'].append({
                'desc': info['desc'], 'hs': info['hs'],
                'fabric': info['fabric'], 'qty': qty
            })
    return [c for c in ctns if c['items']], unmapped

# ── 그룹 빌드 ─────────────────────────────────────────────
def build_groups(ctns):
    desc_order = []; desc_groups = OrderedDict()
    for ctn in ctns:
        main_desc = get_main(ctn['items'])
        main_item = next(it for it in ctn['items'] if it['desc'] == main_desc)
        if main_desc not in desc_groups:
            desc_groups[main_desc] = {
                'hs': main_item['hs'], 'fabric': main_item['fabric'],
                'main_qty': 0, 'ctns': [], 'total_gw': 0, 'sec_map': {}
            }
            desc_order.append(main_desc)
        g = desc_groups[main_desc]
        g['main_qty'] += main_item['qty']
        g['ctns'].append(ctn['ctn_no'])
        if ctn['gw']: g['total_gw'] += ctn['gw']
        for it in ctn['items']:
            if it['desc'] != main_desc:
                sd = it['desc']
                if sd not in g['sec_map']: g['sec_map'][sd] = {'qty': 0, 'ctns': []}
                g['sec_map'][sd]['qty'] += it['qty']
                g['sec_map'][sd]['ctns'].append(ctn['ctn_no'])
    return desc_order, desc_groups

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

# ── Actual Packing List 생성 ──────────────────────────────
def make_actual(desc_order, desc_groups, messrs, destination, date_str):
    wb = Workbook(); ws = wb.active
    ws.title = 'ACTUAL PACKING LIST'

    # 컬럼 너비 (A~M 기준 실제 양식)
    widths = {'A':16,'B':34,'C':30,'D':9,'E':52,'F':14,'G':11,'H':9,'I':9,'J':9,'K':9,'L':9,'M':9}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

    # 헤더 섹션
    ws.merge_cells('A1:H1')
    c = ws.cell(row=1, column=1, value='ACTUAL PACKING LIST')
    c.font = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    c.fill = hdr_fill; c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 20

    for r, txt in [
        (2, 'SPECIALGUEST\u00AE'),
        (3, 'postal code: 12923  / 22, Misagangbyeonhangang-ro 346beon-gil, Hanam-si, Gyeonggi-do, Republic of Korea'),
        (4, 'Tel : +82 7077643333       Email : specialguest.co.kr@gmail.com'),
    ]:
        ws.merge_cells(f'A{r}:H{r}')
        sc(ws, r, 1, txt, bold=(r==2), size=10 if r==2 else 9, align=left_a)

    ws.merge_cells('A6:F6'); sc(ws, 6, 1, f'MESSRS : {messrs}', bold=True)
    ws.merge_cells('G6:H6'); sc(ws, 6, 7, f'DATE : {date_str}', size=8)
    ws.merge_cells('A7:F7'); sc(ws, 7, 1, 'SHIPMENT FROM :   Republic of Korea', bold=True)
    ws.merge_cells('A8:H8'); sc(ws, 8, 1, f'FINAL DESTINATION :  {destination}', bold=True)
    sc(ws, 9, 8, CBM_PER_BOX, bold=True, align=center)

    # 컬럼 헤더 (실제 양식과 동일)
    hdrs = ['HS code','Description of goods','Fabric ratio','Ǫuantity','** Comments','Dimension (cm)','Weight(kg)','CBM']
    for i, h in enumerate(hdrs, 1):
        sc(ws, 10, i, h, bold=True, fill=col_fill, align=center, border=tb())
    ws.row_dimensions[10].height = 14

    cur = 11; use_alt = False
    grand_qty = grand_ctns = grand_gw = 0

    for desc in desc_order:
        g = desc_groups[desc]
        n = len(g['ctns']); gw = g['total_gw']
        fill = alt_fill if use_alt else no_fill; use_alt = not use_alt

        sc(ws, cur, 1, g['hs'],         bold=True, fill=fill, align=center, border=tb())
        sc(ws, cur, 2, desc,            bold=True, fill=fill, align=left_a,  border=tb())
        sc(ws, cur, 3, g['fabric'],     bold=True, fill=fill, align=left_a,  border=tb())
        sc(ws, cur, 4, g['main_qty'],   bold=True, fill=fill, align=center, border=tb())
        sc(ws, cur, 5, n,               bold=True, fill=fill, align=center, border=tb())  # CTNS (E열)
        sc(ws, cur, 6, '55 x 40 x 40', bold=True, fill=fill, align=center, border=tb())
        sc(ws, cur, 7, gw,              bold=True, fill=fill, align=center, border=tb())
        # CBM = CTNS × 0.088 수식
        cbm_cell = ws.cell(row=cur, column=8, value=f'=E{cur}*$H$9')
        cbm_cell.font = Font(name='Arial', bold=True, size=9)
        cbm_cell.fill = fill; cbm_cell.alignment = center; cbm_cell.border = tb()
        ws.row_dimensions[cur].height = 14

        grand_qty += g['main_qty']; grand_ctns += n; grand_gw += gw
        cur += 1

        # packed with 코멘트 행
        for sd, info in g['sec_map'].items():
            unique = list(dict.fromkeys(info['ctns']))
            refs   = ', '.join(f'CTN no.{c}' for c in unique)
            comment = f"** packed with '{sd}' ({refs}) {info['qty']}ea"
            for col in range(1, 9):
                ws.cell(row=cur, column=col).fill = fill
                ws.cell(row=cur, column=col).border = tb()
            sc(ws, cur, 4, info['qty'], bold=True,  fill=fill, align=center, border=tb())
            sc(ws, cur, 5, comment,     bold=False, fill=fill, align=left_a,  border=tb())
            ws.row_dimensions[cur].height = 14
            cur += 1

    # 하단
    cur += 1
    sc(ws, cur, 1, 'Please check the ** comments.', bold=True); cur += 1
    sc(ws, cur, 1, 'N = Nylon  /  P = Polyester  /  PU = Polyurethane  /  C = Cotton  /  R = Rayon   /  A = Acrylic', bold=True)
    cur += 2
    for i, h in enumerate(['Ǫuantity','CTNS','Dimension (cm)','Weight(kg)','CBM'], start=4):
        sc(ws, cur, i, h, bold=True, align=center)
    cur += 1
    ws.merge_cells(f'A{cur}:C{cur}'); sc(ws, cur, 1, 'Total', bold=True)
    sc(ws, cur, 4, grand_qty,                          bold=True, align=center)
    sc(ws, cur, 5, grand_ctns,                         bold=True, align=center)
    sc(ws, cur, 6, '55 x 40 x 40',                    bold=True, align=center)
    sc(ws, cur, 7, grand_gw,                           bold=True, align=center)
    sc(ws, cur, 8, round(grand_ctns * CBM_PER_BOX, 3), bold=True, align=center)
    cur += 2
    ws.merge_cells(f'D{cur}:H{cur}'); sc(ws, cur, 4, 'MADE IN KOREA', bold=True)

    # Sheet1 (액세서리 매핑 테이블 — 실제 양식 그대로)
    ws2 = wb.create_sheet('Sheet1')
    acc_data = [
        ('6505.00-0000', 'BEANIE',        'Acryl 100'),
        ('6505.00-0000', 'CAP',           'Shell C100 / Lining P100 / '),
        ('',             'BEANIE(Angora)','Shell A 70 + Angora 30'),
        ('6505.00-0000', 'BALACLAVA',     'Shell C 88 / S 12'),
        ('6505.00-0000', 'HOOD WARMER N', 'P100'),
    ]
    for r_idx, (hs, name, fab) in enumerate(acc_data, 1):
        ws2.cell(row=r_idx, column=1, value=hs)
        ws2.cell(row=r_idx, column=2, value=name)
        ws2.cell(row=r_idx, column=3, value=fab)

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── 카테고리별 패킹리스트 생성 ────────────────────────────
def make_category(pack_bytes, ctns):
    try:
        wb_orig = load_workbook(io.BytesIO(pack_bytes), data_only=True, keep_vba=False)
    except:
        wb_orig = load_workbook(io.BytesIO(pack_bytes), data_only=True)
    PACK_SHEET_NAMES = ['중국', 'Sheet1', '한국', '일본', '미국', '중국시트']
    ws_orig = None
    for sh in PACK_SHEET_NAMES:
        if sh in wb_orig.sheetnames:
            ws_orig = wb_orig[sh]; break
    if not ws_orig: ws_orig = wb_orig.active
    headers_orig = [ws_orig.cell(row=1, column=i).value for i in range(1, 12)]

    ctn_main = {}
    for ctn in ctns:
        totals = {}
        for it in ctn['items']: totals[it['desc']] = totals.get(it['desc'], 0) + it['qty']
        mx = max(totals.values())
        cands = [d for d, q in totals.items() if q == mx]
        ctn_main[ctn['ctn_no']] = sorted(cands, key=get_prio)[0]

    SHEET_ORDER = [
        'Snowboard Pants','Snowboard Jacket','Insulated Snowboard Jacket',
        'Sweatpants','Sweatshirt, Hooded Sweatshirt','Long-sleeve Tee','T-shirt',
        'Beanie','Beanie(Angora)','Cap','HOOD WARMER N','BALACLAVA','Snowboard Stomppad'
    ]
    COLORS = {
        'Snowboard Pants':'DCE6F1','Snowboard Jacket':'FEF9E7',
        'Insulated Snowboard Jacket':'EBF5FB','Sweatpants':'F9F2FF',
        'Sweatshirt, Hooded Sweatshirt':'FEF5E7','Long-sleeve Tee':'E9F7EF',
        'T-shirt':'FDF2F8','Beanie':'F0FFF4','Beanie(Angora)':'FFFDE7',
        'Cap':'F3E5F5','HOOD WARMER N':'E8EAF6','BALACLAVA':'FCE4EC',
        'Snowboard Stomppad':'E8F5E9',
    }

    sheet_rows = {sh: [] for sh in SHEET_ORDER}
    current_ctn = None
    for row in ws_orig.iter_rows(min_row=2, max_row=ws_orig.max_row):
        ctn_no = str(row[0].value).strip() if row[0].value else ''
        style  = str(row[1].value).strip() if row[1].value else ''
        if ctn_no and ctn_no not in ['', 'CTN NO.']: current_ctn = ctn_no
        if not current_ctn or not style: continue
        main = ctn_main.get(current_ctn)
        if main and main in sheet_rows:
            sheet_rows[main].append([cell.value for cell in row])

    wb_cat = Workbook()
    # 기본 시트 이름을 임시로 바꿔두고 마지막에 삭제
    wb_cat.active.title = '_temp'
    first_sheet = True
    for sh in SHEET_ORDER:
        if not sheet_rows.get(sh): continue
        if first_sheet:
            ws = wb_cat.active
            ws.title = sh[:31]
            first_sheet = False
        else:
            ws = wb_cat.create_sheet(sh[:31])
        sh_fill = PatternFill('solid', fgColor=COLORS.get(sh, 'FFFFFF'))

        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 44
        ws.column_dimensions['C'].width = 22
        for col in ['D','E','F','G','H','I']: ws.column_dimensions[col].width = 7
        ws.column_dimensions['J'].width = 8
        ws.column_dimensions['K'].width = 10

        for i, h in enumerate(headers_orig, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = Font(name='Arial', bold=True, size=9)
            c.fill = col_fill; c.alignment = center; c.border = tb()

        cur_row = 2
        for row_data in sheet_rows[sh]:
            ctn_val  = row_data[0]
            row_fill = sh_fill if ctn_val else no_fill
            for i, val in enumerate(row_data, 1):
                c = ws.cell(row=cur_row, column=i, value=val)
                c.font = Font(name='Arial', size=9, bold=(bool(ctn_val) and i == 1))
                c.alignment = center if i >= 4 else left_a
                c.border = tb(); c.fill = row_fill
            cur_row += 1

        total_qty = 0
        for r in sheet_rows[sh]:
            for idx in range(3, 9):
                try: total_qty += int(r[idx]) if r[idx] else 0
                except: pass
        cur_row += 1
        ws.merge_cells(f'A{cur_row}:B{cur_row}')
        c = ws.cell(row=cur_row, column=1, value='TOTAL')
        c.font = Font(name='Arial', bold=True, size=9); c.fill = col_fill
        tc = ws.cell(row=cur_row, column=10, value=total_qty)
        tc.font = Font(name='Arial', bold=True, size=9)

    buf = io.BytesIO(); wb_cat.save(buf); buf.seek(0)
    return buf

# ════════════════════════════════════════════════════════
# UI
# ════════════════════════════════════════════════════════
st.divider()

col1, col2 = st.columns([1, 1])
with col1:
    st.subheader("📁 파일 업로드")
    pack_file = st.file_uploader("패킹리스트 (.xlsx / .xlsm)", type=['xlsx', 'xlsm'])
    inv_file  = st.file_uploader("인보이스 (.xlsx)", type=['xlsx'])

with col2:
    st.subheader("⚙️ 설정")
    messrs      = st.text_input("거래처명 (MESSRS)", placeholder="예: 傲刃堂（成都）...")
    destination = st.text_input("Final Destination", value="China")
    date_str    = st.text_input("DATE", placeholder="예: 2026/04/09")

st.divider()

if pack_file and inv_file:
    with st.spinner("처리 중..."):
        try:
            pack_bytes = pack_file.read()
            inv_bytes  = inv_file.read()

            entries = load_invoice_map(inv_bytes)
            ctns, unmapped = parse_packing(pack_bytes, entries)
            desc_order, desc_groups = build_groups(ctns)

            total_qty  = sum(g['main_qty']     for g in desc_groups.values())
            total_ctns = sum(len(g['ctns'])     for g in desc_groups.values())
            total_gw   = sum(g['total_gw']      for g in desc_groups.values())
            total_cbm  = round(total_ctns * CBM_PER_BOX, 3)

            st.success(f"✅ 파싱 완료 —  **{total_ctns} CTNs  /  {total_qty} pcs  /  {total_gw} kg  /  CBM {total_cbm}**")

            if unmapped:
                with st.expander(f"⚠️ 매핑 안된 스타일 {len(unmapped)}개 (인보이스 확인 필요)", expanded=True):
                    for s in sorted(unmapped): st.text(f"  - {s}")

            with st.expander("📊 품목별 요약"):
                for desc in desc_order:
                    g = desc_groups[desc]
                    st.write(f"**{desc}** — {g['main_qty']}pcs / {len(g['ctns'])}CTNs / {g['total_gw']}kg")

            st.divider()
            st.subheader("📥 다운로드")
            dl1, dl2 = st.columns(2)

            with dl1:
                act_buf = make_actual(desc_order, desc_groups, messrs, destination, date_str)
                fname_act = f"ACTUAL_PACKING_LIST_{date_str.replace('/', '')}.xlsx"
                st.download_button(
                    "⬇️ Actual Packing List",
                    data=act_buf, file_name=fname_act,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with dl2:
                cat_buf = make_category(pack_bytes, ctns)
                fname_cat = f"PACKING_LIST_BY_CATEGORY_{date_str.replace('/', '')}.xlsx"
                st.download_button(
                    "⬇️ 카테고리별 패킹리스트",
                    data=cat_buf, file_name=fname_cat,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"❌ 오류: {e}")
            import traceback; st.code(traceback.format_exc())

elif pack_file and not inv_file:
    st.info("👆 인보이스 파일도 업로드해주세요.")
else:
    st.info("👆 패킹리스트와 인보이스 파일을 업로드해주세요.")
