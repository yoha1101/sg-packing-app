# ════════════════════════════════════════════════════════
# TAB 4: 💱 국가별 가격 생성기
# ════════════════════════════════════════════════════════

def make_country_pricing_sheet(product_file, country_list, exchange_rates):
    """
    프로덕트 시트에서 기본가격을 읽고 국가별 가격 시트 생성
    
    Args:
        product_file: 프로덕트 시트 Excel 파일
        country_list: 국가 정보 리스트 [{'code':'USD', 'name':'United States', ...}]
        exchange_rates: {'USD': 1.0, 'EUR': 1.1, ...} 형태의 환율 딕셔너리
    """
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import io
    
    # 프로덕트 시트 로드
    prod_bytes = product_file.read() if hasattr(product_file, 'read') else open(product_file, 'rb').read()
    wb_prod = load_workbook(io.BytesIO(prod_bytes), data_only=True)
    ws_prod = wb_prod.active
    
    # 데이터 추출
    products = []
    for row in range(4, ws_prod.max_row + 1):
        cat = ws_prod.cell(row, 1).value
        prod_name = ws_prod.cell(row, 2).value
        style_no = ws_prod.cell(row, 3).value
        color = ws_prod.cell(row, 4).value
        
        # KRW 가격 (컬럼 13)
        price_krw = ws_prod.cell(row, 13).value
        
        if prod_name and style_no and price_krw:
            try:
                price_krw = float(price_krw)
                products.append({
                    'category': cat,
                    'product_name': prod_name,
                    'style_no': style_no,
                    'color': color,
                    'price_krw': price_krw,
                })
            except:
                pass
    
    # 새 워크북 생성
    wb_out = Workbook()
    ws_out = wb_out.active
    ws_out.title = 'Country Pricing'
    
    # 스타일 정의
    thin = Side(style='thin')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # 헤더 설정
    headers = ['Category', 'Product Name', 'Style No.', 'Color', 'Base Price (KRW)']
    headers += [f"{c['code']} ({c['symbol']})" for c in country_list]
    
    for col, header in enumerate(headers, 1):
        cell = ws_out.cell(row=1, column=col, value=header)
        cell.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='006FC0')
        cell.alignment = center_align
        cell.border = border_all
    
    # 컬럼 너비
    ws_out.column_dimensions['A'].width = 15
    ws_out.column_dimensions['B'].width = 40
    ws_out.column_dimensions['C'].width = 16
    ws_out.column_dimensions['D'].width = 20
    ws_out.column_dimensions['E'].width = 18
    
    for col in range(6, 6 + len(country_list)):
        ws_out.column_dimensions[get_column_letter(col)].width = 16
    
    # 데이터 입력
    for row_idx, prod in enumerate(products, 2):
        # Category
        cell = ws_out.cell(row=row_idx, column=1, value=prod['category'])
        cell.font = Font(name='Arial', size=9)
        cell.alignment = left_align
        cell.border = border_all
        
        # Product Name
        cell = ws_out.cell(row=row_idx, column=2, value=prod['product_name'])
        cell.font = Font(name='Arial', size=9)
        cell.alignment = left_align
        cell.border = border_all
        
        # Style No.
        cell = ws_out.cell(row=row_idx, column=3, value=prod['style_no'])
        cell.font = Font(name='Arial', size=9)
        cell.alignment = center_align
        cell.border = border_all
        
        # Color
        cell = ws_out.cell(row=row_idx, column=4, value=prod['color'])
        cell.font = Font(name='Arial', size=9)
        cell.alignment = left_align
        cell.border = border_all
        
        # Base Price (KRW)
        cell = ws_out.cell(row=row_idx, column=5, value=int(prod['price_krw']))
        cell.font = Font(name='Arial', size=9)
        cell.alignment = center_align
        cell.border = border_all
        cell.number_format = '#,##0'
        
        # 국가별 가격 계산
        for col_idx, country in enumerate(country_list, 6):
            code = country['code']
            adjustment = country['adjustment']
            exchange = exchange_rates.get(code, 1.0)
            
            # 가격 = KRW 기본가 / 환율 * (1 + 조정%)
            price_local = (prod['price_krw'] / exchange) * (1 + adjustment)
            
            cell = ws_out.cell(row=row_idx, column=col_idx, value=price_local)
            cell.font = Font(name='Arial', size=9)
            cell.alignment = center_align
            cell.border = border_all
            cell.number_format = '#,##0.00'
    
    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf


# ── TAB 4: 국가별 가격 생성기
with tab4:
    st.caption("🌍 프로덕트 시트 + 환율 + 조정비율 → 국가별 가격 시트 자동 생성")
    st.divider()
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("📁 파일 업로드")
        pricing_product_file = st.file_uploader(
            "프로덕트 시트 Excel (.xlsx)",
            type=['xlsx'],
            key="pricing_product"
        )
        st.caption("📌 예: _SPECIALGUEST____NEO_DAWN___2026SS____PRODUCT_LIST.xlsx")
    
    with col2:
        st.subheader("⚙️ 기본 설정")
        selected_countries = st.multiselect(
            "국가 선택",
            ['USD', 'EUR', 'GBP', 'JPY', 'CNY', 'AED', 'CAD', 'AUD', 'NZD', 'CHF', 'SEK', 'NOK', 'DKK', 'HUF', 'PLN', 'CZK'],
            default=['USD', 'EUR', 'GBP', 'JPY'],
            key="select_countries"
        )
    
    st.divider()
    
    if pricing_product_file:
        st.subheader("🌐 환율 및 조정비율 입력")
        
        # 기본 환율 (KRW 기준)
        default_rates = {
            'USD': 1250,
            'EUR': 1350,
            'GBP': 1580,
            'JPY': 8.3,
            'CNY': 172,
            'AED': 340,
            'CAD': 920,
            'AUD': 810,
            'NZD': 780,
            'CHF': 1400,
            'SEK': 115,
            'NOK': 118,
            'DKK': 181,
            'HUF': 3.3,
            'PLN': 310,
            'CZK': 54,
        }
        
        # 기본 조정비율
        default_adjustments = {
            'USD': 68.3,
            'EUR': 0,
            'GBP': 54.0,
            'JPY': 39.7,
            'CNY': 39.7,
            'AED': 29.8,
            'CAD': 39.7,
            'AUD': 30.9,
            'NZD': 38.6,
            'CHF': 36.4,
            'SEK': 60.6,
            'NOK': 56.2,
            'DKK': 60.6,
            'HUF': 61.7,
            'PLN': 57.3,
            'CZK': 55.1,
        }
        
        currency_symbols = {
            'USD': '$', 'EUR': '€', 'GBP': '£', 'JPY': '¥',
            'CNY': '¥', 'AED': 'د.إ', 'CAD': 'C$', 'AUD': 'A$',
            'NZD': 'NZ$', 'CHF': 'CHF', 'SEK': 'kr', 'NOK': 'kr',
            'DKK': 'kr', 'HUF': 'Ft', 'PLN': 'zł', 'CZK': 'Kč'
        }
        
        rates = {}
        adjustments = {}
        
        # 선택된 국가만 입력폼 표시
        if selected_countries:
            cols = st.columns(len(selected_countries))
            
            for idx, curr in enumerate(selected_countries):
                with cols[idx]:
                    st.write(f"**{curr} {currency_symbols.get(curr, '')}**")
                    
                    rates[curr] = st.number_input(
                        f"{curr} 환율",
                        value=default_rates.get(curr, 1.0),
                        step=1.0,
                        key=f"rate_{curr}",
                        label_visibility="collapsed",
                        help="KRW 기준 환율 (예: 1 USD = 1250 KRW)"
                    )
                    
                    adj_pct = st.number_input(
                        f"{curr} 조정 %",
                        value=int(default_adjustments.get(curr, 0)),
                        step=1,
                        key=f"adj_{curr}",
                        label_visibility="collapsed",
                        help="조정 비율 (%)"
                    )
                    adjustments[curr] = adj_pct / 100
        else:
            st.warning("⚠️ 최소 1개 국가를 선택해주세요.")
        
        st.divider()
        
        if st.button("🔄 국가별 가격 시트 생성", use_container_width=True, type="primary"):
            if not selected_countries:
                st.error("❌ 국가를 선택해주세요.")
            else:
                with st.spinner("국가별 가격 계산 중..."):
                    try:
                        # 국가 정보 구성
                        country_data = []
                        for code in selected_countries:
                            country_data.append({
                                'code': code,
                                'name': code,
                                'symbol': currency_symbols.get(code, code),
                                'adjustment': adjustments.get(code, 0),
                            })
                        
                        # 가격 시트 생성
                        buf = make_country_pricing_sheet(
                            pricing_product_file,
                            country_data,
                            rates
                        )
                        
                        st.success("✅ 국가별 가격 시트 생성 완료!")
                        
                        date_str = datetime.now().strftime("%Y%m%d")
                        st.download_button(
                            "⬇️ 국가별 가격 시트 다운로드",
                            data=buf,
                            file_name=f"Country_Pricing_{date_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            type="primary"
                        )
                        
                    except Exception as e:
                        st.error(f"❌ 오류: {e}")
                        import traceback
                        st.code(traceback.format_exc())
    else:
        st.info("👆 프로덕트 시트 파일을 업로드해주세요.")

st.divider()
st.caption("💡 환율 및 조정비율은 필요에 따라 자유롭게 조정할 수 있습니다!")
