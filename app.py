import streamlit as st
import pandas as pd
import numpy as np
import os
import io
from datetime import datetime, timedelta

st.set_page_config(page_title="올리브영 수주 자동화 시스템", layout="wide")
st.title("📦 올리브영 납품확인서 수주업로드 및 재고 매핑 시스템")

# ---------------------------------------------------------
# 1. 마스터 파일 연동 설정
# ---------------------------------------------------------
MASTER_FILE_NAME = "oliveyoung_master.xlsx"

def clean_barcode(val):
    """바코드 소수점(.0) 및 공백 정제"""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def is_valid_text(val):
    """NaN, None, 빈 값 체크"""
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s != "" and s != "nan" and s != "none"

@st.cache_data(ttl=60)
def load_master_data(file_path):
    """마스터 파일의 [제품명] 및 [배송처] 시트 정보 읽기"""
    # 1) [제품명] 시트 (바코드 ➔ MEcode, 상품명 ➔ MEcode)
    df_product = pd.read_excel(file_path, sheet_name='제품명')
    barcode_col = '상품바코드' if '상품바코드' in df_product.columns else df_product.columns[0]
    mecode_col = '상품코드' if '상품코드' in df_product.columns else df_product.columns[2]
    name_col = '상품명' if '상품명' in df_product.columns else df_product.columns[1]

    barcode_to_mecode = {}
    name_to_mecode = {}

    for _, row in df_product.iterrows():
        b_code = clean_barcode(row.get(barcode_col, ''))
        m_code = str(row.get(mecode_col, '')).strip()
        p_name = str(row.get(name_col, '')).strip()

        if is_valid_text(m_code):
            if b_code:
                barcode_to_mecode[b_code] = m_code
            if is_valid_text(p_name):
                name_to_mecode[p_name] = m_code

    # 2) [배송처] 시트 (센터명 ➔ 배송코드)
    df_delivery = pd.read_excel(file_path, sheet_name='배송처')
    deliv_name_col = '배송처' if '배송처' in df_delivery.columns else df_delivery.columns[0]
    deliv_code_col = '배송코드' if '배송코드' in df_delivery.columns else df_delivery.columns[1]

    delivery_map = {}
    for _, row in df_delivery.dropna(subset=[deliv_name_col, deliv_code_col]).iterrows():
        d_name = str(row[deliv_name_col]).strip()
        d_code = str(row[deliv_code_col]).strip()
        if d_code.endswith('.0'):
            d_code = d_code[:-2]
        delivery_map[d_name] = d_code

    return barcode_to_mecode, name_to_mecode, delivery_map

# ---------------------------------------------------------
# 2. 사이드바 및 파일 업로드
# ---------------------------------------------------------
st.sidebar.header("📁 일일 작업 파일 업로드")

if os.path.exists(MASTER_FILE_NAME):
    st.sidebar.success(f"✅ 마스터 서식 연동 완료 (`{MASTER_FILE_NAME}`)")
    try:
        barcode_to_mecode, name_to_mecode, delivery_map = load_master_data(MASTER_FILE_NAME)
    except Exception as e:
        st.sidebar.error(f"마스터 파일 읽기 오류: {e}")
else:
    st.sidebar.error(f"❌ 깃허브에서 `{MASTER_FILE_NAME}` 파일을 찾을 수 없습니다.")

order_file = st.sidebar.file_uploader("1. 납품확인서 목록 파일", type=["xlsx"])
wms_file = st.sidebar.file_uploader("2. WMS 일일재고 파일", type=["xlsx"])

# ---------------------------------------------------------
# 3. 데이터 수주업로드 및 재고 매핑 처리
# ---------------------------------------------------------
if order_file and wms_file and os.path.exists(MASTER_FILE_NAME):
    try:
        df_order = pd.read_excel(order_file)
        df_wms_raw = pd.read_excel(wms_file, header=None)

        # WMS 파일 헤더 및 파싱 (2번째 행부터 데이터 시작 구조 대응)
        df_wms = df_wms_raw.iloc[2:].copy() if len(df_wms_raw) > 2 else df_wms_raw.copy()
        
        # WMS 필요한 필드 동적 생성 (D열: ME코드, G열: LOT, N열: 유효일자, R열: 입수량, AF열: 수량, AH열: 바코드)
        wms_stock_data = pd.DataFrame({
            'ME코드': df_wms[3].astype(str).str.strip() if 3 in df_wms.columns else '',
            '상품명': df_wms[4].astype(str).str.strip() if 4 in df_wms.columns else '',
            '화주LOT': df_wms[6].astype(str).str.strip() if 6 in df_wms.columns else '',
            '유효일자': pd.to_datetime(df_wms[13], errors='coerce') if 13 in df_wms.columns else pd.NaT,
            '입수량_BOX': pd.to_numeric(df_wms[17], errors='coerce').fillna(1) if 17 in df_wms.columns else 1,
            '합계수량': pd.to_numeric(df_wms[31], errors='coerce').fillna(0) if 31 in df_wms.columns else 0,
            '상품바코드': df_wms[33].astype(str).str.replace('.0', '', regex=False).str.strip() if 33 in df_wms.columns else ''
        })

        results = []
        unmapped_items = []
        today = datetime.now()

        for idx, row in df_order.iterrows():
            item_name = str(row.get('상품명', '')).strip()
            
            # 🚫 상품명이 nan, None인 경우 제외
            if not is_valid_text(item_name):
                continue

            item_barcode = clean_barcode(row.get('상품코드', '')) # I열 (13자리 바코드)
            center_name = str(row.get('센터', '')).strip()       # H열 (센터명)
            delivery_code = delivery_map.get(center_name, "미등록배송처")

            # 📌 N열: 발주수량(EA)
            raw_qty = row.get('발주수량\n(EA)', row.get('발주수량', 0))
            try:
                order_qty = int(float(raw_qty))
            except:
                order_qty = 0

            # 📌 T열: 원단가
            raw_price = row.get('원단가', 0)
            try:
                unit_price = int(float(raw_price))
            except:
                unit_price = 0

            # 📌 U열: 원가금액 (발주금액)
            raw_amount = row.get('원가금액', 0)
            try:
                total_amount = int(float(raw_amount))
            except:
                total_amount = unit_price * order_qty

            # 입고예정일 정제
            raw_arr_date = row.get('입고예정일', today)
            try:
                order_date = pd.to_datetime(raw_arr_date)
                arr_date_str = order_date.strftime('%Y%m%d')
            except:
                order_date = today
                arr_date_str = today.strftime('%Y%m%d')

            # 🔍 MEcode 찾기 (1차: 바코드 매칭, 2차: 상품명 매칭)
            mecode = barcode_to_mecode.get(item_barcode, None)
            if not mecode:
                mecode = name_to_mecode.get(item_name, None)

            if not mecode:
                unmapped_items.append({'바코드': item_barcode, '상품명': item_name})
                status = "검토필요 (신규상품 - 마스터 미등록)"
                selected_lot, selected_exp = "", ""
            else:
                # 🏢 WMS 재고 파악 (올리브영 출고 조건 적용)
                sub_stock = wms_stock_data[
                    (wms_stock_data['ME코드'] == mecode) | 
                    (wms_stock_data['상품바코드'] == item_barcode) |
                    (wms_stock_data['상품명'] == item_name)
                ].copy()

                # 조건 1: 유효일자 1년 6개월(547일) 이상 남은 재고만 필터링
                min_valid_date = order_date + timedelta(days=547)
                valid_stock = sub_stock[sub_stock['유효일자'] >= min_valid_date].copy()

                # 조건 2: 박스 입수량 미만 재고(단수 재고) 차단
                valid_stock = valid_stock[valid_stock['합계수량'] >= valid_stock['입수량_BOX']]

                # FEFO (유효일자 임박순 정렬)
                valid_stock = valid_stock.sort_values(by='유효일자', ascending=True)

                selected_lot = ""
                selected_exp = ""
                status = "정상"

                if not valid_stock.empty:
                    # 조건 3: 단일 LOT로 발주수량을 전량 출고할 수 있는지 검증
                    single_lot_match = valid_stock[valid_stock['합계수량'] >= order_qty]

                    if not single_lot_match.empty:
                        best_match = single_lot_match.iloc[0]
                        selected_lot = str(best_match['화주LOT'])
                        if pd.notnull(best_match['유효일자']):
                            selected_exp = best_match['유효일자'].strftime('%Y%m%d')
                    else:
                        # 총 재고는 있으나 1개 LOT로 불가능할 때
                        total_avail_qty = valid_stock['합계수량'].sum()
                        if total_avail_qty >= order_qty:
                            status = "검토필요 (LOT 분할 필요)"
                        else:
                            status = "검토필요 (유효재고 부족)"
                else:
                    status = "검토필요 (출고가능 재고없음)"

            results.append({
                '입고예정일': arr_date_str,
                '발주처코드': '86100000',
                '배송코드': delivery_code,
                'MEcode': mecode if mecode else "미등록",
                '상품명': item_name,
                '수량': order_qty,
                '단가': unit_price,
                '발주금액': total_amount,
                'LOT': selected_lot,
                '유효일자': selected_exp,
                '매핑상태': status
            })

        df_result = pd.DataFrame(results)

        # ---------------------------------------------------------
        # 4. 결과 출력 및 스타일링 (정상: 초록, 검토필요: 빨강)
        # ---------------------------------------------------------
        if unmapped_items:
            st.error(f"⚠️ [제품명] 시트에 미등록된 상품 {len(unmapped_items)}건이 발견되었습니다.")
            st.dataframe(pd.DataFrame(unmapped_items))

        st.subheader("📋 수주 업로드 최종 결과")

        # 스타일링 함수
        def style_rows(row):
            status = str(row['매핑상태'])
            if '정상' in status:
                return ['background-color: #e6ffe6; color: #008000; font-weight: bold;'] * len(row)
            else:
                return ['background-color: #ffe6e6; color: #cc0000; font-weight: bold;'] * len(row)

        styled_df = df_result.style.apply(style_rows, axis=1)
        st.dataframe(styled_df, use_container_width=True)

        # ---------------------------------------------------------
        # 5. 엑셀 다운로드 버튼
        # ---------------------------------------------------------
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
        excel_data = excel_buffer.getvalue()

        st.download_button(
            label="📥 최종 수주업로드 결과 엑셀 다운로드",
            data=excel_data,
            file_name=f"올리브영_수주업로드_완료_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("왼쪽 사이드바에서 '납품확인서 목록'과 'WMS 일일재고' 엑셀 파일 2개를 올려주세요.")
