import streamlit as st
import pandas as pd
import numpy as np
import os
import io
from datetime import datetime, timedelta

st.set_page_config(page_title="올리브영 3PL WMS 수주 자동화", layout="wide")

# ---------------------------------------------------------
# 1. 마스터 파일 연동 설정
# ---------------------------------------------------------
MASTER_FILE_NAME = "oliveyoung_master.xlsx"

def clean_barcode(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def is_valid_text(val):
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s != "" and s != "nan" and s != "none"

@st.cache_data(ttl=60)
def load_master_data(file_path):
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
# 2. 사이드바 구성
# ---------------------------------------------------------
st.sidebar.title("📦 수주 자동화 메뉴")

if os.path.exists(MASTER_FILE_NAME):
    st.sidebar.success(f"✅ 마스터 연동 완료 (`{MASTER_FILE_NAME}`)")
    try:
        barcode_to_mecode, name_to_mecode, delivery_map = load_master_data(MASTER_FILE_NAME)
    except Exception as e:
        st.sidebar.error(f"마스터 읽기 오류: {e}")
else:
    st.sidebar.error(f"❌ `{MASTER_FILE_NAME}` 파일 없음")

order_file = st.sidebar.file_uploader("1. 납품확인서 목록 파일", type=["xlsx"])
wms_file = st.sidebar.file_uploader("2. WMS 일일재고 파일", type=["xlsx"])

# ---------------------------------------------------------
# 3. 데이터 수주업로드 및 재고 매핑 처리
# ---------------------------------------------------------
if order_file and wms_file and os.path.exists(MASTER_FILE_NAME):
    try:
        df_order = pd.read_excel(order_file)
        df_wms_raw = pd.read_excel(wms_file, header=None)
        df_wms = df_wms_raw.iloc[2:].copy() if len(df_wms_raw) > 2 else df_wms_raw.copy()
        
        wms_stock_data = pd.DataFrame({
            'ME코드': df_wms[3].astype(str).str.strip() if 3 in df_wms.columns else '',
            '상품명': df_wms[4].astype(str).str.strip() if 4 in df_wms.columns else '',
            '화주LOT': df_wms[6].astype(str).str.strip() if 6 in df_wms.columns else '',
            '유효일자': pd.to_datetime(df_wms[13], errors='coerce') if 13 in df_wms.columns else pd.NaT,
            '입수량_BOX': pd.to_numeric(df_wms[17], errors='coerce').fillna(1) if 17 in df_wms.columns else 1,
            '합계수량': pd.to_numeric(df_wms[31], errors='coerce').fillna(0) if 31 in df_wms.columns else 0,
            '상품바코드': df_wms[33].astype(str).str.replace('.0', '', regex=False).str.strip() if 33 in df_wms.columns else ''
        })

        wms_upload_list = []
        today = datetime.now()

        for idx, row in df_order.iterrows():
            item_name = str(row.get('상품명', '')).strip()
            if not is_valid_text(item_name):
                continue

            item_barcode = clean_barcode(row.get('상품코드', ''))
            center_name = str(row.get('센터', '')).strip()
            delivery_code = delivery_map.get(center_name, "미등록배송처")

            # 수주일자 / 납품일자 (입고예정일)
            raw_order_date = row.get('발주일자', today)
            try:
                order_date_str = pd.to_datetime(raw_order_date).strftime('%Y-%m-%d')
            except:
                order_date_str = today.strftime('%Y-%m-%d')

            raw_arr_date = row.get('입고예정일', today)
            try:
                arr_dt = pd.to_datetime(raw_arr_date)
                arr_date_str = arr_dt.strftime('%Y-%m-%d')
            except:
                arr_dt = today
                arr_date_str = today.strftime('%Y-%m-%d')

            # 수량/단가/합계/부가세(10%)
            try:
                order_qty = int(float(row.get('발주수량\n(EA)', row.get('발주수량', 0))))
            except:
                order_qty = 0

            try:
                unit_price = int(float(row.get('원단가', 0)))
            except:
                unit_price = 0

            try:
                total_amount = int(float(row.get('원가금액', 0)))
            except:
                total_amount = unit_price * order_qty

            vat_amount = int(total_amount * 0.1)

            # 상품코드(MEcode) 매핑
            mecode = barcode_to_mecode.get(item_barcode, None)
            if not mecode:
                mecode = name_to_mecode.get(item_name, None)

            selected_lot = ""
            selected_exp = ""
            box_pack = 1  # 기본 입수량
            status = "정상"

            if not mecode:
                status = "마스터미등록"
            else:
                sub_stock = wms_stock_data[
                    (wms_stock_data['ME코드'] == mecode) | 
                    (wms_stock_data['상품바코드'] == item_barcode) |
                    (wms_stock_data['상품명'] == item_name)
                ].copy()

                if not sub_stock.empty:
                    box_pack = int(sub_stock.iloc[0]['입수량_BOX'])

                # 조건 1: 유효일자 1년 6개월(547일) 이상 남은 재고
                min_valid_date = arr_dt + timedelta(days=547)
                valid_stock = sub_stock[sub_stock['유효일자'] >= min_valid_date].copy()

                # 조건 2: 박스 입수량 이상 재고 (입수부족 체크)
                valid_stock = valid_stock[valid_stock['합계수량'] >= valid_stock['입수량_BOX']]
                valid_stock = valid_stock.sort_values(by='유효일자', ascending=True)

                if not sub_stock.empty and valid_stock.empty:
                    # 전체 재고는 있으나 유효일자 1.5년 미달 또는 입수량 부족인 경우
                    status = "유효일자 미달"
                elif valid_stock.empty:
                    status = "재고부족"
                else:
                    # 조건 3: 단일 LOT 충족 여부
                    single_lot_match = valid_stock[valid_stock['합계수량'] >= order_qty]
                    if not single_lot_match.empty:
                        best_match = single_lot_match.iloc[0]
                        selected_lot = str(best_match['화주LOT'])
                        if pd.notnull(best_match['유효일자']):
                            selected_exp = best_match['유효일자'].strftime('%Y-%m-%d')
                    else:
                        if valid_stock['합계수량'].sum() >= order_qty:
                            status = "LOT분할필요"
                        else:
                            status = "재고부족"

            wms_upload_list.append({
                '출고구분': 0,
                '수주일자': order_date_str,
                '납품일자': arr_date_str,
                '발주처코드': '86100000',
                '발주처': 'CJ올리브영',
                '배송코드': delivery_code,
                '배송지': center_name,
                '상품코드': mecode if mecode else "미등록",
                '상품명': item_name,
                '입수량': box_pack,
                '수량': order_qty,
                '단가': unit_price,
                '합계': total_amount,
                '부가세': vat_amount,
                'LOT': selected_lot,
                '유효일자': selected_exp,
                '매핑상태': status
            })

        df_result = pd.DataFrame(wms_upload_list)

        # ---------------------------------------------------------
        # 4. 화면 대시보드 UI 구성
        # ---------------------------------------------------------
        st.title("📦 3PL WMS 수주 업로드 변환 센터")

        total_cnt = len(df_result)
        normal_cnt = len(df_result[df_result['매핑상태'] == '정상'])
        check_cnt = total_cnt - normal_cnt

        col1, col2, col3, col4 = st.columns([1, 1, 1, 2])
        col1.metric("총 처리 건수", f"{total_cnt} 건")
        col2.metric("✅ 자동 정상 매핑", f"{normal_cnt} 건")
        col3.metric("⚠️ 검토 필요 건수", f"{check_cnt} 건", delta_color="inverse")

        # 3PL WMS 다운로드 버퍼 생성
        wms_pure_cols = ['출고구분', '수주일자', '납품일자', '발주처코드', '발주처', '배송코드', '배송지', '상품코드', '상품명', '수량', '단가', '합계', '부가세', 'LOT', '유효일자']
        df_wms_pure = df_result[wms_pure_cols].copy()

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_wms_pure.to_excel(writer, index=False, sheet_name='WMS업로드_최종')
            df_result.to_excel(writer, index=False, sheet_name='전체상세검토')
        
        col4.download_button(
            label="📥 WMS 업로드용 엑셀 즉시 다운로드",
            data=excel_buffer.getvalue(),
            file_name=f"3PL_WMS_수주업로드_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        st.markdown("---")

        # ---------------------------------------------------------
        # 5. 탭 구성
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["🔍 전체 데이터 확인", "📋 3PL WMS 복사용 양식"])

        # 📌 [탭 1] 전체 데이터 확인 (단가/합계/부가세 삭제, 입수량 추가)
        with tab1:
            st.caption("📌 핵심 처리 내역입니다. (검토필요/오류 항목은 빨간색으로 강조 표시됩니다)")
            
            show_cols = ['납품일자', '배송코드', '배송지', '상품코드', '상품명', '입수량', '수량', 'LOT', '유효일자', '매핑상태']
            df_show = df_result[show_cols].copy()

            # 매핑상태에 따른 빨간색 하이라이트 함수
            def highlight_status(row):
                if str(row['매핑상태']) != '정상':
                    return ['background-color: #f8d7da; color: #dc3545; font-weight: bold;'] * len(row)
                return [''] * len(row)

            styled_show = df_show.style.apply(highlight_status, axis=1)
            st.dataframe(styled_show, height=500, use_container_width=True, hide_index=True)

        # 📌 [탭 2] 3PL WMS 복사용 양식
        with tab2:
            st.caption("📌 **3PL WMS 시스템 입력용 표준 양식입니다. 복사(`Ctrl+C`) 시 순수 데이터만 깔끔하게 복사됩니다.**")
            st.dataframe(df_wms_pure, height=500, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👈 왼쪽 사이드바에서 납품확인서와 WMS 일일재고 엑셀 파일을 업로드해 주세요.")
