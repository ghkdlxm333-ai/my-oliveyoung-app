import streamlit as st
import pandas as pd
import os
import io

st.set_page_config(page_title="올리브영 수주 자동화 시스템", layout="wide")
st.title("📦 올리브영 납품확인서 수주업로드 및 재고 매핑 시스템")

# ---------------------------------------------------------
# 1. 깃허브 저장소 마스터 파일 설정
# ---------------------------------------------------------
MASTER_FILE_NAME = "oliveyoung_master.xlsx"

def clean_barcode(val):
    """바코드에서 소수점(.0) 및 공백을 제거하여 순수 문자열 바코드로 정제"""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def is_valid_text(val):
    """NaN, None, 빈 값 체크 함수"""
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s != "" and s != "nan" and s != "none"

@st.cache_data(ttl=60)
def load_master_data(file_path):
    """마스터 파일에서 [제품명] 시트(바코드, MEcode) 및 [배송처] 시트(배송처, 배송코드) 정보를 읽어옵니다."""
    # 1) [제품명] 시트 매핑
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

    # 2) [배송처] 시트 매핑 (센터명 ➔ 배송코드)
    df_delivery = pd.read_excel(file_path, sheet_name='배송처')
    delivery_map = {}
    
    deliv_name_col = '배송처' if '배송처' in df_delivery.columns else df_delivery.columns[0]
    deliv_code_col = '배송코드' if '배송코드' in df_delivery.columns else df_delivery.columns[1]

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
        df_wms = pd.read_excel(wms_file)

        # WMS 이중 헤더 정리
        if '순번' in df_wms.iloc[0].values or '상품' in df_wms.iloc[0].values:
            df_wms.columns = df_wms.iloc[0]
            df_wms = df_wms[1:].reset_index(drop=True)

        wms_code_col = '상품' if '상품' in df_wms.columns else ('상품코드' if '상품코드' in df_wms.columns else None)

        results = []
        unmapped_items = []

        for idx, row in df_order.iterrows():
            item_name = str(row.get('상품명', '')).strip()
            
            # 🚫 상품명이 nan, None, 빈 값인 유효하지 않은 행은 결과 목록에서 제외
            if not is_valid_text(item_name):
                continue

            item_barcode = clean_barcode(row.get('상품코드', '')) # 납품확인서의 상품코드 = 13자리 바코드
            center_name = str(row.get('센터', '')).strip()       # 납품확인서의 센터 (H열)
            
            # 🚚 배송코드 매칭 ([배송처] 시트 참조)
            delivery_code = delivery_map.get(center_name, "미등록배송처")

            # 발주수량 (EA)
            order_qty = row.get('발주수량\n(EA)', row.get('발주수량', row.get('수량', 0)))
            try:
                order_qty = float(order_qty)
            except:
                order_qty = 0

            # 🔍 1차: 상품바코드 ➔ MEcode 매칭
            mecode = barcode_to_mecode.get(item_barcode, None)
            
            # 🔍 2차 (Fallback): 상품바코드로 안 찾아지면 상품명으로 매칭
            if not mecode:
                mecode = name_to_mecode.get(item_name, None)

            # 매핑 실패 처리
            if not mecode:
                unmapped_items.append({'바코드': item_barcode, '상품명': item_name})
                status = "검토필요 (신규상품 - [제품명] 시트 미등록)"
                lot, expiry = "", ""
            else:
                # WMS에서 해당 MEcode 재고 조회
                if wms_code_col:
                    wms_match = df_wms[df_wms[wms_code_col].astype(str).str.strip() == mecode].copy()
                else:
                    wms_match = df_wms[df_wms['상품명'].astype(str).str.strip() == item_name].copy()

                # FEFO 정렬 (유효일자 임박순)
                if '유효일자' in wms_match.columns:
                    wms_match = wms_match.sort_values(by='유효일자', ascending=True)

                stock_col = '정상수량' if '정상수량' in wms_match.columns else ('합계수량' if '합계수량' in wms_match.columns else None)
                total_stock = pd.to_numeric(wms_match[stock_col], errors='coerce').sum() if stock_col else 0

                if total_stock >= order_qty and len(wms_match) > 0:
                    status = "정상"
                    lot_val = wms_match.iloc[0].get('화주LOT', '')
                    lot = str(lot_val) if pd.notna(lot_val) and is_valid_text(lot_val) else ""
                    
                    exp_val = str(wms_match.iloc[0].get('유효일자', ''))
                    expiry = exp_val[:10].replace('-', '').replace('.', '')
                else:
                    status = "검토필요 (출고가능 재고없음)"
                    lot = ""
                    expiry = ""

            # 입고예정일 날짜 형태 정제 (YYYYMMDD)
            arr_date = str(row.get('입고예정일', ''))[:10].replace('-', '').replace('.', '')

            results.append({
                '입고예정일': arr_date,
                '발주처코드': '86100000',
                '배송코드': delivery_code,
                'MEcode': mecode if mecode else "미등록",
                '상품명': item_name,
                '수량': order_qty,
                '단가': row.get('원단가', row.get('단가', 0)),
                '발주금액': row.get('원가금액', row.get('발주금액', 0)),
                'LOT': lot,
                '유효일자': expiry,
                '매핑상태': status
            })

        df_result = pd.DataFrame(results)

        # 미등록 신규 상품 경고 출력
        if unmapped_items:
            st.error(f"⚠️ [제품명] 시트에 미등록된 바코드/상품이 {len(unmapped_items)}건 발견되었습니다.")
            st.warning("`oliveyoung_master.xlsx` 의 [제품명] 시트에 바코드와 MEcode(상품코드)를 등록해 주세요.")
            st.dataframe(pd.DataFrame(unmapped_items))

        st.subheader("📋 수주 업로드 최종 결과")
        st.dataframe(df_result, use_container_width=True)

        # ---------------------------------------------------------
        # 4. 엑셀 다운로드 버튼 생성
        # ---------------------------------------------------------
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='Sheet1')
        excel_data = excel_buffer.getvalue()

        st.download_button(
            label="📥 최종 수주업로드 결과 엑셀 다운로드",
            data=excel_data,
            file_name="올리브영_수주업로드_완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("왼쪽 사이드바에서 납품확인서 목록 파일과 WMS 일일재고 엑셀 파일을 업로드해 주세요.")
