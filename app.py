import streamlit as st
import pandas as pd
import os

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

@st.cache_data(ttl=60)
def load_master_data(file_path):
    """마스터 파일의 [제품명] 시트에서 (상품바코드 ➔ MEcode) 및 (상품명 ➔ MEcode) 맵을 생성합니다."""
    df_product = pd.read_excel(file_path, sheet_name='제품명')
    
    # 컬럼명 자동 파악 ('상품바코드', '상품코드'(MEcode), '상품명')
    barcode_col = '상품바코드' if '상품바코드' in df_product.columns else df_product.columns[0]
    mecode_col = '상품코드' if '상품코드' in df_product.columns else df_product.columns[2]
    name_col = '상품명' if '상품명' in df_product.columns else df_product.columns[1]

    barcode_to_mecode = {}
    name_to_mecode = {}

    for _, row in df_product.iterrows():
        b_code = clean_barcode(row.get(barcode_col, ''))
        m_code = str(row.get(mecode_col, '')).strip()
        p_name = str(row.get(name_col, '')).strip()

        if m_code and m_code.lower() != 'nan':
            if b_code:
                barcode_to_mecode[b_code] = m_code
            if p_name:
                name_to_mecode[p_name] = m_code

    return barcode_to_mecode, name_to_mecode

# ---------------------------------------------------------
# 2. 사이드바 및 파일 업로드
# ---------------------------------------------------------
st.sidebar.header("📁 일일 작업 파일 업로드")

if os.path.exists(MASTER_FILE_NAME):
    st.sidebar.success(f"✅ 마스터 서식 연동 완료 (`{MASTER_FILE_NAME}`)")
    try:
        barcode_to_mecode, name_to_mecode = load_master_data(MASTER_FILE_NAME)
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
            item_barcode = clean_barcode(row.get('상품코드', '')) # 납품확인서의 상품코드 = 13자리 바코드
            
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
                    lot = str(lot_val) if pd.notna(lot_val) and str(lot_val).lower() != 'nan' else ""
                    
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

        if unmapped_items:
            st.error(f"⚠️ [제품명] 시트에 미등록된 바코드/상품이 {len(unmapped_items)}건 발견되었습니다.")
            st.warning("`oliveyoung_master.xlsx` 의 [제품명] 시트에 바코드와 MEcode(상품코드)를 등록해 주세요.")
            st.dataframe(pd.DataFrame(unmapped_items))

        st.subheader("📋 수주 업로드 최종 결과")
        st.dataframe(df_result, use_container_width=True)

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("왼쪽 사이드바에서 납품확인서 목록 파일과 WMS 일일재고 엑셀 파일을 업로드해 주세요.")
