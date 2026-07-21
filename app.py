import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="올리브영 수주 자동화 시스템", layout="wide")
st.title("📦 올리브영 수주업로드 및 재고 매핑 시스템")

# ---------------------------------------------------------
# 1. 깃허브 저장소 마스터 파일 이름
# ---------------------------------------------------------
MASTER_FILE_NAME = "oliveyoung_master.xlsx"

@st.cache_data(ttl=60)
def load_master_data(file_path):
    """마스터 서식 파일에서 제품명 및 배송처 매핑 정보를 가져옵니다."""
    df_product = pd.read_excel(file_path, sheet_name='제품명')
    
    # '상품코드' 열 찾기 ('상품코드' 또는 '상품' 컬럼 지원)
    code_col = [c for c in df_product.columns if '코드' in str(c)]
    code_col_name = code_col[0] if code_col else '상품코드'
    
    # '상품명' 열 찾기
    name_col = [c for c in df_product.columns if '상품명' in str(c)]
    name_col_name = name_col[1] if len(name_col) > 1 else name_col[0] # 두 번째 상품명 컬럼이 텍스트 상품명인 경우가 많음
    
    df_clean = df_product.dropna(subset=[code_col_name]).copy()
    
    # 상품명 -> MEcode 매핑 딕셔너리
    mecode_map = {}
    for _, r in df_clean.iterrows():
        p_name = str(r[name_col_name]).strip()
        p_code = str(r[code_col_name]).strip()
        if p_name and p_code:
            mecode_map[p_name] = p_code
            
    df_delivery = pd.read_excel(file_path, sheet_name='배송처')
    return mecode_map, df_delivery

# ---------------------------------------------------------
# 2. 파일 연동 및 사이드바
# ---------------------------------------------------------
st.sidebar.header("📁 일일 작업 파일 업로드")

if os.path.exists(MASTER_FILE_NAME):
    st.sidebar.success(f"✅ 마스터 서식 연동 완료 (`{MASTER_FILE_NAME}`)")
    try:
        mecode_map, df_delivery_master = load_master_data(MASTER_FILE_NAME)
    except Exception as e:
        st.sidebar.error(f"마스터 파일 읽기 오류: {e}")
else:
    st.sidebar.error(f"❌ 깃허브에서 `{MASTER_FILE_NAME}` 파일을 찾을 수 없습니다.")

order_file = st.sidebar.file_uploader("1. 올리브영 발주서 (Raw DATA)", type=["xlsx"])
wms_file = st.sidebar.file_uploader("2. WMS 일일재고 파일", type=["xlsx"])

# ---------------------------------------------------------
# 3. 데이터 매핑 실행
# ---------------------------------------------------------
if order_file and wms_file and os.path.exists(MASTER_FILE_NAME):
    try:
        df_order = pd.read_excel(order_file)
        df_wms = pd.read_excel(wms_file)
        
        # WMS 헤더 정리 (이중 헤더 처리)
        if '순번' in df_wms.iloc[0].values or '상품' in df_wms.iloc[0].values:
            df_wms.columns = df_wms.iloc[0]
            df_wms = df_wms[1:].reset_index(drop=True)
            
        # WMS 내 상품코드 컬럼명 자동 파악 ('상품' 또는 '상품코드')
        wms_code_col = '상품' if '상품' in df_wms.columns else ('상품코드' if '상품코드' in df_wms.columns else None)
        
        results = []
        missing_products = [] # [제품명] 시트에 없는 신규 상품 모음

        for idx, row in df_order.iterrows():
            item_name = str(row.get('상품명', '')).strip()
            
            # 발주수량 가져오기 (컬럼명 유연 처리)
            order_qty = row.get('발주수량\n(EA)', row.get('발주수량', row.get('수량', 0)))
            try:
                order_qty = float(order_qty)
            except:
                order_qty = 0
            
            # [A] 깃허브 마스터 서식 [제품명] 시트에서 MEcode 매핑
            mecode = mecode_map.get(item_name, None)
            
            if not mecode:
                missing_products.append(item_name)
                status = "검토필요 (신규상품 - [제품명] 시트 미등록)"
                lot, expiry = "", ""
            else:
                # [B] WMS 재고 조회 (상품명 또는 MEcode/상품 컬럼 조건)
                cond_name = df_wms['상품명'].astype(str).str.strip() == item_name
                if wms_code_col:
                    cond_code = df_wms[wms_code_col].astype(str).str.strip() == mecode
                    wms_match = df_wms[cond_name | cond_code].copy()
                else:
                    wms_match = df_wms[cond_name].copy()
                
                # 유효일자 임박순 정렬 (FEFO)
                if '유효일자' in wms_match.columns:
                    wms_match = wms_match.sort_values(by='유효일자', ascending=True)
                
                # 재고 수량 계산 (합계수량/정상수량/EA 중 유효한 수량)
                stock_col = '정상수량' if '정상수량' in wms_match.columns else ('합계수량' if '합계수량' in wms_match.columns else None)
                
                if stock_col:
                    total_stock = pd.to_numeric(wms_match[stock_col], errors='coerce').sum()
                else:
                    total_stock = 0
                
                if total_stock >= order_qty and len(wms_match) > 0:
                    status = "정상"
                    lot_val = wms_match.iloc[0].get('화주LOT', '')
                    lot = str(lot_val) if pd.notna(lot_val) else ""
                    
                    exp_val = str(wms_match.iloc[0].get('유효일자', ''))
                    expiry = exp_val[:10].replace('-', '').replace('.', '')
                else:
                    status = "검토필요 (출고가능 재고없음)"
                    lot = ""
                    expiry = ""
            
            results.append({
                '입고예정일': str(row.get('입고예정일', ''))[:10].replace('-', '').replace('.', ''),
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

        # 4) 신규 상품 경고 및 매핑 결과 표시
        if missing_products:
            st.error(f"⚠️ [제품명] 시트에 없는 신규 상품이 {len(set(missing_products))}건 발견되었습니다!")
            st.warning("`oliveyoung_master.xlsx` 파일의 [제품명] 시트에 아래 상품을 추가해서 깃허브에 다시 올려주세요.")
            st.dataframe(pd.DataFrame({'미등록 신규 상품명': list(set(missing_products))}))

        st.subheader("📋 수주 업로드 최종 결과")
        st.dataframe(df_result, use_container_width=True)

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("왼쪽 사이드바에서 발주서와 WMS 일일재고 엑셀 파일 2개를 올려주세요.")
