import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="올리브영 수주 자동화 시스템", layout="wide")
st.title("📦 올리브영 수주업로드 및 재고 매핑 시스템")

# ---------------------------------------------------------
# 1. 깃허브 저장소에 올라가 있는 고정 서식 파일 이름
# ---------------------------------------------------------
MASTER_FILE_NAME = "oliveyoung_master.xlsx"

@st.cache_data(ttl=60) # 깃허브에서 파일이 변경되면 60초 후 자동으로 새 데이터를 반영합니다.
def load_master_data(file_path):
    """깃허브에 등록된 마스터 서식파일에서 제품명 및 배송처 정보를 읽어옵니다."""
    # 1) [제품명] 시트 로드
    df_product = pd.read_excel(file_path, sheet_name='제품명')
    
    # '상품코드' 열이 있는 행만 가져오기
    df_product_clean = df_product.dropna(subset=['상품코드']).copy()
    
    # 상품명과 상품코드(MEcode) 앞뒤 공백 제거 후 딕셔너리로 변환
    mecode_map = dict(zip(
        df_product_clean['상품명'].astype(str).str.strip(), 
        df_product_clean['상품코드'].astype(str).str.strip()
    ))
    
    # 2) [배송처] 시트 로드
    df_delivery = pd.read_excel(file_path, sheet_name='배송처')
    
    return mecode_map, df_delivery

# ---------------------------------------------------------
# 2. 마스터 파일 연동 확인 및 일일 작업 파일 업로드
# ---------------------------------------------------------
st.sidebar.header("📁 작업 파일 업로드")

# 깃허브 서버에 파일이 잘 올라가 있는지 확인
if os.path.exists(MASTER_FILE_NAME):
    st.sidebar.success(f"✅ 마스터 서식 연동 완료 (`{MASTER_FILE_NAME}`)")
    mecode_map, df_delivery_master = load_master_data(MASTER_FILE_NAME)
else:
    st.sidebar.error(f"❌ 깃허브에서 `{MASTER_FILE_NAME}` 파일을 찾을 수 없습니다.")
    st.sidebar.info("깃허브 저장소에 `oliveyoung_master.xlsx` 파일을 업로드해 주세요.")

# 작업 시 매번 올려줄 파일 2개
order_file = st.sidebar.file_uploader("1. 올리브영 발주서 (Raw DATA)", type=["xlsx"])
wms_file = st.sidebar.file_uploader("2. WMS 일일재고 파일", type=["xlsx"])

# ---------------------------------------------------------
# 3. 데이터 매핑 로직 실행
# ---------------------------------------------------------
if order_file and wms_file and os.path.exists(MASTER_FILE_NAME):
    try:
        # 발주서 및 WMS 재고 로드
        df_order = pd.read_excel(order_file)
        df_wms = pd.read_excel(wms_file)
        
        # WMS 헤더 정리
        if '상품명' not in df_wms.columns:
            df_wms.columns = df_wms.iloc[0]
            df_wms = df_wms[1:].reset_index(drop=True)

        results = []
        missing_products = [] # [제품명] 시트에 없는 신규 상품 저장 목록

        for idx, row in df_order.iterrows():
            item_name = str(row.get('상품명', '')).strip()
            order_qty = row.get('발주수량\n(EA)', row.get('발주수량', 0))
            
            # [A] 깃허브 마스터 서식 [제품명] 시트에서 MEcode 끌어오기
            mecode = mecode_map.get(item_name, None)
            
            if not mecode:
                # [제품명] 시트에 상품이 없는 경우
                missing_products.append(item_name)
                status = "검토필요 (신규상품 - [제품명] 시트 미등록)"
                lot, expiry = "", ""
            else:
                # [B] WMS 재고 조회 및 유효일자(FEFO) 기준 매핑
                wms_match = df_wms[
                    (df_wms['상품명'].astype(str).str.strip() == item_name) | 
                    (df_wms['상품코드'].astype(str).str.strip() == mecode)
                ]
                
                # 유효일자 임박순 정렬
                wms_match = wms_match.sort_values(by='유효일자', ascending=True)
                
                total_stock = pd.to_numeric(wms_match['정상수량'], errors='coerce').sum() if '정상수량' in wms_match.columns else 0
                
                if total_stock >= order_qty and len(wms_match) > 0:
                    status = "정상"
                    lot = wms_match.iloc[0]['화주LOT']
                    expiry = str(wms_match.iloc[0]['유효일자'])[:10].replace('-', '')
                else:
                    status = "검토필요 (출고가능 재고없음)"
                    lot = ""
                    expiry = ""
            
            results.append({
                '입고예정일': str(row.get('입고예정일', ''))[:10].replace('-', ''),
                '발주처코드': '86100000',
                'MEcode': mecode if mecode else "미등록",
                '상품명': item_name,
                '수량': order_qty,
                '단가': row.get('원단가', 0),
                '발주금액': row.get('원가금액', 0),
                'LOT': lot,
                '유효일자': expiry,
                '매핑상태': status
            })

        df_result = pd.DataFrame(results)

        # 4) 신규 상품 경고 및 매핑 결과 출력
        if missing_products:
            st.error(f"⚠️ [제품명] 시트에 없는 신규 상품이 {len(set(missing_products))}건 발견되었습니다!")
            st.warning("깃허브(GitHub)에 올려둔 `oliveyoung_master.xlsx` 파일의 [제품명] 시트에 아래 상품을 추가하고 새로 덮어씌워 업로드해주세요.")
            st.dataframe(pd.DataFrame({'미등록 신규 상품명': list(set(missing_products))}))

        st.subheader("📋 수주 업로드 최종 결과")
        st.dataframe(df_result, use_container_width=True)

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("왼쪽 사이드바에서 발주서와 WMS 일일재고 엑셀 파일 2개를 올려주세요.")
