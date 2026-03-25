import streamlit as st
import pandas as pd
import io
import re

# 1. 페이지 기본 설정 (가장 위에 와야 함)
st.set_page_config(page_title="올리브영 수주업로드 자동 입력 시스템", page_icon="🌿", layout="wide")

# ==========================================
# 🎨 사이드바 (Sidebar) 디자인: 로고 및 업로드 창
# ==========================================
with st.sidebar:
    # 인터넷에 있는 멘소래담 로고 이미지 주소 (바로 렌더링됨)
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    
    st.markdown("---")
    st.header("⚙️ 작업 설정")
    uploaded_file = st.file_uploader("올리브영 발주 엑셀 업로드", type=['xlsx'])
    
    st.markdown("---")
    st.caption("💡 자동 부분 할당 및 재고 차감 적용")
    st.caption("Developed by Jay")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("올리브영 수주업로드 자동 입력 시스템")
st.markdown("Mentholatum : Moving The Heart")

# 파일이 업로드 되었을 때만 아래 로직 실행
if uploaded_file:
    try:
        # 데이터 불러오기
        df_order = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv = pd.read_excel(uploaded_file, sheet_name='재고', header=2)

        # 불필요한 열 날리기
        if '잔여일수' in df_order.columns:
            start_idx = list(df_order.columns).index('잔여일수')
            cols_to_drop = df_order.columns[start_idx:]
            df_order = df_order.drop(columns=cols_to_drop)

        if 'LOT' not in df_order.columns: df_order['LOT'] = ''
        if '유효일자' not in df_order.columns: df_order['유효일자'] = ''

        # 데이터 정제
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.replace(r'[^a-zA-Z0-9]', '', regex=True).str.upper()

        df_order['수량'] = pd.to_numeric(df_order['수량'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
        df_inv['환산'] = pd.to_numeric(df_inv['환산'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0)
        
        df_inv['유효일자'] = pd.to_datetime(df_inv['유효일자'], errors='coerce').dt.normalize()
        df_inv['유효일자_보존'] = df_inv['유효일자'].fillna(pd.Timestamp('2099-12-31'))

        box_col_candidates = [col for col in df_inv.columns if 'BOX' in col.upper() or '입수량' in col]
        box_col_name = box_col_candidates[0] if box_col_candidates else None

        # 상품별 '진짜' 1박스 입수량 찾기
        product_box_unit = {}
        if box_col_name:
            for mecode, group in df_inv.groupby('상품'):
                box_vals_clean = group[box_col_name].astype(str).str.replace(r'[^\d.]', '', regex=True)
                box_vals = pd.to_numeric(box_vals_clean, errors='coerce').dropna()
                box_vals = box_vals[box_vals > 0]
                if not box_vals.empty:
                    product_box_unit[mecode] = int(box_vals.min())

        # 불량 재고 걸러내기
        idx_pmm = (df_inv['상품'] == 'ME00621PMM') & (df_inv['유효일자'].dt.year != 2028)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].fillna('').astype(str).str.contains('분리배출'))
        df_inv_valid = df_inv[~(idx_pmm | idx_oc2)].copy()

        # 그룹핑
        if not df_inv_valid.empty:
            df_inv_sorted = df_inv_valid.sort_values(by=['상품', '유효일자_보존', '환산'], ascending=[True, True, False])
            agg_dict = {'환산': 'sum', '화주LOT': 'first', '유효일자': 'first'}
            inv_grouped = df_inv_sorted.groupby(['상품', '유효일자_보존']).agg(agg_dict).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '환산', '화주LOT', '유효일자'])

        df_order['할당상태'] = ''
        df_order['부족시_최대가능수량'] = None
        df_order['부족시_LOT'] = ''
        df_order['부족시_유효일자'] = ''

        # 할당 로직
        with st.spinner('실시간 재고 차감 및 최적화 중... 🚀'):
            for i, row in df_order.iterrows():
                mecode = row['MECODE']
                order_qty = row['수량']
                
                if pd.isna(mecode) or str(mecode) == 'NAN' or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                    
                available_inv = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]
                
                if available_inv.empty:
                    df_order.at[i, 'LOT'] = '재고없음'
                    df_order.at[i, '유효일자'] = '재고없음'
                    df_order.at[i, '할당상태'] = '재고없음'
                    continue

                full_match_inv = available_inv[available_inv['환산'] >= order_qty]

                if not full_match_inv.empty:
                    best_match = full_match_inv.sort_values(by='유효일자_보존').iloc[0]
                else:
                    best_match = available_inv.sort_values(by='유효일자_보존').iloc[0]

                best_idx = best_match.name
                max_qty = best_match['환산']
                lot_str = best_match['화주LOT']
                date_str = best_match['유효일자'].strftime('%Y-%m-%d') if pd.notna(best_match['유효일자']) else '일자없음'
                
                box_unit = product_box_unit.get(mecode, 1)
                    
                if max_qty >= order_qty:
                    allocated_boxes = int(order_qty // box_unit)
                    allocated_qty = allocated_boxes * box_unit
                    state = "정상할당" if allocated_qty == order_qty else f"부분할당({allocated_boxes}BOX)"
                else:
                    allocated_boxes = int(max_qty // box_unit)
                    allocated_qty = allocated_boxes * box_unit
                    state = f"부분할당({allocated_boxes}BOX)" if allocated_qty > 0 else "박스단위부족"

                if allocated_qty > 0:
                    df_order.at[i, '수량'] = allocated_qty
                    df_order.at[i, 'LOT'] = lot_str
                    df_order.at[i, '유효일자'] = date_str
                    df_order.at[i, '할당상태'] = state
                    
                    inv_grouped.at[best_idx, '환산'] -= allocated_qty
                    
                    if allocated_qty < order_qty:
                        df_order.at[i, '부족시_최대가능수량'] = max_qty
                        df_order.at[i, '부족시_LOT'] = lot_str
                        df_order.at[i, '부족시_유효일자'] = date_str
                else:
                    df_order.at[i, 'LOT'] = '박스단위부족'
                    df_order.at[i, '유효일자'] = '박스단위부족'
                    df_order.at[i, '할당상태'] = '박스단위부족'
                    df_order.at[i, '부족시_최대가능수량'] = max_qty
                    df_order.at[i, '부족시_LOT'] = lot_str
                    df_order.at[i, '부족시_유효일자'] = date_str

            if '발주원가' in df_order.columns:
                df_order['발주원가'] = pd.to_numeric(df_order['발주원가'], errors='coerce').fillna(0)
                df_order['발주금액'] = df_order['수량'] * df_order['발주원가']

        # 성공 메시지
        st.success("✅ 모든 데이터 매핑 및 수량 할당이 완료되었습니다!")

        # 결과 확인 및 다운로드 영역
        st.subheader("📊 작업 결과 미리보기")
        preview_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태', '부족시_최대가능수량']
        st.dataframe(df_order[[c for c in preview_cols if c in df_order.columns]].head(15), use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        # 다운로드 버튼을 큼직하게 강조
        st.download_button(
            label="💾 최종 완성 엑셀 파일 다운로드",
            data=buffer.getvalue(),
            file_name="올리브영_수주업로드_자동할당완료.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary" # 버튼을 파란색으로 눈에 띄게!
        )

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다: {e}")
