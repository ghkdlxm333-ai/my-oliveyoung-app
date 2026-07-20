import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 페이지 기본 설정 (깃허브 로고 적용)
# ==========================================
st.set_page_config(
    page_title="올리브영 수주업로드 자동 입력 시스템", 
    page_icon="https://raw.githubusercontent.com/paak1010/mentholatum_oliveyoung/main/logo.png",
    layout="wide"
)

# ==========================================
# 2. 커스텀 CSS (사이드바 흰색 & 기본 메뉴 숨기기)
# ==========================================
custom_css = """
<style>
/* 우측 상단 기본 햄버거 메뉴 및 Deploy 헤더 숨기기 */
[data-testid="stHeader"] {
    visibility: hidden;
}

/* 하단 Streamlit 워터마크 숨기기 */
footer {
    visibility: hidden;
}

/* 사이드바 배경색을 완전히 흰색(#FFFFFF)으로 변경 */
[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 🎨 사이드바 디자인
# ==========================================
with st.sidebar:
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    st.markdown("---")
    st.header("⚙️ 작업 설정")
    uploaded_file = st.file_uploader("올리브영 발주 엑셀 업로드", type=['xlsx'])
    st.markdown("---")
    st.caption("💡 자동 부분 할당 및 재고 차감 적용")
    st.caption("✔️ 잔여 유효일자 548일 이하 제외")
    st.caption("Developed by Jay")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("올리브영 수주업로드 자동 입력 시스템")
st.markdown("Mentholatum : Moving The Heart")

def to_safe_float(series):
    """어떤 타입이 들어와도 숫자만 추출하여 float로 변환"""
    cleaned = series.astype(str).str.replace(r'[^0-9.]', '', regex=True)
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

if uploaded_file:
    try:
        # 데이터 읽기
        df_order_raw = pd.read_excel(uploaded_file, sheet_name='서식(수주업로드)', header=1)
        df_inv_raw = pd.read_excel(uploaded_file, sheet_name='재고', header=2)
        
        df_order = df_order_raw.copy()
        df_inv = df_inv_raw.copy()

        # ==========================================
        # 🛡️ [재고 시트] 열 순서 및 이름 변경 방어 코드 (Robust 매칭)
        # ==========================================
        rename_dict = {}
        for col in df_inv.columns:
            col_str = str(col).replace(" ", "").upper()
            if '상품' in col_str and '상품명' not in col_str:
                rename_dict[col] = '상품'
            elif 'LOT' in col_str:
                rename_dict[col] = '화주LOT'
            elif '유효일자' in col_str or '유통기한' in col_str:
                rename_dict[col] = '유효일자'
            elif '환산' in col_str:
                rename_dict[col] = '환산'
        
        df_inv.rename(columns=rename_dict, inplace=True)
        # ==========================================

        # 불필요한 열 제거
        if '잔여일수' in df_order.columns:
            start_idx = list(df_order.columns).index('잔여일수')
            cols_to_drop = df_order.columns[start_idx:]
            df_order = df_order.drop(columns=cols_to_drop)

        # 결과 컬럼 초기화 (범용 타입 지정)
        new_cols = ['LOT', '유효일자', '할당상태', '부족시_최대가능수량', '부족시_LOT', '부족시_유효일자']
        for col in new_cols:
            df_order[col] = ""
            df_order[col] = df_order[col].astype(object)

        # 데이터 정제 (매핑된 컬럼명을 안전하게 호출)
        df_order['MECODE'] = df_order['MECODE'].astype(str).str.strip().str.upper()
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        df_order['수량'] = to_safe_float(df_order['수량']).astype(float)
        df_inv['환산'] = to_safe_float(df_inv['환산']).astype(float)
        
        # 유효일자 처리 (시간 제거)
        df_inv['유효일자_DT'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')
        df_inv['유효일자_보존'] = df_inv['유효일자_DT'].fillna(pd.Timestamp('2099-12-31'))
        df_inv['유효일자_STR'] = df_inv['유효일자_DT'].dt.strftime('%Y-%m-%d').fillna('')

        # [박스 입수량 계산] 열 이름에 'BOX'나 '입수량'이 들어간 아무 열이나 동적으로 찾음
        box_col_candidates = [col for col in df_inv.columns if 'BOX' in str(col).upper() or '입수량' in str(col)]
        box_col_name = box_col_candidates[0] if box_col_candidates else None
        product_box_unit = {}
        if box_col_name:
            for mecode, group in df_inv.groupby('상품'):
                box_vals = to_safe_float(group[box_col_name])
                box_vals = box_vals[box_vals > 0]
                if not box_vals.empty:
                    product_box_unit[mecode] = int(box_vals.min())

        # ==========================================
        # 🔥 [정확한 일수 적용] 재고 필터링 조건 강화
        # ==========================================
        today = pd.Timestamp.today().normalize()
        cutoff_date = today + pd.Timedelta(days=548)
        idx_short_shelf_life = (df_inv['유효일자_보존'] <= cutoff_date)

        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].astype(str).str.contains('분리배출'))
        
        df_inv_valid = df_inv[~(idx_oc2 | idx_short_shelf_life)].copy()

        # [재고 그룹핑]
        df_inv_valid['화주LOT'] = df_inv_valid['화주LOT'].astype(str)
        if not df_inv_valid.empty:
            inv_grouped = df_inv_valid.groupby(['상품', '유효일자_보존']).agg({
                '환산': 'sum', 
                '화주LOT': 'first', 
                '유효일자_STR': 'first'
            }).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '환산', '화주LOT', '유효일자_STR'])

        # 🚀 할당 로직
        with st.spinner('재고 매칭 중...'):
            for i, row in df_order.iterrows():
                mecode = str(row['MECODE'])
                order_qty = float(row['수량'])
                
                if mecode in ['NAN', '', 'NONE'] or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "제외"
                    continue
                    
                available_inv = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]
                
                if available_inv.empty:
                    df_order.at[i, 'LOT'], df_order.at[i, '유효일자'], df_order.at[i, '할당상태'] = '재고없음', '재고없음', '재고없음'
                    continue

                full_match_inv = available_inv[available_inv['환산'] >= order_qty]
                best_match = full_match_inv.sort_values(by='유효일자_보존').iloc[0] if not full_match_inv.empty else available_inv.sort_values(by='유효일자_보존').iloc[0]

                best_idx = best_match.name
                max_qty = float(best_match['환산'])
                lot_str = str(best_match['화주LOT'])
                date_str = str(best_match['유효일자_STR']) 
                
                box_unit = product_box_unit.get(mecode, 1)
                potential_qty = min(order_qty, max_qty)
                allocated_boxes = int(potential_qty // box_unit)
                allocated_qty = float(allocated_boxes * box_unit)

                if allocated_qty > 0:
                    df_order.at[i, '수량'] = allocated_qty
                    df_order.at[i, 'LOT'] = lot_str
                    df_order.at[i, '유효일자'] = date_str
                    df_order.at[i, '할당상태'] = "정상할당" if allocated_qty == order_qty else f"부분할당({allocated_boxes}BOX)"
                    inv_grouped.at[best_idx, '환산'] -= allocated_qty
                else:
                    df_order.at[i, '할당상태'] = '박스단위부족'
                    df_order.at[i, '부족시_최대가능수량'] = max_qty
                    df_order.at[i, '부족시_LOT'] = lot_str
                    df_order.at[i, '부족시_유효일자'] = date_str

        # ==========================================
        # 📊 화면 표시용 미리보기
        # ==========================================
        st.success("✅ 처리가 완료되었습니다!")
        
        st.subheader("📊 작업 결과 미리보기 (상위 100건)")
        view_cols = ['MECODE', '상품명', '수량', 'LOT', '유효일자', '할당상태']
        existing_view_cols = [c for c in view_cols if c in df_order.columns]
        
        df_display = df_order[existing_view_cols].head(100).copy()
        df_safe_display = pd.DataFrame(
            df_display.to_numpy().astype(str), 
            columns=df_display.columns
        )
        
        st.dataframe(df_safe_display, use_container_width=True, hide_index=True)

        # ==========================================
        # 💾 엑셀 다운로드
        # ==========================================
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            workbook = writer.book
            worksheet = writer.sheets['서식(수주업로드)']
            text_format = workbook.add_format({'num_format': '@'}) 
            
            for target_col in ['유효일자', '부족시_유효일자']:
                if target_col in df_order.columns:
                    idx = df_order.columns.get_loc(target_col)
                    worksheet.set_column(idx, idx, 15, text_format)

        st.download_button(
            label="💾 최종 완성본 엑셀 다운로드", 
            data=buffer.getvalue(), 
            file_name="올리브영_자동할당완료.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            type="primary"
        )

    except Exception as e:
        st.error(f"오류 발생: {e}")
