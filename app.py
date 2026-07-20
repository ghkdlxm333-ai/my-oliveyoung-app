import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. 페이지 설정 및 커스텀 CSS
# ==========================================
st.set_page_config(
    page_title="올리브영 수주업로드 자동 입력 시스템", 
    page_icon="https://raw.githubusercontent.com/paak1010/mentholatum_oliveyoung/main/logo.png",
    layout="wide"
)

custom_css = """
<style>
[data-testid="stHeader"] { visibility: hidden; }
footer { visibility: hidden; }
[data-testid="stSidebar"] { background-color: #FFFFFF !important; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 2. 센터별 배송코드 매핑 데이터
# ==========================================
DELIVERY_CODE_MAP = {
    '부곡센터': '86100086',
    '중부센터': '86100118',
    '양지온라인센터': '86101125',
    '양지센터': '86101126',
    '경산센터': '81032980'
}

def clean_barcode(val):
    """바코드 숫자 13자리 정제 (지수표기법 및 .0 방지)"""
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if 'e' in val_str.lower():
        try:
            val_str = f"{float(val_str):.0f}"
        except:
            pass
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

def to_safe_float(series):
    cleaned = series.astype(str).str.replace(r'[^0-9.]', '', regex=True)
    return pd.to_numeric(cleaned, errors='coerce').fillna(0)

# ==========================================
# 🎨 사이드바 파일 업로드
# ==========================================
with st.sidebar:
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    st.markdown("---")
    st.header("⚙️ 파일 업로드")
    
    order_file = st.file_uploader("1️⃣ 올영 주문 리스트 (납품확인서 목록)", type=['xlsx', 'xls'])
    inv_file = st.file_uploader("2️⃣ WMS 일일재고 엑셀", type=['xlsx', 'xls'])
    
    st.markdown("---")
    st.caption("💡 두 원본 파일만 업로드하면 자동으로 서식 구성 및 수주 매핑 진행")
    st.caption("✔️ 잔여 유효일자 548일 이하 제외")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("올리브영 수주업로드 자동 입력 시스템")
st.markdown("Mentholatum : Moving The Heart")

if order_file and inv_file:
    try:
        # ------------------------------------------
        # 1. 일일재고 데이터 로드 및 정제
        # ------------------------------------------
        df_inv_raw = pd.read_excel(inv_file, header=0)
        # 헤더 위치 자동 감지 (1번 행이 헤더인 경우 처리)
        if '상품' not in df_inv_raw.columns and '순번' in df_inv_raw.iloc[0].values:
            df_inv_raw = pd.read_excel(inv_file, header=1)

        df_inv = df_inv_raw.copy()

        # 재고 컬럼 매핑
        inv_col_map = {}
        for c in df_inv.columns:
            c_str = str(c).replace(" ", "").upper()
            if c_str == '상품':
                inv_col_map[c] = '상품'
            elif 'LOT' in c_str:
                inv_col_map[c] = '화주LOT'
            elif '유효' in c_str or '유통' in c_str:
                inv_col_map[c] = '유효일자'
            elif '합계' in c_str or '환산' in c_str or '수량' in c_str:
                if '합계수량' in c_str or '환산' in c_str:
                    inv_col_map[c] = '환산'
            elif '상품바코드' in c_str or '바코드' in c_str:
                inv_col_map[c] = '상품바코드'
            elif 'BOX' in c_str or '입수' in c_str:
                inv_col_map[c] = '입수량(BOX)'

        df_inv.rename(columns=inv_col_map, inplace=True)

        # 재고 기반 바코드 -> MECODE 매핑 테이블 자동 생성
        barcode_to_mecode = {}
        if '상품바코드' in df_inv.columns and '상품' in df_inv.columns:
            for _, r in df_inv.iterrows():
                bc = clean_barcode(r['상품바코드'])
                mc = str(r['상품']).strip().upper()
                if bc and mc and mc not in ['NAN', 'NONE']:
                    barcode_to_mecode[bc] = mc

        # ------------------------------------------
        # 2. 올영 주문 데이터(납품확인서 목록) 로드 및 서식 자동 구성
        # ------------------------------------------
        df_ord_raw = pd.read_excel(order_file)
        
        df_order = pd.DataFrame()
        df_order['발주처코드'] = '86100000'
        
        # 입고예정일 날짜 변환
        df_order['입고예정일'] = pd.to_datetime(df_ord_raw['입고예정일'], errors='coerce').dt.strftime('%Y-%m-%d 00:00:00')
        
        # 배송코드 매핑
        def get_delivery_code(center):
            c_str = str(center)
            for k, v in DELIVERY_CODE_MAP.items():
                if k in c_str:
                    return v
            return ""
        
        df_order['배송코드'] = df_ord_raw['센터'].apply(get_delivery_code)
        df_order['ORDER #'] = ""
        df_order['상품명'] = df_ord_raw['상품명']
        df_order['바코드'] = df_ord_raw['상품코드'].apply(clean_barcode)
        
        # 재고 파일에서 추출한 바코드 매핑 적용
        df_order['MECODE'] = df_order['바코드'].map(barcode_to_mecode).fillna("")
        
        df_order['수량'] = to_safe_float(df_ord_raw['발주수량\n(EA)'])
        
        # 원단가 및 금액
        unit_price_col = [c for c in df_ord_raw.columns if '원단가' in str(c)]
        if unit_price_col:
            df_order['발주원가'] = to_safe_float(df_ord_raw[unit_price_col[0]])
            df_order['발주금액'] = df_order['수량'] * df_order['발주원가']
        else:
            df_order['발주원가'] = 0
            df_order['발주금액'] = 0

        # 결과 저장 컬럼 생성
        new_cols = ['LOT', '유효일자', '할당상태', '부족시_최대가능수량', '부족시_LOT', '부족시_유효일자']
        for col in new_cols:
            df_order[col] = ""
            df_order[col] = df_order[col].astype(object)

        # ------------------------------------------
        # 3. 재고 전처리 및 유효일자(548일) 필터링
        # ------------------------------------------
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        df_inv['환산'] = to_safe_float(df_inv['환산'])
        df_inv['유효일자_DT'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')
        df_inv['유효일자_보존'] = df_inv['유효일자_DT'].fillna(pd.Timestamp('2099-12-31'))
        df_inv['유효일자_STR'] = df_inv['유효일자_DT'].dt.strftime('%Y-%m-%d 00:00:00').fillna('')

        # 입수량 파싱
        product_box_unit = {}
        if '입수량(BOX)' in df_inv.columns:
            for mecode, group in df_inv.groupby('상품'):
                box_vals = to_safe_float(group['입수량(BOX)'])
                box_vals = box_vals[box_vals > 0]
                if not box_vals.empty:
                    product_box_unit[mecode] = int(box_vals.min())

        # 유효일자 548일 기준 차감
        today = pd.Timestamp.today().normalize()
        cutoff_date = today + pd.Timedelta(days=548)
        idx_short_life = (df_inv['유효일자_보존'] <= cutoff_date)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].astype(str).str.contains('분리배출'))

        df_inv_valid = df_inv[~(idx_oc2 | idx_short_life)].copy()

        if not df_inv_valid.empty:
            inv_grouped = df_inv_valid.groupby(['상품', '유효일자_보존']).agg({
                '환산': 'sum',
                '화주LOT': 'first',
                '유효일자_STR': 'first'
            }).reset_index()
        else:
            inv_grouped = pd.DataFrame(columns=['상품', '유효일자_보존', '환산', '화주LOT', '유효일자_STR'])

        # ------------------------------------------
        # 4. 재고 매칭 및 할당 로직 수행
        # ------------------------------------------
        with st.spinner('자동 변환 및 재고 매칭 중...'):
            for i, row in df_order.iterrows():
                mecode = str(row['MECODE'])
                order_qty = float(row['수량'])

                if mecode in ['NAN', '', 'NONE'] or order_qty <= 0:
                    df_order.at[i, '할당상태'] = "MECODE미매핑"
                    continue

                available_inv = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]

                if available_inv.empty:
                    df_order.at[i, 'LOT'], df_order.at[i, '유효일자'], df_order.at[i, '할당상태'] = '재고없음', '재고없음', '재고없음'
                    continue

                full_match = available_inv[available_inv['환산'] >= order_qty]
                best_match = full_match.sort_values(by='유효일자_보존').iloc[0] if not full_match.empty else available_inv.sort_values(by='유효일자_보존').iloc[0]

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

        # ------------------------------------------
        # 5. 결과 화면 출력 및 엑셀 다운로드
        # ------------------------------------------
        st.success("✅ 주문 원본 파일 자동 변환 및 수주 매핑이 완료되었습니다!")

        st.subheader("📊 서식(수주업로드) 자동 작성 결과 미리보기 (상위 100건)")
        
        display_cols = ['발주처코드', '입고예정일', '배송코드', '상품명', '바코드', 'MECODE', '수량', '발주원가', '발주금액', 'LOT', '유효일자', '할당상태']
        existing_cols = [c for c in display_cols if c in df_order.columns]
        
        st.dataframe(df_order[existing_cols].head(100), use_container_width=True, hide_index=True)

        # 서식 다운로드 생성
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_order.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            workbook = writer.book
            worksheet = writer.sheets['서식(수주업로드)']
            text_format = workbook.add_format({'num_format': '@'})

            for target_col in ['유효일자', '부족시_유효일자']:
                if target_col in df_order.columns:
                    idx = df_order.columns.get_loc(target_col)
                    worksheet.set_column(idx, idx, 20, text_format)

        st.download_button(
            label="💾 완성본 엑셀 다운로드 (서식 형태)",
            data=buffer.getvalue(),
            file_name="올리브영_자동생성_수주업로드.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    except Exception as e:
        st.error(f"오류 발생: {e}")
else:
    st.info("👈 사이드바에서 [1️⃣ 올영 주문 리스트]와 [2️⃣ WMS 일일재고 엑셀] 두 파일을 모두 업로드해주세요.")
