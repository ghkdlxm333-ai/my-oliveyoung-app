import streamlit as st
import pandas as pd
import io
from datetime import datetime

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
# 2. 센터별 배송코드 매핑 데이터 및 유틸 함수
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

def clean_date_str(val):
    """00:00:00 시간을 제거하고 YYYY-MM-DD 날짜만 추출"""
    if pd.isna(val) or not val:
        return ""
    dt = pd.to_datetime(val, errors='coerce')
    if pd.isna(dt):
        return str(val).split(' ')[0]
    return dt.strftime('%Y-%m-%d')

# ==========================================
# 🎨 사이드바 파일 통합 업로드 (한 곳에 드래그&드롭)
# ==========================================
with st.sidebar:
    st.image("https://static.wikia.nocookie.net/mycompanies/images/d/de/Fe328a0f-a347-42a0-bd70-254853f35374.jpg/revision/latest?cb=20191117172510", use_container_width=True)
    st.markdown("---")
    st.header("⚙️ 파일 업로드")
    
    # 두 파일을 한 곳에서 한 번에 받도록 통합
    uploaded_files = st.file_uploader(
        "📂 엑셀 파일 2개 한 번에 드래그&드롭\n(주문서 리스트 & WMS 일일재고)", 
        type=['xlsx', 'xls'], 
        accept_multiple_files=True
    )
    
    st.markdown("---")
    st.caption("💡 파일을 한 곳에 모두 넣어주시면 자동으로 종류를 판별합니다.")
    st.caption("✔️ 잔여 유효일자 548일 이하 자동 제외")

# ==========================================
# 메인 화면 디자인
# ==========================================
st.title("올리브영 수주업로드 자동 입력 시스템")

if uploaded_files:
    order_file = None
    inv_file = None

    # 자동 파일 유형 판별 로직
    for file in uploaded_files:
        try:
            df_temp = pd.read_excel(file, nrows=5)
            cols_str = " ".join([str(c) for c in df_temp.columns])
            
            # 주문서 판별 (입고예정일, 발주수량, 센터 등)
            if '입고예정일' in cols_str or '발주수량' in cols_str or '센터' in cols_str:
                order_file = file
            # 일일재고 판별 (화주LOT, 합계수량, 환산, 상품바코드 등)
            elif '화주LOT' in cols_str or '환산' in cols_str or '합계수량' in cols_str or '일일재고' in file.name:
                inv_file = file
            else:
                # 2번째 행 헤더 확인
                df_temp2 = pd.read_excel(file, header=1, nrows=5)
                cols_str2 = " ".join([str(c) for c in df_temp2.columns])
                if '화주LOT' in cols_str2 or '환산' in cols_str2 or '상품' in cols_str2:
                    inv_file = file
        except Exception:
            pass

    if not order_file or not inv_file:
        st.warning("⚠️ 주문서 파일(납품확인서 목록)과 WMS 일일재고 파일 2개가 모두 인식되어야 합니다. 파일 내용을 확인해주세요.")
    else:
        try:
            # ------------------------------------------
            # 1. 일일재고 데이터 로드 및 정제
            # ------------------------------------------
            df_inv_raw = pd.read_excel(inv_file, header=0)
            
            if len(df_inv_raw) > 0 and '순번' in df_inv_raw.columns and str(df_inv_raw.iloc[0]['순번']).strip() == '순번':
                df_inv_raw = df_inv_raw.iloc[1:].reset_index(drop=True)

            df_inv = df_inv_raw.copy()

            # 정밀한 열 매핑
            inv_col_map = {}
            for c in df_inv.columns:
                c_clean = str(c).replace(" ", "").upper()
                if c_clean == '상품':
                    inv_col_map[c] = '상품'
                elif c_clean == '화주LOT':
                    inv_col_map[c] = '화주LOT'
                elif c_clean == '유효일자':
                    inv_col_map[c] = '유효일자'
                elif c_clean in ['합계수량', '환산']:
                    inv_col_map[c] = '환산'
                elif c_clean == '상품바코드':
                    inv_col_map[c] = '상품바코드'
                elif c_clean in ['입수량(BOX)', 'BOX입수량']:
                    inv_col_map[c] = '입수량(BOX)'

            df_inv.rename(columns=inv_col_map, inplace=True)

            # 바코드 -> MECODE 매핑
            barcode_to_mecode = {}
            if '상품바코드' in df_inv.columns and '상품' in df_inv.columns:
                for _, r in df_inv.iterrows():
                    bc = clean_barcode(r['상품바코드'])
                    mc = str(r['상품']).strip().upper()
                    if bc and mc and mc not in ['NAN', 'NONE']:
                        barcode_to_mecode[bc] = mc

            # ------------------------------------------
            # 2. 올영 주문 데이터(납품확인서) 로드 및 [수정] 양식 표준 구성
            # ------------------------------------------
            df_ord_raw = pd.read_excel(order_file)
            
            df_order = pd.DataFrame()
            df_order['orig_idx'] = df_ord_raw.index  # 원본 순서 보존
            
            # [수정] 시트 컬럼 양식대로 생성
            df_order['출고구분'] = 0
            df_order['수주일자'] = datetime.today().strftime('%Y-%m-%d')
            df_order['납품일자'] = df_ord_raw['입고예정일'].apply(clean_date_str)
            df_order['발주처코드'] = '86100000'
            df_order['발주처'] = '올리브영'
            
            def get_delivery_code(center):
                c_str = str(center)
                for k, v in DELIVERY_CODE_MAP.items():
                    if k in c_str:
                        return v
                return ""
            
            df_order['배송코드'] = df_ord_raw['센터'].apply(get_delivery_code)
            df_order['배송처'] = df_ord_raw['센터'].fillna("")
            
            df_order['바코드'] = df_ord_raw['상품코드'].apply(clean_barcode)
            df_order['MECODE'] = df_order['바코드'].map(barcode_to_mecode).fillna("")
            df_order['상품명'] = df_ord_raw['상품명']
            df_order['수량'] = to_safe_float(df_ord_raw['발주수량\n(EA)'])
            
            unit_price_col = [c for c in df_ord_raw.columns if '원단가' in str(c)]
            if unit_price_col:
                df_order['발주원가'] = to_safe_float(df_ord_raw[unit_price_col[0]])
                df_order['발주금액'] = df_order['수량'] * df_order['발주원가']
            else:
                df_order['발주원가'] = 0
                df_order['발주금액'] = 0

            df_order['LOT'] = ""
            df_order['유효일자'] = ""
            df_order['할당상태'] = ""

            # ------------------------------------------
            # 3. 재고 전처리 및 548일 필터링
            # ------------------------------------------
            df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
            df_inv['환산'] = to_safe_float(df_inv['환산'])
            df_inv['유효일자_DT'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')
            df_inv['유효일자_보존'] = df_inv['유효일자_DT'].fillna(pd.Timestamp('2099-12-31'))
            df_inv['유효일자_STR'] = df_inv['유효일자_DT'].dt.strftime('%Y-%m-%d').fillna('')  # 시간 제거 YYYY-MM-DD

            product_box_unit = {}
            if '입수량(BOX)' in df_inv.columns:
                for mecode, group in df_inv.groupby('상품'):
                    box_vals = to_safe_float(group['입수량(BOX)'])
                    box_vals = box_vals[box_vals > 0]
                    if not box_vals.empty:
                        product_box_unit[mecode] = int(box_vals.min())

            today = pd.Timestamp.today().normalize()
            cutoff_date = today + pd.Timedelta(days=548)
            idx_short_life = (df_inv['유효일자_보존'] <= cutoff_date)
            idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].astype(str).str.contains('분리배출', na=False))

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
            # 4. 재고 매칭 및 할당 (원본 행 순서 보장)
            # ------------------------------------------
            with st.spinner('자동 변환 및 재고 매칭 중...'):
                for i in range(len(df_order)):
                    mecode = str(df_order.at[i, 'MECODE'])
                    order_qty = float(df_order.at[i, '수량'])

                    if mecode in ['NAN', '', 'NONE'] or order_qty <= 0:
                        df_order.at[i, '할당상태'] = "MECODE미매핑"
                        continue

                    available_inv = inv_grouped[(inv_grouped['상품'] == mecode) & (inv_grouped['환산'] > 0)]

                    if available_inv.empty:
                        df_order.at[i, 'LOT'], df_order.at[i, '유효일자'], df_order.at[i, '할당상태'] = '', '', '재고없음'
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
                        df_order.at[i, '수량'] = int(allocated_qty)
                        df_order.at[i, 'LOT'] = lot_str
                        df_order.at[i, '유효일자'] = date_str  # YYYY-MM-DD
                        df_order.at[i, '할당상태'] = "정상할당" if allocated_qty == order_qty else f"부분할당({allocated_boxes}BOX)"
                        inv_grouped.at[best_idx, '환산'] -= allocated_qty
                    else:
                        df_order.at[i, '할당상태'] = '박스단위부족'

            # 원본 순서 재정렬
            df_order = df_order.sort_values(by='orig_idx').drop(columns=['orig_idx', '바코드']).reset_index(drop=True)

            # [수정] 시트의 정확한 15개 컬럼 순서 지정
            final_cols = [
                '출고구분', '수주일자', '납품일자', '발주처코드', '발주처', 
                '배송코드', '배송처', 'MECODE', '상품명', '수량', 
                '발주원가', '발주금액', 'LOT', '유효일자', '할당상태'
            ]
            df_final = df_order[final_cols]

            # ------------------------------------------
            # 5. 결과 화면 출력 및 엑셀 다운로드
            # ------------------------------------------
            st.success("✅ 주문 원본 순서 및 [수정] 시트 양식 매핑이 완료되었습니다!")

            st.subheader("📊 수주 매핑 결과 미리보기 ([수정] 시트 동일 양식)")
            st.dataframe(df_final, use_container_width=True, hide_index=True)

            # 다운로드 파일 생성
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
                workbook = writer.book
                worksheet = writer.sheets['서식(수주업로드)']
                text_format = workbook.add_format({'num_format': '@'})

                # 날짜 및 텍스트 컬럼 서식 적용
                for target_col in ['수주일자', '납품일자', '유효일자', '배송코드', '발주처코드']:
                    if target_col in df_final.columns:
                        idx = df_final.columns.get_loc(target_col)
                        worksheet.set_column(idx, idx, 15, text_format)

            st.download_button(
                label="💾 완성본 엑셀 다운로드 (수주업로드 서식)",
                data=buffer.getvalue(),
                file_name="올리브영_수주업로드_완성본.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"처리 중 오류 발생: {e}")
else:
    st.info("👈 사이드바의 업로드 창에 [주문서 파일]과 [일일재고 파일] 2개를 함께 드래그&드롭 해주세요.")
