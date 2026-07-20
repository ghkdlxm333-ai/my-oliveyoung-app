import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime

st.set_page_config(page_title="주문 자동화 프로그램", layout="wide")
st.title("📦 주문 처리 자동화 시스템")

# ---------------------------------------------------------
# 1. 파일 업로드 UI
# ---------------------------------------------------------
col1, col2 = st.columns(2)
with col1:
    file_order = st.file_uploader("1. 발주일 파일 업로드 (.xlsx, .xls, .csv)", type=['xlsx', 'xls', 'csv'])
with col2:
    file_inv = st.file_uploader("2. 일일재고 파일 업로드 (.xlsx, .xls)", type=['xlsx', 'xls'])

if file_order and file_inv:
    try:
        # ---------------------------------------------------------
        # 2. 일일재고 데이터 로드 (무조건 2행 기준: header=1)
        # ---------------------------------------------------------
        df_inv_raw = pd.read_excel(file_inv, header=1)
        
        # 컬럼명 전처리 (공백 제거 및 대문자 변환)
        df_inv_raw.columns = [str(col).strip() for col in df_inv_raw.columns]
        
        # 첫 번째 데이터 행이 헤더 텍스트의 잔여물인 경우 제외
        if not df_inv_raw.empty and str(df_inv_raw.iloc[0].get('순번', '')).strip() == '순번':
            df_inv_raw = df_inv_raw.iloc[1:].reset_index(drop=True)
            
        df_inv = df_inv_raw.copy()

        # 재고 컬럼 표준화 매핑
        inv_col_map = {}
        for c in df_inv.columns:
            c_clean = str(c).replace(" ", "").upper()
            if c_clean in ['상품', '상품코드', '품목코드']:
                inv_col_map[c] = '상품'
            elif c_clean == '화주LOT':
                inv_col_map[c] = '화주LOT'
            elif c_clean == '유효일자':
                inv_col_map[c] = '유효일자'
            elif c_clean in ['합계수량', '환산', '수량', '정상수량']:
                inv_col_map[c] = '환산'
            elif c_clean in ['상품바코드', '바코드']:
                inv_col_map[c] = '상품바코드'
            elif c_clean in ['입수량(BOX)', 'BOX입수량', '입수량']:
                inv_col_map[c] = '입수량(BOX)'

        df_inv = df_inv.rename(columns=inv_col_map)

        # 필수 컬럼 검증
        if '상품' not in df_inv.columns:
            st.error("일일재고 파일의 2행에서 '상품' 컬럼을 찾을 수 없습니다. 컬럼명을 확인해 주세요.")
            st.stop()

        # 데이터 기본 정제
        df_inv['상품'] = df_inv['상품'].astype(str).str.strip().str.upper()
        df_inv['화주LOT'] = df_inv['화주LOT'].astype(str).str.strip()
        df_inv['환산'] = pd.to_numeric(df_inv['환산'], errors='coerce').fillna(0)

        # 유효일자 날짜 변환
        df_inv['유효일자_DT'] = pd.to_datetime(df_inv['유효일자'], errors='coerce')
        df_inv['유효일자_STR'] = df_inv['유효일자_DT'].dt.strftime('%Y-%m-%d').fillna('')
        df_inv['유효일자_보존'] = df_inv['유효일자_STR']

        # 재고 필터링 조건 (소비기한 1년 6개월 / 548일 이하 및 특정 조건 제외)
        today = pd.Timestamp.today().normalize()
        limit_date = today + pd.Timedelta(days=548)

        idx_short_life = (df_inv['유효일자_DT'].fillna(pd.Timestamp('2099-12-31')) <= limit_date)
        idx_oc2 = (df_inv['상품'] == 'ME90621OC2') & (~df_inv['화주LOT'].astype(str).str.contains('분리배출', na=False))

        # 유효한 재고만 추출
        df_inv_valid = df_inv[~(idx_oc2 | idx_short_life)].copy()

        # 상품별/유효일자별 재고 집계
        inv_grouped = df_inv_valid.groupby(['상품', '유효일자_보존'], as_index=False).agg({
            '환산': 'sum',
            '화주LOT': 'first',
            '유효일자_STR': 'first'
        })

        # ---------------------------------------------------------
        # 3. 발주 데이터 로드
        # ---------------------------------------------------------
        if str(file_order.name).endswith('.csv'):
            try:
                df_order = pd.read_csv(file_order, encoding='utf-8-sig')
            except:
                df_order = pd.read_csv(file_order, encoding='cp949')
        else:
            df_order = pd.read_excel(file_order)

        # 오늘 날짜 강제 적용 (Order Date)
        today_str = datetime.now().strftime('%Y-%m-%d')
        if '주문일자' in df_order.columns or 'A' in df_order.columns:
            df_order.iloc[:, 0] = today_str

        st.success("✅ 파일 데이터 로드 및 2행 헤더 정제가 완료되었습니다.")

        # ---------------------------------------------------------
        # 4. 결과 출력 및 탭 구성 (상세 내역 & 최종 결과)
        # ---------------------------------------------------------
        tab1, tab2 = st.tabs(["📋 상세 내역", "📊 최종 결과 요약"])

        with tab1:
            st.subheader("일일재고 정제 데이터 (유효 재고)")
            st.dataframe(inv_grouped)

        with tab2:
            st.subheader("발주 데이터 현황")
            st.dataframe(df_order.head(50))

    except Exception as e:
        st.error(f"처리 중 오류 발생: {e}")
