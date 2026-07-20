import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(page_title="올리브영 수주 매핑 자동화", layout="wide")

st.title("📦 올리브영 발주 - 3PL WMS 일일재고 자동 매핑 솔루션")
st.caption("1년 6개월 이상 잔여 유효기간 & 박스 입수량 이상 재고만 자동 계산하여 단일 LOT를 매핑합니다.")

# 센터명 -> 배송코드 매핑 사전
CENTER_MAP = {
    '[LA02] 양지센터': '86101126',
    '[L002] 부곡센터': '86100086',
    '[L003] 중부센터': '86100118',
    '[L001] 수도권센터': '86100000'
}

# ---------------------------------------------------------
# Sidebar File Upload
# ---------------------------------------------------------
st.sidebar.header("📁 엑셀 파일 업로드")
stock_file = st.sidebar.file_uploader("1. WMS 일일재고 파일 (.xlsx)", type=["xlsx"])
order_file = st.sidebar.file_uploader("2. 올리브영 납품확인서 목록 (.xlsx)", type=["xlsx"])

# ---------------------------------------------------------
# Main Logic
# ---------------------------------------------------------
if stock_file and order_file:
    try:
        # =========================================================
        # 1. 일일재고 파일 읽기 (D열 및 인덱스 위치 기반 파싱)
        # =========================================================
        df_stock_raw = pd.read_excel(stock_file, header=None)
        
        # 데이터는 3번째 행(인덱스 2)부터 시작
        df_stock = df_stock_raw.iloc[2:].copy()
        
        # Col 3: D열 (ME상품코드), Col 6: G열 (화주LOT), Col 13: N열 (유효일자),
        # Col 17: R열 (입수량(BOX)), Col 31: AF열 (합계수량), Col 33: AH열 (상품바코드)
        stock_data = pd.DataFrame({
            'ME코드': df_stock[3].astype(str).str.strip(),
            '화주LOT': df_stock[6].astype(str).str.strip(),
            '유효일자': pd.to_datetime(df_stock[13], errors='coerce'),
            '입수량_BOX': pd.to_numeric(df_stock[17], errors='coerce').fillna(1),
            '합계수량': pd.to_numeric(df_stock[31], errors='coerce').fillna(0),
            '상품바코드': df_stock[33].astype(str).str.replace('.0', '', regex=False).str.strip()
        })
        
        # =========================================================
        # 2. 올리브영 납품확인서 목록 파일 읽기
        # =========================================================
        df_order_raw = pd.read_excel(order_file, header=None)
        
        order_header_idx = 0
        for i, row in df_order_raw.iterrows():
            if '상품코드' in row.values:
                order_header_idx = i
                break
                
        df_order = pd.read_excel(order_file, header=order_header_idx)
        df_order.columns = [str(c).strip() for c in df_order.columns]
        
        df_order['상품코드_str'] = df_order['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_order['입고예정일'] = pd.to_datetime(df_order['입고예정일'], errors='coerce')
        df_order['발주수량\n(EA)'] = pd.to_numeric(df_order['발주수량\n(EA)'], errors='coerce').fillna(0)
        df_order['원단가'] = pd.to_numeric(df_order.get('원단가', 0), errors='coerce').fillna(0)
        df_order['원가금액'] = pd.to_numeric(df_order.get('원가금액', 0), errors='coerce').fillna(0)
        
        # =========================================================
        # 3. 매핑 및 출고 제약조건 처리
        # =========================================================
        results = []
        today = datetime.now()
        
        for idx, row in df_order.iterrows():
            barcode = row['상품코드_str']
            order_qty = row['발주수량\n(EA)']
            order_date = row['입고예정일'] if pd.notnull(row['입고예정일']) else today
            center_name = str(row.get('센터', ''))
            unit_price = row.get('원단가', 0)
            total_amount = row.get('원가금액', 0)
            
            # 1) 바코드 또는 ME코드로 재고 탐색
            sub_stock = stock_data[
                (stock_data['상품바코드'] == barcode) | 
                (stock_data['ME코드'] == barcode)
            ].copy()
            
            # 2) 🛑 [제약 1] 유효기간 1년 6개월(547일) 미만 재고 차단
            min_valid_date = order_date + timedelta(days=547)
            valid_stock = sub_stock[sub_stock['유효일자'] >= min_valid_date].copy()
            
            # 3) 🛑 [제약 2] 박스 입수량 미만 재고(단수 재고) 차단
            valid_stock = valid_stock[valid_stock['합계수량'] >= valid_stock['입수량_BOX']]
            
            # 4) FEFO 정렬
            valid_stock = valid_stock.sort_values(by='유효일자', ascending=True)
            
            selected_lot = ""
            selected_exp_date = ""
            status_msg = "정상"
            me_code = ""
            
            if not valid_stock.empty:
                me_code = valid_stock.iloc[0]['ME코드']
                fefo_match = valid_stock[valid_stock['합계수량'] >= order_qty]
                
                if not fefo_match.empty:
                    best_lot = fefo_match.iloc[0]
                    selected_lot = best_lot['화주LOT']
                    selected_exp_date = best_lot['유효일자'].strftime('%Y-%m-%d') if pd.notnull(best_lot['유효일자']) else ""
                else:
                    total_avail_qty = valid_stock['합계수량'].sum()
                    if total_avail_qty >= order_qty:
                        status_msg = "검토필요 (LOT 분할 필요)"
                    else:
                        status_msg = "검토필요 (유효재고 부족)"
            else:
                status_msg = "검토필요 (출고가능 재고없음)"
                
            shipping_code = CENTER_MAP.get(center_name, '')
            
            # 요청하신 지정 순서대로 dict 생성
            out_row = {
                '입고예정일': order_date.strftime('%Y-%m-%d') if pd.notnull(order_date) else '',
                '발주처코드': '86100000',
                '배송코드': shipping_code,
                'MEcode': me_code,
                '상품명': row.get('상품명', ''),
                '수량': order_qty,
                '단가': unit_price,
                '발주원가': unit_price,
                '발주금액': total_amount,
                'LOT': selected_lot,
                '유효일자': selected_exp_date,
                '매핑상태': status_msg
            }
            results.append(out_row)
            
        df_result = pd.DataFrame(results)
        
        # =========================================================
        # 4. 결과 지표 및 표 하이라이팅 연출
        # =========================================================
        normal_cnt = (df_result['매핑상태'] == '정상').sum()
        review_cnt = len(df_result) - normal_cnt
        
        col1, col2, col3 = st.columns(3)
        col1.metric("📊 전체 수주 건수", f"{len(df_result)} 건")
        col2.metric("🟢 정상 매핑 건수", f"{normal_cnt} 건")
        col3.metric("🔴 검토 필요 건수", f"{review_cnt} 건")
        
        st.subheader("📋 수주 매핑 최종 결과 목록")
        
        # 스타일 지정 함수 (검토필요 항목 강조)
        def highlight_review(row):
            if '검토필요' in str(row['매핑상태']):
                return ['background-color: #ffe6e6; color: #cc0000; font-weight: bold;'] * len(row)
            return [''] * len(row)
            
        styled_df = df_result.style.apply(highlight_review, axis=1)
        
        st.dataframe(styled_df, use_container_width=True)
        
        # 엑셀 다운로드 버튼
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
            
        st.download_button(
            label="📥 결과 엑셀 파일 다운로드",
            data=output.getvalue(),
            file_name=f"올리브영_수주업로드_완료_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👈 좌측 사이드바에 '일일재고' 및 '납품확인서 목록' 엑셀 파일을 업로드해주세요.")
