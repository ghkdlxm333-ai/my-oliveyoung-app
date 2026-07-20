import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="올리브영 자동 매핑 솔루션", layout="wide")

st.title("📦 올리브영 발주 - WMS 재고 자동 매핑 시스템")
st.caption("1년 6개월 이상 잔여 유효기간 & 박스 입수량 이상 재고만 자동 계산하여 단일 LOT를 매핑합니다.")

# ---------------------------------------------------------
# Sidebar File Upload
# ---------------------------------------------------------
st.sidebar.header("📁 엑셀 파일 업로드")
stock_file = st.sidebar.file_uploader("1. WMS 일일재고 파일 (.xlsx)", type=["xlsx"])
order_file = st.sidebar.file_uploader("2. 올리브영 납품확인서 목록 (.xlsx)", type=["xlsx"])

# Delivery Center Code Mapping
CENTER_MAP = {
    '[LA02] 양지센터': '86101126',
    '[L002] 부곡센터': '86100086',
    '[L003] 중부센터': '86100118',
    '[L001] 수도권센터': '86100000' # 필요시 확장 가능
}

# ---------------------------------------------------------
# Processing Logic
# ---------------------------------------------------------
if stock_file and order_file:
    try:
        # Load Raw Data
        df_stock_raw = pd.read_excel(stock_file, header=None)
        df_order_raw = pd.read_excel(order_file, header=None)
        
        # Parse Stock Data Header (Row 0/1 area)
        stock_header_idx = 0
        for i, row in df_stock_raw.iterrows():
            if '상품코드' in row.values or '상품' in row.values:
                stock_header_idx = i
                break
        
        df_stock = pd.read_excel(stock_file, header=stock_header_idx)
        df_stock.columns = [str(c).strip() for c in df_stock.columns]
        
        # Parse Order Data Header
        order_header_idx = 0
        for i, row in df_order_raw.iterrows():
            if '상품코드' in row.values:
                order_header_idx = i
                break
        
        df_order = pd.read_excel(order_file, header=order_header_idx)
        df_order.columns = [str(c).strip() for c in df_order.columns]
        
        # Preprocessing Barcodes & Clean strings (Force clean string without .0)
        df_stock['상품코드_str'] = df_stock['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_stock['상품바코드_str'] = df_stock['상품바코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        df_order['상품코드_str'] = df_order['상품코드'].astype(str).str.replace('.0', '', regex=False).str.strip()
        
        # Parse Dates
        df_stock['유효일자'] = pd.to_datetime(df_stock['유효일자'], errors='coerce')
        df_order['입고예정일'] = pd.to_datetime(df_order['입고예정일'], errors='coerce')
        
        # Preprocessing Numbers
        df_stock['합계수량'] = pd.to_numeric(df_stock['합계수량'], errors='coerce').fillna(0)
        df_stock['입수량(BOX)'] = pd.to_numeric(df_stock['입수량(BOX)'], errors='coerce').fillna(1)
        
        df_order['발주수량\n(EA)'] = pd.to_numeric(df_order['발주수량\n(EA)'], errors='coerce').fillna(0)
        df_order['BOX\n입수'] = pd.to_numeric(df_order['BOX\n입수'], errors='coerce').fillna(1)
        
        mapped_rows = []
        review_rows = []
        
        today = datetime.now()
        
        # Process Mapping Loop
        for idx, row in df_order.iterrows():
            barcode = row['상품코드_str']
            order_qty = row['발주수량\n(EA)']
            box_in_qty = row['BOX\n입수']
            order_date = row['입고예정일'] if pd.notnull(row['입고예정일']) else today
            center_name = str(row.get('센터', ''))
            
            # Find Matching Stock (Match by Barcode or ME Code)
            stock_sub = df_stock[
                (df_stock['상품바코드_str'] == barcode) | 
                (df_stock['상품코드_str'] == barcode)
            ].copy()
            
            # 🛑 [제약 1] 유효기간 1년 6개월(547.5일) 이상 남은 재고만 선별
            # 기준일: 오늘 또는 입고예정일
            min_valid_date = order_date + timedelta(days=547)
            valid_stock = stock_sub[stock_sub['유효일자'] >= min_valid_date].copy()
            
            # 🛑 [제약 2] 재고 수량이 박스 입수량보다 적은 재고(단수) 차단
            # 재고 수량이 입수량(BOX) 이상인 것만 허용
            valid_stock = valid_stock[valid_stock['합계수량'] >= valid_stock['입수량(BOX)']]
            
            # FEFO 정렬 (유효일자 임박순)
            valid_stock = valid_stock.sort_values(by='유효일자', ascending=True)
            
            selected_lot = ""
            selected_exp_date = ""
            status_msg = "정상"
            me_code = ""
            
            if not valid_stock.empty:
                me_code = valid_stock.iloc[0]['상품코드']
                
                # FEFO 첫 번째 LOT 수량이 충분한가?
                fefo_match = valid_stock[valid_stock['합계수량'] >= order_qty]
                
                if not fefo_match.empty:
                    # 단일 LOT로 전량 커버 가능
                    best_lot = fefo_match.iloc[0]
                    selected_lot = best_lot['화주LOT']
                    selected_exp_date = best_lot['유효일자'].strftime('%Y-%m-%d') if pd.notnull(best_lot['유효일자']) else ""
                else:
                    # 유효재고 전체 합산하여 총량 확인
                    total_avail_qty = valid_stock['합계수량'].sum()
                    if total_avail_qty >= order_qty:
                        status_msg = "🔴 경고: 단일 LOT 수량 부족 (LOT 분할 필요)"
                    else:
                        status_msg = "🔴 경고: 출고가능 유효재고 부족 (1.5년 미만/단수제외 후 부족)"
            else:
                status_msg = "🔴 경고: 출고 가능한 재고 없음 (1년 6개월 미만 또는 박스 미만 재고)"
                
            # 배송코드 매핑
            shipping_code = CENTER_MAP.get(center_name, '')
            
            # 결과 행 구성 (서식(수주업로드) 형태)
            out_row = {
                '발주처코드': '86100000',
                '입고예정일': order_date.strftime('%Y-%m-%d') if pd.notnull(order_date) else '',
                '배송코드': shipping_code,
                'ORDER #': row.get('입고전표', ''),
                '상품명': row.get('상품명', ''),
                '바코드': barcode,
                'MECODE': me_code,
                '수량': order_qty,
                '발주원가': row.get('원단가', 0),
                '발주금액': row.get('원가금액', 0),
                'LOT': selected_lot,
                '유효일자': selected_exp_date,
                '매핑상태': status_msg
            }
            
            if selected_lot != "" and status_msg == "정상":
                mapped_rows.append(out_row)
            else:
                review_rows.append(out_row)
                
        df_success = pd.DataFrame(mapped_rows)
        df_review = pd.DataFrame(review_rows)
        
        # ---------------------------------------------------------
        # Display Results
        # ---------------------------------------------------------
        col1, col2 = st.columns(2)
        col1.metric("🟢 정상 매핑 성공 건수", len(df_success))
        col2.metric("🔴 예외/검토 필요 건수", len(df_review))
        
        tab1, tab2 = st.tabs(["🟢 자동 매핑 완료 목록", "🔴 수기 검토 필요 목록 (제약 미달 / 재고 쪼개짐)"])
        
        with tab1:
            st.dataframe(df_success, use_container_width=True)
            if not df_success.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_success.to_excel(writer, index=False, sheet_name='서식(수주업로드)')
                st.download_button(
                    label="📥 완료된 수주업로드 엑셀 다운로드",
                    data=output.getvalue(),
                    file_name=f"올리브영_수주업로드_완료_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        with tab2:
            st.dataframe(df_review, use_container_width=True)
            if not df_review.empty:
                st.warning("⚠️ 위 건들은 1년 6개월 미만 유효일자 재고, 박스 입수량 미만 재고, 또는 단일 LOT 수량 부족 건입니다. 재고를 수기 확인해주세요.")

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👈 좌측 사이드바에 '일일재고'와 '납품확인서 목록' 엑셀 파일을 올려주세요.")
