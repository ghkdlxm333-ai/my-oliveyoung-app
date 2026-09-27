from datetime import datetime
from io import BytesIO
import re
import pandas as pd
import streamlit as st

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="올리브영 자동 출고 & LOT 매핑 (그레이스 3PL)",
    page_icon="📦",
    layout="wide",
)

MIN_SHELF_LIFE_DAYS = 547  # 올리브영 납품 기준 잔여 유통기한 (1년 6개월 = 547일)

# ============================================================
# DATA CLEANING HELPERS
# ============================================================
def clean_str(val):
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str

def parse_number(val):
    try:
        if pd.isna(val):
            return 0
        cleaned = re.sub(r"[^\d.-]", "", str(val))
        return float(cleaned) if cleaned else 0
    except Exception:
        return 0

# ============================================================
# AUTOMATIC FILE PARSER
# ============================================================
def load_uploaded_file(uploaded_file):
    raw_bytes = uploaded_file.getvalue()
    
    # 1. Excel (openpyxl / xlrd) 시도
    try:
        return pd.read_excel(BytesIO(raw_bytes), header=None)
    except Exception:
        pass

    # 2. HTML Table 형태의 .xls 엑셀 파일 시도
    try:
        html_dfs = pd.read_html(BytesIO(raw_bytes))
        if html_dfs:
            return html_dfs[0]
    except Exception:
        pass

    # 3. CSV 시도
    for encoding in ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]:
        try:
            return pd.read_csv(BytesIO(raw_bytes), encoding=encoding, header=None)
        except Exception:
            continue

    raise ValueError("파일을 표준 형태로 읽을 수 없습니다. 확장자(.xlsx)를 확인해 주세요.")

# ============================================================
# WMS & DELIVERY FILE PARSERS
# ============================================================
def parse_grace_wms(df_raw):
    """
    그레이스 3PL WMS 재고 현황 파싱
    - B열(1): 상품코드
    - C열(2): 상품명
    - D열(3): LOT
    - E열(4): 유통기한
    - G열(6): 바코드
    - K열(10): 가용재고
    """
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str))
        if "상품코드" in row_str and "LOT" in row_str:
            header_idx = idx
            break

    if header_idx is None:
        header_idx = 1

    df = df_raw.iloc[header_idx + 1:].copy()
    
    parsed_rows = []
    for _, row in df.iterrows():
        wms_sku = clean_str(row.iloc[1]) if len(row) > 1 else ""
        product_name = clean_str(row.iloc[2]) if len(row) > 2 else ""
        lot = clean_str(row.iloc[3]) if len(row) > 3 else "NO_LOT"
        exp_date_str = clean_str(row.iloc[4]) if len(row) > 4 else ""
        barcode = clean_str(row.iloc[6]) if len(row) > 6 else ""
        avail_qty = parse_number(row.iloc[10]) if len(row) > 10 else 0

        if not wms_sku and not barcode and not product_name:
            continue
        if avail_qty <= 0:
            continue

        exp_date = pd.to_datetime(exp_date_str, errors="coerce")

        parsed_rows.append({
            "WMS상품코드": wms_sku,
            "WMS상품명": product_name,
            "LOT": lot,
            "유통기한": exp_date,
            "바코드": barcode,
            "가용재고": avail_qty,
            "가용재고_남은수량": avail_qty,
        })

    return pd.DataFrame(parsed_rows)


def parse_oliveyoung_delivery(df_raw):
    """
    올리브영 납품확인서 목록 파싱
    """
    if "상품코드" in df_raw.columns:
        df = df_raw.copy()
    else:
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.dropna().astype(str))
            if "상품코드" in row_str and "발주수량" in row_str:
                header_idx = idx
                break

        df = df_raw.copy()
        df.columns = df.iloc[header_idx].map(clean_str)
        df = df.iloc[header_idx + 1:].reset_index(drop=True)

    df = df[df["상품코드"].notna() & (df["상품코드"] != "")].copy()
    df["상품코드"] = df["상품코드"].apply(clean_str)
    
    req_col = [c for c in df.columns if "발주수량" in c]
    box_col = [c for c in df.columns if "BOX" in c and "입수" in c]
    
    df["출고수량"] = df[req_col[0]].apply(parse_number) if req_col else 0
    df["BOX입수"] = df[box_col[0]].apply(parse_number) if box_col else 1
    df["입고예정일"] = pd.to_datetime(df["입고예정일"], errors="coerce")

    return df

# ============================================================
# AUTO MATCHING & ALLOCATION LOGIC
# ============================================================
def allocate_inventory(delivery_df, wms_df, min_days=MIN_SHELF_LIFE_DAYS):
    """
    올영 바코드 = WMS G열(바코드) 매칭
    상품코드 = WMS B열, 상품명 = WMS C열 적용
    """
    wms_work = wms_df.copy()
    allocated_results = []

    for idx, d_row in delivery_df.iterrows():
        barcode = d_row.get("상품코드", "")
        req_qty = d_row.get("출고수량", 0)
        box_in = d_row.get("BOX입수", 1)
        target_date = d_row.get("입고예정일", pd.Timestamp.now())

        if pd.isna(target_date):
            target_date = pd.Timestamp.now()

        # 올영 바코드 = WMS G열(바코드) 또는 B열 매칭
        matched_wms = wms_work[(wms_work["바코드"] == barcode) | (wms_work["WMS상품코드"] == barcode)].copy()

        if not matched_wms.empty:
            matched_wms["잔여일수"] = (matched_wms["유통기한"] - target_date).dt.days
            # FEFO 유통기한 빠른 순 정렬
            valid_wms = matched_wms[matched_wms["잔여일수"] >= min_days].sort_values("유통기한")
            invalid_wms = matched_wms[matched_wms["잔여일수"] < min_days]
        else:
            valid_wms = pd.DataFrame()
            invalid_wms = pd.DataFrame()

        remaining_to_pick = req_qty
        picked_lots = []
        status = "NORMAL"

        if valid_wms.empty:
            if not invalid_wms.empty:
                status = "INVALID_SHELF_LIFE"
            else:
                status = "NO_STOCK"
        else:
            for w_idx, w_row in valid_wms.iterrows():
                if remaining_to_pick <= 0:
                    break
                
                avail = w_row["가용재고_남은수량"]
                if avail <= 0:
                    continue

                take = min(avail, remaining_to_pick)
                wms_work.loc[w_idx, "가용재고_남은수량"] -= take
                remaining_to_pick -= take

                exp_str = w_row["유통기한"].strftime("%Y-%m-%d") if pd.notna(w_row["유통기한"]) else "N/A"
                picked_lots.append({
                    "LOT": w_row["LOT"],
                    "유통기한": exp_str,
                    "수량": take,
                    "WMS코드": w_row["WMS상품코드"],
                    "WMS상품명": w_row["WMS상품명"]
                })

            if remaining_to_pick > 0:
                status = "SHORTAGE"

        # WMS B열 상품코드 및 WMS C열 상품명 매핑
        wms_code_display = picked_lots[0]["WMS코드"] if picked_lots else (matched_wms["WMS상품코드"].iloc[0] if not matched_wms.empty else "-")
        wms_name_display = picked_lots[0]["WMS상품명"] if picked_lots else (matched_wms["WMS상품명"].iloc[0] if not matched_wms.empty else "-")

        lot_summary = [f"{p['LOT']}({int(p['수량'])}개)" for p in picked_lots]
        exp_summary = list(set([p["유통기한"] for p in picked_lots]))
        is_split = len(picked_lots) > 1

        status_flag = "🟢 정상출고"
        if status == "NO_STOCK":
            status_flag = "🔴 재고없음"
        elif status == "INVALID_SHELF_LIFE":
            status_flag = "⛔ [출고불가] 유통기한 1년6개월 미만"
        elif status == "SHORTAGE":
            status_flag = f"🟡 수량부족 ({int(remaining_to_pick)}EA)"
        elif is_split:
            status_flag = "⚠️ [주의] LOT 쪼개짐"

        box_check = "OK"
        if box_in > 0 and (req_qty % box_in != 0):
            box_check = f"❌ 박스미달 (입수:{int(box_in)})"

        allocated_results.append({
            "바코드": barcode,
            "상품코드": wms_code_display,
            "상품명": wms_name_display,
            "BOX입수": int(box_in),
            "출고수량": int(req_qty),
            "LOT": " / ".join(lot_summary) if lot_summary else "-",
            "유통기한": " / ".join(exp_summary) if exp_summary else "-",
            "매핑상태": status_flag,
            "박스입수체크": box_check,
        })

    return pd.DataFrame(allocated_results), wms_work

# ============================================================
# STREAMLIT UI
# ============================================================
st.title("📦 올리브영 자동 출고 & LOT 매핑 시스템 (그레이스 3PL)")
st.caption("그레이스 WMS 재고와 올리브영 납품확인서를 바코드 기반으로 자동 매핑하여 LOT 및 유통기한을 배정합니다.")

st.sidebar.header("⚙️ 설정 옵션")
min_months = st.sidebar.slider("올리브영 납품 가능 최소 유통기한 (개월)", 6, 24, 18, 1)
min_days_limit = int(min_months * 30.4375)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. 올리브영 납품확인서 목록 파일")
    st.info("💡 올리브영 납품확인서 엑셀 다운로드 파일을 업로드해주세요.")
    delivery_file = st.file_uploader("납품확인서 목록 파일 (.xlsx)", key="delivery")

with col2:
    st.subheader("2. 그레이스(wms) 재고파일")
    st.warning("⚠️ 확장자가 .xls인 경우 업로드 시 오류가 발생할 수 있습니다. 확장자를 .xlsx로 변경하시거나, Microsoft Excel 2007 버전으로 다운받은 파일을 업로드해 주세요.")
    wms_file = st.file_uploader("WMS 재고파일 (.xlsx, .xls)", key="wms")

if wms_file and delivery_file:
    try:
        wms_raw = load_uploaded_file(wms_file)
        delivery_raw = load_uploaded_file(delivery_file)

        wms_df = parse_grace_wms(wms_raw)
        delivery_df = parse_oliveyoung_delivery(delivery_raw)

        st.success(f"✅ 파일 읽기 성공! (WMS 가용재고: {len(wms_df)}건 / 올리브영 납품항목: {len(delivery_df)}건)")

        if st.button("🚀 LOT 자동 매핑 및 출고 검증 실행", type="primary"):
            result_df, updated_wms = allocate_inventory(delivery_df, wms_df, min_days=min_days_limit)

            st.subheader("📊 자동 매핑 및 출고 검증 결과")

            # KPI 요약
            k1, k2, k3, k4 = st.columns(4)
            normal_cnt = sum(result_df["매핑상태"].str.contains("정상출고"))
            split_cnt = sum(result_df["매핑상태"].str.contains("LOT 쪼개짐"))
            invalid_cnt = sum(result_df["매핑상태"].str.contains("출고불가"))
            shortage_cnt = sum(result_df["매핑상태"].str.contains("부족|재고없음"))

            k1.metric("🟢 정상 출고 가능", f"{normal_cnt} 건")
            k2.metric("⚠️ LOT 쪼개짐 항목", f"{split_cnt} 건")
            k3.metric("⛔ 유통기한 부적합", f"{invalid_cnt} 건")
            k4.metric("🔴 재고 부족/없음", f"{shortage_cnt} 건")

            # 필터 기능
            status_filter = st.multiselect("상태별 필터링", result_df["매핑상태"].unique(), default=result_df["매핑상태"].unique())
            filtered_result = result_df[result_df["매핑상태"].isin(status_filter)]

            # [바코드, 상품코드, 상품명, BOX입수, 출고수량, LOT, 유통기한] 순서로 출력
            st.dataframe(filtered_result, use_container_width=True, height=500)

            # Excel 다운로드
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                filtered_result.to_excel(writer, index=False, sheet_name="올리브영_LOT매핑결과")
            excel_data = output.getvalue()

            st.download_button(
                label="📥 매핑 결과 엑셀 다운로드",
                data=excel_data,
                file_name=f"올리브영_LOT매핑결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")

st.markdown("---")

# ============================================================
# PROGRAM USAGE GUIDE & ERROR HANDLING
# ============================================================
with st.expander("📖 올리브영 자동 출고 매핑 프로그램 사용 방법 및 오류 해결 안내", expanded=True):
    st.markdown("""
    ### 1. 사용 방법
    1. **올리브영 납품확인서 목록 파일**과 **그레이스 WMS 재고 파일**을 각각 상단 업로드 영역에 첨부합니다.
    2. **[🚀 LOT 자동 매핑 및 출고 검증 실행]** 버튼을 클릭합니다.
    3. 결과 표에서 **`매핑상태`** 및 **`박스입수체크`** 항목을 확인합니다.
    4. **[📥 매핑 결과 엑셀 다운로드]** 버튼을 눌러 결과 보고서를 저장합니다.

    ---

    ### 2. 매핑 및 검증 기준 안내
    * **바코드 매칭**: 올리브영 납품확인서의 바코드와 그레이스 WMS G열(바코드)을 자동 매칭합니다.
    * **상품코드/상품명**: 그레이스 WMS B열(상품코드) 및 C열(상품명) 정보를 가져와 표시합니다.
    * **유통기한 기준 (1년 6개월)**: 입고예정일 기준 잔여 유통기한이 **547일(18개월) 미만**인 재고는 출고 대상에서 자동 제외되며 `⛔ [출고불가]`로 표기됩니다.
    * **LOT 쪼개짐 태그**: single LOT 재고 부족으로 한 납품 품목에 2개 이상의 LOT가 나누어 매핑된 경우 `⚠️ [주의] LOT 쪼개짐`으로 표기됩니다.
    * **박스 입수 체크**: 출고 수량이 박스 입수 단위로 깔끔하게 떨어지지 않을 경우 `❌ 박스미달` 경고가 발생합니다.

    ---

    ### 3. 파일 업로드 오류 해결 방법 (`application/vnd.ms-excel` 등)
    * **원인**: WMS 시스템에서 내려받은 구형 `.xls` 파일은 보안/형식 문제로 웹 업로드가 거부될 수 있습니다.
    * **해결책 (택 1)**:
      1. WMS에서 **Export To Excel(저장)** 시 하단 **파일 형식(T)**을 **`Microsoft Excel 2007`**로 선택하고 저장 후 업로드합니다.
      2. 다운로드받은 파일 이름 끝의 확장자를 **`.xls`에서 `.xlsx`로 변경**한 후 업로드합니다.
    """)
