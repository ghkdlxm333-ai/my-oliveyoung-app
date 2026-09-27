from datetime import datetime
from io import BytesIO
import re
import pandas as pd
import streamlit as st

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="올리브영 출고 LOT자동 매핑 시스템(그레이스3PL)",
    page_icon="🫒",
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
    - H열(7) '정상창고' 재고만 선택
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
        warehouse = clean_str(row.iloc[7]) if len(row) > 7 else ""
        avail_qty = parse_number(row.iloc[10]) if len(row) > 10 else 0

        if not wms_sku and not barcode and not product_name:
            continue
        if avail_qty <= 0:
            continue
        # H열 정상창고 재고만 출고가능 재고로 인정
        if warehouse != "정상창고":
            continue

        exp_date = pd.to_datetime(exp_date_str, errors="coerce")

        parsed_rows.append({
            "WMS상품코드": wms_sku,
            "WMS상품명": product_name,
            "LOT": lot,
            "유통기한": exp_date,
            "바코드": barcode,
            "창고": warehouse,
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


def parse_grace_delivery_confirmation(file_obj):
    """
    그레이스 3PL 출고확인 파일 파싱 (센터별 시트 자동 처리)
    """
    xls = pd.ExcelFile(BytesIO(file_obj.getvalue()))
    grace_rows = []
    
    for sheet in xls.sheet_names:
        try:
            df_sheet = pd.read_excel(xls, sheet_name=sheet)
            if df_sheet.empty:
                continue
            
            barcode_col = None
            qty_col = None
            exp_col = None
            lot_col = None
            center_col = None
            
            for col in df_sheet.columns:
                c_str = str(col).replace("\n", "").strip()
                if "바코드" in c_str: barcode_col = col
                elif "출고수량" in c_str or "수량" in c_str: qty_col = col
                elif "유통기한" in c_str: exp_col = col
                elif "제조일자" in c_str or "LOT" in c_str: lot_col = col
                elif "센터" in c_str: center_col = col
                
            for _, row in df_sheet.iterrows():
                raw_barcode = row.get(barcode_col, "") if barcode_col else ""
                if pd.isna(raw_barcode) or not str(raw_barcode).strip():
                    continue
                try:
                    barcode = str(int(raw_barcode))
                except Exception:
                    barcode = str(raw_barcode).strip()
                    if barcode.endswith(".0"): barcode = barcode[:-2]
                    
                raw_center = str(row.get(center_col, sheet)).strip() if center_col else str(sheet).strip()
                if "양지온라인" in raw_center: center = "양지온라인센터"
                elif "양지" in raw_center: center = "양지센터"
                elif "경산" in raw_center: center = "경산센터"
                else: center = raw_center
                
                out_qty = parse_number(row.get(qty_col, 0)) if qty_col else 0
                
                exp_val = pd.to_datetime(row.get(exp_col, None), errors="coerce") if exp_col else pd.NaT
                exp_str = exp_val.strftime("%Y-%m-%d") if pd.notna(exp_val) else ""
                
                lot_val = str(row.get(lot_col, "")).strip() if lot_col and pd.notna(row.get(lot_col)) else ""
                if lot_val.endswith(".0"): lot_val = lot_val[:-2]
                
                grace_rows.append({
                    "센터": center,
                    "바코드": barcode,
                    "출고수량": out_qty,
                    "유통기한": exp_str,
                    "LOT": lot_val
                })
        except Exception:
            continue
            
    df_grace = pd.DataFrame(grace_rows)
    if df_grace.empty:
        return pd.DataFrame()
        
    grouped = df_grace.groupby(["센터", "바코드"]).agg({
        "출고수량": "sum",
        "유통기한": lambda x: " / ".join(sorted(list(set([s for s in x if s])))),
        "LOT": lambda x: " / ".join(sorted(list(set([s for s in x if s]))))
    }).reset_index()
    
    return grouped

# ============================================================
# AUTO MATCHING & ALLOCATION LOGIC
# ============================================================
def allocate_inventory(delivery_df, wms_df, grace_confirm_df=None, min_days=MIN_SHELF_LIFE_DAYS):
    wms_work = wms_df.copy()
    allocated_results = []

    # 납품확인서 총 요구량 수집
    delivery_totals = delivery_df.copy()
    delivery_totals["센터명"] = delivery_totals["센터"].apply(
        lambda c: "양지온라인센터" if "양지온라인" in str(c) else ("양지센터" if "양지" in str(c) else ("경산센터" if "경산" in str(c) else str(c)))
    )
    delivery_totals_grp = delivery_totals.groupby(["센터명", "상품코드"])["출고수량"].sum().to_dict()

    for idx, d_row in delivery_df.iterrows():
        barcode = d_row.get("상품코드", "")
        req_qty = d_row.get("출고수량", 0)
        box_in = d_row.get("BOX입수", 1)
        oy_product_name = str(d_row.get("상품명", "")).strip()
        target_date = d_row.get("입고예정일", pd.Timestamp.now())

        if pd.isna(target_date):
            target_date = pd.Timestamp.now()

        raw_center = str(d_row.get("센터", "")).strip()
        if "양지온라인" in raw_center: center = "양지온라인센터"
        elif "양지" in raw_center: center = "양지센터"
        elif "경산" in raw_center: center = "경산센터"
        else: center = raw_center if raw_center else "기타센터"

        # 바코드/SKU 매칭
        matched_wms = wms_work[(wms_work["바코드"] == barcode) | (wms_work["WMS상품코드"] == barcode)].copy()

        if not matched_wms.empty:
            matched_wms["잔여일수"] = (matched_wms["유통기한"] - target_date).dt.days
            valid_wms = matched_wms[matched_wms["잔여일수"] >= min_days].sort_values("유통기한")
            invalid_wms = matched_wms[matched_wms["잔여일수"] < min_days].sort_values("유통기한", ascending=False)
        else:
            valid_wms = pd.DataFrame()
            invalid_wms = pd.DataFrame()

        picked_lots = []
        status = "NORMAL"

        if valid_wms.empty:
            if not invalid_wms.empty:
                status = "INVALID_SHELF_LIFE"
                for _, inv_row in invalid_wms.iterrows():
                    exp_str = inv_row["유통기한"].strftime("%Y-%m-%d") if pd.notna(inv_row["유통기한"]) else "N/A"
                    picked_lots.append({
                        "LOT": inv_row["LOT"],
                        "유통기한": exp_str,
                        "창고재고수량": int(inv_row["가용재고"]),
                        "WMS코드": inv_row["WMS상품코드"]
                    })
            else:
                status = "NO_STOCK"
        else:
            # 1. 단일 LOT 탐색
            single_sufficient = valid_wms[valid_wms["가용재고_남은수량"] >= req_qty]

            if not single_sufficient.empty:
                s_row = single_sufficient.iloc[0]
                w_idx = s_row.name
                wms_work.loc[w_idx, "가용재고_남은수량"] -= req_qty

                exp_str = s_row["유통기한"].strftime("%Y-%m-%d") if pd.notna(s_row["유통기한"]) else "N/A"
                picked_lots.append({
                    "LOT": s_row["LOT"],
                    "유통기한": exp_str,
                    "창고재고수량": int(s_row["가용재고"]),
                    "WMS코드": s_row["WMS상품코드"]
                })
                status = "NORMAL"
            else:
                # 2. LOT 분할
                remaining_to_pick = req_qty
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
                        "창고재고수량": int(w_row["가용재고"]),
                        "WMS코드": w_row["WMS상품코드"]
                    })

                status = "SHORTAGE" if remaining_to_pick > 0 else "SPLIT"

        wms_code_display = picked_lots[0]["WMS코드"] if picked_lots else (matched_wms["WMS상품코드"].iloc[0] if not matched_wms.empty else "-")

        lot_summary = [f"{p['LOT']}({p['창고재고수량']})" for p in picked_lots]
        exp_summary = list(set([p["유통기한"] for p in picked_lots]))
        sys_exp_str = " / ".join(exp_summary) if exp_summary else "-"
        is_split = len(picked_lots) > 1

        status_flag = "🟢 정상출고"
        if status == "NO_STOCK":
            status_flag = "🔴 재고없음"
        elif status == "INVALID_SHELF_LIFE":
            status_flag = "⛔ [출고불가] 유통기한 1년6개월 미만"
        elif status == "SHORTAGE":
            status_flag = f"🟡 수량부족"
        elif is_split or status == "SPLIT":
            status_flag = "⚠️ LOT 분할"

        # 그레이스 3PL 교차 검증
        grace_check = "-"
        if grace_confirm_df is not None and not grace_confirm_df.empty:
            g_match = grace_confirm_df[(grace_confirm_df["센터"] == center) & (grace_confirm_df["바코드"] == barcode)]
            if g_match.empty:
                grace_check = "⚠️ 3PL 미확인"
            else:
                g_qty = int(g_match.iloc[0]["출고수량"])
                g_exp = str(g_match.iloc[0]["유통기한"])
                total_req_for_barcode = delivery_totals_grp.get((center, barcode), req_qty)
                
                qty_matched = (int(req_qty) == g_qty) or (int(total_req_for_barcode) == g_qty)
                exp_matched = (sys_exp_str == g_exp) or (sys_exp_str == "-") or (g_exp == "")
                
                diffs = []
                if not qty_matched:
                    diffs.append(f"수량차이(시스템:{int(req_qty)} / 3PL:{g_qty})")
                if not exp_matched:
                    diffs.append(f"유통기한차이(시스템:{sys_exp_str} / 3PL:{g_exp})")
                    
                if not diffs:
                    grace_check = "✅ 일치"
                else:
                    grace_check = f"❌ 불일치 ({', '.join(diffs)})"

        allocated_results.append({
            "납품센터": center,
            "바코드": barcode,
            "상품코드": wms_code_display,
            "상품명": oy_product_name,
            "BOX입수": int(box_in),
            "출고수량": int(req_qty),
            "LOT (현재재고)": " / ".join(lot_summary) if lot_summary else "-",
            "유통기한": sys_exp_str,
            "매핑상태": status_flag,
            "3PL 검증": grace_check,
        })

    return pd.DataFrame(allocated_results), wms_work

# ============================================================
# ROW HIGHLIGHTING FUNCTION
# ============================================================
def style_dataframe(df):
    """
    - 재고없음: 빨강색 (#FFE6E6)
    - LOT 분할: 파랑색 (#E6F2FF)
    - 유통기한 미달: 노랑색 (#FFF9E6)
    """
    def highlight_rows(row):
        status = str(row.get("매핑상태", ""))
        if "재고없음" in status:
            return ["background-color: #ffe6e6"] * len(row)
        elif "LOT 분할" in status:
            return ["background-color: #e6f2ff"] * len(row)
        elif "출고불가" in status or "유통기한" in status:
            return ["background-color: #fff9e6"] * len(row)
        return [""] * len(row)

    return df.style.apply(highlight_rows, axis=1)

# ============================================================
# SIDEBAR WITH OLIVE YOUNG LOGO (최신 Streamlit 파라미터 적용)
# ============================================================
st.sidebar.image(
    "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Olive_Young_Logo.svg/2560px-Olive_Young_Logo.svg.png",
    use_container_width=True,
)
st.sidebar.markdown("---")
st.sidebar.header("⚙️ 설정 옵션")
min_months = st.sidebar.slider("올리브영 납품 가능 최소 유통기한 (개월)", 6, 24, 18, 1)
min_days_limit = int(min_months * 30.4375)

# ============================================================
# STREAMLIT UI MAIN
# ============================================================
st.title("🫒 올리브영 출고 LOT자동 매핑 시스템(그레이스3PL)")
st.caption("그레이스 WMS 정상창고 재고와 올리브영 납품확인서를 바코드 기반으로 자동 매핑하고, 3PL 출고파일과 교차 검증합니다.")

# 3개 파일 업로드 영역
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. 올리브영 납품확인서")
    delivery_file = st.file_uploader(
        "올리브영 납품확인서 엑셀(.xlsx)",
        key="delivery",
        help="올리브영 시스템에서 다운로드받은 납품확인서 목록 파일입니다."
    )

with col2:
    st.subheader("2. 그레이스 WMS 재고")
    wms_file = st.file_uploader(
        "그레이스 WMS 재고파일 (.xlsx, .xls)",
        key="wms",
        help="※ .xls 업로드 오류 시 Microsoft Excel 2007 버전으로 저장 후 업로드하세요."
    )

with col3:
    st.subheader("3. 그레이스 3PL 출고확인 (선택)")
    grace_file = st.file_uploader(
        "그레이스 3PL 출고확인 파일 (.xlsx)",
        key="grace_confirm",
        help="그레이스 3PL에서 송부받은 센터별 시트 분리 출고확인 파일입니다."
    )

# 자동 매핑 및 교차 검증 실행
if wms_file and delivery_file:
    try:
        wms_raw = load_uploaded_file(wms_file)
        delivery_raw = load_uploaded_file(delivery_file)

        wms_df = parse_grace_wms(wms_raw)
        delivery_df = parse_oliveyoung_delivery(delivery_raw)

        grace_confirm_df = None
        if grace_file is not None:
            grace_confirm_df = parse_grace_delivery_confirmation(grace_file)

        result_df, updated_wms = allocate_inventory(delivery_df, wms_df, grace_confirm_df, min_days=min_days_limit)

        st.markdown("---")
        st.subheader("📊 전체 출고 매핑 & 3PL 교차 검증 결과")

        # KPI 요약
        k1, k2, k3, k4 = st.columns(4)
        normal_cnt = sum(result_df["매핑상태"].str.contains("정상출고"))
        split_cnt = sum(result_df["매핑상태"].str.contains("LOT 분할"))
        invalid_cnt = sum(result_df["매핑상태"].str.contains("출고불가"))
        shortage_cnt = sum(result_df["매핑상태"].str.contains("부족|재고없음"))

        k1.metric("🟢 정상 출고 가능", f"{normal_cnt} 건")
        k2.metric("⚠️ LOT 분할 항목", f"{split_cnt} 건")
        k3.metric("⛔ 유통기한 부적합", f"{invalid_cnt} 건")
        k4.metric("🔴 재고 부족/없음", f"{shortage_cnt} 건")

        # 필터 기능
        status_filter = st.multiselect("상태별 필터링", result_df["매핑상태"].unique(), default=result_df["매핑상태"].unique())
        filtered_result = result_df[result_df["매핑상태"].isin(status_filter)]

        styled_result = style_dataframe(filtered_result)
        st.dataframe(styled_result, use_container_width=True, height=520)

        # Excel 다운로드
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            filtered_result.to_excel(writer, index=False, sheet_name="올리브영_3PL검증결과")
        excel_data = output.getvalue()

        st.download_button(
            label="📥 매핑 & 교차 검증 결과 엑셀 다운로드",
            data=excel_data,
            file_name=f"올리브영_3PL검증결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
else:
    st.info("💡 1번 올리브영 납품확인서와 2번 그레이스 WMS 재고파일을 올려주시면 즉시 매핑 표가 나타납니다. 3번 3PL 출고확인 파일을 함께 올리시면 3PL 실시간 교차 검증 결과까지 표시됩니다.")

st.markdown("---")

# ============================================================
# PROGRAM USAGE GUIDE & ERROR HANDLING
# ============================================================
with st.expander("📖 올리브영 자동 출고 매핑 & 3PL 검증기 사용 방법 및 색상 안내", expanded=False):
    st.markdown("""
    ### 1. 주요 기능 및 색상 구분 안내
    * **🔴 빨간색 배경**: 재고가 아예 없는 항목 (`🔴 재고없음`)
    * **⚠️ 파란색 배경**: 단일 LOT 수량이 부족하여 나누어 출고되는 항목 (`⚠️ LOT 분할`)
    * **⛔ 노란색 배경**: 잔여 유통기한 1년 6개월(547일) 미만으로 올리브영 입고 불가한 항목 (`⛔ [출고불가] 유통기한 1년6개월 미만`)
    * **`3PL 검증` 컬럼**: 3번 파일(그레이스 3PL 출고확인)을 함께 첨부하면 **센터별 시트(양지/경산 등)**를 자동 분석하여 출고수량 및 유통기한의 일치 여부를 `✅ 일치` / `❌ 불일치`로 표시합니다.
    * **`LOT (현재재고)` 표기**: `LOT명(숫자)`으로 표기되어 각 LOT별 보유 잔여 재고 확인이 용이합니다.

    ---

    ### 2. `.xls` 파일 업로드 오류 해결책
    * WMS에서 다운로드(Export To Excel) 시 파일 형식을 **`Microsoft Excel 2007`**로 선택 후 저장하시거나 확장자를 **`.xlsx`**로 수정하여 업로드하세요.
    """)
