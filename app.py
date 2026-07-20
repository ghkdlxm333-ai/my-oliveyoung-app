import io
import pandas as pd
import requests
import streamlit as st

# ==========================================
# 1. 페이지 기본 설정 및 CSS
# ==========================================
st.set_page_config(
    page_title="올리브영 수주업로드 자동 입력 시스템",
    page_icon="https://raw.githubusercontent.com/paak1010/mentholatum_oliveyoung/main/logo.png",
    layout="wide",
)

custom_css = """
<style>
[data-testid="stHeader"] { visibility: hidden; }
footer { visibility: hidden; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# GitHub Raw 양식 URL (레포지토리에 올리신 서식파일)
TEMPLATE_URL = "https://raw.githubusercontent.com/paak1010/mentholatum_oliveyoung/main/%EC%98%AC%EB%A6%AC%EB%B8%8C%EC%96%81%20%EC%84%9C%EC%8B%9D%ED%8C%8C%EC%9D%BC(Final)_New%20System%20260720%20%EC%96%91%EC%A7%80%201%20(1).xlsx"


@st.cache_data
def load_template():
  """GitHub에서 올리브영 표준 양식 및 매핑 테이블 로드"""
  try:
    response = requests.get(TEMPLATE_URL)
    xls = pd.ExcelFile(io.BytesIO(response.content))
    df_tpl = pd.read_excel(xls, sheet_name="서식(수주업로드)", header=0)
    df_deliv = pd.read_excel(xls, sheet_name="배송처")
    df_prod = pd.read_excel(xls, sheet_name="제품명")
    return df_tpl, df_deliv, df_prod
  except Exception as e:
    st.error(
        f"GitHub 양식 파일을 불러오는데 실패했습니다: {e}\n(파일명/경로를 확인해주세요)"
    )
    return None, None, None


def to_safe_float(series):
  """전처리용: 숫자 이외 문자 제거 후 float 변환"""
  cleaned = series.astype(str).str.replace(r"[^0-9.]", "", regex=True)
  return pd.to_numeric(cleaned, errors="coerce").fillna(0)


# ==========================================
# 2. 메인 화면 UI (2개 파일 업로드)
# ==========================================
st.title("올리브영 수주업로드 자동 입력 시스템")
st.markdown("---")

col1, col2 = st.columns(2)

with col1:
  uploaded_inv_file = st.file_uploader(
      "📁 [1] 일일재고 파일 업로드 (.xlsx)", type=["xlsx"], key="inv_file"
  )

with col2:
  uploaded_order_file = st.file_uploader(
      "📁 [2] 올리브영 발주(Raw) 파일 업로드 (.xlsx)",
      type=["xlsx"],
      key="order_file",
  )

st.caption("💡 자동 부분 할당 및 재고 차감 적용 | ✔️ 잔여 유효일자 548일 이하 제외")
st.markdown("---")

# ==========================================
# 3. 데이터 처리 및 자동 할당 로직
# ==========================================
if uploaded_inv_file and uploaded_order_file:
  try:
    df_tpl_raw, df_deliv, df_prod = load_template()

    if df_tpl_raw is not None:
      # Raw 발주 및 일일재고 파일 읽기
      df_order_raw = pd.read_excel(uploaded_order_file)
      df_inv_raw = pd.read_excel(uploaded_inv_file)

      # ------------------------------------
      # A. 배송처 및 상품 코드(MECODE) 매핑
      # ------------------------------------
      df_order = df_order_raw.copy()

      # 바코드 정제 (소수점 제거)
      df_order["상품코드"] = (
          df_order["상품코드"].astype(str).str.split(".").str[0].str.strip()
      )

      # 배송코드 매핑 (센터명 기준)
      deliv_dict = dict(zip(df_deliv["배송처"], df_deliv["배송코드"]))
      df_order["배송코드"] = df_order["센터"].map(deliv_dict)

      # MECODE 및 단가 매핑 (제품명 시트 기준)
      df_prod["상품명_clean"] = (
          df_prod["상품명 "].astype(str).str.split(".").str[0].str.strip()
      )
      mecode_dict = dict(zip(df_prod["상품명_clean"], df_prod["상품코드"]))
      df_order["MECODE"] = df_order["상품코드"].map(mecode_dict)

      # ------------------------------------
      # B. 재고 시트 정제
      # ------------------------------------
      df_inv = df_inv_raw.copy()

      # 컬럼명 표준화
      rename_dict = {}
      for col in df_inv.columns:
        col_str = str(col).replace(" ", "").upper()
        if "상품" in col_str and "상품명" not in col_str:
          rename_dict[col] = "상품"
        elif "LOT" in col_str:
          rename_dict[col] = "화주LOT"
        elif "유효일자" in col_str or "유통기한" in col_str:
          rename_dict[col] = "유효일자"
        elif "환산" in col_str:
          rename_dict[col] = "환산"

      df_inv.rename(columns=rename_dict, inplace=True)

      df_inv["상품"] = df_inv["상품"].astype(str).str.strip().str.upper()
      df_inv["환산"] = to_safe_float(df_inv["환산"]).astype(float)

      # 유효일자 및 소비기한(548일 이하) 필터링
      df_inv["유효일자_DT"] = pd.to_datetime(
          df_inv["유효일자"], errors="coerce"
      )
      df_inv["유효일자_보존"] = df_inv["유효일자_DT"].fillna(
          pd.Timestamp("2099-12-31")
      )
      df_inv["유효일자_STR"] = (
          df_inv["유효일자_DT"].dt.strftime("%Y-%m-%d").fillna("")
      )

      today = pd.Timestamp.today().normalize()
      cutoff_date = today + pd.Timedelta(days=548)
      idx_short_shelf = df_inv["유효일자_보존"] <= cutoff_date
      idx_oc2 = (df_inv["상품"] == "ME90621OC2") & (
          ~df_inv["화주LOT"].astype(str).str.contains("분리배출")
      )

      df_inv_valid = df_inv[~(idx_oc2 | idx_short_shelf)].copy()
      df_inv_valid["화주LOT"] = df_inv_valid["화주LOT"].astype(str)

      # 박스 입수량 파악
      box_col_candidates = [
          col
          for col in df_inv.columns
          if "BOX" in str(col).upper() or "입수량" in str(col)
      ]
      box_col_name = box_col_candidates[0] if box_col_candidates else None
      product_box_unit = {}
      if box_col_name:
        for mecode, group in df_inv.groupby("상품"):
          box_vals = to_safe_float(group[box_col_name])
          box_vals = box_vals[box_vals > 0]
          if not box_vals.empty:
            product_box_unit[mecode] = int(box_vals.min())

      # 재고 그룹화
      if not df_inv_valid.empty:
        inv_grouped = (
            df_inv_valid.groupby(["상품", "유효일자_보존"])
            .agg({"환산": "sum", "화주LOT": "first", "유효일자_STR": "first"})
            .reset_index()
        )
      else:
        inv_grouped = pd.DataFrame(
            columns=["상품", "유효일자_보존", "환산", "화주LOT", "유효일자_STR"]
        )

      # ------------------------------------
      # C. 자동 할당 매칭
      # ------------------------------------
      df_order["발주수량\n(EA)"] = to_safe_float(df_order["발주수량\n(EA)"])
      df_order["LOT"] = ""
      df_order["유효일자_결과"] = ""
      df_order["할당상태"] = ""

      with st.spinner("재고 매칭 및 자동 수주 업로드 서식 작성 중..."):
        for i, row in df_order.iterrows():
          mecode = str(row["MECODE"])
          order_qty = float(row["발주수량\n(EA)"])

          if mecode in ["NAN", "", "NONE"] or order_qty <= 0:
            df_order.at[i, "할당상태"] = "제외"
            continue

          available_inv = inv_grouped[
              (inv_grouped["상품"] == mecode) & (inv_grouped["환산"] > 0)
          ]

          if available_inv.empty:
            (
                df_order.at[i, "LOT"],
                df_order.at[i, "유효일자_결과"],
                df_order.at[i, "할당상태"],
            ) = ("재고없음", "재고없음", "재고없음")
            continue

          full_match = available_inv[available_inv["환산"] >= order_qty]
          best_match = (
              full_match.sort_values(by="유효일자_보존").iloc[0]
              if not full_match.empty
              else available_inv.sort_values(by="유효일자_보존").iloc[0]
          )

          best_idx = best_match.name
          max_qty = float(best_match["환산"])
          lot_str = str(best_match["화주LOT"])
          date_str = str(best_match["유효일자_STR"])

          box_unit = product_box_unit.get(mecode, 1)
          potential_qty = min(order_qty, max_qty)
          allocated_boxes = int(potential_qty // box_unit)
          allocated_qty = float(allocated_boxes * box_unit)

          if allocated_qty > 0:
            df_order.at[i, "발주수량\n(EA)"] = allocated_qty
            df_order.at[i, "LOT"] = lot_str
            df_order.at[i, "유효일자_결과"] = date_str
            df_order.at[i, "할당상태"] = (
                "정상할당"
                if allocated_qty == order_qty
                else f"부분할당({allocated_boxes}BOX)"
            )
            inv_grouped.at[best_idx, "환산"] -= allocated_qty
          else:
            df_order.at[i, "할당상태"] = "박스단위부족"

      # ------------------------------------
      # D. 최종 올리브영 서식(수주업로드) 데이터 생성
      # ------------------------------------
      final_df = pd.DataFrame()
      final_df["발주처코드"] = 86100000
      final_df["입고예정일"] = df_order["입고예정일"]
      final_df["배송코드"] = df_order["배송코드"]
      final_df["ORDER #"] = df_order["입고전표"]
      final_df["상품명"] = df_order["상품명"]
      final_df["바코드"] = df_order["상품코드"]
      final_df["MECODE"] = df_order["MECODE"]
      final_df["수량"] = df_order["발주수량\n(EA)"]
      final_df["발주원가"] = df_order["원단가"]
      final_df["발주금액"] = df_order["원가금액"]
      final_df["LOT"] = df_order["LOT"]
      final_df["유효일자"] = df_order["유효일자_결과"]
      final_df["할당상태"] = df_order["할당상태"]

      # ------------------------------------
      # E. 결과 화면 및 다운로드
      # ------------------------------------
      st.success("🎉 올리브영 서식 변환 및 자동 재고 할당이 완료되었습니다!")

      st.subheader("📊 작성된 수주업로드 양식 미리보기")
      st.dataframe(final_df.head(100), use_container_width=True, hide_index=True)

      # 엑셀 다운로드 파일 생성
      buffer = io.BytesIO()
      with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, index=False, sheet_name="서식(수주업로드)")
        workbook = writer.book
        worksheet = writer.sheets["서식(수주업로드)"]
        text_format = workbook.add_format({"num_format": "@"})

        # 바코드, 배송코드, 날짜 열을 문자열(텍스트) 포맷 처리
        for target_col in ["발주처코드", "배송코드", "바코드", "유효일자"]:
          if target_col in final_df.columns:
            idx = final_df.columns.get_loc(target_col)
            worksheet.set_column(idx, idx, 15, text_format)

      st.download_button(
          label="💾 완성된 올리브영 수주업로드 엑셀 다운로드",
          data=buffer.getvalue(),
          file_name="올리브영_수주업로드_완성본.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          type="primary",
      )

  except Exception as e:
    st.error(f"데이터 처리 중 오류가 발생했습니다: {e}")
