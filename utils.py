# utils.py - 데이터 처리, 세무 계산, 데모 데이터 유틸리티
# 국군복지단 주류 세무/회계 자동화 시스템

import pandas as pd
import numpy as np
from io import BytesIO

# ============================================================
# 열 이름 매핑 테이블 (유연한 열 인식)
# - 사용자가 열 이름을 다르게 작성해도 내부 표준 키로 자동 변환
# ============================================================

PURCHASE_COL_MAP = {
    "product_name": ["주류명", "품명", "상품명", "제품명", "품목명", "품목", "주류", "상품"],
    "spec":         ["규격", "용량", "단위", "사이즈", "규격용량", "규격/용량"],
    "quantity":     ["수량", "입고수량", "매입수량", "구매수량", "개수", "병수"],
    "unit_price":   ["단가", "매입단가", "매입가", "구매단가", "단가원"],
    "total_amount": ["금액", "매입금액", "총금액", "합계금액", "공급가액", "총액", "매입총액"],
    "vat_type":     ["부가세여부", "과세구분", "세금구분", "과세면세", "구분", "세구분", "과세여부"],
    "vat_amount":   ["부가세", "VAT", "세액", "부가가치세", "부가세액", "세금"],
}

SALES_COL_MAP = {
    "product_name":   ["주류명", "품명", "상품명", "제품명", "품목명", "품목", "주류", "상품"],
    "spec":           ["규격", "용량", "단위", "사이즈", "규격용량", "규격/용량"],
    "regular_qty":    ["일반판매수량", "과세수량", "일반수량", "과세판매수량", "일반", "과세"],
    "exempt_qty":     ["면세판매수량", "면세수량", "면세판매량", "면세"],
    "unit_price":     ["판매단가", "단가", "판매가", "판매가격", "단가원"],
    "regular_amount": ["일반판매금액", "과세금액", "일반금액", "과세매출", "일반매출금액"],
    "exempt_amount":  ["면세판매금액", "면세금액", "면세매출", "면세매출금액"],
}


# ============================================================
# 열 이름 표준화
# ============================================================

def normalize_columns(df: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """
    데이터프레임 열 이름을 표준 영문 키로 매핑합니다.
    1단계: 정확 일치 / 2단계: 앞부분 일치 순으로 탐색합니다.
    하나의 표준 키는 최초 매칭된 열 하나에만 할당됩니다.
    """
    rename_dict: dict = {}
    assigned_keys: set = set()

    def _clean(text: str) -> str:
        return str(text).strip().replace(" ", "")

    # 1단계: 정확 일치
    for col in df.columns:
        if col in rename_dict:
            continue
        col_c = _clean(col)
        for key, candidates in col_map.items():
            if key in assigned_keys:
                continue
            if any(col_c == _clean(c) for c in candidates):
                rename_dict[col] = key
                assigned_keys.add(key)
                break

    # 2단계: 앞부분 일치 (정확 매칭 실패한 열에 한해)
    for col in df.columns:
        if col in rename_dict:
            continue
        col_c = _clean(col)
        for key, candidates in col_map.items():
            if key in assigned_keys:
                continue
            if any(col_c.startswith(_clean(c)) for c in candidates):
                rename_dict[col] = key
                assigned_keys.add(key)
                break

    return df.rename(columns=rename_dict)


def _to_num(series: pd.Series) -> pd.Series:
    """쉼표·공백이 포함된 문자열 금액을 숫자형으로 변환"""
    return pd.to_numeric(
        series.astype(str)
              .str.replace(",", "", regex=False)
              .str.replace(" ", "", regex=False),
        errors="coerce",
    ).fillna(0)


# ============================================================
# 데모 데이터 생성
# ============================================================

def generate_demo_purchase_data() -> pd.DataFrame:
    """데모용 매입 데이터 (주류 10종)"""
    data = {
        "주류명": ["참이슬", "테라", "카스", "클라우드", "막걸리(서울)", "복분자주", "처음처럼", "하이트", "맥스", "산사춘"],
        "규격":   ["360ml", "500ml", "500ml", "500ml", "750ml", "375ml", "360ml", "500ml", "500ml", "375ml"],
        "수량":   [500, 300, 400, 200, 150, 100, 400, 300, 250, 80],
        "단가":   [850, 1_200, 1_100, 1_300, 1_100, 2_500, 850, 1_150, 1_250, 2_200],
        # 과세구분: 막걸리는 면세(전통주), 나머지는 과세
        "부가세여부": ["과세","과세","과세","과세","면세","과세","과세","과세","과세","과세"],
    }
    df = pd.DataFrame(data)
    df["금액"] = df["수량"] * df["단가"]
    # 매입 부가세: 공급가액(금액) × 10% (면세는 0)
    df["부가세"] = df.apply(
        lambda r: round(r["금액"] * 0.10) if r["부가세여부"] == "과세" else 0,
        axis=1,
    )
    return df


def generate_demo_sales_data() -> pd.DataFrame:
    """데모용 판매 데이터 (주류 10종)"""
    data = {
        "주류명":       ["참이슬", "테라", "카스", "클라우드", "막걸리(서울)", "복분자주", "처음처럼", "하이트", "맥스", "산사춘"],
        "규격":         ["360ml", "500ml", "500ml", "500ml", "750ml", "375ml", "360ml", "500ml", "500ml", "375ml"],
        # 일반판매수량: 일반(과세) 고객 판매량
        "일반판매수량": [280, 180, 250, 120,   0, 60, 220, 170, 150, 40],
        # 면세판매수량: 군인(면세) 판매량
        "면세판매수량": [150,  90, 120,  60, 120, 30, 120,  90,  80, 30],
        # 판매단가는 공급대가(VAT 포함) 기준
        "판매단가":     [1_500, 2_200, 2_000, 2_400, 1_800, 4_500, 1_500, 2_100, 2_300, 3_800],
    }
    df = pd.DataFrame(data)
    # 판매금액 = 공급대가 (VAT 포함, 일반판매 기준)
    df["일반판매금액"] = df["일반판매수량"] * df["판매단가"]
    df["면세판매금액"] = df["면세판매수량"] * df["판매단가"]
    return df


# ============================================================
# 핵심 세무 계산 로직
# ============================================================

def calculate_tax(purchase_df: pd.DataFrame, sales_df: pd.DataFrame) -> dict:
    """
    매입·판매 데이터를 기반으로 부가세 및 이익을 계산합니다.

    부가세 계산 원칙:
      - 일반판매 매출세액 = 공급대가(일반판매금액) × (10/110)
      - 면세판매 매출세액 = 0
      - 매입세액 = 과세 매입 시 지불한 부가세 합계
      - 납부할 세액 = 총 매출세액 − 총 매입세액

    Returns:
        {
          'summary'        : dict  (세무 요약 수치),
          'sales_detail'   : DataFrame (판매 상세),
          'purchase_detail': DataFrame (매입 상세),
        }
    """
    # ── 열 이름 표준화 ──────────────────────────────────────
    purchase = normalize_columns(purchase_df.copy(), PURCHASE_COL_MAP)
    sales    = normalize_columns(sales_df.copy(),    SALES_COL_MAP)

    if "product_name" not in purchase.columns:
        raise ValueError("매입 데이터에서 '주류명(품명)' 열을 찾을 수 없습니다.\n"
                         "열 이름 가이드를 참고하여 파일을 수정해 주세요.")
    if "product_name" not in sales.columns:
        raise ValueError("판매 데이터에서 '주류명(품명)' 열을 찾을 수 없습니다.\n"
                         "열 이름 가이드를 참고하여 파일을 수정해 주세요.")

    # ── 숫자형 변환 ─────────────────────────────────────────
    for col in ["quantity", "unit_price", "total_amount", "vat_amount"]:
        if col in purchase.columns:
            purchase[col] = _to_num(purchase[col])

    for col in ["regular_qty", "exempt_qty", "unit_price", "regular_amount", "exempt_amount"]:
        if col in sales.columns:
            sales[col] = _to_num(sales[col])

    # ── 매입금액·단가 보완 ───────────────────────────────────
    if "total_amount" not in purchase.columns:
        if "quantity" in purchase.columns and "unit_price" in purchase.columns:
            purchase["total_amount"] = purchase["quantity"] * purchase["unit_price"]
        else:
            purchase["total_amount"] = 0.0

    if "unit_price" not in purchase.columns:
        if "total_amount" in purchase.columns and "quantity" in purchase.columns:
            purchase["unit_price"] = (
                purchase["total_amount"] / purchase["quantity"].replace(0, np.nan)
            ).fillna(0)
        else:
            purchase["unit_price"] = 0.0

    # ── 매입 부가세 계산 ─────────────────────────────────────
    # 우선순위: ① 직접 입력된 부가세 열 ② 과세구분 열 기반 계산 ③ 전체 과세 가정
    if "vat_amount" not in purchase.columns:
        TAXABLE_FLAGS = {"과세", "과세품", "y", "yes", "1", "예", "true"}
        if "vat_type" in purchase.columns:
            purchase["vat_amount"] = purchase.apply(
                lambda r: r["total_amount"] * 0.10
                if str(r["vat_type"]).strip().lower() in TAXABLE_FLAGS
                else 0.0,
                axis=1,
            )
        else:
            # 과세구분 열이 없으면 전체 과세로 처리 (보수적 기본값)
            purchase["vat_amount"] = purchase["total_amount"] * 0.10

    # ── 판매금액 보완 (수량 × 단가) ─────────────────────────
    if "regular_amount" not in sales.columns:
        sales["regular_amount"] = (
            sales.get("regular_qty", pd.Series(0, index=sales.index)) *
            sales.get("unit_price",  pd.Series(0, index=sales.index))
        )
    if "exempt_amount" not in sales.columns:
        sales["exempt_amount"] = (
            sales.get("exempt_qty", pd.Series(0, index=sales.index)) *
            sales.get("unit_price", pd.Series(0, index=sales.index))
        )

    # ── 매출세액 계산 ────────────────────────────────────────
    # 공급대가(일반판매금액)에서 세액 및 공급가액을 역산
    sales["regular_supply_price"] = sales["regular_amount"] * (100 / 110)  # 공급가액
    sales["regular_vat"]          = sales["regular_amount"] * (10  / 110)  # 매출세액
    sales["exempt_vat"]           = 0.0                                      # 면세 매출세액
    sales["total_sales"]          = sales["regular_amount"] + sales["exempt_amount"]

    # ── 판매·매입 병합 (이익 계산) ──────────────────────────
    # 주류명(+규격)을 기준으로 매입단가와 판매 데이터를 결합
    merge_keys = ["product_name"]
    if "spec" in purchase.columns and "spec" in sales.columns:
        merge_keys.append("spec")

    purch_for_merge = (
        purchase[merge_keys + ["unit_price"]]
        .rename(columns={"unit_price": "purchase_unit_price"})
    )

    sales_detail = sales.merge(purch_for_merge, on=merge_keys, how="left")

    reg_qty = sales_detail.get("regular_qty", pd.Series(0, index=sales_detail.index))
    exm_qty = sales_detail.get("exempt_qty",  pd.Series(0, index=sales_detail.index))

    sales_detail["sold_qty"]      = reg_qty + exm_qty
    pu_price = sales_detail.get("purchase_unit_price", pd.Series(0, index=sales_detail.index)).fillna(0)
    sales_detail["cost_of_goods"] = sales_detail["sold_qty"] * pu_price
    sales_detail["gross_profit"]  = sales_detail["total_sales"] - sales_detail["cost_of_goods"]
    sales_detail["profit_margin"] = np.where(
        sales_detail["total_sales"] > 0,
        (sales_detail["gross_profit"] / sales_detail["total_sales"] * 100).round(1),
        0.0,
    )

    # ── 세무 요약 집계 ────────────────────────────────────────
    total_regular_sales = float(sales["regular_amount"].sum())
    total_exempt_sales  = float(sales["exempt_amount"].sum())
    total_output_vat    = float(sales["regular_vat"].sum())
    total_input_vat     = float(purchase["vat_amount"].sum())

    return {
        "summary": {
            "total_regular_sales": total_regular_sales,
            "total_exempt_sales":  total_exempt_sales,
            "total_sales":         total_regular_sales + total_exempt_sales,
            "total_output_vat":    total_output_vat,
            "total_input_vat":     total_input_vat,
            "vat_payable":         total_output_vat - total_input_vat,
        },
        "sales_detail":    sales_detail,
        "purchase_detail": purchase,
    }


# ============================================================
# 엑셀 보고서 생성
# ============================================================

# 판매 상세 열 이름 한글 매핑
_SALES_COL_KO = {
    "product_name":        "주류명",
    "spec":                "규격",
    "regular_qty":         "일반판매수량",
    "exempt_qty":          "면세판매수량",
    "sold_qty":            "총판매수량",
    "unit_price":          "판매단가",
    "regular_amount":      "일반판매금액(공급대가)",
    "regular_supply_price":"공급가액",
    "regular_vat":         "매출세액",
    "exempt_amount":       "면세판매금액",
    "total_sales":         "총판매금액",
    "purchase_unit_price": "매입단가",
    "cost_of_goods":       "매입원가",
    "gross_profit":        "매출총이익",
    "profit_margin":       "이익률(%)",
}

# 매입 상세 열 이름 한글 매핑
_PURCH_COL_KO = {
    "product_name": "주류명",
    "spec":         "규격",
    "quantity":     "수량",
    "unit_price":   "매입단가",
    "total_amount": "매입금액(공급가액)",
    "vat_type":     "과세구분",
    "vat_amount":   "매입부가세",
}


def create_excel_report(results: dict) -> bytes:
    """
    회계 분석 결과를 서식 있는 엑셀(.xlsx) 파일로 변환합니다.
    시트 구성: ① 세무 요약  ② 판매 상세  ③ 매입 상세
    """
    output = BytesIO()
    s = results["summary"]

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        # ── 공통 서식 정의 ───────────────────────────────────
        FMT = {
            "header":      wb.add_format({"bold": True, "bg_color": "#1F4E79", "font_color": "white",
                                          "border": 1, "align": "center", "valign": "vcenter"}),
            "money":       wb.add_format({"num_format": "#,##0", "border": 1}),
            "pct":         wb.add_format({"num_format": '0.0"%"', "border": 1, "align": "center"}),
            "normal":      wb.add_format({"border": 1}),
            "title":       wb.add_format({"bold": True, "font_size": 14, "font_color": "#1F4E79"}),
            "subheader":   wb.add_format({"bold": True, "border": 1, "bg_color": "#BDD7EE"}),
            "bold_money":  wb.add_format({"bold": True, "border": 1, "bg_color": "#BDD7EE",
                                          "num_format": "#,##0"}),
        }

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 시트 1: 세무 요약
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        ws1 = wb.add_worksheet("세무 요약")
        writer.sheets["세무 요약"] = ws1

        ws1.write("A1", "국군복지단 주류 부가세 신고 요약서", FMT["title"])
        ws1.write("A2", f"작성일: {pd.Timestamp.now().strftime('%Y년 %m월 %d일')}")

        ws1.write_row(3, 0, ["항  목", "금액 (원)"], FMT["header"])

        # (레이블, 값, 레이블서식, 값서식)
        rows_summary = [
            ("── 매출 내역 ──",                            None,                               FMT["subheader"],  None),
            ("  일반판매(과세) 공급대가 합계",             s["total_regular_sales"],           FMT["normal"],     FMT["money"]),
            ("  일반판매 공급가액  (공급대가 × 100/110)", s["total_regular_sales"] * 100/110, FMT["normal"],     FMT["money"]),
            ("  면세판매 금액",                            s["total_exempt_sales"],            FMT["normal"],     FMT["money"]),
            ("  총 매출액 합계",                           s["total_sales"],                   FMT["subheader"],  FMT["bold_money"]),
            ("",                                           None,                               FMT["normal"],     None),
            ("── 부가세 계산 ──",                          None,                               FMT["subheader"],  None),
            ("  매출세액  (일반판매 공급대가 × 10/110)",  s["total_output_vat"],              FMT["normal"],     FMT["money"]),
            ("  매입세액  (과세매입 부가세 합계)",         s["total_input_vat"],               FMT["normal"],     FMT["money"]),
            ("  납부할 세액  (매출세액 − 매입세액)",       s["vat_payable"],                   FMT["subheader"],  FMT["bold_money"]),
        ]

        for i, (lbl, val, lf, vf) in enumerate(rows_summary):
            ws1.write(4 + i, 0, lbl, lf)
            if val is not None and vf:
                ws1.write(4 + i, 1, val, vf)

        ws1.set_column("A:A", 48)
        ws1.set_column("B:B", 22)

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 시트 2: 판매 상세
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        sd = results["sales_detail"].copy()
        disp_cols = [c for c in _SALES_COL_KO if c in sd.columns]
        sd_display = sd[disp_cols].rename(columns=_SALES_COL_KO)
        sd_display.to_excel(writer, sheet_name="판매 상세", index=False)

        ws2 = writer.sheets["판매 상세"]
        for j, col_name in enumerate(sd_display.columns):
            ws2.write(0, j, col_name, FMT["header"])
        ws2.set_column("A:Z", 17)

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 시트 3: 매입 상세
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        pd_ = results["purchase_detail"].copy()
        disp_cols2 = [c for c in _PURCH_COL_KO if c in pd_.columns]
        pd_display = pd_[disp_cols2].rename(columns=_PURCH_COL_KO)
        pd_display.to_excel(writer, sheet_name="매입 상세", index=False)

        ws3 = writer.sheets["매입 상세"]
        for j, col_name in enumerate(pd_display.columns):
            ws3.write(0, j, col_name, FMT["header"])
        ws3.set_column("A:Z", 17)

    output.seek(0)
    return output.getvalue()
