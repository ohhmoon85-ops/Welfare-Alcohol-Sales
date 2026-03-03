# app.py - 국군복지단 주류 세무/회계 자동화 시스템 메인 UI
# Streamlit 기반 웹 애플리케이션

import traceback

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from utils import (
    calculate_tax,
    create_excel_report,
    generate_demo_purchase_data,
    generate_demo_sales_data,
)

# ============================================================
# 페이지 기본 설정 (반드시 첫 번째 Streamlit 호출)
# ============================================================
st.set_page_config(
    page_title="국군복지단 주류 회계",
    page_icon="🍶",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 전역 CSS 스타일
# ============================================================
st.markdown(
    """
<style>
/* KPI 카드 공통 */
.kpi-card {
    padding: 18px 16px;
    border-radius: 12px;
    color: white;
    text-align: center;
    margin-bottom: 6px;
    box-shadow: 0 3px 10px rgba(0,0,0,.18);
}
.kpi-blue   { background: linear-gradient(135deg, #1F4E79, #2E75B6); }
.kpi-green  { background: linear-gradient(135deg, #1E5631, #2D8653); }
.kpi-orange { background: linear-gradient(135deg, #843C0C, #C55A11); }
.kpi-gray   { background: linear-gradient(135deg, #44546A, #6E8095); }
.kpi-label  { font-size: 13px; opacity: .85; margin-bottom: 5px; }
.kpi-value  { font-size: 24px; font-weight: 700; letter-spacing: -.5px; }

/* 섹션 타이틀 */
.sec-title {
    border-left: 5px solid #1F4E79;
    padding: 4px 0 4px 12px;
    font-size: 17px;
    font-weight: 700;
    margin: 22px 0 10px;
    color: #1F4E79;
}

/* 파일 업로더 외곽 */
[data-testid="stFileUploader"] {
    border: 1.5px dashed #2E75B6;
    border-radius: 8px;
    padding: 6px;
}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# 헬퍼 함수
# ============================================================

def won(amount: float) -> str:
    """금액을 한국 원화 포맷(₩1,234)으로 변환"""
    return f"₩{amount:,.0f}"


def kpi_card(label: str, value: str, color: str = "blue") -> None:
    """KPI 카드 HTML 렌더링"""
    st.markdown(
        f'<div class="kpi-card kpi-{color}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f"</div>",
        unsafe_allow_html=True,
    )


def sec_title(text: str) -> None:
    """섹션 구분선 타이틀"""
    st.markdown(f'<div class="sec-title">{text}</div>', unsafe_allow_html=True)


# ============================================================
# 열 이름 가이드 패널
# ============================================================

def render_col_guide() -> None:
    with st.expander("📋 엑셀 열 이름 인식 가이드 — 열 이름이 다를 때 참고하세요 (클릭하여 펼치기)"):
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("##### 📥 매입 데이터 열 이름")
            st.dataframe(
                pd.DataFrame(
                    {
                        "표준 열 이름": ["주류명", "규격", "수량", "단가", "금액", "과세구분", "부가세"],
                        "대체 가능한 이름": [
                            "품명 · 상품명 · 제품명 · 품목",
                            "용량 · 단위 · 사이즈",
                            "입고수량 · 매입수량 · 개수",
                            "매입단가 · 매입가 · 구매단가",
                            "매입금액 · 총금액 · 공급가액",
                            "부가세여부 · 세금구분 · 구분",
                            "VAT · 세액 · 부가가치세",
                        ],
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )
        with c2:
            st.markdown("##### 📤 판매 데이터 열 이름")
            st.dataframe(
                pd.DataFrame(
                    {
                        "표준 열 이름": [
                            "주류명", "규격", "일반판매수량", "면세판매수량",
                            "판매단가", "일반판매금액", "면세판매금액",
                        ],
                        "대체 가능한 이름": [
                            "품명 · 상품명 · 제품명 · 품목",
                            "용량 · 단위 · 사이즈",
                            "과세수량 · 일반수량 · 과세판매수량",
                            "면세수량 · 면세판매량",
                            "단가 · 판매가 · 판매가격",
                            "과세금액 · 일반금액 · 과세매출",
                            "면세금액 · 면세매출",
                        ],
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )
        st.caption(
            "※ 공백·대소문자는 무시됩니다. 위 목록에 없는 이름은 직접 수정 후 업로드하세요."
        )


# ============================================================
# 대시보드 렌더링
# ============================================================

def render_dashboard(results: dict) -> None:
    s  = results["summary"]
    sd = results["sales_detail"]

    # ── 행 1: 핵심 KPI ────────────────────────────────────────
    sec_title("📊 세무 요약 대시보드")
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    with r1c1: kpi_card("💰 총 매출액",         won(s["total_sales"]),         "blue")
    with r1c2: kpi_card("🏪 일반(과세) 매출",   won(s["total_regular_sales"]), "blue")
    with r1c3: kpi_card("🎖️ 면세 매출",         won(s["total_exempt_sales"]),  "green")
    with r1c4:
        color = "orange" if s["vat_payable"] > 0 else "green"
        kpi_card("🧾 납부할 세액", won(s["vat_payable"]), color)

    st.write("")

    # ── 행 2: 부가세 상세 ────────────────────────────────────
    r2c1, r2c2, r2c3, _ = st.columns(4)
    with r2c1: kpi_card("📤 매출세액",          won(s["total_output_vat"]),              "gray")
    with r2c2: kpi_card("📥 매입세액 (공제)",   won(s["total_input_vat"]),               "green")
    with r2c3: kpi_card("📦 일반판매 공급가액", won(s["total_regular_sales"] * 100/110), "gray")

    st.write("")

    # ── 차트 행 ──────────────────────────────────────────────
    ch1, ch2 = st.columns(2)

    with ch1:
        st.markdown("**판매 구성 (과세 vs 면세)**")
        fig_pie = go.Figure(
            go.Pie(
                labels=["일반판매(과세)", "면세판매"],
                values=[s["total_regular_sales"], s["total_exempt_sales"]],
                hole=0.44,
                marker_colors=["#2E75B6", "#2D8653"],
                textinfo="label+percent",
                hovertemplate="%{label}<br>%{value:,.0f}원<extra></extra>",
            )
        )
        fig_pie.update_layout(
            showlegend=False,
            height=280,
            margin=dict(t=8, b=8, l=0, r=0),
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with ch2:
        st.markdown("**부가세 계산 흐름 (워터폴 차트)**")
        fig_wf = go.Figure(
            go.Waterfall(
                orientation="v",
                measure=["absolute", "relative", "total"],
                x=["매출세액", "매입세액 차감", "납부할 세액"],
                y=[s["total_output_vat"], -s["total_input_vat"], 0],
                text=[won(s["total_output_vat"]),
                      f"−{won(s['total_input_vat'])}",
                      won(s["vat_payable"])],
                textposition="outside",
                increasing=dict(marker=dict(color="#ED7D31")),
                decreasing=dict(marker=dict(color="#2D8653")),
                totals=dict(marker=dict(color="#2E75B6")),
                connector=dict(line=dict(color="#aaa", width=1.5)),
            )
        )
        fig_wf.update_layout(
            height=280,
            margin=dict(t=8, b=8, l=0, r=0),
            yaxis_tickformat=",.0f",
            yaxis_title="금액 (원)",
        )
        st.plotly_chart(fig_wf, use_container_width=True)

    # ── 상품별 바 차트 ────────────────────────────────────────
    sec_title("🍶 상품별 판매 현황")

    if "product_name" in sd.columns:
        bar_df = pd.DataFrame({
            "주류명":   sd["product_name"],
            "일반판매": sd.get("regular_amount", pd.Series(0, index=sd.index)).fillna(0),
            "면세판매": sd.get("exempt_amount",  pd.Series(0, index=sd.index)).fillna(0),
        })
        bar_melt = bar_df.melt(id_vars="주류명", var_name="구분", value_name="금액")
        fig_bar = px.bar(
            bar_melt,
            x="주류명",
            y="금액",
            color="구분",
            barmode="stack",
            color_discrete_map={"일반판매": "#2E75B6", "면세판매": "#2D8653"},
            height=380,
            labels={"금액": "판매금액 (원)"},
        )
        fig_bar.update_yaxes(tickformat=",.0f")
        fig_bar.update_layout(
            legend_title_text="",
            margin=dict(t=10, b=10),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # ── 이익률 차트 ───────────────────────────────────────────
    if "profit_margin" in sd.columns and "product_name" in sd.columns:
        sec_title("📈 상품별 이익률")
        fig_margin = px.bar(
            sd,
            x="product_name",
            y="profit_margin",
            color="profit_margin",
            color_continuous_scale=["#C55A11", "#F7C948", "#2D8653"],
            range_color=[0, 50],
            height=300,
            labels={"product_name": "주류명", "profit_margin": "이익률 (%)"},
            text="profit_margin",
        )
        fig_margin.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig_margin.update_layout(
            coloraxis_showscale=False,
            margin=dict(t=10, b=10),
        )
        st.plotly_chart(fig_margin, use_container_width=True)

    # ── 판매 상세 테이블 ─────────────────────────────────────
    sec_title("📋 판매 상세 내역")

    _COL_KO = {
        "product_name":        "주류명",
        "spec":                "규격",
        "regular_qty":         "일반판매수량",
        "exempt_qty":          "면세판매수량",
        "sold_qty":            "총판매수량",
        "unit_price":          "판매단가",
        "regular_amount":      "일반판매금액",
        "exempt_amount":       "면세판매금액",
        "regular_vat":         "매출세액",
        "total_sales":         "총판매금액",
        "purchase_unit_price": "매입단가",
        "cost_of_goods":       "매입원가",
        "gross_profit":        "매출총이익",
        "profit_margin":       "이익률(%)",
    }
    disp_cols = [c for c in _COL_KO if c in sd.columns]
    df_show   = sd[disp_cols].rename(columns=_COL_KO)

    fmt: dict = {}
    for col in ["일반판매금액", "면세판매금액", "매출세액", "총판매금액", "매입원가", "매출총이익"]:
        if col in df_show.columns:
            fmt[col] = "₩{:,.0f}"
    for col in ["판매단가", "매입단가"]:
        if col in df_show.columns:
            fmt[col] = "₩{:,.0f}"
    if "이익률(%)" in df_show.columns:
        fmt["이익률(%)"] = "{:.1f}%"

    st.dataframe(df_show.style.format(fmt), use_container_width=True, hide_index=True)

    # ── 세무 계산 근거 ────────────────────────────────────────
    with st.expander("🔍 세무 계산 근거 상세 보기"):
        st.markdown(
            f"""
| 항목 | 계산식 | 금액 |
|---|---|---|
| 일반판매 공급대가 합계 | 합계 | **{won(s['total_regular_sales'])}** |
| 일반판매 공급가액 | 공급대가 × (100÷110) | {won(s['total_regular_sales'] * 100/110)} |
| **매출세액** | 공급대가 × (10÷110) | **{won(s['total_output_vat'])}** |
| **매입세액** | 과세매입 부가세 합계 | **{won(s['total_input_vat'])}** |
| **납부할 세액** | 매출세액 − 매입세액 | **{won(s['vat_payable'])}** |

> ※ 면세판매(예: 군인 전용 품목)는 부가세 계산에서 제외됩니다.
> ※ 과세구분 열이 없는 매입 데이터는 **전체 과세**로 처리됩니다.
"""
        )


# ============================================================
# 메인 애플리케이션
# ============================================================

def main() -> None:

    # ── 사이드바 ──────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 🍶 국군복지단\n### 주류 세무/회계 시스템")
        st.divider()
        st.markdown(
            """
**사용 방법**

1. 매입 엑셀 업로드
2. 판매 엑셀 업로드
3. **분석 실행** 클릭
4. 결과 확인 후 엑셀 다운로드

파일이 없으면 **데모 실행** 버튼을 먼저 누르세요.
"""
        )
        st.divider()
        st.markdown(
            """
**지원 파일 형식**
- `.xlsx` / `.xls`
- 첫 번째 행 = 열 이름
- 금액: 숫자 (쉼표 허용)
"""
        )
        st.divider()
        st.caption("ver 1.0  ·  국군복지단 전용")

    # ── 헤더 ─────────────────────────────────────────────────
    st.title("🍶 국군복지단 주류 세무/회계 자동화 시스템")
    st.markdown(
        "매입·판매 엑셀을 업로드하면 **부가세 자동 계산**과 "
        "**회계 보고서 생성**을 즉시 처리합니다."
    )

    # ── 열 이름 가이드 ────────────────────────────────────────
    render_col_guide()
    st.divider()

    # ── 데이터 입력 영역 ──────────────────────────────────────
    st.markdown("### 📁 데이터 입력")

    demo_col, _ = st.columns([2, 5])
    with demo_col:
        demo_btn = st.button(
            "🎮  데모 데이터로 즉시 실행",
            type="primary",
            use_container_width=True,
            help="주류 10종 샘플 데이터로 전체 기능을 즉시 체험합니다.",
        )

    st.write("")

    # ── 파일 업로더 ───────────────────────────────────────────
    up1, up2 = st.columns(2)
    purchase_df: pd.DataFrame | None = None
    sales_df:    pd.DataFrame | None = None

    with up1:
        st.markdown("**📥 매입 데이터 업로드**")
        pu_file = st.file_uploader(
            "매입 엑셀 파일 선택",
            type=["xlsx", "xls"],
            key="pu",
            label_visibility="collapsed",
        )
        if pu_file:
            try:
                purchase_df = pd.read_excel(pu_file)
                st.success(f"✅ 매입 데이터 {len(purchase_df)}행 로드 완료")
                with st.expander("미리보기 (최대 5행)"):
                    st.dataframe(purchase_df.head(), use_container_width=True)
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

    with up2:
        st.markdown("**📤 판매 데이터 업로드**")
        sa_file = st.file_uploader(
            "판매 엑셀 파일 선택",
            type=["xlsx", "xls"],
            key="sa",
            label_visibility="collapsed",
        )
        if sa_file:
            try:
                sales_df = pd.read_excel(sa_file)
                st.success(f"✅ 판매 데이터 {len(sales_df)}행 로드 완료")
                with st.expander("미리보기 (최대 5행)"):
                    st.dataframe(sales_df.head(), use_container_width=True)
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

    # ── 실행 조건 결정 ────────────────────────────────────────
    run_analysis = False

    if demo_btn:
        # 데모 버튼 클릭: 샘플 데이터 생성 후 즉시 분석
        purchase_df = generate_demo_purchase_data()
        sales_df    = generate_demo_sales_data()
        run_analysis = True
        st.info("📌 데모 데이터(주류 10종)를 사용하여 분석합니다.")

    elif purchase_df is not None and sales_df is not None:
        # 양쪽 파일 모두 업로드된 경우 분석 버튼 표시
        run_col, _ = st.columns([2, 5])
        with run_col:
            if st.button("🚀  분석 실행", type="primary", use_container_width=True):
                run_analysis = True

    elif purchase_df is not None or sales_df is not None:
        # 한쪽만 업로드된 경우 안내 메시지
        st.warning(
            "매입 데이터와 판매 데이터를 **모두** 업로드해야 분석을 실행할 수 있습니다."
        )

    # ── 분석 실행 및 결과 표시 ────────────────────────────────
    if not run_analysis:
        return  # 아직 실행 조건 미충족

    st.divider()

    try:
        with st.spinner("분석 중... 잠시만 기다려 주세요."):
            results = calculate_tax(purchase_df, sales_df)

        # 대시보드 렌더링
        render_dashboard(results)

        # ── 다운로드 ─────────────────────────────────────────
        st.divider()
        st.markdown("### 💾 보고서 다운로드")

        dl1, dl2, _ = st.columns([1, 1, 2])

        with dl1:
            xlsx_bytes = create_excel_report(results)
            st.download_button(
                label="📊  Excel 보고서 다운로드",
                data=xlsx_bytes,
                file_name=f"주류회계보고서_{pd.Timestamp.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with dl2:
            # 판매 상세 CSV (BOM 포함 → 엑셀에서 한글 깨짐 방지)
            _COL_CSV = {
                "product_name": "주류명", "spec": "규격",
                "regular_qty": "일반판매수량", "exempt_qty": "면세판매수량",
                "unit_price": "판매단가", "regular_amount": "일반판매금액",
                "exempt_amount": "면세판매금액", "regular_vat": "매출세액",
                "total_sales": "총판매금액",
            }
            sd = results["sales_detail"]
            disp = [c for c in _COL_CSV if c in sd.columns]
            csv_data = (
                sd[disp].rename(columns=_COL_CSV)
                .to_csv(index=False, encoding="utf-8-sig")
            )
            st.download_button(
                label="📄  판매 상세 CSV 다운로드",
                data=csv_data,
                file_name=f"판매상세_{pd.Timestamp.now():%Y%m%d_%H%M}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    except ValueError as ve:
        st.error(f"❌ 데이터 오류: {ve}")
    except Exception:
        st.error("❌ 분석 중 예상치 못한 오류가 발생했습니다.")
        with st.expander("🔧 오류 상세 정보 (개발자용)"):
            st.code(traceback.format_exc())


# ============================================================
if __name__ == "__main__":
    main()
