"""streamlit_app.py — LF몰 제휴 일자별 실적 집계 (경량, UI 전용)

kimhyemin 본 대시보드에서 내보낸 (날짜×제휴사) 집계 JSON만 읽어
날짜×제휴사 / 제휴사×월 피벗과 '분석일 특이점 코멘트'를 즉시 보여준다.
순수 로직은 pivot_core.py에 있고 여기서는 컨트롤·표·HTML만 담당한다.
컨벤션(회사): 전년비 상승 ▼(초록) / 하락 △(빨강). ▲는 쓰지 않는다.
"""
from __future__ import annotations
import os

import numpy as np
import pandas as pd
import streamlit as st

import pivot_core as C

st.set_page_config(page_title="LF몰 제휴 일자별 실적 집계", page_icon="🗓️", layout="wide")

UP_COLOR = "#059669"    # 상승(초록)
DOWN_COLOR = "#dc2626"  # 하락(빨강)
BUNDLED = os.path.join(os.path.dirname(__file__), "data", "daily_pivot.json")


@st.cache_data(show_spinner=False)
def load(raw: bytes):
    import json
    payload = json.loads(raw)
    df = C.prepare(pd.DataFrame(payload.get("records", [])))
    return df, payload


# ── 스타일 ──────────────────────────────────────────────
def style_values(mat, fmt_fn):
    return (mat.style
            .format(fmt_fn)
            .apply(lambda s: ["font-weight:700" if s.name == "합계" else "" for _ in s], axis=1)
            .set_properties(subset=["합계"], **{"font-weight": "700"}))


def style_yoy(mat):
    def f(v):
        return "" if C.na(v) else f"{'▼' if v >= 0 else '△'} {v:+.1f}%"

    def color(v):
        return "" if C.na(v) else (f"color:{UP_COLOR}" if v >= 0 else f"color:{DOWN_COLOR}")

    return mat.style.format(f).map(color)


def _cspan(text, up):
    return f'<span style="color:{UP_COLOR if up else DOWN_COLOR};font-weight:600">{text}</span>'


def _yoy_html(p):
    if p is None:
        return '<span style="color:#9ca3af">전년비 –</span>'
    return "전년비 " + _cspan(f"{'▼' if p >= 0 else '△'} {p:+.1f}%", p >= 0)


def _wd_html(p):
    if p is None:
        return '<span style="color:#9ca3af">동요일평균 –</span>'
    return "동요일평균 대비 " + _cspan(f"{'▼' if p >= 0 else '△'} {p:+.1f}%", p >= 0)


def comment_html(res: dict) -> str:
    parts = []
    for metric, v in res["metrics"].items():
        fmt_fn = C.FMT[C.METRICS[metric]["fmt"]]
        parts.append(
            f'<div style="margin:.15rem 0"><b>{metric}</b> {fmt_fn(v["cur"])} '
            f'&nbsp;·&nbsp; {_yoy_html(v["yoy"])} &nbsp;·&nbsp; {_wd_html(v["wd_dev"])}</div>'
        )
    html = "".join(parts)
    icon = {"uv_spike": "🔎", "cert_vs_total": "⚖️", "wd_dev": "📌"}
    if res["flags"]:
        html += ('<div style="margin-top:.5rem;padding:.5rem .7rem;background:#f9fafb;'
                 'border-left:3px solid #6366f1;border-radius:4px">')
        html += "".join(f'<div style="margin:.2rem 0">{icon.get(fl["code"], "•")} {fl["text"]}</div>'
                        for fl in res["flags"])
        html += "</div>"
    else:
        html += '<div style="margin-top:.4rem;color:#9ca3af">특이 플래그 없음 — 전년·동요일 흐름과 대체로 일치.</div>'
    return html


def build_html(df, cur_all, metric, pay, has_prev, fmt_fn, asof) -> str:
    """현재 지표·결제구분 기준 주요 표를 정적 HTML 1개(자체완결)로. Styler.to_html 사용."""
    import html as _h
    pay_l = "순결제" if pay == "net" else "총결제"
    months = sorted(cur_all.month.unique())
    blocks = []

    # 1) 제휴사 × 월 (당년 + 전년비)
    idx = list(C.agg_value(cur_all, "affiliate", metric, pay).sort_values(ascending=False).index)
    m1 = C.value_matrix(cur_all, "affiliate", "month", metric, pay, idx, months)
    d1 = m1.copy()
    d1.columns = ["합계" if c == "합계" else C.m_label(c) for c in d1.columns]
    blocks.append(("제휴사 × 월", style_values(d1, fmt_fn).to_html()))
    if has_prev:
        dp = df[df.year_tag == "prev"].copy()
        dp["month"] = dp["month"].map(lambda s: C.m_shift(s, 1))
        pm = C.value_matrix(dp, "affiliate", "month", metric, pay, idx, months)
        yy = C.yoy_matrix(m1, pm)
        yy.columns = d1.columns
        blocks.append(("제휴사 × 월 (전년비 %)", style_yoy(yy).to_html()))

    # 2) 날짜 × 제휴사 (최근월)
    last = months[-1]
    dc = cur_all[cur_all.month == last]
    aff = list(C.agg_value(dc, "affiliate", metric, pay).sort_values(ascending=False).index)
    m2 = C.value_matrix(dc, "day", "affiliate", metric, pay, sorted(dc.day.unique()), aff)
    d2 = m2.copy()
    d2.index = ["합계" if i == "합계" else f"{int(i)}일" for i in d2.index]
    blocks.append((f"날짜 × 제휴사 · {last.replace('-', '.')}", style_values(d2, fmt_fn).to_html()))

    # 3) 전체 제휴사 · 일자 × 월
    m3 = C.value_matrix(cur_all, "day", "month", metric, pay, sorted(cur_all.day.unique()), months)
    d3 = m3.copy()
    d3.index = ["합계" if i == "합계" else f"{int(i)}일" for i in d3.index]
    d3.columns = ["합계" if c == "합계" else C.m_label(c) for c in d3.columns]
    blocks.append(("전체 제휴사 · 일자 × 월", style_values(d3, fmt_fn).to_html()))

    body = "\n".join(f"<h2>{_h.escape(t)}</h2>{h}" for t, h in blocks)
    return (
        '<!doctype html><html lang="ko"><head><meta charset="utf-8">'
        "<title>LF몰 제휴 일자별 실적 집계</title><style>"
        "body{font-family:system-ui,'Malgun Gothic',sans-serif;margin:24px;color:#111}"
        "h1{font-size:1.3rem;margin-bottom:.2rem}"
        "h2{font-size:1rem;margin:1.6rem 0 .4rem;border-left:4px solid #6366f1;padding-left:8px}"
        "table{border-collapse:collapse;font-size:.78rem}"
        "td,th{border:1px solid #e5e7eb;padding:3px 7px;text-align:right;white-space:nowrap}"
        ".cap{color:#6b7280;font-size:.8rem}</style></head><body>"
        "<h1>🗓️ LF몰 제휴 일자별 실적 집계</h1>"
        f'<div class="cap">지표: {_h.escape(metric)} · {pay_l} · 기준 {_h.escape(asof or "-")} '
        "· 전년비 상승 ▼(초록)/하락 △(빨강)</div>"
        f"{body}"
        '<div class="cap" style="margin-top:1.6rem">※ (날짜×제휴사) 집계 기반 · 세그·목표비 제외 · 개인정보 없음</div>'
        "</body></html>"
    )


# ══════════════════════════════════════════════════════════
st.title("🗓️ LF몰 제휴 일자별 실적 집계")
st.caption("본 대시보드(kimhyemin)에서 내보낸 (날짜×제휴사) 집계만 읽는 경량 앱 · "
           "전년비 상승 ▼(초록)/하락 △(빨강)")

with st.sidebar:
    st.header("데이터")
    up = st.file_uploader("피벗 JSON 업로드 (daily_pivot_*.json)", type=["json"])
    st.caption("본 대시보드 사이드바 → **일자별 피벗 JSON (경량 앱용)** 버튼으로 내려받은 파일.")

raw, src = None, ""
if up is not None:
    raw, src = up.getvalue(), "업로드"
elif os.path.exists(BUNDLED):
    with open(BUNDLED, "rb") as f:
        raw, src = f.read(), "레포 내장(data/daily_pivot.json)"

if raw is None:
    st.info("① 본 대시보드(kimhyemin)에서 **일자별 피벗 JSON**을 내려받아 왼쪽에 업로드하세요.\n\n"
            "→ (날짜×제휴사) 집계만 담긴 파일이라 개인정보가 없고, 업로드하면 즉시 피벗됩니다.")
    st.stop()

df, payload = load(raw)
if df.empty or "cur" not in set(df.year_tag.unique()):
    st.warning("레코드가 비어 있거나 당년(cur) 데이터가 없습니다. JSON을 확인해 주세요.")
    st.stop()

_asof = payload.get("generated_mtd_end", "")
st.sidebar.success(f"로드: {src}" + (f" · 기준 {_asof}" if _asof else ""))

cur_all = df[df.year_tag == "cur"]
has_prev = "prev" in set(df.year_tag.unique())

with st.sidebar:
    st.header("지표")
    metric = st.selectbox("지표 선택", list(C.METRICS.keys()), index=2)
    pay_on = C.METRICS[metric]["pay"]
    pay = "net" if st.radio("결제 구분", ["순결제", "총결제"], index=0, horizontal=True,
                            disabled=not pay_on,
                            help="거래액·고객수·객단가에만 적용(UV·인증자수는 무관)") == "순결제" else "tot"
    show_yoy = st.checkbox("전년비 %로 보기", value=False, disabled=not has_prev,
                           help="켜면 값 대신 전년 동기 대비 증감률(▼초록/△빨강)로 표시")
    fmt_fn = C.FMT[C.METRICS[metric]["fmt"]]

with st.sidebar:
    st.header("내보내기")
    try:
        _report = build_html(df, cur_all, metric, pay, has_prev, fmt_fn, _asof)
        st.download_button(
            "⬇️ HTML 내보내기", data=_report.encode("utf-8"),
            file_name=f"lf_daily_{(_asof or 'export').replace('-', '')}.html",
            mime="text/html", use_container_width=True,
            help="현재 지표·결제구분 기준 주요 표(제휴사×월+전년비, 날짜×제휴사 최근월, "
                 "전체 일자×월)를 정적 HTML 1개로. 앱 없이 열람·공유 가능.")
    except Exception as _he:
        st.caption(f"⚠️ HTML 생성 오류: {_he}")

tab1, tab_af, tab2, tab3 = st.tabs(
    ["📅 날짜 × 제휴사", "🔎 제휴사 · 일자×월", "🏢 제휴사 × 월", "💬 분석일 코멘트"])

# ── 뷰1: 날짜 × 제휴사 (한 달) ──
with tab1:
    months = sorted(cur_all.month.unique())
    sel_m = st.selectbox("월", months, index=len(months) - 1,
                         format_func=lambda m: m.replace("-", "년 ") + "월")
    dcur = cur_all[cur_all.month == sel_m]
    aff_tot = C.agg_value(dcur, "affiliate", metric, pay).sort_values(ascending=False)
    col_order = [a for a in aff_tot.index if not C.na(aff_tot[a])]
    idx_order = sorted(dcur.day.unique())
    cur_mat = C.value_matrix(dcur, "day", "affiliate", metric, pay, idx_order, col_order)
    disp = cur_mat.copy()
    disp.index = ["합계" if i == "합계" else f"{int(i)}일" for i in disp.index]
    st.markdown(f"**{sel_m.replace('-', '.')} · {metric}** · 행=일 / 열=제휴사 "
                f"(제휴사는 {'순결제' if pay=='net' else '총결제'} 기준 큰 순)")
    if show_yoy and has_prev:
        dprev = df[(df.year_tag == "prev") & (df.month == C.m_shift(sel_m, -1))]
        prev_mat = C.value_matrix(dprev, "day", "affiliate", metric, pay, idx_order, col_order)
        yy = C.yoy_matrix(cur_mat, prev_mat)
        yy.index = disp.index
        st.dataframe(style_yoy(yy), use_container_width=True, height=560)
    else:
        st.dataframe(style_values(disp, fmt_fn), use_container_width=True, height=560)

# ── 뷰: 제휴사 · 일자 × 월 (한 제휴사 필터 + 월 다중선택) ──
with tab_af:
    all_months = sorted(cur_all.month.unique())
    ac1, ac2 = st.columns([1.2, 2])
    with ac1:
        aff_order = list(C.agg_value(cur_all, "affiliate", metric, pay)
                         .sort_values(ascending=False).index)
        af = st.selectbox("제휴사", ["전체"] + aff_order, index=0, key="afdx_af")
    with ac2:
        sel_ms = st.multiselect("월 (다중 선택)", all_months,
                                default=all_months[-3:] if len(all_months) >= 3 else all_months,
                                format_func=lambda m: m.replace("-", "년 ") + "월", key="afdx_ms")
    if not sel_ms:
        sel_ms = all_months[-1:]
    sel_ms = sorted(sel_ms)
    d = cur_all[cur_all.month.isin(sel_ms)]
    if af != "전체":
        d = d[d.affiliate == af]
    idx_order = sorted(d.day.unique())
    cur_mat = C.value_matrix(d, "day", "month", metric, pay, idx_order, sel_ms)
    disp = cur_mat.copy()
    disp.index = ["합계" if i == "합계" else f"{int(i)}일" for i in disp.index]
    disp.columns = ["합계" if c == "합계" else C.m_label(c) for c in disp.columns]
    st.markdown(f"**{af} · {metric}** · 행=일 / 열=월 "
                f"({'순결제' if pay=='net' else '총결제'} 기준) · "
                f"{' · '.join(m.replace('-', '.') for m in sel_ms)}")
    if show_yoy and has_prev:
        dprev = df[df.year_tag == "prev"].copy()
        if af != "전체":
            dprev = dprev[dprev.affiliate == af]
        dprev["month"] = dprev["month"].map(lambda s: C.m_shift(s, 1))  # 전년 → 당년 라벨
        dprev = dprev[dprev.month.isin(sel_ms)]
        prev_mat = C.value_matrix(dprev, "day", "month", metric, pay, idx_order, sel_ms)
        yy = C.yoy_matrix(cur_mat, prev_mat)
        yy.index, yy.columns = disp.index, disp.columns
        st.dataframe(style_yoy(yy), use_container_width=True, height=620)
    else:
        st.dataframe(style_values(disp, fmt_fn), use_container_width=True, height=620)
    st.caption("한 제휴사의 여러 달 일자별 흐름을 나란히 비교(예: 삼성카드 5·6·7월 UV). "
               "전년비 토글 시 각 셀=전년 동일 캘린더일 대비.")

# ── 뷰2: 제휴사 × 월 ──
with tab2:
    idx_order = list(C.agg_value(cur_all, "affiliate", metric, pay).sort_values(ascending=False).index)
    col_order = sorted(cur_all.month.unique())
    cur_mat = C.value_matrix(cur_all, "affiliate", "month", metric, pay, idx_order, col_order)
    disp = cur_mat.copy()
    disp.columns = ["합계" if c == "합계" else C.m_label(c) for c in disp.columns]
    st.markdown(f"**{metric}** · 행=제휴사 / 열=월 ({'순결제' if pay=='net' else '총결제'} 기준)")
    if show_yoy and has_prev:
        dprev = df[df.year_tag == "prev"].copy()
        dprev["month"] = dprev["month"].map(lambda s: C.m_shift(s, 1))  # 전년 → 당년 라벨로 정렬
        prev_mat = C.value_matrix(dprev, "affiliate", "month", metric, pay, idx_order, col_order)
        yy = C.yoy_matrix(cur_mat, prev_mat)
        yy.columns = disp.columns
        st.dataframe(style_yoy(yy), use_container_width=True, height=560)
    else:
        st.dataframe(style_values(disp, fmt_fn), use_container_width=True, height=560)

# ── 뷰3: 분석일 코멘트 ──
with tab3:
    dates = sorted(cur_all.date.unique())
    c1, c2 = st.columns([1, 1])
    with c1:
        day_date = st.selectbox("분석일", dates, index=len(dates) - 1)
    with c2:
        affs = ["전체"] + list(C.agg_value(cur_all, "affiliate", "당월인증거래액", "net")
                               .sort_values(ascending=False).index)
        aff_sel = st.selectbox("제휴사", affs, index=0)
    st.markdown(f"#### {day_date} · {aff_sel} · {'순결제' if pay=='net' else '총결제'}")
    res = C.analyze_day(df, day_date, aff_sel, pay)
    st.markdown(comment_html(res), unsafe_allow_html=True)
    st.caption("전년비=전년 동일 캘린더일 · 동요일평균=같은 달 내 같은 요일 평균(전일비 아님) · 세그·목표비 제외")
