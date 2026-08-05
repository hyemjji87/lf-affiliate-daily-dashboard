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


def zero_perf_affiliates(cur_df) -> list:
    """당월인증거래액(순결제) 합계가 0인 제휴사 = 실적 없음. 「리스팅 제외」 기본 포함 후보.
    ※ pivot_core가 아닌 여기(항상 새로 실행되는 메인 스크립트)에 정의 —
    Streamlit 재배포가 import된 헬퍼 모듈을 옛 캐시로 두어 새 함수를 못 찾던 문제 회피."""
    if cur_df.empty or "cert_amt_net" not in cur_df.columns:
        return []
    s = cur_df.groupby("affiliate")["cert_amt_net"].sum()
    return [a for a in s.index if float(s.get(a, 0) or 0) == 0]


def prev_matched(all_df, months, asof, filter_af=None):
    """전년 레코드를 당년과 '같은 기간'으로 맞춰 반환 — 전년비/합계 정합용.
    ① 전년 월 라벨을 +1년 시프트해 당년 월과 맞추고, ② 표시된 월(months)로 제한,
    ③ as-of 월은 as-of 일자까지만 남긴다(MTD 대 MTD).
    ※ value_matrix의 합계는 넘긴 df 전체 월로 합산하므로, 여기서 기간을 미리 잘라야
    합계·전년비가 당년과 같은 기간끼리 비교된다(전년 8~12월이 분모에 섞이던 버그 방지)."""
    dp = all_df[all_df.year_tag == "prev"].copy()
    if filter_af is not None:
        dp = dp[dp.affiliate == filter_af]
    dp["month"] = dp["month"].map(lambda s: C.m_shift(s, 1))
    dp = dp[dp["month"].isin(months)]
    if len(asof) >= 10 and asof[8:10].isdigit():
        am, ad = asof[:7], int(asof[8:10])
        dp = dp[~((dp["month"] == am) & (dp["day"] > ad))]
    return dp


# 금액 포맷은 원단위(콤마)로 — pivot_core.fmt_won의 억/만 축약 대신 여기(항상 새로 실행되는
# 메인)에서 오버라이드해, 재배포 모듈 캐시와 무관하게 즉시 반영되게 한다.
def fmt_won_full(v):
    return "" if C.na(v) else f"{float(v):,.0f}"


# 인증률(%) 포맷도 여기(메인)에서 로컬 정의 — 재배포 시 pivot_core가 옛 캐시로 남아
# C.fmt_pct를 못 찾고 import 단계에서 앱 전체가 죽던 문제 회피(fmt_won_full과 동일 패턴).
def fmt_pct(v):
    return "" if C.na(v) else f"{float(v) * 100:.1f}%"


FMT = {"int": C.fmt_int, "won": fmt_won_full, "won0": C.fmt_int, "pct": fmt_pct}

CUST_METRICS = ("고객수(당월인증)", "객단가(당월인증)")


def cust_month_matrix(mc_year, pay, index_order, col_order):
    """monthly_cust(한 연도)로 (제휴사×월) '월 순수고객수' 매트릭스 + 합계.
    셀(제휴사×월)은 정확. 합계(행/열/모서리)는 월·제휴사 순수고객의 단순합이라
    교차기간·교차제휴사 중복이 있을 수 있음(근사)."""
    col = "cust_net" if pay == "net" else "cust_tot"
    cols = ["합계"] + list(col_order)
    if mc_year is None or mc_year.empty:
        return pd.DataFrame(index=["합계"] + list(index_order), columns=cols, dtype=float)
    piv = mc_year.pivot_table(index="affiliate", columns="month", values=col, aggfunc="sum") \
                 .reindex(index=index_order, columns=col_order)
    row_tot = mc_year.groupby("affiliate")[col].sum().reindex(index_order)
    col_tot = mc_year.groupby("month")[col].sum().reindex(col_order)
    grand = float(mc_year[col].sum())
    body = piv.copy()
    body.insert(0, "합계", row_tot)
    top = pd.DataFrame([[grand] + [col_tot.get(c, np.nan) for c in col_order]],
                       index=["합계"], columns=body.columns)
    return pd.concat([top, body])


# ── 스타일 ──────────────────────────────────────────────
def style_values(mat, fmt_fn):
    return (mat.style
            .format(fmt_fn, na_rep="")
            .apply(lambda s: ["font-weight:700" if s.name == "합계" else "" for _ in s], axis=1)
            .set_properties(subset=["합계"], **{"font-weight": "700"}))


def style_yoy(mat):
    def f(v):
        return "" if C.na(v) else f"{'▼' if v >= 0 else '△'} {v:+.1f}%"

    def color(v):
        return "" if C.na(v) else (f"color:{UP_COLOR}" if v >= 0 else f"color:{DOWN_COLOR}")

    return mat.style.format(f, na_rep="").map(color)


def _hl_row(styler, label):
    """일자 매트릭스에서 분석일자('N일') 행을 배경색으로 강조."""
    if not label:
        return styler
    return styler.apply(
        lambda s: ["background-color:#fff3bf" if s.name == label else "" for _ in s], axis=1)


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


def _pct1(v):
    return "–" if v is None else f"{v:.1f}%"


def _returns_html(r: dict) -> str:
    """당월인증거래액 총결제·순결제 대조 + 취소·환불(취반품) 요약 한 줄."""
    if r is None or (C.na(r.get("net_cur")) and C.na(r.get("tot_cur"))):
        return ""
    tot, net = fmt_won_full(r["tot_cur"]), fmt_won_full(r["net_cur"])
    seg = (f'<b>당월인증거래액 총/순 대조</b> 총결제 {tot} '
           f'({_yoy_html(r["tot_yoy"])}) &nbsp;·&nbsp; 순결제 {net} ({_yoy_html(r["net_yoy"])})')
    if not C.na(r.get("cancel_cur")):
        cr = r.get("cancel_ratio_cur")
        crp = r.get("cancel_ratio_prev")
        crtxt = f"취소율 {_pct1(cr * 100 if cr is not None else None)}"
        if crp is not None:
            crtxt += f" (전년 {_pct1(crp * 100)})"
        seg += f' &nbsp;·&nbsp; 취소·환불 {fmt_won_full(r["cancel_cur"])} · {crtxt}'
    return f'<div style="margin:.15rem 0;color:#374151">↳ {seg}</div>'


def _lfshare_html(ls: dict) -> str:
    """LF몰 전체거래액 대비 순결제 당월인증거래액 비중 한 줄(순결제일 때만 의미)."""
    if ls is None or ls.get("share_cur") is None:
        return ""
    body = f'<b>LF몰 대비 비중</b> {_pct1(ls["share_cur"])}'
    if ls.get("share_prev") is not None:
        body += f' &nbsp;·&nbsp; 전년 {_pct1(ls["share_prev"])}'
        dpp = ls.get("dpp")
        if dpp is not None:
            body += " &nbsp;·&nbsp; " + _cspan(f"{'▼' if dpp >= 0 else '△'} {dpp:+.2f}%p", dpp >= 0)
    return f'<div style="margin:.15rem 0;color:#374151">↳ {body} <span style="color:#9ca3af">(순결제·당월인증거래액 기준)</span></div>'


def comment_html(res: dict) -> str:
    parts = []
    for metric, v in res["metrics"].items():
        fmt_fn = FMT[C.METRICS[metric]["fmt"]]
        parts.append(
            f'<div style="margin:.15rem 0"><b>{metric}</b> {fmt_fn(v["cur"])} '
            f'&nbsp;·&nbsp; {_yoy_html(v["yoy"])} &nbsp;·&nbsp; {_wd_html(v["wd_dev"])}</div>'
        )
        # 당월인증거래액 바로 아래에 총/순 대조·LF몰 비중을 붙여 맥락을 준다.
        if metric == "당월인증거래액":
            parts.append(_returns_html(res.get("returns")))
            parts.append(_lfshare_html(res.get("lf_share")))
    html = "".join(parts)
    icon = {"uv_spike": "🔎", "cert_vs_total": "⚖️", "cancel": "🔁", "wd_dev": "📌"}
    if res["flags"]:
        html += ('<div style="margin-top:.5rem;padding:.5rem .7rem;background:#f9fafb;'
                 'border-left:3px solid #6366f1;border-radius:4px">')
        html += "".join(f'<div style="margin:.2rem 0">{icon.get(fl["code"], "•")} {fl["text"]}</div>'
                        for fl in res["flags"])
        html += "</div>"
    else:
        html += '<div style="margin-top:.4rem;color:#9ca3af">특이 플래그 없음 — 전년·동요일 흐름과 대체로 일치.</div>'
    return html


def build_html(df, cur_all, has_prev, asof, pay="net", lfmall_daily=None) -> str:
    """UV·인증자수·당월인증거래액(총결제) 리포트를 정적 HTML 1개로. Total 코멘트 포함.
    각 지표 표는 그 지표 합계가 0인 제휴사를 리스팅에서 제외한다.
    Total 코멘트는 사이드바 결제구분(pay)을 그대로 따르고, 총/순 대조·LF몰 비중을 함께 담는다."""
    import html as _h
    months = sorted(cur_all.month.unique())
    asof_day = int(asof[8:10]) if len(asof) >= 10 and asof[8:10].isdigit() else None
    hl = f"{asof_day}일" if asof_day else None  # 일자 매트릭스 분석일자 강조 라벨
    pay_label = "순결제" if pay == "net" else "총결제"

    sections = []

    # 0) Total 코멘트 (최근 분석일 · 전체 · 사이드바 결제구분)
    last_date = max(cur_all.date)
    res = C.analyze_day(df, last_date, "전체", pay, lfmall_daily)
    sections.append((f"📌 Total 코멘트 · {last_date} · 전체 · {pay_label}", comment_html(res)))

    # 지표별: 당월인증거래액은 총결제(tot), UV·인증자수는 결제구분 무관(net)
    for metric, pay in [("UV", "net"), ("인증자수", "net"), ("당월인증거래액", "tot")]:
        fmt_fn = FMT[C.METRICS[metric]["fmt"]]
        suffix = " (총결제)" if metric == "당월인증거래액" else ""
        order = list(C.agg_value(cur_all, "affiliate", metric, pay)
                     .sort_values(ascending=False).index)

        # 제휴사 × 월 (당년 + 전년비)
        m1 = C.value_matrix(cur_all, "affiliate", "month", metric, pay, order, months)
        d1 = m1.copy()
        d1.columns = ["합계" if c == "합계" else C.m_label(c) for c in d1.columns]
        sections.append((f"{metric}{suffix} · 제휴사 × 월", style_values(d1, fmt_fn).to_html()))
        if has_prev:
            dp = prev_matched(df, months, asof)
            pm = C.value_matrix(dp, "affiliate", "month", metric, pay, order, months)
            yy = C.yoy_matrix(m1, pm)
            yy.columns = d1.columns
            sections.append((f"{metric}{suffix} · 제휴사 × 월 (전년비 %)", style_yoy(yy).to_html()))

        # 전체 제휴사 · 일자 × 월 (일자별 흐름)
        m3 = C.value_matrix(cur_all, "day", "month", metric, pay, sorted(cur_all.day.unique()), months)
        d3 = m3.copy()
        d3.index = ["합계" if i == "합계" else f"{int(i)}일" for i in d3.index]
        d3.columns = ["합계" if c == "합계" else C.m_label(c) for c in d3.columns]
        sections.append((f"{metric}{suffix} · 전체 일자 × 월",
                         _hl_row(style_values(d3, fmt_fn), hl).to_html()))

    body = "\n".join(f"<h2>{_h.escape(t)}</h2>{h}" for t, h in sections)
    return (
        '<!doctype html><html lang="ko"><head><meta charset="utf-8">'
        "<title>LF몰 제휴 일자별 실적 집계</title><style>"
        "body{font-family:system-ui,'Malgun Gothic',sans-serif;margin:24px;color:#111;font-size:14px;line-height:1.45}"
        "h1{font-size:1.45rem;margin:0 0 .3rem}"
        "h2{font-size:1.05rem;margin:1.9rem 0 .5rem;border-left:4px solid #6366f1;padding-left:9px}"
        "table{border-collapse:collapse;font-size:.85rem;font-variant-numeric:tabular-nums;margin:.3rem 0 1.1rem}"
        "th,td{border:1px solid #e5e7eb;padding:5px 10px;white-space:nowrap}"
        ".data{text-align:right}"
        ".col_heading,.blank{background:#eef1f6;text-align:center;font-weight:600}"
        ".row_heading{text-align:left;font-weight:500}"
        "tbody tr:hover .data{background:#f6f7f9}"
        ".cap{color:#6b7280;font-size:.82rem;margin:.15rem 0}</style></head><body>"
        "<h1>🗓️ LF몰 제휴 일자별 실적 집계</h1>"
        f'<div class="cap">지표: UV · 인증자수 · 당월인증거래액(총결제) · 기준 {_h.escape(asof or "-")} '
        "· 전년비 상승 ▼(초록)/하락 △(빨강) · 제휴사 리스팅은 UV·당월인증거래액 합계 0 제외</div>"
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

# (year_tag, month, affiliate) 월 순수고객수 — 있으면 월별 뷰의 고객수·객단가를 정확히 계산.
# 없으면(구 v1 JSON) 일별 nunique 합으로 대체(월 순수고객 과다계상 가능).
mc = pd.DataFrame(payload.get("monthly_cust", []))
HAS_MC = not mc.empty

# 일자별 LF몰 전체거래액 {date: 원}. 있으면 분석일 코멘트에 LF몰 대비 비중을 표시.
# 없으면(구 JSON) 비중 줄은 자동 생략 — kimhyemin 재내보내기 필요.
LFMALL_DAILY = payload.get("lfmall_daily", {}) or {}

# 표·합계에서 제외할 비리스팅 파트너. 기본값은 내보낸 JSON의 exclude_listing(비공개 앱에서 지정).
# 위젯 상태 지속 문제(한 번 굳으면 이후 default 무시)를 피하려고, 데이터가 바뀌면
# 세션 상태에 기본 제외목록을 강제 주입한 뒤 위젯을 key로 바인딩한다.
_avail_af = sorted(df.affiliate.unique())
_excl_listing = payload.get("exclude_listing", []) or []
# 기본 제외 = ① JSON의 비리스팅 파트너(exclude_listing) + ② 당월인증거래액(순결제) 실적 0인 제휴사.
# ②는 숨기지 않고 칩으로 노출해 사용자가 빼고 싶으면 뺄 수 있게 한다.
_zero_af = zero_perf_affiliates(df[df.year_tag == "cur"])
_excl_default = [a for a in _avail_af if a in set(_excl_listing) | set(_zero_af)]
_excl_sig = f"{_asof}|{len(df)}|{len(_avail_af)}"
if st.session_state.get("_excl_sig") != _excl_sig:
    st.session_state["_excl_sig"] = _excl_sig
    st.session_state["excl_ms"] = _excl_default
with st.sidebar:
    st.header("제외 제휴사")
    excl = st.multiselect(
        "리스팅 제외", _avail_af, key="excl_ms",
        help="선택 제휴사를 UV·인증자수·당월인증거래액 등 모든 표·합계에서 제외. 기본값 = "
             "내보낸 JSON의 exclude_listing + 당월인증거래액 실적 0인 제휴사(빼고 싶으면 칩에서 제거).")
    _n_zero = len([a for a in _zero_af if a in _avail_af])
    _cap = []
    if _excl_listing:
        _miss = [a for a in _excl_listing if a not in _avail_af]
        _cap.append(f"기본 제외목록 {len(_excl_listing)}개" + (f"(미적용: {', '.join(_miss)})" if _miss else ""))
    else:
        _cap.append("⚠️ JSON에 기본 제외목록 없음 — kimhyemin 재내보내기 권장")
    _cap.append(f"실적(당월인증거래액) 0 제휴사 {_n_zero}곳 자동 포함")
    st.caption(" · ".join(_cap))
if excl:
    df = df[~df.affiliate.isin(excl)]
    if HAS_MC:
        mc = mc[~mc.affiliate.isin(excl)]

cur_all = df[df.year_tag == "cur"]
has_prev = "prev" in set(df.year_tag.unique())

# 분석일자(as-of) 강조용 행 라벨('N일'). 일자 매트릭스에서 해당 일 행을 컬러 강조.
_asof_day = int(_asof[8:10]) if len(_asof) >= 10 and _asof[8:10].isdigit() else None
HL_LABEL = f"{_asof_day}일" if _asof_day else None

with st.sidebar:
    st.header("지표")
    metric = st.selectbox("지표 선택", list(C.METRICS.keys()), index=2)
    pay_on = C.METRICS[metric]["pay"]
    pay = "net" if st.radio("결제 구분", ["순결제", "총결제"], index=0, horizontal=True,
                            disabled=not pay_on,
                            help="거래액·고객수·객단가에만 적용(UV·인증자수는 무관)") == "순결제" else "tot"
    show_yoy = st.checkbox("전년비 %로 보기", value=False, disabled=not has_prev,
                           help="켜면 값 대신 전년 동기 대비 증감률(▼초록/△빨강)로 표시")
    fmt_fn = FMT[C.METRICS[metric]["fmt"]]

with st.sidebar:
    st.header("내보내기")
    try:
        _report = build_html(df, cur_all, has_prev, _asof, pay, LFMALL_DAILY)
        st.download_button(
            "⬇️ HTML 내보내기", data=_report.encode("utf-8"),
            file_name=f"lf_daily_{(_asof or 'export').replace('-', '')}.html",
            mime="text/html", use_container_width=True,
            help="Total 코멘트 + UV·인증자수·당월인증거래액(총결제) 리포트(제휴사×월+전년비, "
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
        dprev = prev_matched(df, [sel_m], _asof)
        prev_mat = C.value_matrix(dprev, "day", "affiliate", metric, pay, idx_order, col_order)
        yy = C.yoy_matrix(cur_mat, prev_mat)
        yy.index = disp.index
        st.dataframe(_hl_row(style_yoy(yy), HL_LABEL), use_container_width=True, height=560)
    else:
        st.dataframe(_hl_row(style_values(disp, fmt_fn), HL_LABEL), use_container_width=True, height=560)

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
        dprev = prev_matched(df, sel_ms, _asof, filter_af=(None if af == "전체" else af))
        prev_mat = C.value_matrix(dprev, "day", "month", metric, pay, idx_order, sel_ms)
        yy = C.yoy_matrix(cur_mat, prev_mat)
        yy.index, yy.columns = disp.index, disp.columns
        st.dataframe(_hl_row(style_yoy(yy), HL_LABEL), use_container_width=True, height=620)
    else:
        st.dataframe(_hl_row(style_values(disp, fmt_fn), HL_LABEL), use_container_width=True, height=620)
    st.caption("한 제휴사의 여러 달 일자별 흐름을 나란히 비교(예: 삼성카드 5·6·7월 UV). "
               "전년비 토글 시 각 셀=전년 동일 캘린더일 대비.")

    # 옵션 라인차트 — x=일 / 라인=월. 시리즈가 '월'뿐이라 겹쳐 보기 좋음(표는 그대로 원본).
    # 값 모드에서만 노출(전년비 %는 표가 더 읽기 쉬움).
    if not (show_yoy and has_prev):
        with st.expander("📈 일자 추이 차트 (x=일 / 라인=월)", expanded=False):
            chart_piv = C.pivot_values(d, "day", "month", metric, pay)
            if chart_piv is None or chart_piv.empty:
                st.caption("표시할 데이터가 없습니다.")
            else:
                chart_piv = chart_piv.reindex(index=idx_order, columns=sel_ms)
                chart_piv.columns = [C.m_label(c) for c in chart_piv.columns]
                chart_piv.index = [int(i) for i in chart_piv.index]
                chart_piv.index.name = "일"
                st.line_chart(chart_piv)
                st.caption(f"{af} · {metric}({'순결제' if pay=='net' else '총결제'}) · "
                           "x=일 / 라인=월 · 값 기준(전년비% 모드에선 숨김)")
                if len(sel_ms) >= 5:
                    st.caption("⚠️ 선택 월이 많아 라인이 겹칩니다 — 월을 3~4개로 줄이면 더 잘 보입니다.")

# ── 뷰2: 제휴사 × 월 ──
with tab2:
    col_order = sorted(cur_all.month.unique())
    # 고객수·객단가는 월 순수고객수(monthly_cust)로 계산 — 일별 nunique 합의 월 과다계상 방지.
    use_mc = metric in CUST_METRICS and HAS_MC
    cust_col = "cust_net" if pay == "net" else "cust_tot"

    def _mc_year(tag, shift=False):
        m = mc[mc.year_tag == tag].copy()
        if shift:
            m["month"] = m["month"].map(lambda s: C.m_shift(s, 1))
            m = m[m["month"].isin(col_order)]
        return m

    if use_mc:
        mc_cur = _mc_year("cur")
        idx_order = list(mc_cur.groupby("affiliate")[cust_col].sum().sort_values(ascending=False).index)
        cust_cur = cust_month_matrix(mc_cur, pay, idx_order, col_order)
        if metric == "고객수(당월인증)":
            cur_mat = cust_cur
        else:  # 객단가 = 당월인증거래액 / 월 순수고객수
            amt_cur = C.value_matrix(cur_all, "affiliate", "month", "당월인증거래액", pay, idx_order, col_order)
            cur_mat = amt_cur / cust_cur.replace(0, np.nan)
    else:
        idx_order = list(C.agg_value(cur_all, "affiliate", metric, pay)
                         .sort_values(ascending=False).index)
        cur_mat = C.value_matrix(cur_all, "affiliate", "month", metric, pay, idx_order, col_order)
    disp = cur_mat.copy()
    disp.columns = ["합계" if c == "합계" else C.m_label(c) for c in disp.columns]
    _note = " · 고객수=월 순수고객(합계는 단순합·중복가능)" if use_mc else ""
    st.markdown(f"**{metric}** · 행=제휴사 / 열=월 ({'순결제' if pay=='net' else '총결제'} 기준){_note}")
    if show_yoy and has_prev:
        if use_mc:
            cust_prev = cust_month_matrix(_mc_year("prev", shift=True), pay, idx_order, col_order)
            if metric == "고객수(당월인증)":
                prev_mat = cust_prev
            else:
                amt_prev = C.value_matrix(prev_matched(df, col_order, _asof), "affiliate", "month",
                                          "당월인증거래액", pay, idx_order, col_order)
                prev_mat = amt_prev / cust_prev.replace(0, np.nan)
        else:
            dprev = prev_matched(df, col_order, _asof)
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
    res = C.analyze_day(df, day_date, aff_sel, pay, LFMALL_DAILY)
    st.markdown(comment_html(res), unsafe_allow_html=True)
    _lf_note = "" if LFMALL_DAILY else " · LF몰 비중은 JSON에 LF몰 일자별 거래액이 있어야 표시(kimhyemin 재내보내기)"
    st.caption("전년비=전년 동일 캘린더일 · 동요일평균=같은 달 내 같은 요일 평균(전일비 아님) · "
               "총/순 대조로 취반품 영향 판별 · 세그·목표비 제외" + _lf_note)
