"""
LF몰 제휴 월마감 분석 대시보드 (Streamlit)

실행:  streamlit run app.py   (monthly 폴더 기준)
설계: 일별 대시보드(../app.py)의 총결제/순결제·전년비(▼상승/△하락)·목표비 개념을 승계.
탭 4종: 개요(OverView) / 주차별 / 제휴사별 / BPU별.  차트 없이 표 중심.

업로드 5종
 ① 당년 raw (복수)      : 유입실적/인증회원/인증거래액/피벗 시트를 담은 월별 raw
 ② 전년 유입인증        : 전년도 UV·인증자수
 ③ 전년 매출분기 (복수) : 전년도 거래(인증거래액) 분기 파일
 ④ 목표                 : 월별 UV/인증자수/당월인증거래액/거래액 목표 (세그: 신규/윈백/기존)
 ⑤ LF몰 전체거래액      : 월별 목표/실제
"""
import json
import datetime as dt

import numpy as np
import pandas as pd
import streamlit as st

import lf_analysis as A
import monthly_core as M

st.set_page_config(page_title="LF몰 제휴 월마감", page_icon="📆", layout="wide")

EXPORT_SECTIONS: list[dict] = []
COLOR_MAP = {"success": "#1a7f37", "danger": "#cf222e", "muted": "#6e7781", "accent": "#0969da"}


# ----------------------------------------------------------------------------------
# 로더 (캐싱)
# ----------------------------------------------------------------------------------
def _sig(files) -> tuple:
    return tuple((f.name, f.size) for f in files) if files else tuple()


@st.cache_data(show_spinner=False)
def _load_raw(list_of_bytes: tuple) -> dict:
    if not list_of_bytes:
        return {"amt": pd.DataFrame(), "uv": pd.DataFrame(), "cert": pd.DataFrame(), "pivot": None}
    merged = A.merge_workbooks(list(list_of_bytes))
    return merged


@st.cache_data(show_spinner=False)
def _load_target_bytes(b: bytes):
    return M.load_target(b)


@st.cache_data(show_spinner=False)
def _load_lf_bytes(b: bytes):
    return M.load_lfmall(b)


def _read_files(files):
    """업로드 파일 → bytes 튜플. DRM(Softcamp 등) 암호화 파일은 걸러내고 경고 목록 반환."""
    good, bad = [], []
    for f in files or []:
        b = f.getvalue()
        if b[:2] == b"PK":  # 정상 xlsx(zip)
            good.append(b)
        else:
            bad.append(f.name)
    return tuple(good), bad


# ----------------------------------------------------------------------------------
# 포맷 유틸 (일별 대시보드와 동일 컨벤션)
# ----------------------------------------------------------------------------------
def fmt(n, unit=1):
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "-"
    return f"{round(n / unit):,}"


def pct_ratio(cur, prev):
    """전년비: 상승 ▼(초록)·하락 △(빨강) — 회사 내부 컨벤션."""
    if prev is None or prev == 0 or pd.isna(prev) or cur is None or pd.isna(cur):
        if prev is None or (isinstance(prev, float) and pd.isna(prev)):
            return "신규", "accent", None
        return "-", "muted", None
    ratio = cur / prev - 1
    pct = round(ratio * 100)
    if pct < 0:
        return f"△{abs(pct)}%", "danger", ratio
    return f"▼{pct}%", "success", ratio


def target_ratio(cur, target):
    if (target is None or (isinstance(target, float) and pd.isna(target)) or target == 0
            or cur is None or (isinstance(cur, float) and pd.isna(cur))):
        return "-", "muted", None
    pct = round(cur / target * 100)
    return f"{pct}%", ("success" if pct >= 100 else "danger"), cur / target


def _export_target() -> list:
    if not EXPORT_SECTIONS:
        EXPORT_SECTIONS.append({"name": "개요", "html": []})
    return EXPORT_SECTIONS[-1]["html"]


def export_tab(name: str):
    EXPORT_SECTIONS.append({"name": name, "html": []})


def md(html: str):
    st.markdown(html, unsafe_allow_html=True)
    _export_target().append(html)


def sub(title: str, note: str = ""):
    extra = f'<span style="font-size:12px;color:#6e7781;font-weight:400;margin-left:8px;">{note}</span>' if note else ""
    html = f"<div style='font-size:16px;font-weight:700;margin:20px 0 6px;'>{title}{extra}</div>"
    st.markdown(html, unsafe_allow_html=True)
    _export_target().append(html)


def render_grouped_table(header_groups, rows, color_rows=None, first_col_label="구분"):
    ths_top = ['<th rowspan="2" style="text-align:center;vertical-align:middle;position:sticky;left:0;'
               'border:1px solid #e1e4e8;padding:6px 10px;background:#f6f8fa;">' + first_col_label + '</th>']
    for g, subs in header_groups:
        ths_top.append(f'<th colspan="{len(subs)}" style="text-align:center;'
                       f'border:1px solid #e1e4e8;padding:6px 10px;background:#f6f8fa;">{g}</th>')
    ths_sub = []
    for g, subs in header_groups:
        for s in subs:
            ths_sub.append(f'<th style="text-align:center;border:1px solid #e1e4e8;'
                           f'padding:4px 10px;background:#fafbfc;font-weight:500;font-size:12px;">{s}</th>')
    body = []
    for ri, row in enumerate(rows):
        cells = [f'<td style="text-align:center;border:1px solid #e1e4e8;padding:4px 10px;'
                 f'white-space:nowrap;position:sticky;left:0;background:#fff;font-weight:600;">{row[0]}</td>']
        for ci, v in enumerate(row[1:]):
            style = "text-align:center;border:1px solid #e1e4e8;padding:4px 10px;white-space:nowrap;"
            if color_rows is not None and color_rows[ri][ci] in ("success", "danger", "accent"):
                style += f"color:{COLOR_MAP[color_rows[ri][ci]]};font-weight:600;"
            cells.append(f'<td style="{style}">{v}</td>')
        body.append(f'<tr>{"".join(cells)}</tr>')
    return ('<div style="overflow-x:auto;"><table style="border-collapse:collapse;width:100%;font-size:13px;">'
            f'<thead><tr>{"".join(ths_top)}</tr><tr>{"".join(ths_sub)}</tr></thead>'
            f'<tbody>{"".join(body)}</tbody></table></div>')


def render_simple_table(columns, rows, color_rows=None):
    ths = [f'<th style="text-align:center;border:1px solid #e1e4e8;padding:6px 10px;'
           f'background:#f6f8fa;">{c}</th>' for c in columns]
    body = []
    for ri, row in enumerate(rows):
        cells = [f'<td style="text-align:center;border:1px solid #e1e4e8;padding:4px 10px;'
                 f'white-space:nowrap;font-weight:600;">{row[0]}</td>']
        for ci, v in enumerate(row[1:]):
            style = "text-align:center;border:1px solid #e1e4e8;padding:4px 10px;white-space:nowrap;"
            if color_rows is not None and color_rows[ri][ci] in ("success", "danger", "accent"):
                style += f"color:{COLOR_MAP[color_rows[ri][ci]]};font-weight:600;"
            cells.append(f'<td style="{style}">{v}</td>')
        body.append(f'<tr>{"".join(cells)}</tr>')
    return ('<div style="overflow-x:auto;"><table style="border-collapse:collapse;width:100%;font-size:13px;">'
            f'<thead><tr>{"".join(ths)}</tr></thead><tbody>{"".join(body)}</tbody></table></div>')


def render_kpi_cards(cards) -> str:
    cells = []
    for label, val, yoy, yoy_c, extra in cards:
        yoy_css = f"color:{COLOR_MAP.get(yoy_c, '#24292f')};font-weight:600;"
        cells.append(
            '<div style="flex:1;min-width:170px;border:1px solid #e1e4e8;border-radius:8px;'
            'padding:14px 16px;background:#fff;">'
            f'<div style="font-size:13px;color:#6e7781;margin-bottom:4px;">{label}</div>'
            f'<div style="font-size:22px;font-weight:700;">{val}</div>'
            f'<div style="font-size:12px;margin-top:6px;white-space:nowrap;">'
            f'<span style="{yoy_css}">전년비 {yoy}</span>{extra}</div></div>')
    return f'<div style="display:flex;gap:12px;flex-wrap:wrap;margin:8px 0 18px;">{"".join(cells)}</div>'


# ==================================================================================
# 사이드바 — 업로드 & 설정
# ==================================================================================
st.sidebar.header("📁 Step 1. 파일 업로드")
up_cur = st.sidebar.file_uploader("① 당년 raw (복수 가능)", type=["xlsx"], accept_multiple_files=True, key="cur")
up_ly_inflow = st.sidebar.file_uploader("② 전년 유입·인증", type=["xlsx"], accept_multiple_files=True, key="lyi")
up_ly_amt = st.sidebar.file_uploader("③ 전년 매출분기 (복수 가능)", type=["xlsx"], accept_multiple_files=True, key="lya")
up_target = st.sidebar.file_uploader("④ 목표", type=["xlsx"], key="tgt")
up_lf = st.sidebar.file_uploader("⑤ LF몰 전체거래액", type=["xlsx"], key="lf")

st.sidebar.divider()

cur_bytes, cur_bad = _read_files(up_cur)
if not cur_bytes:
    st.title("📆 LF몰 제휴 월마감 분석")
    if cur_bad:
        st.error("⚠ 업로드 파일이 DRM 암호화(Softcamp 등) 상태로 보입니다. "
                 f"DRM 해제 후 다시 업로드해 주세요: {', '.join(cur_bad)}")
    st.info("왼쪽에서 **① 당년 raw** 파일을 업로드하면 분석이 시작됩니다.")
    st.stop()

cur_data = _load_raw(cur_bytes)
maps = M.parse_pivot(cur_data.get("pivot"))
M.enrich(cur_data, maps)

ly_i_bytes, _ = _read_files(up_ly_inflow)
ly_a_bytes, _ = _read_files(up_ly_amt)
ly_inflow = _load_raw(ly_i_bytes)
ly_amt = _load_raw(ly_a_bytes)
M.enrich(ly_inflow, maps)
M.enrich(ly_amt, maps)
prev_data = {"amt": ly_amt.get("amt", pd.DataFrame()),
             "uv": ly_inflow.get("uv", pd.DataFrame()),
             "cert": ly_inflow.get("cert", pd.DataFrame())}

target_df = _load_target_bytes(up_target.getvalue()) if up_target else pd.DataFrame()
lf_df = _load_lf_bytes(up_lf.getvalue()) if up_lf else pd.DataFrame()

months = M.available_months(cur_data)
if not months:
    st.error("업로드한 raw에서 월(마감) 정보를 찾지 못했습니다. 피벗 시트/정산일시일을 확인해 주세요.")
    st.stop()

st.sidebar.header("⚙ Step 2. 분석 설정")
sel_months = st.sidebar.multiselect("분석 월 (마감월)", months, default=months,
                                    format_func=lambda m: m.replace("-", "년") + "월")
if not sel_months:
    sel_months = months
primary = st.sidebar.selectbox("기준 월 (KPI/세그 표시)", sel_months, index=len(sel_months) - 1,
                               format_func=lambda m: m.replace("-", "년") + "월")

close_mode = st.sidebar.radio("기간 기준", ["월마감(전체)", "MTD(월초~기준일)"], index=0)
mtd_day = None
if close_mode.startswith("MTD"):
    mtd_day = st.sidebar.number_input("MTD 기준일 (일)", min_value=1, max_value=31, value=15, step=1)

paytype = st.sidebar.radio("결제 구분", ["순결제", "총결제"], index=0, horizontal=True)
pt = "net" if paytype == "순결제" else "tot"

warns = []
if cur_bad:
    warns.append(f"DRM 제외 {len(cur_bad)}개")
if not (up_ly_inflow or up_ly_amt):
    warns.append("전년 미업로드→전년비 제한")

st.title("📆 LF몰 제휴 월마감 분석")
st.caption(f"분석월 {', '.join(sel_months)} · 기준월 {primary} · {close_mode} · {paytype} 기준"
           + (" · ⚠ " + " / ".join(warns) if warns else ""))


# ----------------------------------------------------------------------------------
# 공통 집계 헬퍼
# ----------------------------------------------------------------------------------
def prev_ym(ym: str) -> str:
    y, m = ym.split("-")
    return f"{int(y) - 1}-{m}"


def cur_dates(ym):
    return M.month_dates(cur_data, ym, maps, mtd_day)


def prev_dates(ym):
    return M.month_dates(prev_data, prev_ym(ym), maps, mtd_day)


def bundle_cur(ym):
    return M.metric_bundle(cur_data, cur_dates(ym))


def bundle_prev(ym):
    return M.metric_bundle(prev_data, prev_dates(ym))


def aov(bundle, kind, seg="전체"):
    """객단가: 거래액 / 고객수. kind: 'amt_all'|'amt_cert'."""
    b = bundle[kind][seg]
    c = b[f"cust_{pt}"]
    return b[pt] / c if c else 0


def cr(bundle):
    """당월인증 CR = 당월인증고객수 / 인증자수."""
    cust = bundle["amt_cert"]["전체"][f"cust_{pt}"]
    cert = bundle["cert"]["전체"]
    return cust / cert if cert else 0


# 지표 정의: (라벨, getter(bundle)->값, 목표지표명 or None, 포맷단위)
METRICS = [
    ("UV", lambda b: b["uv"], None, 1),   # UV는 목표비 미표시(전년비만)
    ("인증자수", lambda b: b["cert"]["전체"], "인증자수", 1),
    ("당월인증고객수", lambda b: b["amt_cert"]["전체"][f"cust_{pt}"], None, 1),
    ("당월인증CR", lambda b: cr(b) * 100, None, 1),   # % 표시
    ("당월인증거래액", lambda b: b["amt_cert"]["전체"][pt], "당월인증거래액", 1_000_000),
    ("객단가", lambda b: aov(b, "amt_cert"), None, 1),
    ("전체거래액", lambda b: b["amt_all"]["전체"][pt], "거래액", 1_000_000),
]


tabs = st.tabs(["📊 개요", "📅 주차별", "🤝 제휴사별", "🏷 BPU별"])

# ==================================================================================
# 개요 탭
# ==================================================================================
with tabs[0]:
    export_tab("개요")
    bp = bundle_cur(primary)
    bpp = bundle_prev(primary)

    # KPI 카드 (기준월)
    def _kpi(label, cur_v, prev_v, tgt_metric, seg="전체", unit=1, suffix=""):
        d, dc, _ = pct_ratio(cur_v, prev_v)
        extra = ""
        if tgt_metric:
            t = M.target_value(target_df, tgt_metric, primary, seg)
            td, tc, _ = target_ratio(cur_v, t)
            extra = f'&nbsp;·&nbsp;<span style="color:{COLOR_MAP.get(tc)};font-weight:600;">목표 {td}</span>'
        return (label, fmt(cur_v, unit) + suffix, d, dc, extra)

    sub("핵심 지표 (KPI)", f"{primary} · {paytype}")
    cards = [
        _kpi("UV", bp["uv"], bpp["uv"], None),
        _kpi("인증자수", bp["cert"]["전체"], bpp["cert"]["전체"], "인증자수"),
        _kpi("당월인증거래액 (백만)", bp["amt_cert"]["전체"][pt], bpp["amt_cert"]["전체"][pt], "당월인증거래액", unit=1_000_000),
        _kpi("전체거래액 (백만)", bp["amt_all"]["전체"][pt], bpp["amt_all"]["전체"][pt], "거래액", unit=1_000_000),
    ]
    md(render_kpi_cards(cards))
    # LF몰 대비 비중
    lf_act = M.lfmall_value(lf_df, primary, "실제")
    if lf_act:
        share_all = bp["amt_all"]["전체"][pt] / lf_act * 100
        share_cert = bp["amt_cert"]["전체"][pt] / lf_act * 100
        md(f'<div style="font-size:13px;color:#57606a;margin:-8px 0 4px;">LF몰 전체거래액 대비 · '
           f'전체거래액 <b>{share_all:.1f}%</b> · 당월인증거래액 <b>{share_cert:.1f}%</b></div>')

    # 월별 실적 (행=월, 그룹=지표 실적/전년비)
    sub("월별 실적", f"{paytype} · 거래액/객단가 단위: 당월인증거래액·전체거래액=백만, 객단가=원")
    groups = [(lbl, ["실적", "전년비"]) for lbl, *_ in METRICS]
    rows, crows = [], []
    for ym in sel_months:
        bc, bpv = bundle_cur(ym), bundle_prev(ym)
        row, cr_ = [ym.replace("-", ".")], []
        for lbl, getter, _tm, unit in METRICS:
            cv, pv = getter(bc), getter(bpv)
            suffix = "%" if lbl == "당월인증CR" else ""
            disp = (f"{cv:.1f}%" if lbl == "당월인증CR" else fmt(cv, unit))
            d, dc, _ = pct_ratio(cv, pv)
            row += [disp, d]
            cr_ += [None, dc]
        rows.append(row)
        crows.append(cr_)
    md(render_grouped_table(groups, rows, crows, first_col_label="월"))

    # 세그(신규/윈백/기존) — 기준월
    sub(f"세그별 실적 · {primary}", "인증자수 / 당월인증거래액 / 전체거래액 (실적·비중·전년비·목표비)")
    for metric_label, kind, tgt_metric in [
        ("인증자수", "cert", "인증자수"),
        ("당월인증거래액(백만)", "amt_cert", "당월인증거래액"),
        ("전체거래액(백만)", "amt_all", "거래액"),
    ]:
        unit = 1 if kind == "cert" else 1_000_000
        total = bp["cert"]["전체"] if kind == "cert" else bp[kind]["전체"][pt]
        srows, scr = [], []
        for seg in ["신규", "WIN-BACK", "기존"]:
            if kind == "cert":
                cv = bp["cert"][seg]
                pv = bpp["cert"][seg]
            else:
                cv = bp[kind][seg][pt]
                pv = bpp[kind][seg][pt]
            share = f"{cv / total * 100:.1f}%" if total else "-"
            d, dc, _ = pct_ratio(cv, pv)
            t = M.target_value(target_df, tgt_metric, primary, M.SEG_LABEL[seg])
            td, tc, _ = target_ratio(cv, t)
            srows.append([M.SEG_LABEL[seg], fmt(cv, unit), share, d, td])
            scr.append([None, None, dc, tc])
        md(f'<div style="font-size:13px;font-weight:600;margin-top:10px;">{metric_label}</div>')
        md(render_simple_table(["세그", "실적", "비중", "전년비", "목표비"], srows, scr))

    # LF몰 대비 비중 (월별)
    if not lf_df.empty:
        sub("LF몰 전체 대비 비중 (월별)")
        rows, crows = [], []
        for ym in sel_months:
            bc = bundle_cur(ym)
            lf_a = M.lfmall_value(lf_df, ym, "실제")
            lf_t = M.lfmall_value(lf_df, ym, "목표")
            s_all = f"{bc['amt_all']['전체'][pt] / lf_a * 100:.2f}%" if lf_a else "-"
            s_cert = f"{bc['amt_cert']['전체'][pt] / lf_a * 100:.2f}%" if lf_a else "-"
            rows.append([ym.replace("-", "."), fmt(lf_a, 1_000_000), fmt(lf_t, 1_000_000), s_all, s_cert])
            crows.append([None, None, None, None])
        md(render_simple_table(["월", "LF몰실제(백만)", "LF몰목표(백만)", "전체거래액 비중", "당월인증 비중"], rows, crows))


# ==================================================================================
# 주차별 탭 (목표비 제외)
# ==================================================================================
with tabs[1]:
    export_tab("주차별")
    weeks = M.weeks_of_month(cur_data, primary)
    if not weeks:
        st.info("주차 정보를 찾지 못했습니다(피벗 <주차기준> 확인). 기준월을 바꿔보세요.")
    else:
        sub(f"주차별 실적 · {primary}", f"{paytype} · 전년비는 전년 동월 주차 대비")
        prev_weeks = M.weeks_of_month(prev_data, prev_ym(primary))
        # 전년 주차 매칭: 같은 주차번호(YY_M_W의 W) 기준
        prev_by_wk = {M._parse_week_label(w)[1]: w for w in prev_weeks}
        groups = [(lbl, ["실적", "전년비"]) for lbl, *_ in METRICS]
        rows, crows = [], []
        for w in weeks:
            dts = M.week_dates(cur_data, w)
            bc = M.metric_bundle(cur_data, dts)
            wn = M._parse_week_label(w)[1]
            pw = prev_by_wk.get(wn)
            bpv = M.metric_bundle(prev_data, M.week_dates(prev_data, pw)) if pw else None
            row, cr_ = [w], []
            for lbl, getter, _tm, unit in METRICS:
                cv = getter(bc)
                pv = getter(bpv) if bpv else None
                disp = (f"{cv:.1f}%" if lbl == "당월인증CR" else fmt(cv, unit))
                d, dc, _ = pct_ratio(cv, pv)
                row += [disp, d]
                cr_ += [None, dc]
            rows.append(row)
            crows.append(cr_)
        md(render_grouped_table(groups, rows, crows, first_col_label="주차"))
        st.caption("단위: 당월인증거래액·전체거래액=백만, 객단가=원, 당월인증CR=%")


# ==================================================================================
# 제휴사별 탭 (목표비 제외)
# ==================================================================================
with tabs[2]:
    export_tab("제휴사별")
    # UV·인증·거래액이 발생한 전 제휴사
    def _partners(data, dates):
        s = set()
        amt = data.get("amt")
        if amt is not None and not amt.empty:
            sub_ = A._filter_amt(amt, dates, None, False)
            s |= set(sub_["제휴사"].dropna().unique())
        uv = data.get("uv")
        if uv is not None and not uv.empty:
            s |= set(uv[uv["★일자일"].dt.date.isin(dates)]["제휴사명"].dropna().unique())
        cert = data.get("cert")
        if cert is not None and not cert.empty:
            s |= set(cert[cert["인증일시일"].dt.date.isin(dates)]["제휴사"].dropna().unique())
        return sorted(x for x in s if x and str(x) not in ("nan", ""))

    plist = _partners(cur_data, cur_dates(primary))
    if not plist:
        st.info("해당 월에 활동한 제휴사가 없습니다.")
    else:
        sel_p = st.selectbox("제휴사 선택 (UV·인증·거래액 발생 전 제휴사)", plist)
        sub(f"제휴사 · {sel_p} · 월별 실적", f"{paytype} · 전년비 포함")
        groups = [(lbl, ["실적", "전년비"]) for lbl, *_ in METRICS]
        rows, crows = [], []
        for ym in sel_months:
            dts, pdts = cur_dates(ym), prev_dates(ym)
            bc = M.metric_bundle(cur_data, dts, partners=[sel_p])
            bpv = M.metric_bundle(prev_data, pdts, partners=[sel_p])
            row, cr_ = [ym.replace("-", ".")], []
            for lbl, getter, _tm, unit in METRICS:
                cv, pv = getter(bc), getter(bpv)
                disp = (f"{cv:.1f}%" if lbl == "당월인증CR" else fmt(cv, unit))
                d, dc, _ = pct_ratio(cv, pv)
                row += [disp, d]
                cr_ += [None, dc]
            rows.append(row)
            crows.append(cr_)
        md(render_grouped_table(groups, rows, crows, first_col_label="월"))

        # 제휴사 랭킹 (기준월)
        sub(f"제휴사 랭킹 · {primary}", "당월인증거래액 내림차순")
        pf = A.partner_full_table(cur_data["amt"], cur_data["uv"], cur_data["cert"], cur_dates(primary), plist)
        rrows = []
        for _, r in pf.head(30).iterrows():
            cust = r[f"cust_{pt}"]
            a = r[f"all_{pt}"] / cust if cust else 0
            rrows.append([r["제휴사"], fmt(r["UV"]), fmt(r["인증수"]),
                          fmt(r[f"cert_{pt}"], 1_000_000), fmt(r[f"all_{pt}"], 1_000_000),
                          fmt(cust), fmt(a)])
        md(render_simple_table(["제휴사", "UV", "인증자수", "당월인증거래액(백만)", "전체거래액(백만)", "고객수", "객단가"], rrows))


# ==================================================================================
# BPU별 탭 (목표비 제외 + 기여도/추가분석)
# ==================================================================================
with tabs[3]:
    export_tab("BPU별")
    dts, pdts = cur_dates(primary), prev_dates(primary)
    cert_only = st.radio("거래 범위", ["당월인증거래액", "전체거래액"], index=0, horizontal=True) == "당월인증거래액"

    # BPU별 실적
    sub(f"BPU별 실적 · {primary}", f"{paytype} · {'당월인증' if cert_only else '전체'} 기준")
    bpu_cur = A.group_amt_table(cur_data["amt"], dts, None, "BPU", cert_only=cert_only)
    bpu_prev = (A.group_amt_table(prev_data["amt"], pdts, None, "BPU", cert_only=cert_only)
                if not prev_data["amt"].empty else pd.DataFrame(columns=["BPU", "tot", "net"]))
    prev_map = {r["BPU"]: r for _, r in bpu_prev.iterrows()}
    total_cur = bpu_cur[pt].sum()
    rows, crows = [], []
    for _, r in bpu_cur.sort_values(pt, ascending=False).iterrows():
        p = prev_map.get(r["BPU"])
        pv = p[pt] if p is not None else None
        share = f"{r[pt] / total_cur * 100:.1f}%" if total_cur else "-"
        d, dc, _ = pct_ratio(r[pt], pv)
        rows.append([r["BPU"], fmt(r[pt], 1_000_000), share, d])
        crows.append([None, None, dc])
    md(render_simple_table(["BPU", "거래액(백만)", "비중", "전년비"], rows, crows))

    # 기여도 — 카테고리 / 브랜드
    for gcol, gname in [("물리대카테", "카테고리"), ("Admin브랜드명", "브랜드")]:
        sub(f"상승·하락 기여 {gname} TOP", "기여도 = 그룹 증감 ÷ 전체 증감")
        ctab = M.contribution_table(cur_data["amt"], dts, prev_data["amt"], pdts, gcol, pt=pt, cert_only=cert_only, top=12)
        rows, crows = [], []
        for _, r in ctab.iterrows():
            d, dc, _ = pct_ratio(r["cur"], r.get("prev"))
            contrib = f"{r['contrib'] * 100:+.1f}%" if pd.notna(r["contrib"]) else "-"
            rows.append([r[gcol], fmt(r["cur"], 1_000_000), f"{r['share'] * 100:.1f}%", d, contrib])
            crows.append([None, None, None, dc, ("success" if r["delta"] >= 0 else "danger")])
        md(render_simple_table([gname, "거래액(백만)", "비중", "전년비", "기여도"], rows, crows))

    # 추가분석 1) 객단가 분해
    sub("객단가 분해 (성장 동인)", "거래액 증감 = 고객수 효과 + 객단가 효과")
    dec = M.bpu_decomposition(cur_data["amt"], dts, prev_data["amt"], pdts, pt=pt, cert_only=cert_only)
    rows = []
    for _, r in dec.head(10).iterrows():
        driver = "고객수" if abs(r["cust_effect"]) >= abs(r["aov_effect"]) else "객단가"
        rows.append([r["BPU"], fmt(r["rev_c"], 1_000_000), fmt(r["aov_c"]),
                     fmt(r["cust_effect"], 1_000_000), fmt(r["aov_effect"], 1_000_000), driver])
    md(render_simple_table(["BPU", "거래액(백만)", "객단가", "고객수효과(백만)", "객단가효과(백만)", "주동인"], rows))

    # 추가분석 2) 반품률
    sub("반품률", "(총결제 − 순결제) ÷ 총결제")
    rr = M.return_rate_table(cur_data["amt"], dts, "BPU", cert_only=cert_only)
    rows = [[r["BPU"], fmt(r["tot"], 1_000_000), fmt(r["net"], 1_000_000), f"{r['반품률'] * 100:.1f}%"]
            for _, r in rr.head(10).iterrows()]
    md(render_simple_table(["BPU", "총결제(백만)", "순결제(백만)", "반품률"], rows))

    # 추가분석 3) 신규 침투율
    sub("신규 침투율", "신규 거래액 ÷ 전체 거래액")
    npt = M.new_penetration_table(cur_data["amt"], dts, "BPU", pt=pt, cert_only=cert_only)
    rows = [[r["BPU"], fmt(r["전체"], 1_000_000), fmt(r["신규"], 1_000_000), f"{r['침투율'] * 100:.1f}%"]
            for _, r in npt.head(10).iterrows()]
    md(render_simple_table(["BPU", "전체(백만)", "신규(백만)", "침투율"], rows))

    # 추가분석 4) Pareto
    par = M.pareto_summary(cur_data["amt"], dts, "Admin브랜드명", pt=pt, cert_only=cert_only)
    if par:
        sub("브랜드 집중도 (Pareto)", f"거래 브랜드 {par.get('total_n', 0)}개")
        md(render_simple_table(["구간", "누적 비중"],
                               [["상위 5개", f"{par['top5'] * 100:.1f}%"],
                                ["상위 10개", f"{par['top10'] * 100:.1f}%"],
                                ["상위 20개", f"{par['top20'] * 100:.1f}%"]]))


# ==================================================================================
# 내보내기
# ==================================================================================
st.sidebar.divider()
st.sidebar.header("⤓ Step 3. 내보내기")

_btns, _panels = [], []
for i, sec in enumerate(EXPORT_SECTIONS):
    _btns.append(f'<button class="lf-tabbtn{" active" if i == 0 else ""}" onclick="lfTab({i})">{sec["name"]}</button>')
    _panels.append(f'<div class="lf-panel" id="lf-p-{i}" style="display:{"block" if i == 0 else "none"};">{"".join(sec["html"])}</div>')

export_html = (
    "<!doctype html><html lang='ko'><head><meta charset='utf-8'>"
    f"<title>LF몰 제휴 월마감 - {primary}</title><style>"
    "body{font-family:-apple-system,'Malgun Gothic',sans-serif;max-width:1280px;margin:24px auto;padding:0 16px;color:#1f2328;}"
    ".lf-tabbar{display:flex;gap:4px;border-bottom:2px solid #d0d7de;margin:16px 0 20px;flex-wrap:wrap;}"
    ".lf-tabbtn{border:none;background:none;padding:10px 16px;font-size:14px;font-weight:600;color:#57606a;cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-2px;}"
    ".lf-tabbtn.active{color:#0969da;border-bottom-color:#0969da;}"
    "</style></head><body>"
    "<h1 style='font-size:24px;margin-bottom:4px;'>📆 LF몰 제휴 월마감 분석</h1>"
    f"<p style='color:#57606a;margin-top:0;'>분석월 {', '.join(sel_months)} · 기준월 {primary} · {paytype}</p>"
    f'<div class="lf-tabbar">{"".join(_btns)}</div>' + "".join(_panels)
    + "<script>function lfTab(i){document.querySelectorAll('.lf-panel').forEach((e,x)=>e.style.display=x===i?'block':'none');"
      "document.querySelectorAll('.lf-tabbtn').forEach((e,x)=>e.classList.toggle('active',x===i));}</script>"
    "</body></html>")
st.sidebar.download_button("⤓ HTML 다운로드", export_html, file_name=f"lf_monthly_{primary}.html", mime="text/html")

# JSON 스냅샷
snap = {"분석월": sel_months, "기준월": primary, "결제구분": paytype, "기간기준": close_mode, "지표": {}}
for ym in sel_months:
    bc = bundle_cur(ym)
    snap["지표"][ym] = {
        "UV": bc["uv"], "인증자수": bc["cert"]["전체"],
        "당월인증고객수": bc["amt_cert"]["전체"][f"cust_{pt}"],
        "당월인증CR%": round(cr(bc) * 100, 2),
        "당월인증거래액": bc["amt_cert"]["전체"][pt], "전체거래액": bc["amt_all"]["전체"][pt],
    }
st.sidebar.download_button("⤓ JSON 다운로드", json.dumps(snap, ensure_ascii=False, indent=2),
                           file_name=f"lf_monthly_{primary}.json", mime="application/json")
