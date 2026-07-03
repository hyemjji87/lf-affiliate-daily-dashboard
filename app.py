"""
LF몰 일반제휴 실적 대시보드 (Streamlit)

실행:  streamlit run app.py
필요 패키지: requirements.txt 참고

설계 요약 (대화에서 확정된 사양)
- 업로드 3종: ① 당해년도 실적  ② 전년도 인증 실적  ③ 전년도 거래 실적(분기, 복수 업로드 가능)
- 분석 기준일: (1) 특정 일자 복수선택  (2) 기준일 MTD(월초~기준일)
- 전년 동일자 매칭: 캘린더 동일 날짜 / 요일기준(52주 전) / 전년 특정 일자 복수선택 중 선택
- 모든 지표(UV·인증자수·거래액·고객수·객단가)의 모집단은 "선택 기간에 당월인증거래액이 발생한 제휴사"로 통일
- SSNGCD03(배너성 AF코드) UV 제외 옵션
- 총결제/순결제는 한 표에 섞지 않고 서브탭으로 분리하며, 표시 순서는 순결제 -> 총결제
- 전년비 표기: 상승 "▼", 하락 "△"(빨강) - 특정 회사 내부 컨벤션이라 이 순서를 그대로 사용
- 화면에는 전년비(%)만 표기하고, 전년도 원본 값은 엑셀 다운로드에만 포함
- 신규/윈백/기존 구분은 카테고리·브랜드 탭을 제외한 모든 탭에 적용
"""
import os
import datetime as dt

import numpy as np
import pandas as pd
import streamlit as st

import lf_analysis as A

st.set_page_config(page_title="LF몰 일반제휴 실적", page_icon="📊", layout="wide")

# 일자별 목표(순결제 기준) 파일 - 연말까지 고정값이라 업로드 대신 리포에 고정 파일로 관리한다.
# 목표가 갱신되면 이 파일(target_data.xlsx)만 교체하면 된다 (별도 요청 시 반영).
TARGET_FILE = os.path.join(os.path.dirname(__file__), "target_data.xlsx")

# ----------------------------------------------------------------------------------
# 캐싱된 로더 (같은 파일이면 재계산하지 않음)
# ----------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_single(file_bytes: bytes) -> dict:
    return A.classify_workbook(file_bytes)


@st.cache_data(show_spinner=False)
def _load_multi(list_of_bytes: tuple) -> dict:
    return A.merge_workbooks(list(list_of_bytes))


@st.cache_data(show_spinner=False)
def _load_target_from_disk(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame()
    with open(path, "rb") as f:
        return A.load_target_file(f.read())


# ----------------------------------------------------------------------------------
# 포맷 유틸
# ----------------------------------------------------------------------------------
def fmt(n):
    if n is None or (isinstance(n, float) and pd.isna(n)):
        return "-"
    return f"{round(n):,}"


def sub(title: str):
    """st.subheader는 표 폰트(13px) 대비 지나치게 크므로, 표보다는 크되 절제된 크기의
    섹션 제목으로 대체."""
    st.markdown(f"<div style='font-size:16px;font-weight:700;margin:18px 0 6px;'>{title}</div>",
                unsafe_allow_html=True)


def pct_ratio(cur, prev):
    """엑셀/화면 공통으로 쓸 (표시문자열, 색상, 원시비율) 반환."""
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
    """목표비: 실적/목표 달성률(%) 숫자만 보여준다 - 전년비와 달리 ▼/△ 기호는 아예 쓰지 않는다
    (목표비는 절대 음수가 될 수 없어서 방향 기호가 의미가 없음). 100% 이상(달성·초과)이면 초록,
    100% 미만(미달)이면 빨강으로 색만 구분한다. 목표가 아예 없으면 '-'로 표기한다."""
    if (target is None or (isinstance(target, float) and pd.isna(target)) or target == 0
            or cur is None or (isinstance(cur, float) and pd.isna(cur))):
        return "-", "muted", None
    ratio = cur / target
    pct = round(ratio * 100)
    color = "success" if pct >= 100 else "danger"
    return f"{pct}%", color, ratio


COLOR_MAP = {"success": "#1a7f37", "danger": "#cf222e", "muted": "#6e7781", "accent": "#0969da"}


def render_grouped_table(header_groups: list, rows: list, color_rows: list | None = None, first_col_label: str = "일자"):
    """일자별 탭처럼 '상위 카테고리(전체/신규/윈백/기존) + 하위 지표(실적/전년비/목표비)'를
    2단 헤더로 묶어서 보여주는 HTML 표. 첫 컬럼(일자)은 1·2행을 세로 병합하고, 모든 셀은
    가운데 정렬한다.
    - header_groups: [(그룹명, [하위컬럼명, ...]), ...]
    - rows: [[일자표시값, 값1, 값2, ...], ...] (값은 이미 fmt()된 문자열)
    - color_rows: rows와 동일한 shape(첫 컬럼 제외)으로 각 셀의 색상키('success'/'danger'/'accent'/None)를
      담은 리스트. None이면 색상 없이 렌더링.
    """
    ths_top = ['<th rowspan="2" style="text-align:center;vertical-align:middle;'
               'border:1px solid #e1e4e8;padding:6px 10px;background:#f6f8fa;">'
               f'{first_col_label}</th>']
    for g, subs in header_groups:
        ths_top.append(f'<th colspan="{len(subs)}" style="text-align:center;'
                        f'border:1px solid #e1e4e8;padding:6px 10px;background:#f6f8fa;">{g}</th>')
    ths_sub = []
    for g, subs in header_groups:
        for s in subs:
            ths_sub.append(f'<th style="text-align:center;border:1px solid #e1e4e8;'
                            f'padding:4px 10px;background:#fafbfc;font-weight:500;">{s}</th>')

    body_rows = []
    for ri, row in enumerate(rows):
        cells = [f'<td style="text-align:center;border:1px solid #e1e4e8;'
                 f'padding:4px 10px;white-space:nowrap;">{row[0]}</td>']
        for ci, v in enumerate(row[1:]):
            style = "text-align:center;border:1px solid #e1e4e8;padding:4px 10px;white-space:nowrap;"
            if color_rows is not None:
                key = color_rows[ri][ci]
                if key in ("success", "danger", "accent"):
                    style += f"color:{COLOR_MAP[key]};font-weight:600;"
            cells.append(f'<td style="{style}">{v}</td>')
        body_rows.append(f'<tr>{"".join(cells)}</tr>')

    html = (
        '<div style="overflow-x:auto;">'
        '<table style="border-collapse:collapse;width:100%;font-size:13px;">'
        f'<thead><tr>{"".join(ths_top)}</tr><tr>{"".join(ths_sub)}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        '</table></div>'
    )
    return html


def render_simple_table(columns: list, rows: list, color_rows: list | None = None):
    """반복되는 지표 그룹이 없는 단순 표에도 일자별 탭과 동일한 보더/가운데정렬/폰트 스타일을
    적용하기 위한 단일 헤더행 HTML 표.
    - columns: 전체 컬럼명 목록(첫 컬럼 포함)
    - rows: [[첫컬럼값, 값1, 값2, ...], ...]
    - color_rows: rows와 동일 행 수, 첫 컬럼 제외한 각 셀의 색상키 리스트
    """
    ths = [f'<th style="text-align:center;border:1px solid #e1e4e8;padding:6px 10px;'
           f'background:#f6f8fa;">{c}</th>' for c in columns]
    body_rows = []
    for ri, row in enumerate(rows):
        cells = [f'<td style="text-align:center;border:1px solid #e1e4e8;'
                 f'padding:4px 10px;white-space:nowrap;">{row[0]}</td>']
        for ci, v in enumerate(row[1:]):
            style = "text-align:center;border:1px solid #e1e4e8;padding:4px 10px;white-space:nowrap;"
            if color_rows is not None:
                key = color_rows[ri][ci]
                if key in ("success", "danger", "accent"):
                    style += f"color:{COLOR_MAP[key]};font-weight:600;"
            cells.append(f'<td style="{style}">{v}</td>')
        body_rows.append(f'<tr>{"".join(cells)}</tr>')
    html = (
        '<div style="overflow-x:auto;">'
        '<table style="border-collapse:collapse;width:100%;font-size:13px;">'
        f'<thead><tr>{"".join(ths)}</tr></thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        '</table></div>'
    )
    return html


# ----------------------------------------------------------------------------------
# 사이드바: 파일 업로드 & 분석 설정
# ----------------------------------------------------------------------------------
st.sidebar.header("📁 Step 1. 파일 업로드")
f_cur = st.sidebar.file_uploader("① 당해년도 실적", type=["xlsx"], key="f_cur")
f_ly_cert = st.sidebar.file_uploader("② 전년도 인증 실적", type=["xlsx"], key="f_ly_cert")
f_ly_amt = st.sidebar.file_uploader(
    "③ 전년도 거래 실적 (분기별, 여러 파일 업로드 가능)",
    type=["xlsx"], accept_multiple_files=True, key="f_ly_amt",
)

st.sidebar.divider()
st.sidebar.header("⚙ Step 2. 분석 설정")

if f_cur is None:
    st.sidebar.info("① 당해년도 실적 파일을 올리면 분석 설정이 열립니다.")
    st.title("📊 LF몰 일반제휴 실적")
    st.info("왼쪽 사이드바에서 ① 당해년도 실적 파일을 먼저 업로드해 주세요.")
    st.stop()

cur_data = _load_single(f_cur.getvalue())
if cur_data["amt"].empty:
    st.error("업로드한 파일에서 '인증거래액' 형태의 시트를 찾지 못했습니다. 파일을 확인해 주세요.")
    st.stop()

available_dates = sorted(cur_data["amt"]["정산일시일"].dropna().dt.date.unique().tolist())
min_d, max_d = available_dates[0], available_dates[-1]

mode = st.sidebar.radio("분석 기준일", ["특정 일자 복수선택", "기준일 MTD"], index=1)

if mode == "특정 일자 복수선택":
    sel_dates = st.sidebar.multiselect(
        "날짜 선택 (드롭다운, 여러 개 선택 가능)",
        options=available_dates, default=available_dates,
        format_func=lambda d: d.strftime("%Y-%m-%d"),
    )
else:
    base_date = st.sidebar.date_input(
        "MTD 마감일", value=max_d, min_value=min_d, max_value=max_d,
    )
    month_start = base_date.replace(day=1)
    sel_dates = [d for d in available_dates if month_start <= d <= base_date]

if not sel_dates:
    st.warning("분석할 날짜를 하나 이상 선택해 주세요.")
    st.stop()

excl_ssng = st.sidebar.checkbox("SSNGCD03(삼성배너) UV 제외", value=False)
yoy_mode_label = st.sidebar.radio(
    "전년 동일자 매칭",
    ["캘린더 동일 날짜", "요일기준(52주 전)", "전년 특정 일자 복수선택"],
    index=0,
)
ly_manual_dates = None
if yoy_mode_label == "전년 특정 일자 복수선택":
    yoy_mode = "manual"
    ly_candidates = A.ly_month_candidates(sel_dates)
    _default_ly = [d for d in A.ly_dates(sel_dates, "calendar") if d in ly_candidates]
    ly_manual_dates = st.sidebar.multiselect(
        "전년도 비교 일자 선택 (해당 월 전체 중 복수선택 가능)",
        options=ly_candidates, default=_default_ly,
        format_func=lambda d: d.strftime("%Y-%m-%d"),
    )
    if not ly_manual_dates:
        st.sidebar.caption("⚠ 전년도 비교 일자를 선택하지 않으면 전년비가 계산되지 않습니다.")
else:
    yoy_mode = "calendar" if yoy_mode_label == "캘린더 동일 날짜" else "weekday"
exclude_af = "SSNGCD03" if excl_ssng else None

# ----------------------------------------------------------------------------------
# 전년도 데이터 로딩 (선택사항 - 없으면 전년비는 "-"로 표기)
# ----------------------------------------------------------------------------------
ly_cert_data = _load_single(f_ly_cert.getvalue()) if f_ly_cert else {"amt": pd.DataFrame(), "uv": pd.DataFrame(), "cert": pd.DataFrame()}
ly_amt_bytes = tuple(f.getvalue() for f in f_ly_amt) if f_ly_amt else tuple()
ly_amt_data = _load_multi(ly_amt_bytes) if ly_amt_bytes else {"amt": pd.DataFrame(), "uv": pd.DataFrame(), "cert": pd.DataFrame()}
target_df = _load_target_from_disk(TARGET_FILE)

ly_all_dates = ly_manual_dates if yoy_mode == "manual" else A.ly_dates(sel_dates, yoy_mode)

# [주의] qual_partners(당월인증거래액>0인 제휴사)는 "제휴사별 실적" 탭에 어떤 제휴사를
# 리스팅할지 정하는 기준일 뿐이다. UV/인증자수/거래액 등 다른 모든 탭의 집계는 이 리스트로
# 제한하지 않고 업로드된 전체 제휴사(partners=None) 기준으로 계산한다.
qual_partners = A.qualifying_partners(cur_data["amt"], sel_dates)

# ----------------------------------------------------------------------------------
# 헤더
# ----------------------------------------------------------------------------------
warn_bits = []
if f_ly_cert is None:
    warn_bits.append("전년도 인증 실적 미업로드(UV·인증자수 전년비 불가)")
if not f_ly_amt:
    warn_bits.append("전년도 거래 실적 미업로드(거래액 전년비 불가)")

st.title("📊 LF몰 일반제휴 실적")
period_txt = f"{sel_dates[0]:%Y-%m-%d} ~ {sel_dates[-1]:%Y-%m-%d}" if len(sel_dates) > 1 else f"{sel_dates[0]:%Y-%m-%d}"
st.caption(
    f"분석기간 {period_txt} · UV·인증자수·거래액 등은 전체 제휴사 기준 · "
    f"제휴사별 실적 탭 리스팅 대상 {len(qual_partners)}개(당월인증거래액 발생 기준) · "
    f"전년 매칭: {yoy_mode_label}" + ("" if not warn_bits else " · ⚠ " + " / ".join(warn_bits))
)

json_col, _ = st.columns([1, 5])
with json_col:
    pass  # JSON 내보내기 버튼은 각 탭 하단에서 개별 제공 (탭별 데이터가 다르므로)

tab_ov, tab_daily, tab_partner, tab_cat, tab_brand = st.tabs(
    ["개요", "일자별", "제휴사별 실적", "카테고리별 실적", "브랜드별 실적"]
)

# ====================================================================================
# 개요 탭
# ====================================================================================
with tab_ov:
    uv_cur = A.uv_total(cur_data["uv"], sel_dates, None, exclude_af=exclude_af)
    uv_prev = A.uv_total(ly_cert_data["uv"], ly_all_dates, None) if not ly_cert_data["uv"].empty else None
    cert_cur = A.cert_summary(cur_data["cert"], sel_dates, None)
    cert_prev = A.cert_summary(ly_cert_data["cert"], ly_all_dates, None) if not ly_cert_data["cert"].empty else {"전체": None, "기존": None, "WIN-BACK": None, "신규": None}

    c1, c2 = st.columns(2)
    with c1:
        st.metric("UV", fmt(uv_cur), pct_ratio(uv_cur, uv_prev)[0])
    with c2:
        st.metric("인증자수", fmt(cert_cur["전체"]), pct_ratio(cert_cur["전체"], cert_prev["전체"])[0])

    sub("인증자수 구성 (비중은 기준일 기준)")
    rows, color_rows = [], []
    for k, label in [("신규", "신규"), ("WIN-BACK", "윈백"), ("기존", "기존")]:
        share = f"{cert_cur[k] / cert_cur['전체'] * 100:.1f}%" if cert_cur["전체"] else "-"
        d, d_color, _ = pct_ratio(cert_cur[k], cert_prev.get(k))
        rows.append([label, fmt(cert_cur[k]), share, d])
        color_rows.append([None, None, d_color])
    st.markdown(render_simple_table(["구분", "기준일", "비중", "전년비"], rows, color_rows), unsafe_allow_html=True)

    paytype = st.radio("결제 구분", ["순결제", "총결제"], horizontal=True, key="ov_paytype")
    pt = "net" if paytype == "순결제" else "tot"

    def _amt_block(label, title):
        cur = A.amt_summary(cur_data["amt"], sel_dates, None, cert_only=(label == "당월인증"))
        prev = (A.amt_summary(ly_amt_data["amt"], ly_all_dates, None, cert_only=(label == "당월인증"))
                if not ly_amt_data["amt"].empty else {s: {"tot": None, "net": None, "cust_tot": None, "cust_net": None} for s in ["전체"] + A.STATUSES})
        cur_total = cur["전체"][pt]
        prev_total = prev["전체"][pt]
        d_all, d_all_color, _ = pct_ratio(cur_total, prev_total)
        show_target = pt == "net" and not target_df.empty
        group = "cert" if label == "당월인증" else "all"
        d_all_css = f"color:{COLOR_MAP.get(d_all_color, '#000')};font-weight:600" if d_all_color in COLOR_MAP else ""
        header = (f"**{title}**  \n### {fmt(cur_total)}  "
                  f"<span style='font-size:13px;{d_all_css}'>전년비 {d_all}</span>")
        if show_target:
            tgt_total = A.target_sum(target_df, sel_dates, group, "total")
            t_all, t_all_color, _ = target_ratio(cur_total, tgt_total)
            t_all_css = f"color:{COLOR_MAP.get(t_all_color, '#000')};font-weight:600" if t_all_color in COLOR_MAP else ""
            header += f"  <span style='font-size:13px;{t_all_css}'>목표비 {t_all}</span>"
        st.markdown(header, unsafe_allow_html=True)
        rows, color_rows = [], []
        status_target_key = {"신규": "신규", "WIN-BACK": "윈백", "기존": "기존"}
        for k, lbl in [("신규", "신규"), ("WIN-BACK", "윈백"), ("기존", "기존")]:
            c = cur[k][pt]
            p = prev[k][pt]
            share = f"{c / cur_total * 100:.1f}%" if cur_total else "-"
            d, d_color, _ = pct_ratio(c, p)
            row = [lbl, fmt(c), share, d]
            color_row = [None, None, d_color]
            if show_target:
                tgt_v = A.target_sum(target_df, sel_dates, group, status_target_key[k])
                t, t_color, _ = target_ratio(c, tgt_v)
                row.append(t)
                color_row.append(t_color)
            rows.append(row)
            color_rows.append(color_row)
        cols = ["구분", "기준일", "비중", "전년비"] + (["목표비"] if show_target else [])
        st.markdown(render_simple_table(cols, rows, color_rows), unsafe_allow_html=True)
        return cur, prev

    cur_all, prev_all = _amt_block("전체", "전체거래액")
    cur_cert, prev_cert = _amt_block("당월인증", "당월인증거래액")

    def _cust_aov_table(cur, prev, title):
        sub(title)
        cust_c, cust_p = cur["전체"][f"cust_{pt}"], prev["전체"][f"cust_{pt}"]
        aov_c = cur["전체"][pt] / cust_c if cust_c else 0
        aov_p = (prev["전체"][pt] / cust_p) if cust_p else None
        d1, c1, _ = pct_ratio(cust_c, cust_p)
        d2, c2, _ = pct_ratio(aov_c, aov_p)
        rows = [["고객수", fmt(cust_c), d1], ["객단가", fmt(aov_c), d2]]
        color_rows = [[None, c1], [None, c2]]
        st.markdown(render_simple_table(["구분", "기준일", "전년비"], rows, color_rows), unsafe_allow_html=True)

    _cust_aov_table(cur_all, prev_all, "전체 고객수 · 객단가")
    _cust_aov_table(cur_cert, prev_cert, "당월인증 고객수 · 객단가")

# ====================================================================================
# 일자별 탭 (화면 = 값 + 전년비만, 전년 원본값은 엑셀에만)
# ====================================================================================
with tab_daily:
    daily_uc = A.daily_uv_cert_table(cur_data["uv"], cur_data["cert"], sel_dates, None, exclude_af=exclude_af)
    daily_uc_ly = (A.daily_uv_cert_table(ly_cert_data["uv"], ly_cert_data["cert"], ly_all_dates, None)
                   if (not ly_cert_data["uv"].empty or not ly_cert_data["cert"].empty) else None)

    sub("UV")
    uv_rows, uv_color_rows = [], []
    for i, d in enumerate(sorted(sel_dates)):
        prev_v = daily_uc_ly.iloc[i]["UV"] if daily_uc_ly is not None and i < len(daily_uc_ly) else None
        disp, color, _ = pct_ratio(daily_uc.iloc[i]["UV"], prev_v)
        uv_rows.append([d.strftime("%Y-%m-%d"), fmt(daily_uc.iloc[i]["UV"]), disp])
        uv_color_rows.append([None, color])
    st.markdown(render_simple_table(["일자", "UV", "전년비"], uv_rows, uv_color_rows), unsafe_allow_html=True)

    sub("인증자수 (전체·신규·윈백·기존)")
    cert_rows, cert_color_rows = [], []
    for i, d in enumerate(sorted(sel_dates)):
        r = daily_uc.iloc[i]
        pr = daily_uc_ly.iloc[i] if daily_uc_ly is not None and i < len(daily_uc_ly) else None
        row = [d.strftime("%Y-%m-%d")]
        color_row = []
        for col in ["인증_전체", "인증_신규", "인증_윈백", "인증_기존"]:
            pv = pr[col] if pr is not None else None
            d_disp, d_color, _ = pct_ratio(r[col], pv)
            row += [fmt(r[col]), d_disp]
            color_row += [None, d_color]
        cert_rows.append(row)
        cert_color_rows.append(color_row)
    cert_groups = [(s, ["인증수", "전년비"]) for s in ["전체", "신규", "윈백", "기존"]]
    st.markdown(render_grouped_table(cert_groups, cert_rows, cert_color_rows), unsafe_allow_html=True)

    paytype_d = st.radio("결제 구분", ["순결제", "총결제"], horizontal=True, key="daily_paytype")
    ptd = "net" if paytype_d == "순결제" else "tot"

    def _daily_amt_block(cert_only, title):
        cur_d = A.daily_amt_table(cur_data["amt"], sel_dates, None, cert_only=cert_only)
        ly_d = (A.daily_amt_table(ly_amt_data["amt"], ly_all_dates, None, cert_only=cert_only)
                if not ly_amt_data["amt"].empty else None)
        sub(f"{title} (전체·신규·윈백·기존)")
        show_target = ptd == "net" and not target_df.empty
        group = "cert" if cert_only else "all"
        status_target_key = {"전체": "total", "신규": "신규", "WIN-BACK": "윈백", "기존": "기존"}
        rows, color_rows = [], []
        for i in range(len(cur_d)):
            r = cur_d.iloc[i]
            pr = ly_d.iloc[i] if ly_d is not None and i < len(ly_d) else None
            row = [r["일자"].strftime("%Y-%m-%d")]
            color_row = []
            for s in ["전체"] + A.STATUSES:
                c = r[f"{s}_{ptd}"]
                p = pr[f"{s}_{ptd}"] if pr is not None else None
                d_disp, d_color, _ = pct_ratio(c, p)
                row += [fmt(c), d_disp]
                color_row += [None, d_color]
                if show_target:
                    tgt_v = A.target_lookup(target_df, r["일자"], group, status_target_key[s])
                    t_disp, t_color, _ = target_ratio(c, tgt_v)
                    row.append(t_disp)
                    color_row.append(t_color)
            rows.append(row)
            color_rows.append(color_row)
        subcols = ["실적", "전년비"] + (["목표비"] if show_target else [])
        groups = [(s, subcols) for s in ["전체", "신규", "윈백", "기존"]]
        st.markdown(render_grouped_table(groups, rows, color_rows), unsafe_allow_html=True)
        return cur_d, ly_d

    all_cur_d, all_ly_d = _daily_amt_block(False, "전체거래액")
    cert_cur_d, cert_ly_d = _daily_amt_block(True, "당월인증거래액")

    def _daily_cust_aov(cur_d, ly_d, title):
        sub(title)
        rows, color_rows = [], []
        for i in range(len(cur_d)):
            r = cur_d.iloc[i]
            pr = ly_d.iloc[i] if ly_d is not None and i < len(ly_d) else None
            cust_c = r[f"전체_cust_{ptd}"]
            aov_c = r[f"전체_{ptd}"] / cust_c if cust_c else 0
            cust_p = pr[f"전체_cust_{ptd}"] if pr is not None else None
            aov_p = (pr[f"전체_{ptd}"] / cust_p) if (pr is not None and cust_p) else None
            d1, c1, _ = pct_ratio(cust_c, cust_p)
            d2, c2, _ = pct_ratio(aov_c, aov_p)
            rows.append([r["일자"].strftime("%Y-%m-%d"), fmt(cust_c), d1, fmt(aov_c), d2])
            color_rows.append([None, c1, None, c2])
        groups = [("고객수", ["실적", "전년비"]), ("객단가", ["실적", "전년비"])]
        st.markdown(render_grouped_table(groups, rows, color_rows, first_col_label="일자"), unsafe_allow_html=True)

    _daily_cust_aov(cert_cur_d, cert_ly_d, "고객수 · 객단가 (당월인증 기준)")
    _daily_cust_aov(all_cur_d, all_ly_d, "고객수 · 객단가 (전체거래액 기준)")

# ====================================================================================
# 제휴사별 실적 탭
# ====================================================================================
with tab_partner:
    paytype_p = st.radio("결제 구분", ["순결제", "총결제"], horizontal=True, key="partner_paytype")
    ptp = "net" if paytype_p == "순결제" else "tot"

    pf_cur = A.partner_full_table(cur_data["amt"], cur_data["uv"], cur_data["cert"], sel_dates, qual_partners, exclude_af=exclude_af)
    pf_ly = (A.partner_full_table(ly_amt_data["amt"], ly_cert_data["uv"], ly_cert_data["cert"], ly_all_dates, None)
             if not ly_amt_data["amt"].empty else None)
    ly_map = {r["제휴사"]: r for _, r in pf_ly.iterrows()} if pf_ly is not None else {}

    sub(f"전체 제휴사 (당월인증거래액 발생 기준, 내림차순 · {len(pf_cur)}개)")
    rows, color_rows = [], []
    for _, r in pf_cur.iterrows():
        p = ly_map.get(r["제휴사"])
        cert_c, all_c = r[f"cert_{ptp}"], r[f"all_{ptp}"]
        cust_c = r[f"cust_{ptp}"]
        aov_c = all_c / cust_c if cust_c else 0
        if p is not None:
            cert_p, all_p, cust_p = p[f"cert_{ptp}"], p[f"all_{ptp}"], p[f"cust_{ptp}"]
            aov_p = all_p / cust_p if cust_p else None
            uv_p, cert_cnt_p = p["UV"], p["인증수"]
        else:
            cert_p = all_p = cust_p = aov_p = uv_p = cert_cnt_p = None
        d_uv, c_uv, _ = pct_ratio(r["UV"], uv_p)
        d_cert, c_cert, _ = pct_ratio(r["인증수"], cert_cnt_p)
        d_certamt, c_certamt, _ = pct_ratio(cert_c, cert_p)
        d_all, c_all, _ = pct_ratio(all_c, all_p)
        d_cust, c_cust, _ = pct_ratio(cust_c, cust_p)
        d_aov, c_aov, _ = pct_ratio(aov_c, aov_p)
        rows.append([
            r["제휴사"], fmt(r["UV"]), d_uv, fmt(r["인증수"]), d_cert,
            fmt(cert_c), d_certamt, fmt(all_c), d_all, fmt(cust_c), d_cust, fmt(aov_c), d_aov,
        ])
        color_rows.append([None, c_uv, None, c_cert, None, c_certamt, None, c_all, None, c_cust, None, c_aov])
    partner_groups = [(g, ["실적", "전년비"]) for g in ["UV", "인증수", "당월인증거래액", "전체거래액", "고객수", "객단가"]]
    st.markdown(render_grouped_table(partner_groups, rows, color_rows, first_col_label="제휴사"), unsafe_allow_html=True)
    st.caption("전년비는 전년도 동일 제휴사 코호트 기준(전년도 데이터가 없으면 \"신규\" 표기).")

    if not qual_partners:
        st.info("선택한 기간에 당월인증거래액이 발생한 제휴사가 없어 리스팅할 제휴사가 없습니다.")
    else:
        sub("제휴사 선택 → 일자별 상세")
        sel_partner = st.selectbox("제휴사", qual_partners)
        pd_cur = A.partner_daily_table(cur_data["amt"], cur_data["uv"], cur_data["cert"], sel_dates, sel_partner, exclude_af=exclude_af)
        show_cols = ["일자", "UV", "인증_전체",
                     f"당월인증_{ptp}", f"전체거래액_{ptp}", f"고객수_{ptp}"]
        disp = pd_cur[show_cols].rename(columns={
            "인증_전체": "인증수", f"당월인증_{ptp}": "당월인증거래액",
            f"전체거래액_{ptp}": "전체거래액", f"고객수_{ptp}": "고객수"})
        disp["객단가"] = (disp["전체거래액"] / disp["고객수"].replace(0, np.nan)).fillna(0)
        detail_rows = []
        for _, r in disp.iterrows():
            detail_rows.append([
                r["일자"].strftime("%Y-%m-%d"), fmt(r["UV"]), fmt(r["인증수"]),
                fmt(r["당월인증거래액"]), fmt(r["전체거래액"]), fmt(r["고객수"]), fmt(r["객단가"]),
            ])
        st.markdown(render_simple_table(
            ["일자", "UV", "인증수", "당월인증거래액", "전체거래액", "고객수", "객단가"], detail_rows,
        ), unsafe_allow_html=True)

# ====================================================================================
# 카테고리별 실적 탭
# ====================================================================================
with tab_cat:
    paytype_c = st.radio("결제 구분", ["순결제", "총결제"], horizontal=True, key="cat_paytype")
    ptc = "net" if paytype_c == "순결제" else "tot"

    cat_cert = A.group_amt_table(cur_data["amt"], sel_dates, None, "물리대카테", cert_only=True).rename(columns={"tot": "cert_tot", "net": "cert_net"})
    cat_all = A.group_amt_table(cur_data["amt"], sel_dates, None, "물리대카테", cert_only=False).rename(columns={"tot": "all_tot", "net": "all_net"})
    cat = pd.merge(cat_cert, cat_all, on="물리대카테", how="outer").fillna(0)
    cat = cat.sort_values("cert_tot", ascending=False).reset_index(drop=True)

    cat_ly = None
    if not ly_amt_data["amt"].empty:
        cat_cert_ly = A.group_amt_table(ly_amt_data["amt"], ly_all_dates, None, "물리대카테", cert_only=True).rename(columns={"tot": "cert_tot", "net": "cert_net"})
        cat_all_ly = A.group_amt_table(ly_amt_data["amt"], ly_all_dates, None, "물리대카테", cert_only=False).rename(columns={"tot": "all_tot", "net": "all_net"})
        cat_ly = pd.merge(cat_cert_ly, cat_all_ly, on="물리대카테", how="outer").fillna(0).set_index("물리대카테")

    sub(f"전체 카테고리 (당월인증거래액 내림차순 · {len(cat)}개)")
    rows, color_rows = [], []
    for _, r in cat.iterrows():
        p = cat_ly.loc[r["물리대카테"]] if (cat_ly is not None and r["물리대카테"] in cat_ly.index) else None
        cert_c, all_c = r[f"cert_{ptc}"], r[f"all_{ptc}"]
        cert_p = p[f"cert_{ptc}"] if p is not None else None
        all_p = p[f"all_{ptc}"] if p is not None else None
        d_cert, c_cert, _ = pct_ratio(cert_c, cert_p)
        d_all, c_all, _ = pct_ratio(all_c, all_p)
        rows.append([r["물리대카테"], fmt(cert_c), d_cert, fmt(all_c), d_all])
        color_rows.append([None, c_cert, None, c_all])
    cat_groups = [("당월인증거래액", ["실적", "전년비"]), ("전체거래액", ["실적", "전년비"])]
    st.markdown(render_grouped_table(cat_groups, rows, color_rows, first_col_label="카테고리"), unsafe_allow_html=True)

# ====================================================================================
# 브랜드별 실적 탭
# ====================================================================================
with tab_brand:
    paytype_b = st.radio("결제 구분", ["순결제", "총결제"], horizontal=True, key="brand_paytype")
    ptb = "net" if paytype_b == "순결제" else "tot"

    br_cert = A.group_amt_table(cur_data["amt"], sel_dates, None, "Admin브랜드명", cert_only=True).rename(columns={"tot": "cert_tot", "net": "cert_net"})
    br_all = A.group_amt_table(cur_data["amt"], sel_dates, None, "Admin브랜드명", cert_only=False).rename(columns={"tot": "all_tot", "net": "all_net"})
    br = pd.merge(br_cert, br_all, on="Admin브랜드명", how="outer").fillna(0)
    br = br.sort_values("cert_tot", ascending=False).head(25).reset_index(drop=True)

    br_ly = None
    if not ly_amt_data["amt"].empty:
        br_cert_ly = A.group_amt_table(ly_amt_data["amt"], ly_all_dates, None, "Admin브랜드명", cert_only=True).rename(columns={"tot": "cert_tot", "net": "cert_net"})
        br_all_ly = A.group_amt_table(ly_amt_data["amt"], ly_all_dates, None, "Admin브랜드명", cert_only=False).rename(columns={"tot": "all_tot", "net": "all_net"})
        br_ly = pd.merge(br_cert_ly, br_all_ly, on="Admin브랜드명", how="outer").fillna(0).set_index("Admin브랜드명")

    sub("브랜드 TOP 25 (당월인증거래액 내림차순)")
    rows, color_rows = [], []
    for _, r in br.iterrows():
        p = br_ly.loc[r["Admin브랜드명"]] if (br_ly is not None and r["Admin브랜드명"] in br_ly.index) else None
        cert_c, all_c = r[f"cert_{ptb}"], r[f"all_{ptb}"]
        cert_p = p[f"cert_{ptb}"] if p is not None else None
        all_p = p[f"all_{ptb}"] if p is not None else None
        d_cert, c_cert, _ = pct_ratio(cert_c, cert_p)
        d_all, c_all, _ = pct_ratio(all_c, all_p)
        rows.append([r["Admin브랜드명"], fmt(cert_c), d_cert, fmt(all_c), d_all])
        color_rows.append([None, c_cert, None, c_all])
    br_groups = [("당월인증거래액", ["실적", "전년비"]), ("전체거래액", ["실적", "전년비"])]
    st.markdown(render_grouped_table(br_groups, rows, color_rows, first_col_label="브랜드"), unsafe_allow_html=True)

# ----------------------------------------------------------------------------------
# 전체 JSON 내보내기 (사이드바 하단)
# ----------------------------------------------------------------------------------
st.sidebar.divider()
if st.sidebar.button("⤓ 현재 화면 JSON 스냅샷 만들기"):
    snapshot = {
        "기간": [str(d) for d in sel_dates],
        "UV(전체 제휴사 기준)": uv_cur,
        "인증자수(전체 제휴사 기준)": cert_cur,
        "제휴사별_실적_탭_리스팅_대상(당월인증거래액>0)": qual_partners,
    }
    import json
    st.sidebar.download_button("JSON 다운로드", json.dumps(snapshot, ensure_ascii=False, indent=2),
                                file_name="lf_affiliate_snapshot.json", mime="application/json")
