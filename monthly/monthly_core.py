"""
LF몰 제휴 월마감 대시보드 - 월별 집계 & 피벗 기반 파생 로직

핵심: 업로드 raw의 `당월인증`·`제휴사`·`거래액_VAT제외`·UV `제휴사명`은 원본이 '피벗 시트를
참조하는 엑셀 수식'이라, 파일에 계산 캐시가 없으면 값이 비어 보인다. 그래서 이 모듈은
엑셀 수식을 파이썬으로 그대로 재현(피벗에서 매핑)하여, 캐시 유무와 무관하게 항상 정확한
값을 만든다.

재현하는 수식 (인증거래액 시트 기준)
- 당월인증(AC)  = IF(LEFT(정산일,7)=LEFT(최초인증일,7),"Y","N")   → 정산 YYYY-MM == 최초인증 YYYY-MM
- 제휴사(AB)     = VLOOKUP(제휴처구분3, 피벗!B:C, 2)                 → 코드→약칭
- 거래액_VAT제외 = 거래액 / 1.1
- 인증회원 제휴사 = VLOOKUP(제휴처구분4, 피벗!B:C, 2)
- 유입 제휴사명   = AF코드 → 피벗 I:J(930명) → 피벗 B:C(약칭)
- 주차           = 피벗 <주차기준> 일자 → 라벨(YY_M_W)  →  마감(년,월) & 주차번호
"""
from __future__ import annotations
import re
import datetime as dt

import numpy as np
import pandas as pd

import lf_analysis as A

STATUSES = A.STATUSES  # ["신규", "WIN-BACK", "기존"]
SEG_LABEL = {"신규": "신규", "WIN-BACK": "윈백", "기존": "기존"}


# ======================================================================================
# 1) 피벗 파싱: 3개 블록(제휴사 구분 / 주차기준 / 제휴사명 정제)을 각각 dict로
# ======================================================================================
def parse_pivot(pivot_df: pd.DataFrame | None) -> dict:
    """피벗 시트(위치 기반)를 파싱해 매핑 dict 반환.
    반환: code_to_short, af_to_930, date_to_label, date_to_close(=(y,m)), date_to_week(int)
    """
    maps = {"code_to_short": {}, "af_to_930": {}, "date_to_label": {},
            "date_to_close": {}, "date_to_week": {}}
    if pivot_df is None or pivot_df.empty:
        return maps
    p = pivot_df.reset_index(drop=True)

    def col(i):
        return p.iloc[:, i] if (i is not None and 0 <= i < p.shape[1]) else pd.Series([], dtype=object)

    # 리더(calamine/openpyxl)에 따라 빈 선두열 트림 여부가 달라 열 인덱스가 흔들린다.
    # 따라서 헤더 마커(<제휴사 구분>/<주차기준>/<제휴사명 정제>)의 열 위치를 찾아 상대 참조한다.
    a_start = b_start = c_start = None
    for ri in range(min(3, len(p))):
        for j in range(p.shape[1]):
            s = str(p.iat[ri, j])
            if a_start is None and "제휴사" in s and "구분" in s:
                a_start = j
            if b_start is None and "주차" in s:
                b_start = j
            if c_start is None and "정제" in s:
                c_start = j
        if a_start is not None and b_start is not None and c_start is not None:
            break

    # 블록 A (제휴사 구분): [start]=코드(000127-SKT / 930007-현대3F), [start+1]=약칭(SKT)
    for code, short in zip(col(a_start), col(a_start + 1 if a_start is not None else None)):
        if pd.notna(code) and pd.notna(short) and str(short).strip() and not str(code).startswith("<"):
            maps["code_to_short"][str(code).strip()] = str(short).strip()

    # 블록 C (제휴사명 정제): [start]=AF코드, [start+1]=930명
    for af, n930 in zip(col(c_start), col(c_start + 1 if c_start is not None else None)):
        if pd.notna(af) and pd.notna(n930) and not str(af).startswith("<"):
            maps["af_to_930"][str(af).strip()] = str(n930).strip()

    # 블록 B (주차기준): [start]=일자, [start+2]=주차라벨(YY_M_W)
    dates = pd.to_datetime(col(b_start), errors="coerce")
    labels = col(b_start + 2 if b_start is not None else None)
    for d, lab in zip(dates, labels):
        if pd.isna(d) or pd.isna(lab) or str(lab).startswith("<"):
            continue
        key = d.date()
        lab = str(lab).strip()
        maps["date_to_label"][key] = lab
        ym, wk = _parse_week_label(lab)
        if ym:
            maps["date_to_close"][key] = ym
            maps["date_to_week"][key] = wk
    return maps


def _parse_week_label(lab: str):
    """'26_1_1' → ((2026,1), 1). 실패 시 (None, None)."""
    m = re.match(r"^\s*(\d{2,4})[_\-.](\d{1,2})[_\-.](\d{1,2})", str(lab))
    if not m:
        return None, None
    y = int(m.group(1))
    if y < 100:
        y += 2000
    return (y, int(m.group(2))), int(m.group(3))


# ======================================================================================
# 2) enrich: 파생 컬럼을 실제 값으로 채움 (수식 재현). 이미 값이 있으면 보존(전년 파일 대비).
# ======================================================================================
def enrich(data: dict, maps: dict) -> dict:
    amt, uv, cert = data.get("amt"), data.get("uv"), data.get("cert")
    c2s = maps.get("code_to_short", {})
    af930 = maps.get("af_to_930", {})
    d2close = maps.get("date_to_close", {})
    d2week = maps.get("date_to_week", {})
    d2label = maps.get("date_to_label", {})

    if amt is not None and not amt.empty:
        # 거래액_VAT제외 = 거래액/1.1 (캐시 비었으면 재계산)
        if "거래액" in amt.columns:
            gross = pd.to_numeric(amt["거래액"], errors="coerce")
            vat = pd.to_numeric(amt.get("거래액_VAT제외"), errors="coerce")
            need = vat.isna() | (vat == 0)
            amt["거래액_VAT제외"] = vat.where(~need, gross / 1.1).fillna(0.0)
        # 당월인증: 정산 YYYY-MM == 최초인증 YYYY-MM
        cur = (amt["당월인증"].astype(str).str.strip().str.upper()
               if "당월인증" in amt.columns else pd.Series("", index=amt.index))
        has = cur.isin(["Y", "N"])
        if not has.all() and "최초인증일일" in amt.columns:
            s = pd.to_datetime(amt["정산일시일"], errors="coerce").dt.strftime("%Y-%m")
            fst = pd.to_datetime(amt["최초인증일일"], errors="coerce").dt.strftime("%Y-%m")
            derived = pd.Series(np.where((s == fst) & s.notna(), "Y", "N"), index=amt.index)
            amt["당월인증"] = cur.where(has, derived)
        else:
            amt["당월인증"] = cur.where(has, "N")
        # 제휴사: 코드(제휴처구분3)→약칭
        if "제휴처구분3" in amt.columns:
            need_p = "제휴사" not in amt.columns or amt["제휴사"].isna().all()
            if need_p:
                key = amt["제휴처구분3"].astype(str).str.strip()
                amt["제휴사"] = key.map(c2s).fillna(key.str.replace(r"^[0-9]+-", "", regex=True))
        # 마감월 / 주차 (피벗 기준, 없으면 캘린더 폴백)
        dser = pd.to_datetime(amt["정산일시일"], errors="coerce").dt.date
        amt["_마감"] = dser.map(lambda d: _ym_str(d2close.get(d)) or _cal_ym(d))
        amt["_주차"] = dser.map(lambda d: d2label.get(d) or "")
        amt["_주차N"] = dser.map(lambda d: d2week.get(d) or 0)

    if cert is not None and not cert.empty:
        code_col = "제휴처구분4" if "제휴처구분4" in cert.columns else "제휴처구분3"
        if code_col in cert.columns:
            need_p = "제휴사" not in cert.columns or cert["제휴사"].isna().all()
            if need_p:
                key = cert[code_col].astype(str).str.strip()
                cert["제휴사"] = key.map(c2s).fillna(key.str.replace(r"^[0-9]+-", "", regex=True))
        dser = pd.to_datetime(cert["인증일시일"], errors="coerce").dt.date
        cert["_마감"] = dser.map(lambda d: _ym_str(d2close.get(d)) or _cal_ym(d))
        cert["_주차"] = dser.map(lambda d: d2label.get(d) or "")

    if uv is not None and not uv.empty:
        need_p = "제휴사명" not in uv.columns or uv["제휴사명"].isna().all() or (uv["제휴사명"].astype(str).str.strip() == "").all()
        if need_p and "AF코드" in uv.columns:
            af = uv["AF코드"].astype(str).str.strip()
            uv["제휴사명"] = af.map(lambda a: c2s.get(af930.get(a, ""), af930.get(a, a)))
        dser = pd.to_datetime(uv["★일자일"], errors="coerce").dt.date
        uv["_마감"] = dser.map(lambda d: _ym_str(d2close.get(d)) or _cal_ym(d))
        uv["_주차"] = dser.map(lambda d: d2label.get(d) or "")

    return data


def _ym_str(ym):
    return f"{ym[0]}-{ym[1]:02d}" if ym else None


def _cal_ym(d):
    return f"{d.year}-{d.month:02d}" if isinstance(d, dt.date) else None


# ======================================================================================
# 3) 기간 유틸: 마감월별 날짜 목록 (월마감 / MTD)
# ======================================================================================
def available_months(*datasets) -> list:
    ms = set()
    for d in datasets:
        for df in (d.get("amt"), d.get("uv"), d.get("cert")):
            if df is not None and not df.empty and "_마감" in df.columns:
                ms |= set(df["_마감"].dropna().unique().tolist())
    return sorted(m for m in ms if m)


def month_dates(data: dict, ym: str, maps: dict, mtd_day: int | None = None) -> list:
    """해당 마감월(ym)에 속하는 실제 날짜 목록. 피벗 date_to_close 우선, 없으면 데이터에서 수집.
    mtd_day가 주어지면 캘린더 일자 <= mtd_day 만."""
    dates = set()
    for d, close in maps.get("date_to_close", {}).items():
        if _ym_str(close) == ym:
            dates.add(d)
    if not dates:  # 피벗에 없으면 데이터에서
        for df, dc in ((data.get("amt"), "정산일시일"), (data.get("uv"), "★일자일"), (data.get("cert"), "인증일시일")):
            if df is not None and not df.empty and "_마감" in df.columns:
                sub = df[df["_마감"] == ym]
                dates |= set(pd.to_datetime(sub[dc], errors="coerce").dt.date.dropna().tolist())
    out = sorted(d for d in dates if d)
    if mtd_day is not None:
        out = [d for d in out if d.day <= mtd_day]
    return out


def weeks_of_month(data: dict, ym: str) -> list:
    """해당 마감월의 주차 라벨 목록(정렬)."""
    labels = set()
    amt = data.get("amt")
    for df in (data.get("amt"), data.get("uv"), data.get("cert")):
        if df is not None and not df.empty and "_마감" in df.columns and "_주차" in df.columns:
            sub = df[df["_마감"] == ym]
            labels |= set(x for x in sub["_주차"].dropna().unique().tolist() if x)
    return sorted(labels, key=lambda l: _parse_week_label(l)[1] or 0)


def week_dates(data: dict, label: str) -> list:
    dates = set()
    for df, dc in ((data.get("amt"), "정산일시일"), (data.get("uv"), "★일자일"), (data.get("cert"), "인증일시일")):
        if df is not None and not df.empty and "_주차" in df.columns:
            sub = df[df["_주차"] == label]
            dates |= set(pd.to_datetime(sub[dc], errors="coerce").dt.date.dropna().tolist())
    return sorted(d for d in dates if d)


# ======================================================================================
# 4) 핵심 지표 묶음 (한 기간에 대한 UV/인증/당월인증거래액/전체거래액/고객수/객단가/CR)
# ======================================================================================
def metric_bundle(data: dict, dates: list, partners=None, exclude_af=None) -> dict:
    """한 기간(dates)에 대한 지표 묶음. 세그(신규/윈백/기존) 포함.
    반환 dict 키: uv, cert(전체/신규/WIN-BACK/기존), amt_all, amt_cert (각 tot/net/cust_tot/cust_net, 세그별)"""
    uv = A.uv_total(data.get("uv", pd.DataFrame()), dates, partners, exclude_af=exclude_af) if data.get("uv") is not None else 0
    cert = A.cert_summary(data.get("cert", pd.DataFrame()), dates, partners) if data.get("cert") is not None else {"전체": 0, "기존": 0, "WIN-BACK": 0, "신규": 0}
    amt_all = A.amt_summary(data.get("amt", pd.DataFrame()), dates, partners, cert_only=False)
    amt_cert = A.amt_summary(data.get("amt", pd.DataFrame()), dates, partners, cert_only=True)
    return {"uv": uv, "cert": cert, "amt_all": amt_all, "amt_cert": amt_cert}


# ======================================================================================
# 5) 기여도(상승/하락) & BPU 추가분석
# ======================================================================================
def contribution_table(amt_cur, dates_cur, amt_prev, dates_prev, groupcol, pt="net",
                       cert_only=True, top=12) -> pd.DataFrame:
    """그룹(카테고리/브랜드)별 전년비·비중·기여도. pt: 'net'(순결제)/'tot'(총결제).
    기여도 = 그룹 증감 / 전체 증감 (합계 100%). |증감| 큰 순으로 top개."""
    cur = A.group_amt_table(amt_cur, dates_cur, None, groupcol, cert_only=cert_only)
    cur = _pick(cur, pt).rename(columns={"val": "cur"})
    if amt_prev is not None and not amt_prev.empty:
        prev = A.group_amt_table(amt_prev, dates_prev, None, groupcol, cert_only=cert_only)
        prev = _pick(prev, pt).rename(columns={"val": "prev"})
        m = pd.merge(cur, prev, on=groupcol, how="outer").fillna(0.0)
    else:
        m = cur.copy()
        m["prev"] = np.nan
    m["delta"] = m["cur"] - m["prev"].fillna(0)
    tot_cur = m["cur"].sum()
    tot_delta = m["delta"].sum()
    m["share"] = np.where(tot_cur != 0, m["cur"] / tot_cur, 0)
    m["yoy"] = np.where(m["prev"].fillna(0) != 0, m["cur"] / m["prev"] - 1, np.nan)
    m["contrib"] = np.where(tot_delta != 0, m["delta"] / tot_delta, 0)
    m = m.reindex(m["delta"].abs().sort_values(ascending=False).index)
    return m.head(top).reset_index(drop=True)


def _pick(df, pt):
    if df.empty:
        return pd.DataFrame(columns=[df.columns[0] if len(df.columns) else "grp", "val"])
    gc = df.columns[0]
    return df[[gc, pt]].rename(columns={pt: "val"})


def bpu_decomposition(amt_cur, dates_cur, amt_prev, dates_prev, pt="net", cert_only=True) -> pd.DataFrame:
    """BPU별 객단가 분해: 거래액 증감을 '고객수 효과'와 '객단가 효과'로 분리.
    ΔRev = Δcust·aov_prev + cust_prev·Δaov + Δcust·Δaov (교차항은 객단가 효과에 합산)."""
    def _agg(amt, dates):
        sub = A._filter_amt(amt, dates, None, cert_only)
        if sub.empty:
            return pd.DataFrame(columns=["BPU", "rev", "cust"])
        base = sub[sub["정산구분"] == "판매"] if pt == "tot" else sub
        rev = base.groupby("BPU")["거래액_VAT제외"].sum()
        cust = base.groupby("BPU")["고객번호"].nunique()
        return pd.DataFrame({"rev": rev, "cust": cust}).reset_index()

    cur = _agg(amt_cur, dates_cur).rename(columns={"rev": "rev_c", "cust": "cust_c"})
    if amt_prev is not None and not amt_prev.empty:
        prev = _agg(amt_prev, dates_prev).rename(columns={"rev": "rev_p", "cust": "cust_p"})
        m = pd.merge(cur, prev, on="BPU", how="outer").fillna(0.0)
    else:
        m = cur.copy()
        m["rev_p"] = np.nan
        m["cust_p"] = np.nan
    m["aov_c"] = np.where(m["cust_c"] != 0, m["rev_c"] / m["cust_c"], 0)
    m["aov_p"] = np.where(m["cust_p"].fillna(0) != 0, m["rev_p"] / m["cust_p"], np.nan)
    m["d_rev"] = m["rev_c"] - m["rev_p"].fillna(0)
    m["cust_effect"] = (m["cust_c"] - m["cust_p"].fillna(0)) * m["aov_p"].fillna(0)
    m["aov_effect"] = m["d_rev"] - m["cust_effect"]
    return m.sort_values("rev_c", ascending=False).reset_index(drop=True)


def return_rate_table(amt_cur, dates_cur, groupcol="BPU", cert_only=True) -> pd.DataFrame:
    """반품률 = (총결제 - 순결제) / 총결제, 그룹별."""
    t = A.group_amt_table(amt_cur, dates_cur, None, groupcol, cert_only=cert_only)
    if t.empty:
        return t
    t["반품률"] = np.where(t["tot"] != 0, (t["tot"] - t["net"]) / t["tot"], 0)
    return t.sort_values("tot", ascending=False).reset_index(drop=True)


def new_penetration_table(amt_cur, dates_cur, groupcol="BPU", pt="net", cert_only=True) -> pd.DataFrame:
    """신규 침투율 = 신규 거래액 / 전체 거래액, 그룹별."""
    sub = A._filter_amt(amt_cur, dates_cur, None, cert_only)
    if sub.empty:
        return pd.DataFrame(columns=[groupcol, "전체", "신규", "침투율"])
    base = sub[sub["정산구분"] == "판매"] if pt == "tot" else sub
    total = base.groupby(groupcol)["거래액_VAT제외"].sum()
    new = base[base["기존/win-back/신규"] == "신규"].groupby(groupcol)["거래액_VAT제외"].sum()
    out = pd.DataFrame({"전체": total, "신규": new}).fillna(0.0).reset_index()
    out["침투율"] = np.where(out["전체"] != 0, out["신규"] / out["전체"], 0)
    return out.sort_values("전체", ascending=False).reset_index(drop=True)


def pareto_summary(amt_cur, dates_cur, groupcol="Admin브랜드명", pt="net", cert_only=True) -> dict:
    """상위 5/10/20 브랜드가 거래액에서 차지하는 누적 비중."""
    t = A.group_amt_table(amt_cur, dates_cur, None, groupcol, cert_only=cert_only)
    if t.empty:
        return {}
    vals = t[pt].sort_values(ascending=False).reset_index(drop=True)
    total = vals.sum()
    if total == 0:
        return {}
    out = {"total_n": int((vals > 0).sum())}
    for n in (5, 10, 20):
        out[f"top{n}"] = float(vals.head(n).sum() / total)
    return out


# ======================================================================================
# 6) 목표 / LF몰 전체거래액 로더 (월별 wide 템플릿)
# ======================================================================================
def _norm_ym(v) -> str | None:
    if isinstance(v, (dt.date, dt.datetime)):
        return f"{v.year}-{v.month:02d}"
    s = str(v).strip()
    digits = re.sub(r"[^0-9]", "", s)
    if len(digits) >= 6:
        return f"{digits[:4]}-{int(digits[4:6]):02d}"
    m = re.match(r"^(\d{4})[-./](\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    return None


def load_target(file_bytes: bytes) -> pd.DataFrame:
    """목표 파일 로더. 기대 구조: 열 [지표, 세그, <월들...>].
    지표 ∈ {UV, 인증자수, 당월인증거래액, 거래액}, 세그 ∈ {전체, 신규, 윈백, 기존}.
    반환(long): 지표, 세그, ym, 목표값. '전체'가 없으면 세그 합으로 생성."""
    sheets = A._read_all_sheets(file_bytes)
    data = None
    for sn, d in sheets.items():
        if d and any("지표" in str(c) for c in (d[0] or [])):
            data = d
            break
    if data is None:
        data = next(iter(sheets.values()), None)
    if not data:
        return pd.DataFrame(columns=["지표", "세그", "ym", "목표"])
    header = [str(c).strip() for c in data[0]]
    # 월 컬럼 위치
    ym_cols = {j: _norm_ym(header[j]) for j in range(len(header)) if _norm_ym(header[j])}
    try:
        i_metric = next(j for j, c in enumerate(header) if "지표" in c)
    except StopIteration:
        i_metric = 0
    i_seg = next((j for j, c in enumerate(header) if "세그" in c or "구분" in c), 1)
    rows = []
    for r in data[1:]:
        if not r or i_metric >= len(r) or r[i_metric] in (None, ""):
            continue
        metric = str(r[i_metric]).strip()
        seg = str(r[i_seg]).strip() if i_seg < len(r) and r[i_seg] not in (None, "") else "전체"
        seg = {"win-back": "윈백", "WIN-BACK": "윈백"}.get(seg, seg)
        for j, ym in ym_cols.items():
            if j < len(r) and r[j] not in (None, ""):
                try:
                    rows.append({"지표": metric, "세그": seg, "ym": ym, "목표": float(r[j])})
                except (TypeError, ValueError):
                    pass
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    # 전체 보강
    out = [df]
    for (metric, ym), g in df.groupby(["지표", "ym"]):
        if "전체" not in set(g["세그"]):
            segs = g[g["세그"].isin(["신규", "윈백", "기존"])]
            if not segs.empty:
                out.append(pd.DataFrame([{"지표": metric, "세그": "전체", "ym": ym, "목표": segs["목표"].sum()}]))
    return pd.concat(out, ignore_index=True)


def target_value(target_df, 지표, ym, 세그="전체"):
    if target_df is None or target_df.empty:
        return None
    sub = target_df[(target_df["지표"] == 지표) & (target_df["ym"] == ym) & (target_df["세그"] == 세그)]
    return float(sub["목표"].iloc[0]) if not sub.empty else None


def load_lfmall(file_bytes: bytes) -> pd.DataFrame:
    """LF몰 전체거래액 로더. 기대 구조: 열 [구분, <월들...>], 구분 ∈ {목표, 실제}.
    (또는 열 [월, 목표, 실제]). 반환: ym, 목표, 실제."""
    sheets = A._read_all_sheets(file_bytes)
    data = next(iter(sheets.values()), None)
    if not data:
        return pd.DataFrame(columns=["ym", "목표", "실제"])
    header = [str(c).strip() for c in data[0]]
    ym_cols = {j: _norm_ym(header[j]) for j in range(len(header)) if _norm_ym(header[j])}
    result = {}
    if ym_cols:  # 가로형: 구분 x 월
        i_kind = next((j for j, c in enumerate(header) if "구분" in c or "목표" in c or "실제" in c), 0)
        for r in data[1:]:
            if not r or i_kind >= len(r) or r[i_kind] in (None, ""):
                continue
            kind = "목표" if "목표" in str(r[i_kind]) else ("실제" if ("실제" in str(r[i_kind]) or "실적" in str(r[i_kind])) else None)
            if kind is None:
                continue
            for j, ym in ym_cols.items():
                if j < len(r) and r[j] not in (None, ""):
                    result.setdefault(ym, {})[kind] = _f(r[j])
    else:  # 세로형: 월 | 목표 | 실제
        i_ym = 0
        i_t = next((j for j, c in enumerate(header) if "목표" in c), 1)
        i_a = next((j for j, c in enumerate(header) if "실제" in c or "실적" in c), 2)
        for r in data[1:]:
            ym = _norm_ym(r[i_ym]) if r and len(r) > i_ym else None
            if not ym:
                continue
            result.setdefault(ym, {})["목표"] = _f(r[i_t]) if len(r) > i_t else None
            result[ym]["실제"] = _f(r[i_a]) if len(r) > i_a else None
    rows = [{"ym": ym, "목표": v.get("목표"), "실제": v.get("실제")} for ym, v in result.items()]
    return pd.DataFrame(rows)


def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def lfmall_value(lf_df, ym, kind="실제"):
    if lf_df is None or lf_df.empty:
        return None
    sub = lf_df[lf_df["ym"] == ym]
    if sub.empty:
        return None
    v = sub[kind].iloc[0]
    return None if pd.isna(v) else float(v)
