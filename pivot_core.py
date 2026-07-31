"""pivot_core.py — 경량 피벗 앱의 순수 로직(Streamlit 비의존).

지표 정의·집계·피벗·전년비 매트릭스·분석일 진단(특이점 플래그)을 담는다.
UI(streamlit_app.py)와 분리해 단위 테스트가 가능하도록 한다.
컨벤션: 전년비 상승 ▼ / 하락 △ (색은 UI에서 초록/빨강).
"""
from __future__ import annotations
import datetime as dt

import numpy as np
import pandas as pd

# 지표 정의 — base: 합산 가능한 원천 컬럼(pay별). aov = 당월인증거래액/당월인증 고객수.
METRICS = {
    "UV":               {"kind": "base", "net": "uv",           "tot": "uv",           "pay": False, "fmt": "int"},
    "인증자수":          {"kind": "base", "net": "cert_total",   "tot": "cert_total",   "pay": False, "fmt": "int"},
    "당월인증거래액":     {"kind": "base", "net": "cert_amt_net", "tot": "cert_amt_tot", "pay": True,  "fmt": "won"},
    "전체거래액":        {"kind": "base", "net": "all_amt_net",  "tot": "all_amt_tot",  "pay": True,  "fmt": "won"},
    "고객수(당월인증)":   {"kind": "base", "net": "cust_net",     "tot": "cust_tot",     "pay": True,  "fmt": "int"},
    "객단가(당월인증)":   {"kind": "aov",                                                "pay": True,  "fmt": "won0"},
}
KEY_METRICS_FOR_COMMENT = ["UV", "인증자수", "당월인증거래액", "전체거래액", "객단가(당월인증)"]


def na(v) -> bool:
    return v is None or (isinstance(v, float) and np.isnan(v))


def fmt_int(v):
    return "" if na(v) else f"{int(round(float(v))):,}"


def fmt_won(v):
    if na(v):
        return ""
    v = float(v)
    a = abs(v)
    if a >= 1e8:
        return f"{v/1e8:,.2f}억"
    if a >= 1e4:
        return f"{v/1e4:,.0f}만"
    return f"{v:,.0f}"


FMT = {"int": fmt_int, "won": fmt_won, "won0": fmt_int}


def m_shift(mlabel: str, k: int) -> str:
    y, mm = mlabel.split("-")
    return f"{int(y)+k}-{mm}"


def m_label(mlabel: str) -> str:
    return mlabel[2:4] + "." + mlabel[5:7]  # 2026-07 -> 26.07


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """레코드 DataFrame에 month/day/mmdd 파생 + 숫자형 보정."""
    if df.empty:
        return df
    df = df.copy()
    df["month"] = df["date"].str.slice(0, 7)
    df["day"] = df["date"].str.slice(8, 10).astype(int)
    df["mmdd"] = df["date"].str.slice(5, 10)
    num = [c for c in df.columns if c not in ("year_tag", "date", "affiliate", "month", "mmdd")]
    df[num] = df[num].apply(pd.to_numeric, errors="coerce").fillna(0)
    return df


def base_col(metric, pay):
    m = METRICS[metric]
    if not m.get("pay"):
        return m["net"]
    return m["net"] if pay == "net" else m["tot"]


def agg_value(d, by, metric, pay):
    """by=None → 스칼라 총계, by=컬럼명 → 그 축 Series. 객단가는 sum(거래액)/sum(고객수)."""
    m = METRICS[metric]
    if m["kind"] == "aov":
        an, cn = f"cert_amt_{pay}", f"cust_{pay}"
        if by is None:
            den = d[cn].sum()
            return (d[an].sum() / den) if den else np.nan
        num = d.groupby(by)[an].sum()
        den = d.groupby(by)[cn].sum().replace(0, np.nan)
        return num / den
    col = base_col(metric, pay)
    return d[col].sum() if by is None else d.groupby(by)[col].sum()


def pivot_values(d, index, columns, metric, pay):
    if d.empty:
        return pd.DataFrame()
    m = METRICS[metric]
    if m["kind"] == "aov":
        num = d.pivot_table(index=index, columns=columns, values=f"cert_amt_{pay}", aggfunc="sum")
        den = d.pivot_table(index=index, columns=columns, values=f"cust_{pay}", aggfunc="sum")
        return num.div(den.replace(0, np.nan))
    return d.pivot_table(index=index, columns=columns, values=base_col(metric, pay), aggfunc="sum")


def value_matrix(d, index, columns, metric, pay, index_order, col_order):
    """index_order×col_order 고정 + 합계(상단 행/좌측 열, 모서리=총계)로 감싼 값 매트릭스."""
    piv = pivot_values(d, index, columns, metric, pay)
    piv = piv.reindex(index=index_order, columns=col_order) if not piv.empty \
        else pd.DataFrame(index=index_order, columns=col_order, dtype=float)
    row_tot = agg_value(d, index, metric, pay)
    col_tot = agg_value(d, columns, metric, pay)
    row_tot = row_tot.reindex(index_order) if hasattr(row_tot, "reindex") else pd.Series(dtype=float)
    col_tot = col_tot.reindex(col_order) if hasattr(col_tot, "reindex") else pd.Series(dtype=float)
    grand = agg_value(d, None, metric, pay)
    body = piv.copy()
    body.insert(0, "합계", row_tot)
    top = pd.DataFrame([[grand] + [col_tot.get(c, np.nan) for c in col_order]],
                       index=["합계"], columns=body.columns)
    return pd.concat([top, body])


def yoy_matrix(cur_mat, prev_mat):
    prev = prev_mat.reindex(index=cur_mat.index, columns=cur_mat.columns)
    with np.errstate(divide="ignore", invalid="ignore"):
        yy = (cur_mat.astype(float) / prev.astype(float) - 1.0) * 100.0
    return yy.replace([np.inf, -np.inf], np.nan)


def _pct(cur_v, ref_v):
    if na(ref_v) or ref_v == 0 or na(cur_v):
        return None
    return (cur_v / ref_v - 1) * 100


def analyze_day(df: pd.DataFrame, day_date: str, affiliate: str, pay: str,
                lfmall_daily: dict | None = None) -> dict:
    """분석일 하루를 전년 동일자·동요일평균과 대조해 지표값과 특이점 플래그를 산출(순수).

    반환: {"metrics": {지표: {cur,prev,wd,yoy,wd_dev}}, "flags": [...],
           "returns": {총/순/취소 대조}, "lf_share": {LF몰 비중} | None}
    - 헤드라인 지표값은 인자 `pay`(사이드바 결제구분)를 그대로 따른다.
    - `returns`: 당월인증거래액을 총결제(tot)·순결제(net) 모두 산출 → 취소·환불(취반품)
      영향인지, 순수 매출 변화인지 판별용. `pay`와 무관하게 항상 둘 다 담는다.
    - `lf_share`: lfmall_daily(일자별 LF몰 전체거래액)가 주어지면 순결제 당월인증거래액의
      LF몰 대비 비중을 산출(LF몰 전체거래액은 순결제만 존재 → 순결제 기준으로만 의미).
    플래그(회사 규칙):
      - uv_spike: UV YoY가 인증 YoY보다 크게 높음 → 봇 단정 금지(리워드·앱푸시 비회원 유입)
      - cert_vs_total: 당월인증거래액 급감 + 전체거래액 유지/증가 → 신규 인증 유입 착시
      - cancel: 순결제·총결제 전년비 괴리 → 취소·환불(취반품)이 실적을 밀거나 방어
      - wd_dev: 동요일평균 대비 ±30% 이상 이탈(요일 편차 큰 특성상 전일비 아닌 동요일 기준)
    """
    cur = df[df.year_tag == "cur"]
    prev = df[df.year_tag == "prev"]
    if affiliate != "전체":
        cur = cur[cur.affiliate == affiliate]
        prev = prev[prev.affiliate == affiliate]
    month = day_date[:7]
    mmdd = day_date[5:]
    wd = dt.date.fromisoformat(day_date).weekday()
    day_sub = cur[cur.date == day_date]
    prev_sub = prev[prev.mmdd == mmdd]

    def wd_avg(metric):
        gm = cur[cur.month == month]
        if gm.empty:
            return np.nan
        vals = [agg_value(gm[gm.date == d], None, metric, pay)
                for d in sorted(gm.date.unique())
                if dt.date.fromisoformat(d).weekday() == wd and d != day_date]
        vals = [v for v in vals if not na(v)]
        return float(np.mean(vals)) if vals else np.nan

    metrics = {}
    for metric in KEY_METRICS_FOR_COMMENT:
        cv = agg_value(day_sub, None, metric, pay) if not day_sub.empty else np.nan
        pv = agg_value(prev_sub, None, metric, pay) if not prev_sub.empty else np.nan
        wa = wd_avg(metric)
        metrics[metric] = {"cur": cv, "prev": pv, "wd": wa,
                           "yoy": _pct(cv, pv), "wd_dev": _pct(cv, wa)}

    # ── 총결제·순결제 대조 (취반품 판별) — 당월인증거래액을 두 기준 모두 산출 ──
    def _cav(sub, p):
        return agg_value(sub, None, "당월인증거래액", p) if not sub.empty else np.nan
    net_c, tot_c = _cav(day_sub, "net"), _cav(day_sub, "tot")
    net_p, tot_p = _cav(prev_sub, "net"), _cav(prev_sub, "tot")
    cancel_c = (tot_c - net_c) if not (na(tot_c) or na(net_c)) else np.nan  # 취소·환불액
    cancel_p = (tot_p - net_p) if not (na(tot_p) or na(net_p)) else np.nan
    net_yoy, tot_yoy = _pct(net_c, net_p), _pct(tot_c, tot_p)
    returns = {
        "net_cur": net_c, "tot_cur": tot_c, "cancel_cur": cancel_c,
        "net_prev": net_p, "tot_prev": tot_p, "cancel_prev": cancel_p,
        "net_yoy": net_yoy, "tot_yoy": tot_yoy,
        "cancel_ratio_cur": (cancel_c / tot_c) if (not na(cancel_c) and tot_c) else None,
        "cancel_ratio_prev": (cancel_p / tot_p) if (not na(cancel_p) and tot_p) else None,
    }

    # ── LF몰 대비 비중 (순결제 당월인증거래액 ÷ LF몰 전체거래액, 같은 일자) ──
    lf_share = None
    if lfmall_daily:
        prevdate = f"{int(day_date[:4]) - 1}-{mmdd}"
        lf_c = lfmall_daily.get(day_date)
        lf_p = lfmall_daily.get(prevdate)
        share_c = (net_c / lf_c * 100) if (lf_c and not na(net_c)) else None
        share_p = (net_p / lf_p * 100) if (lf_p and not na(net_p)) else None
        lf_share = {"share_cur": share_c, "share_prev": share_p, "lf_cur": lf_c, "lf_prev": lf_p,
                    "dpp": (share_c - share_p) if (share_c is not None and share_p is not None) else None}

    flags = []
    uv_y, cert_y = metrics["UV"]["yoy"], metrics["인증자수"]["yoy"]
    # "UV만" 급증: UV가 크게 오르되 인증은 평탄/하락(같이 크게 오르면 정상 성장이므로 제외)
    if uv_y is not None and cert_y is not None and uv_y >= 30 and cert_y <= 15 and uv_y - cert_y >= 30:
        flags.append({"code": "uv_spike",
                      "text": f"UV만 급증(전년비 ▼ +{uv_y:.0f}%, 인증자수는 "
                              f"{'▼ +' if cert_y >= 0 else '△ '}{cert_y:.0f}%) — 봇으로 단정하지 말 것. "
                              "리워드·앱푸시로 인한 비회원 유입일 수 있음. UV의 인증 전환 여부가 핵심."})
    cav, aav = metrics["당월인증거래액"]["yoy"], metrics["전체거래액"]["yoy"]
    if cav is not None and aav is not None and cav <= -30 and aav >= 0:
        flags.append({"code": "cert_vs_total",
                      "text": f"당월인증거래액 급감(전년비 △ {cav:.0f}%)이 곧 부진은 아님 — "
                              f"전체거래액은 ▼ +{aav:.0f}%. 신규 인증 유입 감소에 따른 착시일 수 있으니 "
                              "전체거래액과 함께 판단."})
    # 취반품 판별: 순결제(취소 차감)와 총결제(판매만)의 전년비가 크게 벌어지면
    # 취소·환불 자체가 실적을 밀거나 방어한 것. ±10%p 이상 벌어질 때만 플래그.
    if net_yoy is not None and tot_yoy is not None:
        gap = net_yoy - tot_yoy
        if gap <= -10:
            flags.append({"code": "cancel",
                          "text": f"순결제(전년비 {'▼ +' if net_yoy >= 0 else '△ '}{net_yoy:.0f}%)가 "
                                  f"총결제({'▼ +' if tot_yoy >= 0 else '△ '}{tot_yoy:.0f}%)보다 부진 — "
                                  "취소·환불(취반품) 증가가 실적을 깎은 것. 판매 자체보다 취반품이 하방 요인."})
        elif gap >= 10:
            flags.append({"code": "cancel",
                          "text": f"순결제(전년비 {'▼ +' if net_yoy >= 0 else '△ '}{net_yoy:.0f}%)가 "
                                  f"총결제({'▼ +' if tot_yoy >= 0 else '△ '}{tot_yoy:.0f}%)보다 선방 — "
                                  "취소·환불(취반품)이 전년보다 줄어 순결제를 방어. 판매 개선과는 별개."})
    devs = []
    for metric in KEY_METRICS_FOR_COMMENT:
        p = metrics[metric]["wd_dev"]
        if p is not None and abs(p) >= 30:
            devs.append(f"{metric} {'▼ +' if p >= 0 else '△ '}{p:.0f}%")
    if devs:
        flags.append({"code": "wd_dev",
                      "text": "동요일평균 대비 큰 이탈: " + " / ".join(devs)
                              + " (요일 편차가 커 전일비 아닌 동요일 평균 기준)"})
    return {"metrics": metrics, "flags": flags, "returns": returns, "lf_share": lf_share}
