"""
SSP_CS Performance Dashboard (Streamlit)

Reads the Google Sheet "OFM(CS)_SSP Performance Tracking" directly, so the dashboard
follows the sheet: the data is re-downloaded every REFRESH_SECONDS (and on the
Refresh button) and the page re-renders itself on that schedule.

Sheet tabs used
  CS_By Date          daily tracker: Call, Chat (Line), Chat (Facebook), Email,
                      Case Management, Complaint
  SSP_Campaign Plan   campaign calendar (period, main campaign, sub-campaign)

Requirements: streamlit>=1.37, pandas, plotly, openpyxl, requests

The sheet must be shared as "Anyone with the link can view" (it is today). For a
private sheet, see load_workbook_bytes() for the service-account option.
"""
from __future__ import annotations

import datetime as dt
import io
import os
import re
from collections import OrderedDict

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import openpyxl

# ----------------------------------------------------------------------------- settings
SHEET_ID = "1-25DimfROhlDz379xZUVQ4SCitavqRolfm5Xwh2y_J0"
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit"
EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
REFRESH_SECONDS = 300          # how often the page re-downloads the sheet (5 minutes)
LOCAL_XLSX = os.environ.get("SSP_LOCAL_XLSX")  # optional: path to an offline copy for testing

CHANNELS = OrderedDict([
    ("Call", "call"), ("Chat (Line)", "line"), ("Chat (Facebook)", "facebook"),
    ("Email", "email"), ("Case Management", "case"), ("Complaint", "complaint"),
])
SEG_NAME = {"total": "Total", "call": "Call", "line": "Chat (Line)", "facebook": "Chat (Facebook)",
            "email": "Email", "case": "Case Management", "complaint": "Complaint"}
CONTACT = ["call", "line", "facebook", "email"]
COLOR = {"call": "#4C78A8", "line": "#54A24B", "facebook": "#F58518", "email": "#B279A2",
         "case": "#72B7B2", "complaint": "#E45756", "total": "#333333"}
KIND_COLOR = {"Mega Sale": "rgba(228,87,86,0.18)", "Brand Day": "rgba(76,120,168,0.18)", "Seasonal": "rgba(84,162,75,0.18)"}

# Customer service standards: (Excellent, minimum, higher_is_better, format)
STD = {
    "abandon_call":       ("Call abandon rate",            0.03, 0.05, False, "pct"),
    "abandon_chat":       ("Chat abandon rate",            0.03, 0.05, False, "pct"),
    "abandon_email":      ("Email unanswered rate",        0.01, 0.02, False, "pct"),
    "abandon_all":        ("Abandon rate, all channels",   0.03, 0.05, False, "pct"),
    "wait":               ("Average call waiting time",    20,   30,   False, "sec"),
    "csat":               ("CSAT",                         0.95, 0.90, True,  "pct"),
    "case_complete":      ("Case completion rate",         0.95, 0.90, True,  "pct"),
    "complaint_resolved": ("Complaint resolution rate",    0.90, 0.80, True,  "pct"),
}

# ----------------------------------------------------------------------------- loading
@st.cache_data(ttl=REFRESH_SECONDS, show_spinner="Loading the Google Sheet…")
def load_workbook_bytes() -> tuple[bytes, dt.datetime]:
    """Download the sheet as xlsx. Cached for REFRESH_SECONDS, so every reader within
    that window shares one download and the data is never older than that."""
    if LOCAL_XLSX:
        with open(LOCAL_XLSX, "rb") as f:
            return f.read(), dt.datetime.now()
    # Private sheet alternative (uncomment, add a service account JSON to st.secrets):
    #   import gspread; gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])
    #   content = gc.export(SHEET_ID, format=gspread.utils.ExportFormat.EXCEL)
    r = requests.get(EXPORT_URL, timeout=60)
    r.raise_for_status()
    if not r.content.startswith(b"PK"):
        raise RuntimeError("Google did not return a workbook. Check that the sheet is shared as 'Anyone with the link can view'.")
    return r.content, dt.datetime.now()


def _num(v):
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, dt.time):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, dt.datetime):
        return None
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        try:
            return float(s.replace("%", "")) / 100.0 if "%" in s else float(s)
        except ValueError:
            return None
    return float(v)


def _pct(v):
    x = _num(v)
    if x is None:
        return None
    if 1.5 < x <= 100:
        x = x / 100.0
    return None if x > 1.5 else x


def parse_daily(wb) -> pd.DataFrame:
    """CS_By Date -> one row per (date, channel)."""
    ws = wb["CS_By Date"]
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[2]
    col_date = {ci: v.date() for ci, v in enumerate(hdr) if isinstance(v, dt.datetime)}
    if not col_date:  # dates exported as text such as "1-Jan" (year on row 2)
        m = re.search(r"(\d{4})", str(rows[1][2] or ""))
        year = int(m.group(1)) if m else dt.date.today().year
        for ci, v in enumerate(hdr):
            if isinstance(v, str) and re.match(r"^\d{1,2}-[A-Za-z]{3}$", v.strip()):
                col_date[ci] = dt.datetime.strptime(f"{v.strip()}-{year}", "%d-%b-%Y").date()
    out: dict = OrderedDict()
    section = None
    for r in rows[3:]:
        lab = r[1]
        if not isinstance(lab, str):
            continue
        lab = lab.strip()
        if lab in CHANNELS:
            section = CHANNELS[lab]
            continue
        if section is None:
            continue
        L = lab.lower()
        key = None
        if L.startswith("answer"): key = "answered"
        elif L.startswith("abandoned") and "%" in L: key = "abandon_rate"
        elif L.startswith("abandoned"): key = "abandoned"
        elif L.startswith("short abandoned %"): key = "short_abandon_rate"
        elif L.startswith("short abandoned"): key = "short_abandoned"
        elif L.startswith("out of service"): key = "out_of_service"
        elif L.startswith("avg waiting"): key = "avg_wait_sec"
        elif L.startswith("csat"): key = "csat"
        elif L.startswith("hc"): key = "hc"
        elif L.startswith("complete"): key = "complete"
        elif L.startswith("pending"): key = "pending"
        if key is None:
            continue
        for ci, d in col_date.items():
            if ci >= len(r):
                continue
            v = r[ci]
            if key in ("csat", "abandon_rate", "short_abandon_rate"):
                x = _pct(v)
            else:
                x = _num(v)
                if key == "avg_wait_sec" and isinstance(v, (int, float)) and not isinstance(v, bool):
                    x = float(v) * 86400 if float(v) < 1 else float(v)
            if x is None:
                continue
            k = (d, section)
            out.setdefault(k, {"date": d, "channel": section})
            out[k][key] = x
    df = pd.DataFrame(list(out.values()))
    for c in ["answered", "abandoned", "short_abandoned", "out_of_service", "avg_wait_sec", "csat", "hc", "complete", "pending"]:
        if c not in df.columns:
            df[c] = pd.NA
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["date"] = pd.to_datetime(df["date"])
    # keep only dates that have volume (future dates hold planned HC only)
    df = df[df["answered"].notna() | df["complete"].notna()].sort_values(["date", "channel"]).reset_index(drop=True)
    return df


_MON = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def _parse_period(txt: str, year: int, month: int):
    t = txt.strip().lower().replace("–", "-").replace("—", "-")
    parts = [p.strip() for p in t.split("-")]

    def part(p):
        m = re.match(r"(\d+)\s*([a-z]*)", p)
        return int(m.group(1)), (_MON.get(m.group(2)[:3]) if m.group(2) else None)

    d1, m1 = part(parts[0])
    d2, m2 = part(parts[1]) if len(parts) > 1 else (d1, m1)
    if m1 is None: m1 = m2 if m2 is not None else month
    if m2 is None: m2 = m1
    return dt.date(year, m1, d1), dt.date(year if m2 >= m1 else year + 1, m2, d2)


def parse_campaigns(wb) -> pd.DataFrame:
    """SSP_Campaign Plan -> one row per campaign window (adjacent same-name windows merged)."""
    if "SSP_Campaign Plan" not in wb.sheetnames:
        return pd.DataFrame(columns=["id", "start", "end", "days", "campaign", "kind", "offer", "sub_campaign"])
    ws = wb["SSP_Campaign Plan"]
    rows, year, month = [], None, None
    for r in ws.iter_rows(values_only=True):
        a, b, c = (list(r) + [None, None, None])[:3]
        if isinstance(a, dt.datetime):
            year, month = a.year, a.month
            continue
        if year is None or not isinstance(a, str) or a.strip().lower() == "period" or not isinstance(b, str):
            continue
        try:
            s, e = _parse_period(a, year, month)
        except Exception:
            continue
        lines = [x.strip() for x in b.replace("\r", "").split("\n") if x.strip()]
        name = re.sub(r"superspots", "Supersports", lines[0], flags=re.I).strip()
        kind = "Mega Sale" if "mega sale" in name.lower() else "Brand Day" if "brand day" in name.lower() else "Seasonal"
        sub = " ".join(x.strip() for x in c.replace("\r", "").split("\n") if x.strip()) if isinstance(c, str) else ""
        rows.append(dict(start=s, end=e, campaign=name, kind=kind, offer=" / ".join(lines[1:]), sub_campaign=sub))
    rows.sort(key=lambda x: (x["start"], x["end"]))
    merged = []
    for r in rows:
        if merged and merged[-1]["campaign"].lower() == r["campaign"].lower() and r["start"] <= merged[-1]["end"] + dt.timedelta(days=1):
            merged[-1]["end"] = max(merged[-1]["end"], r["end"])
            merged[-1]["sub_campaign"] = merged[-1]["sub_campaign"] or r["sub_campaign"]
            continue
        merged.append(dict(r))
    for r in merged:
        r["id"] = f"c{r['start'].isoformat()}"
        r["days"] = (r["end"] - r["start"]).days + 1
    df = pd.DataFrame(merged)
    df["start"] = pd.to_datetime(df["start"]); df["end"] = pd.to_datetime(df["end"])
    return df


@st.cache_data(ttl=REFRESH_SECONDS)
def load_data():
    content, fetched_at = load_workbook_bytes()
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    return parse_daily(wb), parse_campaigns(wb), fetched_at


# ----------------------------------------------------------------------------- metrics
def fmt(kind, v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    if kind == "pct": return f"{v * 100:.1f}%"
    if kind == "pct2": return f"{v * 100:.2f}%"
    if kind == "sec": return f"{int(v // 60)}m {int(round(v % 60))}s" if v >= 60 else f"{v:.0f}s"
    if kind == "int": return f"{v:,.0f}"
    return str(v)


def grade(key, v):
    if v is None or pd.isna(v):
        return "na"
    _, exc, mn, higher, _ = STD[key]
    if higher:
        return "ok" if v >= exc else "warn" if v >= mn else "bad"
    return "ok" if v <= exc else "warn" if v <= mn else "bad"


BADGE = {"ok": "🟢 Excellent", "warn": "🟡 Watch", "bad": "🔴 Below standard", "na": "⚪ No data"}


def rate(a, b):
    if a is None or b is None or pd.isna(a) or pd.isna(b) or (a + b) == 0:
        return None
    return a / (a + b)


def seg_metrics(df: pd.DataFrame, seg: str) -> dict:
    r = df[df["channel"] == seg]
    s = lambda c: (None if r[c].isna().all() else float(r[c].sum()))
    m = lambda c: (None if r[c].isna().all() else float(r[c].mean()))
    answered, abandoned = s("answered"), s("abandoned")
    complete, pending = s("complete"), s("pending")
    hc = s("hc")
    return dict(
        answered=answered, abandoned=abandoned, abandon_rate=rate(abandoned, answered),
        short_abandoned=s("short_abandoned"), out_of_service=s("out_of_service"),
        wait=m("avg_wait_sec"), csat=m("csat"), complete=complete, pending=pending,
        total_cases=(complete or 0) + (pending or 0), complete_rate=rate(complete, pending),
        hc_days=hc, per_agent_day=(answered / hc) if hc and answered is not None else None, days=int(r["date"].nunique()),
    )


def totals(df: pd.DataFrame) -> dict:
    seg = {c: seg_metrics(df, c) for c in CHANNELS.values()}
    answered = sum((seg[c]["answered"] or 0) for c in CONTACT)
    abandoned = sum((seg[c]["abandoned"] or 0) for c in CONTACT)
    days = int(df["date"].nunique())
    return dict(seg=seg, answered=answered, abandoned=abandoned, abandon_rate=rate(abandoned, answered),
                answer_rate=(1 - rate(abandoned, answered)) if rate(abandoned, answered) is not None else None,
                wait=seg["call"]["wait"], csat_call=seg["call"]["csat"], csat_chat=seg["line"]["csat"],
                case=seg["case"], complaint=seg["complaint"], days=days, per_day=(answered / days) if days else None)


def gran_for(start, end):
    n = (end - start).days + 1
    return "M" if n > 400 else "W" if n > 95 else "D"


def bucket(df, gran):
    d = df["date"]
    if gran == "D": return d
    if gran == "W": return d.dt.to_period("W-SUN").dt.start_time
    return d.dt.to_period("M").dt.start_time


def campaign_at(camps: pd.DataFrame, d) -> str:
    hit = camps[(camps["start"] <= d) & (camps["end"] >= d)]
    return "" if hit.empty else hit.iloc[0]["campaign"]


def add_campaign_bands(fig, camps, start, end):
    """Shade each campaign window that overlaps [start, end]."""
    for _, c in camps.iterrows():
        if c["end"] < start or c["start"] > end:
            continue
        x0 = max(c["start"], start) - pd.Timedelta(hours=12)
        x1 = min(c["end"], end) + pd.Timedelta(hours=12)
        fig.add_vrect(x0=x0, x1=x1, fillcolor=KIND_COLOR.get(c["kind"], KIND_COLOR["Seasonal"]), line_width=0,
                      annotation_text=c["campaign"], annotation_position="top left", annotation_font_size=10, layer="below")


def base_layout(fig, h=320, ytitle=None, yfmt=None):
    fig.update_layout(height=h, margin=dict(l=10, r=10, t=30, b=10), legend=dict(orientation="h", y=-0.2),
                      hovermode="x unified", plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    fig.update_yaxes(title=ytitle, tickformat=yfmt, gridcolor="rgba(128,128,128,0.2)")
    fig.update_xaxes(showgrid=False)
    return fig


# ----------------------------------------------------------------------------- suggestions
def suggestions(key, seg, m, off=None):
    k = lambda c: fmt("int", m.get(c))
    if key == "abandon_call":
        return [f"{k('abandoned')} calls were lost. Call back every abandoned caller the same day and log the reason.",
                "Check the daily table for days with waiting time above 30 seconds and align shift starts and breaks with the morning peak.",
                "Offer a Line chat deflection message in the IVR queue when the expected wait exceeds 20 seconds.",
                "On days with HC below 5, add a floater from the chat team during 10:00 to 14:00."]
    if key == "wait":
        return ["Target 80% of calls answered within 20 seconds; match the days above 30 seconds with headcount.",
                "Shorten after-call work with case templates for the top reasons (delivery status, how-to, refund follow-up).",
                "Route recorded answers for marketplace and fake-page enquiries through the IVR."]
    if key == "csat":
        extra = ("Chat CSAT comes from very few ratings per day; add a rating prompt at chat close so one low score does not dominate."
                 if seg in ("line", "facebook") else "Pair agents under 95% with the top scorer for side-by-side listening.")
        return ["Pull the days with CSAT under 90% from the daily table and review the low-score tickets in a weekly coaching slot.", extra,
                "Check whether low-CSAT days line up with pending cases over SLA; proactive status updates lift CSAT fastest."]
    if key == "abandon_chat":
        return [f"{k('abandoned')} chats abandoned against {k('answered')} answered; set an auto-greeting with the expected reply time.",
                f"{k('out_of_service')} chats arrived out of service hours; add an out-of-hours auto-reply that collects the order number.",
                "Let the call team pick up Line chat when the call queue is empty, and the reverse on chat peaks.",
                "Use quick replies for the top reasons (order information, delivery status, stock)."]
    if key == "abandon_email":
        return [f"{k('abandoned')} emails went unanswered. Clear the backlog, then set a same-day first-reply target.",
                "Use templates for returns and cancellation, the two largest email reasons, to keep replies under 24 hours."]
    if key == "case_complete":
        return [f"{k('pending')} cases are pending. Review pending cases older than 3 days every morning and escalate delivery and refund cases.",
                "Set an internal SLA of 2 to 3 days to close a case and report ageing weekly.",
                "Reduce inbound volume at the source with proactive shipping and refund notifications."]
    if key == "complaint_resolved":
        return [f"{k('pending')} complaints still open. Assign an owner and a resolution date for each and close the loop within 48 hours.",
                "Group complaints by root cause (branch, out of stock, defect, late delivery) and send a weekly summary to the owners.",
                "Call the customer back after resolution and ask for a satisfaction rating."]
    if key == "abandon_all":
        lift = f" against {off['per_day']:.0f} per day on non-campaign days" if off and off.get("per_day") else ""
        return [f"Campaign window handled {m['per_day']:.0f} contacts per day{lift}; {k('abandoned')} contacts were lost.",
                "Plan staffing for the next campaign of this type from this window: extra agents on launch day and the last two days.",
                "Publish a campaign FAQ (promo code, on-top discount, delivery time, stock) on Line and Facebook auto-replies before launch.",
                "Ask Marketing for the campaign calendar two weeks ahead so staffing can be prepared."]
    return []


# ----------------------------------------------------------------------------- page
st.set_page_config(page_title="SSP_CS Performance Dashboard", layout="wide", page_icon="📞")


@st.fragment(run_every=dt.timedelta(seconds=REFRESH_SECONDS))
def dashboard():
    # ---- load (cached, auto-expires) ----
    try:
        daily, camps, fetched_at = load_data()
        st.session_state["last_good"] = (daily, camps, fetched_at)
        load_error = None
    except Exception as ex:  # keep showing the last good data if the sheet is unreachable
        load_error = str(ex)
        if "last_good" not in st.session_state:
            st.error(f"Could not load the Google Sheet: {load_error}")
            st.stop()
        daily, camps, fetched_at = st.session_state["last_good"]

    data_min, data_max = daily["date"].min().date(), daily["date"].max().date()

    # ---- header ----
    st.title("SSP_CS Performance Dashboard")
    c1, c2 = st.columns([4, 1])
    c1.caption(f"Source: [OFM(CS)_SSP Performance Tracking]({SHEET_URL}) · data {data_min:%-d %b %Y} to {data_max:%-d %b %Y} · "
               f"loaded {fetched_at:%-d %b %Y %H:%M} · re-checks the sheet every {REFRESH_SECONDS // 60} min")
    if c2.button("🔄 Refresh now", width="stretch"):
        st.cache_data.clear()
        st.rerun()
    if load_error:
        st.warning(f"The sheet could not be re-downloaded ({load_error}); showing the last successful load.")

    # ---- filters ----
    with st.sidebar:
        st.header("Time frame")
        preset = st.radio("Preset", ["Month to date", "Last month", "Last 30 days", "Last 90 days", "Quarter to date", "Year to date", "All data", "Custom"], index=0)
        last = data_max
        m0 = last.replace(day=1)
        presets = {
            "Month to date": (m0, last),
            "Last month": ((m0 - dt.timedelta(days=1)).replace(day=1), m0 - dt.timedelta(days=1)),
            "Last 30 days": (last - dt.timedelta(days=29), last),
            "Last 90 days": (last - dt.timedelta(days=89), last),
            "Quarter to date": (dt.date(last.year, 3 * ((last.month - 1) // 3) + 1, 1), last),
            "Year to date": (dt.date(last.year, 1, 1), last),
            "All data": (data_min, data_max),
        }
        camp_opts = camps[(camps["end"].dt.date >= data_min) & (camps["start"].dt.date <= data_max)].sort_values("start", ascending=False)
        camp_pick = st.selectbox("Jump to a campaign period", ["(none)"] + [f"{r.campaign} · {r.start:%-d %b} to {r.end:%-d %b %Y}" for r in camp_opts.itertuples()])
        if camp_pick != "(none)":
            r = camp_opts.iloc[[f"{r.campaign} · {r.start:%-d %b} to {r.end:%-d %b %Y}" for r in camp_opts.itertuples()].index(camp_pick)]
            start, end = max(r["start"].date(), data_min), min(r["end"].date(), data_max)
        elif preset == "Custom":
            rng = st.date_input("From / To", value=(m0, last), min_value=data_min, max_value=data_max)
            start, end = (rng if isinstance(rng, tuple) and len(rng) == 2 else (m0, last))
        else:
            start, end = presets[preset]
        start, end = max(start, data_min), min(end, data_max)
        if start > end:
            start, end = end, start
        st.caption(f"Showing {start:%-d %b %Y} to {end:%-d %b %Y} ({(end - start).days + 1} days)")
        st.header("Segment")
        seg = st.radio("Segment detail", list(SEG_NAME.keys()), format_func=SEG_NAME.get, index=0)

    s_ts, e_ts = pd.Timestamp(start), pd.Timestamp(end)
    n_days = (end - start).days + 1
    cur = daily[(daily["date"] >= s_ts) & (daily["date"] <= e_ts)].copy()
    prev = daily[(daily["date"] >= s_ts - pd.Timedelta(days=n_days)) & (daily["date"] < s_ts)]
    T, P = totals(cur), totals(prev)
    gran = gran_for(start, end)
    gran_word = {"D": "daily", "W": "weekly", "M": "monthly"}[gran]
    in_camp = camps[(camps["start"] <= e_ts) & (camps["end"] >= s_ts)]

    if cur.empty:
        st.info("No data in the selected period.")
        return

    # campaign chips
    if len(in_camp):
        st.markdown("**Campaigns in this period:** " + " · ".join(f"`{r.campaign}` {r.start:%-d %b} to {r.end:%-d %b}" for r in in_camp.itertuples()))
    else:
        st.markdown("**Campaigns in this period:** none planned")

    # ---- standards check ----
    checks = []  # (title, key, value, grade, metrics)
    def chk(title, key, v, m, segid=None):
        g = grade(key, v); checks.append((title, key, v, g, m, segid)); return g
    g_call = chk("Call", "abandon_call", T["seg"]["call"]["abandon_rate"], T["seg"]["call"], "call")
    g_wait = chk("Call", "wait", T["seg"]["call"]["wait"], T["seg"]["call"], "call")
    g_csat = chk("Call", "csat", T["seg"]["call"]["csat"], T["seg"]["call"], "call")
    g_line = chk("Chat (Line)", "abandon_chat", T["seg"]["line"]["abandon_rate"], T["seg"]["line"], "line")
    g_csat_line = chk("Chat (Line)", "csat", T["seg"]["line"]["csat"], T["seg"]["line"], "line")
    g_fb = chk("Chat (Facebook)", "abandon_chat", T["seg"]["facebook"]["abandon_rate"], T["seg"]["facebook"], "facebook")
    g_csat_fb = chk("Chat (Facebook)", "csat", T["seg"]["facebook"]["csat"], T["seg"]["facebook"], "facebook")
    g_email = chk("Email", "abandon_email", T["seg"]["email"]["abandon_rate"], T["seg"]["email"], "email")
    g_case = chk("Case management", "case_complete", T["case"]["complete_rate"], T["case"], "case")
    g_compl = chk("Complaint", "complaint_resolved", T["complaint"]["complete_rate"], T["complaint"], "complaint")

    # campaign windows inside the period
    camp_rows = []
    for r in in_camp.itertuples():
        cs, ce = max(r.start, s_ts), min(r.end, e_ts)
        sub = cur[(cur["date"] >= cs) & (cur["date"] <= ce)]
        if sub.empty:
            continue
        t = totals(sub)
        camp_rows.append(dict(id=r.id, campaign=r.campaign, kind=r.kind, start=cs, end=ce, clipped=(r.start < s_ts or r.end > e_ts), **{k: v for k, v in t.items() if k != "seg"}))
    cur["campaign"] = cur["date"].map(lambda d: campaign_at(camps, d))
    on_df, off_df = cur[cur["campaign"] != ""], cur[cur["campaign"] == ""]
    ON, OFF = totals(on_df) if len(on_df) else None, totals(off_df) if len(off_df) else None
    for cr in camp_rows:
        g = grade("abandon_all", cr["abandon_rate"])
        if g in ("bad", "warn"):
            checks.append((f"Campaign {cr['campaign']} ({cr['start']:%-d %b} to {cr['end']:%-d %b})", "abandon_all", cr["abandon_rate"], g, cr, "campaign"))

    # ---- KPI row ----
    st.subheader("Key metrics, all channels")
    st.caption(f"Selected period, compared with the {n_days} days just before it")
    def delta(cur_v, prev_v, pct=False, invert=False, sec=False):
        if P["days"] == 0:
            return None
        if cur_v is None or prev_v is None or pd.isna(cur_v) or pd.isna(prev_v):
            return None
        d = cur_v - prev_v
        if pct: return f"{d * 100:+.1f} pts"
        if sec: return f"{d:+.0f}s"
        return f"{(d / prev_v * 100):+.0f}%" if prev_v else f"{d:+,.0f}"
    kpis = [
        ("Contacts handled", fmt("int", T["answered"]), delta(T["answered"], P["answered"]), "normal", None, "Calls, chats and emails answered"),
        ("Abandon rate (all)", fmt("pct2", T["abandon_rate"]), delta(T["abandon_rate"], P["abandon_rate"], pct=True), "inverse", None, None),
        ("Call abandon rate", fmt("pct2", T["seg"]["call"]["abandon_rate"]), delta(T["seg"]["call"]["abandon_rate"], P["seg"]["call"]["abandon_rate"], pct=True), "inverse", g_call, None),
        ("Avg call wait", fmt("sec", T["wait"]), delta(T["wait"], P["wait"], sec=True), "inverse", g_wait, None),
        ("CSAT (call)", fmt("pct", T["csat_call"]), delta(T["csat_call"], P["csat_call"], pct=True), "normal", g_csat, f"Chat CSAT {fmt('pct', T['csat_chat'])}"),
        ("Case completion", fmt("pct", T["case"]["complete_rate"]), delta(T["case"]["complete_rate"], P["case"]["complete_rate"], pct=True), "normal", g_case, f"{fmt('int', T['case']['complete'])} complete · {fmt('int', T['case']['pending'])} pending"),
        ("Complaints resolved", fmt("pct", T["complaint"]["complete_rate"]), delta(T["complaint"]["complete_rate"], P["complaint"]["complete_rate"], pct=True), "normal", g_compl, f"{fmt('int', T['complaint']['total_cases'])} logged · {fmt('int', T['complaint']['pending'])} pending"),
        ("Complaints logged", fmt("int", T["complaint"]["total_cases"]), delta(T["complaint"]["total_cases"], P["complaint"]["total_cases"]), "inverse", None, None),
    ]
    for col, (label, value, d, dc, g, help_) in zip(st.columns(len(kpis)), kpis):
        col.metric(label, value, d, delta_color=dc, help=help_)
        if g:
            col.caption(BADGE[g])

    # ---- suggestions ----
    st.subheader("Service standard check and suggestions")
    flagged = [c for c in checks if c[3] in ("bad", "warn")]
    flagged.sort(key=lambda c: 0 if c[3] == "bad" else 1)
    if not flagged:
        st.success("All metrics with data meet the Excellent customer service standard for this period.")
    for title, key, v, g, m, segid in flagged:
        label, exc, mn, higher, f = STD[key]
        with st.expander(f"{BADGE[g]} · {title}: {label} {fmt(f, v)} (standard {'≥' if higher else '≤'} {fmt(f, exc)})", expanded=(g == "bad")):
            for tip in suggestions(key, segid, m, OFF):
                st.markdown(f"- {tip}")
    with st.expander("Customer service standards used on this page"):
        st.table(pd.DataFrame([{"Metric": v[0], "Excellent": ("≥ " if v[3] else "≤ ") + fmt(v[4], v[1]), "Minimum": ("≥ " if v[3] else "≤ ") + fmt(v[4], v[2])} for v in STD.values()]))
        st.caption("Abandon rate = abandoned ÷ (answered + abandoned), weighted over the period. CSAT and waiting time are means of the daily values. "
                   "Case completion = complete ÷ (complete + pending). Complaint resolution = resolved ÷ (resolved + pending).")

    # ---- by segment ----
    st.subheader("By segment")
    seg_cols = st.columns(7)
    cards = [
        ("total", "Total contacts", T["answered"], [("Abandon rate", fmt("pct2", T["abandon_rate"]), None), ("Cases", fmt("int", T["case"]["total_cases"]), None), ("Complaints", fmt("int", T["complaint"]["total_cases"]), None)]),
        ("call", "Call", T["seg"]["call"]["answered"], [("Abandon rate", fmt("pct2", T["seg"]["call"]["abandon_rate"]), g_call), ("Avg wait", fmt("sec", T["seg"]["call"]["wait"]), g_wait), ("CSAT", fmt("pct", T["seg"]["call"]["csat"]), g_csat), ("Per agent-day", fmt("int", T["seg"]["call"]["per_agent_day"]), None)]),
        ("line", "Chat (Line)", T["seg"]["line"]["answered"], [("Abandon rate", fmt("pct2", T["seg"]["line"]["abandon_rate"]), g_line), ("Out of hours", fmt("int", T["seg"]["line"]["out_of_service"]), None), ("CSAT", fmt("pct", T["seg"]["line"]["csat"]), g_csat_line)]),
        ("facebook", "Chat (Facebook)", T["seg"]["facebook"]["answered"], [("Abandon rate", fmt("pct2", T["seg"]["facebook"]["abandon_rate"]), g_fb), ("Out of hours", fmt("int", T["seg"]["facebook"]["out_of_service"]), None), ("CSAT", fmt("pct", T["seg"]["facebook"]["csat"]), g_csat_fb)]),
        ("email", "Email", T["seg"]["email"]["answered"], [("Unanswered", fmt("int", T["seg"]["email"]["abandoned"]), None), ("Unanswered rate", fmt("pct2", T["seg"]["email"]["abandon_rate"]), g_email)]),
        ("case", "Case management", T["case"]["total_cases"], [("Complete", fmt("int", T["case"]["complete"]), None), ("Pending", fmt("int", T["case"]["pending"]), None), ("Completion", fmt("pct", T["case"]["complete_rate"]), g_case)]),
        ("complaint", "Complaint", T["complaint"]["total_cases"], [("Resolved", fmt("int", T["complaint"]["complete"]), None), ("Pending", fmt("int", T["complaint"]["pending"]), None), ("Resolution", fmt("pct", T["complaint"]["complete_rate"]), g_compl)]),
    ]
    for col, (sid, title, big, rows) in zip(seg_cols, cards):
        with col:
            st.markdown(f"**{title}**")
            st.markdown(f"<span style='font-size:1.8rem'>{fmt('int', big)}</span>", unsafe_allow_html=True)
            for lab, val, g in rows:
                st.caption(f"{lab}: **{val}** {BADGE[g].split(' ')[0] if g else ''}")

    # ---- volume charts ----
    st.subheader("Contact volume")
    st.caption(f"Charts are {gran_word}. Shaded bands are campaign periods (red Mega Sale, blue Brand Day, green Seasonal).")
    cur["bucket"] = bucket(cur, gran)
    vol = cur[cur["channel"].isin(CONTACT)].pivot_table(index="bucket", columns="channel", values="answered", aggfunc="sum").fillna(0)
    fig = go.Figure()
    for c in CONTACT:
        if c in vol.columns:
            fig.add_bar(x=vol.index, y=vol[c], name=SEG_NAME[c], marker_color=COLOR[c])
    fig.update_layout(barmode="stack", title="Contacts handled by channel")
    add_campaign_bands(fig, camps, s_ts, e_ts)
    st.plotly_chart(base_layout(fig, 340, "contacts"), width="stretch")

    v1, v2 = st.columns(2)
    share = pd.Series({SEG_NAME[c]: (T["seg"][c]["answered"] or 0) for c in CONTACT}).sort_values()
    f2 = go.Figure(go.Bar(x=share.values, y=share.index, orientation="h", marker_color=[COLOR[k] for k in [c for c in CONTACT for n in [SEG_NAME[c]] if n in share.index]][: len(share)],
                          text=[f"{v:,.0f} · {v / max(T['answered'], 1):.1%}" for v in share.values], textposition="outside"))
    f2.update_layout(title="Share of contacts by channel")
    v1.plotly_chart(base_layout(f2, 280), width="stretch")
    cs = cur[cur["channel"] == "case"].groupby("bucket")[["complete", "pending"]].sum()
    f3 = go.Figure([go.Bar(x=cs.index, y=cs["complete"], name="Complete", marker_color=COLOR["case"]), go.Bar(x=cs.index, y=cs["pending"], name="Pending", marker_color=COLOR["complaint"])])
    f3.update_layout(barmode="stack", title="Case management: complete vs pending")
    add_campaign_bands(f3, camps, s_ts, e_ts)
    v2.plotly_chart(base_layout(f3, 280, "cases"), width="stretch")

    # ---- campaign periods ----
    st.subheader("Campaign periods")
    if not camp_rows:
        st.info(f"No campaign is planned between {start:%-d %b} and {end:%-d %b %Y}; every day counts as a non-campaign day. "
                "The plan tab is filled for Jan to Mar and Oct to Nov 2026 only.")
    else:
        if ON and OFF and OFF["per_day"]:
            lift = ON["per_day"] / OFF["per_day"] - 1
            st.markdown(f"Campaign days averaged **{ON['per_day']:.0f}** contacts per day against **{OFF['per_day']:.0f}** on non-campaign days "
                        f"(**{lift:+.0%}**); response rate **{fmt('pct', ON['answer_rate'])}** vs **{fmt('pct', OFF['answer_rate'])}**; "
                        f"call CSAT **{fmt('pct', ON['csat_call'])}** vs **{fmt('pct', OFF['csat_call'])}**.")
        elif ON:
            st.markdown("Every day in this period falls inside a campaign; widen the period to compare with a non-campaign baseline.")
        tbl = pd.DataFrame(camp_rows)
        if OFF:
            tbl = pd.concat([tbl, pd.DataFrame([dict(campaign="No campaign days", kind="—", start=s_ts, end=e_ts, clipped=False, **{k: v for k, v in OFF.items() if k != "seg"})])], ignore_index=True)
        if ON and len(camp_rows) > 1:
            tbl = pd.concat([tbl, pd.DataFrame([dict(campaign="All campaign days", kind="—", start=s_ts, end=e_ts, clipped=False, **{k: v for k, v in ON.items() if k != "seg"})])], ignore_index=True)
        tbl["label"] = tbl.apply(lambda r: r["campaign"] + (f" ({r['start']:%-d %b})" if (tbl["campaign"] == r["campaign"]).sum() > 1 and r["kind"] != "—" else "") + (" *" if r["clipped"] else ""), axis=1)
        c1, c2 = st.columns(2)
        f4 = go.Figure(go.Bar(x=tbl["per_day"], y=tbl["label"], orientation="h", marker_color=[{"Mega Sale": "#E45756", "Brand Day": "#4C78A8", "Seasonal": "#54A24B"}.get(k, "#999") for k in tbl["kind"]],
                              text=[f"{v:.0f} · {d} d" for v, d in zip(tbl["per_day"], tbl["days"])], textposition="outside"))
        f4.update_layout(title="Contacts handled per day, by campaign")
        c1.plotly_chart(base_layout(f4, max(240, 32 * len(tbl))), width="stretch")
        f5 = go.Figure(go.Bar(x=tbl["answer_rate"], y=tbl["label"], orientation="h", marker_color=[{"Mega Sale": "#E45756", "Brand Day": "#4C78A8", "Seasonal": "#54A24B"}.get(k, "#999") for k in tbl["kind"]],
                              text=[f"{v:.1%} · {a:,.0f} lost" for v, a in zip(tbl["answer_rate"].fillna(0), tbl["abandoned"])], textposition="outside"))
        f5.add_vline(x=0.97, line_dash="dash", line_color="gray", annotation_text="Excellent 97%")
        f5.update_layout(title="Response rate by campaign (answered ÷ offered)")
        c2.plotly_chart(base_layout(f5, max(240, 32 * len(tbl)), yfmt=".0%"), width="stretch")
        f5.update_xaxes(tickformat=".0%", range=[0, 1.05])
        show = pd.DataFrame({
            "Campaign": tbl["label"], "Type": tbl["kind"],
            "Window": [("selected period" if k == "—" else f"{s:%-d %b} to {e:%-d %b %Y}") for k, s, e in zip(tbl["kind"], tbl["start"], tbl["end"])],
            "Days": tbl["days"], "Contacts": tbl["answered"], "Per day": tbl["per_day"].round(0),
            "Abandon %": tbl["abandon_rate"], "Response %": tbl["answer_rate"], "Avg call wait (s)": tbl["wait"].round(0),
            "Call CSAT": tbl["csat_call"], "Cases": [c["total_cases"] for c in tbl["case"]], "Complaints": [c["total_cases"] for c in tbl["complaint"]],
            "Complaints per 1,000": [(c["total_cases"] / a * 1000) if a else None for c, a in zip(tbl["complaint"], tbl["answered"])],
        })
        st.dataframe(show.style.format({"Abandon %": "{:.2%}", "Response %": "{:.1%}", "Call CSAT": "{:.1%}", "Contacts": "{:,.0f}", "Per day": "{:,.0f}", "Complaints per 1,000": "{:.2f}"}, na_rep="—"),
                     width="stretch", hide_index=True)
        if tbl["clipped"].any():
            st.caption("* window extends beyond the selected period; figures cover the days inside it.")

    # ---- quality trend ----
    st.subheader("Service quality trend")
    q1, q2 = st.columns(2)
    f6 = go.Figure()
    for c in CONTACT:
        g = cur[cur["channel"] == c].groupby("bucket")[["answered", "abandoned"]].sum()
        g = g[(g["answered"] + g["abandoned"]) > 0]
        f6.add_scatter(x=g.index, y=g["abandoned"] / (g["answered"] + g["abandoned"]), name=SEG_NAME[c], mode="lines", line=dict(color=COLOR[c], width=2))
    f6.add_hline(y=0.03, line_dash="dash", line_color="gray", annotation_text="Standard 3%")
    f6.update_layout(title="Abandon rate by channel")
    add_campaign_bands(f6, camps, s_ts, e_ts)
    q1.plotly_chart(base_layout(f6, 300, yfmt=".1%"), width="stretch")
    f7 = go.Figure()
    for c, nm in (("call", "CSAT call"), ("line", "CSAT chat")):
        g = cur[cur["channel"] == c].groupby("bucket")["csat"].mean().dropna()
        f7.add_scatter(x=g.index, y=g, name=nm, mode="lines+markers", line=dict(color=COLOR[c], width=2))
    w = cur[cur["channel"] == "call"].groupby("bucket")["avg_wait_sec"].mean().dropna()
    f7.add_scatter(x=w.index, y=w, name="Avg call wait (s, right)", mode="lines", line=dict(color="gray", dash="dot"), yaxis="y2")
    f7.add_hline(y=0.95, line_dash="dash", line_color="gray", annotation_text="Standard 95%")
    f7.update_layout(title="CSAT and call waiting time", yaxis2=dict(overlaying="y", side="right", showgrid=False, title="seconds"))
    add_campaign_bands(f7, camps, s_ts, e_ts)
    q2.plotly_chart(base_layout(f7, 300, yfmt=".0%"), width="stretch")

    # ---- segment detail ----
    st.subheader(f"Segment detail: {SEG_NAME[seg]}")
    if seg == "total":
        d = cur[cur["channel"].isin(CONTACT)].pivot_table(index="date", columns="channel", values="answered", aggfunc="sum").fillna(0)
        d["Handled"] = d.sum(axis=1)
        d["Abandoned"] = cur[cur["channel"].isin(CONTACT)].groupby("date")["abandoned"].sum()
        d["Abandon %"] = d["Abandoned"] / (d["Handled"] + d["Abandoned"])
        d["Cases"] = cur[cur["channel"] == "case"].groupby("date")[["complete", "pending"]].sum().sum(axis=1)
        d["Complaints"] = cur[cur["channel"] == "complaint"].groupby("date")[["complete", "pending"]].sum().sum(axis=1)
        d.insert(0, "Campaign", [campaign_at(camps, x) for x in d.index])
        d = d.rename(columns=SEG_NAME)
        st.dataframe(d.style.format({"Abandon %": "{:.2%}"}, precision=0, na_rep="—"), width="stretch")
    else:
        r = cur[cur["channel"] == seg].set_index("date").sort_index()
        m = T["seg"][seg]
        c = st.columns(5)
        if seg in ("case", "complaint"):
            key = "case_complete" if seg == "case" else "complaint_resolved"
            c[0].metric("Logged", fmt("int", m["total_cases"])); c[1].metric("Complete" if seg == "case" else "Resolved", fmt("int", m["complete"]))
            c[2].metric("Pending", fmt("int", m["pending"])); c[3].metric(f"Rate · {BADGE[grade(key, m['complete_rate'])]}", fmt("pct", m["complete_rate"]))
            c[4].metric("Per day", fmt("int", m["total_cases"] / max(1, m["days"])))
            fig = go.Figure([go.Bar(x=r.index, y=r["complete"], name="Complete", marker_color=COLOR[seg]), go.Bar(x=r.index, y=r["pending"], name="Pending", marker_color="#E45756")])
            fig.update_layout(barmode="stack")
            cols = ["complete", "pending"]
            r["rate"] = r["complete"] / (r["complete"] + r["pending"])
            cols.append("rate")
        else:
            akey = {"call": "abandon_call", "email": "abandon_email"}.get(seg, "abandon_chat")
            c[0].metric("Answered", fmt("int", m["answered"])); c[1].metric("Abandoned" if seg != "email" else "Unanswered", fmt("int", m["abandoned"]))
            c[2].metric(f"Abandon rate · {BADGE[grade(akey, m['abandon_rate'])]}", fmt("pct2", m["abandon_rate"]))
            if seg == "call":
                c[3].metric(f"Avg wait · {BADGE[grade('wait', m['wait'])]}", fmt("sec", m["wait"]))
            else:
                c[3].metric("Out of service hours" if seg != "email" else "Avg HC", fmt("int", m["out_of_service"] if seg != "email" else (m["hc_days"] / max(1, m["days"]) if m["hc_days"] else None)))
            c[4].metric(f"CSAT · {BADGE[grade('csat', m['csat'])]}" if seg != "email" else "Per agent-day", fmt("pct", m["csat"]) if seg != "email" else fmt("int", m["per_agent_day"]))
            fig = go.Figure([go.Bar(x=r.index, y=r["answered"], name="Answered", marker_color=COLOR[seg]), go.Bar(x=r.index, y=r["abandoned"], name="Abandoned", marker_color="#E45756")])
            fig.update_layout(barmode="stack")
            cols = [x for x in ["answered", "abandoned", "short_abandoned", "avg_wait_sec", "out_of_service", "csat", "hc"] if x in r.columns and r[x].notna().any()]
            r["abandon %"] = r["abandoned"] / (r["answered"] + r["abandoned"])
            cols.insert(2, "abandon %")
        add_campaign_bands(fig, camps, s_ts, e_ts)
        st.plotly_chart(base_layout(fig, 280), width="stretch")
        tbl = r[cols].copy()
        tbl.insert(0, "campaign", [campaign_at(camps, x) for x in tbl.index])
        st.dataframe(tbl.style.format({"abandon %": "{:.2%}", "csat": "{:.1%}", "rate": "{:.1%}", "avg_wait_sec": "{:.0f}"}, precision=0, na_rep="—"), width="stretch")


dashboard()
