"""SO Productivity Dashboard - ACI Premio Plastics
Run locally:  streamlit run dashboard.py
Data: put the latest Excel (3 sheets: Summary Report / SO Activity / Attendance) in the data/ folder.
"""
import base64, glob, os
import numpy as np, pandas as pd
import plotly.express as px, plotly.graph_objects as go
import streamlit as st

# ================= EASY SETTINGS =================
COMPANY = "ACI Premio Plastics"
TITLE = "SO Productivity Dashboard"
DAILY_OUTLET, DAILY_MEMO, SHIFT_H = 20, 8, 11      # daily targets, working hours
OFFICE_START, LATE_AT = 8 * 60 + 30, 8 * 60 + 45   # office 8:30, "late" = 8:45 or after
SCORE_W = (35, 30, 20, 15)                         # score weights: visit, memo, attendance, punctuality
BASE = os.path.dirname(os.path.abspath(__file__))
PAL = ["#6C5DD3", "#FF754C", "#3F8CFF", "#00C48C", "#FFB800", "#9B8AFB"]
RC = {"Excellent": "#00C48C", "Good": "#3F8CFF", "Needs Improvement": "#FFB800", "Critical": "#FF4D4F"}

st.set_page_config(page_title=TITLE, page_icon="📊", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""<style>
#MainMenu,footer,[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stAppDeployButton"],.stDeployButton{display:none!important}
header[data-testid="stHeader"]{background:transparent;height:0}
.block-container{padding:.6rem 1rem 3rem;max-width:1400px}
.hero{display:flex;align-items:center;gap:14px;padding:16px 20px;border-radius:20px;color:#fff;margin-bottom:12px;
 background:linear-gradient(135deg,#0f1729 0%,#3b2f8f 55%,#6C5DD3 100%);box-shadow:0 14px 30px rgba(60,45,140,.35),0 3px 0 #2a2170}
.hero img{width:58px;height:58px;border-radius:50%;background:#fff;padding:3px;box-shadow:0 4px 12px rgba(0,0,0,.4)}
.hero h1{margin:0;font-size:1.55rem;line-height:1.2;color:#fff}.hero p{margin:2px 0 0;opacity:.8;font-size:.85rem}
.kg{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin:8px 0 14px}
.kc{border-radius:16px;padding:14px 16px;color:#fff;transition:.25s;box-shadow:0 10px 18px rgba(30,30,80,.18),0 4px 0 rgba(0,0,0,.18);transform:perspective(600px) rotateX(2deg)}
.kc:hover{transform:perspective(600px) rotateX(0) translateY(-4px);box-shadow:0 16px 26px rgba(30,30,80,.28),0 4px 0 rgba(0,0,0,.18)}
.kc b{display:block;font-size:1.6rem;line-height:1.15}.kc span{font-size:.78rem;opacity:.9}.kc i{display:block;font-style:normal;font-size:.72rem;opacity:.8;margin-top:2px}
.ai{background:#fff;border-left:6px solid #6C5DD3;border-radius:14px;padding:12px 18px;margin:8px 0;box-shadow:0 6px 16px rgba(0,0,0,.08)}
div[data-testid="stPlotlyChart"]{background:#fff;border-radius:16px;padding:6px;box-shadow:0 8px 20px rgba(30,30,80,.10)}
.stTabs [data-baseweb="tab-list"]{overflow-x:auto;gap:4px}.stTabs [data-baseweb="tab"]{border-radius:10px 10px 0 0;padding:8px 12px;white-space:nowrap}
@media(max-width:640px){.hero h1{font-size:1.15rem}.kc b{font-size:1.3rem}.block-container{padding:.4rem .5rem 3rem}}
</style>""", unsafe_allow_html=True)


# ================= DATA =================
def fix(x):  # repair Bengali text that was saved with a wrong encoding
    if not isinstance(x, str) or ("à" not in x and "Ã" not in x):
        return x
    try:
        return b"".join(c.encode("cp1252") if c.encode("cp1252", "ignore") else c.encode("latin-1") for c in x).decode("utf8")
    except Exception:
        return x


def data_file():
    fs = [f for f in glob.glob(os.path.join(BASE, "data", "*.xls*")) if not os.path.basename(f).startswith("~$")]
    return max(fs, key=os.path.getmtime) if fs else None


@st.cache_data(show_spinner="Loading data...")
def load(path, stamp):
    xl = pd.ExcelFile(path)
    sh = {n.strip().lower(): n for n in xl.sheet_names}
    S, A, T = (xl.parse(sh[k]) for k in ("summary report", "so activity", "attendance"))
    for d in (S, A, T):
        d.columns = [str(c).strip() for c in d.columns]
    M = S.rename(columns={"Staff ID": "sid", "Sr Name": "so", "RSM Name": "rsm", "SM Name": "sm", "Group": "grp"})[["sid", "so", "rsm", "sm", "grp"]]
    M = M.astype(str).apply(lambda c: c.str.strip()).drop_duplicates("sid")
    cm = A.dropna(subset=["Staff ID"]).drop_duplicates("Sr Code").set_index("Sr Code")["Staff ID"].astype(str)
    A["sid"] = A["Staff ID"].astype(str).where(A["Staff ID"].notna(), A["Sr Code"].map(cm).fillna(A["Sr Code"]))
    mp = {k: fix(k) for k in A["Market Name"].dropna().unique()}
    A = A.rename(columns={"Number Of Outlet Visit": "outlets", "Number Of Ordered Shop": "memos", "Visit Date": "date", "Market Name": "market",
                          "Upazilla Name": "upa", "District Name": "dist", "Route Name": "route", "Route Day Name": "rday", "Day": "day"})
    A["market"] = A["market"].map(mp)
    A["off"] = A["rday"].astype(str).str.strip() != A["day"].astype(str).str.strip()
    A = A[A.sid.isin(M.sid)]
    T["sid"] = T["Employee Code"].map(cm).fillna(T["Staff Id"].astype(str))
    T = T[T.sid.isin(M.sid)].rename(columns={"Attendance Date": "date"}).drop_duplicates(["sid", "date"])
    T["mins"] = T["In Time Only"].map(lambda v: v.hour * 60 + v.minute + v.second / 60 if hasattr(v, "hour") else np.nan)
    s = T["Attendance Status"].astype(str).str.strip()
    T["cat"] = np.select([s.eq("Absent"), s.str.startswith("On Leave"), T.mins < LATE_AT], ["Absent", "Leave", "Before 8:45"], "At/After 8:45")
    pres = T.cat.isin(["Before 8:45", "At/After 8:45"])
    T["eff"] = np.where(pres, np.clip((OFFICE_START + SHIFT_H * 60 - np.maximum(T.mins.fillna(OFFICE_START), OFFICE_START)) / 60, 0, SHIFT_H), 0)
    for d in (A, T):
        d["month"] = d["date"].dt.strftime("%b %Y")
        d["ym"] = d["date"].dt.to_period("M").astype(str)
    return M, A, T


def so_table(M, A, T):
    days = max(T.date.nunique(), 1)
    g = A.groupby("sid").agg(visits=("outlets", "sum"), memos=("memos", "sum"))
    c = pd.crosstab(T.sid, T.cat).reindex(columns=["Before 8:45", "At/After 8:45", "Absent", "Leave"], fill_value=0)
    c.columns = ["early", "late", "absent", "leave"]
    D = M.set_index("sid").join(g).join(c).join(T.groupby("sid").eff.sum().rename("eff_h")).fillna(0).reset_index()
    D["days"] = days
    D["present"] = D.early + D.late
    D["present_pct"] = (D.present + D.leave).clip(upper=days) / days * 100  # approved leave counts as present, like the Summary sheet
    D["punct_pct"] = np.where(D.present > 0, D.early / D.present.replace(0, 1) * 100, 0)
    D["visit_t"], D["memo_t"] = days * DAILY_OUTLET, days * DAILY_MEMO
    D["visit_pct"], D["memo_pct"] = D.visits / D.visit_t * 100, D.memos / D.memo_t * 100
    D["std_h"] = D.present * SHIFT_H
    D["v_day"] = D.visits / D.present.replace(0, np.nan)
    D["m_day"] = D.memos / D.present.replace(0, np.nan)
    D["conv"] = D.memos / D.visits.replace(0, np.nan) * 100
    w = SCORE_W
    D["score"] = (w[0] * D.visit_pct.clip(upper=100) + w[1] * D.memo_pct.clip(upper=100) + w[2] * D.present_pct.clip(upper=100) + w[3] * D.punct_pct.clip(upper=100)) / 100
    D["rating"] = pd.cut(D.score, [-1, 50, 70, 85, 101], labels=["Critical", "Needs Improvement", "Good", "Excellent"], right=False).astype(str)
    D["punct_cls"] = np.select([D.early >= days, D.punct_pct >= 80, D.punct_pct >= 50], ["All days before 8:45", "Mostly on time (80%+)", "Mixed (50-80%)"], "Mostly late (<50%)")
    return D


def grp_table(D, by):
    D = D.assign(f_p=D.early >= D.days, has_abs=D.absent > 0, m_late=(D.punct_pct < 50) & (D.present > 0), low=D.rating.isin(["Needs Improvement", "Critical"]))
    R = D.groupby(by).agg(SOs=("sid", "count"), Present=("present", "sum"), Absent=("absent", "sum"), Leave=("leave", "sum"), Before845=("early", "sum"),
                          After845=("late", "sum"), Visits=("visits", "sum"), VisitT=("visit_t", "sum"), Memos=("memos", "sum"), MemoT=("memo_t", "sum"),
                          Score=("score", "mean"), All_Before_845_SOs=("f_p", "sum"), SOs_With_Absent=("has_abs", "sum"), Mostly_Late_SOs=("m_late", "sum"),
                          Need_Improve_SOs=("low", "sum"), days=("days", "max")).reset_index()
    tot = R.SOs * R.days
    R["Present %"], R["Absent %"] = (R.Present + R.Leave) / tot * 100, R.Absent / tot * 100
    R["On-time %"] = R.Before845 / R.Present.replace(0, np.nan) * 100
    R["Visit %"], R["Memo %"] = R.Visits / R.VisitT * 100, R.Memos / R.MemoT * 100
    return R.drop(columns="days")


def daily(A, T):
    a = A.groupby("date").agg(Outlets=("outlets", "sum"), Memos=("memos", "sum"), Active=("sid", "nunique"))
    t = T.groupby("date").agg(Present=("cat", lambda s: s.isin(["Before 8:45", "At/After 8:45"]).sum()), EffH=("eff", "sum"))
    d = a.join(t, how="outer").fillna(0)
    d["Present"] = np.where(d.Present > 0, d.Present, d.Active)
    d["OutletT"], d["MemoT"], d["StdH"] = d.Present * DAILY_OUTLET, d.Present * DAILY_MEMO, d.Present * SHIFT_H
    d["Outlet %"], d["Memo %"] = d.Outlets / d.OutletT * 100, d.Memos / d.MemoT * 100
    d["month"] = pd.Series(d.index.strftime("%b %Y"), index=d.index)
    return d.reset_index()


# ================= UI HELPERS =================
def show(fig, h=380):
    fig.update_layout(height=h, margin=dict(l=8, r=8, t=44, b=8), template="plotly_white", colorway=PAL, legend=dict(orientation="h", y=-0.15),
                      font=dict(size=12), title_font_size=15)
    try:
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
    except Exception:
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def table(df, pct=(), cap=100, h=None):
    cfg = {c: st.column_config.ProgressColumn(c, min_value=0, max_value=max(cap, float(df[c].max() or 0)), format="%.1f%%") for c in pct if c in df}
    for c in df.select_dtypes("float").columns:
        if c not in cfg:
            cfg[c] = st.column_config.NumberColumn(c, format="%.1f")
    kw = {"height": h} if h else {}
    try:
        st.dataframe(df, width="stretch", hide_index=True, column_config=cfg, **kw)
    except Exception:
        st.dataframe(df, use_container_width=True, hide_index=True, column_config=cfg, **kw)


def kpis(items):
    cols = ["#6C5DD3,#9B8AFB", "#FF754C,#FFA27E", "#3F8CFF,#7BB2FF", "#00A67E,#3FD9A8", "#E09B00,#FFC93C", "#D4467D,#FF8DB0"]
    h = "".join(f'<div class="kc" style="background:linear-gradient(135deg,{cols[i % 6].split(",")[0]},{cols[i % 6].split(",")[1]})"><span>{a}</span><b>{b}</b><i>{c}</i></div>'
                for i, (a, b, c) in enumerate(items))
    st.markdown(f'<div class="kg">{h}</div>', unsafe_allow_html=True)


def bars3d(labels, series, title):
    """True 3D column chart: one row of columns per category, one colour per metric."""
    I, J, K = [7, 0, 0, 0, 4, 4, 6, 6, 4, 0, 3, 2], [3, 4, 1, 2, 5, 6, 5, 2, 0, 1, 6, 3], [0, 7, 2, 3, 6, 7, 1, 1, 5, 5, 7, 6]
    fig = go.Figure()
    for j, (name, vals) in enumerate(series.items()):
        x, y, z, i, jj, k = [], [], [], [], [], []
        for n, v in enumerate(vals):
            o = len(x)
            x += [n + dx * .7 for dx in (0, 1, 1, 0, 0, 1, 1, 0)]
            y += [j + dy * .7 for dy in (0, 0, 1, 1, 0, 0, 1, 1)]
            z += [0, 0, 0, 0] + [float(v)] * 4
            i += [a + o for a in I]; jj += [a + o for a in J]; k += [a + o for a in K]
        fig.add_trace(go.Mesh3d(x=x, y=y, z=z, i=i, j=jj, k=k, color=PAL[j % 6], opacity=.92, name=name, showlegend=True, flatshading=True,
                                hovertemplate=name + "<extra></extra>"))
    fig.update_layout(title=title, scene=dict(xaxis=dict(tickvals=[n + .35 for n in range(len(labels))], ticktext=[l[:14] for l in labels], title=""),
                                              yaxis=dict(tickvals=[j + .35 for j in range(len(series))], ticktext=list(series), title=""),
                                              zaxis=dict(title="%"), camera=dict(eye=dict(x=1.6, y=-1.6, z=.9))))
    show(fig, 480)


def level_tabs(D, fn):
    for tab, (nm, by) in zip(st.tabs(["🏢 SM wise", "👤 RSM wise", "🏷️ Group wise"]), [("SM", "sm"), ("RSM", "rsm"), ("Group", "grp")]):
        with tab:
            fn(grp_table(D, by), by, nm)


# ================= LOAD + FILTERS =================
path = data_file()
if not path:
    st.error("No Excel file found in the data/ folder."); st.stop()
M0, A0, T0 = load(path, os.path.getmtime(path))
badge = os.path.join(BASE, "assets", "badge.png")
img = f'<img src="data:image/png;base64,{base64.b64encode(open(badge, "rb").read()).decode()}">' if os.path.exists(badge) else ""
st.markdown(f'<div class="hero">{img}<div><h1>{TITLE}</h1><p>{COMPANY} · Attendance · Outlet Visit · Memo · Productivity</p></div></div>', unsafe_allow_html=True)

with st.expander("🔎 Filters (tap to change)"):
    f = st.columns(4)
    months = list(T0.drop_duplicates("ym").sort_values("ym").month)
    mo = f[0].selectbox("Month", ["All"] + months, index=len(months) if len(months) == 1 else 0)
    sg = f[1].multiselect("Group", sorted(M0.grp.unique()))
    ss = f[2].multiselect("SM", sorted(M0.sm.unique()))
    sr = f[3].multiselect("RSM", sorted(M0.rsm[M0.sm.isin(ss)].unique() if ss else M0.rsm.unique()))
mask = pd.Series(True, index=M0.index)
for col_, sel_ in (("grp", sg), ("sm", ss), ("rsm", sr)):
    if sel_:
        mask &= M0[col_].isin(sel_)
M = M0[mask]
A, T = A0[A0.sid.isin(M.sid)], T0[T0.sid.isin(M.sid)]
if mo != "All":
    A, T = A[A.month == mo], T[T.month == mo]
if M.empty or T.empty:
    st.warning("No data for this selection."); st.stop()
D = so_table(M, A, T)
days = int(D.days.iloc[0])
st.caption(f"📅 {T.date.min():%d %b} – {max(T.date.max(), A.date.max()):%d %b %Y} · {days} attendance days · {len(D)} SOs · Targets: {DAILY_OUTLET} outlets + {DAILY_MEMO} memos per day · Late = 8:45 AM or after · Approved leave counts as present")

tabs = st.tabs(["🏠 Overview", "🕗 Attendance", "🛍️ Outlet Visit", "🧾 Memo", "🏆 Top & Bottom 10", "📍 Market Hotspots", "⭐ Group & RSM Scorecard", "🤖 AI Summary", "📋 SO Data"])

# ---------- Overview ----------
with tabs[0]:
    P, tp = D.present.sum(), len(D) * days
    kpis([("Sales Officers", f"{len(D)}", f"{days} working days"), ("Attendance", f"{(P + D.leave.sum()) / tp * 100:.1f}%", f"{int(D.absent.sum())} absent · {int(D.leave.sum())} leave days"),
          ("Before 8:45", f"{D.early.sum() / max(P, 1) * 100:.1f}%", f"{int(D.late.sum())} late days"),
          ("Outlet Visits", f"{int(D.visits.sum()):,}", f"{D.visits.sum() / D.visit_t.sum() * 100:.1f}% of {int(D.visit_t.sum()):,}"),
          ("Memos", f"{int(D.memos.sum()):,}", f"{D.memos.sum() / D.memo_t.sum() * 100:.1f}% of {int(D.memo_t.sum()):,}"),
          ("Work Hours", f"{int(D.std_h.sum()):,}", f"{SHIFT_H}h x present days"),
          ("Visit / present day", f"{D.visits.sum() / max(P, 1):.1f}", f"target {DAILY_OUTLET}"), ("Memo / present day", f"{D.memos.sum() / max(P, 1):.1f}", f"target {DAILY_MEMO}")])
    c1, c2 = st.columns([3, 2])
    with c1:
        show(px.scatter_3d(D, x="visit_pct", y="memo_pct", z="present_pct", color="grp", hover_name="so", hover_data=["rsm", "rating"],
                           labels={"visit_pct": "Visit %", "memo_pct": "Memo %", "present_pct": "Present %"}, title="3D view: every SO (Visit % x Memo % x Present %)"), 470)
    with c2:
        r = D.rating.value_counts().reindex(list(RC)).fillna(0).reset_index()
        r.columns = ["Rating", "SOs"]
        show(px.pie(r, names="Rating", values="SOs", hole=.55, color="Rating", color_discrete_map=RC, title="SO performance rating"), 470)
    d = daily(A, T)
    for col, tg, nm in (("Outlets", "OutletT", "Outlet visits"), ("Memos", "MemoT", "Memos")):
        fg = go.Figure([go.Bar(x=d.date, y=d[col], name=nm, marker_color=PAL[0] if col == "Outlets" else PAL[1]),
                        go.Scatter(x=d.date, y=d[tg], name="Target (present SOs)", mode="lines+markers", line=dict(color="#222", dash="dot"))])
        fg.update_layout(title=f"Daily {nm} vs target")
        show(fg, 330)
    st.subheader("Daily & Monthly summary (outlet · memo · working hour)")
    cols = ["Present", "Outlets", "OutletT", "Outlet %", "Memos", "MemoT", "Memo %", "StdH", "EffH"]
    dd = d.assign(Date=d.date.dt.strftime("%d %b %a"))[["Date"] + cols].rename(columns={"OutletT": "Outlet Target", "MemoT": "Memo Target", "StdH": "Work Hours (11h)", "EffH": "Effective Hours*"})
    mm = d.groupby("month", sort=False)[["Present", "Outlets", "OutletT", "Memos", "MemoT", "StdH", "EffH"]].sum().reset_index()
    mm["Outlet %"], mm["Memo %"] = mm.Outlets / mm.OutletT * 100, mm.Memos / mm.MemoT * 100
    mm = mm.rename(columns={"month": "Month", "OutletT": "Outlet Target", "MemoT": "Memo Target", "StdH": "Work Hours (11h)", "EffH": "Effective Hours*"})
    table(mm, ["Outlet %", "Memo %"], 100)
    table(dd, ["Outlet %", "Memo %"], 100, 360)
    st.caption("Present = SO-days present. Work Hours = present SO-days x 11h. *Effective Hours = estimate from 8:30 office start (11h shift) minus late arrival, because the file has no check-out time.")

# ---------- Attendance ----------
with tabs[1]:
    st.subheader("Month-wise attendance (8:45 AM rule)")
    mt = T0[T0.sid.isin(M.sid)].groupby(["ym", "month"]).agg(Records=("cat", "size"), Before=("cat", lambda s: (s == "Before 8:45").sum()), After=("cat", lambda s: (s == "At/After 8:45").sum()),
                                                            Absent=("cat", lambda s: (s == "Absent").sum()), Leave=("cat", lambda s: (s == "Leave").sum()),
                                                            SOs_Absent=("sid", lambda s: s[T0.loc[s.index, "cat"] == "Absent"].nunique())).reset_index()
    for a, b in (("Before", "Before 8:45 %"), ("After", "At/After 8:45 %"), ("Absent", "Absent %"), ("Leave", "Leave %")):
        mt[b] = mt[a] / mt.Records * 100
    mt["SOs with absent %"] = mt.SOs_Absent / len(M) * 100
    table(mt.drop(columns="ym").rename(columns={"month": "Month", "Before": "Before 8:45", "After": "At/After 8:45", "SOs_Absent": "SOs with absent"}), ["Before 8:45 %", "At/After 8:45 %", "Absent %", "Leave %", "SOs with absent %"])
    c1, c2 = st.columns(2)
    with c1:
        show(px.bar(mt, x="month", y=["Before", "After", "Absent", "Leave"], barmode="stack", title="Attendance records by month", text_auto=True,
                    color_discrete_sequence=["#00C48C", "#FFB800", "#FF4D4F", "#9B8AFB"]))
    with c2:
        pc = D.punct_cls.value_counts().reset_index(); pc.columns = ["Class", "SOs"]
        show(px.pie(pc, names="Class", values="SOs", hole=.5, title="How SOs arrive (SO count)", color_discrete_sequence=["#00C48C", "#3F8CFF", "#FFB800", "#FF4D4F"]))

    def att(R, by, nm):
        show(px.bar(R, y=by, x=["All_Before_845_SOs", "SOs_With_Absent", "Mostly_Late_SOs"], barmode="group", orientation="h", title=f"{nm}-wise: SO count",
                    labels={"value": "SOs", "variable": ""}), max(320, 60 * len(R)))
        out = R[[by, "SOs", "Present", "Present %", "Absent", "Absent %", "Before845", "After845", "On-time %", "All_Before_845_SOs", "SOs_With_Absent", "Mostly_Late_SOs"]]
        table(out.rename(columns={by: nm, "Present": "Present days", "Absent": "Absent days", "Before845": "Days before 8:45", "After845": "Days 8:45+",
                                  "All_Before_845_SOs": f"SOs before 8:45 all {days} days", "SOs_With_Absent": "SOs with absent", "Mostly_Late_SOs": "SOs mostly late (8:45+)"}),
              ["Present %", "Absent %", "On-time %"])
    level_tabs(D, att)

# ---------- Visit / Memo ----------
def target_tab(act, tgt, pct, label, unit):
    P_ = D[act].sum() / D[tgt].sum() * 100
    kpis([(f"Total {label}", f"{int(D[act].sum()):,}", f"target {int(D[tgt].sum()):,}"), ("Achievement", f"{P_:.1f}%", f"gap {int(D[tgt].sum() - D[act].sum()):,}"),
          ("SOs at 100%+", f"{(D[pct] >= 100).sum()}", f"{(D[pct] >= 100).mean() * 100:.0f}% of SOs"), ("SOs below 60%", f"{(D[pct] < 60).sum()}", f"{(D[pct] < 60).mean() * 100:.0f}% of SOs")])
    b = pd.cut(D[pct], [-1, 60, 80, 100, 1e9], labels=["<60%", "60-80%", "80-100%", "100%+"], right=False).value_counts().reindex(["<60%", "60-80%", "80-100%", "100%+"]).reset_index()
    b.columns = ["Band", "SOs"]
    c1, c2 = st.columns(2)
    with c1:
        show(px.bar(b, x="Band", y="SOs", text="SOs", color="Band", title=f"SOs by {label} achievement", color_discrete_sequence=["#FF4D4F", "#FFB800", "#3F8CFF", "#00C48C"]))
    with c2:
        g = grp_table(D, "grp")
        show(px.pie(g, names="grp", values="Visits" if act == "visits" else "Memos", hole=.5, title=f"Group share of {label}"))

    def lv(R, by, nm):
        col = "Visit %" if act == "visits" else "Memo %"
        show(px.bar(R.sort_values(col), x=col, y=by, orientation="h", text=R.sort_values(col)[col].round(1), color=col, color_continuous_scale="RdYlGn",
                    title=f"{nm}-wise {label} achievement %"), max(300, 55 * len(R)))
    level_tabs(D, lv)
    cols = ["so", "rsm", "sm", "grp", "present", act, tgt, pct] + (["conv"] if act == "memos" else ["v_day"])
    table(D.sort_values(pct, ascending=False)[cols].rename(columns={"so": "SO", "rsm": "RSM", "sm": "SM", "grp": "Group", "present": "Present days", act: label, tgt: "Target", pct: "Achievement %",
                                                                    "conv": "Memo per 100 visits", "v_day": "Visits/day"}), ["Achievement %"], 100, 420)


with tabs[2]:
    target_tab("visits", "visit_t", "visit_pct", "Outlet Visits", "outlets")
with tabs[3]:
    target_tab("memos", "memo_t", "memo_pct", "Memos", "memos")

# ---------- Top & Bottom ----------
with tabs[4]:
    mt_ = st.radio("Rank by", ["Overall Score", "Outlet Visit %", "Memo %", "Attendance %", "On-time %"], horizontal=True)
    col = {"Overall Score": "score", "Outlet Visit %": "visit_pct", "Memo %": "memo_pct", "Attendance %": "present_pct", "On-time %": "punct_pct"}[mt_]
    n = min(10, len(D))
    for nm, sub, clr in (("🏆 Top 10", D.nlargest(n, col), "#00C48C"), ("⚠️ Bottom 10", D.nsmallest(n, col), "#FF4D4F")):
        show(px.bar(sub.sort_values(col), x=col, y="so", orientation="h", text=sub.sort_values(col)[col].round(1), hover_data=["rsm", "sm", "visit_pct", "memo_pct", "present_pct", "punct_pct"],
                    title=f"{nm} SO - {mt_}", color_discrete_sequence=[clr]), 400)
    st.subheader("SO analysis by SM / RSM")
    lv_ = st.radio("Level", ["RSM", "SM"], horizontal=True)
    key = "rsm" if lv_ == "RSM" else "sm"
    pick = st.selectbox(f"Choose {lv_}", sorted(D[key].unique()))
    sub = D[D[key] == pick].sort_values("score", ascending=False)
    show(px.bar(sub, x="so", y=["visit_pct", "memo_pct", "present_pct", "punct_pct"], barmode="group", title=f"{pick}: SO-wise Visit / Memo / Attendance / On-time %",
                labels={"value": "%", "variable": "", "so": ""}), 420)
    table(sub[["so", "grp", "visit_pct", "memo_pct", "present_pct", "punct_pct", "score", "rating"]].rename(columns={"so": "SO", "grp": "Group", "visit_pct": "Visit %", "memo_pct": "Memo %",
          "present_pct": "Present %", "punct_pct": "On-time %", "score": "Score", "rating": "Rating"}), ["Visit %", "Memo %", "Present %", "On-time %", "Score"])

# ---------- Hotspots ----------
with tabs[5]:
    st.caption("The file has market-level data (not individual shop names), so 'shop' = market here. Off-route = visited on a day that is not the route's assigned day.")
    rp = A.groupby(["sid", "market", "upa"]).agg(Days=("date", "nunique"), Outlets=("outlets", "sum"), OffRoute=("off", "sum")).reset_index()
    rp = rp.sort_values(["Days", "Outlets"], ascending=False).drop_duplicates("sid").merge(M[["sid", "so", "rsm"]], on="sid")
    st.subheader("SO who went to the same market most often")
    table(rp.head(15)[["so", "rsm", "market", "upa", "Days", "Outlets", "OffRoute"]].rename(columns={"so": "SO", "rsm": "RSM", "market": "Market", "upa": "Upazila", "Days": "Days visited", "Outlets": "Outlets", "OffRoute": "Off-route visits"}))
    mk = A.groupby(["market", "upa"]).agg(Outlets=("outlets", "sum"), Days=("date", "nunique"), SOs=("sid", "nunique")).reset_index().sort_values("Outlets", ascending=False)
    c1, c2 = st.columns(2)
    with c1:
        show(px.bar(mk.head(12).sort_values("Outlets"), x="Outlets", y="market", orientation="h", title="Top 12 markets by outlet visits"), 420)
    with c2:
        show(px.bar(A.groupby("upa").outlets.sum().nlargest(12).sort_values().reset_index(), x="outlets", y="upa", orientation="h", title="Top 12 areas (Upazila)"), 420)
    A2 = A.assign(Week="Week " + ((A.date.dt.day - 1) // 7 + 1).astype(str))
    top = A2.groupby("upa").outlets.sum().nlargest(12).index
    hm = A2[A2.upa.isin(top)].pivot_table(index="upa", columns="Week", values="outlets", aggfunc="sum", fill_value=0).loc[top]
    show(px.imshow(hm, text_auto=True, aspect="auto", color_continuous_scale="Purples", title="Week-wise area heatmap (outlet visits)"), 460)
    wk = A2.groupby(["Week", "market"]).outlets.sum().reset_index().sort_values("outlets", ascending=False).groupby("Week").head(3).sort_values(["Week", "outlets"], ascending=[True, False])
    table(wk.rename(columns={"market": "Top markets", "outlets": "Outlets"}))

# ---------- Scorecard ----------
with tabs[6]:
    st.caption(f"Score = Visit {SCORE_W[0]}% + Memo {SCORE_W[1]}% + Attendance {SCORE_W[2]}% + On-time {SCORE_W[3]}%.  85+ Excellent · 70+ Good · 50+ Needs Improvement · below 50 Critical")
    G, R_ = grp_table(D, "grp").sort_values("Score", ascending=False), grp_table(D, "rsm").sort_values("Score", ascending=False)
    kpis([("Best group", G.iloc[0].grp, f"score {G.iloc[0].Score:.1f}"), ("Weakest group", G.iloc[-1].grp, f"score {G.iloc[-1].Score:.1f}"),
          ("Best RSM team", R_.iloc[0].rsm, f"score {R_.iloc[0].Score:.1f}"), ("Weakest RSM team", R_.iloc[-1].rsm, f"score {R_.iloc[-1].Score:.1f}")])
    bars3d(list(G.grp), {"Visit %": G["Visit %"], "Memo %": G["Memo %"], "Present %": G["Present %"], "On-time %": G["On-time %"].fillna(0)}, "3D: Group performance")
    bars3d(list(R_.rsm), {"Visit %": R_["Visit %"], "Memo %": R_["Memo %"], "Present %": R_["Present %"], "On-time %": R_["On-time %"].fillna(0)}, "3D: RSM team performance")
    show(px.bar(R_.sort_values("Score"), x="Score", y="rsm", orientation="h", color="Score", color_continuous_scale="RdYlGn", range_color=[40, 90], text=R_.sort_values("Score").Score.round(1), title="RSM ranking (average SO score)"), max(320, 50 * len(R_)))
    cols = ["SOs", "Score", "Visit %", "Memo %", "Present %", "On-time %", "Need_Improve_SOs"]
    table(G[["grp"] + cols].rename(columns={"grp": "Group", "Need_Improve_SOs": "SOs needing improvement"}), ["Visit %", "Memo %", "Present %", "On-time %"], 100)
    table(R_[["rsm"] + cols].rename(columns={"rsm": "RSM", "Need_Improve_SOs": "SOs needing improvement"}), ["Visit %", "Memo %", "Present %", "On-time %"], 100)
    st.subheader("SOs who need improvement (by RSM)")
    need = D[D.rating.isin(["Needs Improvement", "Critical"])].sort_values(["rsm", "score"])
    table(need[["rsm", "so", "grp", "visit_pct", "memo_pct", "present_pct", "punct_pct", "score", "rating"]].rename(columns={"rsm": "RSM", "so": "SO", "grp": "Group", "visit_pct": "Visit %", "memo_pct": "Memo %",
          "present_pct": "Present %", "punct_pct": "On-time %", "score": "Score", "rating": "Rating"}), ["Visit %", "Memo %", "Present %", "On-time %"], 100, 400)

# ---------- AI Summary ----------
with tabs[7]:
    G, R_, S_ = grp_table(D, "grp"), grp_table(D, "rsm"), grp_table(D, "sm")
    vp, mp_ = D.visits.sum() / D.visit_t.sum() * 100, D.memos.sum() / D.memo_t.sum() * 100
    lt = D.late.sum() / max(D.present.sum(), 1) * 100
    bg, wg, br, wr = G.loc[G.Score.idxmax()], G.loc[G.Score.idxmin()], R_.loc[R_.Score.idxmax()], R_.loc[R_.Score.idxmin()]
    t1, b1 = D.nlargest(1, "score").iloc[0], D.nsmallest(1, "score").iloc[0]
    lowv, lowm = D[D.visit_pct < 60], D[D.memo_pct < 40]
    conv = D.memos.sum() / D.visits.sum() * 100
    pts = [("Overall", f"{len(D)} SOs achieved <b>{vp:.1f}%</b> of the outlet-visit target and <b>{mp_:.1f}%</b> of the memo target. Attendance is {D.present.sum() / (len(D) * days) * 100:.1f}% and {lt:.1f}% of present days started at 8:45 or later."),
           ("Groups", f"<b>{bg.grp}</b> is leading (score {bg.Score:.1f}); <b>{wg.grp}</b> needs the most support (score {wg.Score:.1f})."),
           ("RSM teams", f"Best: <b>{br.rsm}</b> ({br.Score:.1f}). Needs improvement: <b>{wr.rsm}</b> ({wr.Score:.1f}) with {int(wr.Need_Improve_SOs)} of {int(wr.SOs)} SOs below 70."),
           ("Outlet visit", f"{len(lowv)} SOs are below 60% of visit target. {(D.visit_pct >= 100).sum()} SOs reached 100%+. Average {D.visits.sum() / max(D.present.sum(), 1):.1f} visits per present day vs target {DAILY_OUTLET}."),
           ("Memo", f"Memo conversion is {conv:.1f} memos per 100 outlet visits. {len(lowm)} SOs are below 40% of memo target - they visit outlets but do not convert; coach them on selling and range presentation."),
           ("Attendance", f"{(D.absent > 0).sum()} SOs had at least 1 absent day; {int((D.early >= days).sum())} SOs reached before 8:45 on every one of {days} days; {int(((D.punct_pct < 50) & (D.present > 0)).sum())} SOs were mostly late."),
           ("Stars", f"Top SO: <b>{t1.so}</b> (score {t1.score:.1f}, RSM {t1.rsm}). Lowest SO: <b>{b1.so}</b> (score {b1.score:.1f}, RSM {b1.rsm})."),
           ("Actions", f"1) Review the {int(D.rating.isin(['Critical']).sum())} Critical SOs one-to-one this week. 2) Ask {wr.rsm} for a field-visit plan. 3) Set a 8:45 AM reminder for the mostly-late SOs. 4) Copy the daily routine of {bg.grp} group / {br.rsm} team to weaker teams.")]
    for h, t in pts:
        st.markdown(f'<div class="ai"><b>{h}:</b> {t}</div>', unsafe_allow_html=True)
    st.caption("Insights are generated automatically from the data using fixed rules and update with every new Excel file.")

# ---------- Data ----------
with tabs[8]:
    q = st.text_input("Search SO / RSM / SM")
    X = D[D.apply(lambda r: q.lower() in f"{r.so} {r.rsm} {r.sm}".lower(), axis=1)] if q else D
    out = X[["sid", "so", "grp", "sm", "rsm", "present", "absent", "leave", "early", "late", "punct_pct", "visits", "visit_t", "visit_pct", "memos", "memo_t", "memo_pct", "std_h", "eff_h", "score", "rating"]]
    out = out.rename(columns={"sid": "Staff ID", "so": "SO", "grp": "Group", "sm": "SM", "rsm": "RSM", "present": "Present", "absent": "Absent", "leave": "Leave", "early": "Before 8:45", "late": "8:45+",
                              "punct_pct": "On-time %", "visits": "Visits", "visit_t": "Visit Target", "visit_pct": "Visit %", "memos": "Memos", "memo_t": "Memo Target", "memo_pct": "Memo %",
                              "std_h": "Work Hours", "eff_h": "Effective Hours", "score": "Score", "rating": "Rating"})
    table(out.sort_values("Score", ascending=False), ["On-time %", "Visit %", "Memo %"], 100, 520)
    st.download_button("⬇️ Download this table (CSV)", out.to_csv(index=False).encode("utf-8-sig"), "SO_productivity.csv", "text/csv")
