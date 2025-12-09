
# CWC.py — Coffee & Connect 2026 Schedule App
# -------------------------------------------------
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import json
import re

# -------------------------------------------------
# App config & constants
# -------------------------------------------------
st.set_page_config(
    page_title="Coffee & Connect 2026",
    page_icon="☕",
    layout="wide"
)

APP_TITLE = "Coffee & Connect — 2026 Schedule (Atul + 4)"
IST = ZoneInfo("Asia/Kolkata")  # Chennai timezone
DEFAULT_FILE = "Coffee_Connect_2026_Schedule.xlsx"

# -------------------------------------------------
# Utilities
# -------------------------------------------------
def clean_text(x: str) -> str:
    """Remove Excel artifacts like _x000D_ and stray control chars."""
    if not isinstance(x, str):
        return x
    x = re.sub(r"_x000D_", " ", x)           # remove Excel line-break artifact
    x = re.sub(r"[\u0000-\u001F]+", " ", x)  # strip control chars
    return " ".join(x.split())

def parse_time_range(tstr: str):
    """
    Parse '15:30–16:00 IST' or '15:30-16:00 IST' into (start_time, end_time).
    If end is missing, default to +30 mins duration later in enrich function.
    """
    if not isinstance(tstr, str):
        return None, None
    s = clean_text(tstr).replace(" IST", "").strip()
    s = s.replace("–", "-")
    parts = [p.strip() for p in s.split("-")]
    start_t = None
    end_t = None
    try:
        start_t = datetime.strptime(parts[0], "%H:%M").time()
    except Exception:
        start_t = None
    if len(parts) > 1:
        try:
            end_t = datetime.strptime(parts[1], "%H:%M").time()
        except Exception:
            end_t = None
    return start_t, end_t

def enrich_schedule(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build Start/End datetimes (IST) and clean text fields.

    Expected columns (from your Excel):
    Month #, Month, Week, Date, Day, Time, Location Focus, Team Focus,
    Participants (4), Manager, Notes, Mode of Connect.
    """
    df = df.copy()

    # Clean text columns
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].apply(clean_text)

    # Parse date/time into datetimes
    starts, ends = [], []
    for _, row in df.iterrows():
        date_str = str(row.get("Date", "")).strip()
        time_str = str(row.get("Time", "")).strip()
        try:
            d = datetime.strptime(date_str, "%Y-%m-%d").date()
        except Exception:
            d = None

        stime, etime = parse_time_range(time_str)
        if d and stime:
            start_dt = datetime.combine(d, stime).replace(tzinfo=IST)
        else:
            start_dt = None
        if d and etime:
            end_dt = datetime.combine(d, etime).replace(tzinfo=IST)
        else:
            end_dt = start_dt + timedelta(minutes=30) if start_dt else None

        starts.append(start_dt)
        ends.append(end_dt)

    df["Start (IST)"] = starts
    df["End (IST)"] = ends

    # If Month # or Day missing, reconstruct from datetime
    if "Month #" not in df.columns:
        df["Month #"] = df["Start (IST)"].dt.month
    if "Day" not in df.columns:
        df["Day"] = df["Start (IST)"].dt.strftime("%a")

    # Ensure sorted
    df = df.sort_values(["Start (IST)", "Week"])
    return df

def upcoming_between(df: pd.DataFrame, now: datetime, horizon: timedelta) -> pd.DataFrame:
    """Sessions starting between now and now+horizon."""
    mask = (df["Start (IST)"] > now) & (df["Start (IST)"] <= now + horizon)
    return df.loc[mask].sort_values("Start (IST)")

def toast_events(df: pd.DataFrame, label: str):
    """Show Streamlit toasts (in-app) for upcoming events."""
    for _, r in df.iterrows():
        st.toast(
            f"🔔 {label}: {r['Week']} • "
            f"{r['Start (IST)'].strftime('%b %d, %Y %I:%M %p')} • "
            f"Participants: {r['Participants (4)']} • Mode: {r.get('Mode of Connect','')}"
        )

def js_browser_notifications(events):
    """
    Inject JS to trigger browser notifications 1 day and 30 minutes prior.
    Requires user to click 'Enable browser notifications' and keep the tab open.
    """
    payload = json.dumps(events)
    st.components.v1.html(f"""
    <script>
      const events = {payload};

      function notifyNow(title, body) {{
        try {{
          new Notification(title, {{ body: body }});
        }} catch (e) {{
          console.log("Notification error:", e);
        }}
      }}

      async function setupNotifications() {{
        if (!('Notification' in window)) {{
          alert("Browser notifications are not supported in this browser.");
          return;
        }}
        let perm = Notification.permission;
        if (perm !== 'granted') {{
          try {{
            perm = await Notification.requestPermission();
          }} catch (e) {{
            console.log("Permission request error:", e);
          }}
        }}
        if (perm !== 'granted') {{
          alert("Please allow notifications in your browser.");
          return;
        }}

        const now = Date.now();
        events.forEach(ev => {{
          // 1-day prior
          const t1 = ev.start_ts - 24*60*60*1000 - now;
          if (t1 > 0 && t1 < 26*60*60*1000) {{
            setTimeout(() => notifyNow("⏰ 1-day reminder: " + ev.title, ev.body), t1);
          }} else if (t1 <= 0 && ev.start_ts > now && (ev.start_ts - now) < 24*60*60*1000) {{
            // Already in 24h window, fire immediately
            notifyNow("⏰ 1-day reminder: " + ev.title, ev.body);
          }}

          // 30-min prior
          const t2 = ev.start_ts - 30*60*1000 - now;
          if (t2 > 0 && t2 < 2*60*60*1000) {{
            setTimeout(() => notifyNow("⏳ 30-min reminder: " + ev.title, ev.body), t2);
          }} else if (t2 <= 0 && ev.start_ts > now && (ev.start_ts - now) < 30*60*1000) {{
            notifyNow("⏳ 30-min reminder: " + ev.title, ev.body);
          }}
        }});

        alert("✅ Browser notifications enabled for sessions in the next 24 hours.");
      }}
    </script>
    <div style="margin:0.5rem 0;">
      <button onclick="setupNotifications()" style="padding:0.5rem 0.8rem; border-radius:6px; background:#2b7de9; color:white; border:0;">
        🔔 Enable browser notifications
      </button>
    </div>
    """, height=80)

def ics_from_rows(rows: pd.DataFrame, cal_name="Coffee & Connect 2026"):
    """Generate one ICS string from selected rows."""
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//CoffeeConnect//Schedule//EN",
        f"X-WR-CALNAME:{cal_name}"
    ]
    for _, r in rows.iterrows():
        start = r["Start (IST)"]
        end = r["End (IST)"]
        if not isinstance(start, datetime) or not isinstance(end, datetime):
            continue
        dtstart_utc = start.astimezone(ZoneInfo("UTC")).strftime("%Y%m%dT%H%M%SZ")
        dtend_utc = end.astimezone(ZoneInfo("UTC")).strftime("%Y%m%dT%H%M%SZ")
        uid = f"{dtstart_utc}-{r.get('Week','W')}-coffee-connect@atul"
        summary = f"Coffee & Connect — {r.get('Week','')}"
        desc = (
            f"Manager: {r.get('Manager','Atul Anand')}\n"
            f"Participants: {r.get('Participants (4)','')}\n"
            f"Focus: {r.get('Location Focus','')}\n"
            f"Mode: {r.get('Mode of Connect','')}"
        )
        loc = r.get("Location Focus", "")
        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"DTSTAMP:{dtstart_utc}",
            f"DTSTART:{dtstart_utc}",
            f"DTEND:{dtend_utc}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            f"LOCATION:{loc}",
            "END:VEVENT"
        ]
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)

def download_bytes(name: str, data: bytes, mime: str = "text/calendar"):
    st.download_button(
        label=f"⬇️ Download {name}",
        data=data,
        file_name=name,
        mime=mime
    )

def safe_autorefresh(interval_ms: int = 60_000, key: str = "refresh_key"):
    """
    Calls Streamlit's auto-refresh if available; no-ops otherwise.
    Works across multiple Streamlit versions.
    """
    # Try the importable function (newer releases)
    try:
        from streamlit import st_autorefresh as _st_autorefresh
        _st_autorefresh(interval=interval_ms, key=key)
        return
    except Exception:
        pass

    # Try attributes on st
    st_autorefresh_attr = getattr(st, "autorefresh", None) or getattr(st, "st_autorefresh", None)
    if callable(st_autorefresh_attr):
        st_autorefresh_attr(interval=interval_ms, key=key)
        return

    # Optional: last-resort JS hard reload (commented by default)
    # st.components.v1.html(
    #     f"<script>setTimeout(() => window.parent.location.reload(), {interval_ms});</script>",
    #     height=0
    # )
    return

# -------------------------------------------------
# Sidebar — file & refresh
# -------------------------------------------------
st.sidebar.title("⚙️ Controls")
uploaded = st.sidebar.file_uploader("Upload Coffee_Connect_2026_Schedule.xlsx", type=["xlsx"])
file_source = uploaded if uploaded else DEFAULT_FILE

# Auto-refresh every minute to keep reminders current
st.sidebar.info("Auto-refresh: ON (every 60s) — ensures reminders stay accurate.")
safe_autorefresh(60_000, "refresh_key")

# -------------------------------------------------
# Load file
# -------------------------------------------------
st.title(APP_TITLE)

try:
    df_raw = pd.read_excel(file_source, engine="openpyxl")
except Exception as e:
    st.error(f"Could not read the schedule file: {e}")
    st.stop()

# Confirm expected columns and show a note
expected_cols = [
    "Month #","Month","Week","Date","Day","Time","Location Focus",
    "Team Focus","Participants (4)","Manager","Notes","Mode of Connect"
]
missing = [c for c in expected_cols if c not in df_raw.columns]
if missing:
    st.warning(f"Missing columns: {missing}. The app will still try to process available data.")
else:
    st.caption("Loaded columns match your sheet structure.")

df = enrich_schedule(df_raw)

# -------------------------------------------------
# Filters
# -------------------------------------------------
col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
months = sorted(df["Month"].unique().tolist())
weeks = sorted(df["Week"].unique().tolist())
loc_focus = sorted(df["Location Focus"].unique().tolist())
team_focus = sorted(df["Team Focus"].unique().tolist())
modes = sorted(df["Mode of Connect"].dropna().unique().tolist())

with col_f1:
    f_month = st.multiselect("Month", months, default=months)
with col_f2:
    f_week = st.multiselect("Week", weeks, default=weeks)
with col_f3:
    f_loc = st.multiselect("Location Focus", loc_focus, default=loc_focus)
with col_f4:
    f_team = st.multiselect("Team Focus", team_focus, default=team_focus)
with col_f5:
    f_mode = st.multiselect("Mode of Connect", modes, default=modes)

mask = (
    df["Month"].isin(f_month) &
    df["Week"].isin(f_week) &
    df["Location Focus"].isin(f_loc) &
    df["Team Focus"].isin(f_team) &
    df["Mode of Connect"].isin(f_mode)
)
df_f = df.loc[mask].sort_values(["Start (IST)", "Week"])

# -------------------------------------------------
# KPIs & upcoming sections
# -------------------------------------------------
now = datetime.now(IST)

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.metric("Total sessions (filtered)", len(df_f))
with k2:
    st.metric(
        "Next session (local time)",
        df_f[df_f["Start (IST)"] > now]["Start (IST)"].min().strftime("%b %d, %Y %I:%M %p")
        if (df_f["Start (IST)"] > now).any() else "—"
    )
with k3:
    st.metric("Upcoming within 1 day", len(upcoming_between(df_f, now, timedelta(days=1))))
with k4:
    st.metric("Upcoming within 30 min", len(upcoming_between(df_f, now, timedelta(minutes=30))))

st.subheader("🔜 Upcoming Sessions")
up_1d = upcoming_between(df_f, now, timedelta(days=1))
up_30m = upcoming_between(df_f, now, timedelta(minutes=30))

if len(up_1d):
    st.success("⏰ **Within 1 day**")
    st.dataframe(
        up_1d[["Month","Week","Start (IST)","Participants (4)","Location Focus","Mode of Connect","Manager"]],
        use_container_width=True
    )

if len(up_30m):
    st.warning("⏳ **Within 30 minutes**")
    st.dataframe(
        up_30m[["Month","Week","Start (IST)","Participants (4)","Location Focus","Mode of Connect","Manager"]],
        use_container_width=True
    )

# In-app toasts
toast_events(up_1d, "1-day reminder")
toast_events(up_30m, "30-min reminder")

# Browser notifications for events in current filter
events_payload = []
for _, r in df_f.iterrows():
    if not isinstance(r["Start (IST)"], datetime):
        continue
    start_ts_ms = int(r["Start (IST)"].timestamp() * 1000)
    title = f"{r['Week']} — {r['Start (IST)'].strftime('%b %d, %I:%M %p IST')}"
    body = f"Participants: {r['Participants (4)']} • Focus: {r['Location Focus']} • Mode: {r.get('Mode of Connect','')}"
    events_payload.append({
        "id": f"{r['Start (IST)'].strftime('%Y%m%d%H%M')}-{r['Week']}",
        "title": title,
        "body": body,
        "start_ts": start_ts_ms
    })
js_browser_notifications(events_payload)

# -------------------------------------------------
# Full schedule table
# -------------------------------------------------
st.subheader("📋 Full Schedule (Filtered)")
st.dataframe(
    df_f[[
        "Month #","Month","Week","Day","Date","Start (IST)","End (IST)",
        "Participants (4)","Manager","Location Focus","Team Focus","Mode of Connect","Notes"
    ]],
    use_container_width=True
)

# -------------------------------------------------
# Calendar exports (ICS)
# -------------------------------------------------
st.subheader("📅 Calendar Export")
col_dl1, col_dl2 = st.columns(2)
with col_dl1:
    ics_all = ics_from_rows(df, cal_name="Coffee & Connect 2026 — All")
    download_bytes("Coffee_Connect_2026_All.ics", ics_all.encode("utf-8"))
with col_dl2:
    ics_filtered = ics_from_rows(df_f, cal_name="Coffee & Connect 2026 — Filtered")
    download_bytes("Coffee_Connect_2026_Filtered.ics", ics_filtered.encode("utf-8"))

# -------------------------------------------------
# Info & next session card
# -------------------------------------------------
st.info(
    "🔔 **Notifications**:\n"
    "• In-app toasts display automatically when within the 1-day or 30-min window.\n"
    "• Browser pop-ups require clicking ‘Enable browser notifications’ and keeping the tab open.\n"
    "• For guaranteed reminders even when the app is closed, import the ICS into Outlook/Google Calendar.\n\n"
    "📄 **Data source**: This app expects the exact column structure from your uploaded Excel."
)

next_df = df_f[df_f["Start (IST)"] > now].sort_values("Start (IST)").head(1)
if not next_df.empty:
    r = next_df.iloc[0]
    st.success(
        f"Next session: **{r['Week']}** — "
        f"**{r['Start (IST)'].strftime('%b %d, %Y %I:%M %p IST')}** • "
        f"Participants: {r['Participants (4)']} • Mode: {r.get('Mode of Connect','')}"
    )
else:
    st.info("No future sessions in the current filter.")
