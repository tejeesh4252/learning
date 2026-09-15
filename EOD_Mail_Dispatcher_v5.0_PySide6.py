"""
EOD Mail Dispatcher — v5.0
Multi-Group EOD Status Mailer — PySide6 edition
3-pane workspace: Groups | Excel Content | Live Email Preview
(QWebEngineView — real Chromium rendering, no CSS limitations)
Optimised: Lazy win32com, Splash Screen, Fast Startup
Author  : DomusAI x Tej Automation
"""

import os
import sys
import csv
import traceback
import datetime
from html import escape as html_escape

import pandas as pd
# ✅ win32com is NOT imported at top — loaded lazily only when Send is clicked

# ========================== CONFIGURATION ==========================

EMAIL_SUBJECT = "Daily EOD Status"
CONFIG_SHEET  = "Config"

LOG_FILE_PATH = os.path.join(
    os.path.expanduser("~"), "Desktop",
    "EOD_Dispatcher_Log.csv"
)

# ── Alter Domus Brand Colours ──────────────────────────────────────
AD_BLUE   = "#003865"
AD_LIGHT  = "#e8f0f7"
AD_ACCENT = "#0072CE"
AD_GREEN  = "#28a745"
AD_YELLOW = "#e6a817"
AD_RED    = "#dc3545"
AD_GREY   = "#6c757d"
AD_WHITE  = "#ffffff"

# ── Runtime GROUPS — populated by read_config() ───────────────────
GROUPS = {}


# ========================== LOGGING ================================

def log_message(msg):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{ts}] {msg}")


def write_audit_log(group, sent_date, sent_by=""):
    file_exists = os.path.isfile(LOG_FILE_PATH)
    try:
        with open(LOG_FILE_PATH, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Timestamp", "SentBy", "Group",
                                 "SentDate", "Status"])
            writer.writerow([
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                sent_by, group, sent_date, "Sent"
            ])
    except Exception as e:
        log_message(f"WARNING: Could not write audit log — {e}")


def already_sent_today(group):
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    if not os.path.isfile(LOG_FILE_PATH):
        return False
    try:
        with open(LOG_FILE_PATH, "r", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                if (row.get("Group",    "") == group and
                        row.get("SentDate", "").startswith(today_str) and
                        row.get("Status",   "") == "Sent"):
                    return True
    except Exception:
        pass
    return False


# ========================== FILE ACCESS CHECK ======================

def check_file_accessible(path):
    try:
        with open(path, "rb"):
            pass
        return True, ""
    except PermissionError:
        return False, (
            "Permission denied — the file may be:\n\n"
            "  • Open in Excel (please close it first)\n"
            "  • Locked by OneDrive sync (wait a moment)\n\n"
            f"File: {os.path.basename(path)}"
        )
    except Exception as e:
        return False, str(e)


# ========================== CONFIG READER ==========================

def read_config(excel_path):
    """
    Read the 'Config' sheet and build:
      GROUPS  dict  — { group: {sheet, subject, title, subtitle} }
      recipients    — { group: {to: [...], cc: [...]} }

    Config columns: GROUP | Sheet | Title | Funds | Subject | Role | Email | Active
    """
    global GROUPS
    groups_cfg = {}
    recipients = {}
    errors     = []

    try:
        df = pd.read_excel(excel_path, sheet_name=CONFIG_SHEET, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        df.fillna("", inplace=True)

        for _, row in df.iterrows():
            grp      = str(row.get("GROUP",   "")).strip()
            sheet    = str(row.get("Sheet",   "")).strip()
            title    = str(row.get("Title",   "")).strip()
            subtitle = str(row.get("Funds",   "")).strip()
            subject  = str(row.get("Subject", "")).strip()
            role     = str(row.get("Role",    "")).strip().upper()
            email    = str(row.get("Email",   "")).strip()
            active   = str(row.get("Active",  "")).strip().lower()

            if not grp or grp.lower() in ("nan", "none"):
                continue

            if grp not in groups_cfg:
                row_errors = []
                if not title    or title    in ("nan", "none"): row_errors.append("Title")
                if not subtitle or subtitle in ("nan", "none"): row_errors.append("SubTitle")
                if not subject  or subject  in ("nan", "none"): row_errors.append("Subject")
                if not sheet    or sheet    in ("nan", "none"): row_errors.append("Sheet")
                if row_errors:
                    errors.append(
                        f"[{grp}] Missing config fields: "
                        + ", ".join(row_errors)
                    )
                    continue

                groups_cfg[grp] = {
                    "sheet"    : sheet,
                    "subject"  : subject,
                    "title"    : title,
                    "subtitle" : subtitle,
                }
                recipients[grp] = {"to": [], "cc": []}

            if active not in ("yes", "y", "true", "1"):
                continue
            if "@" not in email:
                continue
            parts = email.split()
            email = next((p for p in parts if "@" in p), email)

            if role == "TO":
                recipients[grp]["to"].append(email)
            elif role == "CC":
                recipients[grp]["cc"].append(email)

    except Exception as e:
        log_message(f"ERROR reading Config: {e}")
        log_message(traceback.format_exc())
        return False, f"❌  Failed to read Config sheet:\n\n{e}", {}, {}

    if errors:
        return False, (
            "❌  Config sheet has errors:\n\n"
            + "\n".join(f"  • {e}" for e in errors)
            + "\n\nPlease fix and reload."
        ), {}, {}

    if not groups_cfg:
        return False, (
            "❌  No valid groups found in Config sheet.\n\n"
            "Please check the Config sheet has data and Active = Yes."
        ), {}, {}

    GROUPS = groups_cfg
    log_message(f"Config loaded — Groups: {list(GROUPS.keys())}")

    try:
        xl          = pd.ExcelFile(excel_path)
        sheet_names = xl.sheet_names
        missing     = []
        for grp, cfg in groups_cfg.items():
            if cfg["sheet"] not in sheet_names:
                missing.append(
                    f"  • Group '{grp}' → Sheet '{cfg['sheet']}' "
                    f"not found.\n    Available: {', '.join(sheet_names)}"
                )
        if missing:
            return False, (
                "❌  Sheet name mismatch:\n\n"
                + "\n".join(missing)
                + "\n\nCheck GROUP name matches tab name exactly."
            ), {}, {}
    except Exception as e:
        log_message(f"WARNING: Could not validate sheet names — {e}")

    return True, "", groups_cfg, recipients


# ========================== DATA READER ============================

def read_group_data(excel_path, group_key):
    cfg        = GROUPS[group_key]
    sheet_name = cfg["sheet"]
    today      = datetime.date.today()

    try:
        xl = pd.ExcelFile(excel_path)
        if sheet_name not in xl.sheet_names:
            return None, f"Sheet '{sheet_name}' not found in file."

        df = pd.read_excel(excel_path, sheet_name=sheet_name)

        cleaned_cols = []
        for c in df.columns:
            c_str = str(c).strip()
            if c_str.lower().startswith("unnamed"):
                c_str = ""
            cleaned_cols.append(c_str)
        df.columns = cleaned_cols
        df = df.loc[:, df.columns != ""]

        col_map = {}
        for col in df.columns:
            cl = col.lower().strip()
            if   "date"     in cl:                      col_map[col] = "Date"
            elif "task"     in cl or "activity" in cl:  col_map[col] = "Task"
            elif "preparer" in cl:                      col_map[col] = "Preparer"
            elif "comment"  in cl:                      col_map[col] = "Comments"
            elif "status"   in cl:                      col_map[col] = "Status"
            elif "fund"     in cl:                      col_map[col] = "Fund"
            elif "client"   in cl:                      col_map[col] = "Client"
        df.rename(columns=col_map, inplace=True)

        if "Date" not in df.columns:
            return None, "No 'Date' column found in sheet."

        raw_dates = df["Date"].dropna().unique()[:5]
        log_message(f"[{group_key}] Sample raw dates: {list(raw_dates)}")

        df["_parsed_date"] = df["Date"].apply(to_python_date)
        log_message(f"[{group_key}] Looking for today: {today}")

        today_df = df[df["_parsed_date"] == today].copy()
        today_df.drop(columns=["_parsed_date"], inplace=True)

        if today_df.empty:
            recent = sorted(
                [d for d in df["_parsed_date"].dropna().unique()
                 if d is not None],
                reverse=True
            )[:5]
            log_message(f"[{group_key}] No rows for today. "
                        f"Most recent: {recent}")
            return pd.DataFrame(), (
                f"No records for {today.strftime('%d %b %Y')}. "
                f"Latest: {recent[0] if recent else 'unknown'}"
            )

        log_message(f"[{group_key}] {len(today_df)} rows found.")
        return today_df, None

    except Exception as e:
        log_message(f"[{group_key}] ERROR in read_group_data: {e}")
        log_message(traceback.format_exc())
        return None, str(e)


# ========================== DATE UTILS =============================

def to_python_date(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = str(val).strip()
    if s.lower() in ("", "nan", "nat", "none"):
        return None
    try:
        serial = int(float(s))
        if 40000 < serial < 110000:
            return (datetime.date(1899, 12, 30)
                    + datetime.timedelta(days=serial))
    except Exception:
        pass
    if " " in s:
        s = s.split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y",
                "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%d %b %y",
                "%b %d, %Y", "%d/%m/%y", "%m/%d/%y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except Exception:
            pass
    log_message(f"WARNING: Could not parse date value: '{val}'")
    return None


# ========================== STATUS BADGE ===========================

STATUS_COLOR = {
    "completed"                                 : AD_GREEN,
    "sent to onshore"                           : AD_GREEN,
    "sent for review"                           : AD_GREEN,
    "shared to onshore"                         : AD_GREEN,
    "shared to client"                          : AD_GREEN,
    "sent to client"                            : AD_GREEN,
    "uploaded"                                  : AD_GREEN,
    "posted"                                    : AD_GREEN,
    "onshore review completed"                  : AD_GREEN,
    "upto date"                                 : AD_GREEN,
    "sent to elvira"                            : AD_GREEN,
    "sent to sara"                              : AD_GREEN,
    "re-shared to onshore"                      : AD_GREEN,
    "completed "                                : AD_GREEN,
    "posted "                                   : AD_GREEN,
    "wip"                                       : AD_YELLOW,
    "in progress"                               : AD_YELLOW,
    "inprogress"                                : AD_YELLOW,
    "under internal review"                     : AD_YELLOW,
    "internal review"                           : AD_YELLOW,
    "awaiting approval"                         : AD_YELLOW,
    "awaiting"                                  : AD_YELLOW,
    "under review"                              : AD_YELLOW,
    "working on updates"                        : AD_YELLOW,
    "wip-shared to onshore via teams"           : AD_YELLOW,
    "wip- internal review"                      : AD_YELLOW,
    "sent for approval"                         : AD_YELLOW,
    "awaiting for approval"                     : AD_YELLOW,
    "partially completed- waiting for approval" : AD_YELLOW,
    "awaiting for final approval"               : AD_YELLOW,
    "yet to start"                              : AD_YELLOW,
    "needs to start"                            : AD_YELLOW,
    "updated"                                   : AD_YELLOW,
    "pending"                                   : AD_RED,
    "on hold"                                   : AD_RED,
    "blocked"                                   : AD_RED,
    "no activity"                               : AD_GREY,
    "no cash activity"                          : AD_GREY,
    "—"                                         : AD_GREY,
}


def status_badge(status):
    s = str(status).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return "<span style='color:#999;'>—</span>"
    color = STATUS_COLOR.get(s.lower(), AD_ACCENT)
    return (f"<span style='background:{color};color:#fff;"
            f"padding:2px 9px;border-radius:10px;"
            f"font-size:11px;font-weight:600;white-space:nowrap;'>"
            f"{s}</span>")


# ========================== REFERENCE PATHS HTML ===================

def build_paths_section(paths):
    if not paths:
        return ""
    rows = "".join(
        f"<tr><td style='padding:5px 0;'>"
        f"<span style='font-family:Consolas,monospace;font-size:12px;"
        f"color:{AD_BLUE};'>📁</span>"
        f"&nbsp;<span style='font-family:Consolas,monospace;"
        f"font-size:12px;color:#333;word-break:break-all;'>{p}</span>"
        f"</td></tr>"
        for p in paths
    )
    return f"""
  <!-- REFERENCE PATHS -->
  <tr>
    <td style='padding:0 30px 24px;'>
      <table width='100%' cellpadding='0' cellspacing='0'
             style='border-top:2px solid {AD_BLUE};margin-top:4px;'>
        <tr>
          <td style='padding:12px 0 8px;'>
            <h3 style='margin:0;color:{AD_BLUE};font-size:13px;
                font-weight:700;text-transform:uppercase;
                letter-spacing:.5px;'>
              📁&nbsp; Reference Paths</h3>
          </td>
        </tr>
        {rows}
      </table>
    </td>
  </tr>"""


# ========================== RECIPIENT BANNER =======================

def build_recipient_banner(to_list, cc_list):
    def pill(addr):
        return (f"<span style='display:inline-block;"
                f"background:#e8f0f7;color:#003865;"
                f"border:1px solid #c2d4e8;"
                f"border-radius:12px;padding:3px 10px;"
                f"font-size:11px;margin:2px 4px 2px 0;"
                f"font-family:Consolas,monospace;'>"
                f"{addr}</span>")

    to_pills = "".join(pill(a) for a in to_list) if to_list else \
               "<span style='color:#999;font-style:italic;" \
               "font-size:11px;'>None configured</span>"
    cc_pills = "".join(pill(a) for a in cc_list) if cc_list else \
               "<span style='color:#999;font-style:italic;" \
               "font-size:11px;'>None</span>"

    return f"""
  <!-- RECIPIENT PREVIEW BANNER -->
  <tr>
    <td style='background:#f7f9fc;border:2px dashed #c2d4e8;
               padding:14px 30px;'>
      <table width='100%' cellpadding='0' cellspacing='0'>
        <tr>
          <td width='30' style='vertical-align:top;padding-top:4px;'>
            <span style='font-size:16px;'>📧</span>
          </td>
          <td>
            <p style='margin:0 0 4px;font-size:10px;font-weight:700;
               color:#003865;letter-spacing:1px;
               text-transform:uppercase;'>
              Preview — This email will be sent to:</p>
            <table cellpadding='0' cellspacing='0'
                   style='width:100%;margin-top:6px;'>
              <tr>
                <td style='width:28px;vertical-align:top;
                    padding-top:5px;'>
                  <span style='background:#003865;color:#fff;
                    font-size:9px;font-weight:700;padding:2px 5px;
                    border-radius:4px;'>TO</span>
                </td>
                <td style='padding-left:6px;'>{to_pills}</td>
              </tr>
              <tr><td colspan='2' style='height:6px;'></td></tr>
              <tr>
                <td style='width:28px;vertical-align:top;
                    padding-top:5px;'>
                  <span style='background:#0072CE;color:#fff;
                    font-size:9px;font-weight:700;padding:2px 5px;
                    border-radius:4px;'>CC</span>
                </td>
                <td style='padding-left:6px;'>{cc_pills}</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>"""


# ========================== BUILD HTML =============================

def build_html(today, df, group_key,
               to_list=None, cc_list=None, ref_paths=None):
    cfg          = GROUPS[group_key]
    display_date = today.strftime("%d %b %Y")
    to_list      = to_list  or []
    cc_list      = cc_list  or []
    ref_paths    = ref_paths or []

    # ── Stats ──────────────────────────────────────────────────────
    # ✅ FIX: no activity rows excluded — don't inflate pending count
    NO_ACTIVITY_KEYS = ["no activity", "no cash activity", "—"]

    total     = len(df)
    completed = sum(1 for s in df["Status"].astype(str)
                    if any(k in s.lower() for k in
                           ["completed", "sent to onshore",
                            "sent for review", "shared to onshore",
                            "uploaded", "posted",
                            "onshore review completed", "upto date",
                            "sent to client"]))
    in_prog   = sum(1 for s in df["Status"].astype(str)
                    if any(k in s.lower() for k in
                           ["wip", "in progress", "inprogress",
                            "internal review", "awaiting"]))
    no_act    = sum(1 for s in df["Status"].astype(str)
                    if any(k in s.strip().lower()
                           for k in NO_ACTIVITY_KEYS))
    pending   = total - completed - in_prog - no_act  # ✅ fixed

    # ── Group rows by Client ───────────────────────────────────────
    client_groups = {}
    for _, row in df.iterrows():
        c = str(row.get("Client", "")).strip()
        if not c or c.lower() in ("nan", "none", ""):
            c = group_key
        client_groups.setdefault(c, []).append(row)

    # ── Build table rows ───────────────────────────────────────────
    rows_html = ""
    row_num   = 0
    for client, c_rows in client_groups.items():
        first = True
        for row in c_rows:
            row_num += 1
            bg       = AD_WHITE if row_num % 2 == 0 else AD_LIGHT
            fund     = str(row.get("Fund",     "")).strip()
            task     = str(row.get("Task",     "")).strip()
            status   = str(row.get("Status",   "")).strip()
            comments = str(row.get("Comments", "")).strip()
            preparer = str(row.get("Preparer", "")).strip()

            if comments.lower() in ("nan", "none", ""): comments = "—"
            if preparer.lower() in ("nan", "none", ""): preparer = "—"

            cc = ""
            if first:
                cc = (f"<td rowspan='{len(c_rows)}' "
                      f"style='padding:8px 10px;vertical-align:middle;"
                      f"font-weight:700;background:{AD_LIGHT};"
                      f"border-right:3px solid {AD_BLUE};"
                      f"color:{AD_BLUE};font-size:12px;'>{client}</td>")
                first = False

            rows_html += f"""
            <tr style='background:{bg};border-bottom:1px solid #e0e0e0;'>
              {cc}
              <td style='padding:7px 10px;font-size:12px;color:#333;'>
                {fund}</td>
              <td style='padding:7px 10px;font-size:12px;color:#333;'>
                {task}</td>
              <td style='padding:7px 10px;text-align:center;'>
                {status_badge(status)}</td>
              <td style='padding:7px 10px;font-size:11px;color:#555;'>
                {comments}</td>
              <td style='padding:7px 10px;font-size:11px;color:#003865;
                font-weight:600;text-align:center;'>{preparer}</td>
            </tr>"""

    if not rows_html:
        rows_html = (
            f"<tr><td colspan='6' style='text-align:center;"
            f"padding:24px;color:{AD_GREY};font-style:italic;'>"
            f"No activity recorded for today.</td></tr>"
        )

    recipient_banner = (
        build_recipient_banner(to_list, cc_list)
        if (to_list or cc_list) else ""
    )
    paths_section = build_paths_section(ref_paths)

    return f"""<!DOCTYPE html><html><head><meta charset='UTF-8'></head>
<body style='margin:0;padding:0;font-family:Calibri,Arial,sans-serif;
background:#f0f4f8;'>
<table width='100%' cellpadding='0' cellspacing='0'
       style='background:#f0f4f8;padding:24px 0;'>
<tr><td align='center'>
<table width='820' cellpadding='0' cellspacing='0'
       style='background:{AD_WHITE};border-radius:10px;
       box-shadow:0 3px 12px rgba(0,0,0,0.12);overflow:hidden;'>
  <tr>
    <td style='background:{AD_BLUE};padding:24px 30px;'>
      <table width='100%' cellpadding='0' cellspacing='0'><tr>
        <td>
          <p style='margin:0;color:{AD_WHITE};font-size:10px;
             letter-spacing:2px;text-transform:uppercase;opacity:.7;'>
            End of Day Status Report</p>
          <h1 style='margin:5px 0 2px;color:{AD_WHITE};font-size:22px;
              font-weight:700;'>{cfg["title"]}</h1>
          <p style='margin:0;color:{AD_WHITE};font-size:12px;opacity:.7;'>
            {cfg["subtitle"]}</p>
        </td>
        <td align='right' style='vertical-align:top;'>
          <p style='margin:0;color:{AD_WHITE};font-size:15px;
             font-weight:600;opacity:.9;'>{display_date}</p>
          <p style='margin:4px 0 0;color:{AD_WHITE};font-size:11px;
             opacity:.65;'>Alter Domus</p>
        </td>
      </tr></table>
    </td>
  </tr>

  {recipient_banner}

  <!-- STATS BAR -->
  <tr>
    <td style='background:{AD_BLUE};padding:14px 30px;opacity:.85;'>
      <table width='100%' cellpadding='0' cellspacing='0'><tr>
        <td align='center'>
          <span style='color:#90ee90;font-size:22px;font-weight:700;'>
            {completed}</span><br>
          <span style='color:{AD_WHITE};font-size:10px;opacity:.8;'>
            COMPLETED</span>
        </td>
        <td style='border-left:1px solid rgba(255,255,255,0.3);'
            align='center'>
          <span style='color:#ffe08a;font-size:22px;font-weight:700;'>
            {in_prog}</span><br>
          <span style='color:{AD_WHITE};font-size:10px;opacity:.8;'>
            IN PROGRESS</span>
        </td>
        <td style='border-left:1px solid rgba(255,255,255,0.3);'
            align='center'>
          <span style='color:#ffaaaa;font-size:22px;font-weight:700;'>
            {pending}</span><br>
          <span style='color:{AD_WHITE};font-size:10px;opacity:.8;'>
            PENDING / ON HOLD</span>
        </td>
      </tr></table>
    </td>
  </tr>

  <!-- ACTIVITY TABLE -->
  <tr>
    <td style='padding:26px 30px;'>
      <h2 style='margin:0 0 14px;color:{AD_BLUE};font-size:14px;
          font-weight:700;text-transform:uppercase;letter-spacing:.5px;
          border-bottom:2px solid {AD_BLUE};padding-bottom:6px;'>
        Today's Activity Log</h2>
      <table width='100%' cellpadding='0' cellspacing='0'
             style='border-collapse:collapse;font-size:13px;
             border:1px solid #dde3ea;'>
        <thead>
          <tr style='background:{AD_BLUE};color:{AD_WHITE};'>
            <th style='padding:10px;text-align:left;width:12%;'>
              Client</th>
            <th style='padding:10px;text-align:left;width:14%;'>
              Fund</th>
            <th style='padding:10px;text-align:left;width:24%;'>
              Task / Activity</th>
            <th style='padding:10px;text-align:center;width:16%;'>
              Status</th>
            <th style='padding:10px;text-align:left;width:22%;'>
              Comments</th>
            <th style='padding:10px;text-align:center;width:12%;'>
              Preparer</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>

      <!-- LEGEND -->
      <table cellpadding='0' cellspacing='0' style='margin-top:12px;'>
        <tr>
          <td style='font-size:11px;color:{AD_GREY};padding-right:8px;
              font-style:italic;'>Legend:</td>
          <td style='padding-right:12px;'>
            <span style='background:{AD_GREEN};color:#fff;
            padding:2px 8px;border-radius:8px;font-size:10px;'>
              ✔ Completed / Sent</span></td>
          <td style='padding-right:12px;'>
            <span style='background:{AD_YELLOW};color:#fff;
            padding:2px 8px;border-radius:8px;font-size:10px;'>
              ⏳ WIP / In Progress</span></td>
          <td style='padding-right:12px;'>
            <span style='background:{AD_RED};color:#fff;
            padding:2px 8px;border-radius:8px;font-size:10px;'>
              ⚠ Pending / On Hold</span></td>
          <td>
            <span style='background:{AD_GREY};color:#fff;
            padding:2px 8px;border-radius:8px;font-size:10px;'>
              — No Activity</span></td>
        </tr>
      </table>
    </td>
  </tr>

  {paths_section}

  <!-- FOOTER -->
  <tr>
    <td style='background:{AD_BLUE};padding:14px 30px;'>
      <p style='margin:0;color:{AD_WHITE};font-size:11px;
         opacity:.65;text-align:center;'>
        Alter Domus &nbsp;·&nbsp; {display_date}
      </p>
    </td>
  </tr>

</table>
</td></tr></table>
</body></html>"""


# ========================== BUILD NO-DATA HTML =====================

def build_no_data_html(today, group_key,
                       to_list=None, cc_list=None, ref_paths=None):
    cfg          = GROUPS[group_key]
    display_date = today.strftime("%d %b %Y")
    to_list      = to_list  or []
    cc_list      = cc_list  or []
    ref_paths    = ref_paths or []

    recipient_banner = (
        build_recipient_banner(to_list, cc_list)
        if (to_list or cc_list) else ""
    )
    paths_section = build_paths_section(ref_paths)

    return f"""<!DOCTYPE html><html><head><meta charset='UTF-8'></head>
<body style='margin:0;padding:0;font-family:Calibri,Arial,sans-serif;
background:#f0f4f8;'>
<table width='100%' cellpadding='0' cellspacing='0'
       style='background:#f0f4f8;padding:24px 0;'>
<tr><td align='center'>
<table width='820' cellpadding='0' cellspacing='0'
       style='background:{AD_WHITE};border-radius:10px;
       box-shadow:0 3px 12px rgba(0,0,0,0.12);overflow:hidden;'>

  <!-- HEADER -->
  <tr>
    <td style='background:{AD_BLUE};padding:24px 30px;'>
      <table width='100%' cellpadding='0' cellspacing='0'><tr>
        <td>
          <p style='margin:0;color:{AD_WHITE};font-size:10px;
             letter-spacing:2px;text-transform:uppercase;opacity:.7;'>
            End of Day Status Report</p>
          <h1 style='margin:5px 0 2px;color:{AD_WHITE};font-size:22px;
              font-weight:700;'>{cfg["title"]}</h1>
          <p style='margin:0;color:{AD_WHITE};font-size:12px;opacity:.7;'>
            {cfg["subtitle"]}</p>
        </td>
        <td align='right' style='vertical-align:top;'>
          <p style='margin:0;color:{AD_WHITE};font-size:15px;
             font-weight:600;opacity:.9;'>{display_date}</p>
          <p style='margin:4px 0 0;color:{AD_WHITE};font-size:11px;
             opacity:.65;'>Alter Domus</p>
        </td>
      </tr></table>
    </td>
  </tr>

  {recipient_banner}

  <tr>
    <td style='padding:30px;text-align:center;'>
      <p style='font-size:15px;color:{AD_GREY};font-style:italic;'>
        No activity entries found for
        <strong>{display_date}</strong>.<br><br>
        Please ensure the EOD Status file is updated before sending.
      </p>
    </td>
  </tr>

  {paths_section}

  <tr>
    <td style='background:{AD_BLUE};padding:14px 30px;'>
      <p style='margin:0;color:{AD_WHITE};font-size:11px;
         opacity:.65;text-align:center;'>
        Alter Domus &nbsp;·&nbsp; {display_date}
      </p>
    </td>
  </tr>
</table>
</td></tr></table>
</body></html>"""


# ========================== SEND EMAIL =============================

def send_outlook_email(html_body, group_key, today,
                       to_list, cc_list,
                       attachments=None, ref_paths=None):
    """
    ✅ win32com loaded lazily — only imported when Send is clicked.
    This keeps startup fast.
    """
    import win32com.client as win32   # ✅ LAZY IMPORT

    attachments = attachments or []
    subject = (f"{EMAIL_SUBJECT} | "
               f"{GROUPS[group_key]['subject']} | "
               f"{today.strftime('%d %b %Y')}")

    outlook = win32.Dispatch("outlook.application")
    mail    = outlook.CreateItem(0)
    mail.Subject  = subject
    mail.HTMLBody = html_body

    for addr in to_list:
        r        = mail.Recipients.Add(addr)
        r.Type   = 1   # ✅ olTo — explicitly set
        r.Resolve()    # ✅ Force Outlook to resolve the address

    for addr in cc_list:
        r        = mail.Recipients.Add(addr)
        r.Type   = 2   # olCC
        r.Resolve()    # ✅ FIX — was missing, causing CC to be dropped

    for fp in attachments:
        mail.Attachments.Add(Source=fp)
        log_message(f"[{group_key}] 📎 Attached: {os.path.basename(fp)}")

    mail.Send()
    log_message(f"[{group_key}] ✅ Email sent — "
                f"TO: {len(to_list)}  CC: {len(cc_list)}  "
                f"Attachments: {len(attachments)}")



# ========================== GUI (PySide6) ===========================

from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QColor, QPixmap, QPainter, QFont
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QMessageBox,
    QScrollArea, QFrame, QCheckBox, QPlainTextEdit, QSplitter,
    QTableWidget, QTableWidgetItem, QAbstractItemView, QTextEdit,
    QSplashScreen,
)
from PySide6.QtWebEngineWidgets import QWebEngineView


class EODDispatcherWindow(QMainWindow):
    BG         = "#f5f6fa"
    PANEL_BG   = "#ffffff"
    WARN_BG    = "#fff7e6"
    HEADER_BG  = AD_BLUE
    HEADER_FG  = "#ffffff"
    BTN_SEND   = "#E83A2C"
    BTN_ACCENT = "#0072CE"
    TEXT_DARK  = "#1a1a2e"
    TEXT_MID   = "#444444"
    GREEN      = AD_GREEN
    YELLOW     = AD_YELLOW
    RED        = AD_RED
    GREY       = AD_GREY

    LOG_COLORS = {
        "green":  "#3fb950",
        "yellow": "#e3b341",
        "red":    "#f85149",
        "blue":   "#58a6ff",
        "grey":   "#8b949e",
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle("EOD Mail Dispatcher  —  v5.0  (PySide6)")
        self.resize(1500, 920)
        self.setStyleSheet(f"background:{self.BG};")

        self.file_path          = ""
        self.recipients         = {}
        self.group_cards        = {}
        self.send_checks        = {}
        self.no_activity_checks = {}
        self.status_labels      = {}
        self.preview_buttons    = {}
        self.attachment_paths   = {}
        self.attachment_labels  = {}
        self.path_edits         = {}
        self.group_dataframes   = {}   # cache of last-loaded df per group
        self.focused_group      = None

        self._build_ui()

    # ── UI BUILD ─────────────────────────────────────────────────

    def _btn_style(self, color):
        return (f"QPushButton {{ background:{color}; color:white; border:none;"
                f"border-radius:4px; padding:6px 12px; font-weight:600; }}"
                f"QPushButton:disabled {{ background:#c7cdd6; color:#8a8f99; }}")

    def _section_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet(
            f"color:{self.BTN_ACCENT}; font-weight:700; font-size:11px; "
            f"padding:4px 0; letter-spacing:.5px;")
        return lbl

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(12, 10, 12, 10)
        root.setSpacing(8)

        header = QLabel("EOD MAIL DISPATCHER   ·   Alter Domus")
        header.setStyleSheet(
            f"background:{self.HEADER_BG}; color:{self.HEADER_FG};"
            f"font-size:16px; font-weight:700; padding:12px 16px;"
            f"border-radius:4px;")
        root.addWidget(header)

        # ── file bar ───────────────────────────────────────────
        file_row = QHBoxLayout()
        self.file_edit = QLineEdit()
        self.file_edit.setReadOnly(True)
        self.file_edit.setPlaceholderText("No file loaded")
        browse_btn = QPushButton("Browse…")
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.setStyleSheet(self._btn_style(self.BTN_ACCENT))
        browse_btn.clicked.connect(self._browse_file)
        file_row.addWidget(self.file_edit, 1)
        file_row.addWidget(browse_btn)
        root.addLayout(file_row)

        self.file_status_lbl = QLabel("No file loaded")
        self.file_status_lbl.setStyleSheet(
            f"color:{self.GREY}; font-style:italic; font-size:11px;")
        root.addWidget(self.file_status_lbl)

        # ── main 3-pane splitter ───────────────────────────────
        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)

        # -- left rail: groups --
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.addWidget(self._section_label("📋  GROUPS"))

        self.group_scroll = QScrollArea()
        self.group_scroll.setWidgetResizable(True)
        self.group_scroll.setStyleSheet("QScrollArea { border: none; }")
        self.group_container = QWidget()
        self.group_container_layout = QVBoxLayout(self.group_container)
        self.group_container_layout.setSpacing(6)
        self.group_container_layout.addStretch(1)
        self.group_scroll.setWidget(self.group_container)
        left_layout.addWidget(self.group_scroll, 1)

        select_row = QHBoxLayout()
        sel_all = QPushButton("✓ Select All")
        sel_all.setStyleSheet(self._btn_style(self.GREY))
        sel_all.clicked.connect(self._select_all)
        clr_all = QPushButton("✗ Clear All")
        clr_all.setStyleSheet(self._btn_style(self.GREY))
        clr_all.clicked.connect(self._clear_all)
        select_row.addWidget(sel_all)
        select_row.addWidget(clr_all)
        left_layout.addLayout(select_row)

        send_btn = QPushButton("📧  Send Selected Emails")
        send_btn.setCursor(Qt.PointingHandCursor)
        send_btn.setStyleSheet(
            f"QPushButton {{ background:{self.BTN_SEND}; color:white;"
            f"font-weight:700; font-size:13px; padding:10px; border:none;"
            f"border-radius:4px; }}")
        send_btn.clicked.connect(self._send_selected)
        left_layout.addWidget(send_btn)

        left_panel.setMinimumWidth(340)
        left_panel.setMaximumWidth(460)

        # -- middle: excel content --
        mid_panel = QWidget()
        mid_layout = QVBoxLayout(mid_panel)
        mid_layout.setContentsMargins(0, 0, 0, 0)
        self.excel_title = self._section_label("📊  EXCEL CONTENT — select a group")
        mid_layout.addWidget(self.excel_title)
        self.excel_table = QTableWidget(0, 0)
        self.excel_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.excel_table.setSelectionMode(QAbstractItemView.NoSelection)
        self.excel_table.horizontalHeader().setStretchLastSection(True)
        self.excel_table.setAlternatingRowColors(True)
        self.excel_table.setStyleSheet(
            "QTableWidget { background:white; gridline-color:#e0e0e0; }"
            f"QHeaderView::section {{ background:{self.HEADER_BG}; color:white; "
            f"padding:6px; font-weight:600; border:none; }}")
        mid_layout.addWidget(self.excel_table, 1)

        # -- right: live preview --
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        self.preview_title = self._section_label("📧  EMAIL PREVIEW — select a group")
        right_layout.addWidget(self.preview_title)
        self.preview_view = QWebEngineView()
        self.preview_view.setHtml(
            "<body style='font-family:Calibri,Arial,sans-serif;color:#999;"
            "padding:40px;'>Load a file and click a group's 👁 icon to see "
            "the email here.</body>")
        right_layout.addWidget(self.preview_view, 1)

        splitter.addWidget(left_panel)
        splitter.addWidget(mid_panel)
        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 1)
        splitter.setSizes([380, 480, 600])

        root.addWidget(splitter, 1)

        # ── activity log ───────────────────────────────────────
        root.addWidget(self._section_label("🗒  ACTIVITY LOG"))
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFixedHeight(140)
        self.log_box.setStyleSheet(
            "background:#1e1e2e; color:#cdd6f4; font-family:Consolas,monospace;"
            "font-size:11px; border:none;")
        root.addWidget(self.log_box)

        self._log("Ready. Browse & Load the EOD Status file to begin.", "grey")

    # ── GROUP CARDS ──────────────────────────────────────────────

    def _build_group_cards(self):
        while self.group_container_layout.count() > 1:
            item = self.group_container_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

        self.group_cards        = {}
        self.send_checks        = {}
        self.no_activity_checks = {}
        self.status_labels      = {}
        self.preview_buttons    = {}
        self.attachment_paths   = {g: [] for g in GROUPS}
        self.attachment_labels  = {}
        self.path_edits         = {}
        self.group_dataframes   = {}
        self.focused_group      = None

        for group in GROUPS:
            card = self._make_group_card(group)
            self.group_container_layout.insertWidget(
                self.group_container_layout.count() - 1, card)
            self.group_cards[group] = card

    def _make_group_card(self, group):
        card = QFrame()
        card.setStyleSheet(
            f"QFrame {{ background:{self.PANEL_BG}; border:1px solid #d8dee8;"
            f"border-radius:4px; }}")
        lay = QVBoxLayout(card)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(4)

        top = QHBoxLayout()
        chk = QCheckBox(group)
        chk.setStyleSheet(f"font-weight:700; color:{self.TEXT_DARK};")
        top.addWidget(chk)
        top.addStretch(1)
        prev_btn = QPushButton("👁")
        prev_btn.setFixedWidth(34)
        prev_btn.setCursor(Qt.PointingHandCursor)
        prev_btn.setStyleSheet(self._btn_style(self.BTN_ACCENT))
        prev_btn.setEnabled(False)
        prev_btn.clicked.connect(lambda _, g=group: self._focus_group(g))
        top.addWidget(prev_btn)
        lay.addLayout(top)

        status_lbl = QLabel("— load file first")
        status_lbl.setStyleSheet(f"color:{self.GREY}; font-size:10px; font-style:italic;")
        lay.addWidget(status_lbl)

        # no-activity confirmation row — hidden unless the sheet has
        # zero rows for this group today; see _set_no_activity_state
        na_row = QFrame()
        na_row.setStyleSheet(f"background:{self.WARN_BG}; border-radius:3px;")
        na_lay = QHBoxLayout(na_row)
        na_lay.setContentsMargins(6, 4, 6, 4)
        na_cb = QCheckBox("⚠  Confirm: genuinely no activity today — OK to send")
        na_cb.setStyleSheet("color:#8a5a00; font-weight:600; font-size:10px;")
        na_cb.stateChanged.connect(lambda _s, g=group: self._on_no_activity_confirm(g))
        na_lay.addWidget(na_cb)
        na_row.setVisible(False)
        lay.addWidget(na_row)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color:#e5e5e5;")
        lay.addWidget(sep)

        att_row = QHBoxLayout()
        att_row.addWidget(QLabel("📎"))
        att_btn = QPushButton("+ Attach")
        att_btn.setStyleSheet(self._btn_style(self.BTN_ACCENT))
        att_btn.clicked.connect(lambda _, g=group: self._browse_attachments(g))
        att_clear = QPushButton("✕")
        att_clear.setFixedWidth(28)
        att_clear.setStyleSheet(self._btn_style(self.GREY))
        att_clear.clicked.connect(lambda _, g=group: self._clear_attachments(g))
        att_files_lbl = QLabel("No files attached")
        att_files_lbl.setStyleSheet(f"color:{self.GREY}; font-size:10px; font-style:italic;")
        att_files_lbl.setWordWrap(True)
        att_row.addWidget(att_btn)
        att_row.addWidget(att_clear)
        att_row.addWidget(att_files_lbl, 1)
        lay.addLayout(att_row)

        path_row = QHBoxLayout()
        path_row.addWidget(QLabel("📁"))
        path_edit = QPlainTextEdit()
        path_edit.setFixedHeight(44)
        path_edit.setPlaceholderText("One reference path per line…")
        path_edit.setStyleSheet("font-family:Consolas,monospace; font-size:10px;")
        path_row.addWidget(path_edit, 1)
        lay.addLayout(path_row)

        self.send_checks[group]        = chk
        self.no_activity_checks[group] = na_cb
        self.status_labels[group]      = status_lbl
        self.preview_buttons[group]    = prev_btn
        self.attachment_labels[group]  = att_files_lbl
        self.path_edits[group]         = path_edit
        card._na_row                   = na_row  # stashed for show/hide

        return card

    # ── NO-ACTIVITY CONFIRMATION GATE ───────────────────────────

    def _on_no_activity_confirm(self, group):
        cb  = self.no_activity_checks[group]
        chk = self.send_checks[group]
        confirmed = cb.isChecked()
        chk.setChecked(confirmed)
        chk.setEnabled(confirmed)
        if confirmed:
            self._log(f"[{group}] No-activity send confirmed by user.", "blue")
        else:
            self._log(f"[{group}] No-activity confirmation withdrawn.", "yellow")

    def _set_no_activity_state(self, group, is_no_activity):
        """Show/lock or hide/unlock the no-activity confirmation row.
        A zero-row group can never auto-send — the send checkbox
        stays disabled until this box is explicitly ticked."""
        card   = self.group_cards[group]
        na_row = card._na_row
        chk    = self.send_checks[group]
        na_cb  = self.no_activity_checks[group]

        na_cb.blockSignals(True)
        na_cb.setChecked(False)
        na_cb.blockSignals(False)

        if is_no_activity:
            na_row.setVisible(True)
            chk.setChecked(False)
            chk.setEnabled(False)
        else:
            na_row.setVisible(False)
            chk.setEnabled(True)

    # ── FILE BROWSE & LOAD ───────────────────────────────────────

    def _browse_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select EOD Status Excel File", "",
            "Excel Files (*.xlsx *.xls);;All Files (*.*)")
        if path:
            self.file_path = path
            self.file_edit.setText(path)
            self._load_file()

    def _load_file(self):
        path = self.file_path
        if not path or not os.path.isfile(path):
            QMessageBox.warning(self, "Invalid File",
                                "Please select a valid Excel file.")
            return

        ok, reason = check_file_accessible(path)
        if not ok:
            QMessageBox.critical(self, "File Locked", reason)
            self.file_status_lbl.setText(
                "❌  File is locked — close Excel / wait for OneDrive sync")
            self.file_status_lbl.setStyleSheet(f"color:{self.RED}; font-size:11px;")
            self._log(f"File locked: {reason}", "red")
            return

        self._log(f"Loading: {os.path.basename(path)}", "blue")

        ok, err_msg, groups_cfg, recipients = read_config(path)
        if not ok:
            QMessageBox.critical(self, "Config Error", err_msg)
            self.file_status_lbl.setText("❌  Config error — see details")
            self.file_status_lbl.setStyleSheet(f"color:{self.RED}; font-size:11px;")
            self._log(f"Config error: {err_msg}", "red")
            return

        self.recipients = recipients
        self._build_group_cards()

        today     = datetime.date.today()
        found_any = False

        for group in GROUPS:
            to_list = recipients.get(group, {}).get("to", [])
            status_lbl = self.status_labels[group]

            if not to_list:
                status_lbl.setText("⚠️  No recipients configured — will be skipped")
                status_lbl.setStyleSheet(f"color:{self.YELLOW}; font-size:10px;")
                self.preview_buttons[group].setEnabled(False)
                self._set_no_activity_state(group, False)
                self.send_checks[group].setChecked(False)
                self._log(f"[{group}] ⚠️ No recipients found — auto-unchecked.",
                          "yellow")
                continue

            df, err = read_group_data(path, group)
            self.group_dataframes[group] = df
            sent = already_sent_today(group)

            if sent:
                status_lbl.setText("✅  Already sent today")
                status_lbl.setStyleSheet(f"color:{self.GREEN}; font-size:10px;")
                self.preview_buttons[group].setEnabled(True)
                self._set_no_activity_state(group, False)
                self.send_checks[group].setChecked(False)
                self._log(f"[{group}] Already sent today.", "yellow")

            elif df is None:
                status_lbl.setText(f"❌  Error — {err}")
                status_lbl.setStyleSheet(f"color:{self.RED}; font-size:10px;")
                self.preview_buttons[group].setEnabled(False)
                self._set_no_activity_state(group, False)
                self.send_checks[group].setChecked(False)
                self._log(f"[{group}] ERROR — {err}", "red")

            elif df.empty:
                status_lbl.setText(f"⚠️  No data for today — {err}")
                status_lbl.setStyleSheet(f"color:{self.YELLOW}; font-size:10px;")
                self.preview_buttons[group].setEnabled(True)
                self._set_no_activity_state(group, True)
                self._log(
                    f"[{group}] No data for today — awaiting manual "
                    f"'no activity' confirmation before it can send.",
                    "yellow")
                found_any = True

            else:
                status_lbl.setText(f"✅  {len(df)} row(s) ready")
                status_lbl.setStyleSheet(f"color:{self.GREEN}; font-size:10px;")
                self.preview_buttons[group].setEnabled(True)
                self._set_no_activity_state(group, False)
                self.send_checks[group].setChecked(True)
                self._log(f"[{group}] {len(df)} row(s) loaded.", "green")
                found_any = True

        if found_any:
            self.file_status_lbl.setText(
                f"✅  Loaded: {os.path.basename(path)}  "
                f"({today.strftime('%d %b %Y')})")
            self.file_status_lbl.setStyleSheet(f"color:{self.GREEN}; font-size:11px;")
        else:
            self.file_status_lbl.setText(
                "⚠️  File loaded but no data found for today.")
            self.file_status_lbl.setStyleSheet(f"color:{self.YELLOW}; font-size:11px;")

        # auto-focus the first previewable group so the split view isn't empty
        for group in GROUPS:
            if self.preview_buttons[group].isEnabled():
                self._focus_group(group)
                break

    # ── LEFT/RIGHT SPLIT: EXCEL CONTENT + LIVE PREVIEW ─────────────

    def _focus_group(self, group):
        self.focused_group = group
        df = self.group_dataframes.get(group)
        if df is None:
            df, _ = read_group_data(self.file_path, group)
            self.group_dataframes[group] = df

        self.excel_title.setText(f"📊  EXCEL CONTENT — {group}")
        self._populate_excel_table(df)

        self.preview_title.setText(f"📧  EMAIL PREVIEW — {group}")
        self._refresh_preview(group)

    def _populate_excel_table(self, df):
        table = self.excel_table
        table.clear()
        if df is None or df.empty:
            table.setColumnCount(1)
            table.setRowCount(1)
            table.setHorizontalHeaderLabels(["Status"])
            table.setItem(0, 0, QTableWidgetItem("No rows for today."))
            return

        cols = list(df.columns)
        table.setColumnCount(len(cols))
        table.setHorizontalHeaderLabels(cols)
        table.setRowCount(len(df))
        for r, (_, row) in enumerate(df.iterrows()):
            for c, col in enumerate(cols):
                val = str(row.get(col, ""))
                if val.lower() in ("nan", "none"):
                    val = ""
                table.setItem(r, c, QTableWidgetItem(val))
        table.resizeColumnsToContents()

    def _refresh_preview(self, group):
        df = self.group_dataframes.get(group)
        if df is None:
            df, _ = read_group_data(self.file_path, group)
            self.group_dataframes[group] = df
        if df is None:
            self.preview_view.setHtml(
                "<body style='font-family:Calibri,Arial,sans-serif;"
                "color:#c0392b;padding:40px;'>"
                "Could not read data for this group.</body>")
            return

        today     = datetime.date.today()
        to_list   = self.recipients.get(group, {}).get("to", [])
        cc_list   = self.recipients.get(group, {}).get("cc", [])
        ref_paths = self._get_ref_paths(group)

        html = (
            build_html(today, df, group, to_list=to_list,
                      cc_list=cc_list, ref_paths=ref_paths)
            if not df.empty else
            build_no_data_html(today, group, to_list=to_list,
                               cc_list=cc_list, ref_paths=ref_paths)
        )
        self.preview_view.setHtml(html)

    # ── ATTACHMENT HELPERS ───────────────────────────────────────

    def _browse_attachments(self, group):
        files, _ = QFileDialog.getOpenFileNames(
            self, f"Select attachment(s) for {group}", "",
            "All Files (*.*);;Excel (*.xlsx *.xls);;PDF (*.pdf);;"
            "Word (*.docx *.doc)")
        if files:
            duplicates, added = [], []
            for f in files:
                if f in self.attachment_paths[group]:
                    duplicates.append(os.path.basename(f))
                else:
                    self.attachment_paths[group].append(f)
                    added.append(os.path.basename(f))
            if added:
                self._refresh_attachment_label(group)
                self._log(f"[{group}] 📎 {len(added)} file(s) added.", "blue")
            if duplicates:
                QMessageBox.warning(
                    self, "Duplicate Attachment",
                    f"Already added for {group} — skipped:\n\n"
                    + "\n".join(f"  • {d}" for d in duplicates))

    def _clear_attachments(self, group):
        self.attachment_paths[group].clear()
        self._refresh_attachment_label(group)
        self._log(f"[{group}] Attachments cleared.", "grey")

    def _refresh_attachment_label(self, group):
        files = self.attachment_paths.get(group, [])
        lbl   = self.attachment_labels[group]
        if not files:
            lbl.setText("No files attached")
            lbl.setStyleSheet(f"color:{self.GREY}; font-size:10px; font-style:italic;")
        else:
            names = ",  ".join(os.path.basename(f) for f in files)
            lbl.setText(names)
            lbl.setStyleSheet(f"color:{self.GREEN}; font-size:10px;")

    def _get_ref_paths(self, group):
        edit = self.path_edits.get(group)
        if edit is None:
            return []
        raw = edit.toPlainText().strip()
        return [p.strip() for p in raw.splitlines() if p.strip()]

    # ── SELECT ALL / CLEAR ALL ──────────────────────────────────

    def _select_all(self):
        for chk in self.send_checks.values():
            if chk.isEnabled():
                chk.setChecked(True)

    def _clear_all(self):
        for chk in self.send_checks.values():
            chk.setChecked(False)

    # ── SEND ─────────────────────────────────────────────────────

    def _send_selected(self):
        path = self.file_path
        if not path:
            QMessageBox.warning(self, "No File",
                                "Please browse and load the EOD Status file first.")
            return

        ok, reason = check_file_accessible(path)
        if not ok:
            QMessageBox.critical(self, "File Locked", reason)
            return

        if not GROUPS:
            QMessageBox.warning(self, "No Groups",
                                "No groups loaded. Please load the EOD Status file first.")
            return

        selected = [g for g, chk in self.send_checks.items() if chk.isChecked()]
        if not selected:
            QMessageBox.warning(self, "Nothing Selected",
                                "Please check at least one group to send.")
            return

        # ── no-activity confirmation guard (defensive re-check) ────
        unconfirmed_no_activity = []
        for group in selected:
            df, _ = read_group_data(path, group)
            if (df is not None and df.empty and
                    not self.no_activity_checks[group].isChecked()):
                unconfirmed_no_activity.append(group)

        if unconfirmed_no_activity:
            QMessageBox.critical(
                self, "Confirmation Required",
                "⛔  Cannot send — these groups have zero rows for today "
                "and have not been confirmed as genuinely 'no activity':\n\n"
                + "\n".join(f"  • {g}" for g in unconfirmed_no_activity)
                + "\n\nTick the confirm box on each group, or update the "
                  "EOD Status file, then try again.")
            return

        # ── attachment validation ────────────────────────────────
        MAX_ATTACHMENT_MB  = 10
        broken_attachments = []
        large_attachments  = []

        for group in selected:
            for fp in self.attachment_paths.get(group, []):
                if not os.path.exists(fp):
                    broken_attachments.append(
                        f"[{group}]  {os.path.basename(fp)} (file not found)")
                else:
                    ok_a, _ = check_file_accessible(fp)
                    if not ok_a:
                        broken_attachments.append(
                            f"[{group}]  {os.path.basename(fp)} (locked/in use)")
                    else:
                        size_mb = os.path.getsize(fp) / (1024 * 1024)
                        if size_mb > MAX_ATTACHMENT_MB:
                            large_attachments.append(
                                f"[{group}]  {os.path.basename(fp)} "
                                f"({size_mb:.1f} MB)")

        if broken_attachments:
            QMessageBox.critical(
                self, "Attachment Error",
                "⛔  Cannot send — files missing or locked:\n\n"
                + "\n".join(f"  • {b}" for b in broken_attachments)
                + "\n\nPlease fix and try again.")
            return

        if large_attachments:
            resp = QMessageBox.question(
                self, "Large Attachments",
                f"⚠️  Files exceed {MAX_ATTACHMENT_MB} MB:\n\n"
                + "\n".join(f"  • {a}" for a in large_attachments)
                + "\n\nSend anyway?")
            if resp != QMessageBox.Yes:
                return

        # ── confirm dialog ────────────────────────────────────────
        confirm_lines = []
        for g in selected:
            line = f"  •  {g}"
            atts = self.attachment_paths.get(g, [])
            refs = self._get_ref_paths(g)
            if atts:
                line += f"\n       📎 {', '.join(os.path.basename(f) for f in atts)}"
            if refs:
                line += f"\n       📁 {len(refs)} path(s)"
            confirm_lines.append(line)

        resp = QMessageBox.question(
            self, "Confirm Send",
            "You are about to send EOD emails for:\n\n"
            + "\n".join(confirm_lines)
            + f"\n\nDate: {datetime.date.today().strftime('%d %b %Y')}"
            + "\n\nProceed?")
        if resp != QMessageBox.Yes:
            return

        today   = datetime.date.today()
        success, skipped, failed = [], [], []

        for group in selected:
            try:
                self._log(f"[{group}] Preparing…", "blue")

                if already_sent_today(group):
                    self._log(f"[{group}] Skipped — already sent today.",
                              "yellow")
                    continue

                df, err = read_group_data(path, group)
                if df is None:
                    raise Exception(err)

                ref_paths   = self._get_ref_paths(group)
                attachments = self.attachment_paths.get(group, [])

                html = (build_html(today, df, group, ref_paths=ref_paths)
                        if not df.empty
                        else build_no_data_html(today, group,
                                                ref_paths=ref_paths))

                to_list = self.recipients.get(group, {}).get("to", [])
                cc_list = self.recipients.get(group, {}).get("cc", [])

                if not to_list:
                    self._log(f"[{group}] ⚠️ Skipped — no recipients.",
                              "yellow")
                    self.status_labels[group].setText("⚠️  Skipped — no recipients")
                    self.status_labels[group].setStyleSheet(
                        f"color:{self.YELLOW}; font-size:10px;")
                    skipped.append(group)
                    continue

                send_outlook_email(
                    html, group, today, to_list, cc_list,
                    attachments=attachments, ref_paths=ref_paths)

                write_audit_log(group, today.strftime("%Y-%m-%d"))

                self.status_labels[group].setText("✅  Sent just now!")
                self.status_labels[group].setStyleSheet(
                    f"color:{self.GREEN}; font-size:10px;")
                self.send_checks[group].setChecked(False)
                self._set_no_activity_state(group, False)
                success.append(group)

                self.attachment_paths[group].clear()
                self._refresh_attachment_label(group)
                self.path_edits[group].clear()

                self._log(
                    f"[{group}] ✅ Sent! "
                    f"(📎 {len(attachments)} attachment(s), "
                    f"📁 {len(ref_paths)} path(s))", "green")

            except Exception as e:
                failed.append(group)
                self._log(f"[{group}] ❌ FAILED — {e}", "red")
                log_message(traceback.format_exc())

        msg = ""
        if success:
            msg += "✅  Sent:\n" + "\n".join(f"  • {g}" for g in success)
        if skipped:
            msg += ("\n\n" if msg else "") + \
                   "⚠️  Skipped:\n" + \
                   "\n".join(f"  • {g}" for g in skipped)
        if failed:
            msg += ("\n\n" if msg else "") + \
                   "❌  Failed:\n" + \
                   "\n".join(f"  • {g}" for g in failed)
        if msg:
            QMessageBox.information(self, "Send Complete", msg.strip())

    # ── HELPERS ──────────────────────────────────────────────────

    def _log(self, message, tag="grey"):
        ts    = datetime.datetime.now().strftime("%H:%M:%S")
        color = self.LOG_COLORS.get(tag, "#8b949e")
        self.log_box.append(
            f"<span style='color:{color};'>[{ts}]  {html_escape(message)}</span>")
        bar = self.log_box.verticalScrollBar()
        bar.setValue(bar.maximum())
        log_message(message)


# ========================== ENTRY POINT ============================

def _make_splash_pixmap():
    pix = QPixmap(420, 130)
    pix.fill(QColor(AD_BLUE))
    painter = QPainter(pix)
    painter.setPen(QColor("#ffffff"))
    painter.setFont(QFont("Calibri", 18, QFont.Bold))
    painter.drawText(pix.rect().adjusted(0, 24, 0, 0),
                     Qt.AlignHCenter | Qt.AlignTop, "EOD Mail Dispatcher")
    painter.setFont(QFont("Calibri", 9))
    painter.setPen(QColor("#90b8d8"))
    painter.drawText(pix.rect().adjusted(0, 62, 0, 0),
                     Qt.AlignHCenter | Qt.AlignTop,
                     "Alter Domus  ·  Starting up, please wait...")
    painter.end()
    return pix


if __name__ == "__main__":
    app = QApplication(sys.argv)

    splash = QSplashScreen(_make_splash_pixmap())
    splash.show()
    app.processEvents()

    window = EODDispatcherWindow()
    splash.finish(window)
    window.show()

    sys.exit(app.exec())
