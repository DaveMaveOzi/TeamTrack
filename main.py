import os
import re
import uuid
import sqlite3
import smtplib
import io
import csv
import base64
import html
from email.message import EmailMessage
from datetime import datetime, timezone, date, timedelta
from pathlib import Path

import bcrypt
import streamlit as st
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

# =========================================================
# TEAMTRACK — MULTI-COMPANY ISSUE TRACKER MVP
# =========================================================

APP_DIR = Path(__file__).parent
DB_PATH = APP_DIR / "teamtrack.db"
UPLOAD_DIR = APP_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
LOGO_PATH = APP_DIR / "teamtrack_logo.png"

STATUSES = ["Open", "Assigned", "In Progress", "On Hold", "Resolved", "Closed"]
PRIORITIES = ["Critical", "High", "Medium", "Low"]
ROLES = ["Admin", "IT Personnel", "Team Lead", "Member", "Viewer"]
DEFAULT_TEAMS = ["IT", "HR", "Marketing", "Operations", "Finance", "Customer Support", "Product", "Design"]

st.set_page_config(
    page_title="TeamTrack",
    page_icon="🎫",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------
# THEME / CSS — ENVOYCE-INSPIRED VENTICK SYSTEM
# -----------------------------
if "teamtrack_theme" not in st.session_state:
    st.session_state.teamtrack_theme = "Light"

DARK_MODE = st.session_state.teamtrack_theme == "Dark"
if DARK_MODE:
    C = {
        "app": "#15151B", "sidebar": "#1D1D25", "surface": "#22222B",
        "surface_alt": "#292933", "text": "#F3F1F4", "muted": "#B7B3BE",
        "border": "#393844", "input": "#292933", "accent": "#A64B6D",
        "accent_hover": "#BD6384", "accent_text": "#FFFFFF", "soft": "#382832",
        "shadow": "0 8px 24px rgba(0,0,0,.18)",
    }
else:
    C = {
        "app": "#F5F5F6", "sidebar": "#FFFFFF", "surface": "#FFFFFF",
        "surface_alt": "#FFFFFF", "text": "#000000", "muted": "#4A4549",
        "border": "#E5E1E4", "input": "#FFFFFF", "accent": "#4D0F28",
        "accent_hover": "#681B39", "accent_text": "#FFFFFF", "soft": "#F4E9EE",
        "shadow": "0 5px 18px rgba(35,20,28,.045)",
    }

# Light-mode status cards: keep the success and invitation-code cards readable.
# Dark mode is intentionally left unchanged.
LIGHT_MODE_CARD_CSS = ""
if not DARK_MODE:
    LIGHT_MODE_CARD_CSS = """
.tt-created-success-card {
  background: #187A45 !important;
  color: #FFFFFF !important;
  border: 1px solid #187A45 !important;
  border-radius: 9px;
  padding: 0.85rem 1rem;
  font-weight: 600;
  margin: 0.25rem 0 0.75rem;
}
.tt-created-success-card * { color: #FFFFFF !important; }
[data-testid="stCodeBlock"] {
  background: #173B63 !important;
  border: 1px solid #173B63 !important;
  border-radius: 9px !important;
}
[data-testid="stCodeBlock"] pre,
[data-testid="stCodeBlock"] code,
[data-testid="stCodeBlock"] span {
  background: transparent !important;
  color: #FFFFFF !important;
}
[data-testid="stCodeBlock"] button,
[data-testid="stCodeBlock"] button svg {
  color: #FFFFFF !important;
  fill: #FFFFFF !important;
}
.tt-workspace-settings-note {
  background: #173B63 !important;
  color: #FFFFFF !important;
  border: 1px solid #173B63 !important;
  border-radius: 9px;
  padding: 0.9rem 1rem;
  line-height: 1.55;
  margin: 0.35rem 0 0.75rem;
}
.tt-workspace-settings-note * { color: #FFFFFF !important; }
"""

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');
:root {{
  --tt-app: {C['app']}; --tt-sidebar: {C['sidebar']}; --tt-surface: {C['surface']};
  --tt-surface-alt: {C['surface_alt']}; --tt-text: {C['text']}; --tt-muted: {C['muted']};
  --tt-border: {C['border']}; --tt-input: {C['input']}; --tt-accent: {C['accent']};
  --tt-accent-hover: {C['accent_hover']}; --tt-accent-text: {C['accent_text']};
  --tt-soft: {C['soft']}; --tt-shadow: {C['shadow']};
}}
html, body, [class*="css"], [data-testid="stAppViewContainer"] {{
  font-family: 'Poppins', sans-serif !important;
}}
.stApp, [data-testid="stAppViewContainer"] {{ background: var(--tt-app) !important; color: var(--tt-text) !important; }}
[data-testid="stHeader"] {{ background: var(--tt-app) !important; }}
[data-testid="stMain"] {{ background: var(--tt-app) !important; }}
[data-testid="stMainBlockContainer"] {{ max-width: 1440px; padding-top: 2.1rem; padding-bottom: 3rem; }}
[data-testid="stSidebar"] {{ background: var(--tt-sidebar) !important; border-right: 1px solid var(--tt-border); min-width: 265px; }}
[data-testid="stSidebar"] > div:first-child {{ background: var(--tt-sidebar) !important; }}
[data-testid="stSidebar"] * {{ color: var(--tt-text); font-family: 'Poppins', sans-serif !important; }}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {{ color: var(--tt-muted); }}
/* Sidebar navigation: identical left edge, width, padding and text alignment. */
[data-testid="stSidebar"] [data-testid="stButton"] {{ width: 100% !important; }}
[data-testid="stSidebar"] [data-testid="stButton"] > button {{
  width: 100% !important; min-height: 2.55rem; display: flex !important;
  align-items: center !important; justify-content: flex-start !important;
  text-align: left !important; padding: 0.55rem 0.85rem !important;
  margin: 0 !important; border-radius: 8px !important;
}}
[data-testid="stSidebar"] [data-testid="stButton"] > button div,
[data-testid="stSidebar"] [data-testid="stButton"] > button p {{
  width: 100% !important; text-align: left !important; margin-left: 0 !important;
}}
[data-testid="stSidebar"] button[kind="primary"] {{ background: var(--tt-soft) !important; color: var(--tt-accent) !important; border: 1px solid var(--tt-border) !important; font-weight: 600; }}
[data-testid="stSidebar"] button[kind="secondary"] {{ background: transparent !important; color: var(--tt-text) !important; border: 1px solid transparent !important; }}
[data-testid="stSidebar"] button[kind="secondary"]:hover {{ background: var(--tt-surface-alt) !important; border-color: var(--tt-border) !important; }}
[data-testid="stSidebar"] [data-testid="stButton"] button p {{ color: inherit !important; }}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2 {{ color: var(--tt-accent) !important; font-weight: 700; letter-spacing: -.04em; }}
h1, h2, h3, h4, h5, h6 {{ color: var(--tt-text) !important; font-family: 'Poppins', sans-serif !important; letter-spacing: -.035em; }}
p, li, label, span, div {{ font-family: 'Poppins', sans-serif; }}
[data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li, [data-testid="stCaptionContainer"] {{ color: var(--tt-text); }}
[data-testid="stCaptionContainer"], .stCaption {{ color: var(--tt-muted) !important; }}
small {{ color: var(--tt-muted) !important; }}
hr {{ border-color: var(--tt-border) !important; }}
/* Inputs and readable form labels */
[data-testid="stTextInput"] label, [data-testid="stTextArea"] label,
[data-testid="stSelectbox"] label, [data-testid="stDateInput"] label,
[data-testid="stFileUploader"] label, [data-testid="stMultiSelect"] label,
[data-testid="stNumberInput"] label, [data-testid="stCheckbox"] label,
[data-testid="stRadio"] label, [data-testid="stToggle"] label {{ color: var(--tt-text) !important; font-weight: 500; }}
input, textarea {{ background: var(--tt-input) !important; color: var(--tt-text) !important; border-color: var(--tt-border) !important; }}
[data-baseweb="select"] > div, [data-baseweb="input"] > div,
[data-baseweb="textarea"] > div {{ background: var(--tt-input) !important; border-color: var(--tt-border) !important; }}
[data-baseweb="select"] *, [data-baseweb="input"] *, [data-baseweb="textarea"] * {{ color: var(--tt-text) !important; }}
[data-baseweb="popover"], [data-baseweb="menu"], [role="listbox"] {{ background: var(--tt-surface) !important; }}
[role="option"] {{ color: var(--tt-text) !important; }}
[data-testid="stTextInput"] input::placeholder, textarea::placeholder {{ color: var(--tt-muted) !important; opacity: .8; }}
/* Buttons */
div.stButton > button, div.stFormSubmitButton > button {{
  border-radius: 8px !important; min-height: 2.55rem; font-family: 'Poppins', sans-serif !important;
  font-weight: 500; border: 1px solid var(--tt-border) !important;
  background: var(--tt-surface) !important; color: var(--tt-text) !important;
}}
div.stButton > button:hover, div.stFormSubmitButton > button:hover {{ border-color: var(--tt-accent) !important; color: var(--tt-accent) !important; }}
div.stButton > button[kind="primary"], div.stFormSubmitButton > button[kind="primary"] {{
  background: var(--tt-accent) !important; border-color: var(--tt-accent) !important; color: #FFFFFF !important;
}}
div.stButton > button[kind="primary"] p, div.stFormSubmitButton > button[kind="primary"] p {{ color: #FFFFFF !important; }}
/* Reports & Analytics export buttons */
[data-testid="stDownloadButton"] button {{
  background: #705563 !important; border: 1px solid #705563 !important;
  color: #FFFFFF !important; border-radius: 8px !important;
  font-family: 'Poppins', sans-serif !important; font-weight: 600 !important;
}}
[data-testid="stDownloadButton"] button p,
[data-testid="stDownloadButton"] button span {{ color: #FFFFFF !important; }}
[data-testid="stDownloadButton"] button:hover {{
  background: #4D0F28 !important; border-color: #4D0F28 !important;
  color: #FFFFFF !important;
}}
[data-testid="stDownloadButton"] button:hover p,
[data-testid="stDownloadButton"] button:hover span {{ color: #FFFFFF !important; }}
/* Panels, metrics, forms and tabs */
[data-testid="stVerticalBlockBorderWrapper"] {{ background: var(--tt-surface) !important; border: 1px solid var(--tt-border) !important; border-radius: 10px !important; box-shadow: var(--tt-shadow, none); }}
[data-testid="stMetric"] {{ background: var(--tt-surface) !important; border: 1px solid var(--tt-border) !important; border-radius: 10px; padding: 16px 18px; }}
[data-testid="stMetricLabel"] *, [data-testid="stMetricValue"] *, [data-testid="stMetricDelta"] * {{ color: var(--tt-text) !important; }}
[data-testid="stForm"] {{ background: var(--tt-surface) !important; border: 1px solid var(--tt-border) !important; border-radius: 10px; padding: 1.2rem; }}
[data-testid="stTabs"] [role="tab"] {{ color: var(--tt-muted) !important; }}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {{ color: var(--tt-accent) !important; }}
[data-testid="stAlert"] {{ border-radius: 9px; }}
.ticket-card {{ background: var(--tt-surface); border: 1px solid var(--tt-border); border-radius: 10px; padding: 15px 17px; margin-bottom: 10px; }}
.muted {{ color: var(--tt-muted) !important; font-size: .88rem; }}
.pill {{ display: inline-block; border-radius: 5px; padding: 3px 9px; background: var(--tt-soft); color: var(--tt-accent) !important; font-size: .76rem; font-weight: 600; }}
[data-testid="stFileUploader"] section {{ background: var(--tt-surface-alt) !important; border-color: var(--tt-border) !important; }}
[data-testid="stFileUploader"] section * {{ color: var(--tt-text) !important; }}
/* Envoyce-inspired VENTICK authentication palette and readable controls */
:root {{ color-scheme: {'dark' if False else 'light'}; }}
[data-testid="stTextInput"] input, [data-testid="stTextArea"] textarea,
[data-testid="stSelectbox"] [data-baseweb="select"] > div {{
  min-height:44px !important; border-radius:8px !important;
}}
[data-testid="stTextInput"] input:focus, [data-testid="stTextArea"] textarea:focus {{
  border-color:var(--tt-accent) !important; box-shadow:0 0 0 1px var(--tt-accent) !important;
}}
[data-testid="stForm"] label, [data-testid="stVerticalBlockBorderWrapper"] label {{
  color:var(--tt-text) !important; opacity:1 !important;
}}
[data-testid="stMarkdownContainer"] p, [data-testid="stCaptionContainer"] p {{ color:var(--tt-text); }}
/* Keep the sidebar open and remove the collapse affordance */
[data-testid="stSidebar"] [data-testid="collapsedControl"],
[data-testid="stSidebar"] button[kind="header"] {{ display: none !important; }}
@media (max-width: 800px) {{ [data-testid="stMainBlockContainer"] {{ padding: 1rem; }} }}
{LIGHT_MODE_CARD_CSS}
</style>
""", unsafe_allow_html=True)


# -----------------------------
# DATABASE
# -----------------------------
def db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def init_db():
    with db() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS companies (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            join_code TEXT UNIQUE NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS users (
            id TEXT PRIMARY KEY,
            company_id TEXT NOT NULL,
            full_name TEXT NOT NULL,
            email TEXT NOT NULL,
            password_hash BLOB NOT NULL,
            role TEXT NOT NULL DEFAULT 'Member',
            team TEXT NOT NULL DEFAULT 'Operations',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT NOT NULL,
            UNIQUE(company_id, email),
            FOREIGN KEY(company_id) REFERENCES companies(id)
        );
        CREATE TABLE IF NOT EXISTS tickets (
            id TEXT PRIMARY KEY,
            company_id TEXT NOT NULL,
            ticket_key TEXT NOT NULL,
            title TEXT NOT NULL,
            description TEXT NOT NULL,
            category TEXT NOT NULL,
            team TEXT NOT NULL,
            priority TEXT NOT NULL,
            status TEXT NOT NULL,
            reporter_id TEXT NOT NULL,
            assignee_id TEXT,
            due_date TEXT,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            resolved_at TEXT,
            UNIQUE(company_id, ticket_key),
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(reporter_id) REFERENCES users(id),
            FOREIGN KEY(assignee_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS comments (
            id TEXT PRIMARY KEY,
            ticket_id TEXT NOT NULL,
            user_id TEXT NOT NULL,
            body TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(ticket_id) REFERENCES tickets(id) ON DELETE CASCADE,
            FOREIGN KEY(user_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS notifications (
            id TEXT PRIMARY KEY,
            company_id TEXT NOT NULL,
            user_id TEXT NOT NULL,
            ticket_id TEXT,
            message TEXT NOT NULL,
            is_read INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            FOREIGN KEY(company_id) REFERENCES companies(id),
            FOREIGN KEY(user_id) REFERENCES users(id),
            FOREIGN KEY(ticket_id) REFERENCES tickets(id) ON DELETE CASCADE
        );
        CREATE TABLE IF NOT EXISTS activity (
            id TEXT PRIMARY KEY,
            ticket_id TEXT NOT NULL,
            user_id TEXT NOT NULL,
            action TEXT NOT NULL,
            created_at TEXT NOT NULL,
            FOREIGN KEY(ticket_id) REFERENCES tickets(id) ON DELETE CASCADE,
            FOREIGN KEY(user_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS teams (
            id TEXT PRIMARY KEY,
            company_id TEXT NOT NULL,
            name TEXT NOT NULL,
            UNIQUE(company_id, name),
            FOREIGN KEY(company_id) REFERENCES companies(id)
        );
        CREATE TABLE IF NOT EXISTS password_reset_tokens (
            id TEXT PRIMARY KEY,
            user_id TEXT NOT NULL,
            token_hash TEXT NOT NULL UNIQUE,
            expires_at TEXT NOT NULL,
            used_at TEXT,
            created_at TEXT NOT NULL,
            FOREIGN KEY(user_id) REFERENCES users(id) ON DELETE CASCADE
        );
        """)
        con.commit()


def q_one(sql, params=()):
    with db() as con:
        return con.execute(sql, params).fetchone()


def q_all(sql, params=()):
    with db() as con:
        return con.execute(sql, params).fetchall()


def execute(sql, params=()):
    with db() as con:
        con.execute(sql, params)
        con.commit()


def password_hash(password):
    return bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt())


def verify_password(password, hashed):
    try:
        return bcrypt.checkpw(password.encode("utf-8"), hashed)
    except Exception:
        return False


# -----------------------------
# NOTIFICATIONS
# -----------------------------
def _setting(name, default=""):
    """Read deployment secrets first, then local environment variables."""
    try:
        value = st.secrets.get(name, None)
        if value not in (None, ""):
            return str(value)
    except Exception:
        pass
    return os.getenv(name, default)


def app_base_url():
    return _setting("TEAMTRACK_APP_URL", "https://teamtrack.streamlit.app/").rstrip("/")


def send_email(to_email, subject, body):
    """Send email when SMTP credentials are configured in secrets or environment."""
    host = _setting("TEAMTRACK_SMTP_HOST")
    port = int(_setting("TEAMTRACK_SMTP_PORT", "587"))
    username = _setting("TEAMTRACK_SMTP_USERNAME")
    password = _setting("TEAMTRACK_SMTP_PASSWORD")
    sender = _setting("TEAMTRACK_FROM_EMAIL", username or "")
    if not all([host, username, password, sender, to_email]):
        return False
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = to_email
    msg.set_content(body)
    try:
        with smtplib.SMTP(host, port, timeout=15) as server:
            server.starttls()
            server.login(username, password)
            server.send_message(msg)
        return True
    except Exception:
        return False


def send_registration_emails(company_id, user_id, is_company_admin=False):
    """Send registration confirmation to the new user and notify the company contact."""
    user = q_one("SELECT full_name, email, team FROM users WHERE id=? AND company_id=?", (user_id, company_id))
    company = q_one("SELECT name FROM companies WHERE id=?", (company_id,))
    if not user or not company:
        return {"user": False, "company": False}
    user_sent = send_email(
        user["email"],
        "Your TeamTrack registration is confirmed",
        f"Hello {user['full_name']},\n\nYour TeamTrack account for {company['name']} has been created. "
        "You can now sign in to your workspace.\n\nTeamTrack"
    )
    if is_company_admin:
        # The company registration contact is the first Admin's registered email.
        company_sent = user_sent
    else:
        admin = q_one("SELECT email, full_name FROM users WHERE company_id=? AND role='Admin' AND active=1 ORDER BY created_at LIMIT 1", (company_id,))
        company_sent = False
        if admin:
            company_sent = send_email(
                admin["email"],
                f"New TeamTrack registration — {company['name']}",
                f"Hello {admin['full_name']},\n\n{user['full_name']} ({user['email']}) has registered in "
                f"the {company['name']} workspace for the {user['team']} department.\n\nTeamTrack"
            )
    return {"user": user_sent, "company": company_sent}


def create_password_reset(user_email, company_code=""):
    sql = "SELECT u.id, u.full_name, u.email, c.name AS company_name, c.join_code FROM users u JOIN companies c ON c.id=u.company_id WHERE lower(u.email)=lower(?) AND u.active=1"
    params = [user_email.strip().lower()]
    if company_code.strip():
        sql += " AND upper(c.join_code)=upper(?)"
        params.append(company_code.strip())
    matches = q_all(sql, tuple(params))
    # Avoid exposing whether an email is registered.
    if len(matches) != 1:
        return False
    user = matches[0]
    raw_token = uuid.uuid4().hex + uuid.uuid4().hex
    token_hash = __import__("hashlib").sha256(raw_token.encode()).hexdigest()
    expires = (datetime.now(timezone.utc) + timedelta(minutes=45)).isoformat(timespec="seconds")
    execute(
        "INSERT INTO password_reset_tokens (id,user_id,token_hash,expires_at,created_at) VALUES (?,?,?,?,?)",
        (str(uuid.uuid4()), user["id"], token_hash, expires, now())
    )
    reset_link = f"{app_base_url()}/?reset={raw_token}"
    sent = send_email(
        user["email"], "Reset your TeamTrack password",
        f"Hello {user['full_name']},\n\nUse the link below to choose a new password. This link expires in 45 minutes and can only be used once.\n\n{reset_link}\n\nIf you did not request this, you can ignore this email.\n\nTeamTrack"
    )
    return sent


def complete_password_reset(raw_token, new_password):
    token_hash = __import__("hashlib").sha256(raw_token.encode()).hexdigest()
    record = q_one(
        "SELECT id,user_id,expires_at,used_at FROM password_reset_tokens WHERE token_hash=?",
        (token_hash,)
    )
    if not record or record["used_at"] or record["expires_at"] < now():
        return False
    with db() as con:
        con.execute("UPDATE users SET password_hash=? WHERE id=?", (password_hash(new_password), record["user_id"]))
        con.execute("UPDATE password_reset_tokens SET used_at=? WHERE id=?", (now(), record["id"]))
        con.execute("UPDATE password_reset_tokens SET used_at=? WHERE user_id=? AND id<>? AND used_at IS NULL", (now(), record["user_id"], record["id"]))
        con.commit()
    return True


def notify_users(company_id, user_ids, ticket_id, message, subject="TeamTrack update"):
    unique_ids = list(dict.fromkeys([x for x in user_ids if x]))
    with db() as con:
        for uid in unique_ids:
            con.execute(
                """INSERT INTO notifications
                   (id, company_id, user_id, ticket_id, message, created_at)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (str(uuid.uuid4()), company_id, uid, ticket_id, message, now())
            )
        con.commit()
    # Email is best-effort and does not affect in-app notification creation.
    for uid in unique_ids:
        user = q_one("SELECT email, full_name FROM users WHERE id=? AND company_id=?", (uid, company_id))
        if user:
            send_email(
                user["email"],
                subject,
                f"Hello {user['full_name']},\n\n{message}\n\nOpen TeamTrack to view the ticket."
            )


def ticket_notification_recipients(company_id, team, reporter_id=None, assignee_id=None, exclude_user_id=None):
    """Notify only the ticket's department and the people directly involved."""
    rows = q_all(
        "SELECT id FROM users WHERE company_id=? AND team=? AND active=1",
        (company_id, team)
    )
    recipients = {row["id"] for row in rows}
    if reporter_id:
        recipients.add(reporter_id)
    if assignee_id:
        recipients.add(assignee_id)
    if exclude_user_id:
        recipients.discard(exclude_user_id)
    return sorted(recipients)


def add_activity(ticket_id, user_id, action):
    execute(
        "INSERT INTO activity (id, ticket_id, user_id, action, created_at) VALUES (?, ?, ?, ?, ?)",
        (str(uuid.uuid4()), ticket_id, user_id, action, now())
    )


def get_company_users(company_id):
    return q_all(
        "SELECT id, full_name, email, role, team FROM users WHERE company_id=? AND active=1 ORDER BY full_name",
        (company_id,)
    )


def get_ticket(ticket_id, company_id):
    return q_one(
        """SELECT t.*, r.full_name AS reporter_name, a.full_name AS assignee_name
           FROM tickets t
           JOIN users r ON r.id=t.reporter_id
           LEFT JOIN users a ON a.id=t.assignee_id
           WHERE t.id=? AND t.company_id=?""",
        (ticket_id, company_id)
    )


# -----------------------------
# AUTH / WORKSPACE SETUP
# -----------------------------
def create_company(company_name, full_name, email, password, team):
    company_id = str(uuid.uuid4())
    user_id = str(uuid.uuid4())
    join_code = uuid.uuid4().hex[:8].upper()
    clean_email = email.strip().lower()
    if q_one("SELECT id FROM companies WHERE lower(name)=lower(?)", (company_name.strip(),)):
        return False, "A company with this name already exists in this local MVP. Try a more specific workspace name."
    try:
        with db() as con:
            con.execute(
                "INSERT INTO companies VALUES (?, ?, ?, ?)",
                (company_id, company_name.strip(), join_code, now())
            )
            con.execute(
                """INSERT INTO users
                   (id, company_id, full_name, email, password_hash, role, team, created_at)
                   VALUES (?, ?, ?, ?, ?, 'Admin', ?, ?)""",
                (user_id, company_id, full_name.strip(), clean_email, password_hash(password), team, now())
            )
            for name in DEFAULT_TEAMS:
                con.execute("INSERT OR IGNORE INTO teams (id, company_id, name) VALUES (?, ?, ?)",
                            (str(uuid.uuid4()), company_id, name))
            con.commit()
        email_status = send_registration_emails(company_id, user_id, is_company_admin=True)
        return True, (company_id, user_id, join_code, email_status)
    except sqlite3.IntegrityError:
        return False, "This email is already registered in that workspace."


def join_company(join_code, full_name, email, password, team):
    company = q_one("SELECT * FROM companies WHERE upper(join_code)=upper(?)", (join_code.strip(),))
    if not company:
        return False, "That company invitation code was not found."
    try:
        user_id = str(uuid.uuid4())
        execute(
            """INSERT INTO users
               (id, company_id, full_name, email, password_hash, role, team, created_at)
               VALUES (?, ?, ?, ?, ?, 'Member', ?, ?)""",
            (user_id, company["id"], full_name.strip(), email.strip().lower(),
             password_hash(password), team, now())
        )
        email_status = send_registration_emails(company["id"], user_id, is_company_admin=False)
        return True, (company["id"], user_id, company["name"], email_status)
    except sqlite3.IntegrityError:
        return False, "This email is already registered in that company."


def login(email, password, company_code=""):
    params = [email.strip().lower()]
    sql = """SELECT u.*, c.name AS company_name, c.join_code
             FROM users u JOIN companies c ON c.id=u.company_id
             WHERE lower(u.email)=lower(?) AND u.active=1"""
    if company_code.strip():
        sql += " AND upper(c.join_code)=upper(?)"
        params.append(company_code.strip())
    rows = q_all(sql, tuple(params))
    for user in rows:
        if verify_password(password, user["password_hash"]):
            return user
    return None


def logout():
    for key in ["user", "page", "selected_ticket"]:
        st.session_state.pop(key, None)
    st.rerun()


# -----------------------------
# APP HELPERS
# -----------------------------
def is_admin(user):
    return user["role"] == "Admin"


def is_it(user):
    return user["role"] == "IT Personnel"


def can_assign_across_departments(user):
    return is_admin(user) or is_it(user)


def can_assign_in_department(user, department):
    if can_assign_across_departments(user):
        return True
    return user["role"] == "Team Lead" and user["team"] == department


def can_manage(user):
    return is_admin(user) or is_it(user) or user["role"] == "Team Lead"


def can_manage_members(user):
    return is_admin(user)


def ticket_key(company_id):
    count = q_one("SELECT COUNT(*) AS n FROM tickets WHERE company_id=?", (company_id,))["n"] + 1
    return f"TT-{count:04d}"


def status_badge(status):
    colors = {
        "Open": "#e9efff", "Assigned": "#e9efff", "In Progress": "#fff2d8",
        "On Hold": "#f1eaff", "Resolved": "#e4f7ed", "Closed": "#eef0f3"
    }
    return f'<span style="background:{colors.get(status, "#eee")};color:#4D0F28;padding:4px 9px;border-radius:20px;font-size:12px;font-weight:700">{status}</span>'


def priority_badge(priority):
    colors = {"Critical": "#ffe5e5", "High": "#fff0e2", "Medium": "#fff8d9", "Low": "#e8f3ff"}
    return f'<span style="background:{colors.get(priority, "#eee")};color:#4D0F28;padding:4px 9px;border-radius:20px;font-size:12px;font-weight:700">{priority}</span>' 

def ticket_rows(user):
    return q_all(
        """SELECT t.*, r.full_name AS reporter_name, a.full_name AS assignee_name
           FROM tickets t JOIN users r ON r.id=t.reporter_id
           LEFT JOIN users a ON a.id=t.assignee_id
           WHERE t.company_id=? ORDER BY
           CASE t.priority WHEN 'Critical' THEN 1 WHEN 'High' THEN 2 WHEN 'Medium' THEN 3 ELSE 4 END,
           t.updated_at DESC""",
        (user["company_id"],)
    )



REPORT_COLUMNS = [
    "Ticket ID", "Title", "Description", "Category", "Department", "Priority",
    "Status", "Reported By", "Assigned To", "Due Date", "Created At",
    "Last Updated", "Resolved At",
]


def build_report_dataframe(tickets):
    """Convert company-scoped ticket rows into a stable export dataframe.

    Supplying columns explicitly is important: an empty result must still have
    the same schema so PDF, Word, Excel and CSV exports can handle zero matches.
    """
    records = []
    for t in tickets:
        records.append({
            "Ticket ID": t["ticket_key"],
            "Title": t["title"],
            "Description": t["description"],
            "Category": t["category"],
            "Department": t["team"],
            "Priority": t["priority"],
            "Status": t["status"],
            "Reported By": t["reporter_name"],
            "Assigned To": t["assignee_name"] or "Unassigned",
            "Due Date": t["due_date"] or "",
            "Created At": t["created_at"],
            "Last Updated": t["updated_at"],
            "Resolved At": t["resolved_at"] or "",
        })
    return pd.DataFrame(records, columns=REPORT_COLUMNS)


def make_excel_report(df, summary, company_name, generated_at):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, index=False, sheet_name="Summary")
        df.to_excel(writer, index=False, sheet_name="Ticket details")
        for sheet_name in ["Summary", "Ticket details"]:
            ws = writer.book[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="4D0F28")
                cell.alignment = Alignment(vertical="center")
            for column_cells in ws.columns:
                letter = column_cells[0].column_letter
                max_length = max((len(str(c.value or "")) for c in column_cells), default=10)
                ws.column_dimensions[letter].width = min(max(max_length + 2, 12), 42)
        ws = writer.book["Summary"]
        ws.insert_rows(1, 3)
        ws["A1"] = company_name
        ws["A1"].font = Font(size=16, bold=True, color="4D0F28")
        ws["A2"] = "TeamTrack report"
        ws["A3"] = f"Generated: {generated_at}"
    output.seek(0)
    return output.getvalue()


def make_word_report(df, summary, company_name, generated_at, period_label):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    title = doc.add_heading(company_name, 0)
    title.runs[0].font.color.rgb = RGBColor(77, 15, 40)
    doc.add_heading("TeamTrack Reports & Analytics", level=1)
    doc.add_paragraph(f"Reporting period: {period_label}")
    doc.add_paragraph(f"Generated: {generated_at}")
    doc.add_heading("Summary", level=2)
    table = doc.add_table(rows=1, cols=2)
    table.style = "Light Shading Accent 1"
    table.rows[0].cells[0].text = "Metric"
    table.rows[0].cells[1].text = "Value"
    for _, row in summary.iterrows():
        cells = table.add_row().cells
        cells[0].text = str(row["Metric"])
        cells[1].text = str(row["Value"])
    doc.add_heading("Ticket details", level=2)
    if df.empty:
        doc.add_paragraph("No tickets match the selected filters.")
    else:
        cols = ["Ticket ID", "Title", "Department", "Priority", "Status", "Assigned To", "Created At"]
        table = doc.add_table(rows=1, cols=len(cols))
        table.style = "Table Grid"
        for i, col in enumerate(cols):
            table.rows[0].cells[i].text = col
        for _, row in df[cols].iterrows():
            cells = table.add_row().cells
            for i, col in enumerate(cols):
                cells[i].text = str(row[col])[:180]
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


def make_pdf_report(df, summary, company_name, generated_at, period_label):
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4), rightMargin=28, leftMargin=28,
                            topMargin=28, bottomMargin=28)
    styles = getSampleStyleSheet()
    story = [Paragraph(company_name, styles["Title"]),
             Paragraph("TeamTrack Reports &amp; Analytics", styles["Heading1"]),
             Paragraph(f"Reporting period: {period_label}", styles["Normal"]),
             Paragraph(f"Generated: {generated_at}", styles["Normal"]), Spacer(1, 12),
             Paragraph("Summary", styles["Heading2"])]
    summary_data = [["Metric", "Value"]] + [[str(r["Metric"]), str(r["Value"])] for _, r in summary.iterrows()]
    summary_table = Table(summary_data, colWidths=[220, 100], repeatRows=1)
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4D0F28")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D4D7")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F4F5")]),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story.extend([summary_table, Spacer(1, 14), Paragraph("Ticket details", styles["Heading2"])])
    cols = ["Ticket ID", "Title", "Department", "Priority", "Status", "Assigned To", "Created At"]
    data = [cols]
    if not df.empty:
        # Reindex defensively so an incomplete dataframe cannot break PDF export.
        pdf_df = df.reindex(columns=cols, fill_value="")
        for _, row in pdf_df.iterrows():
            data.append([str(row[c])[:70] for c in cols])
    if len(data) == 1:
        data.append(["No tickets match the selected filters."] + [""] * (len(cols)-1))
    widths = [62, 190, 85, 65, 75, 105, 105]
    ticket_table = Table(data, colWidths=widths, repeatRows=1)
    ticket_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4D0F28")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9D4D7")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F4F5")]),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("PADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(ticket_table)
    doc.build(story)
    output.seek(0)
    return output.getvalue()

def open_ticket(ticket_id):
    st.session_state.selected_ticket = ticket_id
    st.session_state.page = "Ticket details"
    st.rerun()


# -----------------------------
# LOGIN / COMPANY ONBOARDING
# -----------------------------
init_db()

if "user" not in st.session_state:
    st.session_state.user = None

# A password-reset link opens a dedicated form before the normal sign-in view.
reset_token = st.query_params.get("reset", "")
if reset_token and not st.session_state.user:
    st.markdown("<div style='max-width:460px;margin:8vh auto 0;text-align:center'><h2>Choose a new password</h2><p>Use at least 8 characters.</p></div>", unsafe_allow_html=True)
    with st.container(border=True):
        with st.form("complete_password_reset_form"):
            new_password = st.text_input("New password", type="password")
            confirm_new_password = st.text_input("Confirm new password", type="password")
            reset_submit = st.form_submit_button("Update password", type="primary", use_container_width=True)
        if reset_submit:
            if len(new_password) < 8:
                st.error("Use a password with at least 8 characters.")
            elif new_password != confirm_new_password:
                st.error("The passwords do not match.")
            elif complete_password_reset(reset_token, new_password):
                st.query_params.clear()
                st.session_state.teamtrack_password_reset_success = True
                st.session_state.teamtrack_public_page = "auth"
                st.session_state.teamtrack_auth_mode = "signin"
                st.rerun()
            else:
                st.error("This reset link is invalid, expired or has already been used. Request a new link.")
    st.stop()

if not st.session_state.user:
    if "teamtrack_public_page" not in st.session_state:
        st.session_state.teamtrack_public_page = "landing"
    if "teamtrack_auth_mode" not in st.session_state:
        st.session_state.teamtrack_auth_mode = "signin"

    # Public landing page, replicated from the supplied TeamTrack design.
    if st.session_state.teamtrack_public_page == "landing":
        st.markdown("""
        <style>
        [data-testid="stSidebar"] {display:none !important;}
        [data-testid="stHeader"] {display:none !important;}
        [data-testid="stMainBlockContainer"] {max-width:1440px !important; min-height:100vh !important; padding:0 0 76px !important; box-sizing:border-box !important;}
        [data-testid="stAppViewContainer"], .stApp, [data-testid="stMain"] {background:#FFFFFF !important;}
        .tt-landing-top {height:90px; display:flex; align-items:center; padding-left:32px; border-bottom:1px solid #E7DDE1; box-sizing:border-box;}
        .tt-landing-copy {text-align:center; padding:76px 20px 0; color:#080808;}
        .tt-landing-title {font-size:35px; line-height:1.48; font-weight:700; letter-spacing:-1.45px; margin:0 auto 10px; color:#080808; max-width:850px;}
        .tt-landing-subtitle {font-size:16px; line-height:1.55; color:#777174; margin:0 auto; max-width:720px; font-weight:400;}
        .st-key-landing_get_started, .st-key-landing_sign_in {margin-top:24px !important;}
        .tt-landing-illustration {display:flex; justify-content:center; margin:34px auto 0; width:100%;}
        .tt-landing-illustration img {width:min(470px, 82vw); height:auto; object-fit:contain;}
        .tt-social-proof {display:flex; justify-content:center; align-items:center; gap:12px; margin:24px auto 34px; color:#A19A9E; font-size:13px; font-weight:400;}
        .tt-social-avatars {display:flex; align-items:center; padding-left:4px;}
        .tt-social-avatar {width:30px; height:30px; border-radius:50%; display:block; overflow:hidden; border:2px solid #fff; margin-left:-6px; background:#F0DDE5; box-shadow:0 1px 3px rgba(40,20,28,.12);}
        .tt-social-avatar img {width:100%; height:100%; display:block; object-fit:cover;}
        .tt-social-avatar:nth-child(2) {background:#E8D9CE;}
        .tt-social-avatar:nth-child(3) {background:#D8E4E1;}
        .tt-landing-footer-wrap {position:fixed; z-index:1000; left:0; right:0; bottom:0; width:100%; min-height:58px; display:flex; justify-content:center; align-items:center; gap:14px; padding:14px 18px 16px; margin:0; background:#FFFFFF; border-top:1px solid #F0EAED; box-sizing:border-box;}
        .tt-landing-footer {text-align:center; color:#777174; font-size:11px; margin:0;}
        .tt-landing-legal {display:flex; align-items:center; gap:14px; font-size:11px; color:#4D0F28;}
        .tt-landing-legal span {display:inline-block; color:#4D0F28; padding:0; margin:0;}
        .st-key-landing_get_started button {background:#4D0F28 !important; border:1px solid #4D0F28 !important; color:#FFFFFF !important; border-radius:22px !important; min-height:48px !important; font-size:17px !important; font-weight:500 !important;}
        .st-key-landing_get_started button p {color:#FFFFFF !important;}
        .st-key-landing_sign_in button {background:#FFFFFF !important; border:1px solid #4D0F28 !important; color:#4D0F28 !important; border-radius:22px !important; min-height:48px !important; font-size:17px !important; font-weight:500 !important;}
        .st-key-landing_sign_in button p {color:#4D0F28 !important;}
        @media(max-width:700px) {
          .tt-landing-top {height:72px; padding-left:18px;}
          .tt-landing-copy {padding:54px 18px 0;}
          .tt-landing-title {font-size:26px; line-height:1.35; letter-spacing:-.8px;}
          .tt-landing-subtitle {font-size:14px;}
          .st-key-landing_get_started, .st-key-landing_sign_in {margin-top:20px !important;}
          .tt-landing-illustration {margin-top:32px;}
          .tt-landing-illustration img {width:min(390px, 88vw);}
          .tt-social-proof {margin:20px auto 28px; font-size:12px; gap:9px;}
          .tt-social-avatar {width:27px; height:27px;}
          .tt-landing-footer-wrap {gap:10px; flex-wrap:wrap; padding:12px 12px 14px; min-height:58px;}
          .tt-landing-footer, .tt-landing-legal {font-size:10px;}
          [data-testid="stMainBlockContainer"] {padding:0 0 12px !important;}
        }
        </style>
        """, unsafe_allow_html=True)

        if LOGO_PATH.exists():
            logo_data = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
            logo_markup = f"<img src='data:image/png;base64,{logo_data}' alt='TeamTrack' style='width:54px;height:auto;display:block'>"
        else:
            logo_markup = "<strong style='color:#4D0F28'>TeamTrack</strong>"
        st.markdown(f"<div class='tt-landing-top'>{logo_markup}</div>", unsafe_allow_html=True)
        st.markdown("""
        <div class='tt-landing-copy'>
          <div class='tt-landing-title'>Every task has a place.<br>Every team has a clear direction.</div>
          <div class='tt-landing-subtitle'>From raising requests to tracking resolutions,<br>TeamTrack helps teams stay organised, connected and accountable.</div>
        </div>
        """, unsafe_allow_html=True)

        button_spacer_l, button_area, button_spacer_r = st.columns([1.45, 1.65, 1.45])
        with button_area:
            start_col, signin_col = st.columns(2, gap="large")
            with start_col:
                if st.button("Get Started", key="landing_get_started", type="primary", use_container_width=True):
                    st.session_state.teamtrack_auth_mode = "create"
                    st.session_state.teamtrack_public_page = "auth"
                    st.rerun()
            with signin_col:
                if st.button("Sign In", key="landing_sign_in", use_container_width=True):
                    st.session_state.teamtrack_auth_mode = "signin"
                    st.session_state.teamtrack_public_page = "auth"
                    st.rerun()

        illustration_path = APP_DIR / "landing_illustration.png"
        if illustration_path.exists():
            illustration_data = base64.b64encode(illustration_path.read_bytes()).decode("ascii")
            st.markdown(f"<div class='tt-landing-illustration'><img src='data:image/png;base64,{illustration_data}' alt='A team member organising work at a desk'></div>", unsafe_allow_html=True)
        # Small illustrated profile portraits replace generic person symbols.
        avatar_svgs = [
            "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' fill='#E9D8DE'/><path d='M9 64c2-15 10-23 23-23s21 8 23 23' fill='#4D0F28'/><circle cx='32' cy='25' r='13' fill='#B97858'/><path d='M18 25c0-13 7-19 15-19 10 0 15 8 14 20-4-3-8-8-9-12-5 6-12 9-20 11' fill='#2D1B1B'/></svg>",
            "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' fill='#E8E0D7'/><path d='M8 64c2-15 10-23 24-23s22 8 24 23' fill='#566B70'/><circle cx='32' cy='25' r='13' fill='#D59A72'/><path d='M18 23c1-12 7-18 15-18 9 0 14 7 14 18-5-2-10-6-13-11-4 6-9 9-16 11' fill='#38251E'/></svg>",
            "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'><rect width='64' height='64' fill='#DCE7E3'/><path d='M8 64c2-15 10-23 24-23s22 8 24 23' fill='#92718A'/><circle cx='32' cy='25' r='13' fill='#8D573F'/><path d='M17 27c-1-14 5-22 15-22s16 8 15 22l-4-2-2-11c-7 7-14 10-24 11' fill='#171616'/></svg>"
        ]
        avatar_markup = "".join(
            "<span class='tt-social-avatar'><img alt='TeamTrack community member' src='data:image/svg+xml;base64," +
            base64.b64encode(svg.encode('utf-8')).decode('ascii') + "'></span>"
            for svg in avatar_svgs
        )
        st.markdown(
            "<div class='tt-social-proof'><div class='tt-social-avatars'>" + avatar_markup +
            "</div><span>Used by 200+ companies</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown("<div class='tt-landing-footer-wrap'><div class='tt-landing-footer'>© 2026 TeamTrack. All rights reserved.</div><div class='tt-landing-legal'><span>Terms</span><span>Privacy</span></div></div>", unsafe_allow_html=True)
        st.stop()

    # Keep the authentication experience visually consistent with the
    # Envoyce-inspired VENTICK interface: centred wordmark, compact white
    # card, clear segmented navigation, generous spacing and Poppins.
    st.markdown("""
    <style>
    [data-testid="stSidebar"] {display:none !important;}
    [data-testid="stMainBlockContainer"] {max-width:100% !important; padding:1.2rem 1.5rem 3rem !important;}
    [data-testid="stHeader"] {background:transparent !important;}
    .st-key-teamtrack_auth_card [data-testid="stVerticalBlockBorderWrapper"] {
        background:var(--tt-surface) !important;
        border:1px solid var(--tt-border) !important;
        border-radius:18px !important;
        padding:28px 30px 26px !important;
        box-shadow:var(--tt-shadow) !important;
    }
    .teamtrack-auth-brand {display:flex; align-items:center; min-height:48px; padding:2px 0 6px;}
    .teamtrack-auth-brand img {display:block; width:40px; max-width:40px; height:auto; object-fit:contain;}
    .teamtrack-workspace-wordmark {font-size:25px; line-height:1.2; font-weight:800; letter-spacing:-1.05px; color:var(--tt-accent) !important; margin:2px 0 14px; padding:4px 2px;}
    .teamtrack-auth-page-wordmark {font-size:27px; line-height:1.2; font-weight:800; letter-spacing:-1.1px; color:var(--tt-accent) !important; text-align:center; margin:0 0 14px; padding:2px 0;}
    .teamtrack-auth-heading {text-align:center; margin:5px 0 20px;}
    .teamtrack-auth-heading h2 {font-size:24px; margin:0 0 6px; letter-spacing:-.8px;}
    .teamtrack-auth-heading p {font-size:13px; color:var(--tt-muted) !important; margin:0;}
    .teamtrack-auth-foot {text-align:center; color:var(--tt-muted); font-size:12px; margin-top:18px;}
    .st-key-teamtrack_auth_card [data-testid="stTextInput"] label,
    .st-key-teamtrack_auth_card [data-testid="stSelectbox"] label {color:var(--tt-text) !important; font-weight:500 !important; font-size:13px !important;}
    .st-key-teamtrack_auth_card input {min-height:44px !important; border-radius:8px !important;}
    .st-key-teamtrack_auth_card [data-testid="stForm"] {border:0 !important; padding:8px 0 0 !important; background:transparent !important;}
    .st-key-teamtrack_auth_card [data-testid="stFormSubmitButton"] button {min-height:46px !important; border-radius:8px !important; font-weight:600 !important;}
    .st-key-teamtrack_auth_card [data-testid="stButton"] button {min-height:42px !important; border-radius:8px !important; font-size:13px !important;}
    .st-key-teamtrack_auth_card [data-testid="stCaptionContainer"] p {color:var(--tt-muted) !important;}
    .st-key-teamtrack_auth_card [data-testid="stAlert"] {border-radius:9px !important;}
    @media(max-width:700px) {
      .teamtrack-auth-brand {min-height:44px; padding-top:0;}
      .teamtrack-auth-brand img {width:36px; max-width:36px;}
      .st-key-teamtrack_auth_card [data-testid="stVerticalBlockBorderWrapper"] {padding:20px 16px !important;}
    }
    </style>
    """, unsafe_allow_html=True)

    # Small transparent brand mark at the top-left; theme control stays top-right.
    brand_col, theme_control = st.columns([5, 1], vertical_alignment="center")
    with brand_col:
        if LOGO_PATH.exists():
            st.image(str(LOGO_PATH), width=40)
        else:
            st.markdown("<div class='teamtrack-workspace-wordmark'>TeamTrack</div>", unsafe_allow_html=True)
    with theme_control:
        theme_toggle = st.toggle(
            "Dark mode",
            value=st.session_state.teamtrack_theme == "Dark",
            key="teamtrack_auth_dark_toggle",
        )
    selected_theme = "Dark" if theme_toggle else "Light"
    if selected_theme != st.session_state.teamtrack_theme:
        st.session_state.teamtrack_theme = selected_theme
        st.rerun()

    left, centre, right = st.columns([1.1, 1.45, 1.1])
    with centre:
        st.markdown("<div class='teamtrack-auth-page-wordmark'>TeamTrack</div>", unsafe_allow_html=True)
        with st.container(border=True, key="teamtrack_auth_card"):
            mode = st.session_state.get("teamtrack_auth_mode", "signin")
            nav1, nav2, nav3 = st.columns(3, gap="small")
            with nav1:
                if st.button("Sign in", key="auth_nav_signin", use_container_width=True,
                             type="primary" if mode == "signin" else "secondary"):
                    st.session_state.teamtrack_auth_mode = "signin"
                    st.rerun()
            with nav2:
                if st.button("Create company", key="auth_nav_create", use_container_width=True,
                             type="primary" if mode == "create" else "secondary"):
                    st.session_state.teamtrack_auth_mode = "create"
                    st.rerun()
            with nav3:
                if st.button("Join company", key="auth_nav_join", use_container_width=True,
                             type="primary" if mode == "join" else "secondary"):
                    st.session_state.teamtrack_auth_mode = "join"
                    st.rerun()

            if mode == "signin":
                st.markdown("""
                <div class="teamtrack-auth-heading">
                  <h2>Welcome back</h2>
                  <p>Sign in to access your company workspace.</p>
                </div>
                """, unsafe_allow_html=True)
                with st.form("login_form", clear_on_submit=False):
                    email = st.text_input("Work email", placeholder="you@company.com")
                    company_code = st.text_input(
                        "Company code", placeholder="Optional if your email is unique",
                        help="Use your company code if this email belongs to more than one workspace."
                    )
                    password = st.text_input("Password", type="password", placeholder="Enter your password")
                    submitted = st.form_submit_button("Sign in", type="primary", use_container_width=True)
                if submitted:
                    user = login(email, password, company_code)
                    if user:
                        st.session_state.user = dict(user)
                        st.session_state.page = "Overview"
                        st.rerun()
                    else:
                        st.error("We couldn't verify those login details. Check your email, password and company code.")
                if st.button("Forgot password?", key="forgot_password_link"):
                    st.session_state.teamtrack_auth_mode = "forgot"
                    st.rerun()

            elif mode == "forgot":
                st.markdown("""
                <div class="teamtrack-auth-heading">
                  <h2>Reset your password</h2>
                  <p>Enter your account email. If it is registered, we will send a reset link.</p>
                </div>
                """, unsafe_allow_html=True)
                with st.form("forgot_password_form", clear_on_submit=False):
                    reset_email = st.text_input("Account email", placeholder="you@company.com")
                    reset_company_code = st.text_input("Company code (if needed)", placeholder="Optional")
                    reset_submitted = st.form_submit_button("Send reset link", type="primary", use_container_width=True)
                if reset_submitted:
                    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", reset_email.strip()):
                        st.error("Enter a valid email address.")
                    else:
                        sent = create_password_reset(reset_email, reset_company_code)
                        if sent:
                            st.success("If the account exists and email delivery is configured, a password reset link has been sent.")
                        else:
                            st.info("If the account exists and email delivery is configured, a password reset link will be sent. Check your email shortly.")
                if st.button("Back to sign in", key="forgot_back_signin"):
                    st.session_state.teamtrack_auth_mode = "signin"
                    st.rerun()

            elif mode == "create":
                st.markdown("""
                <div class="teamtrack-auth-heading">
                  <h2>Create your company workspace</h2>
                  <p>Set up your organisation and invite your team to join.</p>
                </div>
                """, unsafe_allow_html=True)
                with st.form("create_company_form", clear_on_submit=False):
                    company_name = st.text_input("Company or organisation name", placeholder="e.g. Acme Technologies")
                    owner_name = st.text_input("Your full name", placeholder="Enter your full name")
                    owner_email = st.text_input("Work email", placeholder="you@company.com")
                    owner_password = st.text_input("Create password", type="password", placeholder="At least 8 characters")
                    owner_password2 = st.text_input("Confirm password", type="password", placeholder="Enter password again")
                    team = st.selectbox("Your department", DEFAULT_TEAMS)
                    submitted = st.form_submit_button("Create company workspace", type="primary", use_container_width=True)
                if submitted:
                    if not all([company_name.strip(), owner_name.strip(), owner_email.strip(), owner_password]):
                        st.error("Complete all required fields.")
                    elif not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", owner_email.strip()):
                        st.error("Enter a valid email address.")
                    elif len(owner_password) < 8:
                        st.error("Use a password with at least 8 characters.")
                    elif owner_password != owner_password2:
                        st.error("The passwords do not match.")
                    else:
                        ok, result = create_company(company_name, owner_name, owner_email, owner_password, team)
                        if ok:
                            company_id, user_id, code, email_status = result
                            st.session_state.teamtrack_created_join_code = code
                            st.session_state.teamtrack_registration_email_status = email_status
                            st.session_state.user = dict(q_one(
                                """SELECT u.*, c.name AS company_name, c.join_code
                                   FROM users u JOIN companies c ON c.id=u.company_id WHERE u.id=?""",
                                (user_id,)
                            ))
                            st.session_state.page = "Overview"
                            st.rerun()
                        else:
                            st.error(result)

            else:
                st.markdown("""
                <div class="teamtrack-auth-heading">
                  <h2>Join your company</h2>
                  <p>Enter the invitation code shared by your company Admin.</p>
                </div>
                """, unsafe_allow_html=True)
                with st.form("join_company_form", clear_on_submit=False):
                    join_code = st.text_input("Company invitation code", placeholder="Enter your 8-character code")
                    full_name = st.text_input("Your full name", placeholder="Enter your full name")
                    join_email = st.text_input("Work email", placeholder="you@company.com")
                    join_password = st.text_input("Create password", type="password", placeholder="At least 8 characters")
                    join_password2 = st.text_input("Confirm password", type="password", placeholder="Enter password again")
                    join_team = st.selectbox("Your department", DEFAULT_TEAMS, key="join_team")
                    submitted = st.form_submit_button("Join company workspace", type="primary", use_container_width=True)
                if submitted:
                    if not all([join_code.strip(), full_name.strip(), join_email.strip(), join_password]):
                        st.error("Complete all required fields.")
                    elif not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", join_email.strip()):
                        st.error("Enter a valid email address.")
                    elif len(join_password) < 8:
                        st.error("Use a password with at least 8 characters.")
                    elif join_password != join_password2:
                        st.error("The passwords do not match.")
                    else:
                        ok, result = join_company(join_code, full_name, join_email, join_password, join_team)
                        if ok:
                            company_id, user_id, company_name, email_status = result
                            st.session_state.teamtrack_registration_email_status = email_status
                            st.session_state.user = dict(q_one(
                                """SELECT u.*, c.name AS company_name, c.join_code
                                   FROM users u JOIN companies c ON c.id=u.company_id WHERE u.id=?""",
                                (user_id,)
                            ))
                            st.session_state.page = "Overview"
                            st.rerun()
                        else:
                            st.error(result)

        if st.session_state.pop("teamtrack_password_reset_success", False):
            st.success("Your password has been updated. You can now sign in.")
        email_status = st.session_state.pop("teamtrack_registration_email_status", None)
        if email_status is not None:
            if email_status.get("user"):
                st.success("Registration confirmation email sent.")
            else:
                st.warning("Your account was created, but confirmation email delivery is not configured or could not be completed. Set the SMTP secrets to enable email.")
            if not email_status.get("company", False) and st.session_state.get("teamtrack_auth_mode") == "join":
                st.info("The company notification could not be sent. Check the SMTP configuration.")
        st.markdown("<div class='teamtrack-auth-foot'>TeamTrack · Your team's work, in one place.</div>", unsafe_allow_html=True)
    st.stop()


# -----------------------------
# AUTHENTICATED APP SHELL
# -----------------------------
user = st.session_state.user
company_id = user["company_id"]
company_name = user["company_name"]
members = get_company_users(company_id)
teams = [r["name"] for r in q_all("SELECT name FROM teams WHERE company_id=? ORDER BY name", (company_id,))]
registration_email_status = st.session_state.pop("teamtrack_registration_email_status", None)
if registration_email_status is not None:
    if registration_email_status.get("user"):
        st.success("Registration confirmation email sent.")
    else:
        st.warning("Your account was created, but email delivery is not configured or could not be completed. Configure the SMTP secrets to enable confirmation emails.")
    if not registration_email_status.get("company", False) and user.get("role") != "Admin":
        st.info("The company notification email could not be sent. Check the SMTP configuration.")
if not teams:
    teams = DEFAULT_TEAMS
all_tickets = ticket_rows(user)
if user["role"] == "Team Lead":
    all_tickets = [ticket for ticket in all_tickets if ticket["team"] == user["team"]]
unread = q_one("SELECT COUNT(*) AS n FROM notifications WHERE user_id=? AND is_read=0", (user["id"],))["n"]

# Show the new workspace invitation code once the first Admin has created
# a company, so they can share it with colleagues.
if st.session_state.get("teamtrack_created_join_code"):
    st.markdown(f"<div class='tt-created-success-card'>{html.escape(company_name)} workspace created successfully.</div>", unsafe_allow_html=True)
    st.markdown("**Your company invitation code**")
    st.code(st.session_state.teamtrack_created_join_code, language=None)
    if st.button("I've saved the invitation code", key="dismiss_created_code"):
        st.session_state.pop("teamtrack_created_join_code", None)
        st.rerun()

with st.sidebar:
    # Compact graphic logo replaces the text wordmark after authentication.
    if LOGO_PATH.exists():
        st.image(str(LOGO_PATH), width=40)
    else:
        st.markdown("<div class='teamtrack-workspace-wordmark'>TeamTrack</div>", unsafe_allow_html=True)
    st.caption(company_name)
    workspace_dark_mode = st.toggle(
        "Dark mode",
        value=st.session_state.teamtrack_theme == "Dark",
        key="workspace_dark_mode_toggle",
        help="Switch the entire TeamTrack workspace between light and dark appearance.",
    )
    selected_theme = "Dark" if workspace_dark_mode else "Light"
    if selected_theme != st.session_state.teamtrack_theme:
        st.session_state.teamtrack_theme = selected_theme
        st.rerun()
    st.markdown("---")
    pages = ["Overview", "All tickets", "My tickets", "Notifications"]
    if user["role"] != "Viewer":
        pages.insert(3, "Create ticket")
    if can_manage(user):
        pages += ["Team & members", "Reports"]
    if is_admin(user):
        pages += ["Workspace settings"]
    current = st.session_state.get("page", "Overview")
    if current not in pages:
        current = "Overview"
    for page in pages:
        label = f"{page}  ({unread})" if page == "Notifications" and unread else page
        if st.button(label, key=f"nav_{page}", use_container_width=True,
                     type="primary" if current == page else "secondary"):
            st.session_state.page = page
            st.session_state.pop("selected_ticket", None)
            st.rerun()
    st.markdown("---")
    st.markdown(f"**{user['full_name']}**")
    st.caption(f"{user['role']} · {user['team']}")
    if st.button("Log out", use_container_width=True):
        logout()


page = st.session_state.get("page", "Overview")

# -----------------------------
# OVERVIEW
# -----------------------------
if page == "Overview":
    st.title("Good to see you, " + user["full_name"].split()[0])
    st.caption("Here's what's happening across your company.")
    total = len(all_tickets)
    open_count = sum(t["status"] in ["Open", "Assigned", "In Progress", "On Hold"] for t in all_tickets)
    in_progress = sum(t["status"] == "In Progress" for t in all_tickets)
    resolved = sum(t["status"] in ["Resolved", "Closed"] for t in all_tickets)
    overdue = sum(bool(t["due_date"]) and t["due_date"] < datetime.now().date().isoformat()
                  and t["status"] not in ["Resolved", "Closed"] for t in all_tickets)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total tickets", total)
    c2.metric("Open", open_count)
    c3.metric("In progress", in_progress)
    c4.metric("Resolved", resolved)
    c5.metric("Overdue", overdue)
    st.markdown("### Recent tickets")
    if not all_tickets:
        st.info("No tickets yet. Create the first ticket to get your team started.")
        if st.button("Create your first ticket", type="primary"):
            st.session_state.page = "Create ticket"
            st.rerun()
    else:
        for t in all_tickets[:8]:
            with st.container(border=True):
                left, middle, right = st.columns([5, 2, 1])
                left.markdown(f"**{t['ticket_key']} · {t['title']}**")
                left.caption(f"{t['team']} · Reported by {t['reporter_name']} · {t['assignee_name'] or 'Unassigned'}")
                middle.markdown(status_badge(t["status"]), unsafe_allow_html=True)
                middle.markdown(priority_badge(t["priority"]), unsafe_allow_html=True)
                if right.button("Open", key=f"overview_{t['id']}"):
                    open_ticket(t["id"])


# -----------------------------
# CREATE TICKET
# -----------------------------
elif page == "Create ticket":
    st.title("Create a ticket")
    st.caption("Describe the issue clearly so the right team can take action.")
    with st.form("new_ticket_form"):
        title = st.text_input("Issue title", placeholder="e.g. Website contact form is not submitting")
        description = st.text_area("Description", height=150, placeholder="What happened? What did you expect to happen? Include any useful context.")
        col1, col2 = st.columns(2)
        category = col1.selectbox("Category", ["Technical issue", "People & HR", "Operations", "Finance", "Marketing", "Customer support", "Product", "Other"])
        team = col2.selectbox("Responsible team", teams)
        col3, col4 = st.columns(2)
        priority = col3.selectbox("Priority", PRIORITIES, index=2)
        due_date = col4.date_input("Due date (optional)", value=None)
        can_assign_here = can_assign_in_department(user, team)
        if can_assign_here:
            assignable_members = [m for m in members if m["team"] == team]
            assignee_options = ["Unassigned"] + [f"{m['full_name']} ({m['team']})" for m in assignable_members]
            assignee_choice = st.selectbox("Assign to (optional)", assignee_options)
        else:
            assignee_choice = "Unassigned"
            st.caption("A Team Lead, IT Personnel member or Company Admin will assign this ticket.")
        attachment = st.file_uploader("Attachment (optional)", type=["png", "jpg", "jpeg", "pdf", "txt", "doc", "docx", "xlsx", "csv"])
        submitted = st.form_submit_button("Submit ticket", type="primary", use_container_width=True)
    if submitted:
        if not title.strip() or not description.strip():
            st.error("Add a title and description before submitting.")
        else:
            ticket_id = str(uuid.uuid4())
            key = ticket_key(company_id)
            assignee_id = None
            if assignee_choice != "Unassigned" and can_assign_in_department(user, team):
                assignee_name = assignee_choice.rsplit(" (", 1)[0]
                found = next((m for m in members if m["full_name"] == assignee_name and m["team"] == team), None)
                assignee_id = found["id"] if found else None
            status = "Assigned" if assignee_id else "Open"
            due = due_date.isoformat() if due_date else None
            execute(
                """INSERT INTO tickets
                   (id, company_id, ticket_key, title, description, category, team, priority, status,
                    reporter_id, assignee_id, due_date, created_at, updated_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (ticket_id, company_id, key, title.strip(), description.strip(), category, team,
                 priority, status, user["id"], assignee_id, due, now(), now())
            )
            if attachment:
                safe_name = re.sub(r"[^A-Za-z0-9_.-]", "_", attachment.name)
                path = UPLOAD_DIR / f"{ticket_id}_{safe_name}"
                path.write_bytes(attachment.getbuffer())
            add_activity(ticket_id, user["id"], f"Ticket created by {user['full_name']}")
            # Notify the selected department and the people directly involved only.
            recipients = ticket_notification_recipients(
                company_id, team, reporter_id=user["id"], assignee_id=assignee_id,
                exclude_user_id=user["id"]
            )
            notify_users(
                company_id, recipients, ticket_id,
                f"New ticket {key}: {title.strip()} · Team: {team} · Priority: {priority}",
                "New department ticket"
            )
            st.success(f"Ticket {key} created. The {team} department has been notified.")
            st.session_state.page = "All tickets"
            st.rerun()


# -----------------------------
# TICKET LISTS
# -----------------------------
elif page in ["All tickets", "My tickets"]:
    st.title(page)
    scoped = all_tickets
    if page == "My tickets":
        scoped = [t for t in scoped if t["reporter_id"] == user["id"] or t["assignee_id"] == user["id"]]
    f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
    search = f1.text_input("Search tickets", placeholder="Search title or ticket ID")
    status_filter = f2.selectbox("Status", ["All"] + STATUSES)
    priority_filter = f3.selectbox("Priority", ["All"] + PRIORITIES)
    team_filter = f4.selectbox("Team", ["All"] + teams)
    for t in scoped:
        if search and search.lower() not in (t["title"] + t["ticket_key"] + t["description"]).lower():
            continue
        if status_filter != "All" and t["status"] != status_filter:
            continue
        if priority_filter != "All" and t["priority"] != priority_filter:
            continue
        if team_filter != "All" and t["team"] != team_filter:
            continue
        with st.container(border=True):
            a, b, c = st.columns([5, 2, 1])
            a.markdown(f"**{t['ticket_key']} · {t['title']}**")
            a.caption(f"{t['team']} · {t['reporter_name']} → {t['assignee_name'] or 'Unassigned'}")
            b.markdown(status_badge(t["status"]), unsafe_allow_html=True)
            b.markdown(priority_badge(t["priority"]), unsafe_allow_html=True)
            if c.button("View", key=f"list_{t['id']}"):
                open_ticket(t["id"])
    if not scoped:
        st.info("No tickets match this view yet.")


# -----------------------------
# TICKET DETAILS
# -----------------------------
elif page == "Ticket details":
    ticket_id = st.session_state.get("selected_ticket")
    t = get_ticket(ticket_id, company_id) if ticket_id else None
    if not t:
        st.warning("Ticket not found.")
        if st.button("Back to tickets"):
            st.session_state.page = "All tickets"
            st.rerun()
    else:
        if st.button("← Back to tickets"):
            st.session_state.page = "All tickets"
            st.rerun()
        st.title(f"{t['ticket_key']} · {t['title']}")
        st.markdown(f"{status_badge(t['status'])} &nbsp; {priority_badge(t['priority'])}", unsafe_allow_html=True)
        left, right = st.columns([2, 1])
        with left:
            st.markdown("### Description")
            st.write(t["description"])
            st.caption(f"Category: {t['category']} · Team: {t['team']}")
            st.markdown("### Updates and comments")
            comments = q_all(
                """SELECT c.*, u.full_name FROM comments c JOIN users u ON u.id=c.user_id
                   WHERE c.ticket_id=? ORDER BY c.created_at""", (ticket_id,)
            )
            if not comments:
                st.caption("No comments yet.")
            for comment in comments:
                with st.container(border=True):
                    st.markdown(f"**{comment['full_name']}**")
                    st.write(comment["body"])
                    st.caption(comment["created_at"][:16].replace("T", " "))
            submitted = False
            comment_body = ""
            if user["role"] != "Viewer":
                with st.form("comment_form"):
                    comment_body = st.text_area("Add a comment or progress update")
                    submitted = st.form_submit_button("Post comment", type="primary")
            else:
                st.caption("Viewer access is read-only.")
            if submitted:
                if comment_body.strip():
                    execute(
                        "INSERT INTO comments (id, ticket_id, user_id, body, created_at) VALUES (?, ?, ?, ?, ?)",
                        (str(uuid.uuid4()), ticket_id, user["id"], comment_body.strip(), now())
                    )
                    add_activity(ticket_id, user["id"], f"{user['full_name']} added a comment")
                    recipients = ticket_notification_recipients(
                        company_id, t["team"], reporter_id=t["reporter_id"],
                        assignee_id=t["assignee_id"], exclude_user_id=user["id"]
                    )
                    notify_users(company_id, recipients, ticket_id,
                                 f"New comment on {t['ticket_key']}: {comment_body.strip()[:120]}",
                                 "Ticket comment")
                    st.rerun()
        with right:
            st.markdown("### Ticket details")
            st.write(f"**Reporter:** {t['reporter_name']}")
            st.write(f"**Assignee:** {t['assignee_name'] or 'Unassigned'}")
            st.write(f"**Created:** {t['created_at'][:10]}")
            st.write(f"**Due:** {t['due_date'] or 'Not set'}")
            st.markdown("---")
            team_lead_can_access = user["role"] == "Team Lead" and user["team"] == t["team"]
            if user["role"] == "Team Lead" and not team_lead_can_access:
                st.error("You can only access tickets belonging to your department.")
                st.stop()
            can_update_work = (user["role"] in ["Admin", "IT Personnel"] or team_lead_can_access or
                               (user["role"] != "Viewer" and user["id"] in [t["reporter_id"], t["assignee_id"]]))
            can_reassign = can_assign_in_department(user, t["team"])
            if t["status"] == "Closed":
                st.info("This ticket is closed and cannot be reopened or edited. If further work is required, create a new ticket and reference this ticket ID.")
            elif can_update_work:
                with st.form("ticket_update_form"):
                    new_status = st.selectbox("Status", STATUSES, index=STATUSES.index(t["status"]))
                    new_priority = st.selectbox("Priority", PRIORITIES, index=PRIORITIES.index(t["priority"]))
                    if can_reassign:
                        if can_assign_across_departments(user):
                            permitted_teams = teams
                        else:
                            permitted_teams = [user["team"]]
                        team_index = permitted_teams.index(t["team"]) if t["team"] in permitted_teams else 0
                        new_team = st.selectbox("Department", permitted_teams, index=team_index)
                        assignable = [m for m in members if m["team"] == new_team]
                        assign_labels = ["Unassigned"] + [f"{m['full_name']} ({m['team']})" for m in assignable]
                        current_assignee = next((f"{m['full_name']} ({m['team']})" for m in assignable if m["id"] == t["assignee_id"]), "Unassigned")
                        new_assignee_label = st.selectbox("Assignee", assign_labels,
                            index=assign_labels.index(current_assignee) if current_assignee in assign_labels else 0)
                    else:
                        new_team = t["team"]
                        new_assignee_label = None
                        st.caption("Only Company Admins, IT Personnel and the relevant Department Team Lead can assign or reassign this ticket.")
                    update = st.form_submit_button("Save changes", type="primary", use_container_width=True)
                if update:
                    new_assignee_id = t["assignee_id"]
                    if can_reassign and new_assignee_label is not None:
                        new_assignee_id = None
                        if new_assignee_label != "Unassigned":
                            new_assignee_name = new_assignee_label.rsplit(" (", 1)[0]
                            new_assignee_id = next((m["id"] for m in assignable if m["full_name"] == new_assignee_name and m["team"] == new_team), None)
                    changed = []
                    if new_status != t["status"]: changed.append(f"status changed to {new_status}")
                    if new_priority != t["priority"]: changed.append(f"priority changed to {new_priority}")
                    if new_assignee_id != t["assignee_id"]: changed.append("assignee updated")
                    resolved_at = now() if new_status in ["Resolved", "Closed"] else None
                    execute(
                        """UPDATE tickets SET status=?, priority=?, team=?, assignee_id=?, updated_at=?,
                           resolved_at=CASE WHEN ? IS NOT NULL THEN ? WHEN ? IN ('Open','Assigned','In Progress','On Hold') THEN NULL ELSE resolved_at END
                           WHERE id=? AND company_id=?""",
                        (new_status, new_priority, new_team, new_assignee_id, now(),
                         resolved_at, resolved_at, new_status, ticket_id, company_id)
                    )
                    if changed:
                        action = f"{user['full_name']} " + ", ".join(changed)
                        add_activity(ticket_id, user["id"], action)
                        recipients = ticket_notification_recipients(
                            company_id, new_team, reporter_id=t["reporter_id"],
                            assignee_id=new_assignee_id, exclude_user_id=user["id"]
                        )
                        notify_users(company_id, recipients, ticket_id,
                                     f"{t['ticket_key']} updated: {', '.join(changed)}",
                                     "Ticket updated")
                    st.success("Ticket updated.")
                    st.rerun()
            else:
                st.info("You can view this ticket but do not have permission to edit it.")
            st.markdown("### Activity")
            activities = q_all(
                """SELECT a.*, u.full_name FROM activity a JOIN users u ON u.id=a.user_id
                   WHERE a.ticket_id=? ORDER BY a.created_at DESC""", (ticket_id,)
            )
            for item in activities:
                st.caption(f"{item['created_at'][:16].replace('T',' ')} · {item['action']}")


# -----------------------------
# NOTIFICATIONS
# -----------------------------
elif page == "Notifications":
    st.title("Notifications")
    notifs = q_all(
        """SELECT n.*, t.ticket_key FROM notifications n
           LEFT JOIN tickets t ON t.id=n.ticket_id
           WHERE n.user_id=? AND n.company_id=? ORDER BY n.created_at DESC LIMIT 100""",
        (user["id"], company_id)
    )
    top1, top2 = st.columns([1, 4])
    if top1.button("Mark all as read"):
        execute("UPDATE notifications SET is_read=1 WHERE user_id=? AND company_id=?", (user["id"], company_id))
        st.rerun()
    if not notifs:
        st.info("You're all caught up. New ticket activity will appear here.")
    for n in notifs:
        with st.container(border=True):
            a, b = st.columns([5, 1])
            prefix = "● " if not n["is_read"] else ""
            a.markdown(f"**{prefix}{n['message']}**")
            a.caption(n["created_at"][:16].replace("T", " "))
            if n["ticket_id"] and b.button("Open ticket", key=f"notif_{n['id']}"):
                execute("UPDATE notifications SET is_read=1 WHERE id=?", (n["id"],))
                open_ticket(n["ticket_id"])
            elif not n["is_read"] and b.button("Mark read", key=f"read_{n['id']}"):
                execute("UPDATE notifications SET is_read=1 WHERE id=?", (n["id"],))
                st.rerun()


# -----------------------------
# TEAM & MEMBERS
# -----------------------------
elif page == "Team & members":
    st.title("Team & members")
    st.caption("Manage departments and the people in your company workspace.")
    st.markdown("### Company invitation code")
    st.code(user["join_code"])
    st.caption("Share this code only with people who should join your company workspace.")
    st.markdown("### Members")
    for m in members:
        with st.container(border=True):
            a, b, c = st.columns([3, 2, 2])
            a.markdown(f"**{m['full_name']}**")
            a.caption(m["email"])
            b.write(m["team"])
            c.write(m["role"])
    if is_admin(user):
        st.markdown("### Manage employee roles")
        st.caption("Assign IT Personnel and Department Team Lead permissions. Team Leads should belong to the department they lead.")
        role_candidates = [m for m in members if m["id"] != user["id"]]
        if role_candidates:
            member_labels = {f"{m['full_name']} · {m['email']} · {m['team']}": m for m in role_candidates}
            selected_label = st.selectbox("Employee", list(member_labels.keys()), key="role_employee_select")
            selected_member = member_labels[selected_label]
            role_col, dept_col = st.columns(2)
            current_role = selected_member["role"] if selected_member["role"] in ROLES else "Member"
            new_role = role_col.selectbox("Role", ROLES, index=ROLES.index(current_role), key="role_value_select")
            new_department = dept_col.selectbox("Department", teams,
                index=teams.index(selected_member["team"]) if selected_member["team"] in teams else 0,
                key="role_department_select")
            if st.button("Save employee access", type="primary", key="save_employee_access"):
                if selected_member["id"] == user["id"]:
                    st.error("You cannot change your own administrator access here.")
                else:
                    execute("UPDATE users SET role=?, team=? WHERE id=? AND company_id=?",
                            (new_role, new_department, selected_member["id"], company_id))
                    st.success(f"Updated {selected_member['full_name']}'s role and department.")
                    st.rerun()
        else:
            st.info("Invite employees before managing their roles.")
    st.markdown("### Departments")
    if not DARK_MODE:
        st.markdown("""
        <style>
        /* Light theme: every department-removal popover trigger is burgundy. */
        [data-testid="stPopover"] button,
        [data-testid="stPopover"] > div button {
            background-color: #4D0F28 !important;
            border: 1px solid #4D0F28 !important;
            color: #FFFFFF !important;
            border-radius: 8px !important;
        }
        [data-testid="stPopover"] button *,
        [data-testid="stPopover"] button svg {
            color: #FFFFFF !important;
            fill: #FFFFFF !important;
        }
        [data-testid="stPopover"] button:hover {
            background-color: #681B39 !important;
            border-color: #681B39 !important;
        }

        /* Streamlit renders popover content in an overlay/portal. Match the
           overlay containing our marker, with fallbacks for Streamlit DOM versions. */
        [data-baseweb="popover"]:has(.tt-remove-dept-marker),
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker),
        [role="dialog"]:has(.tt-remove-dept-marker),
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) > div,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) > div,
        [role="dialog"]:has(.tt-remove-dept-marker) > div {
            background-color: #4D0F28 !important;
            border-color: #4D0F28 !important;
            color: #FFFFFF !important;
        }
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) *,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) *,
        [role="dialog"]:has(.tt-remove-dept-marker) *,
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) *::before,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) *::before,
        [role="dialog"]:has(.tt-remove-dept-marker) *::before {
            color: #FFFFFF !important;
        }
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) [data-testid="stCheckbox"] label,
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) [data-testid="stCheckbox"] label *,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) [data-testid="stCheckbox"] label,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) [data-testid="stCheckbox"] label *,
        [role="dialog"]:has(.tt-remove-dept-marker) [data-testid="stCheckbox"] label,
        [role="dialog"]:has(.tt-remove-dept-marker) [data-testid="stCheckbox"] label * {
            color: #FFFFFF !important;
        }
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) button,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) button,
        [role="dialog"]:has(.tt-remove-dept-marker) button {
            background-color: #681B39 !important;
            border-color: #FFFFFF !important;
            color: #FFFFFF !important;
        }
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) button *,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) button *,
        [role="dialog"]:has(.tt-remove-dept-marker) button * {
            color: #FFFFFF !important;
            fill: #FFFFFF !important;
        }
        /* The final destructive-action confirmation uses a white button
           with burgundy text, while the surrounding popover stays burgundy. */
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button,
        [role="dialog"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button {
            background-color: #FFFFFF !important;
            background: #FFFFFF !important;
            border: 1px solid #FFFFFF !important;
            color: #4D0F28 !important;
        }
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button *,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button *,
        [role="dialog"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button * {
            color: #4D0F28 !important;
            fill: #4D0F28 !important;
        }
        [data-baseweb="popover"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button:hover,
        [data-testid="stPopoverBody"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button:hover,
        [role="dialog"]:has(.tt-remove-dept-marker) [data-testid="stButton"] button:hover {
            background-color: #F4E9ED !important;
            background: #F4E9ED !important;
            border-color: #FFFFFF !important;
        }
        </style>
        """, unsafe_allow_html=True)
    for team_name in teams:
        dept_col, action_col = st.columns([4, 1.4])
        dept_col.write("• " + team_name)
        if is_admin(user):
            with action_col:
                with st.popover("Remove department"):
                    st.markdown('<span class="tt-remove-dept-marker" aria-hidden="true"></span>', unsafe_allow_html=True)
                    st.caption(f"Remove **{team_name}** from this company workspace.")
                    confirm_remove = st.checkbox(
                        "I understand this cannot be undone.",
                        key=f"confirm_remove_department_{team_name}"
                    )
                    if st.button("Confirm removal", key=f"remove_department_{team_name}",
                                 type="primary", disabled=not confirm_remove):
                        assigned_members = q_one(
                            "SELECT COUNT(*) AS total FROM users WHERE company_id=? AND team=?",
                            (company_id, team_name)
                        )["total"]
                        related_tickets = q_one(
                            "SELECT COUNT(*) AS total FROM tickets WHERE company_id=? AND team=?",
                            (company_id, team_name)
                        )["total"]
                        if assigned_members or related_tickets:
                            st.error(
                                f"Cannot remove this department yet: {assigned_members} active member(s) "
                                f"and {related_tickets} ticket(s) are still linked to it. "
                                "Move members and resolve or reassign its tickets first."
                            )
                        else:
                            execute("DELETE FROM teams WHERE company_id=? AND name=?",
                                    (company_id, team_name))
                            st.success(f"{team_name} department removed.")
                            st.rerun()
    if is_admin(user):
        st.markdown("### Add a department")
        with st.form("add_team_form"):
            new_team = st.text_input("Department name")
            submitted = st.form_submit_button("Add department")
        if submitted:
            if new_team.strip():
                try:
                    execute("INSERT INTO teams (id, company_id, name) VALUES (?, ?, ?)",
                            (str(uuid.uuid4()), company_id, new_team.strip()))
                    st.success("Department added.")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("That department already exists.")


# -----------------------------
# REPORTS & ANALYTICS
# -----------------------------
elif page == "Reports":
    st.title("Reports & Analytics")
    st.caption("Explore ticket activity and export company reports in PDF, Word, Excel or CSV format.")

    today = date.today()
    default_start = today - timedelta(days=30)
    f1, f2, f3 = st.columns(3)
    with f1:
        start_date = st.date_input("From", value=default_start, key="report_start_date")
    with f2:
        end_date = st.date_input("To", value=today, key="report_end_date")
    with f3:
        report_team = st.selectbox("Department", ["All departments"] + teams, key="report_team_filter")

    f4, f5, f6 = st.columns(3)
    with f4:
        report_status = st.selectbox("Status", ["All statuses"] + STATUSES, key="report_status_filter")
    with f5:
        report_priority = st.selectbox("Priority", ["All priorities"] + PRIORITIES, key="report_priority_filter")
    with f6:
        member_labels = {"All employees": None}
        for m in members:
            member_labels[f"{m['full_name']} ({m['team']})"] = m["id"]
        selected_member_label = st.selectbox("Employee", list(member_labels.keys()), key="report_member_filter")
        selected_member_id = member_labels[selected_member_label]

    if start_date > end_date:
        st.error("The start date must be on or before the end date.")
        filtered_tickets = []
    else:
        filtered_tickets = []
        for ticket in all_tickets:
            created_day = (ticket["created_at"] or "")[:10]
            if not created_day or not (start_date.isoformat() <= created_day <= end_date.isoformat()):
                continue
            if report_team != "All departments" and ticket["team"] != report_team:
                continue
            if report_status != "All statuses" and ticket["status"] != report_status:
                continue
            if report_priority != "All priorities" and ticket["priority"] != report_priority:
                continue
            if selected_member_id and selected_member_id not in [ticket["reporter_id"], ticket["assignee_id"]]:
                continue
            filtered_tickets.append(ticket)

    report_df = build_report_dataframe(filtered_tickets)
    total = len(filtered_tickets)
    active_count = sum(t["status"] not in ["Resolved", "Closed"] for t in filtered_tickets)
    resolved_count = sum(t["status"] in ["Resolved", "Closed"] for t in filtered_tickets)
    overdue_count = sum(bool(t["due_date"]) and t["due_date"] < today.isoformat()
                        and t["status"] not in ["Resolved", "Closed"] for t in filtered_tickets)
    resolution_rate = round((resolved_count / total) * 100, 1) if total else 0
    summary_df = pd.DataFrame([
        {"Metric": "Total tickets", "Value": total},
        {"Metric": "Active tickets", "Value": active_count},
        {"Metric": "Resolved or closed", "Value": resolved_count},
        {"Metric": "Overdue tickets", "Value": overdue_count},
        {"Metric": "Resolution rate", "Value": f"{resolution_rate}%"},
    ])

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total tickets", total)
    m2.metric("Active", active_count)
    m3.metric("Resolved / closed", resolved_count)
    m4.metric("Overdue", overdue_count)
    m5.metric("Resolution rate", f"{resolution_rate}%")

    if total:
        status_counts = report_df["Status"].value_counts().reindex(STATUSES, fill_value=0)
        priority_counts = report_df["Priority"].value_counts().reindex(PRIORITIES, fill_value=0)
        team_counts = report_df["Department"].value_counts()
        chart1, chart2 = st.columns(2)
        with chart1:
            st.markdown("### Tickets by status")
            st.bar_chart(status_counts)
        with chart2:
            st.markdown("### Tickets by priority")
            st.bar_chart(priority_counts)
        st.markdown("### Department workload")
        st.bar_chart(team_counts)
        st.markdown("### Ticket details")
        visible_columns = ["Ticket ID", "Title", "Department", "Priority", "Status", "Reported By", "Assigned To", "Due Date", "Created At"]
        st.dataframe(report_df[visible_columns], use_container_width=True, hide_index=True)
    else:
        st.info("No tickets match the selected reporting filters. You can still export an empty report with its summary.")

    st.markdown("---")
    st.markdown("### Export report")
    generated_at = datetime.now().strftime("%d %b %Y, %H:%M")
    period_label = f"{start_date.strftime('%d %b %Y')} – {end_date.strftime('%d %b %Y')}" if start_date <= end_date else "Invalid date range"
    file_stub = re.sub(r"[^A-Za-z0-9_-]+", "_", company_name).strip("_") or "Company"
    export1, export2, export3, export4 = st.columns(4)
    with export1:
        try:
            pdf_bytes = make_pdf_report(report_df, summary_df, company_name, generated_at, period_label)
            st.download_button("Download PDF", data=pdf_bytes,
                               file_name=f"{file_stub}_TeamTrack_Report.pdf",
                               mime="application/pdf", use_container_width=True,
                               disabled=start_date > end_date)
        except Exception as exc:
            st.error(f"PDF export unavailable: {exc}")
    with export2:
        try:
            word_bytes = make_word_report(report_df, summary_df, company_name, generated_at, period_label)
            st.download_button("Download Word", data=word_bytes,
                               file_name=f"{file_stub}_TeamTrack_Report.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                               use_container_width=True, disabled=start_date > end_date)
        except Exception as exc:
            st.error(f"Word export unavailable: {exc}")
    with export3:
        try:
            excel_bytes = make_excel_report(report_df, summary_df, company_name, generated_at)
            st.download_button("Download Excel", data=excel_bytes,
                               file_name=f"{file_stub}_TeamTrack_Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, disabled=start_date > end_date)
        except Exception as exc:
            st.error(f"Excel export unavailable: {exc}")
    with export4:
        csv_buffer = io.StringIO()
        report_df.to_csv(csv_buffer, index=False)
        st.download_button("Download CSV", data=csv_buffer.getvalue().encode("utf-8-sig"),
                           file_name=f"{file_stub}_TeamTrack_Tickets.csv",
                           mime="text/csv", use_container_width=True,
                           disabled=start_date > end_date)
    st.caption("Exports contain only tickets belonging to the signed-in user's company workspace and matching the selected filters. Date filtering uses ticket creation date.")

# -----------------------------
# WORKSPACE SETTINGS
# -----------------------------
elif page == "Workspace settings":
    st.title("Workspace settings")
    st.caption("Company workspace information and access details.")
    st.text_input("Company name", value=company_name, disabled=True)
    st.text_input("Company invitation code", value=user["join_code"], disabled=True)
    st.markdown("<div class='tt-workspace-settings-note'>Workspace identity and invitation codes are currently read-only in this MVP. Advanced workspace access and management controls will be introduced in a future update.</div>", unsafe_allow_html=True)

st.markdown("---")
st.caption("TeamTrack MVP · Keep work visible. Keep teams moving.")
