import os
import io
import csv
import json
import datetime
import re
import time
import unicodedata
from zoneinfo import ZoneInfo

import libsql_client
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

from flask import Flask, request, abort, Response
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi, ReplyMessageRequest,
    PushMessageRequest, TextMessage, FlexMessage, FlexContainer
)
from linebot.v3.webhooks import (
    MessageEvent, TextMessageContent, ImageMessageContent,
    FileMessageContent, VideoMessageContent, AudioMessageContent
)
from apscheduler.schedulers.background import BackgroundScheduler

# ---------- Config (set these as environment variables on your host) ----------
CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]

OWNER_USER_ID = os.environ.get("OWNER_USER_ID", "")  # who receives the daily digest + can type commands

# Comma-separated LINE user IDs that count as "tagged" (you + any admin, e.g. "Uabc123,Udef456")
TAGGED_USER_IDS = set(x.strip() for x in os.environ.get("TAGGED_USER_IDS", "").split(",") if x.strip())

# Comma-separated group IDs where EVERY tag counts (instructions/FYI), no file required.
# Any group NOT listed here uses "file-sending" mode: a tag only counts if paired with a file.
PM_MODE_GROUP_IDS = set(x.strip() for x in os.environ.get("PM_MODE_GROUP_IDS", "").split(",") if x.strip())

DIGEST_HOUR = int(os.environ.get("DIGEST_HOUR", "18"))   # 24h, Asia/Bangkok
DIGEST_MINUTE = int(os.environ.get("DIGEST_MINUTE", "0"))

OVERDUE_HOUR = int(os.environ.get("OVERDUE_HOUR", "8"))    # morning check for files left over from before today
OVERDUE_MINUTE = int(os.environ.get("OVERDUE_MINUTE", "0"))
TAG_LOOKBACK_MINUTES = int(os.environ.get("TAG_LOOKBACK_MINUTES", "30"))

# Group that pasted Site Around inspection-document lists get posted to after
# an admin confirms the preview (see "Inspection document list" below).
INSPECTION_TARGET_GROUP_ID = os.environ.get("INSPECTION_TARGET_GROUP_ID", "")

# Optional shared secret required as ?key=... on /report/daily.csv (see "Daily
# backend report" below). Leave unset to allow the report link with no key.
REPORT_ACCESS_KEY = os.environ.get("REPORT_ACCESS_KEY", "")

# Public base URL of this service (e.g. https://line-file-tracker-bot.onrender.com).
# Needed to build the report link from the scheduled daily digest job, which
# runs with no incoming request to read the host from; the 1:1 "รายงาน"
# command works either way since it can read the host from its own request.
SERVICE_BASE_URL = os.environ.get("SERVICE_BASE_URL", "").rstrip("/")

# Document replies (team types answers into a posted inspection document).
# LIFF_ID: the LIFF app's ID (from LINE Developers > LINE Login channel > LIFF),
#   endpoint URL set to <SERVICE_BASE_URL>/liff/reply. Without it the cards get
#   no "ตอบกลับ" button, but the in-group shortcut "AR-337 <ข้อความ>" still works.
# LINE_LOGIN_CHANNEL_ID: that LINE Login channel's Channel ID, used to verify
#   the ID token the LIFF page sends, so a reply is always tied to a real LINE user.
LIFF_ID = os.environ.get("LIFF_ID", "").strip()
LINE_LOGIN_CHANNEL_ID = os.environ.get("LINE_LOGIN_CHANNEL_ID", "").strip()

# Turso (hosted libSQL/SQLite) - persistent storage, survives redeploys unlike a local
# sqlite file on Render's free tier (which has no disk and resets on every deploy).
TURSO_DATABASE_URL = os.environ["TURSO_DATABASE_URL"]
TURSO_AUTH_TOKEN = os.environ["TURSO_AUTH_TOKEN"]

def _turso_http_url(url):
    """libsql_client maps the libsql:// scheme to a WebSocket (hrana) connection.
    That WebSocket handshake fails (400 Invalid response status) in some hosting
    environments (e.g. Render). Forcing the https:// scheme makes the client use
    plain HTTP (hrana-over-HTTP) instead, which works reliably there."""
    if url.startswith("libsql://"):
        return "https://" + url[len("libsql://"):]
    return url

TURSO_HTTP_URL = _turso_http_url(TURSO_DATABASE_URL)

app = Flask(__name__)
handler = WebhookHandler(CHANNEL_SECRET)
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)

# ---------- Storage ----------
SCHEMA_STATEMENTS = [
    """
    CREATE TABLE IF NOT EXISTS file_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_id TEXT,
        group_name TEXT,
        sender_id TEXT,
        sender_name TEXT,
        event_type TEXT,      -- 'file' or 'pm'
        msg_type TEXT,
        file_name TEXT,
        text_content TEXT,
        sent_at TEXT,
        tagged INTEGER DEFAULT 0,
        tagged_at TEXT,
        downloaded INTEGER DEFAULT 0,
        downloaded_at TEXT,
        downloaded_by TEXT,
        mention_spans TEXT,
        tagged_user_ids TEXT
    )
    """,
    # pending tag markers, for when the tag message arrives BEFORE the file(s)
    """
    CREATE TABLE IF NOT EXISTS pending_tags (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        group_id TEXT,
        sender_id TEXT,
        text_content TEXT,
        mention_spans TEXT,
        tagged_user_ids TEXT,
        created_at TEXT
    )
    """,
    # one pending "inspection document list" preview per admin, awaiting
    # ส่ง/ยกเลิก confirmation before it's posted to the group
    """
    CREATE TABLE IF NOT EXISTS pending_inspection_previews (
        user_id TEXT PRIMARY KEY,
        items_json TEXT,
        created_at TEXT
    )
    """,
    # simple key/value settings that can be changed from LINE without a
    # redeploy (e.g. which group the inspection list gets posted to)
    """
    CREATE TABLE IF NOT EXISTS bot_settings (
        key TEXT PRIMARY KEY,
        value TEXT
    )
    """,
    # one row per close action ("โหลดแล้ว" / "✅ ..."), so "ย้อนกลับ" can
    # reopen exactly the items that action closed
    """
    CREATE TABLE IF NOT EXISTS close_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id TEXT,
        ids_json TEXT,
        description TEXT,
        created_at TEXT
    )
    """,
    # every inspection document that has been posted to a group, keyed by its
    # short code (e.g. "AR-337"), so the team can reply to it afterwards
    """
    CREATE TABLE IF NOT EXISTS inspection_docs (
        code TEXT PRIMARY KEY,
        doc_number TEXT,
        subject TEXT,
        date TEXT,
        extra_note TEXT,
        group_id TEXT,
        posted_at TEXT
    )
    """,
    # buddy teams: who must answer posted documents (see "Buddy teams")
    """
    CREATE TABLE IF NOT EXISTS buddy_teams (
        label TEXT PRIMARY KEY,
        prefixes TEXT,
        member_ids TEXT,
        group_id TEXT,
        updated_at TEXT
    )
    """,
    # everyone the bot has seen talk in a group: lets "ตั้งบัดดี้" resolve an
    # @mention to a user ID even when that person never added the bot as a friend
    """
    CREATE TABLE IF NOT EXISTS group_members (
        group_id TEXT,
        user_id TEXT,
        display_name TEXT,
        seen_at TEXT,
        PRIMARY KEY (group_id, user_id)
    )
    """,
    # replies typed by the team about a posted document - many per document
    """
    CREATE TABLE IF NOT EXISTS doc_replies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT,
        user_id TEXT,
        user_name TEXT,
        text TEXT,
        source TEXT,          -- 'group' (typed "AR-337 ...") or 'liff'
        created_at TEXT
    )
    """,
]

class _FetchResult:
    """Makes a libsql_client ResultSet behave like a sqlite3 cursor for the
    .fetchall()/.fetchone() calls used throughout this file."""
    def __init__(self, rows):
        self._rows = [tuple(r) for r in rows]

    def fetchall(self):
        return self._rows

    def fetchone(self):
        return self._rows[0] if self._rows else None

class _LibsqlConn:
    """Thin adapter so the rest of the app can keep using the sqlite3-style
    conn.execute(sql, params).fetchall()/.fetchone(), conn.executemany(...),
    conn.commit(), conn.close() calls against a Turso (libSQL) database."""
    def __init__(self, client):
        self._client = client

    def execute(self, sql, params=()):
        rs = self._client.execute(sql, list(params))
        return _FetchResult(rs.rows)

    def executemany(self, sql, seq_of_params):
        stmts = [libsql_client.Statement(sql, list(p)) for p in seq_of_params]
        if stmts:
            self._client.batch(stmts)

    def commit(self):
        pass  # each statement over the Turso HTTP API is already committed

    def close(self):
        self._client.close()

def db():
    client = libsql_client.create_client_sync(url=TURSO_HTTP_URL, auth_token=TURSO_AUTH_TOKEN)
    return _LibsqlConn(client)

def init_db_schema():
    client = libsql_client.create_client_sync(url=TURSO_HTTP_URL, auth_token=TURSO_AUTH_TOKEN)
    try:
        for stmt in SCHEMA_STATEMENTS:
            client.execute(stmt)
        try:  # added later: owner can close a doc by hand ("ปิดเอกสาร AR-337")
            client.execute("ALTER TABLE inspection_docs ADD COLUMN closed_at TEXT")
        except Exception:
            pass  # column already there
    finally:
        client.close()

init_db_schema()

def get_profile_name(group_id, user_id):
    try:
        with ApiClient(configuration) as api_client:
            api = MessagingApi(api_client)
            if group_id:
                prof = api.get_group_member_profile(group_id, user_id)
            else:
                prof = api.get_profile(user_id)
            return prof.display_name
    except Exception:
        return "ไม่ทราบชื่อ"

def get_group_name(group_id):
    try:
        with ApiClient(configuration) as api_client:
            api = MessagingApi(api_client)
            summary = api.get_group_summary(group_id)
            return summary.group_name
    except Exception:
        return group_id


# --- Group member directory (no friendship needed) ---
# LINE puts the sender's user ID on every group message even if they never
# added the bot, but only puts a user ID on an @mention when the mentioned
# person has. So we remember each sender's ID + current display name, and
# "ตั้งบัดดี้" falls back to matching an @name against this list.
_member_refreshed = {}
MEMBER_REFRESH_SECONDS = 6 * 3600

def _norm_name(name):
    return re.sub(r"\s+", " ", (name or "").strip().lstrip("@").strip()).casefold()

def save_group_member(group_id, user_id, display_name):
    if not (group_id and user_id and display_name):
        return
    conn = db()
    try:
        conn.execute(
            "INSERT INTO group_members (group_id, user_id, display_name, seen_at) VALUES (?,?,?,?) "
            "ON CONFLICT(group_id, user_id) DO UPDATE SET display_name = excluded.display_name, "
            "seen_at = excluded.seen_at",
            (group_id, user_id, display_name, bkk_now().isoformat()))
    finally:
        conn.close()

def remember_group_member(group_id, user_id, display_name=None):
    """Record a group message sender (name looked up at most every few hours)."""
    if not (group_id and user_id):
        return
    key = (group_id, user_id)
    now = time.time()
    if display_name is None and now - _member_refreshed.get(key, 0) < MEMBER_REFRESH_SECONDS:
        return
    try:
        name = display_name or get_profile_name(group_id, user_id)
        if name and name != "ไม่ทราบชื่อ":
            save_group_member(group_id, user_id, name)
            _member_refreshed[key] = now
    except Exception:
        app.logger.exception("remember group member failed")

def find_group_member_by_name(group_id, name):
    """user_id whose display name matches @name: this group first, then any
    group. None when unknown or ambiguous."""
    target = _norm_name(name)
    if not target:
        return None
    conn = db()
    try:
        rows = conn.execute("SELECT group_id, user_id, display_name FROM group_members").fetchall()
    finally:
        conn.close()
    for scope in ([r for r in rows if r[0] == group_id], rows):
        ids = {uid for _, uid, dn in scope if _norm_name(dn) == target}
        if len(ids) == 1:
            return ids.pop()
        if len(ids) > 1:
            return None
    return None


def load_member_name_map():
    """normalized display name -> user_id, only for names that point to one person."""
    conn = db()
    try:
        rows = conn.execute("SELECT user_id, display_name FROM group_members").fetchall()
    finally:
        conn.close()
    ids = {}
    for uid, dn in rows:
        ids.setdefault(_norm_name(dn), set()).add(uid)
    return {n: next(iter(u)) for n, u in ids.items() if n and len(u) == 1}

def reply_identity_ids(user_id, user_name, name_map):
    """IDs a reply should count for. A reply sent from the LIFF "ตอบกลับ" page
    carries the LINE Login channel's user ID, which can differ from the ID the
    same person has in the group (different provider) - so the reply also
    counts for the group member with the same display name."""
    out = {user_id} if user_id else set()
    alt = name_map.get(_norm_name(user_name))
    if alt:
        out.add(alt)
    return out

def get_setting(key, default=""):
    try:
        conn = db()
        row = conn.execute("SELECT value FROM bot_settings WHERE key = ?", (key,)).fetchone()
        conn.close()
        return row[0] if row and row[0] else default
    except Exception:
        app.logger.exception("get_setting failed for %s", key)
        return default

def set_setting(key, value):
    conn = db()
    conn.execute(
        "INSERT INTO bot_settings (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )
    conn.close()

def inspection_target_group_id():
    """Group set from LINE with "ตั้งกลุ่มแจ้งเอกสาร" wins; the
    INSPECTION_TARGET_GROUP_ID env var is only the fallback."""
    return get_setting("inspection_target_group_id", INSPECTION_TARGET_GROUP_ID)

BKK_TZ = ZoneInfo("Asia/Bangkok")

def bkk_now():
    """Current time in Thailand, regardless of the server's own timezone (Render runs
    in UTC), returned as a naive datetime whose wall-clock value IS Bangkok time -
    so it stores/sorts/formats exactly like before, just correct."""
    return datetime.datetime.now(BKK_TZ).replace(tzinfo=None)

def bkk_today():
    return bkk_now().date()

def utf16_index_to_py(text, utf16_target):
    """Converts a UTF-16 code-unit offset (as LINE's mention index/length use) into
    a Python string index. Characters outside the Basic Multilingual Plane (such as
    the 🐷 emoji) take 2 UTF-16 units but only 1 Python character."""
    utf16_pos = 0
    for py_i, ch in enumerate(text):
        if utf16_pos >= utf16_target:
            return py_i
        utf16_pos += 2 if ord(ch) > 0xFFFF else 1
    return len(text)

def reply(reply_token, text):
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(ReplyMessageRequest(
            reply_token=reply_token,
            messages=[TextMessage(text=text)]
        ))

def push_to(user_id, text):
    if not user_id:
        return
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.push_message(PushMessageRequest(
            to=user_id,
            messages=[TextMessage(text=text)]
        ))

def push_to_owner(text):
    push_to(OWNER_USER_ID, text)

def reply_flex(reply_token, alt_text, contents_dict):
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(ReplyMessageRequest(
            reply_token=reply_token,
            messages=[FlexMessage(alt_text=alt_text, contents=FlexContainer.from_dict(contents_dict))]
        ))

def push_flex_to(user_id, alt_text, contents_dict):
    if not user_id:
        return
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.push_message(PushMessageRequest(
            to=user_id,
            messages=[FlexMessage(alt_text=alt_text, contents=FlexContainer.from_dict(contents_dict))]
        ))

def push_flex_to_owner(alt_text, contents_dict):
    push_flex_to(OWNER_USER_ID, alt_text, contents_dict)

def reply_flex_and_text(reply_token, alt_text, contents_dict, text):
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(ReplyMessageRequest(
            reply_token=reply_token,
            messages=[
                FlexMessage(alt_text=alt_text, contents=FlexContainer.from_dict(contents_dict)),
                TextMessage(text=text),
            ]
        ))

def reply_text_flex_text(reply_token, lead_alt_text, lead_contents, alt_text, contents_dict, trailing_text):
    """contents_dict may be one bubble or a list of bubbles (sent one per message)."""
    bubbles = contents_dict if isinstance(contents_dict, list) else [contents_dict]
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(ReplyMessageRequest(
            reply_token=reply_token,
            messages=(
                [FlexMessage(alt_text=lead_alt_text, contents=FlexContainer.from_dict(lead_contents))]
                + [FlexMessage(alt_text=alt_text, contents=FlexContainer.from_dict(b)) for b in bubbles]
                + [TextMessage(text=trailing_text)]
            )
        ))

LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"

# Shared folder holding the inspection document files, linked under the
# @All mention. Override with INSPECTION_FILES_URL without a code change.
INSPECTION_FILES_URL = os.environ.get(
    "INSPECTION_FILES_URL",
    "https://firsttechno.storage-readyidc.com/index.php/s/WpypEPtTLXxMkt5",
)

def push_mention_all_to_group(group_id, flex_alt_text, flex_contents, lead_alt_text, lead_contents, mention_text="โปรดตรวจสอบและตอบกลับแอดมิน ส่วนรายการไฟล์เอกสารอยู่ในลิ้งด้านล่างค่ะ"):
    """Pushes a lead announcement Flex (so part of its text can be colored),
    then the documents Flex, then an '@All'-mention text message, to the
    given group. Uses a raw HTTP call instead of the messaging SDK's typed
    message classes because the installed line-bot-sdk version here doesn't
    expose LINE's textV2/mention payload - this builds that payload
    directly, per LINE's documented format."""
    bubbles = flex_contents if isinstance(flex_contents, list) else [flex_contents]
    messages = [
        {"type": "flex", "altText": lead_alt_text, "contents": lead_contents},
        *[{"type": "flex", "altText": flex_alt_text, "contents": b} for b in bubbles],
        {
            "type": "textV2",
            "text": f"{{user1}} {mention_text}"
                    + (f"\n{INSPECTION_FILES_URL}"
                       if INSPECTION_FILES_URL else ""),
            "substitution": {"user1": {"type": "mention", "mentionee": {"type": "all"}}},
        },
    ]
    resp = requests.post(
        LINE_PUSH_URL,
        headers={
            "Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}",
            "Content-Type": "application/json",
        },
        json={"to": group_id, "messages": messages},
        timeout=15,
    )
    app.logger.warning("inspection push to %s -> HTTP %s %s (request-id %s)",
                       group_id, resp.status_code, resp.text[:300],
                       resp.headers.get("X-Line-Request-Id"))
    resp.raise_for_status()

# ---------- Webhook ----------
@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

@handler.add(MessageEvent, message=(ImageMessageContent, FileMessageContent, VideoMessageContent, AudioMessageContent))
def handle_file(event):
    source = event.source
    if source.type != "group":
        return

    group_id = source.group_id
    if group_id in PM_MODE_GROUP_IDS:
        return  # PM-mode groups don't track files

    sender_id = source.user_id
    sender_name = get_profile_name(group_id, sender_id)
    remember_group_member(group_id, sender_id, sender_name)
    group_name = get_group_name(group_id)
    msg = event.message
    file_name = getattr(msg, "file_name", None) or msg.type
    now = bkk_now()
    now_iso = now.isoformat()
    cutoff = (now - datetime.timedelta(minutes=TAG_LOOKBACK_MINUTES)).isoformat()

    conn = db()
    # was this sender already tagged recently (tag arrived BEFORE the file)?
    pending = conn.execute(
        "SELECT id, text_content, mention_spans, tagged_user_ids FROM pending_tags WHERE group_id = ? AND sender_id = ? AND created_at >= ? ORDER BY created_at DESC LIMIT 1",
        (group_id, sender_id, cutoff)
    ).fetchone()
    tagged_now = 1 if pending else 0
    tag_text = pending[1] if pending else None
    tag_mention_spans = pending[2] if pending else None
    tag_user_ids = pending[3] if pending else None

    conn.execute(
        """INSERT INTO file_events
           (group_id, group_name, sender_id, sender_name, event_type, msg_type, file_name, text_content, sent_at, tagged, tagged_at, mention_spans, tagged_user_ids)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (group_id, group_name, sender_id, sender_name, "file", msg.type, file_name, tag_text, now_iso,
         tagged_now, now_iso if tagged_now else None, tag_mention_spans, tag_user_ids)
    )
    conn.commit()
    conn.close()

@handler.add(MessageEvent, message=TextMessageContent)
def handle_text(event):
    source = event.source
    text = event.message.text.strip()

    # --- 1:1 chat with the owner: on-demand commands ---
    if source.type == "user":
        user_id = source.user_id

        # --- Inspection document list: paste "code|doc_number|subject|date"
        # lines (one per document, as Claude in Chrome summarizes them from
        # Site Around) to get a Flex Carousel preview, then confirm/cancel. ---
        if text in ("ส่ง", "ยืนยัน", "confirm", "yes"):
            reply(event.reply_token, handle_inspection_confirm(user_id))
            return
        if text in ("แจ้งส่งผิดกลุ่ม", "ส่งผิดกลุ่ม"):
            reply(event.reply_token, handle_wrong_group_notice())
            return
        if text in ("ยกเลิก", "cancel"):
            if get_inspection_preview(user_id) is not None:
                clear_inspection_preview(user_id)
                reply(event.reply_token, "ยกเลิกรายการที่รอส่งแล้วครับ")
            return

        if text.startswith(TRACK_BACKFILL_CMD):
            reply(event.reply_token, handle_track_backfill(text))
            return

        if "|" in text:
            parsed_items = parse_inspection_paste(text)
            if parsed_items:
                alt_text, contents = build_inspection_carousel_flex(parsed_items)
                lead_alt_text, lead_contents = build_inspection_announce_flex(len(parsed_items))
                save_inspection_preview(user_id, parsed_items)
                reply_text_flex_text(
                    event.reply_token,
                    lead_alt_text, lead_contents,
                    alt_text, contents,
                    'ถ้าข้อมูลถูกต้อง พิมพ์ "ส่ง" เพื่อโพสต์เข้ากลุ่ม หรือ "ยกเลิก" เพื่อยกเลิกครับ'
                )
            else:
                reply(event.reply_token,
                      'รูปแบบไม่ถูกต้องครับ แต่ละบรรทัดต้องมีอย่างน้อย 4 ช่องคั่นด้วย "|" คือ '
                      'code|เลขที่เอกสาร|เรื่อง|วันที่ เช่น\n'
                      'AR-329|RIH-QA-0123|ขออนุมัติวัสดุ Shop Drawing|23/10/2026')
            return

        manual_answer = handle_manual_answer(user_id, text)
        if manual_answer is not None:
            reply(event.reply_token, manual_answer)
            return

        if text in ("สรุป", "สรุปไฟล์", "summary", "เช็คไฟล์"):
            # shared summary: everyone tracked sees the same tagged events,
            # regardless of who specifically got tagged
            alt_text, contents = build_summary_flex()
            if contents is None:
                reply(event.reply_token, alt_text)
            else:
                reply_flex(event.reply_token, alt_text, contents)
        elif text in ("id", "whoami", "ไอดี"):
            reply(event.reply_token, f"user_id ของคุณคือ:\n{source.user_id}")
        elif text in ("โหลดแล้ว", "✅", "👍"):
            reply(event.reply_token, handle_close_command("", source.user_id))
        elif text.startswith("โหลดแล้ว "):
            reply(event.reply_token, handle_close_command(text[len("โหลดแล้ว "):], source.user_id))
        elif text.startswith("✅") or text.startswith("👍"):
            reply(event.reply_token, handle_close_command(text[1:], source.user_id))
        elif text in ("คำตอบ", "ตอบกลับ", "replies"):
            alt_text, contents = build_replies_overview_flex()
            if contents is None:
                reply(event.reply_token, alt_text)
            else:
                reply_flex(event.reply_token, alt_text, contents)
        elif text in ("ติดตาม", "ติดตามคำตอบ"):
            alt_text, contents, who = build_followup_preview()
            if contents is None:
                reply(event.reply_token, alt_text)
            else:
                reply_flex_and_text(event.reply_token, alt_text, contents,
                                    f'{who}\n\nตัวอย่างข้อความติดตามที่จะส่งเข้ากลุ่มอัตโนมัติ {REMIND_HOUR:02d}:{REMIND_MINUTE:02d} น. '
                                    f'({REMIND_DAYS_OF_WEEK}) — ถ้าต้องการส่งตอนนี้เลย พิมพ์ "ส่งติดตาม" ครับ')
        elif text in ("บัดดี้", "ผู้รับผิดชอบ"):
            reply(event.reply_token, build_buddy_list_text())
        elif text.startswith("ลบบัดดี้ "):
            label = text[len("ลบบัดดี้ "):].strip().upper()
            reply(event.reply_token, f"ลบบัดดี้ {label} แล้วครับ" if delete_buddy_team(label)
                  else f"ไม่พบบัดดี้ชื่อ {label} ครับ (พิมพ์ \"บัดดี้\" เพื่อดูรายชื่อ)")
        elif text.startswith("ปิดเอกสาร "):
            reply(event.reply_token, set_doc_closed(text[len("ปิดเอกสาร "):], True))
        elif text.startswith("เปิดเอกสาร "):
            reply(event.reply_token, set_doc_closed(text[len("เปิดเอกสาร "):], False))
        elif text in ("ส่งติดตาม",):
            reply(event.reply_token, send_reply_followup(force=True))
        elif text.startswith("คำตอบ "):
            reply(event.reply_token, build_reply_thread_text(text[len("คำตอบ "):]))
        elif text in ("ย้อนกลับ", "เลิกทำ", "undo"):
            reply(event.reply_token, undo_last_close(source.user_id))
        elif text in ("รายงาน", "report", "export"):
            reply(event.reply_token,
                  f"ดาวน์โหลดรายงานวันนี้ (ไฟล์ Excel จัดสีสวยพร้อมเปิด):\n{daily_report_url()}\n\n"
                  f"หรือแบบ CSV ดิบ:\n{daily_report_url(fmt='csv')}\n\n"
                  f"ดูวันอื่นได้โดยเติม ?date=YYYY-MM-DD ต่อท้ายลิงก์")
        return

    if source.type != "group":
        return

    group_id = source.group_id
    remember_group_member(group_id, source.user_id)

    # utility: anyone can type "groupid" in a group to get its ID for setup
    if text.lower() in ("groupid", "ไอดีกลุ่ม"):
        reply(event.reply_token, f"group_id ของกลุ่มนี้คือ:\n{group_id}")
        return

    # set THIS group as where the inspection document list gets posted.
    # Owner-only, and the bot never replies in the group: the result goes to
    # the owner's 1:1 chat, so after unsending their own command message the
    # group shows no trace of it. Anyone else typing it is silently ignored.
    if text in ("ตั้งกลุ่มแจ้งเอกสาร", "ตั้งกลุ่มนี้แจ้งเอกสาร"):
        sender_id = source.user_id
        allowed = {OWNER_USER_ID} if OWNER_USER_ID else set(TAGGED_USER_IDS)
        if sender_id not in allowed:
            return
        try:
            set_setting("inspection_target_group_id", group_id)
            push_to(sender_id,
                    f"ตั้งกลุ่ม \"{get_group_name(group_id)}\" เป็นกลุ่มปลายทางสำหรับแจ้งเอกสาร Request แล้วค่ะ ✅\n"
                    "กดยกเลิกข้อความคำสั่งในกลุ่มได้เลย")
        except Exception:
            app.logger.exception("set inspection target group failed")
            try:
                push_to(sender_id, "ตั้งค่ากลุ่มไม่สำเร็จค่ะ ลองใหม่อีกครั้ง")
            except Exception:
                app.logger.exception("could not notify owner")
        return

    # set a buddy team (owner only; confirmation goes to the owner's 1:1 chat)
    if text.startswith("ตั้งบัดดี้"):
        try:
            handle_set_buddy_command(event, group_id, text)
        except Exception:
            app.logger.exception("set buddy team failed")
        return

    # shortcut reply: "AR-337 ตรวจแล้ว เหลือแก้รอยต่อฝ้า" saves a reply to
    # that posted document silently (no message back, to save LINE quota).
    # Doesn't return - a reply that also tags someone is still tracked below.
    try:
        record_group_reply_shortcut(group_id, source.user_id, text)
    except Exception:
        app.logger.exception("group reply shortcut failed")

    if not TAGGED_USER_IDS:
        return

    mention = getattr(event.message, "mention", None)
    if not mention or not mention.mentionees:
        return

    is_tagged = any(m.user_id in TAGGED_USER_IDS for m in mention.mentionees if getattr(m, "user_id", None))
    if not is_tagged:
        return

    sender_id = source.user_id
    sender_name = get_profile_name(group_id, sender_id)
    now = bkk_now()
    now_iso = now.isoformat()
    cutoff = (now - datetime.timedelta(minutes=TAG_LOOKBACK_MINUTES)).isoformat()

    # capture the character ranges of the @mention(s) within the message text,
    # so the summary can color the tagged name differently from the rest of the text.
    # LINE gives index/length in UTF-16 code units, but Python strings are indexed by
    # code point, so emoji outside the BMP (like 🐷, which is 2 UTF-16 units but 1
    # Python character) shift everything after them - convert before storing.
    spans = []
    for m in mention.mentionees:
        idx = getattr(m, "index", None)
        length = getattr(m, "length", None)
        if idx is not None and length is not None:
            py_start = utf16_index_to_py(text, idx)
            py_end = utf16_index_to_py(text, idx + length)
            spans.append([py_start, py_end - py_start])
    mention_spans_json = json.dumps(spans) if spans else None

    # which tracked person(s) were actually tagged in this message -
    # each admin's "สรุป" / daily digest only shows events where THEY were tagged
    tagged_ids = [m.user_id for m in mention.mentionees if getattr(m, "user_id", None) in TAGGED_USER_IDS]
    tagged_user_ids_json = json.dumps(tagged_ids)

    conn = db()

    if group_id in PM_MODE_GROUP_IDS:
        # PM mode: the tag message itself is the event, no file needed
        group_name = get_group_name(group_id)
        conn.execute(
            """INSERT INTO file_events
               (group_id, group_name, sender_id, sender_name, event_type, msg_type, text_content, sent_at, tagged, tagged_at, mention_spans, tagged_user_ids)
               VALUES (?,?,?,?,?,?,?,?,1,?,?,?)""",
            (group_id, group_name, sender_id, sender_name, "pm", "text", text, now_iso, now_iso, mention_spans_json, tagged_user_ids_json)
        )
    else:
        # file mode: match backward to recent un-tagged files from this sender,
        # and attach this tag's own message text to those files
        cur = conn.execute(
            """UPDATE file_events SET tagged = 1, tagged_at = ?, text_content = ?, mention_spans = ?, tagged_user_ids = ?
               WHERE group_id = ? AND sender_id = ? AND tagged = 0 AND sent_at >= ? AND event_type = 'file'""",
            (now_iso, text, mention_spans_json, tagged_user_ids_json, group_id, sender_id, cutoff)
        )
        # also leave a pending marker (with the message text) in case file(s) arrive AFTER this tag
        conn.execute(
            "INSERT INTO pending_tags (group_id, sender_id, text_content, mention_spans, tagged_user_ids, created_at) VALUES (?,?,?,?,?,?)",
            (group_id, sender_id, text, mention_spans_json, tagged_user_ids_json, now_iso)
        )

    conn.commit()
    conn.close()

# ---------- Inspection document list (paste from Claude in Chrome) ----------
def parse_inspection_paste(text):
    """Parses pasted lines of 4+ pipe-separated fields per document:
    code|doc_number|subject|date[|extra date/notes...], e.g.:
        AR-329|RIH-QA-0123|ขออนุมัติวัสดุ Shop Drawing|23/10/2026
    Any pipe-separated segments beyond the first 4 are folded into a
    separate extra_note field (joined with " / ") instead of rejecting the
    whole paste - Claude in Chrome sometimes adds a genuinely useful extra
    date (e.g. a concrete-pour date) past the confirmed 4-field template.
    Returns a list of {"code", "doc_number", "subject", "date", "extra_note"}
    dicts (extra_note is "" when the line had no 5th+ field).
    Returns None when NO line contains "|" at all, so a plain command like
    "สรุป" or "ยกเลิก" still falls through to the other 1:1 commands
    untouched. If a line DOES contain "|" but is missing a required field,
    the whole paste is also rejected (None) - the caller then tells the
    admin so, instead of staying silent."""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not lines or not any("|" in ln for ln in lines):
        return None
    items = []
    for ln in lines:
        parts = [p.strip() for p in ln.split("|")]
        if len(parts) < 4:
            return None
        code, doc_number, subject, date_str = parts[0], parts[1], parts[2], parts[3]
        extra_note = " / ".join(p for p in parts[4:] if p)
        if not code or not doc_number or not subject or not date_str:
            return None
        items.append({
            "code": code, "doc_number": doc_number, "subject": subject,
            "date": date_str, "extra_note": extra_note,
        })
    return items

INSPECTION_ANNOUNCE_HIGHLIGHT_COLOR = "#1565C0"  # the "จำนวน N ฉบับ" count phrase

def build_inspection_announce_flex(n):
    """A single-line Flex bubble (not a plain TextMessage) so the "จำนวน N
    ฉบับ" count phrase can be colored - LINE's plain text messages have no
    per-span styling, only Flex text components do."""
    contents = {
        "type": "bubble",
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "lg",
            "contents": [{
                "type": "text", "wrap": True, "contents": [
                    {"type": "span", "text": "📋✨ แจ้งข้อมูลเอกสาร Request จาก บมจ.พรีบิลท์ (PB) "},
                    {"type": "span", "text": f"จำนวน {n} ฉบับ",
                     "color": INSPECTION_ANNOUNCE_HIGHLIGHT_COLOR, "weight": "bold"},
                    {"type": "span", "text": " ดังนี้ค่ะ 📑"},
                ]
            }]
        }
    }
    alt_text = f"แจ้งข้อมูลเอกสาร Request จาก PB จำนวน {n} ฉบับ"
    return alt_text, contents

def save_inspection_preview(user_id, items):
    conn = db()
    conn.execute("DELETE FROM pending_inspection_previews WHERE user_id = ?", (user_id,))
    conn.execute(
        "INSERT INTO pending_inspection_previews (user_id, items_json, created_at) VALUES (?,?,?)",
        (user_id, json.dumps(items, ensure_ascii=False), bkk_now().isoformat())
    )
    conn.commit()
    conn.close()

def get_inspection_preview(user_id):
    conn = db()
    row = conn.execute(
        "SELECT items_json FROM pending_inspection_previews WHERE user_id = ?", (user_id,)
    ).fetchone()
    conn.close()
    if not row:
        return None
    try:
        return json.loads(row[0])
    except (ValueError, TypeError):
        return None

def clear_inspection_preview(user_id):
    conn = db()
    conn.execute("DELETE FROM pending_inspection_previews WHERE user_id = ?", (user_id,))
    conn.commit()
    conn.close()

INSPECTION_LABEL_COLOR = "#90A4AE"
INSPECTION_VALUE_COLOR = "#000000"       # Code / เรื่อง / วันที่ส่งเอกสารเข้าระบบ - plain black
INSPECTION_DOCNUM_COLOR = "#C2571A"      # เลขที่เอกสาร (เช่น AR-333) - bold, burnt orange
INSPECTION_DOCNUM_BG_COLOR = "#FFF8E7"   # warm cream highlight behind it
INSPECTION_DOCNUM_BORDER_COLOR = "#F5DDB5"  # soft beige outline
INSPECTION_DOCNUM_LABEL_COLOR = "#7A2E12"  # label inside the badge - deep brown-red
INSPECTION_EXTRA_NOTE_COLOR = "#EF6C00"  # วันที่เทคอนกรีต/ก่อ/ฉาบ/อื่นๆ - kept colored
INSPECTION_SEPARATOR_COLOR = "#0D47A1"   # divider line between documents - dark blue

INSPECTION_MAX_ITEMS = 30  # safety cap so the flex message doesn't blow past LINE's size limit
INSPECTION_BUBBLE_MAX_BYTES = 28000  # LINE's per-bubble cap is 30 KB; keep some headroom
# a push/reply carries at most 5 messages: lead + document bubbles + closing message
INSPECTION_MAX_BUBBLES = 3

def build_inspection_carousel_flex(items):
    """A SINGLE bubble listing every document stacked vertically, one block
    per document (separated by a divider line), each block's own fields
    also stacked vertically as an emoji + label line followed by the value,
    color-coded per field. Deliberately not a LINE "carousel" (separate
    swipeable cards) - with everything in one bubble, the whole list reads
    top-to-bottom in a single scroll instead of needing to swipe sideways
    between documents."""
    txt = flex_txt

    def field(emoji, label, value, color, bold=False):
        return {
            "type": "box", "layout": "vertical", "margin": "md", "spacing": "xs",
            "contents": [
                txt(f"{emoji} {label}", size="xs", color=INSPECTION_LABEL_COLOR, weight="bold"),
                txt(value, size="xs", color=color, wrap=True, weight="bold" if bold else "regular"),
            ]
        }

    def doc_number_badge(value):
        """เลขที่เอกสาร gets its own highlighted pill (background color box)
        instead of the plain field() layout, to stand out as the key value."""
        return {
            "type": "box", "layout": "vertical", "margin": "md", "spacing": "xs",
            "backgroundColor": INSPECTION_DOCNUM_BG_COLOR, "cornerRadius": "md", "paddingAll": "md",
            "borderWidth": "1px", "borderColor": INSPECTION_DOCNUM_BORDER_COLOR,
            "contents": [
                txt("📌 เลขที่เอกสาร", size="xs", color=INSPECTION_DOCNUM_LABEL_COLOR, weight="bold"),
                txt(value, size="lg", color=INSPECTION_DOCNUM_COLOR, wrap=True, weight="bold"),
            ]
        }

    shown = items[:INSPECTION_MAX_ITEMS]
    doc_blocks = []
    for it in shown:
        body = []
        body.append(doc_number_badge(it["code"]))
        body.append(field("🔖", "Code", it["doc_number"], INSPECTION_VALUE_COLOR))
        body.append(field("📝", "เรื่อง", it["subject"], INSPECTION_VALUE_COLOR))
        body.append(field("📅", "วันที่ส่งเอกสารเข้าระบบ", it["date"], INSPECTION_VALUE_COLOR))
        if it.get("extra_note"):
            body.append(field("🧱", "วันที่เทคอนกรีต/ก่อ/ฉาบ/อื่นๆ", it["extra_note"],
                               INSPECTION_EXTRA_NOTE_COLOR))
        reply_link = doc_reply_liff_url(it["code"])
        if reply_link:
            body.append({
                "type": "text", "text": "💬 ตอบกลับเอกสารนี้", "size": "xs", "weight": "bold",
                "color": REPLY_LINK_COLOR, "margin": "md", "decoration": "underline",
                "action": {"type": "uri", "label": "ตอบกลับ", "uri": reply_link},
            })
        doc_blocks.append(body)

    if len(items) > INSPECTION_MAX_ITEMS:
        doc_blocks[-1] = doc_blocks[-1] + [
            {"type": "separator", "margin": "xl"},
            txt(f"… มีอีก {len(items) - INSPECTION_MAX_ITEMS} รายการที่ไม่ได้แสดง",
                size="xs", color=INSPECTION_LABEL_COLOR, margin="md"),
        ]

    # LINE rejects a bubble whose JSON is over 30 KB (Thai text is 3 bytes a
    # character, so ~16 documents already reach that). Pack documents into as
    # few bubbles as fit; they're sent as separate messages one under another,
    # so the list still reads top-to-bottom without sideways swiping.
    def bubble(blocks):
        body = []
        for i, b in enumerate(blocks):
            if i > 0:
                body.append({"type": "separator", "margin": "xl", "color": INSPECTION_SEPARATOR_COLOR})
            body.extend(b)
        return {"type": "bubble",
                "body": {"type": "box", "layout": "vertical", "spacing": "sm", "contents": body, "paddingAll": "lg"}}

    def size_of(b):
        return len(json.dumps(b, ensure_ascii=False, separators=(",", ":")).encode("utf-8"))

    bubbles, current = [], []
    for block in doc_blocks:
        if current and size_of(bubble(current + [block])) > INSPECTION_BUBBLE_MAX_BYTES:
            bubbles.append(bubble(current))
            current = []
        current.append(block)
    if current:
        bubbles.append(bubble(current))
    if len(bubbles) > INSPECTION_MAX_BUBBLES:
        app.logger.warning("inspection list needs %d bubbles, only %d sent", len(bubbles), INSPECTION_MAX_BUBBLES)
        bubbles = bubbles[:INSPECTION_MAX_BUBBLES]
        bubbles[-1]["body"]["contents"].append(
            txt("… รายการยาวเกินกว่าที่ LINE ส่งได้ในครั้งเดียว กรุณาแบ่งส่งเป็นสองรอบ",
                size="xs", color=OVERDUE_COLOR, margin="md"))

    alt_text = f"เอกสารตรวจสอบ {len(items)} รายการ"
    return alt_text, bubbles

def handle_inspection_confirm(user_id):
    items = get_inspection_preview(user_id)
    if not items:
        return 'ไม่มีรายการที่รอยืนยันอยู่ครับ ลองวางข้อมูลเอกสารรูปแบบ "code|doc_number|subject|date" ก่อนนะครับ'
    target = inspection_target_group_id()
    if not target:
        return ('ยังไม่ได้ตั้งกลุ่มปลายทางครับ ให้พิมพ์ "ตั้งกลุ่มแจ้งเอกสาร" ในกลุ่มที่ต้องการให้โพสต์ '
                'แล้วกลับมาพิมพ์ "ส่ง" อีกครั้ง (รายการยังรออยู่)')
    alt_text, contents = build_inspection_carousel_flex(items)
    lead_alt_text, lead_contents = build_inspection_announce_flex(len(items))
    try:
        push_mention_all_to_group(target, alt_text, contents, lead_alt_text, lead_contents)
    except Exception as e:
        app.logger.exception("inspection push failed")
        detail = ""
        resp = getattr(e, "response", None)
        if resp is not None:
            detail = f"\n(LINE ตอบกลับ {resp.status_code}: {resp.text[:200]})"
        return ("โพสต์เข้ากลุ่มไม่สำเร็จครับ ลองใหม่อีกครั้ง หรือพิมพ์ \"ตั้งกลุ่มแจ้งเอกสาร\" "
                "ในกลุ่มที่ต้องการก่อน" + detail)
    clear_inspection_preview(user_id)
    try:
        save_inspection_docs(items, target)
    except Exception:
        app.logger.exception("could not record posted inspection docs for replies")
    try:
        set_setting("last_inspection_post_group_id", target)
    except Exception:
        app.logger.exception("could not record last inspection post group")
    return f"โพสต์เข้ากลุ่ม \"{get_group_name(target)}\" เรียบร้อยแล้วครับ"

WRONG_GROUP_APOLOGY = ("🙏 ขออภัยค่ะ รายการแจ้งเอกสาร Request ด้านบนส่งผิดกลุ่ม "
                       "รบกวนไม่ต้องดำเนินการค่ะ ขอบคุณค่ะ")

def handle_wrong_group_notice():
    """Pushes an apology to the group the inspection list was last posted
    to. LINE's Messaging API can't unsend a bot's message, so this is the
    closest thing. Posts made before the last group was being recorded fall
    back to the INSPECTION_TARGET_GROUP_ID env var, which is where those
    went."""
    target = get_setting("last_inspection_post_group_id", INSPECTION_TARGET_GROUP_ID)
    if not target:
        return "ไม่พบกลุ่มที่โพสต์ล่าสุดครับ"
    try:
        resp = requests.post(
            LINE_PUSH_URL,
            headers={"Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}",
                     "Content-Type": "application/json"},
            json={"to": target, "messages": [{"type": "text", "text": WRONG_GROUP_APOLOGY}]},
            timeout=15,
        )
        app.logger.warning("wrong-group notice to %s -> HTTP %s %s", target, resp.status_code, resp.text[:300])
        resp.raise_for_status()
    except Exception as e:
        resp = getattr(e, "response", None)
        detail = f"\n(LINE ตอบกลับ {resp.status_code}: {resp.text[:200]})" if resp is not None else ""
        return "ส่งข้อความแจ้งขออภัยไม่สำเร็จครับ" + detail
    return f'ส่งข้อความแจ้งขออภัยเข้ากลุ่ม "{get_group_name(target)}" แล้วครับ'

# ---------- Document replies (team answers to posted inspection docs) ----------
# Each posted document gets a reply thread. Team members add to it either by
# typing "<code> <message>" in the group (e.g. "AR-337 ตรวจแล้ว") or through
# the LIFF page opened from the card's "💬 ตอบกลับเอกสารนี้" link. Nothing is
# pushed when a reply arrives; the owner checks with the 1:1 "คำตอบ" command.
import re
import time

REPLY_LINK_COLOR = "#1565C0"
REPLY_MAX_LEN = 1000
REPLIES_OVERVIEW_DAYS = int(os.environ.get("REPLIES_OVERVIEW_DAYS", "14"))
REPLIES_OVERVIEW_MAX_DOCS = 40

# "AR-337 ข้อความ", "st-802: ข้อความ", "AR 337 - ข้อความ" (code at the very start)
REPLY_SHORTCUT_RE = re.compile(r"^\s*([A-Za-z]{1,5})\s*-?\s*(\d{1,5})\s*[:：\-–]?\s+(.+)$", re.DOTALL)

def normalize_doc_code(raw):
    m = re.match(r"^\s*([A-Za-z]{1,5})\s*-?\s*(\d{1,5})\s*$", raw or "")
    if not m:
        return None
    return f"{m.group(1).upper()}-{int(m.group(2))}"

def doc_reply_liff_url(code):
    if not LIFF_ID:
        return None
    return f"https://liff.line.me/{LIFF_ID}?doc={requests.utils.quote(code)}"

def save_inspection_docs(items, group_id):
    now_iso = bkk_now().isoformat()
    conn = db()
    try:
        rows = []
        for it in items:
            code = normalize_doc_code(it["code"]) or it["code"].strip().upper()
            rows.append((code, it["doc_number"], it["subject"], it["date"],
                         it.get("extra_note", ""), group_id, now_iso))
        conn.executemany(
            "INSERT INTO inspection_docs (code, doc_number, subject, date, extra_note, group_id, posted_at) "
            "VALUES (?,?,?,?,?,?,?) ON CONFLICT(code) DO UPDATE SET "
            "doc_number = excluded.doc_number, subject = excluded.subject, date = excluded.date, "
            "extra_note = excluded.extra_note, group_id = excluded.group_id, posted_at = excluded.posted_at",
            rows,
        )
    finally:
        conn.close()

# --- Add earlier documents to the reply follow-up (1:1, nothing posted) ---
#   ติดตามย้อนหลัง AR-337 ST-802 ...            codes already posted by the bot
#   ติดตามย้อนหลัง 27/10/2026                    first line (date optional = when
#   AR-337|RIH-I02351|เรื่อง|24/10/2026          they were posted to the group),
#   ...                                          then the usual pipe lines
# Existing docs keep their original posted date (so "รอ X วัน" stays right) and
# are reopened if closed; unknown ones are inserted with the given date (or now).
TRACK_BACKFILL_CMD = "ติดตามย้อนหลัง"

def _parse_thai_or_iso_date(raw):
    raw = (raw or "").strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"):
        try:
            d = datetime.datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
        if d.year > 2400:  # Buddhist year
            d = d.replace(year=d.year - 543)
        return d
    return None

def handle_track_backfill(text):
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    head = lines[0][len(TRACK_BACKFILL_CMD):].strip()
    body = lines[1:]
    posted_date = None
    head_refs = []
    for tok in re.split(r"[\s,]+", head):
        if not tok:
            continue
        d = _parse_thai_or_iso_date(tok)
        if d:
            posted_date = d
        else:
            head_refs.append(tok)
    pipe_items = parse_inspection_paste("\n".join(body)) if body else None
    if body and pipe_items is None:
        head_refs += [r for ln in body for r in re.split(r"[\s,]+", ln) if r]
    if not head_refs and not pipe_items:
        return (f'วิธีใช้:\n{TRACK_BACKFILL_CMD} AR-337 ST-802\n'
                f'หรือ (กรณีบอทไม่มีข้อมูลเอกสาร)\n{TRACK_BACKFILL_CMD} 27/10/2026\n'
                'AR-337|RIH-I02351|เรื่อง|24/10/2026\n...\n'
                '(วันที่บรรทัดแรก = วันที่เคยโพสต์เข้ากลุ่ม ไม่ใส่ก็ได้)')
    if posted_date:
        posted_iso = datetime.datetime.combine(posted_date, datetime.time(17, 0)).isoformat()
    else:
        posted_iso = bkk_now().isoformat()
    group_id = inspection_target_group_id() or ""
    added, reopened, kept, unknown = [], [], [], []
    conn = db()
    try:
        for it in (pipe_items or []):
            code = normalize_doc_code(it["code"]) or it["code"].strip().upper()
            row = conn.execute("SELECT closed_at FROM inspection_docs WHERE code = ?", (code,)).fetchone()
            if row:
                conn.execute("UPDATE inspection_docs SET closed_at = NULL WHERE code = ?", (code,))
                (reopened if row[0] else kept).append(code)
            else:
                conn.execute(
                    "INSERT INTO inspection_docs (code, doc_number, subject, date, extra_note, group_id, posted_at) "
                    "VALUES (?,?,?,?,?,?,?)",
                    (code, it["doc_number"], it["subject"], it["date"], it.get("extra_note", ""),
                     group_id, posted_iso))
                added.append(code)
        for ref in head_refs:
            code = normalize_doc_code(ref)
            row = conn.execute("SELECT closed_at FROM inspection_docs WHERE code = ?",
                               (code,)).fetchone() if code else None
            if not row:
                unknown.append(ref)
                continue
            conn.execute("UPDATE inspection_docs SET closed_at = NULL WHERE code = ?", (code,))
            (reopened if row[0] else kept).append(code)
        if posted_date:
            # an explicit date also corrects docs whose posted date was lost or reset
            for code in kept + reopened:
                conn.execute("UPDATE inspection_docs SET posted_at = ? WHERE code = ?", (posted_iso, code))
    finally:
        conn.close()
    touched = set(added + reopened + kept)
    status = {d["code"]: d for d in get_docs_status(REMIND_WINDOW_DAYS) if d["code"] in touched}
    out = [f"📌 เพิ่มเข้ารายการติดตามแล้ว {len(touched)} ฉบับ (ไม่ได้โพสต์เข้ากลุ่ม)"]
    for code in sorted(touched, key=lambda c: (doc_prefix(c), int(c.split("-")[1]) if c.split("-")[1].isdigit() else 0)):
        d = status.get(code)
        tag = " (ใหม่)" if code in added else (" (เปิดคืน)" if code in reopened else "")
        if not d:
            out.append(f"• {code}{tag} — เก่าเกิน {REMIND_WINDOW_DAYS} วัน จะไม่ขึ้นในรายการติดตาม")
        elif d["pending"]:
            out.append(f"• {code}{tag} — 🟠 รอตอบ: " + (team_status_line(d["status"]) if d["status"] else "ยังไม่มีคนตอบ"))
        else:
            out.append(f"• {code}{tag} — ✅ ตอบครบแล้ว" + (f" ({team_status_line(d['status'])})" if d["status"] else ""))
    if unknown:
        out.append(f"\n❓ ไม่พบในระบบบอท: {', '.join(unknown)}\n"
                   f"ให้ส่งแบบมีรายละเอียดแทน:\n{TRACK_BACKFILL_CMD} <วันที่โพสต์>\ncode|เลขที่เอกสาร|เรื่อง|วันที่")
    out.append('\nดูตัวอย่างข้อความติดตาม: พิมพ์ "ติดตาม"')
    return "\n".join(out)

def get_inspection_doc(code):
    conn = db()
    try:
        row = conn.execute(
            "SELECT code, doc_number, subject, date, extra_note, posted_at FROM inspection_docs WHERE code = ?",
            (code,),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return None
    return dict(zip(("code", "doc_number", "subject", "date", "extra_note", "posted_at"), row))

def add_doc_reply(code, user_id, user_name, text, source):
    text = (text or "").strip()
    if not text:
        return False
    conn = db()
    try:
        conn.execute(
            "INSERT INTO doc_replies (code, user_id, user_name, text, source, created_at) VALUES (?,?,?,?,?,?)",
            (code, user_id, user_name, text[:REPLY_MAX_LEN], source, bkk_now().isoformat()),
        )
    finally:
        conn.close()
    return True

def get_doc_replies(code):
    conn = db()
    try:
        rows = conn.execute(
            "SELECT user_name, text, source, created_at, user_id FROM doc_replies WHERE code = ? ORDER BY created_at, id",
            (code,),
        ).fetchall()
    finally:
        conn.close()
    return [dict(zip(("user_name", "text", "source", "created_at", "user_id"), r)) for r in rows]

def record_group_reply_shortcut(group_id, user_id, text):
    """Returns True when the text was saved as a reply. Only codes that were
    actually posted count, so ordinary chat that happens to start with
    something like "EE-1 ..." isn't swallowed by mistake."""
    m = REPLY_SHORTCUT_RE.match(text or "")
    if not m:
        return False
    code = f"{m.group(1).upper()}-{int(m.group(2))}"
    if not get_inspection_doc(code):
        return False
    name = get_profile_name(group_id, user_id) if user_id else "ไม่ทราบชื่อ"
    saved = add_doc_reply(code, user_id, name, m.group(3), "group")
    if saved:
        after_doc_reply(code, user_id, name, m.group(3).strip())
    return saved

def _fmt_dt(iso):
    try:
        return datetime.datetime.fromisoformat(iso).strftime("%d/%m %H:%M")
    except (ValueError, TypeError):
        return ""

def build_replies_overview_flex():
    """1:1 "คำตอบ": recently posted documents, the ones still waiting first,
    with which buddy teams have answered and the latest reply."""
    docs = list(reversed(get_docs_status(REPLIES_OVERVIEW_DAYS)))  # newest first
    if not docs:
        return (f"ยังไม่มีเอกสารที่โพสต์เข้ากลุ่มในช่วง {REPLIES_OVERVIEW_DAYS} วันที่ผ่านมาครับ", None)

    pending = [d for d in docs if d["pending"]]
    done = [d for d in docs if not d["pending"]]
    txt = flex_txt
    body = [txt("💬 คำตอบเอกสาร", weight="bold", size="lg", color="#111111"),
            txt(f"{REPLIES_OVERVIEW_DAYS} วันล่าสุด · ทั้งหมด {len(docs)} ฉบับ · "
                f"ยังรอคำตอบ {len(pending)} ฉบับ",
                size="xs", color="#888888", margin="xs"),
            {"type": "separator", "margin": "md"}]

    shown = 0
    def doc_block(d, color):
        if d["closed"]:
            suffix = "  (🔒 ปิดโดยแอดมิน)"
        else:
            suffix = f"  ({d['n']} คำตอบ)" if d["n"] else "  (ยังไม่มีคำตอบ)"
        block = [two_tone_text(d["code"], color, suffix, size="sm", margin="md")]
        block.append(txt(d["subject"] or "-", size="xs", color="#444444", margin="none"))
        if d["status"]:
            block.append(txt(team_status_line(d["status"]), size="xs", color="#333333", margin="none"))
        if d["latest"]:
            preview = d["latest"].replace("\n", " ")
            block.append(txt(f"↳ {preview[:120]}{'…' if len(preview) > 120 else ''}",
                             size="xs", color="#0B6E4F", margin="none"))
        return block

    if pending:
        body.append(txt("🔴 ยังรอคำตอบ", weight="bold", size="sm", color=OVERDUE_COLOR, margin="md"))
        for d in pending:
            if shown >= REPLIES_OVERVIEW_MAX_DOCS:
                break
            body.extend(doc_block(d, OVERDUE_COLOR))
            shown += 1
    if done and shown < REPLIES_OVERVIEW_MAX_DOCS:
        body.append({"type": "separator", "margin": "lg"})
        body.append(txt("✅ ตอบครบแล้ว", weight="bold", size="sm", color="#0B6E4F", margin="md"))
        for d in done:
            if shown >= REPLIES_OVERVIEW_MAX_DOCS:
                break
            body.extend(doc_block(d, "#0D47A1"))
            shown += 1
    if shown < len(docs):
        body.append(txt(f"… มีอีก {len(docs) - shown} ฉบับที่ไม่ได้แสดง", size="xs", color="#999999", margin="md"))

    body.append({"type": "separator", "margin": "lg"})
    body.append(txt('ดูคำตอบทั้งหมดของเอกสาร: พิมพ์ "คำตอบ AR-337" · ปิดเอง: "ปิดเอกสาร AR-337"',
                    size="xs", color="#888888", margin="md"))
    contents = {"type": "bubble",
                "body": {"type": "box", "layout": "vertical", "spacing": "xs", "contents": body}}
    return (f"คำตอบเอกสาร: ยังรอคำตอบ {len(pending)} จาก {len(docs)} ฉบับ", contents)

def build_reply_thread_text(raw_code):
    code = normalize_doc_code(raw_code)
    if not code:
        return 'พิมพ์รหัสเอกสารให้ถูกรูปแบบครับ เช่น "คำตอบ AR-337"'
    doc = get_inspection_doc(code)
    if not doc:
        return f"ไม่พบเอกสาร {code} ในรายการที่เคยโพสต์เข้ากลุ่มครับ"
    replies = get_doc_replies(code)
    lines = [f"📌 {code} · {doc['subject']}", f"ส่งเข้าระบบ {doc['date']}", ""]
    if not replies:
        lines.append("ยังไม่มีคำตอบครับ")
    for r in replies:
        lines.append(f"👤 {r['user_name']} ({_fmt_dt(r['created_at'])} น.)")
        lines.append("   " + r["text"].replace("\n", "\n   "))
    text = "\n".join(lines)
    return text if len(text) <= 4900 else text[:4900] + "\n…(ยาวเกิน แสดงไม่ครบ)"

def count_unanswered_docs():
    return sum(1 for d in get_docs_status(REPLIES_OVERVIEW_DAYS) if d["pending"])

# --- Buddy teams: who must answer each posted document ---
# The owner sets a team per system in the group, picking names from LINE's
# mention popup so the bot gets each person's user ID (display names can
# change freely - LINE always renders the current name in a mention):
#     ตั้งบัดดี้ AR,ST EE-A @นายA @นายB
#     ตั้งบัดดี้ AR,ST AC @นายC @นายD
# A document whose code prefix (AR, ST, ...) has teams needs an answer from
# every SYSTEM; within a team one member's reply counts for the whole team.
# Teams named "<system>-<building>" (EE-A, EE-B) split one system by building:
# the subject is searched for "อาคาร A"/"ตึก B"/"Tower B"/...; one building found ->
# only that team counts, none or both found -> either team's reply counts.
# A reply from someone outside the doc's teams (and not the owner) is kept
# but doesn't count; the owner gets a 1:1 heads-up and may close the doc
# with "ปิดเอกสาร <code>". Prefixes with no team keep the old rule (any reply).
BUILDING_RE = re.compile(
    r"(?:อาคาร|ตึก|tower|building|bldg\.?|bld\.?|zone|โซน)\s*[-:.]?\s*([AB])(?![A-Za-z])", re.IGNORECASE)
TEAM_SPLIT_RE = re.compile(r"^(.+)-([A-Z0-9])$")

def detect_buildings(subject):
    return {b.upper() for b in BUILDING_RE.findall(subject or "")}

def team_system(label):
    m = TEAM_SPLIT_RE.match(label)
    return (m.group(1), m.group(2)) if m else (label, None)

def doc_prefix(code):
    return (code or "").split("-")[0].upper()

def load_buddy_teams():
    conn = db()
    try:
        rows = conn.execute(
            "SELECT label, prefixes, member_ids, group_id FROM buddy_teams ORDER BY label").fetchall()
    finally:
        conn.close()
    teams = []
    for label, prefixes, member_ids, group_id in rows:
        system, building = team_system(label)
        teams.append({"label": label, "system": system, "building": building,
                      "prefixes": {p for p in (prefixes or "").split(",") if p},
                      "members": json.loads(member_ids or "[]"), "group_id": group_id})
    return teams

def save_buddy_team(label, prefixes, members, group_id):
    conn = db()
    try:
        conn.execute(
            "INSERT INTO buddy_teams (label, prefixes, member_ids, group_id, updated_at) VALUES (?,?,?,?,?) "
            "ON CONFLICT(label) DO UPDATE SET prefixes = excluded.prefixes, member_ids = excluded.member_ids, "
            "group_id = excluded.group_id, updated_at = excluded.updated_at",
            (label, ",".join(sorted(prefixes)), json.dumps(members), group_id, bkk_now().isoformat()))
    finally:
        conn.close()

def delete_buddy_team(label):
    conn = db()
    try:
        found = conn.execute("SELECT 1 FROM buddy_teams WHERE label = ?", (label,)).fetchone()
        conn.execute("DELETE FROM buddy_teams WHERE label = ?", (label,))
    finally:
        conn.close()
    return bool(found)

# Systems every doc of a prefix must be answered by, even if that system's
# buddy team hasn't been set up yet - otherwise one team's reply would make
# the doc look fully answered and it would drop out of the follow-up.
# Format: "AR,ST=EE,AC,SN/FP;XX=EE"
def _parse_required_systems(raw):
    out = {}
    for part in (raw or "").split(";"):
        if "=" not in part:
            continue
        prefixes, systems = part.split("=", 1)
        sys_list = [x.strip().upper() for x in systems.split(",") if x.strip()]
        for pre in prefixes.split(","):
            if pre.strip():
                out[pre.strip().upper()] = sys_list
    return out

REQUIRED_SYSTEMS = _parse_required_systems(os.environ.get("REQUIRED_SYSTEMS", "AR,ST=EE,AC,SN/FP"))

def doc_team_status(code, subject, teams, replier_ids):
    """[{label, done, tag_ids}] per system this doc needs, or None when no
    team covers the doc's prefix (old rule: any reply answers it)."""
    prefix = doc_prefix(code)
    applicable = [t for t in teams if prefix in t["prefixes"]]
    required = REQUIRED_SYSTEMS.get(prefix, [])
    if not applicable and not required:
        return None
    systems = []
    for t in applicable:
        if t["system"] not in systems:
            systems.append(t["system"])
    out = []
    for system in systems:
        subs = [t for t in applicable if t["system"] == system]
        need = subs
        if len(subs) > 1 or subs[0]["building"]:
            found = detect_buildings(subject)
            named = [t for t in subs if t["building"] in found]
            if len(found) == 1 and named:
                need = named
        if len(need) == 1:
            label = need[0]["label"]
        else:
            label = f"{system} ({'/'.join(t['building'] or t['label'] for t in need)})"
        members = [u for t in need for u in t["members"]]
        done = bool(set(members) & replier_ids) or any(team_marker(t["label"]) in replier_ids for t in need)
        out.append({"label": label, "done": done, "tag_ids": [] if done else members,
                    "teams": [t["label"] for t in need]})
    for system in required:
        if system not in systems:
            # no buddy team yet: can only be cleared by "ปิดเอกสาร" or by setting the team
            out.append({"label": f"{system} (ยังไม่ตั้งบัดดี้)", "done": False, "tag_ids": [],
                        "teams": [], "missing": True})
    return out

def team_marker(label):
    """user_id stored on a reply the owner records for a whole team."""
    return f"team:{label}"

def get_docs_status(days):
    """Every doc posted in the last `days` days with its reply state, oldest first."""
    since = (bkk_now() - datetime.timedelta(days=days)).isoformat()
    conn = db()
    try:
        docs = conn.execute(
            "SELECT code, subject, date, posted_at, group_id, closed_at FROM inspection_docs "
            "WHERE posted_at >= ? ORDER BY posted_at, code", (since,)).fetchall()
        replies = conn.execute(
            "SELECT r.code, r.user_id, r.user_name, r.text FROM doc_replies r "
            "JOIN inspection_docs d ON d.code = r.code WHERE d.posted_at >= ? "
            "ORDER BY r.created_at, r.id", (since,)).fetchall()
    finally:
        conn.close()
    by_code = {}
    try:
        name_map = load_member_name_map()
    except Exception:
        app.logger.exception("load member names failed")
        name_map = {}
    for code, uid, name, text in replies:
        info = by_code.setdefault(code, {"ids": set(), "n": 0, "latest": None})
        info["ids"] |= reply_identity_ids(uid, name, name_map)
        info["n"] += 1
        info["latest"] = f"{name}: {text}"
    teams = load_buddy_teams()
    out = []
    for code, subject, date, posted_at, group_id, closed_at in docs:
        info = by_code.get(code, {"ids": set(), "n": 0, "latest": None})
        status = doc_team_status(code, subject, teams, info["ids"])
        if closed_at:
            pending = False
        elif status is None:
            pending = info["n"] == 0
        else:
            pending = not all(s["done"] for s in status)
        out.append({"code": code, "subject": subject, "date": date, "posted_at": posted_at,
                    "group_id": group_id or "", "closed": bool(closed_at), "status": status,
                    "n": info["n"], "latest": info["latest"], "pending": pending})
    return out

def team_status_line(status):
    return "   ".join(f"{'✅' if s['done'] else '⏳'} {s['label']}" for s in status)

def set_doc_closed(raw_code, closed):
    code = normalize_doc_code(raw_code)
    if not code:
        return 'พิมพ์รหัสเอกสารให้ถูกรูปแบบครับ เช่น "ปิดเอกสาร AR-337"'
    if not get_inspection_doc(code):
        return f"ไม่พบเอกสาร {code} ในรายการที่เคยโพสต์เข้ากลุ่มครับ"
    conn = db()
    try:
        conn.execute("UPDATE inspection_docs SET closed_at = ? WHERE code = ?",
                     (bkk_now().isoformat() if closed else None, code))
    finally:
        conn.close()
    if closed:
        return f"🔒 ปิดเอกสาร {code} แล้วครับ จะไม่อยู่ในรายการติดตามอีก (เปิดคืน: \"เปิดเอกสาร {code}\")"
    return f"🔓 เปิดเอกสาร {code} คืนแล้วครับ"

def notify_owner_outsider_reply(code, user_id, user_name, text):
    """Heads-up to the owner when someone outside the doc's buddy teams replies."""
    if not OWNER_USER_ID or user_id == OWNER_USER_ID:
        return
    doc = get_inspection_doc(code)
    if not doc:
        return
    applicable = [t for t in load_buddy_teams() if doc_prefix(code) in t["prefixes"]]
    ids = reply_identity_ids(user_id, user_name, load_member_name_map())
    if not applicable or any(ids & set(t["members"]) for t in applicable):
        return
    preview = text if len(text) <= 300 else text[:300] + "…"
    push_to_owner(
        f"📝 มีคำตอบจากคนนอกทีมผู้รับผิดชอบ\n{code} · {doc['subject']}\n"
        f"👤 {user_name}: {preview}\n\n"
        f"คำตอบนี้ยังไม่นับให้ระบบใด ถ้าตรวจแล้วต้องการปิดเอกสาร พิมพ์ \"ปิดเอกสาร {code}\" ครับ")

def after_doc_reply(code, user_id, user_name, text):
    try:
        notify_owner_outsider_reply(code, user_id, user_name, text)
    except Exception:
        app.logger.exception("outsider reply notice failed")

def handle_set_buddy_command(event, group_id, text):
    """Group: "ตั้งบัดดี้ AR,ST EE-A @a @b" (owner only, confirmation goes 1:1).
    Several teams can be set in one message, one "ตั้งบัดดี้ ..." per line;
    each line only gets the @mentions written on that line."""
    sender_id = event.source.user_id
    allowed = {OWNER_USER_ID} if OWNER_USER_ID else set(TAGGED_USER_IDS)
    if sender_id not in allowed:
        return
    mention = getattr(event.message, "mention", None)
    mentionees = list(mention.mentionees) if mention and mention.mentionees else []
    spans = []  # (py_start, py_end, mentionee)
    for m in mentionees:
        if getattr(m, "index", None) is not None and getattr(m, "length", None) is not None:
            spans.append((utf16_index_to_py(text, m.index),
                          utf16_index_to_py(text, m.index + m.length), m))
    usage = ('รูปแบบ: ตั้งบัดดี้ AR,ST EE-A @ชื่อ1 @ชื่อ2\n'
             '(ตั้งหลายทีมในข้อความเดียวได้ บรรทัดละ 1 ทีม · เลือกชื่อจากรายชื่อที่ LINE เด้งขึ้นมา)')
    results, pos = [], 0
    for line in text.split("\n"):
        line_start, line_end = pos, pos + len(line)
        pos = line_end + 1
        if not line.strip().startswith("ตั้งบัดดี้"):
            continue
        own = [(a, b, m) for a, b, m in spans if line_start <= a < line_end]
        plain = line
        for a, b, _ in sorted(own, key=lambda x: x[0], reverse=True):
            plain = plain[:a - line_start] + " " + plain[b - line_start:]
        tokens = plain.split()
        if len(tokens) < 3:
            results.append(f"❌ \"{line.strip()}\" ข้อมูลไม่ครบ")
            continue
        prefixes = {p.strip().upper() for p in tokens[1].split(",") if p.strip()}
        label = tokens[2].upper()
        if not prefixes or not all(re.fullmatch(r"[A-Z]{1,5}", p) for p in prefixes):
            results.append(f'❌ {label}: รหัสหมวด "{tokens[1]}" ไม่ถูกต้อง')
            continue
        members, unknown = [], []
        for a, b, m in own:
            if getattr(m, "type", "user") != "user":
                continue
            uid = getattr(m, "user_id", None)
            if not uid:  # not a friend of the bot: match the @name to someone seen in the group
                uid = find_group_member_by_name(group_id, text[a:b])
            if uid:
                members.append(uid)
            else:
                unknown.append(text[a:b].strip())
        if unknown or not members:
            who = " ".join(unknown) or "(ไม่มีชื่อที่ @)"
            results.append(f"❌ {label}: บอทยังไม่รู้จัก {who}\n"
                           "   → ให้คนนั้นพิมพ์ข้อความอะไรก็ได้ในกลุ่มนี้ 1 ครั้ง แล้วส่งคำสั่งตั้งบัดดี้ใหม่ "
                           "(ไม่ต้องแอดบอทเป็นเพื่อน)")
            continue
        members = list(dict.fromkeys(members))
        save_buddy_team(label, prefixes, members, group_id)
        names = ", ".join(get_profile_name(group_id, u) for u in members)
        results.append(f"✅ {label} (เอกสาร {', '.join(sorted(prefixes))}): {names}")
    if not results:
        push_to(sender_id, "ตั้งบัดดี้ไม่สำเร็จ ข้อมูลไม่ครบครับ\n" + usage)
        return
    msg = "ผลการตั้งบัดดี้\n" + "\n".join(results)
    if any(r.startswith("❌") for r in results):
        msg += "\n\n" + usage
    msg += '\n\nกดยกเลิกข้อความคำสั่งในกลุ่มได้เลย · ดูทั้งหมดพิมพ์ "บัดดี้" ในแชทนี้'
    push_to(sender_id, msg)

def build_buddy_list_text():
    teams = load_buddy_teams()
    if not teams:
        return ('ยังไม่ได้ตั้งบัดดี้ครับ ตั้งในกลุ่มด้วย:\nตั้งบัดดี้ AR,ST EE-A @ชื่อ1 @ชื่อ2\n'
                '(เลือกชื่อจากรายชื่อที่ LINE เด้งขึ้นมา)')
    lines = ["👥 บัดดี้ผู้รับผิดชอบตอบเอกสาร"]
    for t in teams:
        names = ", ".join(get_profile_name(t["group_id"], u) for u in t["members"])
        lines.append(f"• {t['label']} ({', '.join(sorted(t['prefixes']))}): {names}")
    lines.append('\nลบ: "ลบบัดดี้ EE-A" · เปลี่ยนคน: ส่ง "ตั้งบัดดี้" ชื่อเดิมทับในกลุ่ม')
    return "\n".join(lines)

def group_member_ok(group_id, user_id):
    try:
        resp = requests.get(f"https://api.line.me/v2/bot/group/{group_id}/member/{user_id}",
                            headers={"Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}"}, timeout=10)
        return resp.status_code != 404
    except Exception:
        return True  # can't tell - try the mention anyway

# --- Teams that answer the owner privately (e.g. AC) ---
# The owner (or a tracked admin) records it in the 1:1 chat:
#     ตอบแล้ว AR-325 RIH-I02330      (team defaults to MANUAL_ANSWER_DEFAULT_TEAM)
#     AC ตอบแล้ว AR-325              (or name the team)
#     AC ตอบครบแล้ว                  (every doc still waiting on that team)
#     AC ยกเลิกตอบแล้ว AR-325        (undo)
# Saved as a reply from "team:<label>", which counts for that team.
MANUAL_ANSWER_DEFAULT_TEAM = os.environ.get("MANUAL_ANSWER_DEFAULT_TEAM", "AC").strip().upper()
MANUAL_ANSWER_RE = re.compile(r"^(?:(\S+)\s+)?(ยกเลิกตอบแล้ว|ตอบครบแล้ว|ตอบแล้ว)(?:\s+(.*))?$", re.DOTALL)

def find_doc_code(ref):
    """Doc short code ("AR-325") from a code or a document number
    ("RIH-I02330", or just its tail like "I02330" / "02330")."""
    code = normalize_doc_code(ref)
    if code and get_inspection_doc(code):
        return code
    ref_u = (ref or "").strip().upper()
    if len(ref_u) < 4:
        return None
    conn = db()
    try:
        rows = conn.execute("SELECT code, doc_number FROM inspection_docs").fetchall()
    finally:
        conn.close()
    exact = [c for c, n in rows if (n or "").strip().upper() == ref_u]
    if exact:
        return exact[0]
    tail = [c for c, n in rows if (n or "").strip().upper().endswith(ref_u)]
    return tail[0] if len(tail) == 1 else None

def handle_manual_answer(user_id, text):
    """Returns a reply string, or None if the text isn't this command."""
    m = MANUAL_ANSWER_RE.match(text)
    if not m:
        return None
    allowed = ({OWNER_USER_ID} if OWNER_USER_ID else set()) | set(TAGGED_USER_IDS)
    if allowed and user_id not in allowed:
        return None
    label = (m.group(1) or MANUAL_ANSWER_DEFAULT_TEAM).upper()
    action, rest = m.group(2), (m.group(3) or "").strip()
    teams = {t["label"]: t for t in load_buddy_teams()}
    team = teams.get(label)
    if not team:
        return (f"ไม่พบบัดดี้ชื่อ {label} ครับ ตั้งในกลุ่มก่อนด้วย "
                f"\"ตั้งบัดดี้ AR,ST {label} @ชื่อ1 @ชื่อ2\"")
    marker = team_marker(label)

    if action == "ตอบครบแล้ว":
        targets = [d["code"] for d in get_docs_status(REMIND_WINDOW_DAYS)
                   if d["pending"] and d["status"]
                   and any(label in s["teams"] and not s["done"] for s in d["status"])]
        if not targets:
            return f"ไม่มีเอกสารที่ยังรอทีม {label} ตอบครับ"
        for code in targets:
            add_doc_reply(code, marker, f"ทีม {label} (แจ้งผ่านแอดมิน)", "ตอบแล้ว (ตอบส่วนตัวกับแอดมิน)", "owner")
        return f"✅ บันทึกว่าทีม {label} ตอบครบแล้ว {len(targets)} ฉบับ: {', '.join(targets)}"

    refs = [r for r in re.split(r"[\s,]+", rest) if r]
    if not refs:
        return (f"ใส่รหัสหรือเลขที่เอกสารด้วยครับ เช่น \"{action} AR-325 RIH-I02330\" "
                f"หรือ \"{label} ตอบครบแล้ว\"")
    ok, skipped, unknown = [], [], []
    conn = None
    for ref in refs:
        code = find_doc_code(ref)
        if not code:
            unknown.append(ref)
            continue
        if doc_prefix(code) not in team["prefixes"]:
            skipped.append(code)
            continue
        existing = [r for r in get_doc_replies(code) if r["user_id"] == marker]
        if action == "ยกเลิกตอบแล้ว":
            conn = db()
            try:
                conn.execute("DELETE FROM doc_replies WHERE code = ? AND user_id = ?", (code, marker))
            finally:
                conn.close()
            ok.append(code)
        elif existing:
            ok.append(code)  # already recorded - nothing to add
        else:
            add_doc_reply(code, marker, f"ทีม {label} (แจ้งผ่านแอดมิน)", "ตอบแล้ว (ตอบส่วนตัวกับแอดมิน)", "owner")
            ok.append(code)
    lines = []
    if ok:
        lines.append(("↩️ ยกเลิกสถานะตอบแล้วของทีม " if action == "ยกเลิกตอบแล้ว" else "✅ บันทึกว่าทีม ")
                     + f"{label}{'' if action == 'ยกเลิกตอบแล้ว' else ' ตอบแล้ว'}: {', '.join(ok)}")
    if skipped:
        lines.append(f"ข้าม (ทีม {label} ไม่ได้รับผิดชอบหมวดนี้): {', '.join(skipped)}")
    if unknown:
        lines.append(f"ไม่พบเอกสาร: {', '.join(unknown)}")
    return "\n".join(lines)

# --- Daily follow-up in the group: documents still waiting for a reply ---
# Every working day the bot posts, into the group each document was posted to,
# the Request PB documents from the last REMIND_WINDOW_DAYS days still waiting
# for an answer (oldest first, days waited, which buddy teams are still due),
# then tags exactly the people whose team still owes an answer (@All for docs
# with no team set). Groups with nothing outstanding get nothing.
#   1:1 "ติดตาม"     -> preview what would be posted (reply, free)
#   1:1 "ส่งติดตาม"  -> post it to the group(s) right now
REMIND_HOUR = int(os.environ.get("REMIND_HOUR", "9"))
REMIND_MINUTE = int(os.environ.get("REMIND_MINUTE", "0"))
REMIND_DAYS_OF_WEEK = os.environ.get("REMIND_DAYS_OF_WEEK", "mon-sat")  # APScheduler day_of_week syntax
REMIND_WINDOW_DAYS = int(os.environ.get("REMIND_WINDOW_DAYS", str(REPLIES_OVERVIEW_DAYS)))
REMIND_LATE_DAYS = int(os.environ.get("REMIND_LATE_DAYS", "3"))   # waited this long or more -> red
REMIND_MAX_DOCS = 25
REMIND_MAX_MENTIONS = 20
REMIND_MENTION_TEXT = os.environ.get(
    "REMIND_MENTION_TEXT",
    "รบกวนตรวจสอบและตอบกลับเอกสาร Request PB ที่ยังค้างด้านบนด้วยนะคะ "
    'กด "💬 ตอบกลับ" ใต้เอกสาร หรือพิมพ์รหัสตามด้วยข้อความในกลุ่ม เช่น "AR-337 ตรวจแล้ว" ค่ะ',
)
REMIND_HEADER_COLOR = "#E65100"
REMIND_WAIT_COLOR = "#EF6C00"

def get_unanswered_docs_by_group():
    """{group_id: [doc dict, ...]} for docs still waiting, oldest first."""
    out = {}
    for d in get_docs_status(REMIND_WINDOW_DAYS):
        if d["pending"]:
            out.setdefault(d["group_id"], []).append(d)
    return out

def _days_waiting(posted_at):
    try:
        posted = datetime.datetime.fromisoformat(posted_at).date()
    except (ValueError, TypeError):
        return 0
    return max((bkk_today() - posted).days, 0)

def followup_tag_targets(docs):
    """(user_ids to mention, needs @All) for a list of pending docs."""
    ids, need_all = [], False
    for d in docs:
        if d["status"] is None:
            need_all = True
            continue
        if any(s.get("missing") for s in d["status"]):
            need_all = True  # a system with no buddy team - nobody specific to tag
        for s in d["status"]:
            for u in s["tag_ids"]:
                if u not in ids:
                    ids.append(u)
    if len(ids) > REMIND_MAX_MENTIONS:
        return ([], True)
    return (ids, need_all)

def build_followup_flex(docs):
    txt = flex_txt
    late = sum(1 for d in docs if _days_waiting(d["posted_at"]) >= REMIND_LATE_DAYS)
    body = [
        txt("🔔 ติดตามการตอบกลับเอกสาร Request PB", weight="bold", size="md", color=REMIND_HEADER_COLOR),
        {"type": "text", "wrap": True, "size": "xs", "color": "#666666", "margin": "xs", "contents": [
            {"type": "span", "text": "ยังรอคำตอบ "},
            {"type": "span", "text": f"{len(docs)} ฉบับ", "weight": "bold", "color": OVERDUE_COLOR},
            {"type": "span", "text": (f" · ค้างเกิน {REMIND_LATE_DAYS} วัน {late} ฉบับ" if late else "")
                                     + f" · ข้อมูล ณ {bkk_now().strftime('%d/%m %H:%M')} น."},  # never empty
        ]},
        {"type": "separator", "margin": "md"},
    ]
    for d in docs[:REMIND_MAX_DOCS]:
        code = d["code"]
        waited = _days_waiting(d["posted_at"])
        color = OVERDUE_COLOR if waited >= REMIND_LATE_DAYS else REMIND_WAIT_COLOR
        wait_label = "โพสต์วันนี้" if waited == 0 else f"รอ {waited} วัน"
        body.append(two_tone_text(f"{'🔴' if waited >= REMIND_LATE_DAYS else '🟠'} {code}", color,
                                  f"  ({wait_label})", size="sm", margin="md"))
        body.append(txt(d["subject"] or "-", size="xs", color="#444444", margin="none"))
        if d["status"]:
            body.append(txt(team_status_line(d["status"]), size="xs", color="#333333", margin="xs"))
        meta = [txt(f"ส่งเข้าระบบ {d['date']}" if d["date"] else "-", size="xxs", color="#999999", flex=1)]
        link = doc_reply_liff_url(code)
        if link:
            meta.append({"type": "text", "text": "💬 ตอบกลับ", "size": "xs", "weight": "bold",
                         "color": REPLY_LINK_COLOR, "align": "end", "flex": 0,
                         "action": {"type": "uri", "label": "ตอบกลับ", "uri": link}})
        body.append({"type": "box", "layout": "horizontal", "margin": "none", "contents": meta})
    if len(docs) > REMIND_MAX_DOCS:
        body.append(txt(f"… และอีก {len(docs) - REMIND_MAX_DOCS} ฉบับ", size="xs", color="#999999", margin="md"))
    contents = {"type": "bubble", "size": "giga",
                "body": {"type": "box", "layout": "vertical", "spacing": "xs", "contents": body}}
    return (f"🔔 ติดตามการตอบกลับเอกสาร Request PB: ยังรอคำตอบ {len(docs)} ฉบับ", contents)

def build_followup_mention_message(user_ids, need_all):
    keys, subs = [], {}
    if need_all or not user_ids:
        keys.append("{all}")
        subs["all"] = {"type": "mention", "mentionee": {"type": "all"}}
    for i, uid in enumerate(user_ids, 1):
        keys.append(f"{{u{i}}}")
        subs[f"u{i}"] = {"type": "mention", "mentionee": {"type": "user", "userId": uid}}
    return {"type": "textV2", "text": " ".join(keys) + " " + REMIND_MENTION_TEXT, "substitution": subs}

def push_followup_to_group(group_id, alt_text, contents, user_ids, need_all):
    messages = [{"type": "flex", "altText": alt_text, "contents": contents},
                build_followup_mention_message(user_ids, need_all)]
    resp = requests.post(
        LINE_PUSH_URL,
        headers={"Authorization": f"Bearer {CHANNEL_ACCESS_TOKEN}", "Content-Type": "application/json"},
        json={"to": group_id, "messages": messages},
        timeout=15,
    )
    app.logger.warning("follow-up push to %s -> HTTP %s %s", group_id, resp.status_code, resp.text[:300])
    resp.raise_for_status()

def send_reply_followup(force=False):
    """Posts the follow-up to every group with docs still waiting. Returns a
    short Thai status line. The scheduled run (force=False) sends at most once
    per day, so a restart around the scheduled minute can't double-post."""
    today = bkk_today().isoformat()
    if not force and get_setting("reply_followup_last_sent", "") == today:
        return "วันนี้ส่งติดตามไปแล้วครับ"
    by_group = get_unanswered_docs_by_group()
    if not by_group:
        set_setting("reply_followup_last_sent", today)
        return "ทุกเอกสารตอบครบแล้ว ไม่ต้องส่งติดตามครับ 🎉"
    sent, failed, missing = [], [], []
    for group_id, docs in by_group.items():
        target = group_id or inspection_target_group_id()
        if not target:
            failed.append(f"{len(docs)} ฉบับ (ไม่รู้กลุ่มปลายทาง)")
            continue
        user_ids, need_all = followup_tag_targets(docs)
        present = []
        for uid in user_ids:
            if group_member_ok(target, uid):
                present.append(uid)
            else:
                missing.append(uid)
        if len(present) < len(user_ids):
            need_all = True  # someone who should answer has left - make sure the group still hears it
        alt_text, contents = build_followup_flex(docs)
        try:
            push_followup_to_group(target, alt_text, contents, present, need_all)
            sent.append(f'"{get_group_name(target)}" {len(docs)} ฉบับ')
        except Exception as e:
            app.logger.exception("follow-up push failed")
            failed.append(f'"{get_group_name(target)}" ({e})')
    if sent:
        set_setting("reply_followup_last_sent", today)
    lines = []
    if sent:
        lines.append("ส่งติดตามคำตอบเข้ากลุ่มแล้ว: " + ", ".join(sent))
    if failed:
        lines.append("ส่งไม่สำเร็จ: " + ", ".join(failed))
    if missing:
        teams = load_buddy_teams()
        labels = sorted({t["label"] for t in teams for u in missing if u in t["members"]})
        lines.append(f"⚠️ มีผู้รับผิดชอบ {len(missing)} คนไม่อยู่ในกลุ่มแล้ว (บัดดี้ {', '.join(labels)}) "
                     "จึงแท็ก @All แทน — ตั้งคนใหม่ด้วย \"ตั้งบัดดี้\" ในกลุ่ม")
    return "\n".join(lines)

def scheduled_reply_followup():
    try:
        result = send_reply_followup()
        if "⚠️" in result or "ไม่สำเร็จ" in result:
            push_to_owner(result)
    except Exception:
        app.logger.exception("scheduled reply follow-up failed")

def build_followup_preview():
    by_group = get_unanswered_docs_by_group()
    if not by_group:
        return ("ทุกเอกสารในช่วง %d วันล่าสุดตอบครบแล้วครับ 🎉" % REMIND_WINDOW_DAYS, None, "")
    docs = [d for group_docs in by_group.values() for d in group_docs]
    docs.sort(key=lambda d: (d["posted_at"] or "", d["code"]))
    alt_text, contents = build_followup_flex(docs)
    user_ids, need_all = followup_tag_targets(docs)
    group_for_names = next(iter(by_group)) or inspection_target_group_id()
    who = [get_profile_name(group_for_names, u) for u in user_ids]
    if need_all:
        who.insert(0, "@All")
    return (alt_text, contents, "จะแท็ก: " + (", ".join(who) if who else "@All"))

# --- LIFF page + JSON API ---
LINE_VERIFY_ID_TOKEN_URL = "https://api.line.me/oauth2/v2.1/verify"

def verify_liff_id_token(id_token):
    """Returns (user_id, display_name) for a valid LINE ID token, else (None, None)."""
    if not id_token or not LINE_LOGIN_CHANNEL_ID:
        return (None, None)
    try:
        resp = requests.post(LINE_VERIFY_ID_TOKEN_URL,
                             data={"id_token": id_token, "client_id": LINE_LOGIN_CHANNEL_ID},
                             timeout=10)
        if resp.status_code != 200:
            app.logger.warning("ID token verify failed: %s %s", resp.status_code, resp.text[:200])
            return (None, None)
        data = resp.json()
        return (data.get("sub"), data.get("name") or "ไม่ทราบชื่อ")
    except Exception:
        app.logger.exception("ID token verify error")
        return (None, None)

def _liff_auth():
    auth = request.headers.get("Authorization", "")
    token = auth[len("Bearer "):] if auth.startswith("Bearer ") else ""
    return verify_liff_id_token(token)

@app.route("/api/docs/<code>/replies", methods=["GET", "POST"])
def api_doc_replies(code):
    user_id, user_name = _liff_auth()
    if not user_id:
        return {"error": "unauthorized"}, 401
    code = normalize_doc_code(code)
    doc = get_inspection_doc(code) if code else None
    if not doc:
        return {"error": "not_found"}, 404
    if request.method == "POST":
        payload = request.get_json(silent=True) or {}
        text = (payload.get("text") or "").strip()
        if not text:
            return {"error": "empty"}, 400
        add_doc_reply(code, user_id, user_name, text, "liff")
        after_doc_reply(code, user_id, user_name, text)
    replies = [{"user_name": r["user_name"], "text": r["text"],
                "time": _fmt_dt(r["created_at"]), "mine": r["user_id"] == user_id}
               for r in get_doc_replies(code)]
    return {"doc": doc, "replies": replies}

@app.route("/api/docs", methods=["GET"])
def api_docs():
    user_id, _ = _liff_auth()
    if not user_id:
        return {"error": "unauthorized"}, 401
    since = (bkk_now() - datetime.timedelta(days=REPLIES_OVERVIEW_DAYS)).isoformat()
    conn = db()
    try:
        rows = conn.execute(
            "SELECT d.code, d.subject, d.date, (SELECT COUNT(*) FROM doc_replies r WHERE r.code = d.code) "
            "FROM inspection_docs d WHERE d.posted_at >= ? ORDER BY d.posted_at DESC, d.code",
            (since,),
        ).fetchall()
    finally:
        conn.close()
    return {"docs": [{"code": c, "subject": s, "date": dt, "replies": n} for c, s, dt, n in rows]}

LIFF_REPLY_HTML = r"""<!doctype html>
<html lang="th"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>ตอบกลับเอกสาร</title>
<style>
  :root{--blue:#0D47A1;--pill:#E3F2FD;--grey:#78909C;--line:#E0E0E0;--green:#0B6E4F;--bg:#F5F7FA}
  *{box-sizing:border-box} body{margin:0;font-family:-apple-system,"Sukhumvit Set","Noto Sans Thai",sans-serif;background:var(--bg);color:#111}
  header{background:#fff;padding:14px 16px;border-bottom:1px solid var(--line);position:sticky;top:0;z-index:2}
  .code{display:inline-block;background:var(--pill);color:var(--blue);font-weight:700;padding:3px 10px;border-radius:8px;font-size:14px}
  .subject{margin:8px 0 2px;font-weight:600;font-size:15px;line-height:1.4}
  .meta{color:var(--grey);font-size:12px}
  main{padding:12px 16px 110px}
  .reply{background:#fff;border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin-bottom:10px}
  .reply.mine{border-color:#90CAF9}
  .who{font-size:12px;color:var(--green);font-weight:700}.who span{color:var(--grey);font-weight:400;margin-left:6px}
  .text{margin-top:4px;white-space:pre-wrap;word-break:break-word;font-size:14px;line-height:1.5}
  .empty,.msg{color:var(--grey);text-align:center;padding:28px 0;font-size:14px}
  footer{position:fixed;left:0;right:0;bottom:0;background:#fff;border-top:1px solid var(--line);padding:10px 12px;display:flex;gap:8px}
  textarea{flex:1;border:1px solid #CFD8DC;border-radius:10px;padding:10px;font:inherit;font-size:14px;resize:none;height:44px;max-height:120px}
  button{background:var(--blue);color:#fff;border:0;border-radius:10px;padding:0 18px;font-weight:700;font-size:14px}
  button:disabled{opacity:.5}
  .doc{display:block;background:#fff;border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin-bottom:10px;color:inherit;text-decoration:none}
  .count{float:right;font-size:12px;color:var(--grey)} .count.zero{color:#C62828;font-weight:700}
</style></head>
<body>
<header id="head"><div class="meta">กำลังโหลด…</div></header>
<main id="list"></main>
<footer id="composer" style="display:none">
  <textarea id="input" placeholder="พิมพ์คำตอบ…" maxlength="1000"></textarea>
  <button id="send">ส่ง</button>
</footer>
<script src="https://static.line-scdn.net/liff/edge/2/sdk.js"></script>
<script>
const LIFF_ID = "__LIFF_ID__";
const $ = id => document.getElementById(id);
function el(tag, cls, text){const e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
let token = null, code = null;

async function api(path, opts={}){
  const r = await fetch(path, Object.assign({headers:{"Authorization":"Bearer "+token,"Content-Type":"application/json"}}, opts));
  if(!r.ok) throw new Error(r.status);
  return r.json();
}

function renderThread(data){
  const d = data.doc, head = $("head"); head.innerHTML = "";
  head.append(el("span","code",d.code), el("div","subject",d.subject), el("div","meta","ส่งเข้าระบบ "+d.date+(d.extra_note?" · ดำเนินงาน "+d.extra_note:"")));
  const list = $("list"); list.innerHTML = "";
  if(!data.replies.length){ list.append(el("div","empty","ยังไม่มีคำตอบ เป็นคนแรกที่ตอบได้เลย")); }
  data.replies.forEach(r=>{
    const box = el("div","reply"+(r.mine?" mine":""));
    const who = el("div","who",r.user_name); who.append(el("span",null,r.time+" น."));
    box.append(who, el("div","text",r.text)); list.append(box);
  });
  $("composer").style.display = "flex";
  window.scrollTo(0, document.body.scrollHeight);
}

function renderDocList(data){
  const head = $("head"); head.innerHTML = "";
  head.append(el("div","subject","💬 เลือกเอกสารที่จะตอบกลับ"));
  const list = $("list"); list.innerHTML = "";
  if(!data.docs.length){ list.append(el("div","empty","ยังไม่มีเอกสารที่โพสต์ล่าสุด")); return; }
  data.docs.forEach(d=>{
    const a = el("a","doc"); a.href = "?doc="+encodeURIComponent(d.code);
    a.append(el("span","count"+(d.replies?"":" zero"), d.replies? d.replies+" คำตอบ":"ยังไม่มีคำตอบ"), el("span","code",d.code), el("div","subject",d.subject), el("div","meta","ส่งเข้าระบบ "+d.date));
    list.append(a);
  });
}

async function send(){
  const input = $("input"), text = input.value.trim(); if(!text) return;
  const btn = $("send"); btn.disabled = true;
  try{ renderThread(await api("/api/docs/"+encodeURIComponent(code)+"/replies",{method:"POST",body:JSON.stringify({text})})); input.value=""; }
  catch(e){ alert("ส่งไม่สำเร็จ ลองใหม่อีกครั้ง"); }
  btn.disabled = false;
}

(async function(){
  try{
    await liff.init({liffId: LIFF_ID});
    if(!liff.isLoggedIn()){ liff.login({redirectUri: location.href}); return; }
    token = liff.getIDToken();
    code = new URLSearchParams(location.search).get("doc");
    if(code){ renderThread(await api("/api/docs/"+encodeURIComponent(code)+"/replies")); }
    else { renderDocList(await api("/api/docs")); }
    $("send").onclick = send;
  }catch(e){
    $("head").innerHTML = ""; $("head").append(el("div","msg", e.message==="404" ? "ไม่พบเอกสารนี้ในรายการที่โพสต์" : "โหลดไม่สำเร็จ ลองปิดแล้วเปิดใหม่ ("+e.message+")"));
  }
})();
</script>
</body></html>"""

@app.route("/liff/reply", methods=["GET"])
def liff_reply_page():
    if not LIFF_ID:
        return Response("LIFF_ID is not configured on the server.", status=503, mimetype="text/plain")
    return Response(LIFF_REPLY_HTML.replace("__LIFF_ID__", LIFF_ID), mimetype="text/html")

# ---------- Summary building ----------
def build_summary_text(date=None):
    date = date or bkk_today()
    start = datetime.datetime.combine(date, datetime.time.min).isoformat()
    end = datetime.datetime.combine(date, datetime.time.max).isoformat()

    conn = db()
    rows = conn.execute(
        """SELECT group_name, sender_name, event_type, file_name, text_content, sent_at, tagged, downloaded
           FROM file_events WHERE sent_at BETWEEN ? AND ? AND tagged = 1
           ORDER BY group_name, sender_name, sent_at""",
        (start, end)
    ).fetchall()
    conn.close()

    if not rows:
        return f"วันที่ {date.strftime('%d/%m/%Y')} ยังไม่มีข้อความที่มีคนแท็คคุณครับ"

    by_group = {}
    for group_name, sender_name, event_type, file_name, text_content, sent_at, tagged, downloaded in rows:
        by_group.setdefault(group_name, {}).setdefault(sender_name, []).append(
            (event_type, file_name, text_content, sent_at, downloaded)
        )

    pending_total = 0
    lines = [f"📋 สรุปรายการที่แท็คคุณ วันที่ {date.strftime('%d/%m/%Y')}"]
    for group_name, by_sender in by_group.items():
        lines.append(f"\n━━ 📁 {group_name} ━━")
        for sender_name, items in by_sender.items():
            pending_count = sum(1 for e, _, _, _, d in items if e == "file" and not d)
            pending_total += pending_count
            label = f" (ยังไม่โหลด {pending_count})" if any(e == "file" for e, *_ in items) else ""
            lines.append(f"\n👤 {sender_name}{label}")
            for event_type, file_name, text_content, sent_at, downloaded in items:
                t = datetime.datetime.fromisoformat(sent_at).strftime("%H:%M")
                if event_type == "file":
                    mark = "✅" if downloaded else "🔲"
                    lines.append(f"  {mark} {file_name}  ({t} น.)")
                    if text_content:
                        preview = text_content.replace("\n", " ")
                        if len(preview) > 80:
                            preview = preview[:80] + "..."
                        lines.append(f"      💬 {preview}")
                else:
                    preview = (text_content or "").replace("\n", " ")
                    if len(preview) > 60:
                        preview = preview[:60] + "..."
                    lines.append(f"  💬 {preview}  ({t} น.)")

    if pending_total == 0:
        lines.append("\nไฟล์งานโหลดครบแล้วครับ 🎉")
    else:
        lines.append("\n" + CLOSE_HELP_TEXT)
    return "\n".join(lines)

# ---------- Flex building helpers (shared by the daily summary and the overdue digest) ----------
MENTION_COLOR = "#1E88E5"   # tagged LINE name (was the message color)
MESSAGE_COLOR = "#111111"   # the rest of the message text (was the file name color)
FILE_NAME_COLOR = "#E65100"  # file name (was the tagged-name color)
MAX_PREVIEW = 500

def flex_txt(content, **kwargs):
    d = {"type": "text", "text": content, "wrap": True}
    d.update(kwargs)
    return d

GROUP_NAME_BG_COLOR = "#E8F5E9"    # light green highlight behind the group name
GROUP_NAME_TEXT_COLOR = "#0B6E4F"  # matches the group header's existing green

CLOSE_BUTTON_COLOR = "#0B6E4F"

def close_tap_text(label, command_text):
    """A small tappable text that sends command_text as if the admin typed it
    (works in the bot's 1:1 chat, where the summaries are delivered)."""
    return {
        "type": "text", "text": label, "size": "xs", "weight": "bold",
        "color": CLOSE_BUTTON_COLOR, "align": "end", "gravity": "center", "flex": 0,
        "decoration": "underline",
        "action": {"type": "message", "label": label[:20], "text": command_text},
    }

def group_name_badge(group_name, margin="lg", with_close_button=False):
    """A highlighted pill behind the group name, so it stands out at a
    glance when scanning a summary with several groups in it. With
    with_close_button, adds a "ปิดทั้งกลุ่ม" tap target that sends "✅ <code>"."""
    contents = [flex_txt(f"📁 {group_name}", weight="bold", size="md", color=GROUP_NAME_TEXT_COLOR, flex=1)]
    if with_close_button:
        contents.append(close_tap_text("✅ ปิดทั้งกลุ่ม", f"✅ {group_close_token(group_name)}"))
    return {
        "type": "box", "layout": "horizontal" if with_close_button else "vertical", "margin": margin,
        "backgroundColor": GROUP_NAME_BG_COLOR, "cornerRadius": "md", "paddingAll": "sm",
        "spacing": "sm", "contents": contents
    }

def sender_line(sender_label, group_name, sender_name, with_close_button):
    """The "👤 name (ค้าง N)" line, plus a "✅ ปิด" tap target for just that
    person inside that group when they still have pending items."""
    name_txt = flex_txt(sender_label, weight="bold", size="sm", color="#222222", flex=1)
    if not with_close_button:
        name_txt["margin"] = "sm"
        return name_txt
    return {
        "type": "box", "layout": "horizontal", "margin": "sm", "spacing": "sm",
        "contents": [name_txt, close_tap_text("✅ ปิด", f"✅ {group_close_token(group_name)} {sender_name}")]
    }

CLOSE_HELP_TEXT = ('กด "✅ ปิดทั้งกลุ่ม" / "✅ ปิด" ด้านบน หรือพิมพ์ "✅" ปิดทั้งหมด, '
                   '"✅ REQ" / "✅ BIM" / "✅ PM" ปิดทั้งกลุ่ม, "✅ <ชื่อคน>" ปิดเฉพาะคน, '
                   '"✅ PM <ชื่อคน>" ปิดเฉพาะคนในกลุ่มนั้น — ปิดผิดพิมพ์ "ย้อนกลับ"')

TIME_COLOR = "#888888"  # timestamps, kept visually distinct from the text next to them

def two_tone_text(main_text, main_color, suffix_text, suffix_color=TIME_COLOR, **kwargs):
    """A text component with a bold colored main part and a differently-colored
    suffix (used for "<file name> (<time>)" so the time doesn't blend into the name)."""
    spans = [{"type": "span", "text": main_text, "color": main_color, "weight": "bold"}]
    if suffix_text:
        spans.append({"type": "span", "text": suffix_text, "color": suffix_color})
    d = {"type": "text", "wrap": True, "contents": spans}
    d.update(kwargs)
    return d

def message_text_component(prefix, text_content, mention_spans_json, suffix="", **kwargs):
    """Builds a text component where the @mention substring(s) are colored
    differently (orange, bold) from the rest of the message (blue)."""
    raw = (text_content or "").replace("\n", " ")
    spans = []
    if mention_spans_json:
        try:
            spans = json.loads(mention_spans_json)
        except (ValueError, TypeError):
            spans = []

    valid = []
    for s in spans:
        if isinstance(s, (list, tuple)) and len(s) == 2:
            idx, length = s
            if isinstance(idx, int) and isinstance(length, int) and 0 <= idx < len(raw) and length > 0:
                valid.append((idx, min(idx + length, len(raw))))
    valid.sort()
    merged = []
    for idx, e in valid:
        if merged and idx <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((idx, e))

    truncated = False
    text = raw
    if len(text) > MAX_PREVIEW:
        text = text[:MAX_PREVIEW]
        truncated = True
        merged = [(i, min(e, MAX_PREVIEW)) for i, e in merged if i < MAX_PREVIEW]

    spans_out = []
    if prefix:
        spans_out.append({"type": "span", "text": prefix, "color": MESSAGE_COLOR})
    if not merged:
        spans_out.append({"type": "span", "text": text or " ", "color": MESSAGE_COLOR})
    else:
        pos = 0
        for idx, e in merged:
            if idx > pos:
                spans_out.append({"type": "span", "text": text[pos:idx], "color": MESSAGE_COLOR})
            spans_out.append({"type": "span", "text": text[idx:e], "color": MENTION_COLOR, "weight": "bold"})
            pos = e
        if pos < len(text):
            spans_out.append({"type": "span", "text": text[pos:], "color": MESSAGE_COLOR})
    if truncated:
        spans_out.append({"type": "span", "text": "...", "color": MESSAGE_COLOR})
    if suffix:
        spans_out.append({"type": "span", "text": suffix, "color": "#888888"})

    d = {"type": "text", "wrap": True, "contents": spans_out}
    d.update(kwargs)
    return d

def build_summary_flex(date=None, user_id=None):
    """Returns (alt_text, flex_contents_dict). flex_contents_dict is None when
    there's nothing to show, in which case alt_text is the plain message to send instead.
    When user_id is given, only shows events where THAT specific person was tagged
    (each admin sees only their own); the downloaded ✅/🔲 status is shared data either way."""
    date = date or bkk_today()
    start = datetime.datetime.combine(date, datetime.time.min).isoformat()
    end = datetime.datetime.combine(date, datetime.time.max).isoformat()
    date_str = date.strftime('%d/%m/%Y')

    conn = db()
    rows = conn.execute(
        """SELECT group_name, sender_name, event_type, file_name, text_content, sent_at, tagged, downloaded, mention_spans, tagged_user_ids, downloaded_by
           FROM file_events
           WHERE tagged = 1 AND (
               (sent_at BETWEEN ? AND ? AND NOT downloaded = 1)
               OR (downloaded = 0 AND sent_at < ?)
           )
           ORDER BY group_name, sender_name, sent_at""",
        (start, end, start)
    ).fetchall()
    conn.close()

    if user_id:
        def includes_user(tagged_user_ids_json):
            if not tagged_user_ids_json:
                return True  # unknown (older row) - show rather than silently hide
            try:
                return user_id in json.loads(tagged_user_ids_json)
            except (ValueError, TypeError):
                return True
        rows = [r for r in rows if includes_user(r[9])]

    if not rows:
        return (f"วันที่ {date_str} ยังไม่มีข้อความที่มีคนแท็คคุณครับ", None)

    by_group = {}
    for group_name, sender_name, event_type, file_name, text_content, sent_at, tagged, downloaded, mention_spans, tagged_user_ids, downloaded_by in rows:
        by_group.setdefault(group_name, {}).setdefault(sender_name, []).append(
            (event_type, file_name, text_content, sent_at, downloaded, mention_spans, downloaded_by)
        )

    txt = flex_txt

    body = [txt(f"📋 สรุปรายการที่แท็คคุณ", weight="bold", size="lg", color="#111111")]
    body.append(txt(date_str, size="sm", color="#888888", margin="xs"))
    body.append({"type": "separator", "margin": "md"})

    pending_total = 0
    item_count = 0
    MAX_ITEMS = 60  # safety cap so the flex message doesn't blow past LINE's size limit
    truncated = False

    for group_name, by_sender in by_group.items():
        group_has_pending = any(not it[4] for items in by_sender.values() for it in items)
        body.append(group_name_badge(group_name, with_close_button=group_has_pending))
        for sender_name, items in by_sender.items():
            if item_count >= MAX_ITEMS:
                truncated = True
                break
            pending_count = sum(1 for e, _, _, _, d, *_ in items if not d)
            pending_total += pending_count
            label = f"  (ค้าง {pending_count})" if pending_count else ""
            body.append(sender_line(f"👤 {sender_name}{label}", group_name, sender_name, pending_count > 0))

            for event_type, file_name, text_content, sent_at, downloaded, mention_spans, downloaded_by in items:
                if item_count >= MAX_ITEMS:
                    truncated = True
                    break
                item_count += 1
                sent_dt = datetime.datetime.fromisoformat(sent_at)
                is_overdue = not downloaded and sent_dt.date() < date
                t = sent_dt.strftime("%d/%m %H:%M") if is_overdue else sent_dt.strftime("%H:%M")
                mark = "✅" if downloaded else ("🔴" if is_overdue else "🔲")
                if event_type == "file":
                    name_color = OVERDUE_COLOR if is_overdue else FILE_NAME_COLOR
                    body.append(two_tone_text(f"{mark} {file_name}", name_color, f"  ({t} น.)",
                                               size="sm", margin="xs"))
                    if downloaded and downloaded_by:
                        body.append(txt(f"   ปิดโดย {downloaded_by}", size="xs", color="#0B6E4F", margin="none"))
                    if text_content:
                        body.append(message_text_component(
                            "💬 ", text_content, mention_spans, size="xs", margin="xs"
                        ))
                else:
                    body.append(message_text_component(
                        f"{mark} 💬 ", text_content, mention_spans, suffix=f"  ({t} น.)", size="sm", margin="xs"
                    ))
                    if downloaded and downloaded_by:
                        body.append(txt(f"   ปิดโดย {downloaded_by}", size="xs", color="#0B6E4F", margin="none"))

    if truncated:
        body.append(txt("… มีรายการเพิ่มเติมที่ไม่ได้แสดงครบ", size="xs", color="#999999", margin="md"))

    body.append({"type": "separator", "margin": "lg"})
    if pending_total == 0:
        body.append(txt("ปิดสถานะครบทุกรายการแล้วครับ 🎉", size="sm", color="#0B6E4F", margin="md"))
    else:
        body.append(txt(CLOSE_HELP_TEXT, size="xs", color="#888888", margin="md"))

    contents = {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "xs",
            "contents": body
        }
    }
    alt_text = f"สรุปรายการที่แท็คคุณ วันที่ {date_str} ({item_count} รายการ)"
    return (alt_text, contents)

def _normalize_name_for_match(name):
    # strip whitespace and any surrounding bracket/quote characters people
    # commonly (mis)type around a name, e.g. "<StarbugBank>", "[Starbug]",
    # so "โหลดแล้ว <StarbugBank>" still matches the real display name
    # "StarbugBank" instead of silently matching nothing.
    name = name.strip()
    name = name.strip("<>[](){}「」【】《》\"'‘’“”")
    return name.strip().lower()

def _item_label(file_name, text_content):
    """Short display label for a pending item in mark_downloaded's reply:
    the file name for a file event, or a trimmed snippet of the tagged
    message text for a text/pm event (which has no file_name)."""
    if file_name:
        return file_name
    snippet = (text_content or "").replace("\n", " ").strip()
    if len(snippet) > 40:
        snippet = snippet[:40] + "…"
    return snippet or "ข้อความที่ถูกแท็ค"

# ---------- Closing by group ----------
# Short codes for closing a whole group at once ("✅ REQ", "✅ PM Phichet").
# Each code maps to a fragment of the group's name; a group matches a code when
# that fragment appears in its name after stripping spaces/punctuation/case.
# Override or extend with the GROUP_ALIASES env var as JSON, e.g.
#   {"REQ": "Req't Document", "BIM": "BIM&Shop Dwg", "PM": "Project Management"}
DEFAULT_GROUP_ALIASES = {
    "REQ": "Req't Document",
    "BIM": "BIM&Shop Dwg",
    "PM": "Project Management",
}

def _norm_key(s):
    """Lowercase and keep only letters/digits/combining marks (Thai vowels
    included), so "FTC-AIH_Req't Document" and "ftc aih reqt document" compare equal."""
    return "".join(ch for ch in (s or "").lower() if unicodedata.category(ch)[0] in "LNM")

def _load_group_aliases():
    aliases = dict(DEFAULT_GROUP_ALIASES)
    raw = os.environ.get("GROUP_ALIASES", "").strip()
    if raw:
        try:
            aliases.update({str(k): str(v) for k, v in json.loads(raw).items()})
        except (ValueError, TypeError, AttributeError):
            pass
    return {code.strip().upper(): _norm_key(frag) for code, frag in aliases.items() if _norm_key(frag)}

GROUP_ALIASES = _load_group_aliases()

def group_short_code(group_name):
    key = _norm_key(group_name)
    for code, frag in GROUP_ALIASES.items():
        if frag in key:
            return code
    return None

def group_close_token(group_name):
    """What a button/command uses to point at a group: its short code if it has one."""
    return group_short_code(group_name) or group_name

def _pending_rows(conn):
    return conn.execute(
        "SELECT id, group_name, sender_name, file_name, text_content FROM file_events "
        "WHERE tagged = 1 AND downloaded = 0 ORDER BY group_name, sender_name, sent_at"
    ).fetchall()

def _parse_close_target(arg, pending):
    """Splits the text after "✅"/"โหลดแล้ว" into (group_name, sender_name, error).
    Order: an exact pending sender name wins (so a person whose name happens to
    look like part of a group name still works), then the longest leading run of
    words that names a group (short code, full name or a distinctive part of it),
    with whatever follows treated as a sender name inside that group."""
    arg = arg.strip()
    if not arg:
        return (None, None, None)

    groups = sorted({r[1] for r in pending if r[1]})
    senders = {_normalize_name_for_match(r[2]) for r in pending if r[2]}
    if _normalize_name_for_match(arg) in senders:
        return (None, arg, None)

    words = arg.split()
    for k in range(len(words), 0, -1):
        cand = " ".join(words[:k])
        rest = " ".join(words[k:]).strip() or None
        code = _normalize_name_for_match(cand).upper()
        if code in GROUP_ALIASES:
            matches = [g for g in groups if GROUP_ALIASES[code] in _norm_key(g)]
            if not matches:
                return (None, None, f"กลุ่ม {code} ไม่มีรายการค้างให้ปิดครับ")
        else:
            key = _norm_key(cand)
            if len(key) < 4:
                continue
            matches = [g for g in groups if key == _norm_key(g) or key in _norm_key(g)]
            if not matches:
                continue
        if len(matches) > 1:
            listed = "\n".join(f"• {group_close_token(g)}" for g in matches)
            return (None, None, f'"{cand}" ตรงกับหลายกลุ่ม ยังไม่ได้ปิดอะไรครับ ระบุให้ชัดขึ้นอีกนิด:\n{listed}')
        return (matches[0], rest, None)

    return (None, arg, None)

def _remaining_summary(conn):
    rows = _pending_rows(conn)
    if not rows:
        return "ไม่มีรายการค้างแล้วครับ 🎉"
    counts = {}
    for _, g, *_ in rows:
        counts[g] = counts.get(g, 0) + 1
    parts = [f"{group_close_token(g)} {n}" for g, n in counts.items()]
    return f"ยังค้างอยู่ {len(rows)} รายการ: " + ", ".join(parts)

def _not_found_message(arg, pending):
    lines = [f'ไม่พบ "{arg}" ทั้งในชื่อกลุ่มและชื่อคนที่ยังมีรายการค้างครับ ยังไม่ได้ปิดอะไร']
    groups = sorted({r[1] for r in pending if r[1]})
    names = sorted({r[2] for r in pending if r[2]})
    if groups:
        lines.append("กลุ่มที่ยังค้าง: " + ", ".join(group_close_token(g) for g in groups))
    if names:
        lines.append("คนที่ยังค้าง: " + ", ".join(names))
    lines.append('ตัวอย่าง: "✅ REQ", "✅ Phichet", "✅ PM Phichet"')
    return "\n".join(lines)

def handle_close_command(arg, user_id):
    """Entry point for "โหลดแล้ว" / "✅" with an optional group and/or person."""
    arg = arg.replace("\ufe0f", "").strip()  # drop the emoji variation selector some keyboards add
    marked_by = get_profile_name(None, user_id) if user_id else None
    conn = db()
    try:
        pending = _pending_rows(conn)
        if not pending:
            return "ไม่มีรายการค้างให้ปิดครับ"

        group_name, sender_name, error = _parse_close_target(arg, pending)
        if error:
            return error

        selected = pending
        if group_name:
            selected = [r for r in selected if r[1] == group_name]
        if sender_name:
            want = _normalize_name_for_match(sender_name)
            selected = [r for r in selected if _normalize_name_for_match(r[2] or "") == want]

        if not selected:
            if group_name and sender_name:
                in_group = sorted({r[2] for r in pending if r[1] == group_name})
                return (f'ไม่พบรายการค้างของ "{sender_name}" ในกลุ่ม {group_close_token(group_name)} ครับ\n'
                        f'คนที่ยังค้างในกลุ่มนี้: {", ".join(in_group)}')
            return _not_found_message(arg, pending)

        g_token = group_close_token(group_name) if group_name else None
        if sender_name:
            sender_name = selected[0][2]  # the real display name, not however it was typed
        if group_name and sender_name:
            scope = f"ของ {sender_name} ในกลุ่ม {g_token}"
        elif group_name:
            scope = f"ทั้งกลุ่ม {g_token}"
        elif sender_name:
            scope = f"ของ {sender_name}"
        else:
            scope = "ทั้งหมด"

        # who closed it and how, kept with the item (shows as "ปิดโดย ..." and in the daily report)
        closed_by = f"{marked_by or 'ไม่ทราบชื่อ'} · ปิด {scope}"

        ids = [r[0] for r in selected]
        now_iso = bkk_now().isoformat()
        conn.executemany(
            "UPDATE file_events SET downloaded = 1, downloaded_at = ?, downloaded_by = ? WHERE id = ? AND downloaded = 0",
            [(now_iso, closed_by, i) for i in ids]
        )
        conn.execute(
            "INSERT INTO close_batches (user_id, ids_json, description, created_at) VALUES (?,?,?,?)",
            (user_id, json.dumps(ids), scope, now_iso)
        )

        lines = [f"ปิดสถานะ {scope} {len(ids)} รายการแล้ว ✅", ""]
        last_group = None
        for _, g, sname, fname, tcontent in selected:
            if not group_name and g != last_group:
                lines.append(f"📁 {group_close_token(g)}")
                last_group = g
            label = _item_label(fname, tcontent)
            lines.append(f"✅ {label}" if sender_name else f"✅ {label} ({sname})")
        lines.append("")
        lines.append(_remaining_summary(conn))
        lines.append('ปิดผิด? พิมพ์ "ย้อนกลับ" เพื่อเปิดรายการชุดนี้คืน')
        return "\n".join(lines)
    finally:
        conn.close()

def mark_downloaded(sender_name=None, marked_by_name=None):
    """Kept for backward compatibility; closing now goes through handle_close_command."""
    return handle_close_command(sender_name or "", None)

def undo_last_close(user_id):
    """Reopens the items closed by this user's most recent close action.
    Repeating it steps further back, one action at a time."""
    conn = db()
    try:
        batch = conn.execute(
            "SELECT id, ids_json, description, created_at FROM close_batches WHERE user_id = ? ORDER BY id DESC LIMIT 1",
            (user_id,)
        ).fetchone()
        if not batch:
            return "ไม่มีรายการที่ปิดล่าสุดให้ย้อนกลับครับ"
        batch_id, ids_json, description, created_at = batch
        try:
            ids = json.loads(ids_json) or []
        except (ValueError, TypeError):
            ids = []
        if ids:
            conn.executemany(
                "UPDATE file_events SET downloaded = 0, downloaded_at = NULL, downloaded_by = NULL WHERE id = ?",
                [(i,) for i in ids]
            )
        conn.execute("DELETE FROM close_batches WHERE id = ?", (batch_id,))
        when = datetime.datetime.fromisoformat(created_at).strftime("%d/%m %H:%M")
        return (f"↩️ เปิดสถานะคืน {len(ids)} รายการแล้ว (ที่ปิด {description} เมื่อ {when} น.)\n"
                f"{_remaining_summary(conn)}")
    finally:
        conn.close()

# ---------- Daily backend report (CSV) ----------
REPORT_CSV_COLUMNS = [
    "วันที่สรุป", "กลุ่ม", "ผู้ส่ง", "ประเภท", "ชื่อไฟล์", "ข้อความ",
    "วันที่ส่งจริง", "สถานะ", "โหลดโดย", "วันที่โหลด", "หมายเหตุ",
]

def build_daily_report_rows(date=None):
    """One row per tagged event relevant to `date`: everything actually sent
    ON that date (file or message, downloaded or not - a full record, unlike
    the LINE summary which hides downloaded items), plus any FILE event
    tagged on an earlier date that's still not downloaded as of `date`
    (carried forward, same as the LINE digest's overdue list, flagged in
    the "หมายเหตุ" column instead of colored red)."""
    date = date or bkk_today()
    start = datetime.datetime.combine(date, datetime.time.min).isoformat()
    end = datetime.datetime.combine(date, datetime.time.max).isoformat()

    conn = db()
    rows = conn.execute(
        """SELECT group_name, sender_name, event_type, file_name, text_content, sent_at,
                  downloaded, downloaded_at, downloaded_by
           FROM file_events
           WHERE tagged = 1 AND (
               (sent_at BETWEEN ? AND ?)
               OR (downloaded = 0 AND sent_at < ?)
           )
           ORDER BY group_name, sender_name, sent_at""",
        (start, end, start)
    ).fetchall()
    conn.close()

    out = []
    for group_name, sender_name, event_type, file_name, text_content, sent_at, downloaded, downloaded_at, downloaded_by in rows:
        sent_dt = datetime.datetime.fromisoformat(sent_at)
        carried_over = sent_dt.date() < date
        out.append({
            "วันที่สรุป": date.strftime("%d/%m/%Y"),
            "กลุ่ม": group_name,
            "ผู้ส่ง": sender_name,
            "ประเภท": "ไฟล์" if event_type == "file" else "ข้อความ",
            "ชื่อไฟล์": file_name or "",
            "ข้อความ": (text_content or "").replace("\n", " "),
            "วันที่ส่งจริง": sent_dt.strftime("%d/%m/%Y %H:%M"),
            "สถานะ": "ปิดแล้ว" if downloaded else "ยังไม่ปิด",
            "โหลดโดย": downloaded_by or "",
            "วันที่โหลด": datetime.datetime.fromisoformat(downloaded_at).strftime("%d/%m/%Y %H:%M") if downloaded_at else "",
            "หมายเหตุ": "ยกยอดจากวันก่อน" if carried_over else "ใหม่วันนี้",
        })
    return out

def build_daily_report_csv(date=None):
    rows = build_daily_report_rows(date)
    buf = io.StringIO()
    buf.write("﻿")  # BOM so Excel opens the Thai text as UTF-8 instead of mangling it
    writer = csv.DictWriter(buf, fieldnames=REPORT_CSV_COLUMNS)
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()

@app.route("/report/daily.csv", methods=["GET"])
def report_daily_csv():
    if REPORT_ACCESS_KEY and request.args.get("key", "") != REPORT_ACCESS_KEY:
        abort(403)
    date_str = request.args.get("date", "")
    try:
        date = datetime.date.fromisoformat(date_str) if date_str else bkk_today()
    except ValueError:
        abort(400, "date must be YYYY-MM-DD")
    return Response(
        build_daily_report_csv(date),
        mimetype="text/csv",
        headers={"Content-Disposition": f'attachment; filename="report_{date.isoformat()}.csv"'}
    )

# ---------- Daily backend report (styled Excel) ----------
# Two extra columns that exist only in the Excel view (not in the CSV, and
# not stored in the DB): "Recheck" is a tickbox the person checks by hand
# after physically verifying the item, and "ผลตรวจสอบ" is a live formula
# comparing that tick against the "สถานะ" column.
COL_RECHECK = "Recheck"
COL_RECHECK_RESULT = "ผลตรวจสอบ"
REPORT_XLSX_COLUMNS = (
    REPORT_CSV_COLUMNS[:8]
    + [COL_RECHECK, COL_RECHECK_RESULT]
    + REPORT_CSV_COLUMNS[8:]
)  # ... "สถานะ", Recheck, ผลตรวจสอบ, "โหลดโดย", ...
REPORT_COL_WIDTHS = [12, 26, 16, 10, 40, 42, 18, 12, 13, 14, 14, 18, 18]

NAVY = "1F3864"
TEAL = "0B6E4F"
GREEN_BG = "E8F5E9"
GREEN_TXT = "1B5E20"
RED_BG = "FDECEA"
RED_TXT = "C62828"
ORANGE_TXT = "E65100"
GROUP_BG = "D9EAD6"
ZEBRA_BG = "F7F9FC"
BLUE_TXT = "1565C0"
PURPLE_TXT = "6A1B9A"
GREY_TXT = "616161"
AMBER_BG = "FFC107"
AMBER_TXT = "7A5900"

THIN = Side(style="thin", color="D9D9D9")
CELL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def build_daily_report_xlsx(date=None):
    """A styled, color-coded Excel report: title banner, KPI summary strip,
    a highlighted group header row per กลุ่ม (same visual language as the
    LINE Flex summary's badges), zebra-striped data rows, and the สถานะ
    column colored green/red so it reads at a glance."""
    date = date or bkk_today()
    rows = build_daily_report_rows(date)
    date_str = date.strftime("%d/%m/%Y")

    total = len(rows)
    closed = sum(1 for r in rows if r["สถานะ"] == "ปิดแล้ว")
    pending = total - closed

    wb = Workbook()
    ws = wb.active
    ws.title = "รายงานรายวัน"
    ncols = len(REPORT_XLSX_COLUMNS)
    last_col_letter = get_column_letter(ncols)

    # --- Title banner ---
    ws.merge_cells(f"A1:{last_col_letter}1")
    title_cell = ws["A1"]
    title_cell.value = f"📋 รายงานข้อมูลเอกสาร/ข้อความประจำวันที่ {date_str}"
    title_cell.font = Font(size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor=NAVY)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # --- KPI summary strip (3 colored cards) ---
    third = max(ncols // 3, 1)
    kpi_specs = [
        (1, third, f"ทั้งหมด {total} รายการ", "455A64"),
        (third + 1, third * 2, f"✅ ปิดแล้ว {closed}", TEAL),
        (third * 2 + 1, ncols, f"🔲 ยังไม่ปิด {pending}", "C62828" if pending else "9E9E9E"),
    ]
    for start, end, text, color in kpi_specs:
        start_letter = get_column_letter(start)
        end_letter = get_column_letter(max(end, start))
        ws.merge_cells(f"{start_letter}2:{end_letter}2")
        cell = ws[f"{start_letter}2"]
        cell.value = text
        cell.font = Font(size=12, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    # --- blank spacer row ---
    ws.row_dimensions[3].height = 6

    # --- header row ---
    header_row = 4
    for i, col_name in enumerate(REPORT_XLSX_COLUMNS, start=1):
        is_manual_col = col_name in (COL_RECHECK, COL_RECHECK_RESULT)
        header_label = "✔️ Recheck" if col_name == COL_RECHECK else col_name
        cell = ws.cell(row=header_row, column=i, value=header_label)
        cell.font = Font(bold=True, color=AMBER_TXT if is_manual_col else "FFFFFF")
        cell.fill = PatternFill("solid", fgColor=AMBER_BG if is_manual_col else TEAL)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = CELL_BORDER
    ws.row_dimensions[header_row].height = 22

    # --- grouped data rows ---
    by_group = {}
    group_order = []
    for r in rows:
        g = r["กลุ่ม"]
        if g not in by_group:
            by_group[g] = []
            group_order.append(g)
        by_group[g].append(r)

    r_idx = header_row + 1
    for group_name in group_order:
        ws.merge_cells(f"A{r_idx}:{last_col_letter}{r_idx}")
        gcell = ws.cell(row=r_idx, column=1, value=f"📁 {group_name}")
        gcell.font = Font(bold=True, color=TEAL, size=11)
        gcell.fill = PatternFill("solid", fgColor=GROUP_BG)
        gcell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r_idx].height = 20
        r_idx += 1

        for zebra_i, row_dict in enumerate(by_group[group_name]):
            status_col_letter = get_column_letter(REPORT_XLSX_COLUMNS.index("สถานะ") + 1)
            recheck_col_letter = get_column_letter(REPORT_XLSX_COLUMNS.index(COL_RECHECK) + 1)
            for col_i, col_name in enumerate(REPORT_XLSX_COLUMNS, start=1):
                if col_name == COL_RECHECK:
                    # Plain text cell (not boolean - Excel ignores custom
                    # number formats on TRUE/FALSE values and always shows
                    # the literal word, so a tickbox glyph via number_format
                    # doesn't work). Default "☐" = unticked; pick "✓" from
                    # the cell's dropdown to tick it (see DataValidation
                    # below) - conditional formatting turns it green once
                    # ticked, so it reads like a toggle button.
                    cell = ws.cell(row=r_idx, column=col_i, value="☐")
                    cell.font = Font(size=16, bold=True, color="9E9E9E")
                    cell.fill = PatternFill("solid", fgColor="FFFDE7")
                    cell.border = Border(
                        left=Side(style="medium", color=AMBER_TXT), right=Side(style="medium", color=AMBER_TXT),
                        top=Side(style="medium", color=AMBER_TXT), bottom=Side(style="medium", color=AMBER_TXT),
                    )
                elif col_name == COL_RECHECK_RESULT:
                    cell = ws.cell(
                        row=r_idx, column=col_i,
                        value=f'=IF(({recheck_col_letter}{r_idx}="✓")=({status_col_letter}{r_idx}="ปิดแล้ว"),"ตรงกัน","ไม่ตรงกัน")'
                    )
                    cell.font = Font(bold=True)
                else:
                    cell = ws.cell(row=r_idx, column=col_i, value=row_dict[col_name])

                if col_name != COL_RECHECK:
                    cell.border = CELL_BORDER
                cell.alignment = Alignment(
                    horizontal="left", vertical="center",
                    wrap_text=col_name in ("ชื่อไฟล์", "ข้อความ")
                )
                if zebra_i % 2 == 1 and col_name != COL_RECHECK:
                    cell.fill = PatternFill("solid", fgColor=ZEBRA_BG)

                if col_name in (COL_RECHECK, COL_RECHECK_RESULT):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == "สถานะ":
                    cell.font = Font(bold=True, color=GREEN_TXT if row_dict[col_name] == "ปิดแล้ว" else RED_TXT)
                    cell.fill = PatternFill("solid", fgColor=GREEN_BG if row_dict[col_name] == "ปิดแล้ว" else RED_BG)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == "ประเภท":
                    cell.font = Font(bold=True, color=BLUE_TXT if row_dict[col_name] == "ไฟล์" else PURPLE_TXT)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_name == "หมายเหตุ" and row_dict[col_name] == "ยกยอดจากวันก่อน":
                    cell.font = Font(bold=True, color=ORANGE_TXT)
                elif col_name == "วันที่สรุป" or col_name == "วันที่ส่งจริง" or col_name == "วันที่โหลด":
                    cell.font = Font(color=GREY_TXT)
            ws.row_dimensions[r_idx].height = 22
            r_idx += 1

    if total == 0:
        ws.merge_cells(f"A{r_idx}:{last_col_letter}{r_idx}")
        empty_cell = ws.cell(row=r_idx, column=1, value="ไม่มีรายการสำหรับวันที่นี้ครับ 🎉")
        empty_cell.font = Font(italic=True, color=GREY_TXT)
        empty_cell.alignment = Alignment(horizontal="center", vertical="center")
        r_idx += 1

    if total:
        last_data_row = r_idx - 1
        recheck_col_letter = get_column_letter(REPORT_XLSX_COLUMNS.index(COL_RECHECK) + 1)
        result_col_letter = get_column_letter(REPORT_XLSX_COLUMNS.index(COL_RECHECK_RESULT) + 1)

        # Click-to-toggle dropdown on each Recheck cell: click the cell, pick
        # "✓" (ticked) or "☐" (unticked) from the dropdown arrow - a single
        # pick swaps between the two, so it behaves like a toggle switch.
        dv = DataValidation(
            type="list", formula1='"✓,☐"', allow_blank=True,
            showInputMessage=True, showDropDown=False,
        )
        dv.promptTitle = "Recheck"
        dv.prompt = "คลิกลูกศรแล้วเลือก ✓ เพื่อติ๊ก หรือ ☐ เพื่อยกเลิกติ๊ก"
        recheck_range = f"{recheck_col_letter}{header_row + 1}:{recheck_col_letter}{last_data_row}"
        dv.add(recheck_range)
        ws.add_data_validation(dv)

        # Recheck cell itself turns green/bold once ticked, so it reads
        # like a pressed toggle button rather than a static dropdown.
        ws.conditional_formatting.add(
            recheck_range,
            CellIsRule(operator="equal", formula=['"✓"'],
                       fill=PatternFill("solid", fgColor=GREEN_BG), font=Font(bold=True, size=16, color=GREEN_TXT))
        )

        # Color the formula result green/red once Excel evaluates it.
        result_range = f"{result_col_letter}{header_row + 1}:{result_col_letter}{last_data_row}"
        ws.conditional_formatting.add(
            result_range,
            CellIsRule(operator="equal", formula=['"ตรงกัน"'],
                       fill=PatternFill("solid", fgColor=GREEN_BG), font=Font(bold=True, color=GREEN_TXT))
        )
        ws.conditional_formatting.add(
            result_range,
            CellIsRule(operator="equal", formula=['"ไม่ตรงกัน"'],
                       fill=PatternFill("solid", fgColor=RED_BG), font=Font(bold=True, color=RED_TXT))
        )

    # --- column widths, freeze header, autofilter ---
    for i, width in enumerate(REPORT_COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.freeze_panes = f"A{header_row + 1}"
    if total:
        ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{r_idx - 1}"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

@app.route("/report/daily.xlsx", methods=["GET"])
def report_daily_xlsx():
    if REPORT_ACCESS_KEY and request.args.get("key", "") != REPORT_ACCESS_KEY:
        abort(403)
    date_str = request.args.get("date", "")
    try:
        date = datetime.date.fromisoformat(date_str) if date_str else bkk_today()
    except ValueError:
        abort(400, "date must be YYYY-MM-DD")
    return Response(
        build_daily_report_xlsx(date),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="report_{date.isoformat()}.xlsx"'}
    )

def daily_report_url(date=None, fmt="xlsx"):
    date = date or bkk_today()
    base = SERVICE_BASE_URL
    if not base:
        try:
            base = request.host_url.rstrip("/")
        except RuntimeError:
            return None  # no SERVICE_BASE_URL set and no request context (e.g. the scheduler)
    url = f"{base}/report/daily.{fmt}?date={date.isoformat()}"
    if REPORT_ACCESS_KEY:
        url += f"&key={REPORT_ACCESS_KEY}"
    return url

# ---------- Daily digest ----------
def send_daily_digest():
    # shared digest: every tracked admin gets the same full summary
    alt_text, contents = build_summary_flex()
    recipients = set(TAGGED_USER_IDS)
    if OWNER_USER_ID:
        recipients.add(OWNER_USER_ID)
    for uid in recipients:
        if contents is None:
            push_to(uid, alt_text)
        else:
            push_flex_to(uid, alt_text, contents)

    # folded into the same report message (no extra push) to save LINE quota
    try:
        unanswered = count_unanswered_docs()
    except Exception:
        app.logger.exception("could not count unanswered docs")
        unanswered = 0
    reply_note = (f'\n\n💬 เอกสารที่ยังรอคำตอบ {unanswered} ฉบับ — พิมพ์ "คำตอบ" เพื่อดู'
                  if unanswered else "")

    report_url = daily_report_url()
    if report_url:
        for uid in recipients:
            push_to(uid, f"📄 รายงานข้อมูลวันนี้ (ไฟล์ Excel จัดสีสวยพร้อมเปิด):\n{report_url}{reply_note}")
    elif reply_note:
        for uid in recipients:
            push_to(uid, reply_note.strip())

# ---------- Overdue (leftover from previous days) morning digest ----------
OVERDUE_COLOR = "#C62828"  # red

def build_overdue_flex():
    """Returns (alt_text, flex_contents_dict) for files tagged on a PREVIOUS day
    that are still not marked downloaded, highlighted in red. Returns (None, None)
    when there's nothing overdue, so the caller can skip sending anything."""
    today_start = datetime.datetime.combine(bkk_today(), datetime.time.min).isoformat()

    conn = db()
    rows = conn.execute(
        """SELECT group_name, sender_name, event_type, file_name, text_content, sent_at, mention_spans
           FROM file_events
           WHERE tagged = 1 AND downloaded = 0 AND sent_at < ?
           ORDER BY group_name, sender_name, sent_at""",
        (today_start,)
    ).fetchall()
    conn.close()

    if not rows:
        return (None, None)

    by_group = {}
    for group_name, sender_name, event_type, file_name, text_content, sent_at, mention_spans in rows:
        by_group.setdefault(group_name, {}).setdefault(sender_name, []).append(
            (event_type, file_name, text_content, sent_at, mention_spans)
        )

    txt = flex_txt
    body = [txt("⚠️ รายการค้างจากวันก่อนหน้า (ยังไม่ปิดสถานะ)", weight="bold", size="lg", color=OVERDUE_COLOR)]
    body.append(txt(f"ทั้งหมด {len(rows)} รายการ", size="sm", color=OVERDUE_COLOR, margin="xs"))
    body.append({"type": "separator", "margin": "md"})

    for group_name, by_sender in by_group.items():
        body.append(group_name_badge(group_name, with_close_button=True))
        for sender_name, items in by_sender.items():
            body.append(sender_line(f"👤 {sender_name}  (ค้าง {len(items)})", group_name, sender_name, True))
            for event_type, file_name, text_content, sent_at, mention_spans in items:
                date_str = datetime.datetime.fromisoformat(sent_at).strftime("%d/%m %H:%M")
                if event_type == "file":
                    body.append(two_tone_text(f"🔴 {file_name}", OVERDUE_COLOR, f"  ({date_str} น.)",
                                               size="sm", margin="xs"))
                    if text_content:
                        body.append(message_text_component(
                            "💬 ", text_content, mention_spans, size="xs", margin="xs"
                        ))
                else:
                    body.append(message_text_component(
                        "🔴 💬 ", text_content, mention_spans, suffix=f"  ({date_str} น.)", size="sm", margin="xs"
                    ))

    body.append({"type": "separator", "margin": "lg"})
    body.append(txt(CLOSE_HELP_TEXT, size="xs", color="#888888", margin="md"))

    contents = {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "xs",
            "contents": body
        }
    }
    alt_text = f"⚠️ มีรายการค้างจากวันก่อนหน้ายังไม่ปิดสถานะ {len(rows)} รายการ"
    return (alt_text, contents)

def send_overdue_digest():
    alt_text, contents = build_overdue_flex()
    if contents is None:
        return  # nothing overdue - stay quiet
    recipients = set(TAGGED_USER_IDS)
    if OWNER_USER_ID:
        recipients.add(OWNER_USER_ID)
    for uid in recipients:
        push_flex_to(uid, alt_text, contents)

scheduler = BackgroundScheduler(timezone="Asia/Bangkok")
scheduler.add_job(send_daily_digest, "cron", hour=DIGEST_HOUR, minute=DIGEST_MINUTE)
scheduler.add_job(send_overdue_digest, "cron", hour=OVERDUE_HOUR, minute=OVERDUE_MINUTE)
scheduler.add_job(scheduled_reply_followup, "cron", day_of_week=REMIND_DAYS_OF_WEEK,
                  hour=REMIND_HOUR, minute=REMIND_MINUTE)
scheduler.start()

@app.route("/", methods=["GET"])
def health():
    return "LINE file-tracker bot is running."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
