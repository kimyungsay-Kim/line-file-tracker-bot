import os
import io
import csv
import json
import datetime
from zoneinfo import ZoneInfo

import libsql_client
import requests

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

def reply_text_flex_text(reply_token, lead_text, alt_text, contents_dict, trailing_text):
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(ReplyMessageRequest(
            reply_token=reply_token,
            messages=[
                TextMessage(text=lead_text),
                FlexMessage(alt_text=alt_text, contents=FlexContainer.from_dict(contents_dict)),
                TextMessage(text=trailing_text),
            ]
        ))

LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"

def push_mention_all_to_group(group_id, flex_alt_text, flex_contents, lead_text, mention_text="โปรดตรวจสอบครับ"):
    """Pushes a lead announcement text, then the Flex Carousel, then an
    '@All'-mention text message, to the given group. Uses a raw HTTP call
    instead of the messaging SDK's typed message classes because the
    installed line-bot-sdk version here doesn't expose LINE's textV2/mention
    payload - this builds that payload directly, per LINE's documented
    format."""
    messages = [
        {"type": "text", "text": lead_text},
        {"type": "flex", "altText": flex_alt_text, "contents": flex_contents},
        {
            "type": "textV2",
            "text": f"{{user1}} {mention_text}",
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
        if text in ("ยกเลิก", "cancel"):
            if get_inspection_preview(user_id) is not None:
                clear_inspection_preview(user_id)
                reply(event.reply_token, "ยกเลิกรายการที่รอส่งแล้วครับ")
            return

        if "|" in text:
            parsed_items = parse_inspection_paste(text)
            if parsed_items:
                alt_text, contents = build_inspection_carousel_flex(parsed_items)
                save_inspection_preview(user_id, parsed_items)
                reply_text_flex_text(
                    event.reply_token,
                    inspection_announce_text(len(parsed_items)),
                    alt_text, contents,
                    'ถ้าข้อมูลถูกต้อง พิมพ์ "ส่ง" เพื่อโพสต์เข้ากลุ่ม หรือ "ยกเลิก" เพื่อยกเลิกครับ'
                )
            else:
                reply(event.reply_token,
                      'รูปแบบไม่ถูกต้องครับ แต่ละบรรทัดต้องมีอย่างน้อย 4 ช่องคั่นด้วย "|" คือ '
                      'code|เลขที่เอกสาร|เรื่อง|วันที่ เช่น\n'
                      'AR-329|RIH-QA-0123|ขออนุมัติวัสดุ Shop Drawing|23/10/2026')
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
            reply(event.reply_token, mark_downloaded(marked_by_name=get_profile_name(None, source.user_id)))
        elif text.startswith("โหลดแล้ว "):
            reply(event.reply_token, mark_downloaded(
                text[len("โหลดแล้ว "):].strip(), marked_by_name=get_profile_name(None, source.user_id)
            ))
        elif text.startswith("✅ ") or text.startswith("👍 "):
            reply(event.reply_token, mark_downloaded(
                text[2:].strip(), marked_by_name=get_profile_name(None, source.user_id)
            ))
        elif text in ("รายงาน", "report", "export"):
            reply(event.reply_token,
                  f"ดาวน์โหลดรายงานวันนี้ (CSV เปิดด้วย Excel ได้เลย):\n{daily_report_url()}\n\n"
                  f"ดูวันอื่นได้โดยเติม ?date=YYYY-MM-DD ต่อท้ายลิงก์")
        return

    if source.type != "group":
        return

    group_id = source.group_id

    # utility: anyone can type "groupid" in a group to get its ID for setup
    if text.lower() in ("groupid", "ไอดีกลุ่ม"):
        reply(event.reply_token, f"group_id ของกลุ่มนี้คือ:\n{group_id}")
        return

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

def inspection_announce_text(n):
    return f"📋✨ แจ้งข้อมูลเอกสาร Request จาก บมจ.พรีบิลท์ (PB) จำนวน {n} ฉบับ ดังนี้ครับ 📑"

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
INSPECTION_DOCNUM_COLOR = "#1E88E5"      # เลขที่เอกสาร (เช่น AR-333)
INSPECTION_CODE_COLOR = "#6A1B9A"        # Code (เช่น RIH-I02222)
INSPECTION_SUBJECT_COLOR = "#212121"     # เรื่อง
INSPECTION_DATE_COLOR = "#2E7D32"        # วันที่ส่งเอกสารเข้าระบบ
INSPECTION_EXTRA_NOTE_COLOR = "#EF6C00"  # วันที่เทคอนกรีต/ก่อ/ฉาบ/อื่นๆ

INSPECTION_MAX_ITEMS = 30  # safety cap so the flex message doesn't blow past LINE's size limit

def build_inspection_carousel_flex(items):
    """A SINGLE bubble listing every document stacked vertically, one block
    per document (separated by a divider line), each block's own fields
    also stacked vertically as an emoji + label line followed by the value,
    color-coded per field. Deliberately not a LINE "carousel" (separate
    swipeable cards) - with everything in one bubble, the whole list reads
    top-to-bottom in a single scroll instead of needing to swipe sideways
    between documents."""
    txt = flex_txt

    def field(emoji, label, value, color, value_size="md", bold=False):
        return {
            "type": "box", "layout": "vertical", "margin": "md", "spacing": "xs",
            "contents": [
                txt(f"{emoji} {label}", size="xs", color=INSPECTION_LABEL_COLOR, weight="bold"),
                txt(value, size=value_size, color=color, wrap=True, weight="bold" if bold else "regular"),
            ]
        }

    shown = items[:INSPECTION_MAX_ITEMS]
    body = []
    for i, it in enumerate(shown):
        if i > 0:
            body.append({"type": "separator", "margin": "xl"})
        body.append(field("📌", "เลขที่เอกสาร", it["code"], INSPECTION_DOCNUM_COLOR, value_size="lg", bold=True))
        body.append(field("🔖", "Code", it["doc_number"], INSPECTION_CODE_COLOR))
        body.append(field("📝", "เรื่อง", it["subject"], INSPECTION_SUBJECT_COLOR, value_size="sm"))
        body.append(field("📅", "วันที่ส่งเอกสารเข้าระบบ", it["date"], INSPECTION_DATE_COLOR, value_size="sm"))
        if it.get("extra_note"):
            body.append(field("🧱", "วันที่เทคอนกรีต/ก่อ/ฉาบ/อื่นๆ", it["extra_note"],
                               INSPECTION_EXTRA_NOTE_COLOR, value_size="sm"))

    if len(items) > INSPECTION_MAX_ITEMS:
        body.append({"type": "separator", "margin": "xl"})
        body.append(txt(f"… มีอีก {len(items) - INSPECTION_MAX_ITEMS} รายการที่ไม่ได้แสดง",
                         size="xs", color=INSPECTION_LABEL_COLOR, margin="md"))

    contents = {
        "type": "bubble",
        "body": {"type": "box", "layout": "vertical", "spacing": "sm", "contents": body, "paddingAll": "lg"},
    }
    alt_text = f"เอกสารตรวจสอบ {len(items)} รายการ"
    return alt_text, contents

def handle_inspection_confirm(user_id):
    items = get_inspection_preview(user_id)
    if not items:
        return 'ไม่มีรายการที่รอยืนยันอยู่ครับ ลองวางข้อมูลเอกสารรูปแบบ "code|doc_number|subject|date" ก่อนนะครับ'
    if not INSPECTION_TARGET_GROUP_ID:
        return "ยังไม่ได้ตั้งค่า INSPECTION_TARGET_GROUP_ID ปลายทางครับ แจ้งผู้ดูแลระบบด้วย"
    alt_text, contents = build_inspection_carousel_flex(items)
    try:
        push_mention_all_to_group(INSPECTION_TARGET_GROUP_ID, alt_text, contents, inspection_announce_text(len(items)))
    except Exception:
        return "โพสต์เข้ากลุ่มไม่สำเร็จครับ ลองใหม่อีกครั้ง หรือแจ้งผู้ดูแลระบบ"
    clear_inspection_preview(user_id)
    return "โพสต์เข้ากลุ่มเรียบร้อยแล้วครับ"

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
        lines.append("\nพิมพ์ \"โหลดแล้ว\" หรือแค่ \"✅\" เพื่อติ๊กว่าโหลดครบหมดแล้ว หรือ \"✅ <ชื่อคน>\" เพื่อติ๊กเฉพาะคนนั้น")
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
               (sent_at BETWEEN ? AND ? AND NOT (event_type = 'file' AND downloaded = 1))
               OR (event_type = 'file' AND downloaded = 0 AND sent_at < ?)
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
        body.append(txt(f"📁 {group_name}", weight="bold", size="md", color="#0B6E4F", margin="lg"))
        for sender_name, items in by_sender.items():
            if item_count >= MAX_ITEMS:
                truncated = True
                break
            pending_count = sum(1 for e, _, _, _, d, *_ in items if e == "file" and not d)
            pending_total += pending_count
            label = f"  (ยังไม่โหลด {pending_count})" if any(e == "file" for e, *_ in items) else ""
            sender_line = f"👤 {sender_name}{label}"
            body.append(txt(sender_line, weight="bold", size="sm", color="#222222", margin="sm"))

            for event_type, file_name, text_content, sent_at, downloaded, mention_spans, downloaded_by in items:
                if item_count >= MAX_ITEMS:
                    truncated = True
                    break
                item_count += 1
                sent_dt = datetime.datetime.fromisoformat(sent_at)
                is_overdue = event_type == "file" and not downloaded and sent_dt.date() < date
                t = sent_dt.strftime("%d/%m %H:%M") if is_overdue else sent_dt.strftime("%H:%M")
                if event_type == "file":
                    mark = "✅" if downloaded else ("🔴" if is_overdue else "🔲")
                    name_color = OVERDUE_COLOR if is_overdue else FILE_NAME_COLOR
                    body.append(two_tone_text(f"{mark} {file_name}", name_color, f"  ({t} น.)",
                                               size="sm", margin="xs"))
                    if downloaded and downloaded_by:
                        body.append(txt(f"   โหลดโดย {downloaded_by}", size="xs", color="#0B6E4F", margin="none"))
                    if text_content:
                        body.append(message_text_component(
                            "💬 ", text_content, mention_spans, size="xs", margin="xs"
                        ))
                else:
                    body.append(message_text_component(
                        "💬 ", text_content, mention_spans, suffix=f"  ({t} น.)", size="sm", margin="xs"
                    ))

    if truncated:
        body.append(txt("… มีรายการเพิ่มเติมที่ไม่ได้แสดงครบ", size="xs", color="#999999", margin="md"))

    body.append({"type": "separator", "margin": "lg"})
    if pending_total == 0:
        body.append(txt("ไฟล์งานโหลดครบแล้วครับ 🎉", size="sm", color="#0B6E4F", margin="md"))
    else:
        body.append(txt('พิมพ์ "โหลดแล้ว" หรือแค่ "✅" เพื่อติ๊กว่าโหลดครบหมดแล้ว หรือ "✅ <ชื่อคน>" เพื่อติ๊กเฉพาะคนนั้น',
                         size="xs", color="#888888", margin="md"))

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

def mark_downloaded(sender_name=None, marked_by_name=None):
    conn = db()
    if sender_name:
        pending = conn.execute(
            "SELECT id, sender_name, file_name FROM file_events WHERE tagged = 1 AND downloaded = 0 AND event_type = 'file' AND sender_name = ?",
            (sender_name,)
        ).fetchall()
    else:
        pending = conn.execute(
            "SELECT id, sender_name, file_name FROM file_events WHERE tagged = 1 AND downloaded = 0 AND event_type = 'file'"
        ).fetchall()

    if not pending:
        conn.close()
        return f"ไม่มีไฟล์ค้างให้ติ๊กครับ{' (ของ ' + sender_name + ')' if sender_name else ''}"

    ids = [row[0] for row in pending]
    now_iso = bkk_now().isoformat()
    conn.executemany(
        "UPDATE file_events SET downloaded = 1, downloaded_at = ?, downloaded_by = ? WHERE id = ?",
        [(now_iso, marked_by_name, i) for i in ids]
    )
    conn.commit()
    conn.close()

    target = f"ของ {sender_name} " if sender_name else ""
    lines = [f"ติ๊กว่าโหลดแล้ว {target}{len(pending)} ไฟล์ ✅", ""]
    for _, sname, fname in pending:
        lines.append(f"✅ {fname} ({sname})" if not sender_name else f"✅ {fname}")
    return "\n".join(lines)

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
               OR (event_type = 'file' AND downloaded = 0 AND sent_at < ?)
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
            "สถานะ": ("โหลดแล้ว" if downloaded else "ยังไม่โหลด") if event_type == "file" else "-",
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

def daily_report_url(date=None):
    date = date or bkk_today()
    base = SERVICE_BASE_URL
    if not base:
        try:
            base = request.host_url.rstrip("/")
        except RuntimeError:
            return None  # no SERVICE_BASE_URL set and no request context (e.g. the scheduler)
    url = f"{base}/report/daily.csv?date={date.isoformat()}"
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

    report_url = daily_report_url()
    if report_url:
        for uid in recipients:
            push_to(uid, f"📄 รายงานข้อมูลวันนี้ (CSV เปิดด้วย Excel ได้เลย):\n{report_url}")

# ---------- Overdue (leftover from previous days) morning digest ----------
OVERDUE_COLOR = "#C62828"  # red

def build_overdue_flex():
    """Returns (alt_text, flex_contents_dict) for files tagged on a PREVIOUS day
    that are still not marked downloaded, highlighted in red. Returns (None, None)
    when there's nothing overdue, so the caller can skip sending anything."""
    today_start = datetime.datetime.combine(bkk_today(), datetime.time.min).isoformat()

    conn = db()
    rows = conn.execute(
        """SELECT group_name, sender_name, file_name, text_content, sent_at, mention_spans
           FROM file_events
           WHERE event_type = 'file' AND tagged = 1 AND downloaded = 0 AND sent_at < ?
           ORDER BY group_name, sender_name, sent_at""",
        (today_start,)
    ).fetchall()
    conn.close()

    if not rows:
        return (None, None)

    by_group = {}
    for group_name, sender_name, file_name, text_content, sent_at, mention_spans in rows:
        by_group.setdefault(group_name, {}).setdefault(sender_name, []).append(
            (file_name, text_content, sent_at, mention_spans)
        )

    txt = flex_txt
    body = [txt("⚠️ ไฟล์ค้างจากวันก่อนหน้า (ยังไม่โหลด)", weight="bold", size="lg", color=OVERDUE_COLOR)]
    body.append(txt(f"ทั้งหมด {len(rows)} ไฟล์", size="sm", color=OVERDUE_COLOR, margin="xs"))
    body.append({"type": "separator", "margin": "md"})

    for group_name, by_sender in by_group.items():
        body.append(txt(f"📁 {group_name}", weight="bold", size="md", color="#0B6E4F", margin="lg"))
        for sender_name, items in by_sender.items():
            body.append(txt(f"👤 {sender_name}", weight="bold", size="sm", color="#222222", margin="sm"))
            for file_name, text_content, sent_at, mention_spans in items:
                date_str = datetime.datetime.fromisoformat(sent_at).strftime("%d/%m %H:%M")
                body.append(two_tone_text(f"🔴 {file_name}", OVERDUE_COLOR, f"  ({date_str} น.)",
                                           size="sm", margin="xs"))
                if text_content:
                    body.append(message_text_component(
                        "💬 ", text_content, mention_spans, size="xs", margin="xs"
                    ))

    body.append({"type": "separator", "margin": "lg"})
    body.append(txt('พิมพ์ "โหลดแล้ว" หรือแค่ "✅" เพื่อติ๊กว่าโหลดครบหมดแล้ว หรือ "✅ <ชื่อคน>" เพื่อติ๊กเฉพาะคนนั้น',
                     size="xs", color="#888888", margin="md"))

    contents = {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "xs",
            "contents": body
        }
    }
    alt_text = f"⚠️ มีไฟล์ค้างจากวันก่อนหน้ายังไม่โหลด {len(rows)} ไฟล์"
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
scheduler.start()

@app.route("/", methods=["GET"])
def health():
    return "LINE file-tracker bot is running."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
