import os
import sqlite3
import datetime

from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi, ReplyMessageRequest,
    PushMessageRequest, TextMessage
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
TAG_LOOKBACK_MINUTES = int(os.environ.get("TAG_LOOKBACK_MINUTES", "30"))
DB_PATH = os.environ.get("DB_PATH", "tracker.db")

app = Flask(__name__)
handler = WebhookHandler(CHANNEL_SECRET)
configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)

# ---------- Storage ----------
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
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
            downloaded_at TEXT
        )
    """)
    # pending tag markers, for when the tag message arrives BEFORE the file(s)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS pending_tags (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            group_id TEXT,
            sender_id TEXT,
            text_content TEXT,
            created_at TEXT
        )
    """)
    return conn

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

def reply(reply_token, text):
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.reply_message(ReplyMessageRequest(
            reply_token=reply_token,
            messages=[TextMessage(text=text)]
        ))

def push_to_owner(text):
    if not OWNER_USER_ID:
        return
    with ApiClient(configuration) as api_client:
        api = MessagingApi(api_client)
        api.push_message(PushMessageRequest(
            to=OWNER_USER_ID,
            messages=[TextMessage(text=text)]
        ))

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
    now = datetime.datetime.now()
    now_iso = now.isoformat()
    cutoff = (now - datetime.timedelta(minutes=TAG_LOOKBACK_MINUTES)).isoformat()

    conn = db()
    # was this sender already tagged recently (tag arrived BEFORE the file)?
    pending = conn.execute(
        "SELECT id, text_content FROM pending_tags WHERE group_id = ? AND sender_id = ? AND created_at >= ? ORDER BY created_at DESC LIMIT 1",
        (group_id, sender_id, cutoff)
    ).fetchone()
    tagged_now = 1 if pending else 0
    tag_text = pending[1] if pending else None

    conn.execute(
        """INSERT INTO file_events
           (group_id, group_name, sender_id, sender_name, event_type, msg_type, file_name, text_content, sent_at, tagged, tagged_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (group_id, group_name, sender_id, sender_name, "file", msg.type, file_name, tag_text, now_iso,
         tagged_now, now_iso if tagged_now else None)
    )
    conn.commit()
    conn.close()

@handler.add(MessageEvent, message=TextMessageContent)
def handle_text(event):
    source = event.source
    text = event.message.text.strip()

    # --- 1:1 chat with the owner: on-demand commands ---
    if source.type == "user":
        if text in ("สรุป", "สรุปไฟล์", "summary", "เช็คไฟล์"):
            reply(event.reply_token, build_summary_text())
        elif text in ("id", "whoami", "ไอดี"):
            reply(event.reply_token, f"user_id ของคุณคือ:\n{source.user_id}")
        elif text in ("โหลดแล้ว", "✅", "👍"):
            reply(event.reply_token, mark_downloaded())
        elif text.startswith("โหลดแล้ว "):
            reply(event.reply_token, mark_downloaded(text[len("โหลดแล้ว "):].strip()))
        elif text.startswith("✅ ") or text.startswith("👍 "):
            reply(event.reply_token, mark_downloaded(text[2:].strip()))
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
    now = datetime.datetime.now()
    now_iso = now.isoformat()
    cutoff = (now - datetime.timedelta(minutes=TAG_LOOKBACK_MINUTES)).isoformat()

    conn = db()

    if group_id in PM_MODE_GROUP_IDS:
        # PM mode: the tag message itself is the event, no file needed
        group_name = get_group_name(group_id)
        conn.execute(
            """INSERT INTO file_events
               (group_id, group_name, sender_id, sender_name, event_type, msg_type, text_content, sent_at, tagged, tagged_at)
               VALUES (?,?,?,?,?,?,?,?,1,?)""",
            (group_id, group_name, sender_id, sender_name, "pm", "text", text, now_iso, now_iso)
        )
    else:
        # file mode: match backward to recent un-tagged files from this sender,
        # and attach this tag's own message text to those files
        cur = conn.execute(
            """UPDATE file_events SET tagged = 1, tagged_at = ?, text_content = ?
               WHERE group_id = ? AND sender_id = ? AND tagged = 0 AND sent_at >= ? AND event_type = 'file'""",
            (now_iso, text, group_id, sender_id, cutoff)
        )
        # also leave a pending marker (with the message text) in case file(s) arrive AFTER this tag
        conn.execute(
            "INSERT INTO pending_tags (group_id, sender_id, text_content, created_at) VALUES (?,?,?,?)",
            (group_id, sender_id, text, now_iso)
        )

    conn.commit()
    conn.close()

# ---------- Summary building ----------
def build_summary_text(date=None):
    date = date or datetime.date.today()
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

def mark_downloaded(sender_name=None):
    conn = db()
    if sender_name:
        cur = conn.execute(
            "UPDATE file_events SET downloaded = 1, downloaded_at = ? WHERE tagged = 1 AND downloaded = 0 AND event_type = 'file' AND sender_name = ?",
            (datetime.datetime.now().isoformat(), sender_name)
        )
    else:
        cur = conn.execute(
            "UPDATE file_events SET downloaded = 1, downloaded_at = ? WHERE tagged = 1 AND downloaded = 0 AND event_type = 'file'",
            (datetime.datetime.now().isoformat(),)
        )
    count = cur.rowcount
    conn.commit()
    conn.close()

    if count == 0:
        return f"ไม่มีไฟล์ค้างให้ติ๊กครับ{' (ของ ' + sender_name + ')' if sender_name else ''}"
    target = f"ของ {sender_name} " if sender_name else ""
    return f"ติ๊กว่าโหลดแล้ว {target}{count} ไฟล์ ✅"

# ---------- Daily digest ----------
def send_daily_digest():
    push_to_owner(build_summary_text())

scheduler = BackgroundScheduler(timezone="Asia/Bangkok")
scheduler.add_job(send_daily_digest, "cron", hour=DIGEST_HOUR, minute=DIGEST_MINUTE)
scheduler.start()

@app.route("/", methods=["GET"])
def health():
    return "LINE file-tracker bot is running."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
