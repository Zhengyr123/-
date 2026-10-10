"""阿里邮箱 MCP 服务器：通过 IMAP/SMTP 让 Claude 收发阿里邮箱邮件。

环境变量：
  ALIMAIL_USER       邮箱地址，例如 you@company.com 或 you@aliyun.com
  ALIMAIL_PASSWORD   三方客户端安全密码（不是登录密码）
  ALIMAIL_IMAP_HOST  可选，默认按域名自动判断
  ALIMAIL_SMTP_HOST  可选，默认按域名自动判断
"""

import base64
import email
import imaplib
import os
import re
import smtplib
from datetime import datetime, timedelta
from email.header import decode_header, make_header
from email.message import EmailMessage
from email.utils import parseaddr

try:  # mcp >= 2
    from mcp.server.mcpserver import MCPServer as _Server
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as _Server

USER = os.environ.get("ALIMAIL_USER", "")
PASSWORD = os.environ.get("ALIMAIL_PASSWORD", "")
_personal = USER.lower().endswith("@aliyun.com")
IMAP_HOST = os.environ.get(
    "ALIMAIL_IMAP_HOST", "imap.aliyun.com" if _personal else "imap.qiye.aliyun.com"
)
SMTP_HOST = os.environ.get(
    "ALIMAIL_SMTP_HOST", "smtp.aliyun.com" if _personal else "smtp.qiye.aliyun.com"
)

mcp = _Server("aliyun-mail")


# IMAP 文件夹名使用 modified UTF-7（RFC 3501），中文文件夹需要编解码。
def _utf7_decode(name: str) -> str:
    def repl(m):
        s = m.group(1)
        if s == "":
            return "&"
        s = s.replace(",", "/")
        s += "=" * (-len(s) % 4)
        return base64.b64decode(s).decode("utf-16-be")

    return re.sub(r"&([^-]*)-", repl, name)


def _utf7_encode(name: str) -> str:
    out, buf = [], ""

    def flush():
        nonlocal buf
        if buf:
            b = base64.b64encode(buf.encode("utf-16-be")).decode().rstrip("=")
            out.append("&" + b.replace("/", ",") + "-")
            buf = ""

    for ch in name:
        if 0x20 <= ord(ch) <= 0x7E:
            flush()
            out.append("&-" if ch == "&" else ch)
        else:
            buf += ch
    flush()
    return "".join(out)


def _connect() -> imaplib.IMAP4_SSL:
    if not USER or not PASSWORD:
        raise RuntimeError("请先设置环境变量 ALIMAIL_USER 和 ALIMAIL_PASSWORD")
    conn = imaplib.IMAP4_SSL(IMAP_HOST, 993)
    conn.login(USER, PASSWORD)
    return conn


def _select(conn, folder: str, readonly: bool = True):
    typ, _ = conn.select('"%s"' % _utf7_encode(folder), readonly=readonly)
    if typ != "OK":
        raise RuntimeError(f"无法打开文件夹：{folder}")


def _h(value) -> str:
    return str(make_header(decode_header(value))) if value else ""


def _body(msg) -> str:
    plain, html = None, None
    for part in msg.walk() if msg.is_multipart() else [msg]:
        if part.get_content_maintype() == "multipart" or part.get_filename():
            continue
        payload = part.get_payload(decode=True)
        if payload is None:
            continue
        text = payload.decode(part.get_content_charset() or "utf-8", errors="replace")
        if part.get_content_type() == "text/plain" and plain is None:
            plain = text
        elif part.get_content_type() == "text/html" and html is None:
            html = text
    if plain:
        return plain
    if html:
        html = re.sub(r"(?is)<(script|style).*?</\1>", "", html)
        return re.sub(r"\n\s*\n+", "\n\n", re.sub(r"<[^>]+>", "", html)).strip()
    return ""


@mcp.tool()
def list_folders() -> list[str]:
    """列出邮箱中的所有文件夹。"""
    conn = _connect()
    try:
        _, data = conn.list()
        names = []
        for line in data:
            m = re.search(rb'"([^"]*)"\s*$|(\S+)\s*$', line)
            raw = (m.group(1) or m.group(2)).decode()
            names.append(_utf7_decode(raw))
        return names
    finally:
        conn.logout()


@mcp.tool()
def search_emails(
    folder: str = "INBOX",
    keyword: str = "",
    from_addr: str = "",
    since_days: int = 0,
    unread_only: bool = False,
    limit: int = 20,
) -> list[dict]:
    """搜索邮件，按时间倒序返回摘要。

    keyword 在主题和发件人中匹配（支持中文）；from_addr 按发件人邮箱过滤；
    since_days 只看最近 N 天；返回的 uid 可传给 read_email。
    """
    conn = _connect()
    try:
        _select(conn, folder)
        criteria = []
        if unread_only:
            criteria.append("UNSEEN")
        if since_days:
            day = datetime.now() - timedelta(days=since_days)
            criteria += ["SINCE", day.strftime("%d-%b-%Y")]
        if from_addr:
            criteria += ["FROM", '"%s"' % from_addr]
        _, data = conn.uid("SEARCH", *(criteria or ["ALL"]))
        uids = data[0].split()[::-1]
        # 中文关键词在客户端过滤，最多检查最近 300 封
        if keyword:
            uids = uids[:300]
        results = []
        for i in range(0, len(uids), 50):
            if len(results) >= limit:
                break
            batch = b",".join(uids[i : i + 50])
            _, rows = conn.uid(
                "FETCH", batch, "(FLAGS BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])"
            )
            items = [r for r in rows if isinstance(r, tuple)]
            items.sort(key=lambda r: -int(re.search(rb"UID (\d+)", r[0]).group(1)))
            for meta, raw in items:
                hdr = email.message_from_bytes(raw)
                subject, sender = _h(hdr["Subject"]), _h(hdr["From"])
                if keyword and keyword.lower() not in (subject + sender).lower():
                    continue
                results.append(
                    {
                        "uid": re.search(rb"UID (\d+)", meta).group(1).decode(),
                        "subject": subject,
                        "from": sender,
                        "date": hdr["Date"] or "",
                        "unread": b"\\Seen" not in meta,
                    }
                )
                if len(results) >= limit:
                    break
        return results
    finally:
        conn.logout()


@mcp.tool()
def read_email(uid: str, folder: str = "INBOX", mark_as_read: bool = False) -> dict:
    """读取一封邮件的完整内容（正文和附件名）。"""
    conn = _connect()
    try:
        _select(conn, folder, readonly=not mark_as_read)
        fetch = "(RFC822)" if mark_as_read else "(BODY.PEEK[])"
        _, data = conn.uid("FETCH", uid, fetch)
        raw = next((r[1] for r in data if isinstance(r, tuple)), None)
        if raw is None:
            raise RuntimeError(f"找不到邮件 uid={uid}")
        msg = email.message_from_bytes(raw)
        return {
            "uid": uid,
            "subject": _h(msg["Subject"]),
            "from": _h(msg["From"]),
            "to": _h(msg["To"]),
            "cc": _h(msg["Cc"]),
            "date": msg["Date"] or "",
            "message_id": msg["Message-ID"] or "",
            "attachments": [_h(p.get_filename()) for p in msg.walk() if p.get_filename()],
            "body": _body(msg)[:20000],
        }
    finally:
        conn.logout()


@mcp.tool()
def send_email(
    to: str, subject: str, body: str, cc: str = "", in_reply_to: str = ""
) -> str:
    """发送纯文本邮件。to / cc 可用逗号分隔多个地址；
    回复邮件时把原邮件的 message_id 传给 in_reply_to。"""
    if not USER or not PASSWORD:
        raise RuntimeError("请先设置环境变量 ALIMAIL_USER 和 ALIMAIL_PASSWORD")
    msg = EmailMessage()
    msg["From"] = USER
    msg["To"] = to
    if cc:
        msg["Cc"] = cc
    msg["Subject"] = subject
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
        msg["References"] = in_reply_to
    msg.set_content(body)
    recipients = [parseaddr(a)[1] for a in (to + "," + cc).split(",") if a.strip()]
    with smtplib.SMTP_SSL(SMTP_HOST, 465) as smtp:
        smtp.login(USER, PASSWORD)
        smtp.send_message(msg, to_addrs=recipients)
    return f"已发送给 {', '.join(recipients)}"


if __name__ == "__main__":
    mcp.run()
