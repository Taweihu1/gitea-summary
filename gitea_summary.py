#!/usr/bin/env python3
"""
gitea_summary.py

Reads Gitea commit notification emails from an Outlook 365 folder and either
generates a grouped summary or posts the raw emails to a Claude.ai chat.

Auth: Playwright browser captures the OWA Bearer token once; subsequent
runs reuse auth_state.json (headless).  No app-password or OAuth app needed.

Quick start
-----------
  # Step 1 (once): save Outlook session
  python gitea_summary.py --login

  # Step 2 (once, only needed for --post): save Claude.ai session
  python gitea_summary.py --login-claude

  # Summarise last day (default)
  python gitea_summary.py

  # AI narrative summary via Claude API
  python gitea_summary.py --claude --output summary.md

  # Post raw emails to Claude.ai chat for summarization
  python gitea_summary.py --post

  # List available Outlook folders
  python gitea_summary.py --list-folders
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import time
import urllib.parse
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from pathlib import Path


def _load_dotenv() -> None:
    env_file = Path(__file__).parent / ".env"
    if not env_file.exists():
        return
    with env_file.open(encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key, value = key.strip(), value.strip()
            if key and key not in os.environ:
                os.environ[key] = value


_load_dotenv()

# ── Constants ────────────────────────────────────────────────────────────────

_BROWSER_CANDIDATES = [
    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
    r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/brave-browser",
    "/usr/bin/google-chrome",
]

OUTLOOK_AUTH_FILE  = Path(__file__).parent / "auth_state.json"
OUTLOOK_HOME       = "https://outlook.cloud.microsoft"
CDP_PORT           = 9222
BRAVE_PROFILE_DIR  = Path(__file__).parent / "brave_profile"
OUTLOOK_API       = "https://outlook.office365.com/api/v2.0"
CLAUDE_CHAT_URL   = os.getenv(
    "CLAUDE_CHAT_URL",
    "https://claude.ai/chat/61330a1e-39e4-4717-8704-70904b9f1971",
)
_AUTH_SKIP        = ("/auth/", "/login", "/refresh", "/token", "/oauth")
SENDER_KEYWORD    = "GIT-INFORMER"   # only these emails get summarised + deleted


def _find_browser() -> str | None:
    for p in _BROWSER_CANDIDATES:
        if Path(p).exists():
            print(f"  Browser: {p}")
            return p
    return None


# ── Login: Outlook ───────────────────────────────────────────────────────────

def cmd_login() -> None:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        sys.exit("playwright required:  pip install playwright && playwright install chromium")

    with sync_playwright() as pw:
        exe = _find_browser()
        browser = pw.chromium.launch(headless=False, executable_path=exe)
        ctx = browser.new_context()
        page = ctx.new_page()
        print(f"Opening {OUTLOOK_HOME} — please log in, then press Enter here.")
        page.goto(OUTLOOK_HOME)
        input("Press Enter after you have logged in ... ")
        ctx.storage_state(path=str(OUTLOOK_AUTH_FILE))
        print(f"Outlook session saved -> {OUTLOOK_AUTH_FILE}")
        browser.close()



# ── Capture Outlook Bearer token (headless) ──────────────────────────────────

def _capture_headers() -> dict[str, str]:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        sys.exit("playwright required:  pip install playwright && playwright install chromium")

    if not OUTLOOK_AUTH_FILE.exists():
        sys.exit(f"{OUTLOOK_AUTH_FILE} not found — run with --login first.")

    captures: dict[str, dict] = {}

    def on_request(req):
        if any(s in req.url for s in _AUTH_SKIP):
            return
        auth = req.headers.get("authorization", "")
        if auth.startswith("Bearer "):
            origin = urllib.parse.urlparse(req.url).netloc
            if origin not in captures:
                captures[origin] = dict(req.headers)
                print(f"  Captured token from: {origin}")

    with sync_playwright() as pw:
        exe = _find_browser()
        browser = pw.chromium.launch(headless=True, executable_path=exe)
        ctx = browser.new_context(storage_state=str(OUTLOOK_AUTH_FILE))
        page = ctx.new_page()
        page.on("request", on_request)
        page.goto(OUTLOOK_HOME, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(6000)
        browser.close()

    if not captures:
        sys.exit(
            "No Bearer token captured.\n"
            "Session may have expired — run with --login to refresh."
        )

    return (
        captures.get("outlook.office365.com")
        or captures.get("outlook.cloud.microsoft")
        or next(iter(captures.values()))
    )


# ── Outlook REST helpers ─────────────────────────────────────────────────────

def _get_session(headers: dict[str, str]):
    try:
        import requests as req_mod
    except ImportError:
        sys.exit("requests required:  pip install requests")
    s = req_mod.Session()
    s.headers.update(headers)
    return s


def _find_folder_id(session, folder_name: str) -> str:
    url = f"{OUTLOOK_API}/me/mailFolders?$top=100"
    while url:
        r = session.get(url, timeout=15)
        r.raise_for_status()
        body = r.json()
        for f in body.get("value", []):
            if f.get("DisplayName", "").lower() == folder_name.lower():
                return f["Id"]
        url = body.get("@odata.nextLink")
    sys.exit(f"Folder '{folder_name}' not found in mailbox.")


def _list_folders(session) -> None:
    url = f"{OUTLOOK_API}/me/mailFolders?$top=100"
    print("Available Outlook folders:")
    while url:
        r = session.get(url, timeout=15)
        r.raise_for_status()
        body = r.json()
        for f in body.get("value", []):
            print(f"  {f.get('DisplayName')}  ({f.get('TotalItemCount', '?')} items)")
        url = body.get("@odata.nextLink")


def _fetch_raw_messages(session, folder: str, days: int, unread_only: bool = False) -> list[dict]:
    folder_id = _find_folder_id(session, folder)
    folder_id_enc = urllib.parse.quote(folder_id, safe="")
    since = (datetime.now(timezone.utc) - timedelta(days=days)).strftime(
        "%Y-%m-%dT00:00:00Z"
    )
    clauses = [f"receivedDateTime ge {since}"]
    if unread_only:
        clauses.append("isRead eq false")
    filt = " and ".join(clauses)
    url = (
        f"{OUTLOOK_API}/me/mailFolders/{folder_id_enc}/messages"
        f"?$filter={filt}"
        f"&$select=Id,Subject,ReceivedDateTime,Body,From&$top=50"
    )
    msgs: list[dict] = []
    while url:
        r = session.get(url, timeout=20)
        r.raise_for_status()
        body = r.json()
        msgs.extend(body.get("value", []))
        url = body.get("@odata.nextLink")
        time.sleep(0.1)
    return msgs


# ── Sender filter ─────────────────────────────────────────────────────────────

def _is_target_sender(msg: dict) -> bool:
    name = msg.get("From", {}).get("EmailAddress", {}).get("Name", "")
    return SENDER_KEYWORD in name.upper()


def _filter_by_sender(messages: list[dict]) -> list[dict]:
    """Keep only messages from SENDER_KEYWORD so the summarised set and the
    deleted set are always identical."""
    kept = [m for m in messages if _is_target_sender(m)]
    skipped = len(messages) - len(kept)
    if skipped:
        print(f"  → {skipped} email(s) from other senders ignored")
    return kept


# ── Mark read + delete ──────────────────────────────────────────────────────

def _mark_read_and_delete(session, messages: list[dict]) -> None:
    """Mark each message read, then delete it (moves to Deleted Items).
    Callers MUST pre-filter with _filter_by_sender — this no longer checks
    the sender, so whatever is passed in gets deleted. The mark-read PATCH is
    kept intentionally: if DELETE fails, the message is at least flagged read."""
    if not messages:
        return
    deleted = 0
    for msg in messages:
        msg_id = msg.get("Id") or msg.get("id")
        if not msg_id:
            continue
        msg_id_enc = urllib.parse.quote(msg_id, safe="")
        url = f"{OUTLOOK_API}/me/messages/{msg_id_enc}"
        try:
            session.patch(url, json={"IsRead": True}, timeout=10).raise_for_status()
            session.delete(url, timeout=10).raise_for_status()
            deleted += 1
        except Exception as exc:
            print(f"  [warn] failed to process message {msg_id[:20]}…: {exc}")
        time.sleep(0.05)
    print(f"  → {deleted} email(s) marked read and deleted")


# ── Upstream-sync detection ──────────────────────────────────────────────────

_SYNC_PATTERNS = [
    r"sync\s+(from\s+)?upstream",
    r"upstream\s+sync",
    r"merge\s+(remote[-\s]tracking\s+branch\s+['\"]?)?upstream",
    r"update\s+from\s+upstream",
    r"rebase\s+(on\s+)?upstream",
    r"bump\s+from\s+upstream",
    r"pull\s+from\s+upstream",
    r"follow\s+upstream",
    r"cherry.pick.*upstream",
]
_SYNC_RE = re.compile("|".join(_SYNC_PATTERNS), re.IGNORECASE)


def _is_upstream_sync(subject: str, body: str = "") -> bool:
    return bool(_SYNC_RE.search(subject) or _SYNC_RE.search(body[:800]))


# ── Gitea email parser ───────────────────────────────────────────────────────

_SUBJECT_RE     = re.compile(
    r"^\[(?P<repo>[^\]]+)\]"
    r"\s+(?P<title>.+?)"
    r"(?:\s+\((?P<branch>[^)]+)\))?$"
)
_HASH_RE        = re.compile(r"\b([0-9a-f]{7,40})\b")
_AUTHOR_RE      = re.compile(r"^(?:Author|Committer):\s*(.+)", re.MULTILINE)
_COMMIT_LINE_RE = re.compile(r"^\*\s+([0-9a-f]{6,40})\s+(.+)$", re.MULTILINE)


def _strip_html(html: str) -> str:
    text = re.sub(r"<[^>]+>", " ", html)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _parse_gitea_message(subject: str, body: str, date: str) -> list[dict] | None:
    m = _SUBJECT_RE.match(subject.strip())
    if not m:
        return None

    repo   = m.group("repo").strip()
    title  = m.group("title").strip()
    branch = (m.group("branch") or "").strip()

    commit_lines = _COMMIT_LINE_RE.findall(body)
    if commit_lines:
        author_m = _AUTHOR_RE.search(body)
        return [
            {
                "repo":   repo,
                "branch": branch,
                "hash":   chash[:8],
                "title":  ctitle.strip(),
                "author": author_m.group(1).strip() if author_m else "",
                "date":   date,
            }
            for chash, ctitle in commit_lines
        ]

    hash_m   = _HASH_RE.search(body)
    author_m = _AUTHOR_RE.search(body)
    return [{
        "repo":   repo,
        "branch": branch,
        "hash":   hash_m.group(1)[:8] if hash_m else "?",
        "title":  title,
        "author": author_m.group(1).strip() if author_m else "",
        "date":   date,
    }]


# ── Fetch & filter commits ───────────────────────────────────────────────────

def _fetch_commits(session, folder: str, days: int, verbose: bool = False) -> tuple[list[dict], list[dict]]:
    raw_messages = _fetch_raw_messages(session, folder, days)
    raw_messages = _filter_by_sender(raw_messages)
    print(f"Scanning {len(raw_messages)} email(s) from the last {days} day(s) …")

    commits: list[dict] = []
    skipped_sync = 0
    skipped_non_gitea = 0

    for msg in raw_messages:
        subject    = msg.get("Subject", "")
        body_text  = _strip_html(msg.get("Body", {}).get("Content", ""))
        date       = msg.get("ReceivedDateTime", "")

        parsed = _parse_gitea_message(subject, body_text, date)
        if parsed is None:
            skipped_non_gitea += 1
            if verbose:
                print(f"  [skip/non-gitea] {subject}")
            continue

        for c in parsed:
            if _is_upstream_sync(c["title"], body_text):
                skipped_sync += 1
                if verbose:
                    print(f"  [skip/upstream] {c['repo']} — {c['title']}")
            else:
                commits.append(c)

    print(
        f"  → {len(commits)} commits kept"
        f"  |  {skipped_sync} upstream-sync filtered"
        f"  |  {skipped_non_gitea} non-Gitea emails ignored"
    )
    return commits, raw_messages


# ── Plain-text summary ───────────────────────────────────────────────────────

def _format_summary(commits: list[dict], days: int) -> str:
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    if not commits:
        return (
            f"# Gitea Commit Summary — Last {days} Days\n"
            f"*Generated {now}*\n\n"
            "No qualifying commits found (upstream-sync commits excluded).\n"
        )

    by_repo: dict[str, list[dict]] = defaultdict(list)
    for c in commits:
        by_repo[c["repo"]].append(c)

    lines = [
        f"# Gitea Commit Summary — Last {days} Days",
        f"*Generated {now}*",
        f"*{len(commits)} commits across {len(by_repo)} repositories*",
        "*(Upstream-sync commits excluded)*",
        "",
    ]
    for repo in sorted(by_repo):
        repo_commits = by_repo[repo]
        lines.append(f"## {repo}  ({len(repo_commits)} commit{'s' if len(repo_commits)>1 else ''})")
        for c in repo_commits:
            branch = f" [{c['branch']}]" if c.get("branch") else ""
            author = f"  — {c['author']}" if c.get("author") else ""
            lines.append(f"  - `{c['hash']}`{branch} {c['title']}{author}")
        lines.append("")
    return "\n".join(lines)


# ── Claude API narrative summary ─────────────────────────────────────────────

def _claude_api_summary(commits: list[dict], days: int) -> str:
    try:
        import anthropic
    except ImportError:
        sys.exit("Claude summary requires anthropic:  pip install anthropic")

    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        sys.exit("ANTHROPIC_API_KEY environment variable not set")

    by_repo: dict[str, list[dict]] = defaultdict(list)
    for c in commits:
        by_repo[c["repo"]].append(c)

    commit_list_text = []
    for repo, repo_commits in sorted(by_repo.items()):
        commit_list_text.append(f"Repository: {repo}")
        for c in repo_commits:
            branch = f" (branch: {c['branch']})" if c.get("branch") else ""
            author = f" by {c['author']}" if c.get("author") else ""
            commit_list_text.append(f"  - [{c['hash']}]{branch} {c['title']}{author}")
        commit_list_text.append("")

    prompt = (
        f"以下是過去 {days} 天的 Gitea commit 清單（已排除 upstream 同步的 commits）。\n"
        "請依照 repository 分組，用繁體中文摘要每個 repo 的主要變更，"
        "並在最後提供整體開發進度的簡短總結。\n\n"
        + "\n".join(commit_list_text)
    )

    client = anthropic.Anthropic(api_key=api_key)
    print("Generating AI summary with Claude API …")

    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    ) as stream:
        result = stream.get_final_message()

    header = (
        f"# Gitea Commit Summary — Last {days} Days (AI)\n"
        f"*Generated {datetime.now().strftime('%Y-%m-%d %H:%M')}*\n"
        f"*{len(commits)} commits across {len(by_repo)} repositories — upstream-sync excluded*\n\n"
    )
    return header + result.content[0].text


# ── CDP browser connection (bypasses Cloudflare) ────────────────────────────

def _connect_cdp(pw):
    """Connect to Brave via CDP. Brave must be running with --remote-debugging-port=CDP_PORT.
    Uses http.client directly to bypass corporate proxy."""
    import http.client, json as _json

    try:
        conn = http.client.HTTPConnection("localhost", CDP_PORT, timeout=5)
        conn.request("GET", "/json/version")
        resp = conn.getresponse()
        ws_url = _json.loads(resp.read())["webSocketDebuggerUrl"]
        conn.close()
    except Exception as e:
        sys.exit(
            f"Cannot reach Brave CDP on port {CDP_PORT}: {e}\n"
            f"Launch Brave with:\n"
            f'  & "C:\\Program Files\\BraveSoftware\\Brave-Browser\\Application\\brave.exe"'
            f" --remote-debugging-port={CDP_PORT}"
        )

    print(f"  WebSocket: {ws_url}")
    return pw.chromium.connect_over_cdp(ws_url)


# ── Post raw emails to Claude.ai chat ────────────────────────────────────────

def _clear_clipboard() -> None:
    """Wipe the email contents from the Windows clipboard after pasting."""
    import subprocess
    try:
        subprocess.run(["powershell", "-command", "Set-Clipboard -Value ' '"],
                       check=False, timeout=10)
    except Exception:
        pass



def _format_for_claude_chat(msgs: list[dict], days: int) -> str:
    lines = [
        f"以下是 Outlook GIT 資料夾的未讀信件（共 {len(msgs)} 封）。",
        "請依照對話規則總結這些 Gitea commit 通知。\n",
    ]
    for i, msg in enumerate(msgs, 1):
        subject = msg.get("Subject", "(無主旨)")
        date    = msg.get("ReceivedDateTime", "")[:10]
        sender  = msg.get("From", {}).get("EmailAddress", {}).get("Name", "")
        body    = _strip_html(msg.get("Body", {}).get("Content", ""))
        if len(body) > 2000:
            body = body[:2000] + "\n...(truncated)"
        lines.append(f"=== [{i}] {date} | {subject} | {sender} ===")
        lines.append(body)
        lines.append("")
    return "\n".join(lines)


def _post_to_claude_chat(content: str, chat_url: str) -> bool:
    """Returns True only if the message was confirmed sent. Callers should NOT
    delete source emails unless this returns True."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        sys.exit("playwright required:  pip install playwright && playwright install chromium")

    print(f"Connecting to browser via CDP (port {CDP_PORT}) …")

    with sync_playwright() as pw:
        browser = _connect_cdp(pw)
        ctx = browser.contexts[0] if browser.contexts else browser.new_context()

        # Reuse existing claude.ai tab or open a new one
        page = next((p for p in ctx.pages if "claude.ai" in p.url), None)
        if page is None:
            page = ctx.new_page()

        print(f"  Navigating to chat …")
        page.bring_to_front()
        page.goto(chat_url, wait_until="domcontentloaded", timeout=60_000)
        page.wait_for_timeout(3000)

        # Write content to Windows clipboard via temp file + PowerShell
        import subprocess, tempfile
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".txt", encoding="utf-8", delete=False
        ) as tf:
            tf.write(content)
            tmp = Path(tf.name)
        try:
            subprocess.run(
                ["powershell", "-command",
                 f"Get-Content -Path '{tmp}' -Raw | Set-Clipboard"],
                check=True
            )
            print("  Clipboard ready.")
        finally:
            tmp.unlink(missing_ok=True)

        # Focus editor and paste
        editor = (
            page.query_selector('.ProseMirror[contenteditable="true"]')
            or page.query_selector('[contenteditable="true"]')
        )
        if not editor:
            print("  WARNING: editor not found — message NOT sent.")
            _clear_clipboard()
            return False
        editor.click()
        page.keyboard.press("Control+a")
        page.keyboard.press("Control+v")
        print("  Pasted.")
        page.wait_for_timeout(1500)

        # Submit
        sent = False
        send_btn = (
            page.query_selector('button[aria-label="Send message"]')
            or page.query_selector('button[data-testid="send-button"]')
        )
        if send_btn:
            send_btn.click()
            print("  Message sent.")
            sent = True
        else:
            editor = page.query_selector('[contenteditable="true"]')
            if editor:
                editor.press("Enter")
                print("  Message sent via Enter.")
                sent = True
            else:
                print("  WARNING: could not find send button or editor — message NOT sent.")

        _clear_clipboard()
        if sent:
            print("Done — check the browser for Claude's response.")
        return sent


# ── Entry point ──────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Summarise Gitea commit emails from an Outlook 365 folder.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--login", action="store_true",
        help="Open browser to log in to Outlook and save session")
    parser.add_argument("--test-browser", action="store_true",
        help="Test Brave CDP connection and open the target chat URL, then exit")
    parser.add_argument("--list-folders", action="store_true",
        help="Print all Outlook folders and exit")
    parser.add_argument("--post", action="store_true",
        help="Post raw emails to Claude.ai chat for summarization (Brave must run with --remote-debugging-port=9222)")
    parser.add_argument("--chat", default=CLAUDE_CHAT_URL,
        help="Claude.ai chat URL for --post (or set CLAUDE_CHAT_URL)")
    parser.add_argument("--folder", default="GIT",
        help="Outlook folder name (default: GIT)")
    parser.add_argument("--days", type=int, default=7,
        help="Days to look back for --claude/--output (default: 7)")
    parser.add_argument("--output", metavar="FILE",
        help="Write summary to file instead of stdout (not used with --post)")
    parser.add_argument("--claude", action="store_true",
        help="Generate narrative summary via Claude API (requires ANTHROPIC_API_KEY)")
    parser.add_argument("--verbose", action="store_true",
        help="Show which emails are skipped and why")
    args = parser.parse_args()

    if args.login:
        cmd_login()
        return

    if args.test_browser:
        from playwright.sync_api import sync_playwright
        print(f"Testing Brave CDP connection on port {CDP_PORT} …")
        with sync_playwright() as pw:
            browser = _connect_cdp(pw)
            ctx = browser.contexts[0] if browser.contexts else browser.new_context()
            page = next((p for p in ctx.pages if "claude.ai" in p.url), None)
            if page is None:
                page = ctx.new_page()
            page.bring_to_front()

            # Step 1: check claude.ai login status
            print("  Checking claude.ai login …")
            page.goto("https://claude.ai", wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(2000)
            current = page.url
            title = page.title()
            if "login" in current or "sign" in current.lower():
                print(f"  ✗ Not logged in — redirected to: {current}")
                print("    Please log in to Claude.ai in the Brave window, then re-run.")
                input("Press Enter to close … ")
                return
            print(f"  ✓ Logged in  (title: {title})")

            # Step 2: navigate to target chat
            print(f"  Navigating to target chat …")
            page.goto(args.chat, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(2000)
            chat_title = page.title()
            if "login" in page.url or "sign" in page.url.lower():
                print(f"  ✗ Chat redirected to login: {page.url}")
            else:
                print(f"  ✓ Chat reachable  (title: {chat_title})")
                print(f"  ✓ All checks passed — ready to run normally.")

            input("Press Enter to close … ")
        return

    print("Capturing Outlook auth token …")
    headers = _capture_headers()
    session = _get_session(headers)

    if args.list_folders:
        _list_folders(session)
        return

    if args.post or not (args.claude or args.output or args.list_folders):
        # Default behaviour: fetch unread emails and post to Claude.ai chat
        print(f"Fetching unread emails from '{args.folder}' …")
        msgs = _fetch_raw_messages(session, args.folder, args.days, unread_only=True)
        print(f"  → {len(msgs)} emails fetched")
        msgs = _filter_by_sender(msgs)
        if not msgs:
            print("No matching emails found.")
            return
        content = _format_for_claude_chat(msgs, args.days)
        sent = _post_to_claude_chat(content, args.chat)
        if sent:
            _mark_read_and_delete(session, msgs)
        else:
            print("  Send not confirmed — emails kept (not deleted).")
        return

    # --claude / --output: fetch commits, format/summarize, print or save
    commits, raw_messages = _fetch_commits(session, args.folder, args.days, verbose=args.verbose)

    if args.claude:
        summary = _claude_api_summary(commits, args.days)
    else:
        summary = _format_summary(commits, args.days)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(summary)
        print(f"\nSummary written to {args.output}")
    else:
        print()
        print(summary)

    _mark_read_and_delete(session, raw_messages)


if __name__ == "__main__":
    main()
