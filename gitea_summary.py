#!/usr/bin/env python3
"""
gitea_summary.py

Reads Gitea commit notification emails from an Outlook 365 IMAP folder and
generates a grouped summary, skipping upstream-sync commits.

Auth modes
----------
  --auth password   email + app-password (no extra deps)
  --auth oauth      OAuth2 device-flow  (requires: pip install msal)

Optional AI summary
-------------------
  --claude          Use Claude API for a narrative summary
                    (requires: pip install anthropic  +  ANTHROPIC_API_KEY)

Quick start
-----------
  # List available IMAP folders first (to confirm folder name)
  python gitea_summary.py --email you@company.com --list-folders

  # Summarise last 7 days
  python gitea_summary.py --email you@company.com --days 7

  # Save to file
  python gitea_summary.py --email you@company.com --days 14 --output summary.md
"""

from __future__ import annotations

import email
import getpass
import imaplib
import os
import re
import sys
import argparse
from collections import defaultdict
from datetime import datetime, timedelta
from email.header import decode_header
from textwrap import indent


# ── Upstream-sync commit detection ──────────────────────────────────────────

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


# ── Email / MIME helpers ─────────────────────────────────────────────────────

def _decode_header_str(raw) -> str:
    if not raw:
        return ""
    parts = decode_header(raw)
    out = []
    for fragment, enc in parts:
        if isinstance(fragment, bytes):
            out.append(fragment.decode(enc or "utf-8", errors="replace"))
        else:
            out.append(str(fragment))
    return "".join(out)


def _get_text_body(msg: email.message.Message) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    return payload.decode(charset, errors="replace")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
    return ""


# ── Gitea email parser ───────────────────────────────────────────────────────
#
# Gitea subject formats:
#   [owner/repo] Commit message (branch)
#   [owner/repo] N new commits pushed to branch
#   [owner/repo] [branch] PR title  (PR notification — we skip these)

_SUBJECT_RE = re.compile(
    r"^\[(?P<repo>[^\]]+)\]"          # [owner/repo]
    r"\s+(?P<title>.+?)"              # commit title or push summary
    r"(?:\s+\((?P<branch>[^)]+)\))?$" # optional (branch)
)
_HASH_RE   = re.compile(r"\b([0-9a-f]{7,40})\b")
_AUTHOR_RE = re.compile(r"^(?:Author|Committer):\s*(.+)", re.MULTILINE)
_DATE_RE   = re.compile(r"^Date:\s*(.+)", re.MULTILINE)
_MULTI_COMMIT_RE = re.compile(r"(\d+)\s+new commits?", re.IGNORECASE)

# Lines that look like   * abc1234 commit message
_COMMIT_LINE_RE = re.compile(r"^\*\s+([0-9a-f]{6,40})\s+(.+)$", re.MULTILINE)


def _parse_gitea_email(msg: email.message.Message) -> list[dict] | None:
    """
    Return a list of commit dicts extracted from a Gitea notification email,
    or None if the email doesn't look like a Gitea push notification.
    """
    subject = _decode_header_str(msg.get("Subject", ""))
    body    = _get_text_body(msg)
    date    = msg.get("Date", "")

    m = _SUBJECT_RE.match(subject.strip())
    if not m:
        return None

    repo   = m.group("repo").strip()
    title  = m.group("title").strip()
    branch = (m.group("branch") or "").strip()

    # Multi-commit push: Gitea lists commits in the body as "* <hash> <msg>"
    commit_lines = _COMMIT_LINE_RE.findall(body)
    if commit_lines:
        commits = []
        for chash, ctitle in commit_lines:
            author_m = _AUTHOR_RE.search(body)
            commits.append({
                "repo":    repo,
                "branch":  branch,
                "hash":    chash[:8],
                "title":   ctitle.strip(),
                "author":  author_m.group(1).strip() if author_m else "",
                "date":    date,
                "subject": subject,
                "body":    body,
            })
        return commits

    # Single-commit push
    hash_m   = _HASH_RE.search(body)
    author_m = _AUTHOR_RE.search(body)
    return [{
        "repo":    repo,
        "branch":  branch,
        "hash":    hash_m.group(1)[:8] if hash_m else "?",
        "title":   title,
        "author":  author_m.group(1).strip() if author_m else "",
        "date":    date,
        "subject": subject,
        "body":    body,
    }]


# ── IMAP connection ──────────────────────────────────────────────────────────

IMAP_HOST = "outlook.office365.com"
IMAP_PORT = 993


def _connect_password(email_addr: str, password: str) -> imaplib.IMAP4_SSL:
    conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    conn.login(email_addr, password)
    return conn


def _connect_oauth2(email_addr: str, tenant_id: str, client_id: str) -> imaplib.IMAP4_SSL:
    try:
        import msal
    except ImportError:
        sys.exit("OAuth2 requires msal:  pip install msal")

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )
    scopes = ["https://outlook.office365.com/IMAP.AccessAsUser.All"]

    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        sys.exit(f"Failed to start device flow: {flow}")

    print(f"\n{flow['message']}\n")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        sys.exit(f"OAuth2 failed: {result.get('error_description', result)}")

    token = result["access_token"]
    auth_str = f"user={email_addr}\x01auth=Bearer {token}\x01\x01"
    conn = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    conn.authenticate("XOAUTH2", lambda _: auth_str.encode())
    return conn


# ── Fetch & filter ───────────────────────────────────────────────────────────

def _list_folders(conn: imaplib.IMAP4_SSL) -> None:
    _, raw_list = conn.list()
    print("Available IMAP folders:")
    for item in raw_list or []:
        print(" ", _decode_header_str(item) if isinstance(item, bytes) else item)


def _fetch_commits(
    conn: imaplib.IMAP4_SSL,
    folder: str,
    days: int,
    verbose: bool = False,
) -> list[dict]:
    # Try selecting with and without quotes
    for name in (f'"{folder}"', folder):
        status, _ = conn.select(name)
        if status == "OK":
            break
    else:
        print(f"\nERROR: Cannot open folder '{folder}'.")
        _list_folders(conn)
        sys.exit(1)

    since = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
    _, data = conn.search(None, f'SINCE "{since}"')
    ids = (data[0] or b"").split()

    if not ids:
        print(f"No emails found in '{folder}' since {since}.")
        return []

    print(f"Scanning {len(ids)} email(s) from the last {days} day(s) …")

    commits: list[dict] = []
    skipped_sync = 0
    skipped_non_gitea = 0

    for num in ids:
        _, msg_data = conn.fetch(num, "(RFC822)")
        if not msg_data or not msg_data[0]:
            continue
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)

        parsed = _parse_gitea_email(msg)
        if parsed is None:
            skipped_non_gitea += 1
            if verbose:
                subj = _decode_header_str(msg.get("Subject", ""))
                print(f"  [skip/non-gitea] {subj}")
            continue

        for c in parsed:
            if _is_upstream_sync(c["title"], c.get("body", "")):
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
    return commits


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


# ── Claude AI summary (optional) ─────────────────────────────────────────────

def _claude_summary(commits: list[dict], days: int) -> str:
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

    prompt = f"""以下是過去 {days} 天的 Gitea commit 清單（已排除 upstream 同步的 commits）。
請依照 repository 分組，用繁體中文摘要每個 repo 的主要變更，並在最後提供整體開發進度的簡短總結。

{chr(10).join(commit_list_text)}"""

    client = anthropic.Anthropic(api_key=api_key)
    print("Generating AI summary with Claude …")

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


# ── Entry point ──────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Summarise Gitea commit emails from an Outlook 365 IMAP folder.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--email",
        default=os.getenv("OUTLOOK_EMAIL"),
        help="Your Outlook 365 address (or set OUTLOOK_EMAIL)")
    parser.add_argument("--password",
        default=os.getenv("OUTLOOK_PASSWORD"),
        help="App-password (or set OUTLOOK_PASSWORD; prompted if omitted)")
    parser.add_argument("--auth",
        choices=["password", "oauth"], default="password",
        help="Auth method (default: password)")
    parser.add_argument("--tenant",
        default=os.getenv("AZURE_TENANT_ID", "common"),
        help="Azure tenant ID for OAuth2 (or AZURE_TENANT_ID, default: common)")
    parser.add_argument("--client-id",
        default=os.getenv("AZURE_CLIENT_ID"),
        help="Azure app client ID for OAuth2 (or AZURE_CLIENT_ID)")
    parser.add_argument("--folder", default="GIT",
        help="IMAP folder name (default: GIT)")
    parser.add_argument("--days", type=int, default=7,
        help="Days to look back (default: 7)")
    parser.add_argument("--output", metavar="FILE",
        help="Write summary to file instead of stdout")
    parser.add_argument("--claude", action="store_true",
        help="Generate narrative summary via Claude API (requires ANTHROPIC_API_KEY)")
    parser.add_argument("--list-folders", action="store_true",
        help="Print all IMAP folders and exit")
    parser.add_argument("--verbose", action="store_true",
        help="Show which emails are skipped and why")
    args = parser.parse_args()

    if not args.email:
        sys.exit("Email required: use --email or set OUTLOOK_EMAIL")

    # ── Connect ────────────────────────────────────────────────────────────
    print(f"Connecting to {IMAP_HOST} …")
    if args.auth == "oauth":
        if not args.client_id:
            sys.exit("OAuth2 requires --client-id or AZURE_CLIENT_ID")
        conn = _connect_oauth2(args.email, args.tenant, args.client_id)
    else:
        pwd = args.password
        if not pwd:
            pwd = getpass.getpass("App-password: ")
        conn = _connect_password(args.email, pwd)
    print("Connected.\n")

    if args.list_folders:
        _list_folders(conn)
        conn.logout()
        return

    # ── Fetch & summarise ─────────────────────────────────────────────────
    commits = _fetch_commits(conn, args.folder, args.days, verbose=args.verbose)
    conn.logout()

    if args.claude:
        summary = _claude_summary(commits, args.days)
    else:
        summary = _format_summary(commits, args.days)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(summary)
        print(f"\nSummary written to {args.output}")
    else:
        print()
        print(summary)


if __name__ == "__main__":
    main()
