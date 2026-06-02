# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Tool

```bash
# Step 1 (once): save Outlook session
python gitea_summary.py --login

# Step 2 (once, only needed for --post): save Claude.ai session
python gitea_summary.py --login-claude

# Default: post last 7 days of unread GIT-INFORMER emails to Claude.ai chat,
# then mark them read and delete them (only on confirmed send)
python gitea_summary.py

# Generate AI narrative summary (Traditional Chinese) via Claude API
python gitea_summary.py --claude --output summary.md

# Post raw emails to Claude.ai chat (requires --login-claude; Brave/Chrome must be running)
python gitea_summary.py --post

# List available Outlook folders
python gitea_summary.py --list-folders
```

## Dependencies

```bash
pip install playwright requests       # required
playwright install chromium           # only needed if no system Brave/Chrome
pip install anthropic                 # required only for --claude
```

## Configuration

Set in `.env` or as environment variables:

- `ANTHROPIC_API_KEY` — needed for `--claude` summaries
- `CLAUDE_CHAT_URL` — override default Claude.ai chat URL for `--post`

Session files:
- `auth_state.json` — Outlook session (created by `--login`); re-run `--login` if expired
- `claude_auth_state.json` — Claude.ai session (created by `--login-claude`); needed for `--post`

Claude.ai posting uses **CDP** (Chrome DevTools Protocol) to connect to the real Brave/Chrome
browser instead of a fresh context, bypassing Cloudflare bot detection.
The script auto-launches Brave with `--remote-debugging-port=9222` if the browser is not running.

## Architecture

Single-file CLI (`gitea_summary.py`). Two auth phases, two operating modes:

```
Login phases (once each):
  --login        → headless=False, outlook.cloud.microsoft → auth_state.json
  --login-claude → headless=False, claude.ai              → claude_auth_state.json

Normal run (summarize mode):
  headless=True + auth_state.json → capture OWA Bearer token
  requests + Bearer → Outlook REST API v2.0 → fetch messages
  _parse_gitea_message() + _is_upstream_sync() → filtered commits
  _format_summary() or _claude_api_summary() → print / save

--post mode (raw email → Claude chat):
  Same Outlook fetch (raw messages, no commit parsing)
  headless=False + claude_auth_state.json → open Claude chat
  clipboard + execCommand('paste') → fill ProseMirror editor → submit
```

**Key parsing logic** — `_parse_gitea_message()` extracts repo, branch, author, hash, and title from Gitea push notification email subjects. Format expected: `[owner/repo] commit_title (branch)`. Multi-commit pushes are handled via `* <hash> <title>` lines in the body (HTML stripped before parsing).

**Upstream-sync filter** — `_is_upstream_sync()` drops merge/rebase-from-upstream commits. Patterns live at the top of that function.

**Sender filter + cleanup** — `_filter_by_sender()` keeps only emails whose `From.EmailAddress.Name` contains `SENDER_KEYWORD` (`"GIT-INFORMER"`), applied right after fetch so the summarised set always equals the deleted set. After a run, `_mark_read_and_delete()` marks each message read then DELETEs it (moves to Deleted Items). **Critical invariant:** in `--post` mode, deletion only happens when `_post_to_claude_chat()` returns `True` (confirmed send) — never delete emails that weren't successfully posted.

**Outlook REST API** — `https://outlook.office365.com/api/v2.0`, PascalCase fields (`Subject`, `Body.Content`, `ReceivedDateTime`). Folder IDs containing `/` are URL-encoded before use.

**Claude chat posting** — uses clipboard API + `execCommand('paste')` to fill the ProseMirror contenteditable editor, then clicks the send button (falls back to Enter key).
