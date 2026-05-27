# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Tool

```bash
# List available IMAP folders (useful for finding the right folder name)
python gitea_summary.py --email you@company.com --list-folders

# Summarize commits from the last 7 days
python gitea_summary.py --email you@company.com --days 7

# Generate AI narrative summary (Traditional Chinese) and save to file
python gitea_summary.py --email you@company.com --days 14 --claude --output summary.md

# Use OAuth2 device-flow instead of password
python gitea_summary.py --email you@company.com --auth oauth \
    --client-id <azure-app-id> --tenant <tenant-id>
```

## Dependencies

The core tool requires no external packages for password-based auth. Optional:

```bash
pip install msal        # required only for --auth oauth
pip install anthropic   # required only for --claude
```

## Configuration

Credentials can be set in a `.env` file (see `.env.example`) or via env vars:

- `OUTLOOK_EMAIL` / `OUTLOOK_PASSWORD` — avoids typing `--email`/`--password` each run
- `ANTHROPIC_API_KEY` — needed for `--claude` summaries
- `AZURE_TENANT_ID` / `AZURE_CLIENT_ID` — needed for `--auth oauth`

## Architecture

Single-file CLI (`gitea_summary.py`, ~460 lines). The data flow is:

```
main()
 └─ _connect_password() / _connect_oauth2()   # IMAP to outlook.office365.com:993
 └─ _fetch_commits()                           # searches folder, calls _parse_gitea_email() per message
 └─ _format_summary() / _claude_summary()     # grouped Markdown or Claude API narrative
```

**Key parsing logic** — `_parse_gitea_email()` extracts repo, branch, author, hash, and title from Gitea push notification emails. Subject format expected: `[owner/repo] commit_title (branch)`. Multi-commit pushes are handled via the text body.

**Upstream-sync filter** — `_is_upstream_sync()` uses regex patterns to drop merge/rebase-from-upstream commits from the output. Patterns live at the top of that function.

**Claude summary** — `_claude_summary()` groups commits by repo and sends them to `claude-sonnet-4-6` with a Traditional Chinese (繁體中文) prompt. Requires `ANTHROPIC_API_KEY`.

**IMAP folder** — defaults to `"GIT"`. Override with `--folder`. Use `--list-folders` to discover folder names on the mailbox.
