# gitea-summary

Reads Gitea commit notification emails from an Outlook 365 IMAP folder and generates a grouped summary, skipping upstream-sync commits.

## Features

- Connects to Outlook 365 via IMAP
- Parses Gitea push notification email format
- Filters out upstream-sync commits automatically
- Groups commits by repository
- Optional AI narrative summary in Traditional Chinese via Claude API
- Supports both app-password and OAuth2 (device flow) authentication

## Requirements

```bash
# Core (no extra deps for password auth)
pip install msal        # only for --auth oauth
pip install anthropic   # only for --claude AI summary
```

## Usage

```bash
# List IMAP folders to confirm folder name
python gitea_summary.py --email you@company.com --list-folders

# Summarise last 7 days
python gitea_summary.py --email you@company.com --days 7

# Save to file
python gitea_summary.py --email you@company.com --days 14 --output summary.md

# AI narrative summary in Traditional Chinese
python gitea_summary.py --email you@company.com --days 7 --claude

# OAuth2 device flow
python gitea_summary.py --email you@company.com --auth oauth \
    --client-id <azure-app-id> --tenant <tenant-id>
```

## Auth

| Method | Notes |
|--------|-------|
| App password | Outlook → Settings → Security → App passwords |
| OAuth2 | Requires Azure app with `IMAP.AccessAsUser.All` permission |

## Upstream-sync filter

Commits matching these patterns are excluded:
- `sync from upstream` / `upstream sync`
- `merge upstream` / `merge remote-tracking branch 'upstream/...'`
- `update/rebase/pull/follow from upstream`
