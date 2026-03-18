# AlFred.mcp

> *Every superhero needs an Alfred. Yours logs the Engagements.*

Built by **Fredrik Holmström**, Solution Consultant @ ServiceNow

Connects Claude Desktop directly to your CRM, calendar, email and Teams — using your existing browser session. No Azure app registration. No stored credentials. No CRM admin work ever again.

---

## What it does

| Source | Capabilities |
|--------|-------------|
| **Dynamics 365** | List opportunities, create/update/complete engagements, CRM hygiene sweep |
| **Outlook Calendar** | Show calendar by date range, search meetings |
| **Outlook Email** | Search emails, list inbox/sent, full body, filter unread |
| **Teams** | Get meeting transcripts, post to channels, read chats |
| **Account Insights** | License utilization, renewal dates, upsell/cross-sell detection |

---

## Requirements

- macOS
- [Claude Desktop](https://claude.ai/download)
- Google Chrome
- Node.js (`brew install node`)

---

## Setup

Download [Setup.command](https://github.com/h22fred/alfred.mcp/raw/main/Setup.command) and double-click it. That's it.

> **First time only:** macOS may block it. Right-click → **Open** → **Open** to bypass.

The installer:
- Runs `npm install` and `tsc`
- Registers AlFred in `~/Library/Application Support/Claude/claude_desktop_config.json`
- Creates **Alfred.app** on your Desktop
- Prompts for your Teams incoming webhook URL
- Installs Monday 9:30am hygiene sweep + Friday 2pm meeting review cron jobs

---

## Every session

1. Double-click **Alfred.app** on your Desktop
2. **First time only:** log into Dynamics, Outlook and Teams (ServiceNow SSO) — Alfred remembers you after that
3. Leave Alfred running in the Dock
4. Open Claude Desktop and ask anything

---

## Example prompts

```
List my open opportunities over $100K
Show SITA opportunities with OPTY numbers
Get full context on this account

Run hygiene sweep and post to Teams
Which accounts are missing a Technical Win?

Create a Discovery engagement for SITA Brown Field
Mark the Givaudan Tech Win as complete

Show my calendar this week
Find all meetings with "PMI" in the subject
Search emails for "budget approval"
Get the transcript from my SITA demo last week

Detect post-meeting engagements from this week
```

---

## How it works

Alfred.app launches Chrome with `--remote-debugging-port=9222` using a dedicated profile (`~/.alfred-profile`). The MCP server extracts session cookies and Bearer tokens via raw CDP WebSocket — no credentials stored, no Azure registration needed.

Auth flow:
1. **Dynamics:** reads `CrmOwinAuthC1/C2` cookies via `Network.getCookies`
2. **Outlook/Graph:** reads Bearer token from MSAL cache in the page's storage
3. All tokens cached in-memory for the session duration

---

## Automated jobs

| When | What |
|------|------|
| Monday 9:30am | CRM hygiene sweep — flags missing engagements, posts to Teams |
| Friday 2:00pm | Meeting review — lists this week's customer meetings + missing engagements per matched opp |

To run manually:
```bash
node scripts/hygiene-sweep.mjs
node scripts/post-meeting-sweep.mjs
```

Config lives in `~/.alfred-config.json`:
```json
{ "teamsWebhook": "https://your-webhook-url" }
```

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| Alfred not running | Claude calls `open_chrome_debug` automatically |
| Not logged into Dynamics | Log into `servicenow.crm.dynamics.com` in Alfred window |
| 401 from Dynamics | Session expired — re-login in Alfred |
| Teams not posting | Run `setup.sh` again to reconfigure webhook |

---

*Questions? Ping Fredrik on Teams or open an issue.*
