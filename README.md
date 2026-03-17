# SC Engagement MCP

> Built by **Fredrik Holmström**, Solution Consultant @ ServiceNow

Connects Claude Desktop directly to your CRM, calendar, email and Teams using your existing browser session. No Azure app registration. No stored credentials.

---

## What it does

| Source | Capabilities |
|--------|-------------|
| **Dynamics 365** | List opportunities (with OPTY numbers), create/update/complete engagements, CRM hygiene sweep |
| **Outlook Calendar** | Show calendar by date range, search meetings by keyword |
| **Outlook Email** | Search emails, list inbox/sent, filter unread |
| **Teams** | Get meeting transcripts, post to channels, read chats |
| **Account Insights** | License utilization, renewal dates, upsell/cross-sell/new logo detection |

---

## Requirements

- macOS
- [Claude Desktop](https://claude.ai/download)
- Google Chrome
- Node.js (`brew install node`)

---

## Setup

```bash
# 1. Clone the repo
git clone https://github.com/h22fred/sc-engagement-mcp.git
cd sc-engagement-mcp

# 2. Run the installer
# Double-click Setup.command in Finder, or:
bash setup.sh
```

The installer:
- Runs `npm install` and `tsc`
- Registers the MCP server in `~/Library/Application Support/Claude/claude_desktop_config.json`
- Creates **ChromeDebug.app** on your Desktop
- Prompts for your Teams incoming webhook URL
- Installs a Monday 9:30am cron job for automated hygiene sweep

---

## Every session

1. Double-click **ChromeDebug.app** on your Desktop
2. Log into Dynamics, Outlook and Teams (ServiceNow SSO)
3. Open Claude Desktop

---

## Example prompts

```
List my open opportunities over $100K
Show SITA opportunities with OPTY numbers
Get full context on this opportunity — what does the customer already own?

Run hygiene sweep and post to Teams
Which accounts are missing a Technical Win?

Create a Discovery engagement for SITA Brown Field
Mark the Givaudan Tech Win as complete

Show my calendar this week
Find all meetings with "PMI" in the subject next 2 weeks
Search emails for "budget approval"

Get the transcript from my SITA demo last week
```

---

## How it works

ChromeDebug.app launches Chrome with `--remote-debugging-port=9222`. The MCP server extracts session cookies and Bearer tokens from the browser via raw CDP WebSocket — no credentials stored, no Azure registration needed.

Auth flow:
1. Dynamics: reads `CrmOwinAuthC1/C2` cookies via `Network.getCookies` CDP command
2. Outlook/Graph: intercepts Bearer token from outgoing requests via Playwright `page.route()`
3. All tokens cached in-memory for the session duration

---

## Monday hygiene sweep

Cron fires at 9:30am every Monday. Requires ChromeDebug to be running and logged in. If auth fails, posts a Teams reminder instead. To run manually:

```bash
node scripts/hygiene-sweep.mjs
```

Or in Claude Desktop: *"Run hygiene sweep and post to Teams"*

Teams webhook config lives in `~/.sc-engagement-config.json`:
```json
{ "teamsWebhook": "https://your-webhook-url" }
```

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| ChromeDebug not running | Claude calls `open_chrome_debug` automatically |
| Not logged into Dynamics | Log into `servicenow.crm.dynamics.com` in ChromeDebug window |
| 401 from Dynamics | Session expired — re-login in ChromeDebug |
| Multiple Chrome windows | Only use ChromeDebug.app, not regular Chrome |
| Teams not posting | Run `setup.sh` again to reconfigure webhook |

---

*Questions? Ping Fredrik on Teams or open an issue.*
