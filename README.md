# AIfred.mcp

> *Every superhero needs an Alfred. Yours handles the CRM.*

Built by **Fred** — Solution Consultant @ ServiceNow

Connects Claude Desktop directly to your CRM, calendar, email and Teams — using your existing browser session. No Azure app registration. No stored credentials. No CRM admin work ever again.

Two flavours — one installer:

| Variant | Who | Install folder |
|---------|-----|----------------|
| **Alfred SC** | SC / SSC / Manager | `~/Documents/alfred.sc` |
| **Alfred Sales** | Account Executive | `~/Documents/alfred.sales` |

---

## What it does

### Alfred SC (Solution Consulting)

| Source | Capabilities |
|--------|-------------|
| **Dynamics 365** | List opportunities, create/update/complete engagements, hygiene sweep, Tech Win assessment, delete cancelled engagements |
| **Outlook Calendar** | Show calendar by date range, search meetings |
| **Outlook Email** | Search emails, list inbox/sent, full body, filter unread |
| **Teams** | Get meeting transcripts, post to channels, read chats |
| **Account Insights** | License utilization, renewal dates, upsell/cross-sell detection |

### Alfred Sales (Account Executive)

| Source | Capabilities |
|--------|-------------|
| **Dynamics 365** | Create & update opportunities, assign SC, search accounts/users, add notes |

---

## Requirements

- macOS
- [Claude Desktop](https://claude.ai/download)
- Google Chrome
- Node.js — **installed automatically if missing**

---

## Setup

1. Go to **https://github.com/h22fred/Alfred.mcp** and download `Setup.command`
2. Open Terminal and run:
```bash
bash ~/Downloads/Setup.command
```

> If a popup appears asking to install Command Line Tools, click **Install**, wait for it to finish, then run the same command again.
> If Homebrew is not installed, it will ask for your **Mac login password** once — this is normal.

The installer asks:
- **SC or Sales?** — determines which Alfred is installed
- **Dynamics company name** — your CRM URL (e.g. `servicenow`)
- **Teams webhook** — for automated notifications (optional)
- **SC role** (SC only) — SC / SSC / Manager
- **Engagement types** (SC only) — which types you use
- **Automated jobs** (SC only) — Monday hygiene sweep + Friday meeting review

---

## Every session

1. Double-click **Alfred.app** on your Desktop
2. **First time only:** log into Dynamics, Outlook and Teams (SSO) — Alfred remembers you after that
3. Alfred automatically opens Claude Desktop — you're ready

---

## Example prompts

### SC / SSC / Manager
```
List my open opportunities over $100K
Run hygiene sweep and post to Teams
Which accounts are missing a Technical Win?
Assess the Tech Win for SITA Brown Field
Create a Discovery engagement for Givaudan from my Tuesday meeting
Mark the SITA Tech Win as complete
Show my calendar this week
Get the transcript from my Givaudan demo last Thursday
Detect post-meeting engagements from this week
Delete the cancelled Demo engagement on PMI
```

### Sales AE
```
Create an opportunity for Givaudan — New ITSM, close December 2026
Find the account ID for Roche
Search for Fredrik to get his SC GUID and assign him
Update the PMI opportunity close date to March 2027
Show my open opportunities
Add a note to the SITA opportunity: had intro call, next step is discovery
```

---

## How it works

Alfred.app launches Chrome with `--remote-debugging-port=9222` using a dedicated profile (`~/.alfred-profile`). The MCP server extracts session cookies and Bearer tokens via CDP WebSocket — no credentials stored, no Azure registration needed.

Auth flow:
1. **Dynamics:** reads `CrmOwinAuthC1/C2` cookies via `Network.getCookies`
2. **Outlook/Graph:** reads Bearer token from MSAL cache in page storage
3. All tokens cached in-memory for the session duration

---

## Automated jobs (SC only)

| When | What |
|------|------|
| Monday 9:30am (configurable) | CRM hygiene sweep — flags missing engagements, posts to Teams |
| Friday 2:00pm (configurable) | Meeting review — matches this week's meetings to open opps |

To run manually:
```bash
node scripts/hygiene-sweep.mjs
node scripts/post-meeting-sweep.mjs
```

Config lives in `~/.alfred-config.json`.

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| Alfred not running | Double-click Alfred.app on Desktop |
| Not logged into Dynamics | Log into Dynamics in the Alfred Chrome window |
| 401 from Dynamics | Session expired — re-login in Alfred window |
| Teams not posting | Re-run setup to reconfigure webhook |
| Node.js not found | Re-run setup — it installs automatically |

---

*Questions? Open a [Discussion](https://github.com/h22fred/Alfred.mcp/discussions) or ping Fred on Teams.*
