# AIfred.mcp

> *Every superhero needs an Alfred. Yours handles the CRM.*

Built by **Fred** — Solution Consultant @ ServiceNow

<a href="https://www.buymeacoffee.com/h22fred"><img src="https://img.buymeacoffee.com/button-api/?text=Buy me a coffee&emoji=&slug=h22fred&button_colour=FF5F5F&font_colour=ffffff&font_family=Cookie&outline_colour=000000&coffee_colour=FFDD00" /></a>

[![LinkedIn](https://img.shields.io/badge/LinkedIn-fredholmstrom-0A66C2?logo=linkedin&logoColor=white)](https://www.linkedin.com/in/fredholmstrom/)
[![GitHub followers](https://img.shields.io/github/followers/h22fred?label=Follow%20on%20GitHub&style=social)](https://github.com/h22fred)
[![X](https://img.shields.io/badge/X-h22fred-000000?logo=x&logoColor=white)](http://www.x.com/h22fred)

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

**Option A — Download and run** _(recommended)_
1. **[⬇️ Download Setup.command](https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup.command)**
2. Open Terminal and run:
```bash
bash ~/Downloads/Setup.command
```

**Option B — One-liner** _(for the terminal-comfortable)_
```bash
curl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup.command | bash
```

> If a popup appears asking to install Command Line Tools, click **Install**, wait for it to finish, then run the same command again.

The installer asks:
- **SC or Sales?** — determines which Alfred is installed
- **Dynamics company name** — your CRM URL (e.g. `servicenow`)
- **Teams webhook** — for automated notifications (optional)
- **SC role** (SC only) — SC / SSC / Manager
- **Engagement types** — which milestones you track (SC types or AE milestones depending on role)
- **Automated jobs** — Monday hygiene sweep + Friday meeting review

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

## Automated jobs

Alfred sets up two recurring cron jobs during installation (both optional):

| When | What |
|------|------|
| **Monday 9:30am** | CRM hygiene sweep — checks all open opportunities for missing milestones, posts a summary card to Teams |
| **Friday 2:00pm** | Meeting review — scans this week's calendar, matches meetings to open opportunities, suggests engagements to log |

Both jobs run silently in the background. If Alfred isn't running or your Dynamics session has expired, they post a Teams reminder instead of failing silently.

To run manually at any time:
```bash
node ~/Documents/alfred.sc/scripts/hygiene-sweep.mjs
node ~/Documents/alfred.sc/scripts/post-meeting-sweep.mjs
```

To change the schedule, edit your crontab:
```bash
crontab -e
```

Config lives in `~/.alfred-config.json`.

---

## Teams webhook

The Teams webhook is optional but recommended — it's what allows Alfred to post hygiene reports and reminders directly to a Teams channel.

**How to set it up:**
1. In Teams, go to the channel where you want Alfred to post
2. Click **···** → **Connectors** (or **Manage channel** → **Connectors**)
3. Add an **Incoming Webhook**, give it a name (e.g. *Alfred*) and copy the URL
4. Paste it when the installer asks, or re-run setup to add/update it

**What Alfred posts:**
- Monday hygiene sweep results — list of opportunities with missing milestones
- Friday meeting review — suggested engagements to log
- Auth reminders — if your Dynamics session has expired before an automated job runs

The webhook URL is stored in `~/.alfred-config.json` on your machine only and is never sent anywhere except to post cards to your Teams channel.

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

---

## Security

### Security assessment

A full security review was conducted on this codebase covering credential handling, input validation, API calls, network exposure, and the install scripts. No vulnerabilities were found in the core data handling or CRM integration layer. Installer integrity can be verified via the checksum below.

### Data handling

| What | How |
|------|-----|
| **Credentials** | Never stored. Alfred reads your existing Chrome session via the local debug port — no passwords, no API keys |
| **Tokens** | Cached in memory only, cleared when Alfred restarts |
| **Config file** | `~/.alfred-config.json` — your machine only, permissions 600 |
| **External calls** | Only to your own Dynamics 365, Microsoft Graph (Outlook/Teams), and your Teams webhook |
| **No telemetry** | Alfred sends nothing to third parties |

### Installer verification

The installer runs as your own user — no sudo, no admin rights required. SHA256 checksums are published in [CHECKSUMS.txt](CHECKSUMS.txt).

```bash
curl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup.command -o ~/Downloads/Setup.command
shasum -a 256 ~/Downloads/Setup.command  # compare to CHECKSUMS.txt
bash ~/Downloads/Setup.command
```
