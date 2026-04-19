# AIfred.mcp

> *Every superhero needs a butler. Yours is called Alfred and handles your CRM.*

Built by **Fred** — Solution Consultant @ ServiceNow

<a href="https://www.buymeacoffee.com/h22fred"><img src="https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png" alt="Buy Me A Coffee" height="60"></a>&nbsp;&nbsp;<a href="https://www.linkedin.com/comm/mynetwork/discovery-see-all?usecase=PEOPLE_FOLLOWS&followMember=fredholmstrom"><img src="setup/assets/btn-linkedin.svg" alt="Follow on LinkedIn" height="60"></a>

<a href="https://twitter.com/intent/follow?screen_name=h22fred"><img src="setup/assets/btn-x.svg" alt="Follow on X" height="60"></a>&nbsp;&nbsp;<a href="https://github.com/h22fred"><img src="setup/assets/btn-github.svg" alt="Follow on GitHub" height="60"></a>

Connects Claude Desktop directly to your CRM, calendar, email and Teams — using your existing browser session. No Azure app registration. No stored credentials. No CRM admin work ever again.

Two flavours — one installer:

| Variant | Who | Install folder |
|---------|-----|----------------|
| **Alfred SC** | SC / SSC / Manager | `~/Documents/alfred.sc` |
| **Alfred Sales** | AE / Sales Specialist / Manager | `~/Documents/alfred.sales` |

---

## What it does

### Alfred SC (Solution Consulting)

| Source | Capabilities |
|--------|-------------|
| **Dynamics 365** | List opportunities, create/update/complete engagements, hygiene sweep, Tech Win assessment, delete cancelled engagements |
| **Outlook Calendar** | Show calendar by date range, search meetings |
| **Outlook Email** | Search emails, list inbox/sent/subfolders, full body, filter unread |
| **Teams** | Get meeting transcripts, post to channels, read chats |
| **Account Insights** | License utilization, renewal dates, upsell/cross-sell detection |

### Alfred Sales (Account Executive)

| Source | Capabilities |
|--------|-------------|
| **Dynamics 365** | Create & update opportunities, assign SC, search accounts/users, add notes, territory pipeline overview |
| **Outlook Calendar** | Show calendar by date range, search meetings |
| **Outlook Email** | Search emails, list inbox/sent/subfolders, full body, filter unread |
| **Teams** | Get meeting transcripts, post to channels, read chats |
| **Account Insights** | License utilization, renewal dates, upsell/cross-sell detection |

---

## Requirements

- macOS or Windows 10+
- [Claude Desktop](https://claude.ai/download)
- Google Chrome
- Node.js — **installed automatically if missing** (macOS)

---

## Setup — macOS

**Option A — Download and run** _(recommended)_
1. **[⬇️ Download Setup_macOS.command](https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_macOS.command)**
2. Open Terminal and run:
```bash
bash ~/Downloads/Setup_macOS.command
```

**Option B — One-liner** _(for the terminal-comfortable)_
```bash
curl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_macOS.command | bash
```

> If a popup appears asking to install Command Line Tools, click **Install**, wait for it to finish, then run the same command again.

## Setup — Windows

1. **[⬇️ Download Setup_Windows.bat](https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_Windows.bat)**
2. Double-click the downloaded file (or right-click → Run as administrator if prompted)

> **Prerequisites:** Git and Node.js must be installed before running the Windows installer.
> - Git: [git-scm.com/download/win](https://git-scm.com/download/win)
> - Node.js LTS: [nodejs.org](https://nodejs.org)

## What the installer asks

- **SC or Sales?** — determines which Alfred variant is installed
- **Dynamics company name** — your CRM URL (e.g. `servicenow`)
- **Teams webhook** — for automated notifications (optional)
- **Role:**
  - SC variant: **SC** / **SSC** / **Manager**
  - Sales variant: **AE** / **Sales Specialist** / **Manager**
- **Engagement types** — which milestones you track
- **Automated checks** — optional weekly hygiene sweep + meeting review (you can skip entirely or configure each one)

### Roles explained

| Role | Variant | Pipeline view |
|------|---------|--------------|
| **SC** | SC | Your assigned opportunities only |
| **SSC** | SC | All opportunities (support role, no assigned pipeline) |
| **SC Manager** | SC | Team-wide / territory view |
| **AE** | Sales | Your owned opportunities only |
| **Sales Specialist** | Sales | All opportunities (AE CRM, AE Risk, etc. — no assigned pipeline) |
| **Sales Manager** | Sales | Team-wide / territory view |

---

## Every session

1. Double-click **Alfred** on your Desktop (macOS: `.app`, Windows: `.bat`)
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

### Sales AE / Specialist / Manager
```
Create an opportunity for Givaudan — New ITSM, close December 2026
Find the account ID for Roche
Search for Fredrik to get his SC GUID and assign him
Update the PMI opportunity close date to March 2027
Show my open opportunities
Show me the full territory pipeline
Add a note to the SITA opportunity: had intro call, next step is discovery
```

---

## How it works

Alfred launches Chrome with `--remote-debugging-port=9222` using a dedicated profile (`~/.alfred-profile`). The MCP server extracts session cookies and Bearer tokens via CDP WebSocket — no credentials stored, no Azure registration needed.

Auth flow:
1. **Dynamics:** reads `CrmOwinAuthC1/C2` cookies via `Network.getCookies`
2. **Outlook/Graph:** reads Bearer token from MSAL cache in page storage, falls back to network interception via `Fetch.enable` (catches Teams v2 service worker requests)
3. All tokens cached in-memory for the session duration

---

## Automated jobs

During installation you're asked if you want automated checks. If yes, Alfred sets up two recurring jobs (both individually optional):

| When | What |
|------|------|
| **Monday 9:30am** (default) | CRM hygiene sweep — checks all open opportunities for missing milestones, posts a summary card to Teams |
| **Friday 2:00pm** (default) | Meeting review — scans this week's calendar, matches meetings to open opportunities, suggests engagements to log |

Both jobs run silently in the background. If Alfred isn't running or your Dynamics session has expired, they post a Teams reminder instead of failing silently.

To run manually at any time:
```bash
# macOS
node ~/Documents/alfred.sc/setup/hygiene-sweep.mjs
node ~/Documents/alfred.sc/setup/post-meeting-sweep.mjs
```

To change the schedule:
```bash
# macOS
crontab -e

# Windows — use Task Scheduler or:
schtasks /query /tn "Alfred-HygieneSweep"
schtasks /query /tn "Alfred-MeetingReview"
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

> **Keep the webhook URL private.** Anyone with this URL can post messages to your channel. Never share it or commit it to version control.

---

## Troubleshooting

| Error | Fix |
|-------|-----|
| Alfred not running | Double-click Alfred on Desktop |
| Not logged into Dynamics | Log into Dynamics in the Alfred browser window |
| 401 from Dynamics | Session expired — re-login in Alfred window |
| Teams not posting | Re-run setup to reconfigure webhook |
| Node.js not found | Re-run setup — it installs automatically (macOS) or install from [nodejs.org](https://nodejs.org) (Windows) |
| macOS privacy prompt for Node.js | Normal — click Allow to let Node.js access local project files |
| No opportunities found (Sales AE) | Make sure your role is set correctly in setup. AE CRM / AE Risk users should select "Sales Specialist" |

---

*Questions? Open a [Discussion](https://github.com/h22fred/Alfred.mcp/discussions) or ping Fred on Teams.*

---

## Security

### Security assessment

A full security review was conducted on this codebase covering credential handling, input validation, API calls, network exposure, and the install scripts. No vulnerabilities were found in the core data handling or CRM integration layer. Installer integrity can be verified via the checksum below.

Automated audit run via [Ruflo](https://github.com/h22fred/ruflo) on 2026-03-25:

| Check | Result |
|-------|--------|
| Security scan (full depth) | ✅ 0 issues |
| Secrets / hardcoded credentials | ✅ None found |
| CVEs (npm audit) | ✅ 0 vulnerabilities |
| Prompt injection defence | ✅ 0 detections |
| External dependencies | ✅ 8 only (MCP SDK, zod, Node built-ins) |

`npm ci` runs a vulnerability audit automatically on every install and will report any issues found in dependencies.

### Data handling

| What | How |
|------|-----|
| **Credentials** | Never stored. Alfred reads your existing Chrome session via the local debug port — no passwords, no API keys |
| **Tokens** | Cached in memory only, cleared when Alfred restarts |
| **Config file** | `~/.alfred-config.json` — your machine only, permissions 600 |
| **External calls** | Only to your own Dynamics 365, Microsoft Graph (Outlook/Teams), and your Teams webhook |
| **No telemetry** | Alfred sends nothing to third parties |

### Installer verification

The installer runs as your own user — no sudo, no admin rights required. SHA256 checksums are published in [setup/CHECKSUMS.txt](setup/CHECKSUMS.txt).

```bash
curl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_macOS.command -o ~/Downloads/Setup_macOS.command
shasum -a 256 ~/Downloads/Setup_macOS.command  # compare to setup/CHECKSUMS.txt
bash ~/Downloads/Setup_macOS.command
```
