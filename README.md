# AIfred.mcp

> *Every superhero needs a butler. Yours is called Alfred and handles your CRM.*

Built by **Fred** — Solution Consultant

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
| **Dynamics 365** | List opportunities (incl. colleague pipeline), create/update/complete engagements, create on behalf of colleague, hygiene sweep, Tech Win assessment, collaboration notes, delete cancelled engagements |
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

**Windows only — install these two first:**
- [Git for Windows](https://git-scm.com/download/win) — click Next through all defaults
- [Node.js LTS](https://nodejs.org/en/download) — click Next through all defaults

> macOS: Git and Node.js are installed automatically by the setup script.

---

## Setup — macOS

**Option A — One-liner** _(no security prompts — recommended)_
```bash
curl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_macOS.command | bash
```

> Piping directly to `bash` skips macOS Gatekeeper entirely — the script is never written to disk so no quarantine prompt appears.

**Option B — Download and run**
1. **[⬇️ Download Setup_macOS.command](https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_macOS.command)**
2. Open Terminal and run:
```bash
bash ~/Downloads/Setup_macOS.command
```

> **macOS Sequoia+:** Downloaded files may be blocked ("cannot confirm it is free from malware"). Go to **System Settings → Privacy & Security**, scroll down and click **Open Anyway**, then re-run the command.

> Git (via Xcode Command Line Tools) and Node.js (via nvm) are **installed automatically** if missing.
>
> If a popup appears asking to install Command Line Tools, click **Install**, wait for it to finish, then run the same command again.

## Setup — Windows

**Step 1 — Install Git** _(if not already installed)_
Download and run (click Next through all defaults): [git-scm.com/download/win](https://git-scm.com/download/win)

**Step 2 — Install Node.js** _(if not already installed)_
Download the LTS version and run (click Next through all defaults): [nodejs.org/en/download](https://nodejs.org/en/download)

**Step 3 — Install Alfred**
1. **[⬇️ Download Setup_Windows.bat](https://raw.githubusercontent.com/h22fred/Alfred.mcp/refs/heads/main/Setup_Windows.bat)**
2. Double-click the downloaded file
3. If Windows SmartScreen appears ("Windows protected your PC"), click **More info → Run anyway** — this is expected for unsigned scripts downloaded from the internet
4. When prompted, enter `1` for SC or `2` for Sales

**Prefer a PowerShell one-liner?** Use this safe form (downloads to a file first — does NOT use `IEX`):
```powershell
$f="$env:TEMP\Setup_Windows.bat"; Invoke-WebRequest -Uri "https://raw.githubusercontent.com/h22fred/Alfred.mcp/main/Setup_Windows.bat" -OutFile $f; Start-Process cmd.exe -ArgumentList "/c `"$f`"" -Wait -NoNewWindow
```

> **Enterprise / managed machines:** If your EDR flags the installer, ask Endpoint Security Engineering to add an exclusion for `Setup_Windows.bat` and `setup\setup.ps1` from this repo. Do **not** use `iex (... DownloadString ...)` one-liners — those trigger detections by design and are not the official install method.

## What the installer asks

- **SC or Sales?** — determines which Alfred variant is installed
- **Dynamics company name** — your CRM subdomain (e.g. `mycompany` for `mycompany.crm.dynamics.com`)
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

## Updating Alfred

In Claude Desktop, just say **"update Alfred"** — the `update_alfred` tool handles it automatically.

If you're on an older version where the auto-update fails, run this one-liner in Terminal (once):
```bash
curl -fsSL https://raw.githubusercontent.com/h22fred/Alfred.mcp/main/setup/update.sh | bash
```
After that, `update_alfred` in Claude will work for all future updates.

---

## Every session

1. Open Claude Desktop — Alfred's browser launches automatically in the background
2. **First time (and ~every 8 hours when tokens expire):** log into Dynamics, Outlook and Teams (SSO) in the Alfred browser window — the browser appears briefly, then closes automatically once auth is complete

---

## Example prompts

### SC / SSC / Manager
```
List my open opportunities over $100K
Run hygiene sweep and post to Teams
Which accounts are missing a Technical Win?
Assess the Tech Win for Acme Corp
Create a Discovery engagement for Contoso from my Tuesday meeting
Mark the Acme Tech Win as complete
Show my calendar this week
Get the transcript from my Contoso demo last Thursday
Detect post-meeting engagements from this week
Delete the cancelled Demo engagement on Fabrikam
Show me Stéphane's open pipeline
Create a Discovery placeholder on behalf of Stéphane for the Acme opp
```

### Sales AE / Specialist / Manager
```
Create an opportunity for Contoso — New ITSM, close December 2026
Find the account ID for Fabrikam
Search for John to get his SC GUID and assign him
Update the Fabrikam opportunity close date to March 2027
Show my open opportunities
Show me the full territory pipeline
Add a note to the Acme opportunity: had intro call, next step is discovery
```

---

## How it works

Alfred launches a private Chromium browser (named "Alfred") via Playwright with a dedicated profile (`~/.alfred-pw`). The MCP server reads session cookies directly from the live browser — no credentials stored, no Azure registration needed.

Auth flow:
1. **Dynamics:** reads `CrmOwinAuthC1/C2` cookies from the live session
2. **Outlook/Graph:** reads session cookies from the Outlook tab in the Alfred browser
3. Tokens cached to disk — survive Claude Desktop restarts without re-login
4. Browser closes automatically ~3 seconds after tokens are cached — reopens silently when next needed (typically every ~8 hours when the cache expires)

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
# macOS — SC variant
node ~/Documents/alfred.sc/setup/hygiene-sweep.mjs
node ~/Documents/alfred.sc/setup/post-meeting-sweep.mjs

# macOS — Sales variant
node ~/Documents/alfred.sales/setup/hygiene-sweep.mjs
node ~/Documents/alfred.sales/setup/post-meeting-sweep.mjs
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
| Alfred not running | Restart Claude Desktop — Alfred launches automatically |
| Not logged into Dynamics | Log into Dynamics in the Alfred browser window |
| 401 from Dynamics | Session expired — re-login in Alfred window |
| Teams not posting | Re-run setup to reconfigure webhook |
| Node.js not found | macOS: re-run setup — it installs automatically. Windows: install from [nodejs.org](https://nodejs.org/en/download) then re-run `Setup_Windows.bat` |
| macOS privacy prompt for Node.js | Normal — click Allow to let Node.js access local project files |
| No opportunities found (Sales AE) | Make sure your role is set correctly in setup. AE CRM / AE Risk users should select "Sales Specialist" |

---

*Questions? Open a [Discussion](https://github.com/h22fred/Alfred.mcp/discussions) or ping Fred on Teams.*

---

## Security

### Security assessment

A full security review was conducted on this codebase covering credential handling, input validation, API calls, network exposure, and the install scripts. No vulnerabilities were found in the core data handling or CRM integration layer. Installer integrity can be verified via the checksum below.

Automated audit run via [Ruflo](https://github.com/h22fred/ruflo) on 2026-05-21:

| Check | Result |
|-------|--------|
| Security scan (full depth) | ✅ 0 issues |
| Secrets / hardcoded credentials | ✅ None found |
| CVEs (npm audit) | ✅ 0 vulnerabilities |
| Prompt injection defence | ✅ 0 detections |
| External dependencies | ✅ 3 only (MCP SDK, Playwright, Hono) |

`npm ci` runs a vulnerability audit automatically on every install and will report any issues found in dependencies.

### Data handling

| What | How |
|------|-----|
| **Credentials** | Never stored. Alfred reads your existing browser session via Playwright — no passwords, no API keys |
| **Tokens** | Cached to disk and memory — survive Claude Desktop restarts without re-login |
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
