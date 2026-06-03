# Alfred — Windows Installation Guide

This guide installs Alfred without running any .bat or .ps1 scripts.  
Estimated time: **10–15 minutes**.

---

## Before you start

You need:
- **Claude Desktop** (already installed)
- **Node.js 20+** — download from https://nodejs.org (LTS version, use defaults)
- **Git for Windows** — download from https://git-scm.com/download/win (use defaults, include Git Bash)

> **Why Git?** Alfred uses `git pull` internally to update itself. Installing Git now means
> you can update Alfred from Claude with a single command later — no manual downloads needed.

---

## Step 1 — Install Git for Windows

1. Download Git from https://git-scm.com/download/win
2. Run the installer — use all default options
3. When asked about the default editor, any choice is fine
4. Verify: open **Git Bash** (search in Start menu) and run:
   ```
   git --version
   ```
   You should see `git version 2.x.x`

---

## Step 2 — Install Node.js

1. Download from https://nodejs.org (pick the **LTS** version)
2. Run the installer — use all defaults
3. Verify in **Git Bash**:
   ```
   node --version
   npm --version
   ```

---

## Step 3 — Clone Alfred

Open **Git Bash** and run the correct command for your role:

**For Solution Consultants (alfred.sc):**
```bash
git clone https://github.com/h22fred/Alfred.mcp ~/Documents/alfred.sc
cd ~/Documents/alfred.sc
npm install
npm run build
```

**For Account Executives (alfred.sales):**
```bash
git clone https://github.com/h22fred/Alfred.mcp ~/Documents/alfred.sales
cd ~/Documents/alfred.sales
npm install
npm run build
```

---

## Step 4 — Configure Alfred

Still in Git Bash, create the config file:

```bash
cat > ~/.alfred-config.json << 'EOF'
{
  "dynamicsUrl": "https://YOUR-COMPANY.crm.dynamics.com",
  "role": "sc"
}
EOF
```

Replace `YOUR-COMPANY` with your actual Dynamics subdomain (e.g. `servicenow` → `https://servicenow.crm.dynamics.com`).

Replace `"sc"` with your role if needed: `sc`, `ssc`, `manager`, `sales`, `sales_specialist`, `sales_manager`.

> Alternatively, after completing Step 5, you can ask Alfred in Claude:  
> *"Configure Alfred"* — it will walk you through the setup conversationally.

---

## Step 5 — Add Alfred to Claude Desktop

1. Open Claude Desktop
2. Go to **Settings → Developer → Edit Config** (or open the file directly):
   ```
   %APPDATA%\Claude\claude_desktop_config.json
   ```
3. Add the Alfred entry inside `"mcpServers"`:

**For alfred.sc** (SC role):
```json
{
  "mcpServers": {
    "alfred": {
      "command": "node",
      "args": ["C:/Users/YOUR_USERNAME/Documents/alfred.sc/dist/sc/index.js"]
    }
  }
}
```

**For alfred.sales** (AE role):
```json
{
  "mcpServers": {
    "alfred": {
      "command": "node",
      "args": ["C:/Users/YOUR_USERNAME/Documents/alfred.sales/dist/sales/index.js"]
    }
  }
}
```

Replace `YOUR_USERNAME` with your actual Windows username.

4. Save the file and **restart Claude Desktop**

---

## Step 6 — First launch

1. In Claude Desktop, open a new conversation
2. Alfred will open a Chromium browser window automatically
3. Log in to Dynamics 365 in that window
4. Ask Claude: *"Run a hygiene sweep"* to verify everything works

---

## Updating Alfred

Because you installed with Git, updates work directly from Claude:

> *"Update Alfred"*

Alfred will pull the latest version, rebuild, and ask you to restart Claude Desktop.

---

## Troubleshooting

**"node is not recognized"** — Node.js is not on your PATH. Restart Git Bash after installing Node, or reinstall Node with "Add to PATH" checked.

**"git is not recognized"** — Restart Git Bash after installing Git.

**Alfred browser doesn't open** — Check that the path in claude_desktop_config.json uses forward slashes (`/`) and matches your actual username.

**"Invalid Dynamics URL"** — Check ~/.alfred-config.json — the URL must match `https://COMPANY.crm.dynamics.com` exactly.

**Questions?** Reach out to Fredrik.
