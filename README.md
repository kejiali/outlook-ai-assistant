# Outlook AI Assistant

A Microsoft Outlook Add-in for Mac that drafts email replies powered by AI — supporting both local (free) and cloud routing.

![Office Add-in](https://img.shields.io/badge/Office%20Add--in-Outlook-0078d4?logo=microsoft-outlook)
![Platform](https://img.shields.io/badge/platform-macOS-lightgrey?logo=apple)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

- **One-click reply drafting** — reads the email you're viewing and generates a contextual reply
- **Four rewrite modes** — Make professional, Make concise, Fix grammar, Change tone (Formal / Friendly / Assertive / Empathetic)
- **Dual provider routing** — switch between local (free) and cloud at runtime:
  - 🦞 **OpenClaw** — routes through a local proxy to Claude Sonnet via Amazon Bedrock. Zero API cost if you're already running [OpenClaw](https://openclaw.ai)
  - ⚡ **OpenAI** — routes to GPT-4o-mini via the OpenAI API (uses credits)
- **Live route indicator** — always shows exactly where your traffic is going
- **No build step** — pure HTML + vanilla JS, no webpack, no npm install to run
- **Persistent background server** — runs as a macOS launchd service, starts on login, auto-restarts on crash

---

## Architecture

```
Outlook WebView (HTTPS)
    │
    ▼
https://localhost:3000        ← Node.js HTTPS server (launchd, always-on)
    │
    ├── Static files (taskpane.html, taskpane.js, config.js)
    │
    └── /proxy/v1/*           ← Reverse proxy endpoint
            │
            ▼
    http://127.0.0.1:8402     ← OpenClaw local LLM proxy (Bedrock/Claude)
```

### Key Engineering Decisions

**Why a local HTTPS server instead of a CDN or file:// URLs?**
Office Add-ins require a URL in the manifest — `file://` is not supported. HTTPS is required by Outlook on Mac (even for localhost). Rather than depending on `npx http-server` (fragile, requires a terminal), a minimal Node.js server using only built-in modules (`https`, `http`, `fs`) is registered as a launchd service for reliability.

**Why a reverse proxy for OpenClaw?**
Outlook's WebView enforces mixed-content rules — an HTTPS page cannot make requests to an `http://` endpoint. Since OpenClaw's local proxy runs on plain HTTP (`127.0.0.1:8402`), calls are bridged through the same HTTPS server via a `/proxy/*` route, eliminating the mixed-content block without any changes to OpenClaw itself.

**Why `mkcert` instead of a self-signed cert?**
Self-signed certificates require manual trust per-device and are rejected by Outlook's WebView without additional system config. `mkcert` installs a local CA into the macOS Keychain (trusted system-wide), issuing certs that are accepted transparently — no browser warnings, no per-session exceptions.

**Why vanilla JS and no framework?**
Office Add-ins load inside a restricted WebView. Keeping the stack as plain HTML/JS avoids bundler complexity, reduces attack surface, and makes the add-in auditable at a glance. The entire logic is ~150 lines.

---

## Setup

### Prerequisites

- macOS with Outlook installed (Microsoft 365)
- Node.js (for the background server)
- [mkcert](https://github.com/FiloSottile/mkcert) for trusted local HTTPS

### 1. Install mkcert and generate certs

```bash
brew install mkcert
cd ~/Documents/github/outlook-ai-assistant
mkcert -install          # installs local CA (requires your Mac password)
mkcert localhost 127.0.0.1
```

This generates `localhost+1.pem` and `localhost+1-key.pem` (gitignored).

### 2. Configure your API key(s)

```bash
cp config.example.js config.js
```

Edit `config.js` and fill in whichever provider(s) you want:

```js
// OpenClaw (free if you're running OpenClaw locally)
var OPENCLAW_BASE_URL = "https://localhost:3000/proxy/v1";
var OPENCLAW_MODEL    = "amazon-bedrock/eu.anthropic.claude-sonnet-4-6";

// OpenAI (optional, uses credits)
var OPENAI_API_KEY = "sk-proj-...";
```

### 3. Start the background server

Register the launchd service (starts now and on every login):

```bash
cp com.mrli.outlook-addin-server.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.mrli.outlook-addin-server.plist
```

Verify it's running:

```bash
curl -sk https://localhost:3000/taskpane.html | head -3
```

### 4. Sideload the add-in in Outlook

1. Open Outlook and view any email
2. Go to **Get Add-ins** → **My Add-ins** → **+ Add a custom add-in** → **Add from file…**
3. Select `manifest.xml`
4. Look for **"AI Draft Reply"** in the Home ribbon

---

## Usage

1. Open a received email in Outlook
2. Click **AI Draft Reply** in the ribbon
3. Select your rewrite modes (and tone if needed)
4. Choose your provider: **🦞 OpenClaw** (local, free) or **⚡ OpenAI** (cloud)
5. Optionally add extra instructions
6. Click **Draft Reply** — a compose window opens with the AI-generated reply

---

## Files

| File | Purpose |
|------|---------|
| `manifest.xml` | Office Add-in manifest — sideload this in Outlook |
| `taskpane.html` | Add-in UI |
| `taskpane.js` | Core logic — reads email, calls LLM, opens reply |
| `server.js` | Minimal HTTPS static server + reverse proxy (no dependencies) |
| `config.js` | Your API keys (gitignored) |
| `config.example.js` | Template — copy to `config.js` |
| `commands.html` | Required stub for manifest commands |

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| Add-in error on load | Run `curl -sk https://localhost:3000/taskpane.html` — if empty, restart launchd service |
| `API error: Load failed` | Mixed-content issue — make sure `OPENCLAW_BASE_URL` points to `https://localhost:3000/proxy/v1`, not `http://` |
| `502` on OpenClaw route | OpenClaw isn't running — start it, or switch to OpenAI provider |
| Manifest sideload fails | Validate XML, ensure server is running, try removing and re-adding the add-in |

---

## License

MIT
