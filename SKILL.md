---
name: outlook-ai-assistant
description: >
  Install, configure, and troubleshoot the Outlook AI Assistant add-in.
  Use when the user wants to set up the add-in, restart the server, switch
  providers, update API keys, or diagnose errors in Outlook.
---

# Outlook AI Assistant — OpenClaw Skill

This skill helps you manage the Outlook AI Assistant add-in on macOS.

## Variables

Set these based on the user's machine before running any commands:

```
REPO_DIR  = /Users/<username>/Documents/github/outlook-ai-assistant
PLIST     = ~/Library/LaunchAgents/<plist-filename>
PLIST_SRC = <repo-dir>/<plist-filename>
SERVER    = https://localhost:3000
```

> The plist filename is found in the repo root — it matches the launchd label inside the file. Check with:
> ```bash
> ls ~/Documents/github/outlook-ai-assistant/*.plist
> ```

---

## Install

### 1. Clone the repository

```bash
cd ~/Documents/github
git clone https://github.com/kejiali/outlook-ai-assistant.git
cd outlook-ai-assistant
```

### 2. Generate trusted local HTTPS certificates

```bash
brew install mkcert
mkcert -install        # requires Mac password — installs local CA to Keychain
mkcert localhost 127.0.0.1
```

Expected output: `localhost+1.pem` and `localhost+1-key.pem` in the project folder.

### 3. Configure API keys

```bash
cp config.example.js config.js
```

Edit `config.js`:
- Set `OPENCLAW_MODEL` to match your OpenClaw provider (e.g. `"amazon-bedrock/eu.anthropic.claude-sonnet-4-6"`)
- Set `OPENAI_API_KEY` if you want the OpenAI fallback

### 4. Register the background server (launchd)

```bash
PLIST=$(ls *.plist | head -1)
cp "$PLIST" ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/"$PLIST"
```

### 5. Verify server is running

```bash
curl -sk https://localhost:3000/taskpane.html | head -3
```

Expected: `<!DOCTYPE html>` — if empty or error, see Troubleshooting below.

### 6. Sideload the add-in in Outlook

1. Open Outlook → open any received email
2. **Get Add-ins** → **My Add-ins** → **+ Add a custom add-in** → **Add from file…**
3. Select `manifest.xml` from the repo folder
4. Look for **"AI Draft Reply"** in the Home ribbon

---

## Server Management

### Check if running

```bash
curl -sk https://localhost:3000/taskpane.html | head -1
```

### Restart

```bash
PLIST=$(ls ~/Library/LaunchAgents/*.plist | xargs grep -l "outlook-ai-assistant\|outlook-addin" 2>/dev/null | head -1)
launchctl unload "$PLIST"
launchctl load   "$PLIST"
sleep 2
curl -sk https://localhost:3000/taskpane.html | head -1
```

### View logs

```bash
tail -50 ~/Documents/github/outlook-ai-assistant/server.log
```

### Stop permanently

```bash
PLIST=$(ls ~/Library/LaunchAgents/*.plist | xargs grep -l "outlook-ai-assistant\|outlook-addin" 2>/dev/null | head -1)
launchctl unload "$PLIST"
```

---

## Configuration

### Switch OpenClaw model

Edit `config.js` and update `OPENCLAW_MODEL`. No server restart needed — reopen the task pane in Outlook.

### Update OpenAI API key

Edit `config.js` and replace `OPENAI_API_KEY`. No server restart needed.

### Test OpenClaw proxy

```bash
curl -sk -X POST https://localhost:3000/proxy/v1/chat/completions \
  -H "content-type: application/json" \
  -d '{"model":"YOUR_MODEL","max_tokens":10,"messages":[{"role":"user","content":"hi"}]}'
```

Expected: a JSON response with `choices[0].message.content`.

### Test OpenAI key

```bash
curl -s https://api.openai.com/v1/chat/completions \
  -H "Authorization: Bearer YOUR_KEY" \
  -H "content-type: application/json" \
  -d '{"model":"gpt-4o-mini","max_tokens":10,"messages":[{"role":"user","content":"hi"}]}'
```

---

## Troubleshooting

### Add-in error on load in Outlook

Server isn't running or cert isn't trusted.

```bash
# Check server
curl -sk https://localhost:3000/taskpane.html | head -1

# If empty — find and restart the launchd service
PLIST=$(ls ~/Library/LaunchAgents/*.plist | xargs grep -l "outlook-ai-assistant\|outlook-addin" 2>/dev/null | head -1)
launchctl unload "$PLIST"
launchctl load   "$PLIST"
```

If the server starts but Outlook still rejects it, the cert may not be trusted. Re-run:

```bash
mkcert -install
```

### `API error: Load failed`

Mixed-content block — `OPENCLAW_BASE_URL` in `config.js` must use `https://`, not `http://`:

```js
var OPENCLAW_BASE_URL = "https://localhost:3000/proxy/v1";  // ✅
var OPENCLAW_BASE_URL = "http://127.0.0.1:8402/v1";        // ❌ blocked by WebView
```

### `502` on OpenClaw route

OpenClaw's local proxy (`127.0.0.1:8402`) isn't running. Start OpenClaw, or switch to the OpenAI provider in the task pane.

### API error with non-English message (e.g. token expired / auth failure)

This is typically an environment-specific issue — the API key in `config.js` may be expired, rate-limited, or from a third-party proxy that has its own auth layer. Replace it with a fresh key from your provider's console.

### Manifest sideload fails — "Sorry, we can't complete this operation"

- Confirm server is running: `curl -sk https://localhost:3000/taskpane.html | head -1`
- Remove the existing add-in first: **Get Add-ins** → **My Add-ins** → remove → re-add
- Validate `manifest.xml` has no stray `http://` URLs (all should be `https://`)

### Add-in button doesn't appear in ribbon

The add-in only activates in the **reading pane** (viewing a received email). It will not appear when composing, replying to, or forwarding an email.

---

## Important Notes

- The add-in **only works when reading a received email** — not in compose, reply, or forward windows
- Data privacy depends on your provider: see the `README.md` ⚠️ section
- `config.js` is gitignored — never commit your API keys
- Certs (`localhost+1.pem`, `localhost+1-key.pem`) are gitignored — regenerate per machine with `mkcert`
