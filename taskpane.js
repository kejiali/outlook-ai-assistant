/* global Office */

var currentProvider = "openclaw"; // default

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("draft-btn").addEventListener("click", onDraftReply);
    document.getElementById("mode-tone").addEventListener("change", toggleToneRow);
  }
});

function setProvider(provider) {
  currentProvider = provider;

  // Toggle button styles
  document.getElementById("btn-openclaw").classList.toggle("active", provider === "openclaw");
  document.getElementById("btn-openai").classList.toggle("active", provider === "openai");

  // Update route indicator
  const indicator = document.getElementById("route-indicator");
  const routeText = document.getElementById("route-text");
  indicator.className = provider;

  if (provider === "openclaw") {
    routeText.textContent = "Routing via OpenClaw local proxy → Claude Sonnet (Bedrock)";
  } else {
    routeText.textContent = "Routing via OpenAI API → GPT-4o-mini (cloud, uses credits)";
  }
}

function toggleToneRow() {
  const toneRow = document.getElementById("tone-row");
  if (document.getElementById("mode-tone").checked) {
    toneRow.classList.add("visible");
  } else {
    toneRow.classList.remove("visible");
  }
}

async function onDraftReply() {
  const btn = document.getElementById("draft-btn");
  const spinner = document.getElementById("spinner");
  const btnText = document.getElementById("btn-text");
  const statusEl = document.getElementById("status");

  const modes = [];
  if (document.getElementById("mode-professional").checked) modes.push("Make professional");
  if (document.getElementById("mode-concise").checked) modes.push("Make concise");
  if (document.getElementById("mode-grammar").checked) modes.push("Fix grammar");
  if (document.getElementById("mode-tone").checked) {
    const tone = document.getElementById("tone-select").value;
    modes.push(`Change tone to ${tone}`);
  }

  if (modes.length === 0) {
    showStatus("Please select at least one rewrite mode.", "error");
    return;
  }

  const extraInstructions = document.getElementById("extra-instructions").value.trim();

  btn.disabled = true;
  spinner.style.display = "block";
  btnText.textContent = "Drafting…";
  statusEl.style.display = "none";
  statusEl.className = "";

  try {
    const item = Office.context.mailbox.item;
    const subject = item.subject || "(No subject)";

    item.body.getAsync(Office.CoercionType.Text, async function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        showStatus("Could not read email body: " + result.error.message, "error");
        resetBtn();
        return;
      }

      const emailBody = result.value || "";
      const senderName = item.from ? item.from.displayName : "the sender";
      const modeList = modes.join(", ");

      const systemPrompt = `You are an email assistant. Your job is to draft a professional email reply.

Apply these transformations to the drafted reply: ${modeList}.
${extraInstructions ? `\nAdditional instructions: ${extraInstructions}` : ""}

Rules:
- Draft a reply TO the original email, addressing ${senderName}
- Apply all requested transformations
- Be natural and human-sounding
- Do NOT include a subject line
- Do NOT add placeholders like [Your Name] — just write the body
- Output ONLY the email reply body, nothing else`;

      const userMessage = `Original email from ${senderName}:
Subject: ${subject}

${emailBody}

Please draft a reply.`;

      try {
        const draft = await callLLM(systemPrompt, userMessage);
        openReplyWithDraft(draft);

        const providerLabel = currentProvider === "openclaw"
          ? "Claude Sonnet via OpenClaw (local, free)"
          : "GPT-4o-mini via OpenAI (cloud)";
        showStatus(`✓ Reply drafted by ${providerLabel}. Check your compose window.`, "success");
      } catch (err) {
        showStatus("API error: " + err.message, "error");
      }

      resetBtn();
    });

  } catch (err) {
    showStatus("Error: " + err.message, "error");
    resetBtn();
  }
}

async function callLLM(systemPrompt, userMessage) {
  if (currentProvider === "openclaw") {
    return callOpenClaw(systemPrompt, userMessage);
  } else {
    return callOpenAI(systemPrompt, userMessage);
  }
}

async function callOpenClaw(systemPrompt, userMessage) {
  const baseUrl = window.OPENCLAW_BASE_URL || "http://127.0.0.1:8402/v1";
  const model = window.OPENCLAW_MODEL || "amazon-bedrock/eu.anthropic.claude-sonnet-4-6";

  const response = await fetch(baseUrl + "/chat/completions", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      model: model,
      max_tokens: 1024,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userMessage }
      ]
    })
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `HTTP ${response.status} — is OpenClaw running?`);
  }

  const data = await response.json();
  return data.choices[0].message.content;
}

async function callOpenAI(systemPrompt, userMessage) {
  if (!window.OPENAI_API_KEY || window.OPENAI_API_KEY === "your-key-here") {
    throw new Error("OpenAI API key not set. Edit config.js.");
  }

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": "Bearer " + window.OPENAI_API_KEY,
      "content-type": "application/json"
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      max_tokens: 1024,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userMessage }
      ]
    })
  });

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `HTTP ${response.status}`);
  }

  const data = await response.json();
  return data.choices[0].message.content;
}

function openReplyWithDraft(draftText) {
  const item = Office.context.mailbox.item;
  item.displayReplyForm({
    htmlBody: draftText.replace(/\n/g, "<br>")
  });
}

function showStatus(message, type) {
  const statusEl = document.getElementById("status");
  statusEl.textContent = message;
  statusEl.className = type;
  statusEl.style.display = "block";
}

function resetBtn() {
  const btn = document.getElementById("draft-btn");
  const spinner = document.getElementById("spinner");
  const btnText = document.getElementById("btn-text");
  btn.disabled = false;
  spinner.style.display = "none";
  btnText.textContent = "Draft Reply";
}
