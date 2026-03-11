import express from "express";
import fs from "fs";
import path from "path";
import { createQueueStore } from "./queue-store.js";

export function createApp({ baseDir = process.cwd() } = {}) {
  const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET || "change-me";
  const PIPEDRIVE_BASE_URL = process.env.PIPEDRIVE_BASE_URL || "";
  const PIPEDRIVE_API_TOKEN = process.env.PIPEDRIVE_API_TOKEN || "";
  const PERSON_DM_ELIGIBLE_FIELD_KEY = process.env.PERSON_DM_ELIGIBLE_FIELD_KEY || "";
  const CALL_DISPOSITION_TRIGGER_OPTION_ID = String(process.env.CALL_DISPOSITION_TRIGGER_OPTION_ID || "6");
  const CALL_DISPOSITION_TRIGGER_LABEL = process.env.CALL_DISPOSITION_TRIGGER_LABEL || "LinkedIn Outreach next step";
  const ADMIN_USERNAME = String(process.env.ADMIN_USERNAME || "").trim();
  const ADMIN_PASSWORD = String(process.env.ADMIN_PASSWORD || "").trim();

  const sequencesPath = path.join(baseDir, "data", "sequences.json");
  const queueStore = createQueueStore({ baseDir });
  const queueReady = queueStore.init();

  const app = express();
  app.use((req, res, next) => {
    const origin = String(req.headers.origin || "");
    const allowOrigin = origin.startsWith("chrome-extension://")
      || origin.startsWith("http://localhost")
      || origin.startsWith("https://localhost");

    res.setHeader("Access-Control-Allow-Origin", allowOrigin ? origin : "*");
    res.setHeader("Vary", "Origin");
    res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, x-peak-access-secret");

    if (req.method === "OPTIONS") {
      res.status(204).end();
      return;
    }

    next();
  });
  app.use(express.json({ limit: "1mb" }));

  app.get("/health", (_req, res) => {
    res.json({ ok: true, service: "peak-access-linkedin-backend", now: new Date().toISOString() });
  });

  app.get("/admin", (req, res) => {
    if (!isAdminAuthorized(req, { username: ADMIN_USERNAME, password: ADMIN_PASSWORD })) {
      res.setHeader("WWW-Authenticate", 'Basic realm="Peak Access Admin"');
      res.status(401).send("Unauthorized");
      return;
    }
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.send(renderAdminHtml());
  });

  app.get("/admin/api/sequences", (req, res) => {
    if (!isAdminAuthorized(req, { username: ADMIN_USERNAME, password: ADMIN_PASSWORD })) {
      res.status(401).json({ ok: false, error: "Unauthorized" });
      return;
    }

    const data = readJson(sequencesPath, { version: 1, updated_at: null, sequences: [] });
    res.json({ ok: true, data });
  });

  app.put("/admin/api/sequences", (req, res) => {
    if (!isAdminAuthorized(req, { username: ADMIN_USERNAME, password: ADMIN_PASSWORD })) {
      res.status(401).json({ ok: false, error: "Unauthorized" });
      return;
    }

    const payload = req.body || {};
    const validationError = validateSequencesPayload(payload);
    if (validationError) {
      res.status(400).json({ ok: false, error: validationError });
      return;
    }

    const nextValue = {
      ...payload,
      updated_at: new Date().toISOString()
    };

    try {
      writeJsonAtomic(sequencesPath, nextValue);
      res.json({ ok: true, data: nextValue });
    } catch (error) {
      res.status(500).json({ ok: false, error: `Failed to save sequences: ${error.message || String(error)}` });
    }
  });

  app.get("/sequences", (_req, res) => {
    const data = readJson(sequencesPath, { version: 1, updated_at: null, sequences: [] });
    const summary = data.sequences.map((sequence) => ({
      id: sequence.id,
      name: sequence.name,
      description: sequence.description,
      recommended_start_stage: sequence.recommended_start_stage
    }));

    res.json({ ok: true, version: data.version, updated_at: data.updated_at, sequences: summary });
  });

  app.get("/sequences/:id", (req, res) => {
    const data = readJson(sequencesPath, { sequences: [] });
    const sequence = data.sequences.find((item) => item.id === req.params.id);

    if (!sequence) {
      res.status(404).json({ ok: false, error: "Sequence not found" });
      return;
    }

    res.json({ ok: true, sequence });
  });

  app.get("/templates", (req, res) => {
    const sequenceId = String(req.query.sequence_id || "").trim();
    const stage = req.query.stage ? Number(req.query.stage) : null;

    if (!sequenceId) {
      res.status(400).json({ ok: false, error: "sequence_id query param is required" });
      return;
    }

    const data = readJson(sequencesPath, { sequences: [] });
    const sequence = data.sequences.find((item) => item.id === sequenceId);

    if (!sequence) {
      res.status(404).json({ ok: false, error: "Sequence not found" });
      return;
    }

    const templates = Array.isArray(sequence.templates) ? sequence.templates : [];
    const filtered = Number.isFinite(stage) ? templates.filter((item) => Number(item.stage) === stage) : templates;

    res.json({ ok: true, sequence_id: sequenceId, stage: Number.isFinite(stage) ? stage : null, templates: filtered });
  });

  app.get("/eligible/:personId", async (req, res) => {
    const personId = String(req.params.personId || "").trim();
    await ensureQueueReady(queueReady);
    const item = await queueStore.get(personId);

    res.json({ ok: true, person_id: personId, eligible: Boolean(item?.eligible), item });
  });

  app.post("/pipedrive/webhook", async (req, res) => {
    if (!isWebhookAuthorized(req, WEBHOOK_SECRET)) {
      res.status(401).json({ ok: false, error: "Unauthorized" });
      return;
    }

    const payload = req.body || {};
    const personId = extractPersonId(payload);

    if (!personId) {
      res.status(200).json({ ok: true, ignored: true, reason: "No person_id in payload" });
      return;
    }

    if (!payloadHasTriggerOption(payload, { CALL_DISPOSITION_TRIGGER_OPTION_ID, CALL_DISPOSITION_TRIGGER_LABEL })) {
      res.status(200).json({ ok: true, ignored: true, reason: "Trigger option not present" });
      return;
    }

    await ensureQueueReady(queueReady);
    await queueStore.upsertEligible(personId, summarizePayload(payload), payload);

    if (PIPEDRIVE_BASE_URL && PIPEDRIVE_API_TOKEN && PERSON_DM_ELIGIBLE_FIELD_KEY) {
      try {
        await updatePipedriveEligibleFlag({
          personId,
          eligible: true,
          baseUrl: PIPEDRIVE_BASE_URL,
          apiToken: PIPEDRIVE_API_TOKEN,
          fieldKey: PERSON_DM_ELIGIBLE_FIELD_KEY
        });
      } catch (error) {
        console.error("Failed to update Pipedrive eligible flag", error.message);
      }
    }

    res.status(200).json({ ok: true, person_id: personId, eligible: true });
  });

  return app;
}

function isWebhookAuthorized(req, webhookSecret) {
  const headerSecret = String(req.header("x-peak-access-secret") || "").trim();
  if (headerSecret && headerSecret === webhookSecret) {
    return true;
  }

  const querySecret = String(req.query?.secret || "").trim();
  if (querySecret && querySecret === webhookSecret) {
    return true;
  }

  const basic = parseBasicAuth(req.header("authorization"));
  if (basic.password && basic.password === webhookSecret) {
    return true;
  }

  return false;
}

function readJson(filePath, fallback) {
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch (_error) {
    return fallback;
  }
}

async function ensureQueueReady(queueReady) {
  try {
    await queueReady;
  } catch (error) {
    throw new Error(`Queue storage init failed: ${error.message || String(error)}`);
  }
}

function extractPersonId(payload) {
  const direct = payload?.data?.person_id || payload?.current?.person_id || payload?.person_id;
  if (direct && String(direct).trim()) {
    return String(typeof direct === "object" ? direct.value || direct.id : direct);
  }

  const nested = payload?.data?.person?.id || payload?.current?.person?.id;
  if (nested && String(nested).trim()) {
    return String(nested);
  }

  return "";
}

function payloadHasTriggerOption(payload, triggerConfig) {
  const candidateValues = [payload?.data, payload?.current, payload];

  for (const value of candidateValues) {
    if (!value || typeof value !== "object") continue;

    const matches = flattenValues(value).map((item) => String(item));
    const hasOptionId = matches.some((item) => item === triggerConfig.CALL_DISPOSITION_TRIGGER_OPTION_ID);
    const hasLabel = matches.some((item) =>
      item.toLowerCase().includes(triggerConfig.CALL_DISPOSITION_TRIGGER_LABEL.toLowerCase())
    );

    if (hasOptionId || hasLabel) {
      return true;
    }
  }

  return false;
}

function flattenValues(value) {
  if (Array.isArray(value)) {
    return value.flatMap((entry) => flattenValues(entry));
  }

  if (value && typeof value === "object") {
    return Object.values(value).flatMap((entry) => flattenValues(entry));
  }

  return [value];
}

function parseBasicAuth(headerValue) {
  const value = String(headerValue || "").trim();
  if (!value.toLowerCase().startsWith("basic ")) {
    return { username: "", password: "" };
  }

  try {
    const encoded = value.slice(6).trim();
    const decoded = Buffer.from(encoded, "base64").toString("utf8");
    const separatorIndex = decoded.indexOf(":");
    if (separatorIndex < 0) {
      return { username: decoded, password: "" };
    }

    return {
      username: decoded.slice(0, separatorIndex),
      password: decoded.slice(separatorIndex + 1)
    };
  } catch (_error) {
    return { username: "", password: "" };
  }
}

function summarizePayload(payload) {
  return {
    event: payload?.event || payload?.meta?.action || "unknown",
    activity_id: payload?.data?.id || payload?.current?.id || null,
    subject: payload?.data?.subject || payload?.current?.subject || ""
  };
}

function isAdminAuthorized(req, { username, password }) {
  if (!username || !password) {
    return false;
  }
  const basic = parseBasicAuth(req.header("authorization"));
  return basic.username === username && basic.password === password;
}

function validateSequencesPayload(value) {
  if (!value || typeof value !== "object") {
    return "Payload must be a JSON object.";
  }

  if (!Array.isArray(value.sequences)) {
    return "Payload must include a sequences array.";
  }

  for (const sequence of value.sequences) {
    if (!sequence || typeof sequence !== "object") {
      return "Each sequence must be an object.";
    }
    if (!String(sequence.id || "").trim()) return "Each sequence needs an id.";
    if (!String(sequence.name || "").trim()) return "Each sequence needs a name.";
    if (!Array.isArray(sequence.templates)) return `Sequence '${sequence.id}' must include a templates array.`;

    for (const template of sequence.templates) {
      if (!template || typeof template !== "object") {
        return `Sequence '${sequence.id}' includes an invalid template entry.`;
      }
      if (!String(template.id || "").trim()) return `A template in '${sequence.id}' is missing id.`;
      if (!Number.isFinite(Number(template.stage))) return `Template '${template.id}' in '${sequence.id}' needs numeric stage.`;
      if (!String(template.label || "").trim()) return `Template '${template.id}' in '${sequence.id}' needs label.`;
      if (!String(template.body || "").trim()) return `Template '${template.id}' in '${sequence.id}' needs body.`;
    }
  }

  return "";
}

function writeJsonAtomic(filePath, value) {
  const dir = path.dirname(filePath);
  fs.mkdirSync(dir, { recursive: true });
  const tmpPath = `${filePath}.tmp`;
  fs.writeFileSync(tmpPath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  fs.renameSync(tmpPath, filePath);
}

function renderAdminHtml() {
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Peak Access Admin</title>
    <style>
      :root {
        --bg: #f5f7fb;
        --surface: #ffffff;
        --border: #d9e1ec;
        --text: #1d2a3a;
        --muted: #5f6f86;
        --primary: #4cae4f;
        --primary-dark: #449c47;
        --danger: #b33a3a;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: "Open Sans", "Segoe UI", Arial, sans-serif;
        background: var(--bg);
        color: var(--text);
      }
      .wrap {
        max-width: 1100px;
        margin: 20px auto;
        padding: 0 16px 32px;
      }
      .card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: 10px;
        padding: 14px;
      }
      .head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 8px;
        margin-bottom: 10px;
      }
      h1 {
        margin: 0;
        font-size: 20px;
      }
      .hint {
        color: var(--muted);
        margin: 0 0 10px;
        font-size: 13px;
      }
      textarea {
        width: 100%;
        min-height: 560px;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 10px;
        font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
        font-size: 12px;
        line-height: 1.5;
        resize: vertical;
      }
      .row {
        display: flex;
        gap: 8px;
        align-items: center;
      }
      button {
        border: 1px solid transparent;
        border-radius: 6px;
        padding: 8px 12px;
        font-weight: 700;
        cursor: pointer;
      }
      .btn-primary {
        background: var(--primary);
        border-color: var(--primary-dark);
        color: #fff;
      }
      .btn-secondary {
        background: #fff;
        border-color: var(--border);
        color: var(--text);
      }
      .status {
        margin-top: 10px;
        font-size: 13px;
      }
      .ok { color: #1f7a3b; }
      .err { color: var(--danger); }
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="card">
        <div class="head">
          <h1>Peak Access Template Admin</h1>
          <div class="row">
            <button id="reloadBtn" class="btn-secondary" type="button">Reload</button>
            <button id="formatBtn" class="btn-secondary" type="button">Format JSON</button>
            <button id="saveBtn" class="btn-primary" type="button">Save</button>
          </div>
        </div>
        <p class="hint">Edit shared LinkedIn sequences/templates JSON. Changes are saved to <code>backend/data/sequences.json</code>.</p>
        <textarea id="editor" spellcheck="false"></textarea>
        <div id="status" class="status"></div>
      </div>
    </div>
    <script>
      const editor = document.getElementById("editor");
      const statusEl = document.getElementById("status");
      const reloadBtn = document.getElementById("reloadBtn");
      const formatBtn = document.getElementById("formatBtn");
      const saveBtn = document.getElementById("saveBtn");

      function setStatus(msg, ok = true) {
        statusEl.textContent = msg;
        statusEl.className = "status " + (ok ? "ok" : "err");
      }

      async function loadData() {
        setStatus("Loading...");
        const res = await fetch("/admin/api/sequences", { method: "GET" });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.ok === false) {
          setStatus(data.error || "Failed to load data.", false);
          return;
        }
        editor.value = JSON.stringify(data.data, null, 2);
        setStatus("Loaded.");
      }

      function formatJson() {
        try {
          const parsed = JSON.parse(editor.value);
          editor.value = JSON.stringify(parsed, null, 2);
          setStatus("Formatted.");
        } catch (error) {
          setStatus("Invalid JSON: " + error.message, false);
        }
      }

      async function saveData() {
        let parsed;
        try {
          parsed = JSON.parse(editor.value);
        } catch (error) {
          setStatus("Invalid JSON: " + error.message, false);
          return;
        }

        setStatus("Saving...");
        const res = await fetch("/admin/api/sequences", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(parsed)
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.ok === false) {
          setStatus(data.error || "Save failed.", false);
          return;
        }
        editor.value = JSON.stringify(data.data, null, 2);
        setStatus("Saved.");
      }

      reloadBtn.addEventListener("click", loadData);
      formatBtn.addEventListener("click", formatJson);
      saveBtn.addEventListener("click", saveData);
      loadData();
    </script>
  </body>
</html>`;
}

async function updatePipedriveEligibleFlag({ personId, eligible, baseUrl, apiToken, fieldKey }) {
  const url = new URL(`/api/v1/persons/${personId}`, baseUrl);
  url.searchParams.set("api_token", apiToken);

  const response = await fetch(url.toString(), {
    method: "PUT",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      [fieldKey]: Boolean(eligible)
    })
  });

  if (!response.ok) {
    throw new Error(`Pipedrive update failed (${response.status})`);
  }
}
