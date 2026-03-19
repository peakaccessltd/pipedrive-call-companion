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
  const CONFIG_SYNC_SECRET = String(process.env.CONFIG_SYNC_SECRET || WEBHOOK_SECRET || "").trim();

  const sequencesPath = path.join(baseDir, "data", "sequences.json");
  const extensionConfigPath = path.join(baseDir, "data", "extension-config.json");
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

  app.get("/admin/api/extension-config", (req, res) => {
    if (!isAdminAuthorized(req, { username: ADMIN_USERNAME, password: ADMIN_PASSWORD })) {
      res.status(401).json({ ok: false, error: "Unauthorized" });
      return;
    }

    const data = readJson(extensionConfigPath, normalizeExtensionConfigPayload({}));
    res.json({ ok: true, data });
  });

  app.put("/admin/api/extension-config", (req, res) => {
    if (!isAdminAuthorized(req, { username: ADMIN_USERNAME, password: ADMIN_PASSWORD })) {
      res.status(401).json({ ok: false, error: "Unauthorized" });
      return;
    }

    const payload = normalizeExtensionConfigPayload(req.body || {});
    if (!payload.backendBaseUrl) {
      res.status(400).json({ ok: false, error: "backendBaseUrl is required." });
      return;
    }

    try {
      const nextValue = {
        ...payload,
        updated_at: new Date().toISOString()
      };
      writeJsonAtomic(extensionConfigPath, nextValue);
      res.json({ ok: true, data: nextValue });
    } catch (error) {
      res.status(500).json({ ok: false, error: `Failed to save extension config: ${error.message || String(error)}` });
    }
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

  app.get("/extension-config", (_req, res) => {
    if (!isConfigSyncAuthorized(_req, CONFIG_SYNC_SECRET)) {
      res.status(401).json({ ok: false, error: "Unauthorized" });
      return;
    }
    const data = readJson(extensionConfigPath, null);
    res.json({ ok: true, data });
  });

  app.put("/extension-config", (req, res) => {
    if (!isConfigSyncAuthorized(req, CONFIG_SYNC_SECRET)) {
      res.status(401).json({ ok: false, error: "Unauthorized" });
      return;
    }
    const payload = normalizeExtensionConfigPayload(req.body || {});
    if (!payload.backendBaseUrl) {
      res.status(400).json({ ok: false, error: "backendBaseUrl is required." });
      return;
    }

    try {
      writeJsonAtomic(extensionConfigPath, {
        ...payload,
        updated_at: new Date().toISOString()
      });
      res.json({ ok: true, data: payload });
    } catch (error) {
      res.status(500).json({ ok: false, error: `Failed to save extension config: ${error.message || String(error)}` });
    }
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

function isConfigSyncAuthorized(req, configSyncSecret) {
  if (!configSyncSecret) {
    return false;
  }

  const headerSecret = String(req.header("x-peak-access-secret") || "").trim();
  if (headerSecret && headerSecret === configSyncSecret) {
    return true;
  }

  const querySecret = String(req.query?.secret || "").trim();
  if (querySecret && querySecret === configSyncSecret) {
    return true;
  }

  return false;
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

function normalizeExtensionConfigPayload(value) {
  return {
    apiToken: String(value.apiToken || "").trim(),
    backendBaseUrl: String(value.backendBaseUrl || "").trim().replace(/\/+$/, ""),
    personLinkedinProfileUrlKey: String(value.personLinkedinProfileUrlKey || "").trim(),
    personLinkedinDmSequenceIdKey: String(value.personLinkedinDmSequenceIdKey || "").trim(),
    personLinkedinDmStageKey: String(value.personLinkedinDmStageKey || "").trim(),
    personLinkedinDmLastSentAtKey: String(value.personLinkedinDmLastSentAtKey || "").trim(),
    personLinkedinDmEligibleKey: String(value.personLinkedinDmEligibleKey || "").trim(),
    callDispositionFieldKey: String(value.callDispositionFieldKey || "").trim(),
    callDispositionTriggerOptionId: String(value.callDispositionTriggerOptionId || "").trim(),
    callDispositionTriggerOptionLabel: String(value.callDispositionTriggerOptionLabel || "").trim(),
    emailTemplatesByStage: String(value.emailTemplatesByStage || "").trim(),
    autoOpenPanel: Boolean(value.autoOpenPanel),
    showNotes: Boolean(value.showNotes),
    showActivities: Boolean(value.showActivities)
  };
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
        --soft: #f8fafc;
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
      .row {
        display: flex;
        gap: 8px;
        align-items: center;
        flex-wrap: wrap;
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
      .btn-danger {
        background: #fff;
        border-color: #e7c6c6;
        color: var(--danger);
      }
      .status {
        margin-top: 10px;
        font-size: 13px;
      }
      .ok { color: #1f7a3b; }
      .err { color: var(--danger); }
      .sequences {
        display: grid;
        gap: 12px;
        margin-top: 12px;
      }
      .sequence-card {
        border: 1px solid var(--border);
        border-radius: 10px;
        background: var(--surface);
        padding: 12px;
      }
      .sequence-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 10px;
      }
      .sequence-head-left {
        display: flex;
        align-items: center;
        gap: 8px;
        min-width: 0;
      }
      .sequence-title {
        margin: 0;
        font-size: 16px;
        cursor: pointer;
      }
      .sequence-toggle {
        width: 26px;
        height: 26px;
        border-radius: 6px;
        border: 1px solid var(--border);
        background: #fff;
        color: var(--text);
        font-weight: 800;
        cursor: pointer;
        padding: 0;
      }
      .sequence-body {
        display: grid;
        gap: 0;
      }
      .sequence-body-collapsed {
        display: none;
      }
      .grid {
        display: grid;
        gap: 8px;
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
      .field {
        display: grid;
        gap: 4px;
      }
      .field-full {
        grid-column: 1 / -1;
      }
      label {
        font-size: 12px;
        color: var(--muted);
        font-weight: 700;
      }
      input, textarea {
        width: 100%;
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 8px;
        font-size: 13px;
        font-family: "Open Sans", "Segoe UI", Arial, sans-serif;
      }
      textarea {
        min-height: 120px;
        resize: vertical;
        line-height: 1.5;
      }
      .templates {
        margin-top: 10px;
        display: grid;
        gap: 8px;
      }
      .template-card {
        border: 1px solid var(--border);
        border-radius: 8px;
        background: var(--soft);
        padding: 10px;
      }
      .template-head {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 8px;
      }
      .template-title {
        margin: 0;
        font-size: 14px;
      }
      .actions {
        display: flex;
        gap: 8px;
        margin-top: 8px;
        flex-wrap: wrap;
      }
      .small {
        font-size: 12px;
        color: var(--muted);
      }
      .stack {
        display: grid;
        gap: 14px;
      }
      .divider {
        height: 1px;
        background: var(--border);
        margin: 10px 0;
      }
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="stack">
        <div class="card">
          <div class="head">
            <h1>Peak Access Admin</h1>
            <div class="row">
              <button id="reloadBtn" class="btn-secondary" type="button">Reload</button>
              <button id="addSequenceBtn" class="btn-secondary" type="button">Add Sequence</button>
              <button id="saveBtn" class="btn-primary" type="button">Save Sequences</button>
            </div>
          </div>
          <p class="hint">Manage shared LinkedIn sequences and stage messages here.</p>
          <div id="sequences" class="sequences"></div>
          <div id="status" class="status"></div>
        </div>
        <div class="card">
          <div class="head">
            <h1>Extension Config</h1>
            <div class="row">
              <button id="reloadConfigBtn" class="btn-secondary" type="button">Reload Config</button>
              <button id="saveConfigBtn" class="btn-primary" type="button">Save Config</button>
            </div>
          </div>
          <p class="hint">This is the central companion config. The extension can restore from this backend config using the sync secret.</p>
          <div class="grid">
            <div class="field"><label for="cfgApiToken">Pipedrive API token</label><input id="cfgApiToken" type="text" /></div>
            <div class="field"><label for="cfgBackendBaseUrl">Backend base URL</label><input id="cfgBackendBaseUrl" type="text" /></div>
            <div class="field"><label for="cfgProfileUrlKey">LinkedIn Profile URL key</label><input id="cfgProfileUrlKey" type="text" /></div>
            <div class="field"><label for="cfgSequenceIdKey">DM Sequence ID key</label><input id="cfgSequenceIdKey" type="text" /></div>
            <div class="field"><label for="cfgStageKey">DM Stage key</label><input id="cfgStageKey" type="text" /></div>
            <div class="field"><label for="cfgLastSentKey">DM Last Sent key</label><input id="cfgLastSentKey" type="text" /></div>
            <div class="field"><label for="cfgEligibleKey">DM Eligible key</label><input id="cfgEligibleKey" type="text" /></div>
            <div class="field"><label for="cfgDispositionFieldKey">Call Disposition field key</label><input id="cfgDispositionFieldKey" type="text" /></div>
            <div class="field"><label for="cfgDispositionOptionId">Trigger option ID</label><input id="cfgDispositionOptionId" type="text" /></div>
            <div class="field field-full"><label for="cfgEmailTemplatesByStage">No Answer email templates JSON</label><textarea id="cfgEmailTemplatesByStage"></textarea></div>
            <label class="row"><input id="cfgAutoOpenPanel" type="checkbox" /> Auto-open panel</label>
            <label class="row"><input id="cfgShowNotes" type="checkbox" /> Show notes</label>
            <label class="row"><input id="cfgShowActivities" type="checkbox" /> Show activities</label>
          </div>
          <div id="configStatus" class="status"></div>
        </div>
      </div>
    </div>
    <script>
      const sequencesRoot = document.getElementById("sequences");
      const statusEl = document.getElementById("status");
      const configStatusEl = document.getElementById("configStatus");
      const reloadBtn = document.getElementById("reloadBtn");
      const addSequenceBtn = document.getElementById("addSequenceBtn");
      const saveBtn = document.getElementById("saveBtn");
      const reloadConfigBtn = document.getElementById("reloadConfigBtn");
      const saveConfigBtn = document.getElementById("saveConfigBtn");
      const configRefs = {
        apiToken: document.getElementById("cfgApiToken"),
        backendBaseUrl: document.getElementById("cfgBackendBaseUrl"),
        personLinkedinProfileUrlKey: document.getElementById("cfgProfileUrlKey"),
        personLinkedinDmSequenceIdKey: document.getElementById("cfgSequenceIdKey"),
        personLinkedinDmStageKey: document.getElementById("cfgStageKey"),
        personLinkedinDmLastSentAtKey: document.getElementById("cfgLastSentKey"),
        personLinkedinDmEligibleKey: document.getElementById("cfgEligibleKey"),
        callDispositionFieldKey: document.getElementById("cfgDispositionFieldKey"),
        callDispositionTriggerOptionId: document.getElementById("cfgDispositionOptionId"),
        emailTemplatesByStage: document.getElementById("cfgEmailTemplatesByStage"),
        autoOpenPanel: document.getElementById("cfgAutoOpenPanel"),
        showNotes: document.getElementById("cfgShowNotes"),
        showActivities: document.getElementById("cfgShowActivities")
      };
      let state = { version: 1, updated_at: null, sequences: [] };
      let localCounter = 0;
      const expandedSequences = Object.create(null);

      function setStatus(msg, ok = true) {
        statusEl.textContent = msg;
        statusEl.className = "status " + (ok ? "ok" : "err");
      }

      function setConfigStatus(msg, ok = true) {
        configStatusEl.textContent = msg;
        configStatusEl.className = "status " + (ok ? "ok" : "err");
      }

      function slugify(value) {
        return String(value || "")
          .toLowerCase()
          .replace(/[^a-z0-9]+/g, "-")
          .replace(/^-+|-+$/g, "")
          .slice(0, 48);
      }

      function nextLocalId(prefix) {
        localCounter += 1;
        return prefix + "-" + Date.now() + "-" + localCounter;
      }

      function ensureTemplateMinimum(sequence) {
        const templates = Array.isArray(sequence.templates) ? sequence.templates : [];
        if (templates.length >= 3) return templates;

        const defaults = [
          { stage: 1, label: "Stage 1 Intro", body: "Hi {{personFirstName}}, thanks for connecting..." },
          { stage: 2, label: "Stage 2 Follow-up", body: "Hi {{personFirstName}}, following up with one practical idea..." },
          { stage: 3, label: "Stage 3 Close Loop", body: "Hi {{personFirstName}}, closing the loop on our prior note..." }
        ];

        for (let i = templates.length; i < 3; i += 1) {
          const d = defaults[i];
          templates.push({
            id: nextLocalId("template"),
            stage: d.stage,
            label: d.label,
            body: d.body
          });
        }
        return templates;
      }

      function getSequenceUiKey(sequence, index) {
        if (sequence && sequence._uiKey) return sequence._uiKey;
        const base = String(sequence?.id || "").trim() || ("sequence-" + index);
        if (sequence) {
          sequence._uiKey = base + "-" + nextLocalId("ui");
          return sequence._uiKey;
        }
        return base + "-" + nextLocalId("ui");
      }

      function setSequenceExpanded(uiKey, expanded) {
        expandedSequences[uiKey] = Boolean(expanded);
      }

      function isSequenceExpanded(uiKey) {
        return Boolean(expandedSequences[uiKey]);
      }

      function renderAll() {
        sequencesRoot.innerHTML = "";
        if (!Array.isArray(state.sequences) || state.sequences.length === 0) {
          const empty = document.createElement("div");
          empty.className = "small";
          empty.textContent = "No sequences yet. Click Add Sequence.";
          sequencesRoot.appendChild(empty);
          return;
        }

        state.sequences.forEach((sequence, sequenceIndex) => {
          ensureTemplateMinimum(sequence);
          const uiKey = getSequenceUiKey(sequence, sequenceIndex);
          if (expandedSequences[uiKey] === undefined) {
            setSequenceExpanded(uiKey, false);
          }
          const card = document.createElement("section");
          card.className = "sequence-card";

          const head = document.createElement("div");
          head.className = "sequence-head";
          const headLeft = document.createElement("div");
          headLeft.className = "sequence-head-left";
          const toggle = document.createElement("button");
          toggle.type = "button";
          toggle.className = "sequence-toggle";
          toggle.textContent = isSequenceExpanded(uiKey) ? "-" : "+";
          const h = document.createElement("h2");
          h.className = "sequence-title";
          h.textContent = sequence.name || "Untitled Sequence";
          h.title = "Click to expand/collapse";
          const toggleExpanded = () => {
            setSequenceExpanded(uiKey, !isSequenceExpanded(uiKey));
            renderAll();
          };
          toggle.addEventListener("click", (event) => {
            event.preventDefault();
            event.stopPropagation();
            toggleExpanded();
          });
          h.addEventListener("click", (event) => {
            event.preventDefault();
            toggleExpanded();
          });
          headLeft.appendChild(toggle);
          headLeft.appendChild(h);
          const removeSequenceBtn = document.createElement("button");
          removeSequenceBtn.type = "button";
          removeSequenceBtn.className = "btn-danger";
          removeSequenceBtn.textContent = "Remove Sequence";
          removeSequenceBtn.addEventListener("click", () => {
            state.sequences.splice(sequenceIndex, 1);
            renderAll();
          });
          head.appendChild(headLeft);
          head.appendChild(removeSequenceBtn);
          card.appendChild(head);

          const bodyWrap = document.createElement("div");
          bodyWrap.className = "sequence-body" + (isSequenceExpanded(uiKey) ? "" : " sequence-body-collapsed");
          const grid = document.createElement("div");
          grid.className = "grid";

          grid.appendChild(makeField("Sequence ID", sequence.id || "", (v) => {
            sequence.id = v;
          }));
          grid.appendChild(makeField("Name", sequence.name || "", (v) => {
            sequence.name = v;
            h.textContent = v || "Untitled Sequence";
          }));
          grid.appendChild(makeField("Recommended Start Stage", String(sequence.recommended_start_stage || 1), (v) => {
            sequence.recommended_start_stage = Number(v || 1);
          }, "number"));
          grid.appendChild(makeField("Description", sequence.description || "", (v) => {
            sequence.description = v;
          }));

          bodyWrap.appendChild(grid);
          const divider = document.createElement("div");
          divider.className = "divider";
          bodyWrap.appendChild(divider);

          const templatesWrap = document.createElement("div");
          templatesWrap.className = "templates";
          sequence.templates.forEach((template, templateIndex) => {
            templatesWrap.appendChild(renderTemplateCard(sequence, template, sequenceIndex, templateIndex));
          });
          bodyWrap.appendChild(templatesWrap);

          const actions = document.createElement("div");
          actions.className = "actions";
          const addTemplateBtn = document.createElement("button");
          addTemplateBtn.type = "button";
          addTemplateBtn.className = "btn-secondary";
          addTemplateBtn.textContent = "Add Stage/Message";
          addTemplateBtn.addEventListener("click", () => {
            const nextStage = sequence.templates.reduce((max, t) => Math.max(max, Number(t.stage || 0)), 0) + 1;
            sequence.templates.push({
              id: nextLocalId("template"),
              stage: nextStage,
              label: "Stage " + nextStage + " Message",
              body: "Hi {{personFirstName}}, ..."
            });
            renderAll();
          });
          actions.appendChild(addTemplateBtn);
          bodyWrap.appendChild(actions);
          card.appendChild(bodyWrap);

          sequencesRoot.appendChild(card);
        });
      }

      function renderTemplateCard(sequence, template, _sequenceIndex, templateIndex) {
        const card = document.createElement("article");
        card.className = "template-card";

        const head = document.createElement("div");
        head.className = "template-head";
        const title = document.createElement("h3");
        title.className = "template-title";
        title.textContent = template.label || ("Template " + (templateIndex + 1));
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "btn-danger";
        removeBtn.textContent = "Remove";
        removeBtn.addEventListener("click", () => {
          sequence.templates.splice(templateIndex, 1);
          renderAll();
        });

        const duplicateBtn = document.createElement("button");
        duplicateBtn.type = "button";
        duplicateBtn.className = "btn-secondary";
        duplicateBtn.textContent = "Duplicate";
        duplicateBtn.addEventListener("click", () => {
          const nextStage = sequence.templates.reduce((max, t) => Math.max(max, Number(t.stage || 0)), 0) + 1;
          sequence.templates.splice(templateIndex + 1, 0, {
            id: nextLocalId("template"),
            stage: nextStage,
            label: String(template.label || "Message") + " Copy",
            body: String(template.body || "")
          });
          renderAll();
        });

        const moveUpBtn = document.createElement("button");
        moveUpBtn.type = "button";
        moveUpBtn.className = "btn-secondary";
        moveUpBtn.textContent = "Up";
        moveUpBtn.disabled = templateIndex === 0;
        moveUpBtn.addEventListener("click", () => {
          if (templateIndex <= 0) return;
          const current = sequence.templates[templateIndex];
          sequence.templates[templateIndex] = sequence.templates[templateIndex - 1];
          sequence.templates[templateIndex - 1] = current;
          renderAll();
        });

        const moveDownBtn = document.createElement("button");
        moveDownBtn.type = "button";
        moveDownBtn.className = "btn-secondary";
        moveDownBtn.textContent = "Down";
        moveDownBtn.disabled = templateIndex >= sequence.templates.length - 1;
        moveDownBtn.addEventListener("click", () => {
          if (templateIndex >= sequence.templates.length - 1) return;
          const current = sequence.templates[templateIndex];
          sequence.templates[templateIndex] = sequence.templates[templateIndex + 1];
          sequence.templates[templateIndex + 1] = current;
          renderAll();
        });

        const actions = document.createElement("div");
        actions.className = "actions";
        actions.appendChild(moveUpBtn);
        actions.appendChild(moveDownBtn);
        actions.appendChild(duplicateBtn);
        actions.appendChild(removeBtn);

        head.appendChild(title);
        head.appendChild(actions);
        card.appendChild(head);

        const fields = document.createElement("div");
        fields.className = "grid";

        fields.appendChild(makeField("Template ID", template.id || "", (v) => {
          template.id = v;
        }));
        fields.appendChild(makeField("Stage", String(template.stage || 1), (v) => {
          template.stage = Number(v || 1);
        }, "number"));
        fields.appendChild(makeField("Title", template.label || "", (v) => {
          template.label = v;
          title.textContent = v || ("Template " + (templateIndex + 1));
        }, "text", true));
        card.appendChild(fields);

        const bodyField = document.createElement("div");
        bodyField.className = "field field-full";
        const bodyLabel = document.createElement("label");
        bodyLabel.textContent = "Message";
        const bodyInput = document.createElement("textarea");
        bodyInput.value = template.body || "";
        bodyInput.addEventListener("input", (e) => {
          template.body = e.target.value;
        });
        bodyField.appendChild(bodyLabel);
        bodyField.appendChild(bodyInput);
        card.appendChild(bodyField);

        const linkRow = document.createElement("div");
        linkRow.className = "row";
        const linkInput = document.createElement("input");
        linkInput.type = "text";
        linkInput.placeholder = "https://example.com";
        const insertLinkBtn = document.createElement("button");
        insertLinkBtn.type = "button";
        insertLinkBtn.className = "btn-secondary";
        insertLinkBtn.textContent = "Insert Link";
        insertLinkBtn.addEventListener("click", () => {
          const link = String(linkInput.value || "").trim();
          if (!link) return;
          const existing = String(template.body || "");
          template.body = existing + (existing.endsWith("\\n") || existing.length === 0 ? "" : "\\n") + link;
          bodyInput.value = template.body;
          linkInput.value = "";
        });
        linkRow.appendChild(linkInput);
        linkRow.appendChild(insertLinkBtn);
        card.appendChild(linkRow);

        return card;
      }

      function makeField(labelText, value, onInput, type = "text", full = false) {
        const wrapper = document.createElement("div");
        wrapper.className = "field" + (full ? " field-full" : "");
        const label = document.createElement("label");
        label.textContent = labelText;
        const input = document.createElement("input");
        input.type = type;
        input.value = value;
        input.addEventListener("input", (e) => onInput(e.target.value));
        wrapper.appendChild(label);
        wrapper.appendChild(input);
        return wrapper;
      }

      function normalizeForSave() {
        const sequences = (state.sequences || []).map((sequence) => {
          const safeName = String(sequence.name || "").trim() || "New Sequence";
          const safeId = String(sequence.id || "").trim() || slugify(safeName) || nextLocalId("sequence");
          const templates = (sequence.templates || []).map((template, idx) => {
            const label = String(template.label || "").trim() || ("Stage " + Number(template.stage || idx + 1));
            const stage = Number(template.stage || idx + 1);
            const id = String(template.id || "").trim()
              || slugify(safeId + "-" + label)
              || nextLocalId("template");
            return {
              id,
              stage,
              label,
              body: String(template.body || "").trim()
            };
          }).filter((t) => t.body);

          return {
            id: safeId,
            name: safeName,
            description: String(sequence.description || "").trim(),
            recommended_start_stage: Number(sequence.recommended_start_stage || 1),
            templates
          };
        }).filter((s) => s.templates.length > 0);

        return {
          version: Number(state.version || 1),
          updated_at: state.updated_at || null,
          sequences
        };
      }

      async function loadData() {
        setStatus("Loading...");
        const res = await fetch("/admin/api/sequences", { method: "GET" });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.ok === false) {
          setStatus(data.error || "Failed to load data.", false);
          return;
        }
        state = data.data || { version: 1, updated_at: null, sequences: [] };
        if (!Array.isArray(state.sequences)) {
          state.sequences = [];
        }
        state.sequences.forEach((sequence) => ensureTemplateMinimum(sequence));
        renderAll();
        setStatus("Loaded.");
      }

      async function loadConfigData() {
        setConfigStatus("Loading...");
        const res = await fetch("/admin/api/extension-config", { method: "GET" });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.ok === false) {
          setConfigStatus(data.error || "Failed to load config.", false);
          return;
        }
        applyConfigToForm(data.data || {});
        setConfigStatus("Loaded.");
      }

      async function saveData() {
        const parsed = normalizeForSave();

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
        state = data.data;
        renderAll();
        setStatus("Saved.");
      }

      async function saveConfigData() {
        const payload = getConfigPayload();
        const validationError = validateConfigPayload(payload);
        if (validationError) {
          setConfigStatus(validationError, false);
          return;
        }

        setConfigStatus("Saving...");
        const res = await fetch("/admin/api/extension-config", {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        const data = await res.json().catch(() => ({}));
        if (!res.ok || data.ok === false) {
          setConfigStatus(data.error || "Failed to save config.", false);
          return;
        }
        applyConfigToForm(data.data || {});
        setConfigStatus("Saved.");
      }

      function addSequence() {
        if (!Array.isArray(state.sequences)) {
          state.sequences = [];
        }
        const sequenceNumber = (state.sequences?.length || 0) + 1;
        const sequenceId = nextLocalId("sequence");
        const nextSequence = {
          id: sequenceId,
          name: "Sequence " + sequenceNumber,
          description: "",
          recommended_start_stage: 1,
          templates: [
            { id: nextLocalId("template"), stage: 1, label: "Stage 1 Intro", body: "Hi {{personFirstName}}, thanks for connecting..." },
            { id: nextLocalId("template"), stage: 2, label: "Stage 2 Follow-up", body: "Hi {{personFirstName}}, following up..." },
            { id: nextLocalId("template"), stage: 3, label: "Stage 3 Close Loop", body: "Hi {{personFirstName}}, closing the loop..." }
          ]
        };
        const uiKey = getSequenceUiKey(nextSequence, state.sequences.length);
        setSequenceExpanded(uiKey, true);
        state.sequences.push(nextSequence);
        renderAll();
        setStatus("Added new sequence.");
      }

      function applyConfigToForm(value) {
        configRefs.apiToken.value = value.apiToken || "";
        configRefs.backendBaseUrl.value = value.backendBaseUrl || "";
        configRefs.personLinkedinProfileUrlKey.value = value.personLinkedinProfileUrlKey || "";
        configRefs.personLinkedinDmSequenceIdKey.value = value.personLinkedinDmSequenceIdKey || "";
        configRefs.personLinkedinDmStageKey.value = value.personLinkedinDmStageKey || "";
        configRefs.personLinkedinDmLastSentAtKey.value = value.personLinkedinDmLastSentAtKey || "";
        configRefs.personLinkedinDmEligibleKey.value = value.personLinkedinDmEligibleKey || "";
        configRefs.callDispositionFieldKey.value = value.callDispositionFieldKey || "";
        configRefs.callDispositionTriggerOptionId.value = value.callDispositionTriggerOptionId || "";
        configRefs.emailTemplatesByStage.value = value.emailTemplatesByStage || "";
        configRefs.autoOpenPanel.checked = Boolean(value.autoOpenPanel);
        configRefs.showNotes.checked = Boolean(value.showNotes);
        configRefs.showActivities.checked = Boolean(value.showActivities);
      }

      function getConfigPayload() {
        return {
          apiToken: String(configRefs.apiToken.value || "").trim(),
          backendBaseUrl: String(configRefs.backendBaseUrl.value || "").trim(),
          personLinkedinProfileUrlKey: String(configRefs.personLinkedinProfileUrlKey.value || "").trim(),
          personLinkedinDmSequenceIdKey: String(configRefs.personLinkedinDmSequenceIdKey.value || "").trim(),
          personLinkedinDmStageKey: String(configRefs.personLinkedinDmStageKey.value || "").trim(),
          personLinkedinDmLastSentAtKey: String(configRefs.personLinkedinDmLastSentAtKey.value || "").trim(),
          personLinkedinDmEligibleKey: String(configRefs.personLinkedinDmEligibleKey.value || "").trim(),
          callDispositionFieldKey: String(configRefs.callDispositionFieldKey.value || "").trim(),
          callDispositionTriggerOptionId: String(configRefs.callDispositionTriggerOptionId.value || "").trim(),
          emailTemplatesByStage: String(configRefs.emailTemplatesByStage.value || "").trim(),
          autoOpenPanel: configRefs.autoOpenPanel.checked,
          showNotes: configRefs.showNotes.checked,
          showActivities: configRefs.showActivities.checked
        };
      }

      function validateConfigPayload(payload) {
        if (!payload.backendBaseUrl) {
          return "Backend base URL is required.";
        }
        if (payload.emailTemplatesByStage) {
          try {
            JSON.parse(payload.emailTemplatesByStage);
          } catch (_error) {
            return "No Answer email template JSON is invalid.";
          }
        }
        return "";
      }

      reloadBtn.addEventListener("click", loadData);
      addSequenceBtn.addEventListener("click", addSequence);
      saveBtn.addEventListener("click", saveData);
      reloadConfigBtn.addEventListener("click", loadConfigData);
      saveConfigBtn.addEventListener("click", saveConfigData);
      loadData();
      loadConfigData();
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
