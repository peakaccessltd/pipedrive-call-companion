import express from "express";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 8787);
const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET || "change-me";
const PIPEDRIVE_BASE_URL = process.env.PIPEDRIVE_BASE_URL || "";
const PIPEDRIVE_API_TOKEN = process.env.PIPEDRIVE_API_TOKEN || "";
const PERSON_DM_ELIGIBLE_FIELD_KEY = process.env.PERSON_DM_ELIGIBLE_FIELD_KEY || "";
const CALL_DISPOSITION_TRIGGER_OPTION_ID = String(process.env.CALL_DISPOSITION_TRIGGER_OPTION_ID || "6");
const CALL_DISPOSITION_TRIGGER_LABEL = process.env.CALL_DISPOSITION_TRIGGER_LABEL || "LinkedIn Outreach next step";

const sequencesPath = path.join(__dirname, "data", "sequences.json");
const queuePath = path.join(__dirname, "data", "dm-eligible-queue.json");

const app = express();
app.use(express.json({ limit: "1mb" }));

app.get("/health", (_req, res) => {
  res.json({ ok: true, service: "peak-access-linkedin-backend", now: new Date().toISOString() });
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

app.get("/eligible/:personId", (req, res) => {
  const personId = String(req.params.personId || "").trim();
  const queue = readJson(queuePath, { items: {} });
  const item = queue.items?.[personId] || null;

  res.json({ ok: true, person_id: personId, eligible: Boolean(item?.eligible), item });
});

app.post("/pipedrive/webhook", async (req, res) => {
  if (!isWebhookAuthorized(req)) {
    res.status(401).json({ ok: false, error: "Unauthorized" });
    return;
  }

  const payload = req.body || {};
  const personId = extractPersonId(payload);

  if (!personId) {
    res.status(200).json({ ok: true, ignored: true, reason: "No person_id in payload" });
    return;
  }

  if (!payloadHasTriggerOption(payload)) {
    res.status(200).json({ ok: true, ignored: true, reason: "Trigger option not present" });
    return;
  }

  markEligibleInQueue(personId, payload);

  // Optional passthrough: mark a boolean custom field in Pipedrive when env vars are configured.
  if (PIPEDRIVE_BASE_URL && PIPEDRIVE_API_TOKEN && PERSON_DM_ELIGIBLE_FIELD_KEY) {
    try {
      await updatePipedriveEligibleFlag(personId, true);
    } catch (error) {
      console.error("Failed to update Pipedrive eligible flag", error.message);
    }
  }

  res.status(200).json({ ok: true, person_id: personId, eligible: true });
});

app.listen(PORT, () => {
  console.log(`peak-access-linkedin-backend listening on http://localhost:${PORT}`);
});

function isWebhookAuthorized(req) {
  const provided = String(req.header("x-peak-access-secret") || "").trim();
  return Boolean(provided) && provided === WEBHOOK_SECRET;
}

function readJson(filePath, fallback) {
  try {
    return JSON.parse(fs.readFileSync(filePath, "utf8"));
  } catch (_error) {
    return fallback;
  }
}

function writeJson(filePath, data) {
  fs.writeFileSync(filePath, `${JSON.stringify(data, null, 2)}\n`, "utf8");
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

function payloadHasTriggerOption(payload) {
  const candidateValues = [
    payload?.data,
    payload?.current,
    payload
  ];

  for (const value of candidateValues) {
    if (!value || typeof value !== "object") continue;

    const matches = flattenValues(value).map((item) => String(item));
    const hasOptionId = matches.some((item) => item === CALL_DISPOSITION_TRIGGER_OPTION_ID);
    const hasLabel = matches.some((item) => item.toLowerCase().includes(CALL_DISPOSITION_TRIGGER_LABEL.toLowerCase()));

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

function markEligibleInQueue(personId, payload) {
  const queue = readJson(queuePath, { items: {} });
  queue.items = queue.items || {};
  queue.items[personId] = {
    eligible: true,
    source: "pipedrive_webhook",
    received_at: new Date().toISOString(),
    payload_summary: summarizePayload(payload)
  };

  writeJson(queuePath, queue);
}

function summarizePayload(payload) {
  return {
    event: payload?.event || payload?.meta?.action || "unknown",
    activity_id: payload?.data?.id || payload?.current?.id || null,
    subject: payload?.data?.subject || payload?.current?.subject || ""
  };
}

async function updatePipedriveEligibleFlag(personId, eligible) {
  const url = new URL(`/api/v1/persons/${personId}`, PIPEDRIVE_BASE_URL);
  url.searchParams.set("api_token", PIPEDRIVE_API_TOKEN);

  const response = await fetch(url.toString(), {
    method: "PUT",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      [PERSON_DM_ELIGIBLE_FIELD_KEY]: Boolean(eligible)
    })
  });

  if (!response.ok) {
    throw new Error(`Pipedrive update failed (${response.status})`);
  }
}
