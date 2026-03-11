const refs = {
  refreshBtn: document.getElementById("refreshBtn"),
  linkedinContext: document.getElementById("linkedinContext"),
  matchCard: document.getElementById("matchCard"),
  matchCandidates: document.getElementById("matchCandidates"),
  templateSelect: document.getElementById("templateSelect"),
  stageInput: document.getElementById("stageInput"),
  templatePreview: document.getElementById("templatePreview"),
  useTemplateBtn: document.getElementById("useTemplateBtn"),
  logAdvanceBtn: document.getElementById("logAdvanceBtn"),
  status: document.getElementById("status")
};

const state = {
  linkedinContext: null,
  match: null,
  sequences: [],
  templates: [],
  selectedSequenceId: "",
  selectedTemplateId: "",
  stage: 1
};

refs.refreshBtn.addEventListener("click", refreshAll);
refs.templateSelect.addEventListener("change", onTemplateChanged);
refs.stageInput.addEventListener("change", onStageChanged);
refs.useTemplateBtn.addEventListener("click", onInsertTemplate);
refs.logAdvanceBtn.addEventListener("click", onLogAndAdvance);

initTemplateSelect();
refs.stageInput.value = String(state.stage);
refreshAll();

async function refreshAll() {
  setStatus("Loading LinkedIn context...");

  const contextResp = await sendRuntimeMessage({ type: "LINKEDIN_GET_CONTEXT" });
  if (!contextResp.ok) {
    setStatus(contextResp.error || "Failed to detect LinkedIn context", true);
    return;
  }

  state.linkedinContext = contextResp.data;
  renderLinkedInContext();

  const matchResp = await sendRuntimeMessage({
    type: "LINKEDIN_MATCH_PERSON",
    payload: {
      profileUrl: state.linkedinContext.profileUrl,
      profileName: state.linkedinContext.profileName,
      emailHint: state.linkedinContext.emailHint
    }
  });

  if (!matchResp.ok) {
    setStatus(matchResp.error || "Could not match person", true);
  } else {
    state.match = matchResp.data;
    if (Number.isFinite(Number(state.match?.currentStage)) && Number(state.match.currentStage) > 0) {
      state.stage = Number(state.match.currentStage);
      refs.stageInput.value = String(state.stage);
    }
    renderMatchCard();
  }

  await loadSequences();
  await loadTemplates();
  setStatus("LinkedIn Mode ready.", false, true);
}

function initTemplateSelect() {
  refs.templateSelect.innerHTML = "";
}

function renderLinkedInContext() {
  const context = state.linkedinContext || {};
  renderKeyValueCard(refs.linkedinContext, [
    { label: "Profile URL", value: context.profileUrl || "N/A" },
    { label: "Profile Name", value: context.profileName || "N/A" },
    { label: "Messaging View", value: context.isMessaging ? "Yes" : "No" },
    { label: "Thread", value: context.threadTitle || "N/A" }
  ]);
}

function renderMatchCard() {
  const match = state.match || {};

  if (!match.person) {
    renderKeyValueCard(refs.matchCard, [
      { label: "Match", value: "No direct match yet" },
      { label: "Strategy", value: match.strategy || "none" }
    ]);
  } else {
    renderKeyValueCard(refs.matchCard, [
      { label: "Matched", value: `${match.person.name} (#${match.person.id})` },
      { label: "Org", value: match.person.orgName || "N/A" },
      { label: "DM eligible", value: match.dmEligible ? "Yes" : "No" },
      { label: "Current stage", value: String(match.currentStage || 1) }
    ]);
  }

  refs.matchCandidates.innerHTML = "";
  const candidates = Array.isArray(match.candidates) ? match.candidates : [];

  candidates.forEach((candidate) => {
    const card = document.createElement("div");
    card.className = "sp-template";
    card.textContent = `${candidate.name} (#${candidate.id})${candidate.orgName ? ` - ${candidate.orgName}` : ""}`;

    const confirmBtn = document.createElement("button");
    confirmBtn.className = "sp-btn sp-btn-secondary";
    confirmBtn.type = "button";
    confirmBtn.textContent = "Confirm Match + Save LinkedIn URL";
    confirmBtn.addEventListener("click", async () => {
      const resp = await sendRuntimeMessage({
        type: "LINKEDIN_CONFIRM_MATCH",
        payload: {
          personId: candidate.id,
          profileUrl: state.linkedinContext?.profileUrl || ""
        }
      });

      if (!resp.ok) {
        setStatus(resp.error || "Match confirm failed", true);
        return;
      }

      state.match = resp.data;
      renderMatchCard();
      setStatus("Match confirmed and LinkedIn URL saved.", false, true);
    });

    card.appendChild(confirmBtn);
    refs.matchCandidates.appendChild(card);
  });
}

function onTemplateChanged() {
  state.selectedSequenceId = refs.templateSelect.value || "";
  loadTemplates();
}

function onStageChanged() {
  const stage = Number(refs.stageInput.value || 1);
  state.stage = Number.isFinite(stage) && stage > 0 ? stage : 1;
  refs.stageInput.value = String(state.stage);
  loadTemplates();
}

function getSelectedTemplateText() {
  const template = state.templates.find((item) => item.id === state.selectedTemplateId) || state.templates[0];
  return interpolateTemplate(template?.body || "");
}

async function onInsertTemplate() {
  const text = getSelectedTemplateText().trim();
  if (!text) {
    setStatus("No template text available.", true);
    return;
  }

  const resp = await sendRuntimeMessage({
    type: "LINKEDIN_INSERT_TEMPLATE",
    payload: { text }
  });

  if (!resp.ok) {
    setStatus(resp.error || "Failed to insert template", true);
    return;
  }

  if (resp.data.inserted) {
    setStatus("Template inserted into LinkedIn composer.", false, true);
  } else if (resp.data.copied) {
    setStatus(resp.data.message || "Composer unavailable; copied to clipboard.", false, true);
  } else {
    setStatus(resp.data.message || "Could not insert template.", true);
  }
}

async function onLogAndAdvance() {
  if (!state.match?.person?.id) {
    setStatus("No matched Pipedrive person. Confirm a match first.", true);
    return;
  }

  const composerResp = await sendRuntimeMessage({ type: "LINKEDIN_GET_COMPOSER_TEXT" });
  const composerText = composerResp.ok ? composerResp.data.text : "";
  const dmText = composerText || getSelectedTemplateText();

  if (!dmText) {
    setStatus("No DM text found to log.", true);
    return;
  }

  const resp = await sendRuntimeMessage({
    type: "LINKEDIN_LOG_AND_ADVANCE",
    payload: {
      personId: state.match.person.id,
      profileUrl: state.linkedinContext?.profileUrl || "",
      sequenceId: state.selectedSequenceId || "manual_template_flow",
      templateId: state.selectedTemplateId || "manual",
      dmText,
      currentStage: Number(refs.stageInput.value || state.stage || 1)
    }
  });

  if (!resp.ok) {
    setStatus(resp.error || "Log & Advance failed", true);
    return;
  }

  state.stage = Number(resp.data.nextStage || state.stage + 1);
  refs.stageInput.value = String(state.stage);

  if (resp.data.activityWarning) {
    setStatus(`Logged note + advanced to stage ${resp.data.nextStage}. Activity warning: ${resp.data.activityWarning}`, true);
    return;
  }

  setStatus(`Logged activity + note and advanced to stage ${resp.data.nextStage}.`, false, true);
}

async function loadSequences() {
  const resp = await sendRuntimeMessage({ type: "LINKEDIN_GET_SEQUENCES" });
  if (!resp.ok) {
    state.sequences = [];
    renderSequenceSelect();
    setStatus(resp.error || "Could not load sequences.", true);
    return;
  }

  state.sequences = Array.isArray(resp.data?.sequences) ? resp.data.sequences : [];
  if (!state.selectedSequenceId) {
    const matchSequenceId = String(state.match?.sequenceId || "").trim();
    if (matchSequenceId && state.sequences.some((item) => item.id === matchSequenceId)) {
      state.selectedSequenceId = matchSequenceId;
    }
  }
  if (!state.selectedSequenceId || !state.sequences.some((item) => item.id === state.selectedSequenceId)) {
    state.selectedSequenceId = state.sequences[0]?.id || "";
  }
  renderSequenceSelect();
}

async function loadTemplates() {
  if (!state.selectedSequenceId) {
    state.templates = [];
    state.selectedTemplateId = "";
    renderTemplatePreview();
    return;
  }

  const resp = await sendRuntimeMessage({
    type: "LINKEDIN_GET_TEMPLATES",
    payload: {
      sequenceId: state.selectedSequenceId,
      stage: state.stage
    }
  });

  if (!resp.ok) {
    state.templates = [];
    state.selectedTemplateId = "";
    renderTemplatePreview();
    setStatus(resp.error || "Could not load templates for selected stage.", true);
    return;
  }

  state.templates = Array.isArray(resp.data?.templates) ? resp.data.templates : [];
  const selected = state.templates.find((item) => item.id === state.selectedTemplateId) || state.templates[0];
  state.selectedTemplateId = selected?.id || "";
  renderTemplatePreview();
}

function renderSequenceSelect() {
  refs.templateSelect.innerHTML = "";
  state.sequences.forEach((sequence) => {
    const option = document.createElement("option");
    option.value = sequence.id;
    option.textContent = sequence.name || sequence.id;
    refs.templateSelect.appendChild(option);
  });
  refs.templateSelect.value = state.selectedSequenceId || "";
}

function renderTemplatePreview() {
  refs.templatePreview.innerHTML = "";
  if (!state.templates.length) {
    const empty = document.createElement("div");
    empty.className = "sp-card";
    empty.textContent = "No templates found for this sequence/stage.";
    refs.templatePreview.appendChild(empty);
    return;
  }

  state.templates.forEach((template) => {
    const item = document.createElement("div");
    item.className = "sp-template";

    const title = document.createElement("strong");
    title.textContent = template.label || template.id || "Template";
    item.appendChild(title);

    const body = document.createElement("div");
    body.className = "sp-card";
    body.textContent = interpolateTemplate(template.body || "");
    item.appendChild(body);

    const useBtn = document.createElement("button");
    useBtn.className = "sp-btn sp-btn-secondary";
    useBtn.type = "button";
    useBtn.textContent = "Use Template";
    useBtn.addEventListener("click", () => {
      state.selectedTemplateId = template.id;
      onInsertTemplate();
    });
    item.appendChild(useBtn);

    refs.templatePreview.appendChild(item);
  });
}

function interpolateTemplate(rawBody) {
  const personName = state.match?.person?.name || state.linkedinContext?.profileName || "there";
  const firstName = String(personName).split(/\s+/)[0] || "there";
  const orgName = state.match?.person?.orgName || "your team";

  return String(rawBody || "")
    .replace(/\{\{\s*personFirstName\s*\}\}/g, firstName)
    .replace(/\{\{\s*personName\s*\}\}/g, personName)
    .replace(/\{\{\s*orgName\s*\}\}/g, orgName)
    .replace(/\{\{\s*useCaseOrDeal\s*\}\}/g, "your current initiative")
    .replace(/\{\{\s*valueHook\s*\}\}/g, "a practical next step");
}

function setStatus(message, isError = false, isSuccess = false) {
  refs.status.textContent = message;
  refs.status.className = `sp-card ${isError ? "sp-status-error" : ""} ${isSuccess ? "sp-status-success" : ""}`.trim();
}

function sendRuntimeMessage(message) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message || "Runtime error" });
        return;
      }
      resolve(response || { ok: false, error: "No response" });
    });
  });
}

function renderKeyValueCard(container, rows) {
  if (!container) return;
  container.innerHTML = "";
  rows.forEach((row) => {
    const line = document.createElement("div");
    line.className = "sp-kv-line";

    const label = document.createElement("span");
    label.className = "sp-kv-label";
    label.textContent = `${row.label}: `;

    const value = document.createElement("span");
    value.textContent = row.value || "N/A";

    line.appendChild(label);
    line.appendChild(value);
    container.appendChild(line);
  });
}
