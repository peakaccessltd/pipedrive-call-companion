const refs = {
  refreshBtn: document.getElementById("refreshBtn"),
  linkedinContext: document.getElementById("linkedinContext"),
  matchCard: document.getElementById("matchCard"),
  matchCandidates: document.getElementById("matchCandidates"),
  templateSelect: document.getElementById("templateSelect"),
  stageInput: document.getElementById("stageInput"),
  useTemplateBtn: document.getElementById("useTemplateBtn"),
  logAdvanceBtn: document.getElementById("logAdvanceBtn"),
  status: document.getElementById("status")
};

const TEMPLATE_OPTIONS = [
  { id: "touch_1", label: "Template 1: Intro" },
  { id: "touch_2", label: "Template 2: Value" },
  { id: "touch_3", label: "Template 3: Close Loop" }
];

const state = {
  linkedinContext: null,
  match: null,
  selectedTemplateId: TEMPLATE_OPTIONS[0].id,
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

  setStatus("LinkedIn Mode ready.", false, true);
}

function initTemplateSelect() {
  refs.templateSelect.innerHTML = "";
  TEMPLATE_OPTIONS.forEach((template) => {
    const option = document.createElement("option");
    option.value = template.id;
    option.textContent = template.label;
    refs.templateSelect.appendChild(option);
  });
  refs.templateSelect.value = state.selectedTemplateId;
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
  state.selectedTemplateId = refs.templateSelect.value || TEMPLATE_OPTIONS[0].id;
}

function onStageChanged() {
  const stage = Number(refs.stageInput.value || 1);
  state.stage = Number.isFinite(stage) && stage > 0 ? stage : 1;
  refs.stageInput.value = String(state.stage);
}

function getSelectedTemplateText() {
  return buildTemplateMessage({
    templateId: state.selectedTemplateId,
    stage: state.stage,
    context: state.linkedinContext,
    match: state.match
  });
}

function buildTemplateMessage({ templateId, stage, context, match }) {
  const personName = match?.person?.name || context?.profileName || "there";
  const firstName = String(personName).split(/\s+/)[0] || "there";
  const org = match?.person?.orgName || "your team";

  const intros = {
    1: "quick intro",
    2: "follow-up",
    3: "value add",
    4: "close loop"
  };

  const stageTone = intros[stage] || `stage ${stage} follow-up`;

  if (templateId === "touch_2") {
    return `Hi ${firstName}, sharing one idea for ${org} based on our ${stageTone}. If useful, I can send a short 2-step outline.`;
  }

  if (templateId === "touch_3") {
    return `Hi ${firstName}, just closing the loop on our ${stageTone}. If this is still relevant for ${org}, happy to align on next steps.`;
  }

  return `Hi ${firstName}, thanks for connecting. Reaching out as a ${stageTone} for ${org}. Open to a quick exchange this week?`;
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
      sequenceId: "manual_template_flow",
      templateId: state.selectedTemplateId,
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
