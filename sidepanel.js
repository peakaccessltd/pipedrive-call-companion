const refs = {
  refreshBtn: document.getElementById("refreshBtn"),
  linkedinContext: document.getElementById("linkedinContext"),
  matchCard: document.getElementById("matchCard"),
  matchCandidates: document.getElementById("matchCandidates"),
  sequenceSelect: document.getElementById("sequenceSelect"),
  stageInput: document.getElementById("stageInput"),
  templatesList: document.getElementById("templatesList"),
  talkingPoints: document.getElementById("talkingPoints"),
  talkingPointsToggle: document.getElementById("talkingPointsToggle"),
  draftText: document.getElementById("draftText"),
  insertBtn: document.getElementById("insertBtn"),
  copyBtn: document.getElementById("copyBtn"),
  logAdvanceBtn: document.getElementById("logAdvanceBtn"),
  status: document.getElementById("status")
};

const state = {
  linkedinContext: null,
  match: null,
  sequences: [],
  templates: [],
  talkingPoints: [],
  selectedTemplate: null,
  selectedSequenceId: "",
  stage: 1,
  talkingPointsExpanded: false
};

refs.refreshBtn.addEventListener("click", refreshAll);
refs.sequenceSelect.addEventListener("change", onSequenceChanged);
refs.stageInput.addEventListener("change", onStageChanged);
refs.insertBtn.addEventListener("click", onInsertTemplate);
refs.copyBtn.addEventListener("click", onCopy);
refs.logAdvanceBtn.addEventListener("click", onLogAndAdvance);
refs.talkingPointsToggle.addEventListener("click", onTalkingPointsToggle);

refreshAll();
syncTalkingPointsVisibility();

async function refreshAll() {
  setStatus("Loading LinkedIn Mode context...");

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
    state.stage = Number(state.match.currentStage || 1);
    refs.stageInput.value = String(state.stage);
    renderMatchCard();
    await loadTalkingPoints();
  }

  const sequencesResp = await sendRuntimeMessage({ type: "LINKEDIN_GET_SEQUENCES" });
  if (!sequencesResp.ok) {
    setStatus(sequencesResp.error || "Could not load sequences", true);
    return;
  }

  state.sequences = sequencesResp.data.sequences || [];
  renderSequenceSelect();

  if (state.sequences.length) {
    const recommended = String(state.match?.sequenceId || "").trim();
    const current = String(state.selectedSequenceId || "").trim();
    const hasCurrent = state.sequences.some((sequence) => sequence.id === current);
    const hasRecommended = state.sequences.some((sequence) => sequence.id === recommended);
    const resolved = hasCurrent ? current : hasRecommended ? recommended : state.sequences[0].id;
    state.selectedSequenceId = resolved;
    refs.sequenceSelect.value = resolved;
  }

  await loadTemplates();
  setStatus("LinkedIn Mode ready.", false, true);
}

async function loadTemplates() {
  if (!state.selectedSequenceId) {
    state.templates = [];
    renderTemplates();
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
    setStatus(resp.error || "Failed to load templates", true);
    return;
  }

  state.templates = resp.data.templates || [];
  renderTemplates();
}

async function loadTalkingPoints() {
  state.talkingPoints = [];
  renderTalkingPoints();

  const personId = state.match?.person?.id;
  if (!personId) return;

  const resp = await sendRuntimeMessage({
    type: "LINKEDIN_GET_TALKING_POINTS",
    payload: { personId }
  });

  if (!resp.ok) {
    setStatus(resp.error || "Failed to load talking points", true);
    return;
  }

  state.talkingPoints = resp.data?.cards || [];
  renderTalkingPoints(resp.data?.preCall || null);
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
      { label: "Match strategy", value: match.strategy || "N/A" },
      { label: "DM eligible", value: match.dmEligible ? "Yes" : "No" },
      { label: "Current stage", value: String(match.currentStage || 1) },
      { label: "Sequence", value: match.sequenceId || "(none)" }
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
      await loadTalkingPoints();
      setStatus("Match confirmed and LinkedIn URL saved.", false, true);
    });

    card.appendChild(confirmBtn);
    refs.matchCandidates.appendChild(card);
  });
}

function renderSequenceSelect() {
  refs.sequenceSelect.innerHTML = "";

  if (!state.sequences.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No sequences";
    refs.sequenceSelect.appendChild(option);
    return;
  }

  state.sequences.forEach((sequence) => {
    const option = document.createElement("option");
    option.value = sequence.id;
    option.textContent = sequence.name;
    refs.sequenceSelect.appendChild(option);
  });
}

function renderTemplates() {
  refs.templatesList.innerHTML = "";

  if (!state.templates.length) {
    const card = document.createElement("div");
    card.className = "sp-card";
    card.textContent = "No templates found for this stage. Try another stage or sequence.";
    refs.templatesList.appendChild(card);
    return;
  }

  state.templates.forEach((template) => {
    const wrapper = document.createElement("article");
    wrapper.className = "sp-template";

    const badge = document.createElement("span");
    badge.className = "sp-badge";
    badge.textContent = `Stage ${template.stage}`;

    const title = document.createElement("strong");
    title.textContent = template.label;

    const body = document.createElement("div");
    body.textContent = renderTemplateBody(template.body);

    const selectBtn = document.createElement("button");
    selectBtn.type = "button";
    selectBtn.className = "sp-btn sp-btn-secondary";
    selectBtn.textContent = "Use Template";
    selectBtn.addEventListener("click", () => {
      state.selectedTemplate = template;
      refs.draftText.value = renderTemplateBody(template.body);
      setStatus(`Selected template: ${template.label}`);
    });

    wrapper.appendChild(badge);
    wrapper.appendChild(title);
    wrapper.appendChild(body);
    wrapper.appendChild(selectBtn);
    refs.templatesList.appendChild(wrapper);
  });
}

function renderTalkingPoints(preCall = null) {
  refs.talkingPoints.innerHTML = "";

  if (preCall?.oneLiner) {
    const intro = document.createElement("div");
    intro.className = "sp-card";
    intro.textContent = preCall.oneLiner;
    refs.talkingPoints.appendChild(intro);
  }

  if (!state.talkingPoints.length) {
    const card = document.createElement("div");
    card.className = "sp-card";
    card.textContent = "No talking points yet. Confirm a person match first.";
    refs.talkingPoints.appendChild(card);
    return;
  }

  state.talkingPoints.forEach((cardData) => {
    const card = document.createElement("article");
    card.className = "sp-talk-item";

    const title = document.createElement("strong");
    title.textContent = cardData.title || "Talking point";

    const list = document.createElement("ul");
    list.style.margin = "0";
    list.style.paddingLeft = "18px";
    list.style.display = "grid";
    list.style.gap = "4px";
    (cardData.bullets || []).forEach((bullet) => {
      const li = document.createElement("li");
      li.textContent = bullet;
      list.appendChild(li);
    });

    card.appendChild(title);
    card.appendChild(list);
    refs.talkingPoints.appendChild(card);
  });
}

function onTalkingPointsToggle() {
  state.talkingPointsExpanded = !state.talkingPointsExpanded;
  syncTalkingPointsVisibility();
}

function syncTalkingPointsVisibility() {
  refs.talkingPoints.classList.toggle("sp-collapsed", !state.talkingPointsExpanded);
  refs.talkingPointsToggle.textContent = state.talkingPointsExpanded ? "Hide" : "Show";
  refs.talkingPointsToggle.setAttribute("aria-expanded", state.talkingPointsExpanded ? "true" : "false");
}

function renderTemplateBody(body) {
  const personName = state.match?.person?.name || state.linkedinContext?.profileName || "there";
  const firstName = String(personName).split(/\s+/)[0] || "there";
  const org = state.match?.person?.orgName || "your team";
  const dealTitle = state.match?.person?.dealTitle || "your current initiative";

  return String(body || "")
    .replace(/\{\{\s*personFirstName\s*\}\}/g, firstName)
    .replace(/\{\{\s*personName\s*\}\}/g, personName)
    .replace(/\{\{\s*orgName\s*\}\}/g, org)
    .replace(/\{\{\s*useCaseOrDeal\s*\}\}/g, dealTitle)
    .replace(/\{\{\s*valueHook\s*\}\}/g, "a faster and more predictable follow-up process");
}

async function onInsertTemplate() {
  const text = refs.draftText.value.trim();
  if (!text) {
    setStatus("Select or edit a template first.", true);
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
    setStatus("Composer unavailable; copied to clipboard instead.", false, true);
  } else {
    setStatus("Could not insert template or copy to clipboard.", true);
  }
}

async function onCopy() {
  const text = refs.draftText.value.trim();
  if (!text) {
    setStatus("Nothing to copy.", true);
    return;
  }

  const resp = await sendRuntimeMessage({
    type: "LINKEDIN_COPY_TEXT",
    payload: { text }
  });

  if (!resp.ok || !resp.data?.copied) {
    setStatus(resp.error || "Clipboard copy failed", true);
    return;
  }

  setStatus("Copied to clipboard.", false, true);
}

async function onLogAndAdvance() {
  if (!state.match?.person?.id) {
    setStatus("No matched Pipedrive person. Confirm a match first.", true);
    return;
  }

  const composerResp = await sendRuntimeMessage({ type: "LINKEDIN_GET_COMPOSER_TEXT" });
  const composerText = composerResp.ok ? composerResp.data.text : "";
  const draftText = refs.draftText.value.trim();
  const dmText = composerText || draftText;

  if (!dmText) {
    setStatus("No DM text found to log.", true);
    return;
  }

  const resp = await sendRuntimeMessage({
    type: "LINKEDIN_LOG_AND_ADVANCE",
    payload: {
      personId: state.match.person.id,
      profileUrl: state.linkedinContext?.profileUrl || "",
      sequenceId: state.selectedSequenceId,
      templateId: state.selectedTemplate?.id || "manual",
      dmText,
      currentStage: Number(refs.stageInput.value || state.stage || 1)
    }
  });

  if (!resp.ok) {
    setStatus(resp.error || "Log & Advance failed", true);
    return;
  }

  state.match = {
    ...state.match,
    currentStage: resp.data.nextStage,
    dmEligible: resp.data.dmEligible
  };
  state.stage = Number(resp.data.nextStage || state.stage + 1);
  refs.stageInput.value = String(state.stage);
  renderMatchCard();
  await loadTemplates();

  setStatus(`Logged to Pipedrive and advanced to stage ${resp.data.nextStage}.`, false, true);
}

async function onSequenceChanged() {
  state.selectedSequenceId = refs.sequenceSelect.value;
  await loadTemplates();
}

async function onStageChanged() {
  state.stage = Number(refs.stageInput.value || 1);
  await loadTemplates();
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
