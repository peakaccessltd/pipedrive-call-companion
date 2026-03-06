const HOST_ID = "peak-access-linkedin-host";
const SHADOW_WRAPPER_CLASS = "pa-linkedin-wrapper";
const STATE = {
  mounted: false,
  shadowRoot: null,
  drawerVisible: false,
  fallback: {
    context: null,
    match: null,
    candidates: [],
    sequences: [],
    templates: [],
    talkingPoints: [],
    selectedSequenceId: "",
    selectedTemplate: null,
    stage: 1,
    talkingPointsExpanded: false
  }
};

initLinkedInBridge();

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  handleMessage(message)
    .then((data) => sendResponse({ ok: true, data }))
    .catch((error) => sendResponse({ ok: false, error: error.message || String(error) }));

  return true;
});

async function handleMessage(message) {
  const type = message?.type;

  if (type === "LINKEDIN_DETECT_CONTEXT") {
    return detectLinkedInContext();
  }

  if (type === "LINKEDIN_SET_COMPOSER_TEXT") {
    const text = String(message?.payload?.text || "");
    const inserted = await setComposerText(text);

    if (!inserted) {
      const copied = await fallbackCopyToClipboard(text);
      return { inserted: false, copied };
    }

    return { inserted: true, copied: false };
  }

  if (type === "LINKEDIN_GET_COMPOSER_TEXT") {
    return { text: getComposerText() };
  }

  if (type === "LINKEDIN_COPY_TEXT") {
    const copied = await fallbackCopyToClipboard(String(message?.payload?.text || ""));
    return { copied };
  }

  throw new Error("Unsupported LinkedIn content message.");
}

function initLinkedInBridge() {
  if (STATE.mounted) return;
  STATE.mounted = true;

  injectShadowWidget();
  watchUrlChanges();
}

function injectShadowWidget() {
  let host = document.getElementById(HOST_ID);
  if (!host) {
    host = document.createElement("div");
    host.id = HOST_ID;
    host.style.position = "fixed";
    host.style.inset = "0";
    host.style.pointerEvents = "none";
    host.style.zIndex = "2147483647";
    document.documentElement.appendChild(host);
  }

  const shadow = host.shadowRoot || host.attachShadow({ mode: "open" });
  STATE.shadowRoot = shadow;
  shadow.innerHTML = "";

  const wrapper = document.createElement("div");
  wrapper.className = SHADOW_WRAPPER_CLASS;

  const style = document.createElement("style");
  style.textContent = `
    *, *::before, *::after {
      box-sizing: border-box;
    }

    .${SHADOW_WRAPPER_CLASS} {
      position: fixed;
      right: 14px;
      bottom: 14px;
      z-index: 2147483647 !important;
      pointer-events: auto;
      font-family: "Open Sans", "Segoe UI", "Helvetica Neue", Arial, sans-serif;
      display: flex;
      align-items: center;
      gap: 8px;
      background: rgba(64, 60, 54, 0.95);
      color: #f8f6f2;
      border: 1px solid rgba(214, 106, 43, 0.74);
      border-radius: 999px;
      padding: 7px 11px;
      box-shadow: 0 12px 20px rgba(0, 0, 0, 0.24);
    }

    .${SHADOW_WRAPPER_CLASS} button {
      border: 0;
      border-radius: 999px;
      padding: 6px 11px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      color: #fff;
      background: linear-gradient(180deg, #d66a2b 0%, #b9561d 100%);
    }

    .${SHADOW_WRAPPER_CLASS} .meta {
      font-size: 13px;
      opacity: 0.9;
      white-space: nowrap;
      max-width: 280px;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .pa-drawer {
      position: fixed;
      right: 8px;
      bottom: 60px;
      width: min(430px, calc(100vw - 16px));
      height: min(86vh, calc(100vh - 32px));
      max-height: calc(100vh - 24px);
      background: linear-gradient(180deg, #f8f6f2 0%, #f1ede6 100%);
      border: 1px solid #c9c1b6;
      border-radius: 14px;
      box-shadow: 0 18px 30px rgba(0, 0, 0, 0.26);
      overflow: hidden;
      display: none;
      z-index: 2147483647 !important;
      pointer-events: auto;
      flex-direction: column;
    }

    .pa-drawer-visible {
      display: flex;
    }

    .pa-drawer-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 12px;
      background: #4b4741;
      color: #fff;
      font-size: 15px;
      font-weight: 700;
      letter-spacing: 0.2px;
    }

    .pa-drawer-header button {
      border: 0;
      border-radius: 10px;
      background: #f5f1ea;
      color: #4b4741;
      font-size: 14px;
      font-weight: 700;
      padding: 5px 9px;
      cursor: pointer;
    }

    .pa-body {
      flex: 1 1 auto;
      min-height: 0;
      font-size: 14px;
      color: #2c2a27;
      display: flex;
      flex-direction: column;
      background: transparent;
    }

    .pa-scroll {
      flex: 1 1 auto;
      min-height: 0;
      overflow: auto;
      overscroll-behavior: contain;
      padding: 11px;
      display: grid;
      gap: 11px;
    }

    .pa-footer {
      flex: 0 0 auto;
      border-top: 1px solid #d7d0c5;
      background: linear-gradient(180deg, rgba(248, 246, 242, 0.95) 0%, rgba(241, 237, 230, 0.98) 100%);
      backdrop-filter: blur(4px);
      padding: 11px;
      display: grid;
      gap: 11px;
    }

    .pa-section {
      display: grid;
      gap: 6px;
      min-width: 0;
    }

    .pa-section-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
    }

    .pa-section h3 {
      margin: 0;
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.65px;
      color: #5c5852;
      font-weight: 700;
    }

    .pa-card {
      border: 1px solid #c9c1b6;
      border-radius: 12px;
      background: #fff;
      padding: 9px;
      white-space: pre-wrap;
      line-height: 1.45;
      min-width: 0;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .pa-input, .pa-select, .pa-textarea {
      width: 100%;
      max-width: 100%;
      border: 1px solid #c9c1b6;
      border-radius: 8px;
      background: #fff;
      padding: 9px;
      font-size: 14px;
      font-family: "Open Sans", "Segoe UI", "Helvetica Neue", Arial, sans-serif;
      min-width: 0;
    }

    .pa-textarea {
      min-height: 130px;
      resize: vertical;
      line-height: 1.45;
    }

    .pa-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      min-width: 0;
    }

    .pa-row > * {
      min-width: 0;
    }

    .pa-full {
      grid-column: span 2;
    }

    .pa-list {
      display: grid;
      gap: 7px;
      min-width: 0;
    }

    .pa-collapsed {
      display: none;
    }

    .pa-item {
      border: 1px solid #c9c1b6;
      border-radius: 10px;
      background: #fff;
      padding: 9px;
      display: grid;
      gap: 7px;
      min-width: 0;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .pa-item button, .pa-row button, .pa-top-actions button {
      border: 1px solid transparent;
      border-radius: 8px;
      padding: 8px 10px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      font-family: "Open Sans", "Segoe UI", "Helvetica Neue", Arial, sans-serif;
      transition: background 120ms ease, transform 120ms ease;
    }

    .pa-primary {
      background: #d66a2b;
      color: #fff;
    }

    .pa-secondary {
      background: #f2efea;
      color: #4b4741;
      border-color: #c9c1b6;
    }

    .pa-primary:hover, .pa-secondary:hover {
      transform: translateY(-1px);
    }

    .pa-top-actions {
      display: flex;
      justify-content: flex-end;
      gap: 8px;
    }

    .pa-drawer-brand {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .pa-drawer-logo {
      width: 20px;
      height: 20px;
      object-fit: contain;
      flex: 0 0 auto;
    }
  `;

  const button = document.createElement("button");
  button.type = "button";
  button.textContent = "LinkedIn Mode";
  button.addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "OPEN_LINKEDIN_SIDE_PANEL" }, (response) => {
      if (chrome.runtime.lastError || !response?.ok) {
        toggleFallbackDrawer();
      }
    });
  });

  const meta = document.createElement("div");
  meta.className = "meta";
  meta.textContent = buildMetaText();

  wrapper.appendChild(button);
  wrapper.appendChild(meta);
  shadow.appendChild(style);
  shadow.appendChild(wrapper);
  const drawer = createFallbackDrawer();
  shadow.appendChild(drawer);

  if (STATE.drawerVisible) {
    refreshFallbackData();
  }
}

function buildMetaText() {
  const context = detectLinkedInContext();

  if (context.isMessaging) {
    return `Messaging: ${context.profileName || "thread"}`;
  }

  if (context.profileUrl) {
    return `Profile: ${context.profileName || context.profileUrl}`;
  }

  return "Open a LinkedIn profile or messaging thread";
}

function watchUrlChanges() {
  let currentUrl = location.href;

  setInterval(() => {
    if (location.href !== currentUrl) {
      currentUrl = location.href;
      injectShadowWidget();
    }
  }, 1000);
}

function createFallbackDrawer() {
  const drawer = document.createElement("section");
  drawer.className = `pa-drawer ${STATE.drawerVisible ? "pa-drawer-visible" : ""}`.trim();
  drawer.id = "pa-linkedin-drawer";

  const header = document.createElement("div");
  header.className = "pa-drawer-header";
  const brand = document.createElement("div");
  brand.className = "pa-drawer-brand";
  const logo = document.createElement("img");
  logo.className = "pa-drawer-logo";
  logo.src = chrome.runtime.getURL("icons/logo-source.png");
  logo.alt = "Peak Access";
  const headerTitle = document.createElement("span");
  headerTitle.textContent = "LinkedIn Mode (Fallback Drawer)";
  brand.appendChild(logo);
  brand.appendChild(headerTitle);

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.textContent = "Close";
  closeBtn.addEventListener("click", () => {
    STATE.drawerVisible = false;
    updateDrawerState();
  });

  const body = document.createElement("div");
  body.className = "pa-body";
  body.innerHTML = `
    <div class="pa-scroll">
      <div class="pa-top-actions">
        <button id="paRefreshBtn" class="pa-secondary" type="button">Refresh</button>
      </div>
      <section class="pa-section">
        <h3>Detected LinkedIn Context</h3>
        <div id="paContextCard" class="pa-card"></div>
      </section>
      <section class="pa-section">
        <h3>Pipedrive Match</h3>
        <div id="paMatchCard" class="pa-card"></div>
        <div class="pa-row">
          <input id="paSearchName" class="pa-input" type="text" placeholder="Search Pipedrive person by name" />
          <button id="paSearchBtn" class="pa-secondary" type="button">Search</button>
        </div>
        <div id="paCandidateList" class="pa-list"></div>
      </section>
      <section class="pa-section">
        <h3>Sequence & Templates</h3>
        <select id="paSequenceSelect" class="pa-select"></select>
        <input id="paStageInput" class="pa-input" type="number" min="1" step="1" />
        <div id="paTemplateList" class="pa-list"></div>
      </section>
      <section class="pa-section">
        <div class="pa-section-head">
          <h3>Talking Points</h3>
          <button id="paTalkingToggle" class="pa-secondary" type="button" aria-expanded="false">Show</button>
        </div>
        <div id="paTalkingPoints" class="pa-list pa-collapsed"></div>
      </section>
    </div>
    <div class="pa-footer">
      <section class="pa-section">
        <h3>Composer</h3>
        <textarea id="paDraftText" class="pa-textarea" placeholder="Template text appears here"></textarea>
        <div class="pa-row">
          <button id="paInsertBtn" class="pa-primary" type="button">Insert Template</button>
          <button id="paCopyBtn" class="pa-secondary" type="button">Copy</button>
          <button id="paLogBtn" class="pa-primary pa-full" type="button">Log & Advance</button>
        </div>
      </section>
      <section class="pa-section">
        <h3>Status</h3>
        <div id="paStatus" class="pa-card">Idle.</div>
      </section>
    </div>
  `;

  header.appendChild(brand);
  header.appendChild(closeBtn);
  drawer.appendChild(header);
  drawer.appendChild(body);

  wireFallbackEvents(drawer);
  return drawer;
}

function toggleFallbackDrawer() {
  STATE.drawerVisible = !STATE.drawerVisible;
  updateDrawerState();

  if (STATE.drawerVisible) {
    refreshFallbackData();
  }
}

function updateDrawerState() {
  const shadow = STATE.shadowRoot;
  if (!shadow) return;
  const drawer = shadow.getElementById("pa-linkedin-drawer");
  if (!drawer) return;

  drawer.classList.toggle("pa-drawer-visible", STATE.drawerVisible);
}

function wireFallbackEvents(drawer) {
  drawer.querySelector("#paRefreshBtn")?.addEventListener("click", () => refreshFallbackData());
  drawer.querySelector("#paSequenceSelect")?.addEventListener("change", async (event) => {
    STATE.fallback.selectedSequenceId = String(event.target?.value || "");
    await loadFallbackTemplates();
  });
  drawer.querySelector("#paStageInput")?.addEventListener("change", async (event) => {
    STATE.fallback.stage = Number(event.target?.value || 1);
    await loadFallbackTemplates();
  });
  drawer.querySelector("#paSearchBtn")?.addEventListener("click", async () => {
    await runFallbackPersonSearch(drawer);
  });
  drawer.querySelector("#paSearchName")?.addEventListener("keydown", async (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      await runFallbackPersonSearch(drawer);
    }
  });
  // Delegated fallback in case direct listeners get replaced during reinjection.
  drawer.addEventListener("click", async (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    if (target.id === "paSearchBtn") {
      event.preventDefault();
      await runFallbackPersonSearch(drawer);
      return;
    }

    const confirmBtn = target.closest(".pa-confirm-match");
    if (confirmBtn) {
      event.preventDefault();
      const personId = Number(confirmBtn.getAttribute("data-person-id"));
      if (!Number.isFinite(personId) || personId <= 0) {
        setFallbackStatus("Candidate ID missing; cannot confirm.", true);
        return;
      }
      await runConfirmMatch(drawer, personId);
    }
  });
  drawer.querySelector("#paInsertBtn")?.addEventListener("click", async () => {
    const text = String(drawer.querySelector("#paDraftText")?.value || "").trim();
    if (!text) return setFallbackStatus("Select a template first.", true);
    const ok = await setComposerText(text);
    if (!ok) {
      const copied = await fallbackCopyToClipboard(text);
      return setFallbackStatus(copied ? "Composer unavailable, copied to clipboard." : "Insert failed.", !copied);
    }
    setFallbackStatus("Inserted into LinkedIn composer.", false);
  });
  drawer.querySelector("#paCopyBtn")?.addEventListener("click", async () => {
    const text = String(drawer.querySelector("#paDraftText")?.value || "").trim();
    if (!text) return setFallbackStatus("Nothing to copy.", true);
    const copied = await fallbackCopyToClipboard(text);
    setFallbackStatus(copied ? "Copied." : "Copy failed.", !copied);
  });
  drawer.querySelector("#paLogBtn")?.addEventListener("click", async () => {
    const personId = STATE.fallback.match?.person?.id;
    if (!personId) return setFallbackStatus("No matched person to log against.", true);

    const draftText = String(drawer.querySelector("#paDraftText")?.value || "").trim();
    const dmText = getComposerText() || draftText;
    if (!dmText) return setFallbackStatus("No message text to log.", true);

    const response = await sendRuntimeMessage({
      type: "LINKEDIN_LOG_AND_ADVANCE",
      payload: {
        personId,
        profileUrl: STATE.fallback.context?.profileUrl || "",
        sequenceId: STATE.fallback.selectedSequenceId,
        templateId: STATE.fallback.selectedTemplate?.id || "manual",
        dmText,
        currentStage: Number(STATE.fallback.stage || 1)
      }
    });

    if (!response.ok) return setFallbackStatus(response.error || "Log & Advance failed.", true);

    STATE.fallback.stage = Number(response.data?.nextStage || STATE.fallback.stage + 1);
    const stageInput = drawer.querySelector("#paStageInput");
    if (stageInput) stageInput.value = String(STATE.fallback.stage);
    await loadFallbackTemplates();
    setFallbackStatus(`Logged and advanced to stage ${STATE.fallback.stage}.`, false);
  });
  drawer.querySelector("#paTalkingToggle")?.addEventListener("click", () => {
    STATE.fallback.talkingPointsExpanded = !STATE.fallback.talkingPointsExpanded;
    syncFallbackTalkingVisibility(drawer);
  });
  syncFallbackTalkingVisibility(drawer);
}

async function runFallbackPersonSearch(drawer) {
  const input = drawer.querySelector("#paSearchName");
  const query = String(input?.value || "").trim();
  if (!query) {
    setFallbackStatus("Enter a name to search.", true);
    return;
  }

  setFallbackStatus(`Searching for '${query}'...`);

  const response = await sendRuntimeMessage({
    type: "LINKEDIN_SEARCH_PERSONS",
    payload: { query }
  });

  if (!response.ok) {
    setFallbackStatus(response.error || "Search failed.", true);
    return;
  }

  const candidates = response.data?.candidates || [];
  STATE.fallback.candidates = candidates;
  renderMatch(drawer, null, candidates, "");
  setFallbackStatus(`Found ${candidates.length} candidate(s).`);
}

async function runConfirmMatch(drawer, personId) {
  const profileUrl = STATE.fallback.context?.profileUrl || "";
  if (!profileUrl) {
    setFallbackStatus("No LinkedIn profile URL detected; cannot save match.", true);
    setMatchCardMessage(drawer, "No LinkedIn profile URL detected; cannot save match.");
    return;
  }

  setFallbackStatus("Saving match to Pipedrive...");
  setMatchCardMessage(drawer, "Saving match to Pipedrive...");

  const response = await sendRuntimeMessage({
    type: "LINKEDIN_CONFIRM_MATCH",
    payload: {
      personId,
      profileUrl
    }
  });

  if (!response.ok) {
    setFallbackStatus(response.error || "Confirm failed.", true);
    setMatchCardMessage(drawer, `Confirm failed: ${response.error || "Unknown error"}`);
    return;
  }

  STATE.fallback.match = response.data;
  STATE.fallback.candidates = [];
  renderMatch(drawer, response.data.person, [], "");
  await loadFallbackTalkingPoints(drawer);
  setFallbackStatus(`Match saved: ${response.data.person?.name || `#${personId}`}.`);
  setMatchCardMessage(drawer, `Matched: ${response.data.person?.name || `#${personId}`}`);
}

async function refreshFallbackData() {
  setFallbackStatus("Refreshing context...");
  const drawer = STATE.shadowRoot?.getElementById("pa-linkedin-drawer");
  if (!drawer) return;

  const contextResponse = await sendRuntimeMessage({ type: "LINKEDIN_GET_CONTEXT" });
  if (!contextResponse.ok) {
    renderContextCard(drawer, `Error: ${contextResponse.error}`);
    return setFallbackStatus(contextResponse.error || "Context detection failed.", true);
  }

  STATE.fallback.context = contextResponse.data;
  if (!STATE.fallback.context.profileName) {
    STATE.fallback.context.profileName = inferNameFromProfileUrl(STATE.fallback.context.profileUrl);
  }
  renderContextCard(drawer, [
    `Profile URL: ${contextResponse.data.profileUrl || "N/A"}`,
    `Profile Name: ${STATE.fallback.context.profileName || "N/A"}`,
    `Messaging: ${contextResponse.data.isMessaging ? "Yes" : "No"}`,
    `Thread: ${contextResponse.data.threadTitle || "N/A"}`
  ].join("\n"));

  const matchResponse = await sendRuntimeMessage({
    type: "LINKEDIN_MATCH_PERSON",
    payload: {
      profileUrl: contextResponse.data.profileUrl,
      profileName: STATE.fallback.context.profileName,
      emailHint: contextResponse.data.emailHint
    }
  });

  if (!matchResponse.ok) {
    renderMatch(drawer, null, [], matchResponse.error || "Match failed");
  } else {
    STATE.fallback.match = matchResponse.data;
    STATE.fallback.stage = Number(matchResponse.data.currentStage || 1);
    const stageInput = drawer.querySelector("#paStageInput");
    if (stageInput) stageInput.value = String(STATE.fallback.stage);
    renderMatch(drawer, matchResponse.data.person, matchResponse.data.candidates || [], "");
    const searchNameInput = drawer.querySelector("#paSearchName");
    if (searchNameInput && !searchNameInput.value) {
      searchNameInput.value = STATE.fallback.context.profileName || "";
    }

    // If no direct match was found, auto-run a name search so candidates appear without extra clicks.
    if (!matchResponse.data.person && STATE.fallback.context.profileName) {
      await runFallbackPersonSearch(drawer);
    }

    await loadFallbackTalkingPoints(drawer);
  }

  const sequencesResponse = await sendRuntimeMessage({ type: "LINKEDIN_GET_SEQUENCES" });
  if (!sequencesResponse.ok) {
    return setFallbackStatus(sequencesResponse.error || "Sequence fetch failed.", true);
  }

  STATE.fallback.sequences = sequencesResponse.data?.sequences || [];
  renderSequenceSelect(drawer);
  if (!STATE.fallback.selectedSequenceId && STATE.fallback.sequences.length) {
    STATE.fallback.selectedSequenceId =
      STATE.fallback.match?.sequenceId || STATE.fallback.sequences[0].id;
  }
  const sequenceSelect = drawer.querySelector("#paSequenceSelect");
  if (sequenceSelect) sequenceSelect.value = STATE.fallback.selectedSequenceId;

  await loadFallbackTemplates();
  setFallbackStatus("Ready.");
}

async function loadFallbackTemplates() {
  const drawer = STATE.shadowRoot?.getElementById("pa-linkedin-drawer");
  if (!drawer) return;

  if (!STATE.fallback.selectedSequenceId) {
    STATE.fallback.templates = [];
    renderTemplates(drawer);
    return;
  }

  const response = await sendRuntimeMessage({
    type: "LINKEDIN_GET_TEMPLATES",
    payload: {
      sequenceId: STATE.fallback.selectedSequenceId,
      stage: Number(STATE.fallback.stage || 1)
    }
  });

  if (!response.ok) {
    STATE.fallback.templates = [];
    renderTemplates(drawer);
    return setFallbackStatus(response.error || "Template fetch failed.", true);
  }

  STATE.fallback.templates = response.data?.templates || [];
  renderTemplates(drawer);
}

async function loadFallbackTalkingPoints(drawer) {
  const root = drawer.querySelector("#paTalkingPoints");
  if (!root) return;
  root.innerHTML = "";
  STATE.fallback.talkingPoints = [];

  const personId = STATE.fallback.match?.person?.id;
  if (!personId) {
    root.innerHTML = `<div class=\"pa-card\">No talking points yet. Confirm a person match first.</div>`;
    return;
  }

  const response = await sendRuntimeMessage({
    type: "LINKEDIN_GET_TALKING_POINTS",
    payload: { personId }
  });

  if (!response.ok) {
    root.innerHTML = `<div class=\"pa-card\">${response.error || "Failed to load talking points."}</div>`;
    return;
  }

  const preCall = response.data?.preCall || null;
  const cards = response.data?.cards || [];
  STATE.fallback.talkingPoints = cards;

  if (preCall?.oneLiner) {
    const intro = document.createElement("div");
    intro.className = "pa-card";
    intro.textContent = preCall.oneLiner;
    root.appendChild(intro);
  }

  cards.forEach((cardData) => {
    const item = document.createElement("div");
    item.className = "pa-item";

    const title = document.createElement("strong");
    title.textContent = cardData.title || "Talking point";
    item.appendChild(title);

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
    item.appendChild(list);
    root.appendChild(item);
  });

  if (!preCall?.oneLiner && !cards.length) {
    root.innerHTML = `<div class=\"pa-card\">No talking points available.</div>`;
  }
}

function renderSequenceSelect(drawer) {
  const select = drawer.querySelector("#paSequenceSelect");
  if (!select) return;
  select.innerHTML = "";

  STATE.fallback.sequences.forEach((sequence) => {
    const option = document.createElement("option");
    option.value = sequence.id;
    option.textContent = sequence.name;
    select.appendChild(option);
  });
}

function renderTemplates(drawer) {
  const root = drawer.querySelector("#paTemplateList");
  if (!root) return;
  root.innerHTML = "";

  if (!STATE.fallback.templates.length) {
    root.innerHTML = `<div class=\"pa-card\">No templates found for this stage.</div>`;
    return;
  }

  STATE.fallback.templates.forEach((template) => {
    const item = document.createElement("div");
    item.className = "pa-item";
    item.innerHTML = `<strong>${template.label}</strong><div>${interpolateTemplate(template.body)}</div>`;

    const useBtn = document.createElement("button");
    useBtn.type = "button";
    useBtn.className = "pa-secondary";
    useBtn.textContent = "Use Template";
    useBtn.addEventListener("click", () => {
      STATE.fallback.selectedTemplate = template;
      const draft = drawer.querySelector("#paDraftText");
      if (draft) draft.value = interpolateTemplate(template.body);
      setFallbackStatus(`Selected: ${template.label}`);
    });

    item.appendChild(useBtn);
    root.appendChild(item);
  });
}

function syncFallbackTalkingVisibility(drawer) {
  const root = drawer.querySelector("#paTalkingPoints");
  const toggle = drawer.querySelector("#paTalkingToggle");
  if (!root || !toggle) return;

  root.classList.toggle("pa-collapsed", !STATE.fallback.talkingPointsExpanded);
  toggle.textContent = STATE.fallback.talkingPointsExpanded ? "Hide" : "Show";
  toggle.setAttribute("aria-expanded", STATE.fallback.talkingPointsExpanded ? "true" : "false");
}

function renderContextCard(drawer, text) {
  const card = drawer.querySelector("#paContextCard");
  if (card) card.textContent = text;
}

function renderMatch(drawer, person, candidates, errorText) {
  const matchCard = drawer.querySelector("#paMatchCard");
  if (matchCard) {
    if (errorText) {
      matchCard.textContent = errorText;
    } else if (!person) {
      matchCard.textContent = "No direct match found.";
    } else {
      matchCard.textContent = [
        `Matched: ${person.name} (#${person.id})`,
        `Org: ${person.orgName || "N/A"}`,
        `DM eligible: ${STATE.fallback.match?.dmEligible ? "Yes" : "No"}`,
        `Current stage: ${STATE.fallback.match?.currentStage || 1}`
      ].join("\n");
    }
  }

  const candidateRoot = drawer.querySelector("#paCandidateList");
  if (!candidateRoot) return;
  candidateRoot.innerHTML = "";
  STATE.fallback.candidates = candidates;

  candidates.forEach((candidate) => {
    const item = document.createElement("div");
    item.className = "pa-item";
    item.textContent = `${candidate.name} (#${candidate.id}) ${candidate.orgName ? `- ${candidate.orgName}` : ""}`;

    const confirm = document.createElement("button");
    confirm.type = "button";
    confirm.className = "pa-secondary pa-confirm-match";
    confirm.setAttribute("data-person-id", String(candidate.id));
    confirm.textContent = "Confirm Match + Save URL";
    confirm.addEventListener("click", async (event) => {
      event.preventDefault();
      confirm.disabled = true;
      const originalText = confirm.textContent;
      confirm.textContent = "Saving...";
      try {
        await runConfirmMatch(drawer, candidate.id);
      } finally {
        confirm.disabled = false;
        confirm.textContent = originalText;
      }
    });
    item.appendChild(confirm);
    candidateRoot.appendChild(item);
  });
}

function interpolateTemplate(body) {
  const personName = STATE.fallback.match?.person?.name || STATE.fallback.context?.profileName || "there";
  const first = personName.split(/\s+/)[0] || "there";
  const orgName = STATE.fallback.match?.person?.orgName || "your team";
  return String(body || "")
    .replace(/\{\{\s*personFirstName\s*\}\}/g, first)
    .replace(/\{\{\s*personName\s*\}\}/g, personName)
    .replace(/\{\{\s*orgName\s*\}\}/g, orgName)
    .replace(/\{\{\s*useCaseOrDeal\s*\}\}/g, "your current initiative")
    .replace(/\{\{\s*valueHook\s*\}\}/g, "a cleaner and faster outreach workflow");
}

function setFallbackStatus(message, isError = false) {
  const status = STATE.shadowRoot?.getElementById("paStatus");
  if (!status) return;
  status.textContent = message;
  status.style.color = isError ? "#8b2424" : "#1b4a3b";
}

function setMatchCardMessage(drawer, message) {
  const card = drawer.querySelector("#paMatchCard");
  if (!card) return;
  card.textContent = message;
}

function inferNameFromProfileUrl(profileUrl) {
  try {
    const parsed = new URL(String(profileUrl || ""));
    const parts = parsed.pathname.split("/").filter(Boolean);
    if (parts.length < 2 || parts[0] !== "in") return "";

    const slug = parts[1].split("-").filter(Boolean);
    const cleaned = slug
      .filter((part) => !/^\d+$/.test(part))
      .slice(0, 3)
      .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase());

    return cleaned.join(" ");
  } catch (_error) {
    return "";
  }
}

function sendRuntimeMessage(message) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message || "Runtime error" });
        return;
      }
      resolve(response || { ok: false, error: "No response." });
    });
  });
}

function detectLinkedInContext() {
  return {
    profileUrl: getLinkedInProfileUrl(),
    profileName: getProfileName(),
    profileHeadline: getProfileHeadline(),
    isMessaging: /^\/messaging\/?/i.test(location.pathname),
    threadTitle: getMessagingThreadTitle(),
    composerText: getComposerText(),
    emailHint: getEmailHintFromPage()
  };
}

function getLinkedInProfileUrl() {
  const href = location.href;

  if (/linkedin\.com\/in\//i.test(href)) {
    return canonicalizeProfileUrl(href);
  }

  const anchors = Array.from(document.querySelectorAll('a[href*="linkedin.com/in/"]'));
  for (const anchor of anchors) {
    if (anchor.href) {
      return canonicalizeProfileUrl(anchor.href);
    }
  }

  return "";
}

function canonicalizeProfileUrl(url) {
  try {
    const parsed = new URL(url);
    parsed.search = "";
    parsed.hash = "";
    parsed.pathname = parsed.pathname.replace(/\/+$/, "");
    return `${parsed.origin}${parsed.pathname}`;
  } catch (_error) {
    return url;
  }
}

function getProfileName() {
  const candidates = [
    "h1.text-heading-xlarge",
    "h1",
    ".pv-text-details__left-panel h1"
  ];

  for (const selector of candidates) {
    const el = document.querySelector(selector);
    const text = el?.textContent?.trim();
    if (text) return text;
  }

  return "";
}

function getProfileHeadline() {
  const candidates = [
    ".text-body-medium.break-words",
    ".pv-text-details__left-panel .text-body-medium"
  ];

  for (const selector of candidates) {
    const el = document.querySelector(selector);
    const text = el?.textContent?.trim();
    if (text) return text;
  }

  return "";
}

function getMessagingThreadTitle() {
  const candidates = [
    ".msg-thread__link-to-profile .t-14.t-bold",
    ".msg-conversation-listitem__participant-names",
    ".msg-thread__subject"
  ];

  for (const selector of candidates) {
    const el = document.querySelector(selector);
    const text = el?.textContent?.trim();
    if (text) return text;
  }

  return "";
}

function getComposerEditable() {
  const candidates = [
    "div.msg-form__contenteditable[contenteditable='true']",
    "div[role='textbox'][contenteditable='true']",
    ".msg-form__contenteditable"
  ];

  for (const selector of candidates) {
    const el = document.querySelector(selector);
    if (el && isVisible(el)) return el;
  }

  return null;
}

function isVisible(el) {
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

function ensureMessageComposerOpen() {
  const existing = getComposerEditable();
  if (existing) return existing;

  const openComposerButtons = Array.from(document.querySelectorAll("button, a"));
  const target = openComposerButtons.find((element) => {
    const text = String(element.textContent || "").trim().toLowerCase();
    return text.includes("message") || text.includes("new message") || text.includes("send message");
  });

  if (target) {
    target.click();
  }

  return getComposerEditable();
}

async function setComposerText(text) {
  const composer = ensureMessageComposerOpen();
  if (!composer) return false;

  composer.focus();
  clearContentEditable(composer);

  const lines = String(text || "").split(/\n/);
  for (let i = 0; i < lines.length; i += 1) {
    if (i > 0) {
      document.execCommand("insertLineBreak");
    }
    document.execCommand("insertText", false, lines[i]);
  }

  composer.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true }));
  return true;
}

function clearContentEditable(element) {
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(element);
  selection.removeAllRanges();
  selection.addRange(range);
  document.execCommand("delete", false);
}

function getComposerText() {
  const composer = getComposerEditable();
  if (!composer) return "";
  return String(composer.innerText || composer.textContent || "").trim();
}

async function fallbackCopyToClipboard(text) {
  try {
    await navigator.clipboard.writeText(String(text || ""));
    return true;
  } catch (_error) {
    return false;
  }
}

function getEmailHintFromPage() {
  const mailto = document.querySelector('a[href^="mailto:"]');
  if (mailto?.getAttribute("href")) {
    return mailto.getAttribute("href").replace(/^mailto:/i, "").trim();
  }

  return "";
}
