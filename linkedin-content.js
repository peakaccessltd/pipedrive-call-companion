const HOST_ID = "peak-access-linkedin-host";
const SHADOW_WRAPPER_CLASS = "pa-linkedin-wrapper";
const LINKEDIN_LAUNCHER_TOP_KEY = "linkedInModeLauncherTop";
const STATE = {
  mounted: false,
  shadowRoot: null,
  launcherWrapper: null,
  launcherTop: null,
  launcherHandlersInstalled: false,
  suppressLauncherClick: false,
  drag: {
    active: false,
    startY: 0,
    startTop: 0,
    moved: false
  },
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

  restoreLinkedInLauncherTop();
  installLauncherDragHandlers();
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
  STATE.launcherWrapper = wrapper;

  const style = document.createElement("style");
  style.textContent = `
    :host {
      --pa-font: "Open Sans", "Segoe UI", "Helvetica Neue", Arial, sans-serif;
      --pa-bg: #f4f7fb;
      --pa-surface: #ffffff;
      --pa-surface-soft: #f8fafc;
      --pa-border: #d9e1ec;
      --pa-text: #1d2a3a;
      --pa-text-muted: #5f6f86;
      --pa-primary: #0a5bd8;
      --pa-primary-strong: #0849ae;
      --pa-danger: #c2415d;
      --pa-success: #1e8e5a;
    }

    *, *::before, *::after {
      box-sizing: border-box;
    }

    .${SHADOW_WRAPPER_CLASS} {
      position: fixed;
      right: 0;
      top: 50%;
      transform: translateY(-50%);
      z-index: 2147483647 !important;
      pointer-events: auto;
      font-family: var(--pa-font);
      display: block;
      background: transparent;
      color: #fff;
      border: 0;
      border-radius: 0;
      padding: 0;
      box-shadow: none;
    }

    .${SHADOW_WRAPPER_CLASS} button {
      border: 1px solid #0a4bb0;
      border-right: 0;
      border-radius: 10px 0 0 10px;
      min-width: 38px;
      min-height: 54px;
      padding: 7px 8px;
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.35px;
      cursor: grab;
      color: #fff;
      background: linear-gradient(180deg, #0a5bd8 0%, #0849ae 100%);
      box-shadow: 0 12px 20px rgba(9, 57, 130, 0.3);
      white-space: nowrap;
    }

    .${SHADOW_WRAPPER_CLASS} button:active {
      cursor: grabbing;
    }

    .pa-drawer {
      position: fixed;
      right: 0;
      top: 0;
      width: min(460px, 92vw);
      height: 100vh;
      max-height: 100vh;
      background: linear-gradient(180deg, #f7faff 0%, var(--pa-bg) 100%);
      border-left: 1px solid var(--pa-border);
      border-radius: 0;
      box-shadow: -18px 0 30px rgba(20, 42, 71, 0.18);
      overflow: hidden;
      display: flex;
      z-index: 2147483647 !important;
      pointer-events: none;
      flex-direction: column;
      transform: translateX(100%);
      opacity: 0;
      visibility: hidden;
      transition: transform 180ms ease, opacity 180ms ease;
    }

    .pa-drawer-visible {
      transform: translateX(0);
      opacity: 1;
      visibility: visible;
      pointer-events: auto;
    }

    .pa-drawer-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px 14px;
      border-bottom: 1px solid var(--pa-border);
      background: #ffffff;
      color: #0f2239;
      font-size: 18px;
      font-weight: 800;
      letter-spacing: 0.2px;
    }

    .pa-drawer-header button {
      border: 1px solid var(--pa-border);
      border-radius: 10px;
      background: var(--pa-surface-soft);
      color: #334a65;
      font-size: 14px;
      font-weight: 700;
      padding: 9px 11px;
      cursor: pointer;
    }

    .pa-body {
      flex: 1 1 auto;
      min-height: 0;
      overflow: auto;
      overscroll-behavior: contain;
      padding: 12px 14px 18px;
      font-size: 14px;
      color: var(--pa-text);
      display: grid;
      gap: 12px;
      background: transparent;
    }

    .pa-section {
      display: grid;
      gap: 8px;
      min-width: 0;
      padding-bottom: 10px;
      border-bottom: 1px solid #e6edf5;
    }

    .pa-section:last-child {
      border-bottom: 0;
      padding-bottom: 0;
    }

    .pa-section-head {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
    }

    .pa-section h3 {
      margin: 0;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.8px;
      color: var(--pa-text-muted);
      font-weight: 800;
    }

    .pa-card {
      border: 0;
      border-radius: 0;
      background: transparent;
      padding: 0;
      font-size: 13px;
      white-space: pre-wrap;
      line-height: 1.45;
      min-width: 0;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .pa-input, .pa-select, .pa-textarea {
      width: 100%;
      max-width: 100%;
      border: 1px solid var(--pa-border);
      border-radius: 10px;
      background: #fff;
      padding: 9px;
      font-size: 14px;
      color: var(--pa-text);
      font-family: var(--pa-font);
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
      border: 0;
      border-radius: 0;
      background: transparent;
      padding: 0;
      display: grid;
      gap: 7px;
      min-width: 0;
      overflow-wrap: anywhere;
      word-break: break-word;
    }

    .pa-talk-item {
      padding: 2px 0 10px;
      margin-bottom: 10px;
      border-bottom: 1px solid #e6edf5;
      display: grid;
      gap: 7px;
      min-width: 0;
    }

    .pa-talk-item:last-child {
      margin-bottom: 0;
      padding-bottom: 0;
      border-bottom: 0;
    }

    .pa-item button, .pa-row button, .pa-top-actions button {
      border: 1px solid transparent;
      border-radius: 10px;
      padding: 9px 11px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      font-family: var(--pa-font);
      transition: background 140ms ease, transform 140ms ease, box-shadow 140ms ease;
    }

    .pa-primary {
      background: linear-gradient(180deg, var(--pa-primary) 0%, var(--pa-primary-strong) 100%);
      color: #fff;
      box-shadow: 0 8px 14px rgba(10, 91, 216, 0.22);
    }

    .pa-secondary {
      background: var(--pa-surface-soft);
      color: #334a65;
      border-color: var(--pa-border);
    }

    .pa-primary:hover {
      transform: translateY(-1px);
    }

    .pa-secondary:hover {
      background: #eef4fb;
    }

    .pa-top-actions {
      display: flex;
      justify-content: flex-end;
      gap: 8px;
    }

    .pa-drawer-brand {
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .pa-drawer-logo {
      width: 24px;
      height: 24px;
      object-fit: contain;
      flex: 0 0 auto;
    }

    .pa-kv-line {
      margin-bottom: 3px;
    }

    .pa-kv-line:last-child {
      margin-bottom: 0;
    }

    .pa-kv-label {
      font-weight: 700;
    }
  `;

  const button = document.createElement("button");
  button.type = "button";
  button.textContent = "LI";
  button.title = "LinkedIn Mode";
  button.setAttribute("aria-label", "Open LinkedIn Mode");
  button.addEventListener("click", () => {
    if (STATE.suppressLauncherClick) {
      STATE.suppressLauncherClick = false;
      return;
    }
    toggleFallbackDrawer();
  });
  button.addEventListener("mousedown", onLinkedInLauncherMouseDown);

  wrapper.appendChild(button);
  applyLinkedInLauncherPosition();
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

function installLauncherDragHandlers() {
  if (STATE.launcherHandlersInstalled) return;
  STATE.launcherHandlersInstalled = true;
  window.addEventListener("mousemove", onLinkedInLauncherMouseMove);
  window.addEventListener("mouseup", onLinkedInLauncherMouseUp);
}

function onLinkedInLauncherMouseDown(event) {
  if (event.button !== 0) return;
  STATE.drag.active = true;
  STATE.drag.startY = event.clientY;
  STATE.drag.startTop = getCurrentLinkedInLauncherTop();
  STATE.drag.moved = false;
}

function onLinkedInLauncherMouseMove(event) {
  if (!STATE.drag.active || !STATE.launcherWrapper) return;
  const deltaY = event.clientY - STATE.drag.startY;
  if (Math.abs(deltaY) > 3) {
    STATE.drag.moved = true;
  }

  const nextTop = clampLinkedInLauncherTop(STATE.drag.startTop + deltaY);
  STATE.launcherTop = nextTop;
  STATE.launcherWrapper.style.top = `${Math.round(nextTop)}px`;
  STATE.launcherWrapper.style.transform = "none";
}

function onLinkedInLauncherMouseUp() {
  if (!STATE.drag.active) return;
  STATE.drag.active = false;

  if (STATE.drag.moved) {
    STATE.suppressLauncherClick = true;
    saveLinkedInLauncherTop();
  }
}

function getCurrentLinkedInLauncherTop() {
  if (STATE.launcherTop !== null) return STATE.launcherTop;
  const rect = STATE.launcherWrapper?.getBoundingClientRect();
  if (!rect) return window.innerHeight * 0.5;
  return rect.top + rect.height * 0.5;
}

function clampLinkedInLauncherTop(value) {
  const min = 44;
  const max = Math.max(min, window.innerHeight - 44);
  return Math.min(max, Math.max(min, Number(value) || min));
}

function applyLinkedInLauncherPosition() {
  if (!STATE.launcherWrapper) return;
  if (!Number.isFinite(STATE.launcherTop)) return;
  STATE.launcherWrapper.style.top = `${Math.round(STATE.launcherTop)}px`;
  STATE.launcherWrapper.style.transform = "none";
}

function restoreLinkedInLauncherTop() {
  chrome.storage.local.get([LINKEDIN_LAUNCHER_TOP_KEY], (result) => {
    const saved = Number(result[LINKEDIN_LAUNCHER_TOP_KEY]);
    if (!Number.isFinite(saved)) return;
    STATE.launcherTop = clampLinkedInLauncherTop(saved);
    applyLinkedInLauncherPosition();
  });
}

function saveLinkedInLauncherTop() {
  if (!Number.isFinite(STATE.launcherTop)) return;
  chrome.storage.local.set({ [LINKEDIN_LAUNCHER_TOP_KEY]: Math.round(STATE.launcherTop) });
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
  if (STATE.fallback.sequences.length) {
    const recommended = String(STATE.fallback.match?.sequenceId || "").trim();
    const current = String(STATE.fallback.selectedSequenceId || "").trim();
    const hasCurrent = STATE.fallback.sequences.some((sequence) => sequence.id === current);
    const hasRecommended = STATE.fallback.sequences.some((sequence) => sequence.id === recommended);
    STATE.fallback.selectedSequenceId = hasCurrent
      ? current
      : hasRecommended
        ? recommended
        : STATE.fallback.sequences[0].id;
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
    item.className = "pa-talk-item";

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
  if (!card) return;
  const lines = String(text || "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const idx = line.indexOf(":");
      if (idx <= 0) return { label: "Info", value: line };
      return { label: line.slice(0, idx), value: line.slice(idx + 1).trim() };
    });
  renderPaKeyValueCard(card, lines);
}

function renderMatch(drawer, person, candidates, errorText) {
  const matchCard = drawer.querySelector("#paMatchCard");
  if (matchCard) {
    if (errorText) {
      renderPaKeyValueCard(matchCard, [{ label: "Status", value: errorText }]);
    } else if (!person) {
      renderPaKeyValueCard(matchCard, [{ label: "Match", value: "No direct match found." }]);
    } else {
      renderPaKeyValueCard(matchCard, [
        { label: "Matched", value: `${person.name} (#${person.id})` },
        { label: "Org", value: person.orgName || "N/A" },
        { label: "DM eligible", value: STATE.fallback.match?.dmEligible ? "Yes" : "No" },
        { label: "Current stage", value: String(STATE.fallback.match?.currentStage || 1) }
      ]);
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
  renderPaKeyValueCard(card, [{ label: "Status", value: message }]);
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

function renderPaKeyValueCard(container, rows) {
  if (!container) return;
  container.innerHTML = "";
  rows.forEach((row) => {
    const line = document.createElement("div");
    line.className = "pa-kv-line";

    const label = document.createElement("span");
    label.className = "pa-kv-label";
    label.textContent = `${row.label}: `;

    const value = document.createElement("span");
    value.textContent = row.value || "N/A";

    line.appendChild(label);
    line.appendChild(value);
    container.appendChild(line);
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
