const HOST_ID = "peak-access-linkedin-host";
const SHADOW_WRAPPER_CLASS = "pa-linkedin-wrapper";
const LINKEDIN_LAUNCHER_TOP_KEY = "linkedInModeLauncherTop";
const LINKEDIN_PENDING_TEMPLATE_KEY = "pa_pending_linkedin_template_text";
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
    selectedSequenceId: "",
    templates: [],
    selectedTemplateId: "",
    selectedTemplate: null,
    stage: 1
  },
  pendingInsertBusy: false
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
    return await tryInsertWithPending(text);
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
  startPendingTemplateWatcher();
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
      --pa-primary: #4cae4f;
      --pa-primary-strong: #449c47;
      --pa-accent: #00b4d8;
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
      border: 1px solid var(--pa-primary-strong);
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
      background: linear-gradient(180deg, var(--pa-primary) 0%, var(--pa-primary-strong) 100%);
      box-shadow: 0 12px 20px rgba(53, 106, 55, 0.28);
      white-space: nowrap;
    }

    .${SHADOW_WRAPPER_CLASS} button:active {
      cursor: grabbing;
    }

    .pa-drawer {
      position: fixed;
      right: 12px;
      bottom: 12px;
      top: auto;
      width: min(500px, calc(100vw - 24px));
      height: auto;
      max-height: calc(100vh - 24px);
      background: linear-gradient(180deg, #f7faff 0%, var(--pa-bg) 100%);
      border: 1px solid var(--pa-border);
      border-radius: 14px;
      box-shadow: 0 18px 34px rgba(20, 42, 71, 0.2);
      overflow: hidden;
      display: flex;
      z-index: 2147483647 !important;
      pointer-events: none;
      flex-direction: column;
      transform: translateY(16px);
      opacity: 0;
      visibility: hidden;
      transition: transform 170ms ease, opacity 170ms ease;
    }

    .pa-drawer-visible {
      transform: translateY(0);
      opacity: 1;
      visibility: visible;
      pointer-events: auto;
    }

    .pa-drawer-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 12px;
      border-bottom: 1px solid var(--pa-border);
      background: #ffffff;
      color: #0f2239;
      font-size: 16px;
      font-weight: 800;
      letter-spacing: 0.2px;
    }

    .pa-header-controls {
      display: flex;
      gap: 8px;
    }

    .pa-icon-btn {
      border: 1px solid var(--pa-border);
      border-radius: 10px;
      background: #f7faf7;
      color: #2f4f2f;
      height: 32px;
      padding: 0 10px;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      font-weight: 700;
      font-family: var(--pa-font);
      cursor: pointer;
    }

    .pa-icon-btn:hover {
      background: #edf6ed;
    }

    .pa-icon-glyph {
      font-size: 13px;
      line-height: 1;
    }

    .pa-icon-label {
      font-size: 12px;
    }

    .pa-body {
      flex: 1 1 auto;
      min-height: 0;
      overflow: hidden;
      overscroll-behavior: contain;
      padding: 0;
      font-size: 13px;
      color: var(--pa-text);
      display: flex;
      flex-direction: column;
      background: transparent;
    }

    .pa-scroll {
      flex: 1 1 auto;
      min-height: 0;
      overflow: auto;
      padding: 8px 10px 6px;
      display: grid;
      gap: 8px;
    }

    .pa-sticky-footer {
      flex: 0 0 auto;
      padding: 8px 10px 10px;
      border-top: 1px solid #e6edf5;
      background: linear-gradient(180deg, rgba(247, 250, 255, 0.96) 0%, rgba(244, 247, 251, 0.98) 100%);
      backdrop-filter: blur(8px);
      display: grid;
      gap: 8px;
    }

    .pa-section {
      display: grid;
      gap: 6px;
      min-width: 0;
      padding-bottom: 6px;
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
      font-size: 11px;
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
      padding: 7px 8px;
      font-size: 13px;
      color: var(--pa-text);
      font-family: var(--pa-font);
      min-width: 0;
    }

    .pa-textarea {
      min-height: 88px;
      resize: vertical;
      line-height: 1.45;
    }

    .pa-row {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 8px;
      min-width: 0;
    }

    .pa-row > * {
      min-width: 0;
      flex: 0 1 auto;
    }

    .pa-full {
      width: 100%;
      max-width: 100%;
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

    .pa-action-btn {
      border: 1px solid transparent;
      border-radius: 10px;
      padding: 9px 11px;
      font-size: 14px;
      font-weight: 700;
      line-height: 1.2;
      cursor: pointer;
      font-family: var(--pa-font);
      transition: background 140ms ease, transform 140ms ease, box-shadow 140ms ease;
      appearance: none;
      -webkit-appearance: none;
    }

    .pa-row .pa-action-btn {
      min-width: 104px;
    }

    .pa-row .pa-input,
    .pa-row .pa-select,
    .pa-row .pa-textarea {
      flex: 1 1 180px;
      min-width: 0;
    }

    .pa-primary {
      background: var(--pa-primary);
      border-color: var(--pa-primary-strong);
      color: #fff;
      box-shadow: none;
    }

    .pa-secondary {
      background: #f7faf7;
      color: #2f4f2f;
      border-color: var(--pa-border);
    }

    .pa-primary:hover {
      background: var(--pa-primary-strong);
    }

    .pa-secondary:hover {
      background: #edf6ed;
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
      margin-bottom: 2px;
    }

    .pa-kv-line:last-child {
      margin-bottom: 0;
    }

    .pa-kv-label {
      font-weight: 700;
    }

    .pa-match-hint {
      margin: -2px 0 2px;
      font-size: 12px;
      color: var(--pa-text-muted);
      line-height: 1.4;
    }

    .pa-candidate-card {
      border: 1px solid var(--pa-border);
      border-radius: 10px;
      background: var(--pa-surface);
      padding: 10px;
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 8px 10px;
      align-items: center;
      box-shadow: 0 6px 16px rgba(20, 42, 71, 0.06);
    }

    .pa-candidate-name {
      font-size: 14px;
      font-weight: 700;
      color: #21364f;
      line-height: 1.35;
    }

    .pa-candidate-meta {
      font-size: 12px;
      color: var(--pa-text-muted);
      line-height: 1.4;
      grid-column: 1 / 2;
    }

    .pa-confirm-match {
      justify-self: end;
      align-self: center;
      min-width: 0;
      display: inline-flex;
      align-items: center;
      gap: 6px;
      white-space: nowrap;
      grid-column: 2 / 3;
      grid-row: 1 / span 2;
    }

    .pa-confirm-icon {
      font-size: 11px;
      line-height: 1;
    }

    .pa-confirm-label {
      font-size: 12px;
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

  const controls = document.createElement("div");
  controls.className = "pa-header-controls";

  const refreshBtn = document.createElement("button");
  refreshBtn.type = "button";
  refreshBtn.id = "paRefreshBtn";
  refreshBtn.className = "pa-icon-btn";
  refreshBtn.innerHTML = '<span class="pa-icon-glyph" aria-hidden="true">↻</span><span class="pa-icon-label">Refresh</span>';

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "pa-icon-btn";
  closeBtn.innerHTML = '<span class="pa-icon-glyph" aria-hidden="true">✕</span><span class="pa-icon-label">Hide</span>';
  closeBtn.addEventListener("click", () => {
    STATE.drawerVisible = false;
    updateDrawerState();
  });

  const body = document.createElement("div");
  body.className = "pa-body";
  body.innerHTML = `
    <div class="pa-scroll">
      <section class="pa-section">
        <h3>Detected LinkedIn Context</h3>
        <div id="paContextCard" class="pa-card"></div>
      </section>
      <section class="pa-section">
        <h3>Pipedrive Match</h3>
        <div id="paMatchCard" class="pa-card"></div>
        <div class="pa-match-hint">Search for a person if no direct match is found, then confirm the right record to save the LinkedIn URL back to Pipedrive.</div>
        <div class="pa-row">
          <input id="paSearchName" class="pa-input" type="text" placeholder="Search Pipedrive person by name" />
        <button id="paSearchBtn" class="pa-action-btn pa-secondary" type="button">Search</button>
        </div>
        <div id="paCandidateList" class="pa-list"></div>
      </section>
      <section class="pa-section">
        <h3>Message Templates</h3>
        <select id="paTemplateSelect" class="pa-select"></select>
        <input id="paStageInput" class="pa-input" type="number" min="1" step="1" />
        <div id="paTemplateList" class="pa-list"></div>
      </section>
    </div>
    <div class="pa-sticky-footer">
      <section class="pa-section">
        <h3>Composer</h3>
        <textarea id="paComposer" class="pa-textarea" placeholder="Template text appears here. You can edit before posting."></textarea>
        <div class="pa-row">
          <button id="paPostBtn" class="pa-action-btn pa-primary" type="button">Post to LinkedIn</button>
          <button id="paCopyBtn" class="pa-action-btn pa-secondary" type="button">Copy</button>
        </div>
      </section>
      <section class="pa-section">
        <button id="paLogBtn" class="pa-action-btn pa-primary pa-full" type="button">Log & Advance</button>
      </section>
      <section class="pa-section">
        <h3>Status</h3>
        <div id="paStatus" class="pa-card">Idle.</div>
      </section>
    </div>
  `;

  header.appendChild(brand);
  controls.appendChild(refreshBtn);
  controls.appendChild(closeBtn);
  header.appendChild(controls);
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
  drawer.querySelector("#paTemplateSelect")?.addEventListener("change", async (event) => {
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
  drawer.querySelector("#paPostBtn")?.addEventListener("click", async () => {
    const text = getFallbackComposerText(drawer).trim();
    if (!text) {
      setFallbackStatus("Composer is empty. Choose a template or write a message first.", true);
      return;
    }
    const result = await tryInsertWithPending(text);
    if (!result.inserted) {
      setFallbackStatus(
        result.message || (result.copied
          ? "Could not post directly. Copied to clipboard."
          : "Could not post to LinkedIn composer."),
        !result.copied
      );
      return;
    }
    setFallbackStatus("Composer text posted into LinkedIn composer.");
  });
  drawer.querySelector("#paCopyBtn")?.addEventListener("click", async () => {
    const text = getFallbackComposerText(drawer).trim();
    if (!text) {
      setFallbackStatus("Composer is empty.", true);
      return;
    }
    const copied = await fallbackCopyToClipboard(text);
    setFallbackStatus(copied ? "Copied composer text." : "Copy failed.", !copied);
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
  drawer.querySelector("#paLogBtn")?.addEventListener("click", async () => {
    const personId = STATE.fallback.match?.person?.id;
    if (!personId) return setFallbackStatus("No matched person to log against.", true);

    const selectedTemplateText = STATE.fallback.selectedTemplate
      ? interpolateTemplate(STATE.fallback.selectedTemplate.body)
      : "";
    const dmText = getFallbackComposerText(drawer).trim() || getComposerText() || selectedTemplateText;
    if (!dmText) return setFallbackStatus("No message text to log.", true);

    const response = await sendRuntimeMessage({
      type: "LINKEDIN_LOG_AND_ADVANCE",
      payload: {
        personId,
        leadId: STATE.fallback.match?.leadId || "",
        profileUrl: STATE.fallback.context?.profileUrl || "",
        sequenceId: STATE.fallback.selectedSequenceId || "manual_template_flow",
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
    if (response.data?.activityWarning) {
      setFallbackStatus(
        `Logged note and advanced to stage ${STATE.fallback.stage}. Activity warning: ${response.data.activityWarning}`,
        true
      );
      return;
    }
    setFallbackStatus(`Logged activity + note and advanced to stage ${STATE.fallback.stage}.`, false);
  });
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
  await loadFallbackTemplates();
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

    await loadFallbackSequences();
    await loadFallbackTemplates();
  }
  await loadFallbackSequences();
  await loadFallbackTemplates();
  setFallbackStatus("Ready.");
}

async function loadFallbackSequences() {
  const drawer = STATE.shadowRoot?.getElementById("pa-linkedin-drawer");
  if (!drawer) return;

  const response = await sendRuntimeMessage({ type: "LINKEDIN_GET_SEQUENCES" });
  if (!response.ok) {
    STATE.fallback.sequences = [];
    renderTemplateSelect(drawer);
    setFallbackStatus(response.error || "Failed to load sequences.", true);
    return;
  }

  const sequences = Array.isArray(response.data?.sequences) ? response.data.sequences : [];
  STATE.fallback.sequences = sequences;
  if (!STATE.fallback.selectedSequenceId) {
    const matchedSequenceId = String(STATE.fallback.match?.sequenceId || "").trim();
    if (matchedSequenceId && sequences.some((item) => item.id === matchedSequenceId)) {
      STATE.fallback.selectedSequenceId = matchedSequenceId;
    }
  }
  if (!STATE.fallback.selectedSequenceId || !sequences.some((item) => item.id === STATE.fallback.selectedSequenceId)) {
    STATE.fallback.selectedSequenceId = sequences[0]?.id || "";
  }
  renderTemplateSelect(drawer);
}

async function loadFallbackTemplates() {
  const drawer = STATE.shadowRoot?.getElementById("pa-linkedin-drawer");
  if (!drawer) return;
  const stage = Number(STATE.fallback.stage || 1);

  if (!STATE.fallback.selectedSequenceId) {
    STATE.fallback.templates = [];
    STATE.fallback.selectedTemplate = null;
    renderTemplates(drawer);
    return;
  }

  const response = await sendRuntimeMessage({
    type: "LINKEDIN_GET_TEMPLATES",
    payload: {
      sequenceId: STATE.fallback.selectedSequenceId,
      stage
    }
  });

  if (!response.ok) {
    STATE.fallback.templates = [];
    STATE.fallback.selectedTemplate = null;
    renderTemplates(drawer);
    setFallbackStatus(response.error || "Failed to load templates.", true);
    return;
  }

  STATE.fallback.templates = Array.isArray(response.data?.templates) ? response.data.templates : [];
  renderTemplates(drawer);
  const selected = STATE.fallback.templates.find((item) => item.id === STATE.fallback.selectedTemplateId)
    || STATE.fallback.templates[0];
  STATE.fallback.selectedTemplate = selected || null;
  if (selected?.id) {
    STATE.fallback.selectedTemplateId = selected.id;
  }
}


function renderTemplateSelect(drawer) {
  const select = drawer.querySelector("#paTemplateSelect");
  if (!select) return;
  select.innerHTML = "";
  STATE.fallback.sequences.forEach((sequence) => {
    const option = document.createElement("option");
    option.value = sequence.id;
    option.textContent = sequence.name || sequence.id;
    select.appendChild(option);
  });
  select.value = STATE.fallback.selectedSequenceId || "";
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
    useBtn.className = "pa-action-btn pa-secondary";
    useBtn.textContent = "Use Template";
    useBtn.addEventListener("click", async () => {
      STATE.fallback.selectedTemplateId = template.id;
      STATE.fallback.selectedTemplate = template;
      const text = interpolateTemplate(template.body);
      setFallbackComposerText(drawer, text);
      setFallbackStatus(`${template.label} loaded into composer. Review, then click Post to LinkedIn.`);
    });

    item.appendChild(useBtn);
    root.appendChild(item);
  });
}

function getFallbackComposerText(drawer) {
  return String(drawer?.querySelector("#paComposer")?.value || "");
}

function setFallbackComposerText(drawer, text) {
  const composer = drawer?.querySelector("#paComposer");
  if (!composer) return;
  composer.value = String(text || "");
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
        { label: "Lead", value: STATE.fallback.match?.leadId || "N/A" },
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
    item.className = "pa-candidate-card";

    const name = document.createElement("div");
    name.className = "pa-candidate-name";
    name.textContent = candidate.name || `Person #${candidate.id}`;

    const meta = document.createElement("div");
    meta.className = "pa-candidate-meta";
    meta.textContent = `Pipedrive ID: ${candidate.id}${candidate.orgName ? ` • ${candidate.orgName}` : ""}`;

    const confirm = document.createElement("button");
    confirm.type = "button";
    confirm.className = "pa-action-btn pa-primary pa-confirm-match";
    confirm.setAttribute("data-person-id", String(candidate.id));
    confirm.innerHTML = '<span class="pa-confirm-icon" aria-hidden="true">✓</span><span class="pa-confirm-label">Confirm + Save</span>';
    confirm.addEventListener("click", async (event) => {
      event.preventDefault();
      confirm.disabled = true;
      const originalHtml = confirm.innerHTML;
      confirm.innerHTML = '<span class="pa-confirm-icon" aria-hidden="true">…</span><span class="pa-confirm-label">Saving</span>';
      try {
        await runConfirmMatch(drawer, candidate.id);
      } finally {
        confirm.disabled = false;
        confirm.innerHTML = originalHtml;
      }
    });
    item.appendChild(name);
    item.appendChild(meta);
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

function isVisible(el) {
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

async function setComposerText(text) {
  const nextText = String(text || "").trim();
  if (!nextText) return false;

  const composer = getComposerEditable(false);
  if (composer) {
    return setEditableText(composer, nextText);
  }

  const opened = await ensureComposerOrNoteOpen();
  if (opened) {
    return setEditableText(opened, nextText);
  }

  const connectNoteInput = findOpenConnectNoteInput(false);
  if (connectNoteInput) {
    return setEditableText(connectNoteInput, nextText);
  }

  return false;
}

async function tryInsertWithPending(text) {
  const nextText = String(text || "").trim();
  if (!nextText) {
    return { inserted: false, copied: false, pending: false, message: "Template text is empty." };
  }

  setPendingTemplateText(nextText);
  const inserted = await setComposerText(nextText);
  if (inserted) {
    clearPendingTemplateText(nextText);
    return { inserted: true, copied: false, pending: false };
  }

  const copied = await fallbackCopyToClipboard(nextText);
  return {
    inserted: false,
    copied,
    pending: true,
    message: copied
      ? "Waiting for LinkedIn/Sales Navigator composer. Template copied to clipboard as fallback."
      : "Waiting for LinkedIn/Sales Navigator composer to appear."
  };
}

function setPendingTemplateText(text) {
  try {
    localStorage.setItem(
      LINKEDIN_PENDING_TEMPLATE_KEY,
      JSON.stringify({ text: String(text || ""), ts: Date.now() })
    );
  } catch (_error) {
    // Ignore storage failures.
  }
}

function getPendingTemplateText() {
  try {
    const raw = localStorage.getItem(LINKEDIN_PENDING_TEMPLATE_KEY);
    if (!raw) return "";
    const parsed = JSON.parse(raw);
    const text = String(parsed?.text || "").trim();
    return text;
  } catch (_error) {
    return "";
  }
}

function clearPendingTemplateText(expectedText = "") {
  try {
    if (!expectedText) {
      localStorage.removeItem(LINKEDIN_PENDING_TEMPLATE_KEY);
      return;
    }
    const current = getPendingTemplateText();
    if (!current || current === String(expectedText || "").trim()) {
      localStorage.removeItem(LINKEDIN_PENDING_TEMPLATE_KEY);
    }
  } catch (_error) {
    // Ignore storage failures.
  }
}

function startPendingTemplateWatcher() {
  setInterval(async () => {
    if (STATE.pendingInsertBusy) return;
    const pendingText = getPendingTemplateText();
    if (!pendingText) return;

    STATE.pendingInsertBusy = true;
    try {
      const inserted = await setComposerText(pendingText);
      if (inserted) {
        clearPendingTemplateText(pendingText);
      }
    } finally {
      STATE.pendingInsertBusy = false;
    }
  }, 900);
}

async function ensureComposerOrNoteOpen() {
  if (isSalesNavigatorPage()) {
    const salesComposer = await openSalesNavigatorComposer();
    if (salesComposer) return salesComposer;
  } else {
    const linkedinComposer = await openStandardLinkedInMessageComposer();
    if (linkedinComposer) return linkedinComposer;
  }

  const connectNote = await openConnectNoteAndGetInput();
  if (connectNote) return connectNote;

  return null;
}

function setEditableText(target, nextText) {
  if (!target || !nextText) return false;
  target.focus();

  if ("value" in target) {
    target.value = nextText;
    dispatchComposerInputEvents(target, nextText);
    return readComposerText(target).includes(nextText.slice(0, 8));
  }

  if (tryExecInsert(target, nextText)) {
    const current = readComposerText(target);
    if (current && current.includes(nextText.slice(0, 8))) return true;
  }

  target.textContent = nextText;
  dispatchComposerInputEvents(target, nextText);
  return readComposerText(target).includes(nextText.slice(0, 8));
}

function isSalesNavigatorPage() {
  return /\/sales\//i.test(String(location.pathname || ""));
}

async function openStandardLinkedInMessageComposer() {
  const existing = getComposerEditable(false);
  if (existing) return existing;

  const triggers = getStandardMessageTriggers();
  for (const trigger of triggers) {
    safeClick(trigger);
    const found = await waitForComposer(1800, false);
    if (found) return found;
  }

  return null;
}

async function openSalesNavigatorComposer() {
  const existing = getComposerEditable(true);
  if (existing) return existing;

  const triggers = getSalesNavigatorMessageTriggers();
  for (const trigger of triggers) {
    safeClick(trigger);
    const found = await waitForComposer(2200, true);
    if (found) return found;
  }

  return null;
}

function getStandardMessageTriggers() {
  const allButtons = Array.from(document.querySelectorAll("main button, main a[role='button'], main a"));
  return dedupeNodes(allButtons.filter((el) => {
    if (!isVisible(el)) return false;
    const text = normalizeText(el.textContent);
    const aria = normalizeText(el.getAttribute("aria-label"));
    if (text !== "message" && text !== "send message" && aria !== "message" && aria !== "send message") return false;
    if (text.includes("inmail") || aria.includes("inmail")) return false;
    if (text.includes("sales navigator") || aria.includes("sales navigator")) return false;
    return true;
  }));
}

function getSalesNavigatorMessageTriggers() {
  const allButtons = Array.from(document.querySelectorAll("button, a[role='button'], a"));
  return dedupeNodes(allButtons.filter((el) => {
    if (!isVisible(el)) return false;
    const text = normalizeText(el.textContent);
    const aria = normalizeText(el.getAttribute("aria-label"));
    const dataControl = normalizeText(el.getAttribute("data-control-name"));
    if (text === "message" || text === "send message" || aria.includes("message") || dataControl.includes("message")) {
      if (text.includes("inmail") || aria.includes("inmail")) return false;
      return true;
    }
    return false;
  }));
}

async function openConnectNoteAndGetInput() {
  const connectInput = findOpenConnectNoteInput(false);
  if (connectInput) return connectInput;

  const connectTrigger = getConnectTrigger();
  if (!connectTrigger) return null;

  safeClick(connectTrigger);
  await sleep(450);

  const addNote = findAddNoteButton();
  if (addNote) {
    safeClick(addNote);
  }

  return await waitForConnectNoteInput(1600);
}

function getConnectTrigger() {
  const buttons = Array.from(document.querySelectorAll("button, a[role='button']"));
  return buttons.find((el) => {
    if (!isVisible(el)) return false;
    const text = normalizeText(el.textContent);
    const aria = normalizeText(el.getAttribute("aria-label"));
    return text === "connect" || text === "invite" || aria.includes("connect");
  }) || null;
}

function findAddNoteButton() {
  const buttons = Array.from(document.querySelectorAll("button, a[role='button']"));
  return buttons.find((el) => {
    if (!isVisible(el)) return false;
    const text = normalizeText(el.textContent);
    const aria = normalizeText(el.getAttribute("aria-label"));
    return text.includes("add a note") || aria.includes("add a note");
  }) || null;
}

function safeClick(el) {
  if (!el) return;
  try {
    el.dispatchEvent(new PointerEvent("pointerdown", { bubbles: true, cancelable: true, pointerType: "mouse" }));
  } catch (_error) {
    // Ignore pointer event constructor issues.
  }
  try {
    el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent("mouseup", { bubbles: true, cancelable: true, view: window }));
    el.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
  } catch (_error) {
    // Ignore mouse event issues.
  }
  try {
    if (typeof el.click === "function") el.click();
  } catch (_error) {
    // Ignore direct click issues.
  }
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function dedupeNodes(nodes) {
  return Array.from(new Set(nodes));
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
  const composer = getComposerEditable(true);
  if (!composer) return "";
  return readComposerText(composer);
}

function readComposerText(composer) {
  if (!composer) return "";
  if ("value" in composer) return String(composer.value || "").trim();
  return String(composer.innerText || composer.textContent || "").trim();
}

function dispatchComposerInputEvents(composer, text) {
  try {
    composer.dispatchEvent(new InputEvent("beforeinput", { bubbles: true, cancelable: true, data: text, inputType: "insertText" }));
  } catch (_error) {
    // Ignore if constructor is blocked by browser.
  }

  try {
    composer.dispatchEvent(new InputEvent("input", { bubbles: true, cancelable: true, data: text, inputType: "insertText" }));
  } catch (_error) {
    composer.dispatchEvent(new Event("input", { bubbles: true, cancelable: true }));
  }

  composer.dispatchEvent(new Event("change", { bubbles: true, cancelable: true }));
}

function tryExecInsert(composer, text) {
  try {
    clearContentEditable(composer);
    const lines = String(text || "").split(/\n/);
    for (let i = 0; i < lines.length; i += 1) {
      if (i > 0) {
        document.execCommand("insertLineBreak");
      }
      document.execCommand("insertText", false, lines[i]);
    }
    dispatchComposerInputEvents(composer, text);
    return true;
  } catch (_error) {
    return false;
  }
}

function findOpenConnectNoteInput(includeSalesNavigator) {
  const selectors = [
    "textarea#custom-message",
    "textarea[name='message']",
    "div[contenteditable='true'][aria-label*='add a note' i]",
    "textarea[aria-label*='add a note' i]",
    "div[contenteditable='true'][data-placeholder*='Add a note' i]"
  ];

  if (includeSalesNavigator) {
    selectors.unshift(
      ".compose-message__message-field div[contenteditable='true']",
      ".compose-message__message-field textarea",
      "textarea[placeholder*='Type your message' i]",
      "div[contenteditable='true'][aria-label*='Type your message' i]",
      "div[contenteditable='true'][data-placeholder*='Type your message' i]"
    );
  }

  for (const selector of selectors) {
    const el = document.querySelector(selector);
    if (el && isVisible(el)) return el;
  }

  return null;
}

function getComposerEditable(includeSalesNavigator = false) {
  const candidates = [
    "div.msg-form__contenteditable[contenteditable='true']",
    "div.msg-form__contenteditable[role='textbox']",
    "div[role='textbox'][contenteditable='true']",
    ".msg-form__contenteditable",
    "textarea[name='message']",
    "textarea.msg-form__contenteditable"
  ];

  if (includeSalesNavigator || isSalesNavigatorPage()) {
    candidates.unshift(
      ".compose-message__message-field div[contenteditable='true']",
      ".compose-message__message-field textarea",
      "textarea[placeholder*='Type your message' i]",
      "div[contenteditable='true'][aria-label*='Type your message' i]",
      "div[contenteditable='true'][data-placeholder*='Type your message' i]"
    );
  }

  for (const selector of candidates) {
    const el = document.querySelector(selector);
    if (el && isVisible(el)) return el;
  }

  return null;
}

async function waitForComposer(timeoutMs = 1200, includeSalesNavigator = false) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const composer = getComposerEditable(includeSalesNavigator);
    if (composer) return composer;
    await sleep(90);
  }
  return null;
}

async function waitForConnectNoteInput(timeoutMs = 1200) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const input = findOpenConnectNoteInput(false);
    if (input) return input;
    await sleep(90);
  }
  return null;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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
