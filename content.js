const ROOT_ID = "call-companion-root";
const STORAGE_OPTIONS_KEY = {
  autoOpenPanel: true,
  showNotes: true,
  showActivities: true
};
const QUICK_NOTES_PREFIX = "callCompanion:quickNotes:";

let rootEl = null;
let panelEl = null;
let launcherEl = null;
let currentHref = "";
let currentContext = null;
let currentPayload = null;
let activeState = {
  loading: false,
  error: "",
  context: null,
  brief: null,
  mode: "idle",
  draftStatus: ""
};
let watcherTimer = null;
let notesSaveTimer = null;
let manualPanelHidden = false;
let lastContextKey = "";

init();

function init() {
  ensureUI();
  runDetection(true);
  watcherTimer = window.setInterval(() => runDetection(false), 900);
}

function ensureUI() {
  if (document.getElementById(ROOT_ID)) return;

  rootEl = document.createElement("div");
  rootEl.id = ROOT_ID;

  panelEl = document.createElement("aside");
  panelEl.className = "cc-panel cc-panel-hidden";

  launcherEl = document.createElement("button");
  launcherEl.className = "cc-launcher";
  launcherEl.type = "button";
  launcherEl.textContent = "Call Companion";
  launcherEl.addEventListener("click", () => openPanel());

  rootEl.appendChild(panelEl);
  rootEl.appendChild(launcherEl);
  document.body.appendChild(rootEl);
}

async function runDetection(isInitial) {
  if (window.location.href === currentHref && !isInitial) return;

  currentHref = window.location.href;
  currentContext = parsePageContext(currentHref);
  const contextKey = currentContext ? `${currentContext.type}:${currentContext.id}` : "";

  if (contextKey !== lastContextKey) {
    manualPanelHidden = false;
    lastContextKey = contextKey;
  }

  if (!currentContext) {
    renderContainerShell();
    setState({
      loading: false,
      error: "Open a specific Deal or Person page to load Call Companion context.",
      context: null,
      brief: null,
      mode: "idle",
      draftStatus: ""
    });
    await renderBody();
    closePanel();
    return;
  }

  const options = await getOptions();
  renderContainerShell();

  if (options.autoOpenPanel && !manualPanelHidden) {
    openPanel();
  } else {
    closePanel();
  }

  await loadContextAndBrief(currentContext);
}

function parsePageContext(href) {
  try {
    const { pathname } = new URL(href);
    const patterns = [
      { type: "deal", regex: /^\/(?:deal|deals)\/(\d+)(?:\/|$)/i },
      { type: "person", regex: /^\/(?:person|persons)\/(\d+)(?:\/|$)/i }
    ];

    for (const pattern of patterns) {
      const match = pathname.match(pattern.regex);
      if (match) {
        return {
          type: pattern.type,
          id: Number(match[1])
        };
      }
    }

    return null;
  } catch (_error) {
    return null;
  }
}

function renderContainerShell() {
  if (!panelEl) return;

  panelEl.innerHTML = "";

  const header = document.createElement("div");
  header.className = "cc-header";

  const brand = document.createElement("div");
  brand.className = "cc-brand";

  const logo = document.createElement("img");
  logo.className = "cc-logo";
  logo.src = chrome.runtime.getURL("icons/logo-source.png");
  logo.alt = "Peak Access";

  const title = document.createElement("h2");
  title.className = "cc-title";
  title.textContent = "Call Companion";

  const controls = document.createElement("div");
  controls.className = "cc-header-controls";

  const refreshBtn = document.createElement("button");
  refreshBtn.className = "cc-btn cc-btn-secondary";
  refreshBtn.type = "button";
  refreshBtn.textContent = "Refresh";
  refreshBtn.addEventListener("click", () => {
    if (currentContext) {
      loadContextAndBrief(currentContext);
    }
  });

  const closeBtn = document.createElement("button");
  closeBtn.className = "cc-btn cc-btn-secondary";
  closeBtn.type = "button";
  closeBtn.textContent = "Hide";
  closeBtn.addEventListener("click", () => {
    manualPanelHidden = true;
    closePanel();
  });

  controls.appendChild(refreshBtn);
  controls.appendChild(closeBtn);
  brand.appendChild(logo);
  brand.appendChild(title);
  header.appendChild(brand);
  header.appendChild(controls);
  panelEl.appendChild(header);

  const body = document.createElement("div");
  body.className = "cc-body";
  body.dataset.role = "body";
  panelEl.appendChild(body);

  const footer = document.createElement("div");
  footer.className = "cc-footer";
  footer.dataset.role = "footer";
  panelEl.appendChild(footer);
}

async function loadContextAndBrief(contextRef) {
  setState({ loading: true, error: "", draftStatus: "", mode: "idle" });

  const response = await sendMessage({
    type: "GET_CONTEXT_AND_BRIEF",
    payload: {
      type: contextRef.type,
      id: contextRef.id
    }
  });

  if (!response.ok) {
    setState({ loading: false, error: response.error || "Failed to load context." });
    renderBody();
    return;
  }

  currentPayload = response.data;
  setState({
    loading: false,
    error: "",
    context: response.data.context,
    brief: response.data.brief
  });

  renderBody();
}

function setState(partial) {
  activeState = {
    ...activeState,
    ...partial
  };
}

async function renderBody() {
  const body = panelEl?.querySelector('[data-role="body"]');
  const footer = panelEl?.querySelector('[data-role="footer"]');
  if (!body) return;

  body.innerHTML = "";
  if (footer) footer.innerHTML = "";

  if (activeState.loading) {
    body.appendChild(renderInfo("Loading call context..."));
    return;
  }

  if (activeState.error) {
    body.appendChild(renderInfo(activeState.error, true));
    return;
  }

  const context = activeState.context;
  const brief = activeState.brief;
  if (!context || !brief) {
    body.appendChild(renderInfo("No context available for this page."));
    return;
  }

  const options = await getOptions();

  body.appendChild(renderContextSummary(context, options));
  body.appendChild(renderPreCall(brief.preCall));
  body.appendChild(renderCards(brief.cards));

  if (options.showNotes) {
    body.appendChild(renderNotes(context.notes || []));
  }

  if (footer) {
    footer.appendChild(renderActions(context, true));
  }

  if (activeState.mode === "answered") {
    body.appendChild(await renderQuickNotes());
  }

  if (activeState.mode === "no-answer") {
    body.appendChild(renderNoAnswer(brief.noAnswer));
  }

  if (activeState.draftStatus) {
    body.appendChild(renderInfo(activeState.draftStatus, false, "success"));
  }

  if (needsEmail(context)) {
    body.appendChild(renderPasteEmail(context));
  }

  if (options.showActivities) {
    body.appendChild(renderActivities(context.activities || []));
  }
}

function renderContextSummary(context) {
  const section = createSection("Context");
  const list = document.createElement("ul");
  list.className = "cc-list";

  const personName = context.person?.name || "Unknown person";
  const orgName = context.org?.name || context.person?.org?.name || "Unknown org";
  const dealTitle = context.deal?.title || "N/A";
  const stage = context.deal?.stageId ? `Stage ${context.deal.stageId}` : "N/A";
  const value = context.deal
    ? `${Number(context.deal.value || 0).toLocaleString()} ${context.deal.currency || ""}`.trim()
    : "N/A";
  const title = context.person?.title || "N/A";
  const owner = context.person?.ownerName || context.deal?.ownerName || "N/A";
  const website = context.org?.website || "N/A";
  const linkedIn = context.person?.linkedIn || "N/A";

  list.appendChild(listItem(`Person: ${personName}`));
  list.appendChild(listItem(`Title: ${title}`));
  list.appendChild(listItem(`Org: ${orgName}`));
  list.appendChild(listItem(`Owner: ${owner}`));
  list.appendChild(listItem(`Deal: ${dealTitle}`));
  list.appendChild(listItem(`Stage / Value: ${stage} / ${value}`));
  list.appendChild(listItem(`Org Website: ${website}`));
  list.appendChild(listItem(`LinkedIn Profile: ${linkedIn}`));

  section.appendChild(list);
  return section;
}

function renderPreCall(preCall) {
  const section = createSection("Pre-call brief");

  const oneLiner = document.createElement("p");
  oneLiner.className = "cc-one-liner";
  oneLiner.textContent = preCall?.oneLiner || "No pre-call summary generated.";
  section.appendChild(oneLiner);

  const factsTitle = document.createElement("div");
  factsTitle.className = "cc-subtitle";
  factsTitle.textContent = "Key facts";
  section.appendChild(factsTitle);

  const facts = document.createElement("ul");
  facts.className = "cc-list";
  (preCall?.keyFacts || []).forEach((fact) => facts.appendChild(listItem(fact)));
  section.appendChild(facts);

  const riskTitle = document.createElement("div");
  riskTitle.className = "cc-subtitle";
  riskTitle.textContent = "Risk flags";
  section.appendChild(riskTitle);

  const risks = document.createElement("ul");
  risks.className = "cc-list cc-risk-list";
  (preCall?.riskFlags || []).forEach((flag) => risks.appendChild(listItem(flag)));
  section.appendChild(risks);

  return section;
}

function renderCards(cards) {
  const section = createSection("Talking points");

  (cards || []).forEach((card) => {
    const cardEl = document.createElement("article");
    cardEl.className = "cc-card";

    const title = document.createElement("h4");
    title.className = "cc-card-title";
    title.textContent = card.title || "Card";

    const bullets = document.createElement("ul");
    bullets.className = "cc-list";
    (card.bullets || []).forEach((bullet) => bullets.appendChild(listItem(bullet)));

    cardEl.appendChild(title);
    cardEl.appendChild(bullets);
    section.appendChild(cardEl);
  });

  return section;
}

function renderActions(context, compact = false) {
  const section = compact ? document.createElement("section") : createSection("Actions");
  if (compact) {
    section.className = "cc-section cc-section-compact";
  }

  const row = document.createElement("div");
  row.className = "cc-actions-row";

  const answeredBtn = button("Answered", "cc-btn-primary", () => {
    setState({ mode: "answered", draftStatus: "" });
    renderBody();
  });

  const noAnswerBtn = button("No Answer", "cc-btn-primary", () => {
    setState({ mode: "no-answer", draftStatus: "" });
    renderBody();
  });

  const linkedInBtn = button("Open LinkedIn", "cc-btn-secondary", async () => {
    const url = buildLinkedInUrl(context);
    const response = await sendMessage({
      type: "OPEN_URL",
      payload: { url }
    });

    if (!response.ok) {
      setState({ draftStatus: response.error || "Could not open LinkedIn." });
      renderBody();
    }
  });

  row.appendChild(answeredBtn);
  row.appendChild(noAnswerBtn);
  row.appendChild(linkedInBtn);
  section.appendChild(row);

  return section;
}

function renderNoAnswer(noAnswer) {
  const section = createSection("No Answer follow-up");

  (noAnswer?.emailDrafts || []).forEach((draft) => {
    const card = document.createElement("article");
    card.className = "cc-card cc-email-option";

    const heading = document.createElement("div");
    heading.className = "cc-email-label";
    heading.textContent = draft.label || "Email Option";

    const subject = document.createElement("div");
    subject.className = "cc-email-subject";
    subject.textContent = `Subject: ${draft.subject || ""}`;

    const preview = document.createElement("pre");
    preview.className = "cc-email-preview";
    preview.textContent = draft.body || "";

    const to = activeState.context?.person?.primaryEmail || "";
    const createBtn = button("Create Gmail Draft", "cc-btn-primary", async () => {
      if (!to) {
        setState({ draftStatus: "Cannot create draft until contact email is available." });
        renderBody();
        return;
      }

      setState({ draftStatus: "Creating Gmail draft..." });
      renderBody();

      const response = await sendMessage({
        type: "CREATE_GMAIL_DRAFT",
        payload: {
          to,
          subject: draft.subject,
          body: draft.body
        }
      });

      if (!response.ok) {
        setState({ draftStatus: response.error || "Failed to create Gmail draft." });
      } else {
        setState({ draftStatus: `Draft created successfully. Draft ID: ${response.data.draftId}` });
      }

      renderBody();
    });

    card.appendChild(heading);
    card.appendChild(subject);
    card.appendChild(preview);
    card.appendChild(createBtn);
    section.appendChild(card);
  });

  return section;
}

async function renderQuickNotes() {
  const section = createSection("Quick notes");
  const textarea = document.createElement("textarea");
  textarea.className = "cc-notes";
  textarea.placeholder = "Capture call notes here (local only for MVP)...";

  const key = getNotesKey();
  const existing = await getLocalStorage(key);
  textarea.value = String(existing || "");

  textarea.addEventListener("input", (event) => {
    const value = event.target?.value || "";
    if (notesSaveTimer) clearTimeout(notesSaveTimer);

    notesSaveTimer = window.setTimeout(() => {
      chrome.storage.local.set({ [key]: value });
    }, 250);
  });

  section.appendChild(textarea);
  return section;
}

function renderPasteEmail(context) {
  const section = createSection("Email missing");

  const helper = document.createElement("p");
  helper.className = "cc-helper";
  helper.textContent = "Paste email from your own source and save it back to Pipedrive.";
  section.appendChild(helper);

  const input = document.createElement("input");
  input.type = "email";
  input.className = "cc-input";
  input.placeholder = "name@company.com";

  const feedback = document.createElement("div");
  feedback.className = "cc-small-text";

  const saveBtn = button("Save to Pipedrive", "cc-btn-primary", async () => {
    const email = String(input.value || "").trim();
    if (!isValidEmail(email)) {
      feedback.textContent = "Enter a valid email first.";
      feedback.className = "cc-small-text cc-error-text";
      return;
    }

    const personId = context.person?.id;
    if (!personId) {
      feedback.textContent = "Person ID missing; cannot save.";
      feedback.className = "cc-small-text cc-error-text";
      return;
    }

    feedback.textContent = "Saving email...";
    feedback.className = "cc-small-text";

    const response = await sendMessage({
      type: "SAVE_EMAIL_TO_PIPEDRIVE_PERSON",
      payload: { personId, email }
    });

    if (!response.ok) {
      feedback.textContent = response.error || "Failed to save email.";
      feedback.className = "cc-small-text cc-error-text";
      return;
    }

    feedback.textContent = "Email saved. Refreshing context...";
    feedback.className = "cc-small-text cc-success-text";

    await loadContextAndBrief(currentContext);
  });

  section.appendChild(input);
  section.appendChild(saveBtn);
  section.appendChild(feedback);

  return section;
}

function renderActivities(activities) {
  const section = createSection("Recent activities");
  if (!activities.length) {
    section.appendChild(renderInfo("No activities found."));
    return section;
  }

  const list = document.createElement("ul");
  list.className = "cc-list";

  activities.slice(0, 5).forEach((activity) => {
    list.appendChild(listItem(`${activity.subject} (${activity.dueDate || "no due date"})`));
  });

  section.appendChild(list);
  return section;
}

function renderNotes(notes) {
  const section = createSection("Recent notes");
  if (!notes.length) {
    section.appendChild(renderInfo("No notes found."));
    return section;
  }

  notes.slice(0, 5).forEach((note) => {
    const card = document.createElement("article");
    card.className = "cc-card";

    const date = document.createElement("div");
    date.className = "cc-small-text";
    date.textContent = `Added: ${formatNoteDate(note.addTime)}`;

    const content = document.createElement("div");
    content.className = "cc-note-content";
    content.textContent = note.content || "(empty note)";

    card.appendChild(date);
    card.appendChild(content);
    section.appendChild(card);
  });

  return section;
}

function createSection(titleText) {
  const section = document.createElement("section");
  section.className = "cc-section";

  const title = document.createElement("h3");
  title.className = "cc-section-title";
  title.textContent = titleText;

  section.appendChild(title);
  return section;
}

function listItem(text) {
  const item = document.createElement("li");
  item.textContent = text;
  return item;
}

function renderInfo(message, isError = false, tone = "") {
  const el = document.createElement("div");
  el.className = `cc-info ${isError ? "cc-error-text" : ""} ${tone === "success" ? "cc-success-text" : ""}`.trim();
  el.textContent = message;
  return el;
}

function button(label, className, onClick) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = `cc-btn ${className}`.trim();
  btn.textContent = label;
  btn.addEventListener("click", onClick);
  return btn;
}

function buildLinkedInQuery(context) {
  const personName = context.person?.name || "";
  const orgName = context.org?.name || context.person?.org?.name || "";
  const title = context.person?.title || "";
  const emailDomain = getEmailDomain(context.person?.primaryEmail || "");

  return [personName, title, orgName, emailDomain].filter(Boolean).join(" ");
}

function buildLinkedInUrl(context) {
  const direct = canonicalizeLinkedInUrl(context.person?.linkedIn);
  if (direct) {
    return direct;
  }

  const query = buildLinkedInQuery(context);
  if (!query) {
    return "https://www.linkedin.com/";
  }
  return `https://www.linkedin.com/search/results/people/?keywords=${encodeURIComponent(query)}`;
}

function canonicalizeLinkedInUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";

  const withProtocol = /^https?:\/\//i.test(raw) ? raw : `https://${raw.replace(/^\/+/, "")}`;

  try {
    const parsed = new URL(withProtocol);
    if (!String(parsed.hostname || "").toLowerCase().includes("linkedin.com")) return "";
    parsed.hash = "";
    return parsed.toString().replace(/\/+$/, "");
  } catch (_error) {
    return "";
  }
}

function getEmailDomain(email) {
  const value = String(email || "").trim();
  if (!value.includes("@")) return "";
  return value.split("@")[1] || "";
}

function needsEmail(context) {
  return !context?.person?.primaryEmail;
}

function getNotesKey() {
  if (!currentContext) {
    return `${QUICK_NOTES_PREFIX}unknown`;
  }

  return `${QUICK_NOTES_PREFIX}${currentContext.type}:${currentContext.id}`;
}

function getLocalStorage(key) {
  return new Promise((resolve) => {
    chrome.storage.local.get([key], (result) => resolve(result[key] || ""));
  });
}

function getOptions() {
  return new Promise((resolve) => {
    chrome.storage.sync.get(STORAGE_OPTIONS_KEY, (result) => resolve(result));
  });
}

function sendMessage(message) {
  return new Promise((resolve) => {
    chrome.runtime.sendMessage(message, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message || "Runtime error" });
        return;
      }

      resolve(response || { ok: false, error: "No response from background." });
    });
  });
}

function openPanel() {
  if (!panelEl || !launcherEl) return;
  manualPanelHidden = false;
  panelEl.classList.remove("cc-panel-hidden");
  launcherEl.classList.add("cc-launcher-hidden");
}

function closePanel() {
  if (!panelEl || !launcherEl) return;
  panelEl.classList.add("cc-panel-hidden");
  launcherEl.classList.remove("cc-launcher-hidden");
}

function hideAll() {
  closePanel();
  if (launcherEl) {
    launcherEl.classList.remove("cc-launcher-hidden");
  }
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function formatNoteDate(value) {
  if (!value) return "Unknown date";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString();
}
