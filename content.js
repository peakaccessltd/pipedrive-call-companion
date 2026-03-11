const ROOT_ID = "call-companion-root";
const STORAGE_OPTIONS_KEY = {
  autoOpenPanel: true,
  showNotes: true,
  showActivities: true
};
const QUICK_NOTES_PREFIX = "callCompanion:quickNotes:";
const LAUNCHER_TOP_KEY = "callCompanionLauncherTop";

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
  draftStatus: "",
  focusNoAnswer: false,
  focusTalkingPoints: false,
  savingAnsweredNote: false
};
let watcherTimer = null;
let notesSaveTimer = null;
let manualPanelHidden = false;
let lastContextKey = "";
let lastLeadLinkedInLaunchKey = "";
let launcherTop = null;
let dragState = {
  active: false,
  startY: 0,
  startTop: 0,
  moved: false
};
let suppressLauncherClick = false;

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
  launcherEl.textContent = "CC";
  launcherEl.title = "Call Companion";
  launcherEl.setAttribute("aria-label", "Open Call Companion");
  launcherEl.addEventListener("click", () => {
    if (suppressLauncherClick) {
      suppressLauncherClick = false;
      return;
    }
    openPanel();
  });
  launcherEl.addEventListener("mousedown", onLauncherMouseDown);

  rootEl.appendChild(panelEl);
  rootEl.appendChild(launcherEl);
  document.body.appendChild(rootEl);
  window.addEventListener("mousemove", onLauncherMouseMove);
  window.addEventListener("mouseup", onLauncherMouseUp);
  restoreLauncherTop();
}

async function runDetection(isInitial) {
  if (window.location.href === currentHref && !isInitial) return;

  currentHref = window.location.href;
  currentContext = parsePageContext(currentHref);
  const contextKey = currentContext ? `${currentContext.type}:${currentContext.id}` : "";

  if (contextKey !== lastContextKey) {
    manualPanelHidden = false;
    if (!currentContext || currentContext.type !== "lead") {
      lastLeadLinkedInLaunchKey = "";
    }
    lastContextKey = contextKey;
  }

  if (!currentContext) {
    renderContainerShell();
    setState({
      loading: false,
      error: "Open a specific Deal, Person, or Lead page to load Call Companion context.",
      context: null,
      brief: null,
      mode: "idle",
      draftStatus: "",
      focusNoAnswer: false,
      focusTalkingPoints: false,
      savingAnsweredNote: false
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

  const loadResult = await loadContextAndBrief(currentContext);
  if (loadResult?.ok) {
    await maybeAutoLaunchLinkedIn(currentContext);
  }
}

function parsePageContext(href) {
  try {
    const parsedUrl = new URL(href);
    const { pathname, search, hash } = parsedUrl;
    const patterns = [
      { type: "deal", regex: /^\/(?:deal|deals)\/(\d+)(?:\/|$)/i },
      { type: "person", regex: /^\/(?:person|persons)\/(\d+)(?:\/|$)/i },
      { type: "lead", regex: /^\/leads\/inbox\/([^/?#]+)(?:\/|$)/i },
      { type: "lead", regex: /^\/(?:lead|leads)\/([^/?#]+)(?:\/|$)/i },
      { type: "lead", regex: /^\/leads\/detail\/([^/?#]+)(?:\/|$)/i }
    ];

    for (const pattern of patterns) {
      const match = pathname.match(pattern.regex);
      if (match) {
        if (pattern.type === "lead") {
          const leadCandidate = decodeURIComponent(match[1]);
          if (!isLikelyLeadId(leadCandidate)) {
            continue;
          }
        }
        return {
          type: pattern.type,
          id: pattern.type === "lead" ? decodeURIComponent(match[1]) : Number(match[1])
        };
      }
    }

    const leadIdFromFragments = extractLeadIdFromFragments({
      pathname,
      search,
      hash
    });
    if (leadIdFromFragments) {
      return { type: "lead", id: leadIdFromFragments };
    }

    if (/^\/leads(?:\/|$)/i.test(pathname)) {
      const leadIdFromDom = detectSelectedLeadIdFromDom();
      if (leadIdFromDom) {
        return { type: "lead", id: leadIdFromDom };
      }
    }

    return null;
  } catch (_error) {
    return null;
  }
}

function extractLeadIdFromFragments({ pathname, search, hash }) {
  const candidateSources = [pathname, search, hash];
  const regexes = [
    /\/(?:lead|leads)\/([a-z0-9-]{8,})/i,
    /(?:lead_id|leadId|selectedLeadId)=([a-z0-9-]{8,})/i
  ];

  for (const source of candidateSources) {
    const value = String(source || "");
    for (const regex of regexes) {
      const match = value.match(regex);
      if (match?.[1]) {
        const id = decodeURIComponent(match[1]);
        if (isLikelyLeadId(id)) return id;
      }
    }
  }

  return "";
}

function detectSelectedLeadIdFromDom() {
  const selectors = [
    'a[aria-current="page"][href*="/lead/"]',
    'a[aria-current="true"][href*="/lead/"]',
    'a[href*="/lead/"][data-testid*="selected"]',
    'a[href*="/lead/"][class*="selected"]',
    'a[href*="/lead/"][class*="active"]'
  ];

  for (const selector of selectors) {
    const el = document.querySelector(selector);
    const id = extractLeadIdFromHref(el?.getAttribute("href"));
    if (id) return id;
  }

  const links = Array.from(document.querySelectorAll('a[href*="/lead/"]'));
  for (const link of links) {
    const id = extractLeadIdFromHref(link.getAttribute("href"));
    if (id) return id;
  }

  return "";
}

function extractLeadIdFromHref(href) {
  const raw = String(href || "");
  if (!raw) return "";
  const match = raw.match(/\/lead\/([a-z0-9-]{8,})/i) || raw.match(/\/leads\/([a-z0-9-]{8,})/i);
  if (!match?.[1]) return "";
  const id = decodeURIComponent(match[1]);
  return isLikelyLeadId(id) ? id : "";
}

function isLikelyLeadId(id) {
  const value = String(id || "").trim();
  if (!value) return false;
  if (/^[a-f0-9]{32}$/i.test(value)) return true;
  if (/^[a-f0-9-]{24,}$/i.test(value)) return true;
  return false;
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
  refreshBtn.className = "cc-icon-btn";
  refreshBtn.type = "button";
  refreshBtn.innerHTML = '<span class="cc-icon-glyph" aria-hidden="true">↻</span><span class="cc-icon-label">Refresh</span>';
  refreshBtn.setAttribute("aria-label", "Refresh");
  refreshBtn.addEventListener("click", () => {
    if (currentContext) {
      loadContextAndBrief(currentContext);
    }
  });

  const closeBtn = document.createElement("button");
  closeBtn.className = "cc-icon-btn";
  closeBtn.type = "button";
  closeBtn.innerHTML = '<span class="cc-icon-glyph" aria-hidden="true">✕</span><span class="cc-icon-label">Hide</span>';
  closeBtn.setAttribute("aria-label", "Hide");
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
    await renderBody();
    return { ok: false, error: response.error || "Failed to load context." };
  }

  currentPayload = response.data;
  setState({
    loading: false,
    error: "",
    context: response.data.context,
    brief: response.data.brief
  });

  await renderBody();
  return { ok: true, data: response.data };
}

function setState(partial) {
  activeState = {
    ...activeState,
    ...partial
  };
}

async function renderBody() {
  const body = panelEl?.querySelector('[data-role="body"]');
  if (!body) return;

  body.innerHTML = "";

  if (activeState.loading) {
    body.appendChild(renderInfo("Loading call context..."));
    return;
  }

  if (activeState.error) {
    body.appendChild(renderInfo(activeState.error, true));
    return;
  }

  const context = activeState.context;
  if (!context) {
    body.appendChild(renderInfo("No context available for this page."));
    return;
  }
  body.appendChild(renderContextSummary(context));
  body.appendChild(renderActions(context, true));

  if (activeState.mode === "answered") {
    body.appendChild(await renderQuickNotes());
  }

  if (activeState.draftStatus) {
    body.appendChild(renderInfo(activeState.draftStatus, false, "success"));
  }

  if (needsEmail(context)) {
    body.appendChild(renderPasteEmail(context));
  }
}

function renderContextSummary(context) {
  const section = createSection("Contact");
  const list = document.createElement("ul");
  list.className = "cc-list";

  const personName = context.person?.name || context.lead?.personName || context.lead?.title || "Unknown person";
  const linkedIn = getDirectLinkedInUrl(context) || "N/A";
  list.appendChild(labelValueItem("Person", personName));
  list.appendChild(labelValueItem("LinkedIn Profile", linkedIn));

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
  section.id = "cc-talking-points-section";

  (cards || []).forEach((card) => {
    const cardEl = document.createElement("article");
    cardEl.className = "cc-talk-item";

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
  }, "phone");

  const noAnswerBtn = button("No Answer", "cc-btn-primary", async () => {
    const url = getDirectLinkedInUrl(context);
    if (!url) {
      setState({
        mode: "no-answer",
        draftStatus: "No LinkedIn profile URL found on this record. Add/save the profile URL first."
      });
      renderBody();
      return;
    }
    const response = await sendMessage({
      type: "OPEN_URL",
      payload: { url, openIn: "window" }
    });
    setState({
      mode: "no-answer",
      draftStatus: response.ok ? "Opened LinkedIn in a new window." : response.error || "Could not open LinkedIn."
    });
    renderBody();
  }, "hangup");

  const linkedInBtn = button("Open LinkedIn", "cc-btn-secondary", async () => {
    const url = getDirectLinkedInUrl(context);
    if (!url) {
      setState({ draftStatus: "No LinkedIn profile URL found on this record. Add/save the profile URL first." });
      renderBody();
      return;
    }
    const response = await sendMessage({
      type: "OPEN_URL",
      payload: { url, openIn: "window" }
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
  section.id = "cc-no-answer-section";

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

function queueFocusTalkingPoints(sectionEl) {
  if (!sectionEl) return;
  window.requestAnimationFrame(() => {
    sectionEl.scrollIntoView({ behavior: "smooth", block: "start" });
    sectionEl.setAttribute("tabindex", "-1");
    sectionEl.focus();
  });
}

function queueFocusNoAnswer(sectionEl) {
  if (!sectionEl) return;
  window.requestAnimationFrame(() => {
    sectionEl.scrollIntoView({ behavior: "smooth", block: "start" });
    const firstAction = sectionEl.querySelector("button");
    if (firstAction instanceof HTMLElement) {
      firstAction.focus();
    }
  });
}

async function renderQuickNotes() {
  const section = createSection("Answered call note");
  const textarea = document.createElement("textarea");
  textarea.className = "cc-notes";
  textarea.placeholder = "Capture call notes and save to Pipedrive...";

  const key = getNotesKey();
  const existing = await getLocalStorage(key);
  textarea.value = String(existing || "");
  textarea.disabled = Boolean(activeState.savingAnsweredNote);

  textarea.addEventListener("input", (event) => {
    const value = event.target?.value || "";
    if (notesSaveTimer) clearTimeout(notesSaveTimer);

    notesSaveTimer = window.setTimeout(() => {
      chrome.storage.local.set({ [key]: value });
    }, 250);
  });

  const saveBtn = button(activeState.savingAnsweredNote ? "Saving..." : "Save Note to Pipedrive", "cc-btn-primary", async () => {
    if (activeState.savingAnsweredNote) return;
    const note = String(textarea.value || "").trim();
    if (!note) {
      setState({ draftStatus: "Add a note before saving." });
      renderBody();
      return;
    }

    setState({ savingAnsweredNote: true, draftStatus: "" });
    renderBody();

    const response = await sendMessage({
      type: "SAVE_ANSWERED_CALL_NOTE",
      payload: {
        contextType: currentContext?.type,
        contextId: currentContext?.id,
        personId: activeState.context?.person?.id || null,
        note
      }
    });

    if (!response.ok) {
      setState({ savingAnsweredNote: false, draftStatus: response.error || "Failed to save note to Pipedrive." });
      renderBody();
      return;
    }

    setState({ savingAnsweredNote: false, draftStatus: "Answered note saved to Pipedrive." });
    renderBody();
  });

  section.appendChild(textarea);
  section.appendChild(saveBtn);
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

function renderTimeline(context, options) {
  const includeNotes = Boolean(options?.showNotes);
  const includeActivities = Boolean(options?.showActivities);
  const notes = includeNotes ? context.notes || [] : [];
  const activities = includeActivities ? context.activities || [] : [];
  const items = buildTimelineItems(notes, activities);

  const section = createSection("Timeline");
  const focus = renderNextActionFocus(context, activities);
  if (focus) section.appendChild(focus);

  if (!items.length) {
    section.appendChild(renderInfo("No timeline items found."));
    return section;
  }

  items.slice(0, 5).forEach((item) => {
    const entry = document.createElement("details");
    entry.className = "cc-tl-item";

    const summary = document.createElement("summary");
    summary.className = "cc-tl-summary";

    const chip = document.createElement("span");
    chip.className = `cc-tl-chip cc-tl-chip-${item.kind}`;
    chip.textContent = item.kindLabel;

    const title = document.createElement("span");
    title.className = "cc-tl-title";
    title.textContent = item.title;

    const meta = document.createElement("span");
    meta.className = "cc-tl-time";
    meta.textContent = item.timeLabel;

    summary.appendChild(chip);
    summary.appendChild(title);
    summary.appendChild(meta);
    entry.appendChild(summary);

    const details = document.createElement("div");
    details.className = "cc-tl-details";
    details.textContent = item.details || "No additional details.";
    entry.appendChild(details);

    section.appendChild(entry);
  });

  return section;
}

function buildTimelineItems(notes, activities) {
  const noteItems = (notes || []).map((note) => {
    const clean = normalizeText(note.content || "(empty note)");
    return {
      kind: "note",
      kindLabel: "Note",
      title: truncateText(clean, 96),
      details: clean,
      timestamp: toTimestamp(note.addTime),
      timeLabel: formatNoteDate(note.addTime)
    };
  });

  const activityItems = (activities || []).map((activity) => {
    const rawType = String(activity.type || "activity").toLowerCase();
    const kind = /call|meeting|email|task|demo|follow/.test(rawType) ? "engagement" : "activity";
    const kindLabel = kind === "engagement" ? "Engagement" : toTitleCase(rawType || "activity");
    const title = `${toTitleCase(rawType || "activity")}: ${activity.subject || "No subject"}`;
    const details = [
      activity.dueDate ? `Due: ${activity.dueDate}` : "",
      activity.done ? "Status: done" : "Status: pending",
      normalizeText(activity.note || "")
    ].filter(Boolean).join("\n");

    return {
      kind,
      kindLabel,
      title: truncateText(title, 96),
      details,
      timestamp: toTimestamp(activity.dueDate),
      timeLabel: formatNoteDate(activity.dueDate)
    };
  });

  return [...noteItems, ...activityItems].sort((a, b) => b.timestamp - a.timestamp);
}

function renderNextActionFocus(context, activities) {
  const pending = (activities || [])
    .filter((item) => !item?.done)
    .sort((a, b) => toTimestamp(a?.dueDate) - toTimestamp(b?.dueDate));

  const next = pending.find(Boolean);
  if (!next) return null;

  const focus = document.createElement("article");
  focus.className = "cc-focus-card";

  const label = document.createElement("div");
  label.className = "cc-focus-head";
  label.textContent = "Next action";

  const title = document.createElement("div");
  title.className = "cc-focus-title";
  title.textContent = next.subject || `${toTitleCase(next.type || "activity")} follow-up`;

  const meta = document.createElement("div");
  meta.className = "cc-focus-meta";

  const tag = document.createElement("span");
  tag.className = "cc-focus-chip cc-focus-chip-priority";
  tag.textContent = isOverdueDate(next.dueDate) ? "Overdue" : "Next";

  const owner = document.createElement("span");
  owner.className = "cc-focus-owner";
  owner.textContent = `${context.person?.ownerName || context.deal?.ownerName || "Owner"} · ${formatNoteDate(next.dueDate)}`;

  meta.appendChild(tag);
  meta.appendChild(owner);

  focus.appendChild(label);
  focus.appendChild(title);
  focus.appendChild(meta);
  return focus;
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

function labelValueItem(label, value) {
  const item = document.createElement("li");
  const labelEl = document.createElement("span");
  labelEl.className = "cc-kv-label";
  labelEl.textContent = `${label}: `;
  const valueEl = document.createElement("span");
  valueEl.textContent = value || "N/A";
  item.appendChild(labelEl);
  item.appendChild(valueEl);
  return item;
}

function renderInfo(message, isError = false, tone = "") {
  const el = document.createElement("div");
  el.className = `cc-info ${isError ? "cc-error-text" : ""} ${tone === "success" ? "cc-success-text" : ""}`.trim();
  el.textContent = message;
  return el;
}

function button(label, className, onClick, icon = "") {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = `cc-btn ${className}`.trim();

  if (icon) {
    btn.classList.add("cc-btn-with-icon");
    const iconWrap = document.createElement("span");
    iconWrap.className = "cc-btn-icon";
    iconWrap.innerHTML = getActionIconSvg(icon);
    const labelWrap = document.createElement("span");
    labelWrap.className = "cc-btn-label";
    labelWrap.textContent = label;
    btn.appendChild(iconWrap);
    btn.appendChild(labelWrap);
  } else {
    btn.textContent = label;
  }

  btn.addEventListener("click", onClick);
  return btn;
}

function getActionIconSvg(icon) {
  if (icon === "hangup") {
    return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M3 12.5c2.2-1.5 5.2-2.3 9-2.3s6.8.8 9 2.3" /><path d="M8.5 13.2l-1.7 3.8" /><path d="M15.5 13.2l1.7 3.8" /></svg>';
  }
  return '<svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M6.6 4.5c.4-.4 1-.5 1.4-.1l2.2 1.8c.4.3.5.8.4 1.2l-.7 2.1c-.1.3 0 .7.2.9l3 3c.2.2.6.3.9.2l2.1-.7c.4-.1.9 0 1.2.4l1.8 2.2c.3.4.3 1-.1 1.4l-1.5 1.5c-.6.6-1.4.8-2.2.6-2.4-.6-5-2.2-7.3-4.5s-3.9-4.9-4.5-7.3c-.2-.8 0-1.6.6-2.2z" /></svg>';
}

function buildLinkedInQuery(context) {
  const personName = context.person?.name || "";
  const orgName = context.org?.name || context.person?.org?.name || "";
  const title = context.person?.title || "";
  const leadTitle = context.lead?.title || "";
  const emailDomain = getEmailDomain(context.person?.primaryEmail || "");

  return [personName, title, orgName, leadTitle, emailDomain].filter(Boolean).join(" ");
}

function buildLinkedInUrl(context) {
  const direct = findAnyLinkedInUrlInContext(context);
  if (direct) {
    return direct;
  }

  const query = buildLinkedInQuery(context);
  if (!query) {
    return "https://www.linkedin.com/";
  }
  return `https://www.linkedin.com/search/results/people/?keywords=${encodeURIComponent(query)}`;
}

function getDirectLinkedInUrl(context) {
  return findAnyLinkedInUrlInContext(context);
}

function findAnyLinkedInUrlInContext(context) {
  const seen = new Set();

  function walk(value, depth = 0) {
    if (depth > 5 || value === null || value === undefined) return "";
    if (typeof value === "string") {
      return canonicalizeLinkedInUrl(value);
    }

    if (typeof value !== "object") return "";
    if (seen.has(value)) return "";
    seen.add(value);

    if (Array.isArray(value)) {
      for (const item of value) {
        const found = walk(item, depth + 1);
        if (found) return found;
      }
      return "";
    }

    for (const nested of Object.values(value)) {
      const found = walk(nested, depth + 1);
      if (found) return found;
    }

    return "";
  }

  return walk(context) || "";
}

function canonicalizeLinkedInUrl(value) {
  let raw = String(value || "").trim();
  if (!raw) return "";

  const embeddedPublicUrl = extractPublicLinkedInProfileUrl(raw);
  if (embeddedPublicUrl) return embeddedPublicUrl;

  raw = raw.replace(/\\\//g, "/").replace(/^"+|"+$/g, "");
  if (!raw) return "";

  if (/^\/in\//i.test(raw)) {
    raw = `https://www.linkedin.com${raw}`;
  } else if (/^in\//i.test(raw)) {
    raw = `https://www.linkedin.com/${raw}`;
  } else if (/^[a-z0-9.-]*linkedin\.com\//i.test(raw) && !/^https?:\/\//i.test(raw)) {
    raw = `https://${raw}`;
  } else if (!/^https?:\/\//i.test(raw)) {
    return "";
  }

  try {
    const parsed = new URL(raw);
    const host = String(parsed.hostname || "").toLowerCase();
    if (!host.includes("linkedin.com")) return "";
    const path = String(parsed.pathname || "");
    if (!/^\/in\/[^/]+/i.test(path)) return "";
    parsed.hash = "";
    parsed.search = "";
    parsed.pathname = path.match(/^\/in\/[^/]+/i)?.[0] || path;
    return parsed.toString().replace(/\/+$/, "");
  } catch (_error) {
    return "";
  }
}

function extractPublicLinkedInProfileUrl(rawValue) {
  const input = String(rawValue || "");
  if (!input) return "";
  let decoded = input;
  try {
    decoded = decodeURIComponent(input);
  } catch (_error) {
    decoded = input;
  }
  const match = decoded.match(/https?:\/\/(?:[a-z0-9-]+\.)?linkedin\.com\/in\/[a-z0-9-_%]+/i);
  if (!match?.[0]) return "";

  try {
    const parsed = new URL(match[0]);
    parsed.hash = "";
    parsed.search = "";
    parsed.pathname = (parsed.pathname.match(/^\/in\/[^/]+/i)?.[0] || parsed.pathname).replace(/\/+$/, "");
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

  return `${QUICK_NOTES_PREFIX}${currentContext.type}:${String(currentContext.id)}`;
}

async function maybeAutoLaunchLinkedIn(contextRef) {
  if (!contextRef || contextRef.type !== "lead") return;
  const context = activeState.context;
  if (!context) return;

  const launchKey = `${contextRef.type}:${String(contextRef.id)}`;
  if (lastLeadLinkedInLaunchKey === launchKey) return;

  const url = getDirectLinkedInUrl(context);
  if (!url) return;

  const response = await sendMessage({
    type: "OPEN_URL",
    payload: { url, openIn: "window" }
  });

  if (response.ok) {
    lastLeadLinkedInLaunchKey = launchKey;
  }
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

function onLauncherMouseDown(event) {
  if (!launcherEl || event.button !== 0) return;

  dragState.active = true;
  dragState.startY = event.clientY;
  dragState.startTop = getCurrentLauncherTop();
  dragState.moved = false;
}

function onLauncherMouseMove(event) {
  if (!dragState.active || !launcherEl) return;

  const deltaY = event.clientY - dragState.startY;
  if (Math.abs(deltaY) > 3) {
    dragState.moved = true;
  }

  const nextTop = clampLauncherTop(dragState.startTop + deltaY);
  launcherTop = nextTop;
  launcherEl.style.top = `${Math.round(nextTop)}px`;
  launcherEl.style.transform = "none";
}

function onLauncherMouseUp() {
  if (!dragState.active) return;
  dragState.active = false;

  if (dragState.moved) {
    suppressLauncherClick = true;
    saveLauncherTop();
  }
}

function getCurrentLauncherTop() {
  if (launcherTop !== null) return launcherTop;
  const rect = launcherEl?.getBoundingClientRect();
  if (!rect) return window.innerHeight * 0.5;
  return rect.top + rect.height * 0.5;
}

function clampLauncherTop(value) {
  const min = 44;
  const max = Math.max(min, window.innerHeight - 44);
  return Math.min(max, Math.max(min, Number(value) || min));
}

function restoreLauncherTop() {
  chrome.storage.local.get([LAUNCHER_TOP_KEY], (result) => {
    const saved = Number(result[LAUNCHER_TOP_KEY]);
    if (!Number.isFinite(saved) || !launcherEl) return;

    launcherTop = clampLauncherTop(saved);
    launcherEl.style.top = `${Math.round(launcherTop)}px`;
    launcherEl.style.transform = "none";
  });
}

function saveLauncherTop() {
  if (!Number.isFinite(launcherTop)) return;
  chrome.storage.local.set({ [LAUNCHER_TOP_KEY]: Math.round(launcherTop) });
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function formatNoteDate(value) {
  if (!value) return "Unknown date";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    year: "numeric"
  });
}

function toTimestamp(value) {
  if (!value) return 0;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function truncateText(value, maxLen) {
  const text = String(value || "").trim();
  if (text.length <= maxLen) return text;
  return `${text.slice(0, maxLen - 1).trim()}…`;
}

function normalizeText(value) {
  return String(value || "")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function toTitleCase(value) {
  return String(value || "")
    .split(/[\s_-]+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

function isOverdueDate(value) {
  if (!value) return false;
  const date = new Date(`${value}T23:59:59`);
  if (Number.isNaN(date.getTime())) return false;
  return date.getTime() < Date.now();
}
