import { base64UrlEncode, buildRfc2822Message } from "./lib/email.js";
import { APP_CONFIG, DEFAULT_STORAGE } from "./src/config.js";
import {
  canonicalizeLinkedInUrl,
  createLinkedInDmActivity,
  createLinkedInDmLog,
  fetchPersonById,
  searchPersonByEmail,
  searchPersonByLinkedInUrl,
  searchPersonByName,
  updatePersonFields
} from "./src/lib/pipedrive.js";

const DEFAULT_OPTIONS = {
  ...DEFAULT_STORAGE
};

const DEFAULT_STAGE_TEMPLATES = {
  default: [
    {
      label: "Email #1",
      subject: "Quick follow-up on {{dealTitle}}",
      body: "Hi {{personFirstName}},\n\nSorry we missed each other. I called to align on next steps for {{dealTitle}}.\n\nIf helpful, I can send a concise summary and propose two times for a quick sync.\n\nBest,"
    },
    {
      label: "Email #2",
      subject: "A practical idea for {{orgName}}",
      body: "Hi {{personFirstName}},\n\nI could not catch you live. Based on what we discussed, I put together a practical next-step recommendation for {{orgName}}.\n\nWant me to send it over before we reconnect?\n\nBest,"
    },
    {
      label: "Email #3",
      subject: "Should we reschedule?",
      body: "Hi {{personFirstName}},\n\nNo problem we missed each other. Are you open to a 10-minute reconnect this week regarding {{dealTitle}}?\n\nShare two windows and I will lock it in.\n\nBest,"
    }
  ],
  "stage:1": [
    {
      label: "Email #1",
      subject: "Intro follow-up for {{dealTitle}}",
      body: "Hi {{personFirstName}},\n\nTried to reach you earlier. I wanted to confirm your top priority around {{dealTitle}} and whether this is still active on your side.\n\nHappy to keep it short and focused.\n\nBest,"
    },
    {
      label: "Email #2",
      subject: "Worth a quick intro call?",
      body: "Hi {{personFirstName}},\n\nIf timing is not right, no worries. If this is still relevant for {{orgName}}, I can suggest a simple first step.\n\nWould a brief call next week help?\n\nBest,"
    },
    {
      label: "Email #3",
      subject: "Close loop on this?",
      body: "Hi {{personFirstName}},\n\nChecking whether you want to keep {{dealTitle}} open for now. If yes, I can share options and we can decide quickly.\n\nBest,"
    }
  ]
};

const FALLBACK_LINKEDIN_SEQUENCE = {
  id: "default-linkedin-outreach",
  name: "Default LinkedIn Outreach",
  description: "Human-in-the-loop LinkedIn DM sequence aligned to deal stage progression.",
  recommended_start_stage: 1,
  templates: [
    {
      id: "li-stage1-open",
      stage: 1,
      label: "Stage 1 Intro",
      body: "Hi {{personFirstName}} - thanks for connecting. I work with teams like {{orgName}} on {{useCaseOrDeal}}. Open to a quick exchange this week?"
    },
    {
      id: "li-stage2-value",
      stage: 2,
      label: "Stage 2 Value Follow-up",
      body: "Hi {{personFirstName}}, following up with one practical idea for {{orgName}}: {{valueHook}}. If useful, I can share a concise 2-step plan."
    },
    {
      id: "li-stage3-proof",
      stage: 3,
      label: "Stage 3 Proof + CTA",
      body: "Hi {{personFirstName}}, we recently helped a similar team reduce time-to-outcome. If this is still relevant, would a 10-minute chat next week help?"
    }
  ]
};
let lastKnownPipedriveOrigin = "https://app.pipedrive.com";

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  handleMessage(message, sender)
    .then((data) => sendResponse({ ok: true, data }))
    .catch((error) => {
      sendResponse({ ok: false, error: toUserError(error) });
    });

  return true;
});

if (chrome.action?.onClicked) {
  chrome.action.onClicked.addListener(async (tab) => {
    if (!tab?.id) return;
    if (chrome.sidePanel?.open) {
      await chrome.sidePanel.open({ tabId: tab.id });
    }
  });
}

async function handleMessage(message, sender) {
  const msgType = message?.type;

  if (msgType === "GET_CONTEXT_AND_BRIEF") {
    return handleGetContextAndBrief(message, sender);
  }

  if (msgType === "CREATE_GMAIL_DRAFT") {
    return handleCreateGmailDraft(message?.payload || {});
  }

  if (msgType === "OPEN_URL") {
    return handleOpenUrl(message?.payload || {});
  }

  if (msgType === "SAVE_EMAIL_TO_PIPEDRIVE_PERSON") {
    return handleSaveEmail(message?.payload || {}, sender);
  }
  if (msgType === "SAVE_ANSWERED_CALL_NOTE") {
    return handleSaveAnsweredCallNote(message?.payload || {}, sender);
  }

  if (msgType === "OPEN_LINKEDIN_SIDE_PANEL") {
    return handleOpenLinkedInSidePanel(sender);
  }

  if (msgType === "LINKEDIN_GET_CONTEXT") {
    return handleLinkedInGetContext();
  }

  if (msgType === "LINKEDIN_MATCH_PERSON") {
    return handleLinkedInMatchPerson(message?.payload || {});
  }

  if (msgType === "LINKEDIN_SEARCH_PERSONS") {
    return handleLinkedInSearchPersons(message?.payload || {});
  }

  if (msgType === "LINKEDIN_CONFIRM_MATCH") {
    return handleLinkedInConfirmMatch(message?.payload || {});
  }

  if (msgType === "LINKEDIN_GET_SEQUENCES") {
    return handleLinkedInGetSequences();
  }

  if (msgType === "LINKEDIN_GET_TEMPLATES") {
    return handleLinkedInGetTemplates(message?.payload || {});
  }

  if (msgType === "LINKEDIN_GET_TALKING_POINTS") {
    return handleLinkedInGetTalkingPoints(message?.payload || {});
  }

  if (msgType === "LINKEDIN_INSERT_TEMPLATE") {
    return forwardToActiveLinkedInTab({
      type: "LINKEDIN_SET_COMPOSER_TEXT",
      payload: { text: String(message?.payload?.text || "") }
    });
  }

  if (msgType === "LINKEDIN_GET_COMPOSER_TEXT") {
    return forwardToActiveLinkedInTab({ type: "LINKEDIN_GET_COMPOSER_TEXT" });
  }

  if (msgType === "LINKEDIN_COPY_TEXT") {
    return forwardToActiveLinkedInTab({
      type: "LINKEDIN_COPY_TEXT",
      payload: { text: String(message?.payload?.text || "") }
    });
  }

  if (msgType === "LINKEDIN_LOG_AND_ADVANCE") {
    return handleLinkedInLogAndAdvance(message?.payload || {});
  }

  throw new Error("Unsupported message type.");
}

async function handleGetContextAndBrief(message, sender) {
  const payload = message?.payload || {};
  const contextType = payload.type;
  const rawId = payload.id;

  if (!["deal", "person", "lead"].includes(contextType)) {
    throw new Error("Invalid context payload.");
  }

  const id = contextType === "lead" ? String(rawId || "").trim() : Number(rawId);
  if (contextType === "lead") {
    if (!id) {
      throw new Error("Invalid lead id.");
    }
  } else if (!Number.isFinite(id) || id <= 0) {
    throw new Error("Invalid context payload.");
  }

  const options = await getOptions();

  if (!options.apiToken) {
    throw new Error("Pipedrive API token missing. Open extension options and save your token.");
  }

  const baseOrigin = getBaseOrigin(sender?.url || payload.baseOrigin);
  lastKnownPipedriveOrigin = baseOrigin;
  let context;

  if (contextType === "deal") {
    context = await buildDealContext(id, baseOrigin, options.apiToken, options);
  } else if (contextType === "person") {
    context = await buildPersonContext(id, baseOrigin, options.apiToken, options);
  } else {
    context = await buildLeadContext(id, baseOrigin, options.apiToken, options);
  }

  return {
    context,
    brief: buildDeterministicBrief(context, options)
  };
}

async function handleCreateGmailDraft(payload) {
  const to = String(payload.to || "").trim();
  const subject = String(payload.subject || "").trim();
  const body = String(payload.body || "").trim();

  if (!to || !subject || !body) {
    throw new Error("To, subject, and body are required to create a draft.");
  }

  const token = await getGmailToken();
  const rawMessage = buildRfc2822Message({ to, subject, body });
  const raw = base64UrlEncode(rawMessage);

  const response = await fetch("https://gmail.googleapis.com/gmail/v1/users/me/drafts", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${token}`
    },
    body: JSON.stringify({
      message: { raw }
    })
  });

  if (!response.ok) {
    const data = await safeJson(response);
    const apiError = data?.error?.message || "Failed to create Gmail draft.";

    if (response.status === 401 || response.status === 403) {
      throw new Error(`Gmail authorization failed (${response.status}). ${apiError}`);
    }

    throw new Error(`Gmail API error (${response.status}): ${apiError}`);
  }

  const created = await response.json();

  return {
    draftId: created?.id || "unknown",
    messageId: created?.message?.id || null
  };
}

async function handleOpenUrl(payload) {
  const normalized = normalizeOpenUrl(payload.url);
  if (!normalized) {
    throw new Error("URL is required.");
  }

  const openIn = String(payload?.openIn || "tab").toLowerCase();

  try {
    if (openIn === "window") {
      const win = await chrome.windows.create({
        url: normalized,
        focused: true
      });
      return { opened: true, url: normalized, openIn: "window", windowId: win?.id || null };
    }

    const tab = await chrome.tabs.create({ url: normalized });
    return { opened: true, url: normalized, openIn: "tab", tabId: tab?.id || null };
  } catch (_error) {
    throw new Error(openIn === "window" ? "Failed to open the URL in a new window." : "Failed to open the URL in a new tab.");
  }
}

async function handleSaveEmail(payload, sender) {
  const personId = Number(payload.personId);
  const email = String(payload.email || "").trim();

  if (!Number.isFinite(personId) || personId <= 0) {
    throw new Error("Invalid person id.");
  }

  if (!isValidEmail(email)) {
    throw new Error("Please provide a valid email address.");
  }

  const options = await getOptions();
  if (!options.apiToken) {
    throw new Error("Pipedrive API token missing. Open extension options and save your token.");
  }

  const baseOrigin = getBaseOrigin(sender?.url || payload.baseOrigin);
  lastKnownPipedriveOrigin = baseOrigin;

  const updated = await fetchPipedrive({
    path: `/api/v1/persons/${personId}`,
    method: "PUT",
    baseOrigin,
    apiToken: options.apiToken,
    body: {
      email: [{ value: email, primary: true, label: "work" }]
    }
  });

  return {
    person: pickPersonFields(updated, options)
  };
}

async function handleSaveAnsweredCallNote(payload, sender) {
  const contextType = String(payload.contextType || "").trim().toLowerCase();
  const contextIdRaw = payload.contextId;
  const personId = Number(payload.personId);
  const note = String(payload.note || "").trim();

  if (!["deal", "person", "lead"].includes(contextType)) {
    throw new Error("Invalid context type for note save.");
  }
  if (!note) {
    throw new Error("Note content is required.");
  }

  const contextId = contextType === "lead" ? String(contextIdRaw || "").trim() : Number(contextIdRaw);
  if (contextType === "lead") {
    if (!contextId) throw new Error("Invalid lead id.");
  } else if (!Number.isFinite(contextId) || contextId <= 0) {
    throw new Error("Invalid context id.");
  }

  const options = await getOptions();
  if (!options.apiToken) {
    throw new Error("Pipedrive API token missing. Open extension options and save your token.");
  }

  const baseOrigin = getBaseOrigin(sender?.url || payload.baseOrigin);
  lastKnownPipedriveOrigin = baseOrigin;

  const stamp = new Date().toLocaleString();
  const lines = [
    "Call Outcome: Answered",
    `Timestamp: ${stamp}`,
    "",
    note
  ];
  const content = lines.map((line) => escapeHtml(line)).join("<br>");

  const notePayload = { content };
  if (Number.isFinite(personId) && personId > 0) {
    notePayload.person_id = personId;
  }
  if (contextType === "deal") {
    notePayload.deal_id = contextId;
  } else if (contextType === "person") {
    notePayload.person_id = Number.isFinite(personId) && personId > 0 ? personId : contextId;
  } else if (contextType === "lead") {
    notePayload.lead_id = String(contextId);
  }

  const saved = await fetchPipedrive({
    path: "/api/v1/notes",
    method: "POST",
    baseOrigin,
    apiToken: options.apiToken,
    body: notePayload
  });

  return {
    noteId: saved?.id || null
  };
}

async function buildDealContext(dealId, baseOrigin, apiToken, options) {
  const dealRaw = await fetchPipedrive({
    path: `/api/v1/deals/${dealId}`,
    baseOrigin,
    apiToken
  });
  const deal = pickDealFields(dealRaw);

  const personId = normalizeEntityId(deal.person);
  const orgId = normalizeEntityId(deal.org);

  const [personRaw, orgRaw, activitiesRaw, notesRaw, openDealsRaw, dealFieldsRaw, personFieldsRaw] = await Promise.all([
    personId
      ? fetchPipedrive({ path: `/api/v1/persons/${personId}`, baseOrigin, apiToken })
      : Promise.resolve(null),
    orgId ? fetchPipedrive({ path: `/api/v1/organizations/${orgId}`, baseOrigin, apiToken }) : Promise.resolve(null),
    options.showActivities
      ? fetchPipedrive({
          path: `/api/v1/deals/${dealId}/activities?limit=20&sort=due_date%20DESC`,
          baseOrigin,
          apiToken,
          dataIsArray: true
        })
      : Promise.resolve([]),
    options.showNotes
      ? fetchPipedrive({
          path: `/api/v1/notes?deal_id=${dealId}&limit=20&sort=add_time%20DESC`,
          baseOrigin,
          apiToken,
          dataIsArray: true
        })
      : Promise.resolve([]),
    personId
      ? safeFetchPipedriveOr({
          fallback: [],
          config: {
            path: `/api/v1/persons/${personId}/deals?status=open&limit=20&sort=update_time%20DESC`,
            baseOrigin,
            apiToken,
            dataIsArray: true
          }
        })
      : Promise.resolve([]),
    safeFetchPipedriveOr({
      fallback: [],
      config: {
        path: "/api/v1/dealFields?limit=500",
        baseOrigin,
        apiToken,
        dataIsArray: true
      }
    }),
    safeFetchPipedriveOr({
      fallback: [],
      config: {
        path: "/api/v1/personFields?limit=500",
        baseOrigin,
        apiToken,
        dataIsArray: true
      }
    })
  ]);

  const openDeals = Array.isArray(openDealsRaw) ? openDealsRaw.map(pickDealFields) : [];

  return {
    type: "deal",
    deal,
    person: pickPersonFields(personRaw, options),
    org: pickOrgFields(orgRaw),
    openDeals,
    activities: Array.isArray(activitiesRaw) ? activitiesRaw.map(pickActivityFields) : [],
    notes: Array.isArray(notesRaw) ? notesRaw.map(pickNoteFields) : [],
    customInsights: buildCustomInsights({
      dealRaw,
      personRaw,
      dealFieldDefs: dealFieldsRaw,
      personFieldDefs: personFieldsRaw
    })
  };
}

async function buildPersonContext(personId, baseOrigin, apiToken, options) {
  const personRaw = await fetchPipedrive({
    path: `/api/v1/persons/${personId}`,
    baseOrigin,
    apiToken
  });
  const person = pickPersonFields(personRaw, options);

  const [orgRaw, dealsRaw, activitiesRaw, notesRaw, dealFieldsRaw, personFieldsRaw] = await Promise.all([
    person.org?.id
      ? fetchPipedrive({ path: `/api/v1/organizations/${person.org.id}`, baseOrigin, apiToken })
      : Promise.resolve(null),
    fetchPipedrive({
      path: `/api/v1/persons/${personId}/deals?status=open&limit=20&sort=update_time%20DESC`,
      baseOrigin,
      apiToken,
      dataIsArray: true
    }),
    options.showActivities
      ? fetchPipedrive({
          path: `/api/v1/persons/${personId}/activities?limit=20&sort=due_date%20DESC`,
          baseOrigin,
          apiToken,
          dataIsArray: true
        })
      : Promise.resolve([]),
    options.showNotes
      ? fetchPipedrive({
          path: `/api/v1/notes?person_id=${personId}&limit=20&sort=add_time%20DESC`,
          baseOrigin,
          apiToken,
          dataIsArray: true
        })
      : Promise.resolve([]),
    safeFetchPipedriveOr({
      fallback: [],
      config: {
        path: "/api/v1/dealFields?limit=500",
        baseOrigin,
        apiToken,
        dataIsArray: true
      }
    }),
    safeFetchPipedriveOr({
      fallback: [],
      config: {
        path: "/api/v1/personFields?limit=500",
        baseOrigin,
        apiToken,
        dataIsArray: true
      }
    })
  ]);

  return {
    type: "person",
    person,
    org: pickOrgFields(orgRaw),
    openDeals: Array.isArray(dealsRaw) ? dealsRaw.map(pickDealFields) : [],
    activities: Array.isArray(activitiesRaw) ? activitiesRaw.map(pickActivityFields) : [],
    notes: Array.isArray(notesRaw) ? notesRaw.map(pickNoteFields) : [],
    customInsights: buildCustomInsights({
      dealRaw: null,
      personRaw,
      dealFieldDefs: dealFieldsRaw,
      personFieldDefs: personFieldsRaw
    })
  };
}

async function buildLeadContext(leadId, baseOrigin, apiToken, options) {
  const encodedLeadId = encodeURIComponent(String(leadId));
  let leadRaw = null;
  try {
    leadRaw = await fetchPipedrive({
      path: `/api/v1/leads/${encodedLeadId}`,
      baseOrigin,
      apiToken
    });
  } catch (error) {
    const msg = String(error?.message || "");
    if (msg.includes("resource not found")) {
      throw new Error(`Lead not found in Pipedrive for id '${String(leadId)}'.`);
    }
    throw error;
  }
  const lead = pickLeadFields(leadRaw);
  if (!lead) {
    throw new Error("Lead not found for this page context.");
  }

  const personId = normalizeEntityId(lead.person);
  const orgId = normalizeEntityId(lead.org);

  const [personRaw, orgRaw, activitiesRaw, notesRaw, dealsRaw, dealFieldsRaw, personFieldsRaw] = await Promise.all([
    personId
      ? fetchPipedrive({ path: `/api/v1/persons/${personId}`, baseOrigin, apiToken })
      : Promise.resolve(null),
    orgId ? fetchPipedrive({ path: `/api/v1/organizations/${orgId}`, baseOrigin, apiToken }) : Promise.resolve(null),
    options.showActivities
      ? safeFetchPipedriveOr({
          fallback: [],
          config: {
            path: `/api/v1/activities?lead_id=${encodedLeadId}&limit=20&sort=due_date%20DESC`,
            baseOrigin,
            apiToken,
            dataIsArray: true
          }
        })
      : Promise.resolve([]),
    options.showNotes
      ? safeFetchPipedriveOr({
          fallback: [],
          config: {
            path: `/api/v1/notes?lead_id=${encodedLeadId}&limit=20&sort=add_time%20DESC`,
            baseOrigin,
            apiToken,
            dataIsArray: true
          }
        })
      : Promise.resolve([]),
    personId
      ? safeFetchPipedriveOr({
          fallback: [],
          config: {
            path: `/api/v1/persons/${personId}/deals?status=open&limit=20&sort=update_time%20DESC`,
            baseOrigin,
            apiToken,
            dataIsArray: true
          }
        })
      : Promise.resolve([]),
    safeFetchPipedriveOr({
      fallback: [],
      config: {
        path: "/api/v1/dealFields?limit=500",
        baseOrigin,
        apiToken,
        dataIsArray: true
      }
    }),
    safeFetchPipedriveOr({
      fallback: [],
      config: {
        path: "/api/v1/personFields?limit=500",
        baseOrigin,
        apiToken,
        dataIsArray: true
      }
    })
  ]);

  const person =
    pickPersonFields(personRaw, options) ||
    (lead.person
      ? {
          id: normalizeEntityId(lead.person),
          name: lead.person.name || "Unknown person",
          firstName: firstToken(lead.person.name) || "",
          lastName: lastToken(lead.person.name) || "",
          title: "",
          ownerName: lead.ownerName || "",
          primaryEmail: "",
          primaryPhone: "",
          emails: [],
          phones: [],
          linkedIn: lead.linkedIn || "",
          org: normalizeEntity(lead.org, lead.org?.name),
          updateTime: lead.updateTime || null
        }
      : null);
  if (person && !person.linkedIn && lead.linkedIn) {
    person.linkedIn = lead.linkedIn;
  }
  const org = pickOrgFields(orgRaw);
  const dealLikeLead = {
    id: lead.id,
    title: lead.title,
    value: lead.value,
    currency: lead.currency,
    stageId: null,
    pipelineId: null,
    status: lead.status || "open",
    ownerName: lead.ownerName || person?.ownerName || "Unknown owner",
    updateTime: lead.updateTime || null,
    person:
      (lead.person || person?.id)
        ? { id: person?.id || personId, name: person?.name || lead.person?.name || null }
        : null,
    org:
      (lead.org || org?.id)
        ? { id: org?.id || orgId, name: org?.name || lead.org?.name || null }
        : null
  };

  return {
    type: "lead",
    lead,
    deal: dealLikeLead,
    person,
    org,
    openDeals: Array.isArray(dealsRaw) ? dealsRaw.map(pickDealFields) : [],
    activities: Array.isArray(activitiesRaw) ? activitiesRaw.map(pickActivityFields) : [],
    notes: Array.isArray(notesRaw) ? notesRaw.map(pickNoteFields) : [],
    customInsights: buildCustomInsights({
      dealRaw: null,
      personRaw,
      dealFieldDefs: dealFieldsRaw,
      personFieldDefs: personFieldsRaw
    })
  };
}

function pickDealFields(deal) {
  if (!deal) return null;

  return {
    id: Number(deal.id) || null,
    title: deal.title || "Untitled deal",
    value: Number(deal.value) || 0,
    currency: deal.currency || "",
    stageId: deal.stage_id || null,
    pipelineId: deal.pipeline_id || null,
    status: deal.status || "unknown",
    ownerName: deal.owner_name || deal.owner_id?.name || "Unknown owner",
    updateTime: deal.update_time || null,
    person: normalizeEntity(deal.person_id, deal.person_name),
    org: normalizeEntity(deal.org_id, deal.org_name)
  };
}

function pickLeadFields(lead) {
  if (!lead) return null;

  return {
    id: String(lead.id || ""),
    title: lead.title || "Untitled lead",
    value: Number(lead.value) || 0,
    currency: lead.currency || "",
    status: lead.status || "open",
    ownerName: lead.owner_id?.name || lead.owner_name || "Unknown owner",
    addTime: lead.add_time || null,
    updateTime: lead.update_time || null,
    linkedIn: findLinkedInUrlFromEntity(lead),
    person: normalizeEntity(lead.person_id, lead.person_name),
    org: normalizeEntity(lead.organization_id || lead.org_id, lead.org_name || lead.organization_name)
  };
}

function pickPersonFields(person, options = {}) {
  if (!person) return null;

  const primaryEmail = extractPrimaryValue(person.email);
  const primaryPhone = extractPrimaryValue(person.phone);

  return {
    id: Number(person.id) || null,
    name: person.name || "Unknown person",
    firstName: person.first_name || firstToken(person.name) || "",
    lastName: person.last_name || lastToken(person.name) || "",
    title: person.job_title || person.title || "",
    ownerName: person.owner_id?.name || person.owner_name || "",
    primaryEmail,
    primaryPhone,
    emails: normalizeValueList(person.email),
    phones: normalizeValueList(person.phone),
    linkedIn: findLinkedInUrl(person, options),
    org: normalizeEntity(person.org_id, person.org_name),
    updateTime: person.update_time || null
  };
}

function pickOrgFields(org) {
  if (!org) return null;

  return {
    id: Number(org.id) || null,
    name: org.name || "Unknown organization",
    website: org.website || "",
    address: compactAddress(org)
  };
}

function pickActivityFields(activity) {
  if (!activity) return null;

  return {
    id: Number(activity.id) || null,
    type: activity.type || "activity",
    subject: activity.subject || "No subject",
    dueDate: activity.due_date || null,
    done: Boolean(activity.done),
    note: activity.note || ""
  };
}

function pickNoteFields(note) {
  if (!note) return null;

  const plain = String(note.content || "")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  return {
    id: Number(note.id) || null,
    content: plain,
    addTime: note.add_time || null
  };
}

function buildDeterministicBrief(context, options) {
  const person = context.person;
  const org = context.org;
  const deal = context.deal;
  const openDeals = context.openDeals || [];
  const activities = context.activities || [];
  const notes = context.notes || [];

  const personName = person?.name || "this contact";
  const orgName = org?.name || person?.org?.name || "their organization";
  const hasDeal = Boolean(deal);

  const oneLiner = hasDeal
    ? `${personName} is tied to ${deal.title} at ${orgName}. Focus this call on next concrete commitment.`
    : `${personName} at ${orgName} has ${openDeals.length} open deal(s). Focus on priority and next step.`;

  const latestActivity = activities[0];
  const latestNote = notes[0];
  const signals = buildContextSignals(context);

  const keyFacts = [
    hasDeal
      ? `Deal status: ${deal.status}, value: ${formatMoney(deal.value, deal.currency)}, owner: ${deal.ownerName}.`
      : `Open deals: ${openDeals.length}. Top deal: ${openDeals[0]?.title || "None"}.`,
    `Recent activity: ${latestActivity ? `${latestActivity.subject} (${latestActivity.dueDate || "no due date"})` : "No recent activity logged."}`,
    signals.bestInsight
      ? `Customer context: ${signals.bestInsight.label}: ${signals.bestInsight.value}.`
      : `Contact: email ${person?.primaryEmail || "missing"}, phone ${person?.primaryPhone || "missing"}.`
  ];

  const riskFlags = [];
  if (!person?.primaryEmail) riskFlags.push("Missing contact email.");
  if (!person?.primaryPhone) riskFlags.push("Missing phone number.");
  if (!latestActivity) riskFlags.push("No recent activities found.");
  if (!latestNote) riskFlags.push("No recent notes found.");
  if (hasDeal && deal.status !== "open") riskFlags.push(`Deal is currently marked '${deal.status}'.`);
  if (signals.risks.length) riskFlags.push(...signals.risks.slice(0, 2));
  if (riskFlags.length === 0) riskFlags.push("No critical risks detected.");

  const cards = buildTalkingCards({ context, latestActivity, latestNote, signals });

  const noAnswer = {
    emailDrafts: buildNoAnswerDrafts({
      context,
      personName,
      orgName,
      hasDeal,
      templatesByStage: parseTemplatesByStage(options?.emailTemplatesByStage)
    })
  };

  return {
    preCall: {
      oneLiner,
      keyFacts: keyFacts.slice(0, 3),
      riskFlags
    },
    cards,
    noAnswer
  };
}

function buildTalkingCards({ context, latestActivity, latestNote, signals }) {
  const resolvedSignals = signals || buildContextSignals(context);
  const personName = context.person?.name || "Contact";
  const personFirst = context.person?.firstName || firstToken(personName) || personName;
  const orgName = context.org?.name || context.person?.org?.name || "Organization";
  const dealTitle = context.deal?.title || context.openDeals?.[0]?.title || "current opportunity";
  const stageText = context.deal?.stageId ? `stage ${context.deal.stageId}` : "current stage";
  const openDealsCount = Array.isArray(context.openDeals) ? context.openDeals.length : 0;
  const insightLine = resolvedSignals.bestInsight
    ? `${resolvedSignals.bestInsight.label}: ${resolvedSignals.bestInsight.value}.`
    : `Use latest note signal: ${latestNote?.content?.slice(0, 90) || "capture one blocker and one desired outcome."}`;

  const stageCards = {
    early: [
      {
        title: "Discovery Alignment",
        bullets: [
          `Ask ${personFirst} what changed now that makes ${dealTitle} a priority.`,
          insightLine
        ]
      },
      {
        title: "Qualification Fit",
        bullets: [
          `Validate success criteria and urgency for ${orgName}.`,
          signals.timeline ? `Timeline signal to confirm: ${signals.timeline}.` : "Confirm target timeline and decision process."
        ]
      }
    ],
    mid: [
      {
        title: "Solution Fit and Gaps",
        bullets: [
          `Pressure-test fit against top use case for ${dealTitle}.`,
          signals.valueDriver || "Tie next conversation to measurable business value."
        ]
      },
      {
        title: "Stakeholder Mapping",
        bullets: [
          signals.stakeholder
            ? `Stakeholder signal detected: ${signals.stakeholder}.`
            : `Identify champion, economic buyer, and implementation owner at ${orgName}.`,
          `Confirm who signs off before moving beyond ${stageText}.`
        ]
      }
    ],
    late: [
      {
        title: "Decision Readiness",
        bullets: [
          `Confirm approval path and exact go/no-go criteria for ${dealTitle}.`,
          signals.risks[0] || "Surface legal, procurement, or security blockers explicitly."
        ]
      },
      {
        title: "Commercial Close Plan",
        bullets: [
          signals.budget
            ? `Budget signal to anchor: ${signals.budget}.`
            : "Confirm budget range, pricing fit, and procurement sequence.",
          "End call with owner + date for final commitment."
        ]
      }
    ],
    unknown: []
  };

  const selectedStageCards = stageCards[resolvedSignals.stageBucket] || stageCards.unknown;

  return [
    ...selectedStageCards,
    {
      title: "Execution Risks",
      bullets: [
        latestActivity
          ? `Reference latest activity: ${latestActivity.subject}.`
          : "Ask why activity momentum has paused.",
        resolvedSignals.competitor
          ? `Competitive pressure noted: ${resolvedSignals.competitor}.`
          : latestNote
            ? `Address recent note insight: ${latestNote.content.slice(0, 80)}.`
            : "Capture one concrete blocker before ending call."
      ]
    },
    {
      title: "Close With Commitment",
      bullets: [
        openDealsCount > 1
          ? `Prioritize this deal among ${openDealsCount} open opportunities and set next checkpoint.`
          : "Ask for a date-bound follow-up action.",
        `Confirm decision criteria and owner for the next milestone at ${orgName}.`
      ]
    }
  ].slice(0, 6);
}

function buildNoAnswerDrafts({ context, personName, orgName, hasDeal, templatesByStage }) {
  const tokens = buildTemplateTokens(context);
  const stageId = context.deal?.stageId ? `stage:${context.deal.stageId}` : "default";
  const sourceDrafts = templatesByStage[stageId] || templatesByStage.default || [];

  const mapped = sourceDrafts
    .slice(0, 3)
    .map((draft, index) => ({
      label: draft.label || `Email #${index + 1}`,
      subject: applyTemplateTokens(draft.subject || "", tokens),
      body: applyTemplateTokens(draft.body || "", tokens)
    }))
    .filter((draft) => draft.subject && draft.body);

  if (mapped.length === 3) {
    return mapped;
  }

  const dealTitle = context.deal?.title || context.openDeals?.[0]?.title || "your current priorities";
  const mention = hasDeal ? `about ${dealTitle}` : "to align on next steps";

  return [
    {
      label: "Email #1",
      subject: `Quick follow-up ${mention}`,
      body: `Hi ${personName},\n\nTried to reach you just now ${mention}. I wanted to keep momentum and confirm what would be most useful for you this week.\n\nIf easier, reply with a time that works and I will adjust.\n\nBest,`
    },
    {
      label: "Email #2",
      subject: `Idea for ${orgName}`,
      body: `Hi ${personName},\n\nSorry we missed each other. Based on our progress, I drafted a concise plan that can help ${orgName} move faster on ${dealTitle}.\n\nHappy to walk through it in a 10-minute call.\n\nBest,`
    },
    {
      label: "Email #3",
      subject: "Reschedule this week?",
      body: `Hi ${personName},\n\nNo worries that we missed each other. Would you like to reconnect for 10 minutes so we can confirm next steps on ${dealTitle}?\n\nShare a couple of windows and I will send an invite.\n\nBest,`
    }
  ];
}

async function fetchPipedrive({ path, baseOrigin, apiToken, method = "GET", body, dataIsArray = false }) {
  const url = new URL(path, baseOrigin);
  url.searchParams.set("api_token", apiToken);

  const response = await fetch(url.toString(), {
    method,
    headers: {
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : undefined
  });

  const data = await safeJson(response);

  if (!response.ok) {
    const apiError = data?.error || data?.error_info || `HTTP ${response.status}`;

    if (response.status === 401) {
      throw new Error("Pipedrive authentication failed. Verify your API token.");
    }

    if (response.status === 404) {
      throw new Error("Pipedrive resource not found for this page context.");
    }

    throw new Error(`Pipedrive API error (${response.status}): ${apiError}`);
  }

  if (data?.success === false) {
    throw new Error(data?.error || "Pipedrive request failed.");
  }

  if (dataIsArray) {
    return Array.isArray(data?.data) ? data.data : [];
  }

  return data?.data || null;
}

async function safeFetchPipedriveOr({ config, fallback }) {
  try {
    return await fetchPipedrive(config);
  } catch (_error) {
    return fallback;
  }
}

function getGmailToken() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        reject(new Error(`OAuth failed: ${chrome.runtime.lastError.message}`));
        return;
      }

      if (!token) {
        reject(new Error("OAuth did not return an access token."));
        return;
      }

      resolve(token);
    });
  });
}

function getOptions() {
  return new Promise((resolve) => {
    chrome.storage.sync.get(DEFAULT_OPTIONS, (result) => resolve(result));
  });
}

function getBaseOrigin(rawUrl) {
  const fallback = "https://app.pipedrive.com";

  try {
    const parsed = new URL(rawUrl || fallback);

    if (!/\.pipedrive\.com$/i.test(parsed.hostname)) {
      throw new Error("Not on pipedrive host.");
    }

    return `${parsed.protocol}//${parsed.hostname}`;
  } catch (_error) {
    return fallback;
  }
}

function safeJson(response) {
  return response
    .json()
    .catch(() => ({}));
}

function normalizeOpenUrl(rawUrl) {
  const value = String(rawUrl || "").trim();
  if (!value) return "";

  const publicLinkedInUrl = extractPublicLinkedInProfileUrl(value);
  if (publicLinkedInUrl) return publicLinkedInUrl;

  const withProtocol = /^https?:\/\//i.test(value) ? value : `https://${value.replace(/^\/+/, "")}`;

  try {
    const parsed = new URL(withProtocol);
    if (!["http:", "https:"].includes(parsed.protocol)) return "";
    const host = String(parsed.hostname || "").toLowerCase();
    if (host.includes("linkedin.com")) {
      const path = String(parsed.pathname || "");
      if (/^\/sales\//i.test(path)) return "";
      if (!/^\/in\/[^/]+/i.test(path)) return "";
      parsed.search = "";
      parsed.hash = "";
      parsed.pathname = path.match(/^\/in\/[^/]+/i)?.[0] || path;
      return parsed.toString().replace(/\/+$/, "");
    }
    return parsed.toString();
  } catch (_error) {
    return "";
  }
}

function normalizeEntity(entity, fallbackName) {
  if (!entity) return null;

  if (typeof entity === "object") {
    return {
      id: Number(entity.value ?? entity.id) || null,
      name: entity.name || fallbackName || null
    };
  }

  return {
    id: Number(entity) || null,
    name: fallbackName || null
  };
}

function normalizeEntityId(entity) {
  if (!entity) return null;

  if (typeof entity === "object") {
    return Number(entity.id || entity.value) || null;
  }

  return Number(entity) || null;
}

function normalizeValueList(values) {
  if (!Array.isArray(values)) return [];

  return values
    .map((entry) => {
      if (typeof entry === "string") return entry;
      if (entry && typeof entry === "object") return entry.value || "";
      return "";
    })
    .map((value) => String(value || "").trim())
    .filter(Boolean);
}

function extractPrimaryValue(values) {
  if (!Array.isArray(values) || values.length === 0) return "";

  const primary = values.find((item) => item?.primary);
  const source = primary || values[0];

  if (typeof source === "string") return source;
  return String(source?.value || "").trim();
}

function formatMoney(value, currency) {
  const amount = Number(value);

  if (!Number.isFinite(amount)) {
    return "Unknown";
  }

  if (!currency) {
    return `${amount}`;
  }

  return `${amount.toLocaleString()} ${currency}`;
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function toUserError(error) {
  if (!error) return "Unknown error.";
  return String(error.message || error);
}

function buildCustomInsights({ dealRaw, personRaw, dealFieldDefs, personFieldDefs }) {
  const dealInsights = extractInsightsFromFields({
    rawEntity: dealRaw,
    fieldDefs: dealFieldDefs,
    scope: "Deal"
  });
  const personInsights = extractInsightsFromFields({
    rawEntity: personRaw,
    fieldDefs: personFieldDefs,
    scope: "Person"
  });

  return [...dealInsights, ...personInsights].slice(0, 10);
}

function extractInsightsFromFields({ rawEntity, fieldDefs, scope }) {
  if (!rawEntity || !Array.isArray(fieldDefs) || fieldDefs.length === 0) return [];

  const categoryMatchers = {
    budget: /(budget|price|pricing|cost|acv|arr|amount|commercial)/i,
    timeline: /(timeline|deadline|target date|go.?live|launch|decision date|start date)/i,
    competitor: /(competitor|incumbent|alternative|vs\\.|versus)/i,
    stakeholder: /(stakeholder|decision maker|economic buyer|champion|procurement|legal|security)/i,
    useCase: /(use case|pain|priority|initiative|goal|objective|problem)/i
  };

  const insights = [];
  for (const field of fieldDefs) {
    const key = field?.key;
    const label = String(field?.name || "").trim();
    if (!key || !label) continue;

    const value = normalizeInsightValue(rawEntity[key]);
    if (!value) continue;

    const category = Object.keys(categoryMatchers).find((name) => categoryMatchers[name].test(label)) || "";
    if (!category) continue;

    insights.push({
      scope,
      label,
      category,
      value: truncate(value, 140)
    });
  }

  return insights;
}

function normalizeInsightValue(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value).trim();
  }
  if (Array.isArray(value)) {
    return value
      .map((item) => normalizeInsightValue(item))
      .filter(Boolean)
      .join(", ")
      .trim();
  }
  if (typeof value === "object") {
    if (value.name) return String(value.name).trim();
    if (value.value) return String(value.value).trim();
    return String(Object.values(value).filter((v) => v !== null && v !== undefined).join(", ")).trim();
  }
  return "";
}

function buildContextSignals(context) {
  const notes = Array.isArray(context.notes) ? context.notes : [];
  const activities = Array.isArray(context.activities) ? context.activities : [];
  const customInsights = Array.isArray(context.customInsights) ? context.customInsights : [];
  const textCorpus = [
    ...notes.map((note) => note.content || ""),
    ...activities.map((activity) => `${activity.subject || ""} ${activity.note || ""}`)
  ]
    .join(" ")
    .toLowerCase();

  const stageBucket = getStageBucket(context.deal?.stageId);
  const budget = firstMatch(customInsights, "budget") || keywordMatch(textCorpus, /(budget|price|pricing|cost|discount)/i);
  const timeline = firstMatch(customInsights, "timeline") || keywordMatch(textCorpus, /(deadline|timeline|go live|launch|this quarter|next month)/i);
  const competitor = firstMatch(customInsights, "competitor") || keywordMatch(textCorpus, /(competitor|incumbent|alternative|vs )/i);
  const stakeholder = firstMatch(customInsights, "stakeholder") || keywordMatch(textCorpus, /(decision maker|procurement|legal|security|cfo|cto|ceo|vp)/i);
  const useCase = firstMatch(customInsights, "useCase");

  const risks = [];
  if (/procurement|legal|security/.test(textCorpus)) risks.push("Procurement or compliance review appears in recent activity.");
  if (/no response|unresponsive|stalled|delayed|postpone/.test(textCorpus)) risks.push("Momentum risk: notes suggest response delays.");
  if (/budget cut|budget freeze|too expensive|pricing concern/.test(textCorpus)) risks.push("Budget risk surfaced in notes.");

  const valueDriver = useCase
    ? `Tie value to the captured use case: ${useCase}.`
    : "Tie value to a concrete operational or revenue outcome from this quarter.";

  const bestInsight = customInsights[0] || null;

  return {
    stageBucket,
    budget,
    timeline,
    competitor,
    stakeholder,
    valueDriver,
    bestInsight,
    risks
  };
}

function getStageBucket(stageId) {
  const id = Number(stageId);
  if (!Number.isFinite(id) || id <= 0) return "unknown";
  if (id <= 2) return "early";
  if (id <= 4) return "mid";
  return "late";
}

function firstMatch(insights, category) {
  const found = insights.find((item) => item.category === category);
  if (!found) return "";
  return `${found.label} = ${found.value}`;
}

function keywordMatch(text, regex) {
  const match = String(text || "").match(regex);
  return match ? match[0] : "";
}

function truncate(value, maxLength) {
  const text = String(value || "");
  if (text.length <= maxLength) return text;
  return `${text.slice(0, maxLength - 3)}...`;
}

function parseTemplatesByStage(rawTemplates) {
  if (!rawTemplates || !String(rawTemplates).trim()) {
    return DEFAULT_STAGE_TEMPLATES;
  }

  try {
    const parsed = JSON.parse(rawTemplates);
    if (!parsed || typeof parsed !== "object") {
      return DEFAULT_STAGE_TEMPLATES;
    }

    return parsed;
  } catch (_error) {
    return DEFAULT_STAGE_TEMPLATES;
  }
}

function buildTemplateTokens(context) {
  const personName = context.person?.name || "there";
  const personFirstName = context.person?.firstName || firstToken(personName) || "there";
  const orgName = context.org?.name || context.person?.org?.name || "your organization";
  const deal = context.deal || context.openDeals?.[0] || null;
  const dealTitle = deal?.title || "your current priorities";
  const dealValue = deal ? formatMoney(deal.value, deal.currency) : "N/A";
  const dealStage = deal?.stageId ? String(deal.stageId) : "";
  const jobTitle = context.person?.title || "";

  return {
    personName,
    personFirstName,
    orgName,
    dealTitle,
    dealValue,
    dealStage,
    jobTitle
  };
}

function applyTemplateTokens(template, tokens) {
  return String(template || "").replace(/\{\{\s*([a-zA-Z0-9_]+)\s*\}\}/g, (_match, tokenName) => {
    return String(tokens[tokenName] ?? "");
  });
}

function firstToken(name) {
  const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
  return parts[0] || "";
}

function lastToken(name) {
  const parts = String(name || "").trim().split(/\s+/).filter(Boolean);
  return parts.length > 1 ? parts[parts.length - 1] : "";
}

function compactAddress(org) {
  if (!org) return "";
  const parts = [org.address, org.address_locality, org.address_country].filter(Boolean);
  return parts.join(", ");
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function findLinkedInUrl(person, options = {}) {
  if (!person) return "";

  const customFieldKey = String(options?.personLinkedinProfileUrlKey || "").trim();
  const fromCustomField = customFieldKey ? String(person?.[customFieldKey] || "").trim() : "";

  if (fromCustomField) {
    return normalizeLinkedInUrl(fromCustomField);
  }

  if (Array.isArray(person.im)) {
    const profile = person.im.find((entry) =>
      String(entry?.value || "").toLowerCase().includes("linkedin.com")
    );
    const fromIm = String(profile?.value || "").trim();
    if (fromIm) {
      return normalizeLinkedInUrl(fromIm);
    }
  }

  const rawWebsite = String(person.website || person.linkedin || "").trim();
  const normalizedWebsite = normalizeLinkedInUrl(rawWebsite);
  if (normalizedWebsite) return normalizedWebsite;

  // Final fallback: scan all person fields, including custom fields.
  const discovered = findLinkedInUrlFromEntity(person);
  if (discovered) return discovered;

  return "";
}

function findLinkedInUrlFromEntity(entity, depth = 0) {
  if (!entity || depth > 3) return "";

  if (typeof entity === "string") {
    return normalizeLinkedInUrl(entity);
  }

  if (Array.isArray(entity)) {
    for (const item of entity) {
      const found = findLinkedInUrlFromEntity(item, depth + 1);
      if (found) return found;
    }
    return "";
  }

  if (typeof entity === "object") {
    for (const [key, value] of Object.entries(entity)) {
      const keyName = String(key || "").toLowerCase();
      if (keyName.includes("linkedin")) {
        const found = findLinkedInUrlFromEntity(value, depth + 1);
        if (found) return found;
      }
      if (typeof value === "string") {
        const normalized = normalizeLinkedInUrl(value);
        if (normalized) return normalized;
      }
    }

    for (const value of Object.values(entity)) {
      const found = findLinkedInUrlFromEntity(value, depth + 1);
      if (found) return found;
    }
  }

  return "";
}

function normalizeLinkedInUrl(value) {
  const extracted = extractPublicLinkedInProfileUrl(value);
  if (extracted) return extracted;

  const raw = normalizeLinkedInCandidate(value);
  if (!raw) return "";

  try {
    const parsed = new URL(raw);
    const host = String(parsed.hostname || "").toLowerCase();
    if (!host.includes("linkedin.com")) {
      return "";
    }
    const path = String(parsed.pathname || "");
    if (!/^\/in\/[^/]+/i.test(path)) {
      return "";
    }
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

function normalizeLinkedInCandidate(value) {
  let raw = String(value || "").trim();
  if (!raw) return "";

  raw = raw.replace(/\\\//g, "/").replace(/^"+|"+$/g, "");
  if (!raw) return "";

  if (/^\/in\//i.test(raw)) {
    return `https://www.linkedin.com${raw}`;
  }
  if (/^in\//i.test(raw)) {
    return `https://www.linkedin.com/${raw}`;
  }
  if (/^[a-z0-9.-]*linkedin\.com\//i.test(raw) && !/^https?:\/\//i.test(raw)) {
    return `https://${raw}`;
  }
  if (/^https?:\/\//i.test(raw)) {
    return raw;
  }

  return "";
}

async function handleOpenLinkedInSidePanel(sender) {
  const tabId = sender?.tab?.id;
  if (!tabId) {
    throw new Error("No tab context available to open side panel.");
  }

  if (!chrome.sidePanel?.open) {
    throw new Error("Side Panel API is not available in this Chrome version. Use the in-page LinkedIn widget fallback.");
  }

  await chrome.sidePanel.open({ tabId });
  return { opened: true, tabId };
}

async function handleLinkedInGetContext() {
  const tab = await getActiveTab();
  if (!tab?.id) {
    throw new Error("No active tab found.");
  }

  if (!isLinkedInUrl(tab.url)) {
    throw new Error("Open a LinkedIn tab first to use LinkedIn Mode.");
  }

  const response = await sendTabMessage(tab.id, { type: "LINKEDIN_DETECT_CONTEXT" });
  if (!response?.ok) {
    throw new Error(response?.error || "Failed to detect LinkedIn context.");
  }

  return {
    tabId: tab.id,
    ...response.data
  };
}

async function handleLinkedInMatchPerson(payload) {
  const options = await getOptions();
  assertLinkedInOptions(options);

  const tab = await getActiveTab();
  if (!tab?.id || !isLinkedInUrl(tab.url)) {
    throw new Error("Switch to a LinkedIn tab and retry.");
  }

  const baseOrigin = getBaseOriginFromActiveTabOrFallback();
  const profileUrl = canonicalizeLinkedInUrl(String(payload.profileUrl || ""));
  const emailHint = String(payload.emailHint || "").trim();
  const profileName = String(payload.profileName || "").trim();

  let person = null;
  let strategy = "none";

  if (profileUrl && options.personLinkedinProfileUrlKey) {
    person = await searchPersonByLinkedInUrl({
      baseOrigin,
      apiToken: options.apiToken,
      linkedInUrl: profileUrl,
      linkedInFieldKey: options.personLinkedinProfileUrlKey
    });
    if (person) strategy = "linkedin_profile_url";
  }

  if (!person && emailHint) {
    person = await searchPersonByEmail({
      baseOrigin,
      apiToken: options.apiToken,
      email: emailHint
    });
    if (person) strategy = "email";
  }

  let candidates = [];
  if (!person && profileName) {
    const results = await searchPersonByName({
      baseOrigin,
      apiToken: options.apiToken,
      name: profileName
    });
    candidates = results.map(toMatchedPersonSummary);
    strategy = candidates.length ? "name_candidates" : "none";
  }

  let dmEligible = false;
  let currentStage = 1;
  let sequenceId = "";
  let enrichedPerson = null;

  if (person) {
    enrichedPerson = toMatchedPersonSummary(person);
    dmEligible = readBooleanField(person[options.personLinkedinDmEligibleKey]);
    currentStage = Number(person[options.personLinkedinDmStageKey]) || 1;
    sequenceId = String(person[options.personLinkedinDmSequenceIdKey] || "");

    const backendEligible = await fetchDmEligibilityFromBackend(options.backendBaseUrl, enrichedPerson.id);
    if (backendEligible?.eligible === true) {
      dmEligible = true;
    }
  }

  return {
    strategy,
    person: enrichedPerson,
    candidates,
    dmEligible,
    currentStage,
    sequenceId
  };
}

async function handleLinkedInConfirmMatch(payload) {
  const personId = Number(payload.personId);
  if (!Number.isFinite(personId) || personId <= 0) {
    throw new Error("Invalid personId for match confirmation.");
  }

  const options = await getOptions();
  assertLinkedInOptions(options);

  const profileUrl = canonicalizeLinkedInUrl(String(payload.profileUrl || ""));
  if (!profileUrl) {
    throw new Error("LinkedIn profile URL is required to confirm match.");
  }

  const baseOrigin = getBaseOriginFromActiveTabOrFallback();
  await updatePersonFields({
    baseOrigin,
    apiToken: options.apiToken,
    personId,
    fields: {
      [options.personLinkedinProfileUrlKey]: profileUrl
    }
  });

  const person = await fetchPersonById({
    baseOrigin,
    apiToken: options.apiToken,
    personId
  });

  return {
    strategy: "manual_confirm",
    person: toMatchedPersonSummary(person),
    candidates: [],
    dmEligible: readBooleanField(person?.[options.personLinkedinDmEligibleKey]),
    currentStage: Number(person?.[options.personLinkedinDmStageKey]) || 1,
    sequenceId: String(person?.[options.personLinkedinDmSequenceIdKey] || "")
  };
}

async function handleLinkedInSearchPersons(payload) {
  const options = await getOptions();
  assertLinkedInOptions(options);

  const query = String(payload.query || "").trim();
  if (!query) {
    throw new Error("Search query is required.");
  }

  const baseOrigin = getBaseOriginFromActiveTabOrFallback();
  const results = await searchPersonByName({
    baseOrigin,
    apiToken: options.apiToken,
    name: query
  });

  return {
    candidates: results.map(toMatchedPersonSummary)
  };
}

async function handleLinkedInGetSequences() {
  const options = await getOptions();
  const backendBaseUrl = normalizeBackendBaseUrl(options.backendBaseUrl);
  try {
    const data = await fetchBackendJson(`${backendBaseUrl}/sequences`);
    if (Array.isArray(data?.sequences) && data.sequences.length) {
      return data;
    }
  } catch (_error) {
    // Fall through to built-in sequence fallback.
  }

  return {
    ok: true,
    version: 1,
    updated_at: null,
    sequences: [
      {
        id: FALLBACK_LINKEDIN_SEQUENCE.id,
        name: FALLBACK_LINKEDIN_SEQUENCE.name,
        description: FALLBACK_LINKEDIN_SEQUENCE.description,
        recommended_start_stage: FALLBACK_LINKEDIN_SEQUENCE.recommended_start_stage
      }
    ]
  };
}

async function handleLinkedInGetTemplates(payload) {
  const options = await getOptions();
  const backendBaseUrl = normalizeBackendBaseUrl(options.backendBaseUrl);
  const sequenceId = String(payload.sequenceId || "").trim();
  if (!sequenceId) {
    throw new Error("sequenceId is required.");
  }

  const stage = Number(payload.stage || 0);
  const params = new URLSearchParams({ sequence_id: sequenceId });
  if (Number.isFinite(stage) && stage > 0) {
    params.set("stage", String(stage));
  }

  try {
    const data = await fetchBackendJson(`${backendBaseUrl}/templates?${params.toString()}`);
    if (Array.isArray(data?.templates) && data.templates.length) {
      return data;
    }
  } catch (_error) {
    // Fall through to built-in template fallback.
  }

  const fallbackTemplates = FALLBACK_LINKEDIN_SEQUENCE.templates
    .filter((template) => !Number.isFinite(stage) || stage <= 0 || Number(template.stage) === stage)
    .map((template) => ({ ...template }));

  return {
    ok: true,
    sequence_id: sequenceId,
    stage: Number.isFinite(stage) && stage > 0 ? stage : null,
    templates: fallbackTemplates
  };
}

async function handleLinkedInGetTalkingPoints(payload) {
  const personId = Number(payload.personId);
  if (!Number.isFinite(personId) || personId <= 0) {
    throw new Error("Valid personId is required for talking points.");
  }

  const options = await getOptions();
  assertLinkedInOptions(options);
  const baseOrigin = getBaseOriginFromActiveTabOrFallback();

  const context = await buildPersonContext(personId, baseOrigin, options.apiToken, options);
  const brief = buildDeterministicBrief(context, options);

  return {
    preCall: brief.preCall,
    cards: brief.cards || []
  };
}

async function handleLinkedInLogAndAdvance(payload) {
  const options = await getOptions();
  assertLinkedInOptions(options);

  const personId = Number(payload.personId);
  const currentStage = Number(payload.currentStage || 1);
  const nextStage = currentStage + 1;
  const profileUrl = canonicalizeLinkedInUrl(String(payload.profileUrl || ""));
  const sequenceId = String(payload.sequenceId || "");
  const templateId = String(payload.templateId || "manual");
  const dmText = String(payload.dmText || "").trim();

  if (!Number.isFinite(personId) || personId <= 0) {
    throw new Error("Valid personId is required.");
  }

  if (!dmText) {
    throw new Error("No DM text to log.");
  }

  const baseOrigin = getBaseOriginFromActiveTabOrFallback();
  const now = new Date().toISOString();
  const today = now.slice(0, 10);
  const subject = `LinkedIn DM sent (Stage ${currentStage})`;

  const noteResult = await createLinkedInDmLog({
    baseOrigin,
    apiToken: options.apiToken,
    personId,
    linkedInUrl: profileUrl,
    stage: currentStage,
    sequenceId,
    templateId,
    timestamp: now,
    messageBody: dmText
  });

  let activityResult = null;
  let activityWarning = "";
  try {
    activityResult = await createLinkedInDmActivity({
      baseOrigin,
      apiToken: options.apiToken,
      personId,
      subject,
      type: "task",
      dueDate: today,
      note: dmText
    });
  } catch (error) {
    activityWarning = toUserError(error);
  }

  await updatePersonFields({
    baseOrigin,
    apiToken: options.apiToken,
    personId,
    fields: {
      [options.personLinkedinDmStageKey]: nextStage,
      [options.personLinkedinDmSequenceIdKey]: sequenceId || null,
      [options.personLinkedinDmLastSentAtKey]: now,
      [options.personLinkedinDmEligibleKey]: false,
      ...(profileUrl ? { [options.personLinkedinProfileUrlKey]: profileUrl } : {})
    }
  });

  return {
    ok: true,
    nextStage,
    dmEligible: false,
    noteId: noteResult?.id || null,
    activityId: activityResult?.id || null,
    activityWarning
  };
}

async function fetchDmEligibilityFromBackend(backendBaseUrl, personId) {
  if (!personId) return null;
  const base = normalizeBackendBaseUrl(backendBaseUrl);

  try {
    return await fetchBackendJson(`${base}/eligible/${personId}`);
  } catch (_error) {
    return null;
  }
}

async function forwardToActiveLinkedInTab(message) {
  const tab = await getActiveTab();
  if (!tab?.id || !isLinkedInUrl(tab.url)) {
    throw new Error("Open a LinkedIn tab to run this action.");
  }

  const response = await sendTabMessage(tab.id, message);
  if (!response?.ok) {
    throw new Error(response?.error || "LinkedIn tab action failed.");
  }

  return response.data;
}

async function getActiveTab() {
  const tabs = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  return tabs[0] || null;
}

function sendTabMessage(tabId, message) {
  return new Promise((resolve) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      if (chrome.runtime.lastError) {
        resolve({ ok: false, error: chrome.runtime.lastError.message });
        return;
      }

      resolve(response || { ok: false, error: "No response from tab." });
    });
  });
}

function isLinkedInUrl(rawUrl) {
  try {
    const parsed = new URL(String(rawUrl || ""));
    const host = String(parsed.hostname || "").toLowerCase();
    return host === "linkedin.com" || host.endsWith(".linkedin.com");
  } catch (_error) {
    return false;
  }
}

function readBooleanField(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    return normalized === "true" || normalized === "1" || normalized === "yes";
  }

  return false;
}

function normalizeBackendBaseUrl(raw) {
  const fallback = APP_CONFIG.backendBaseUrl;
  const value = String(raw || fallback).trim();
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

async function fetchBackendJson(url) {
  const response = await fetch(url, { method: "GET" });
  const data = await safeJson(response);

  if (!response.ok || data?.ok === false) {
    throw new Error(data?.error || `Backend request failed (${response.status})`);
  }

  return data;
}

function toMatchedPersonSummary(person) {
  const linkedInFromIm = Array.isArray(person?.im)
    ? person.im.find((entry) =>
        String(entry?.value || "").toLowerCase().includes("linkedin.com")
      )?.value || ""
    : "";

  return {
    id: Number(person?.id) || null,
    name: person?.name || "Unknown",
    orgName: person?.org_id?.name || person?.org_name || "",
    email: extractPrimaryValue(person?.email),
    linkedInUrl: person?.linkedin_profile_url || linkedInFromIm,
    dealTitle: person?.open_deals_count ? `${person.open_deals_count} open deal(s)` : ""
  };
}

function assertLinkedInOptions(options) {
  if (!options?.apiToken) {
    throw new Error("Pipedrive API token missing in extension options.");
  }
  if (!options?.personLinkedinProfileUrlKey) {
    throw new Error("Set personLinkedinProfileUrlKey in extension options.");
  }
  if (!options?.personLinkedinDmStageKey || !options?.personLinkedinDmSequenceIdKey || !options?.personLinkedinDmLastSentAtKey || !options?.personLinkedinDmEligibleKey) {
    throw new Error("Set LinkedIn DM custom field keys in extension options.");
  }
}

function getBaseOriginFromActiveTabOrFallback() {
  return lastKnownPipedriveOrigin;
}
