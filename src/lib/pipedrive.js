function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function safeJson(response) {
  try {
    return await response.json();
  } catch (_error) {
    return {};
  }
}

function normalizeApiError(response, body) {
  const detail = body?.error || body?.error_info || body?.data?.error || "Unknown API error";
  return new Error(`Pipedrive API error (${response.status}): ${detail}`);
}

async function fetchWithBackoff(url, init, maxAttempts = 4) {
  let lastError = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await fetch(url, init);
      if (response.status !== 429 && response.status < 500) {
        return response;
      }

      const retryAfter = Number(response.headers.get("retry-after")) || 0;
      const wait = retryAfter > 0 ? retryAfter * 1000 : attempt * 500;
      await sleep(wait);
    } catch (error) {
      lastError = error;
      await sleep(attempt * 500);
    }
  }

  throw lastError || new Error("Request failed after retries.");
}

function buildApiUrl(baseOrigin, path, apiToken, query = {}) {
  const url = new URL(path, baseOrigin);
  url.searchParams.set("api_token", apiToken);

  Object.entries(query).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") {
      url.searchParams.set(key, String(value));
    }
  });

  return url.toString();
}

async function pipedriveRequest({ baseOrigin, apiToken, path, method = "GET", query = {}, body }) {
  const url = buildApiUrl(baseOrigin, path, apiToken, query);
  const response = await fetchWithBackoff(url, {
    method,
    headers: {
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : undefined
  });

  const data = await safeJson(response);

  if (!response.ok || data?.success === false) {
    throw normalizeApiError(response, data);
  }

  return data?.data;
}

async function fetchPersonById({ baseOrigin, apiToken, personId }) {
  if (!personId) return null;

  return pipedriveRequest({
    baseOrigin,
    apiToken,
    path: `/api/v1/persons/${personId}`
  });
}

function normalizeItems(items) {
  if (!Array.isArray(items)) return [];

  return items
    .map((item) => {
      const person = item?.item;
      const id = Number(person?.id || item?.item?.id || item?.id);
      if (!Number.isFinite(id) || id <= 0) return null;
      return {
        id,
        name: person?.name || item?.title || "Unknown",
        detail: person || item
      };
    })
    .filter(Boolean);
}

async function itemSearch({ baseOrigin, apiToken, term, fields, exact = false, limit = 5 }) {
  const data = await pipedriveRequest({
    baseOrigin,
    apiToken,
    path: "/api/v1/itemSearch",
    query: {
      term,
      item_types: "person",
      fields,
      exact_match: exact ? 1 : 0,
      limit
    }
  });

  return normalizeItems(data?.items || []);
}

async function searchPersonByLinkedInUrl({ baseOrigin, apiToken, linkedInUrl, linkedInFieldKey }) {
  if (!linkedInUrl || !linkedInFieldKey) return null;

  const candidates = await itemSearch({
    baseOrigin,
    apiToken,
    term: linkedInUrl,
    fields: "custom_fields",
    exact: false,
    limit: 10
  });

  for (const candidate of candidates) {
    const person = await fetchPersonById({ baseOrigin, apiToken, personId: candidate.id });
    const fieldValue = String(person?.[linkedInFieldKey] || "").trim();

    if (fieldValue && canonicalizeLinkedInUrl(fieldValue) === canonicalizeLinkedInUrl(linkedInUrl)) {
      return person;
    }
  }

  return null;
}

async function searchPersonByEmail({ baseOrigin, apiToken, email }) {
  if (!email) return null;

  const candidates = await itemSearch({
    baseOrigin,
    apiToken,
    term: email,
    fields: "email",
    exact: true,
    limit: 3
  });

  if (!candidates.length) return null;
  return fetchPersonById({ baseOrigin, apiToken, personId: candidates[0].id });
}

async function searchPersonByName({ baseOrigin, apiToken, name }) {
  if (!name) return [];

  const candidates = await itemSearch({
    baseOrigin,
    apiToken,
    term: name,
    fields: "name",
    exact: false,
    limit: 5
  });

  const details = await Promise.all(
    candidates.map((candidate) => fetchPersonById({ baseOrigin, apiToken, personId: candidate.id }))
  );

  return details.filter(Boolean);
}

async function updatePersonFields({ baseOrigin, apiToken, personId, fields }) {
  if (!personId) {
    throw new Error("personId is required for person update.");
  }

  return pipedriveRequest({
    baseOrigin,
    apiToken,
    path: `/api/v1/persons/${personId}`,
    method: "PUT",
    body: fields
  });
}

async function createLinkedInDmLog({
  baseOrigin,
  apiToken,
  personId,
  leadId,
  linkedInUrl,
  stage,
  sequenceId,
  templateId,
  messageBody,
  timestamp
}) {
  if (!personId) {
    throw new Error("personId is required for LinkedIn DM logging.");
  }

  const noteLines = [
    "LinkedIn Outreach DM Logged",
    `Timestamp: ${timestamp}`,
    `LinkedIn URL: ${linkedInUrl || "N/A"}`,
    `Sequence ID: ${sequenceId || "N/A"}`,
    `Template ID: ${templateId || "N/A"}`,
    `Stage: ${stage || "N/A"}`,
    "",
    "Message:",
    messageBody || "(empty)"
  ];

  const notePayload = {
    person_id: personId,
    content: noteLines.map((line) => escapeHtml(line)).join("<br>")
  };

  if (leadId) {
    notePayload.lead_id = String(leadId);
  }

  const note = await pipedriveRequest({
    baseOrigin,
    apiToken,
    path: "/api/v1/notes",
    method: "POST",
    body: notePayload
  });

  return note;
}

async function createLinkedInDmActivity({
  baseOrigin,
  apiToken,
  personId,
  subject,
  type = "task",
  note,
  dueDate
}) {
  if (!personId) {
    throw new Error("personId is required for LinkedIn DM activity logging.");
  }

  const normalizedDueDate = String(dueDate || new Date().toISOString().slice(0, 10));
  const activityPayload = {
    person_id: personId,
    subject: String(subject || "LinkedIn outreach message sent"),
    type: String(type || "task"),
    done: true,
    due_date: normalizedDueDate,
    note: note || ""
  };

  return pipedriveRequest({
    baseOrigin,
    apiToken,
    path: "/api/v1/activities",
    method: "POST",
    body: activityPayload
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function canonicalizeLinkedInUrl(url) {
  const extracted = extractPublicLinkedInProfileUrl(url);
  if (extracted) return extracted.toLowerCase();

  try {
    const parsed = new URL(String(url || ""));
    const host = String(parsed.hostname || "").toLowerCase();
    const path = String(parsed.pathname || "");
    if (!host.includes("linkedin.com")) return "";
    if (!/^\/in\/[^/]+/i.test(path)) return "";
    parsed.search = "";
    parsed.hash = "";
    parsed.pathname = (path.match(/^\/in\/[^/]+/i)?.[0] || path).replace(/\/+$/, "");
    return `${parsed.origin}${parsed.pathname}`.toLowerCase();
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
    const path = String(parsed.pathname || "");
    if (!/^\/in\/[^/]+/i.test(path)) return "";
    parsed.search = "";
    parsed.hash = "";
    parsed.pathname = (path.match(/^\/in\/[^/]+/i)?.[0] || path).replace(/\/+$/, "");
    return `${parsed.origin}${parsed.pathname}`;
  } catch (_error) {
    return "";
  }
}

export {
  pipedriveRequest,
  fetchPersonById,
  searchPersonByLinkedInUrl,
  searchPersonByEmail,
  searchPersonByName,
  updatePersonFields,
  createLinkedInDmLog,
  createLinkedInDmActivity
};
