const DEFAULTS = {
  apiToken: "",
  autoOpenPanel: true,
  showNotes: true,
  showActivities: true,
  emailTemplatesByStage: "",
  backendBaseUrl: "https://backend-production-9816.up.railway.app",
  configSyncSecret: "cd58bda4728861c653f8c0749315b438162f9d26233b72ef8edec8fb20105ac4",
  personLinkedinProfileUrlKey: "person.linkedin_profile_url",
  personLinkedinDmSequenceIdKey: "person.linkedin_dm_sequence_id",
  personLinkedinDmStageKey: "person.linkedin_dm_stage",
  personLinkedinDmLastSentAtKey: "person.linkedin_dm_last_sent_at",
  personLinkedinDmEligibleKey: "person.linkedin_dm_eligible",
  callDispositionFieldKey: "activity.call_disposition",
  callDispositionTriggerOptionId: "6",
  callDispositionTriggerOptionLabel: "LinkedIn Outreach next step"
};

const refs = {
  apiToken: document.getElementById("apiToken"),
  backendBaseUrl: document.getElementById("backendBaseUrl"),
  configSyncSecret: document.getElementById("configSyncSecret"),
  personLinkedinProfileUrlKey: document.getElementById("personLinkedinProfileUrlKey"),
  personLinkedinDmSequenceIdKey: document.getElementById("personLinkedinDmSequenceIdKey"),
  personLinkedinDmStageKey: document.getElementById("personLinkedinDmStageKey"),
  personLinkedinDmLastSentAtKey: document.getElementById("personLinkedinDmLastSentAtKey"),
  personLinkedinDmEligibleKey: document.getElementById("personLinkedinDmEligibleKey"),
  callDispositionFieldKey: document.getElementById("callDispositionFieldKey"),
  callDispositionTriggerOptionId: document.getElementById("callDispositionTriggerOptionId"),
  emailTemplatesByStage: document.getElementById("emailTemplatesByStage"),
  autoOpenPanel: document.getElementById("autoOpenPanel"),
  showNotes: document.getElementById("showNotes"),
  showActivities: document.getElementById("showActivities"),
  save: document.getElementById("save"),
  pullRemote: document.getElementById("pullRemote"),
  reset: document.getElementById("reset"),
  status: document.getElementById("status")
};

init();

function init() {
  restore();

  refs.save.addEventListener("click", save);
  refs.pullRemote.addEventListener("click", pullFromBackend);
  refs.reset.addEventListener("click", resetDefaults);
}

function restore() {
  chrome.storage.sync.get(DEFAULTS, (result) => {
    applyValuesToForm(result);

    if (shouldHydrateFromBackend(result)) {
      pullFromBackend({ silentOnEmpty: true });
    }
  });
}

async function save() {
  const payload = getFormValues();
  const validationError = validateValues(payload);
  if (validationError) {
    setStatus(validationError, true);
    return;
  }

  chrome.storage.sync.set(
    payload,
    () => {
      if (chrome.runtime.lastError) {
        setStatus(`Failed to save: ${chrome.runtime.lastError.message}`, true);
        return;
      }

      setStatus("Settings saved locally.");
    }
  );
}

function resetDefaults() {
  chrome.storage.sync.set({ ...DEFAULTS }, () => {
    restore();
    setStatus("Defaults restored.");
  });
}

async function pullFromBackend({ silentOnEmpty = false } = {}) {
  const backendBaseUrl = String(refs.backendBaseUrl.value || DEFAULTS.backendBaseUrl || "").trim();
  if (!backendBaseUrl) {
    setStatus("Enter a backend base URL first.", true);
    return;
  }

  try {
    const remote = await fetchRemoteConfig(backendBaseUrl);
    if (!remote) {
      if (!silentOnEmpty) {
        setStatus("No config found on backend yet.", true);
      }
      return;
    }

    const merged = { ...DEFAULTS, ...remote, backendBaseUrl };
    applyValuesToForm(merged);
    chrome.storage.sync.set(merged, () => {
      if (chrome.runtime.lastError) {
        setStatus(`Pulled from backend, but local save failed: ${chrome.runtime.lastError.message}`, true);
        return;
      }
      setStatus("Pulled config from backend.");
    });
  } catch (error) {
    if (!silentOnEmpty) {
      setStatus(`Failed to pull backend config: ${error.message || String(error)}`, true);
    }
  }
}

function getFormValues() {
  return {
    apiToken: String(refs.apiToken.value || "").trim(),
    backendBaseUrl: String(refs.backendBaseUrl.value || "").trim(),
    configSyncSecret: String(refs.configSyncSecret.value || "").trim(),
    personLinkedinProfileUrlKey: String(refs.personLinkedinProfileUrlKey.value || "").trim(),
    personLinkedinDmSequenceIdKey: String(refs.personLinkedinDmSequenceIdKey.value || "").trim(),
    personLinkedinDmStageKey: String(refs.personLinkedinDmStageKey.value || "").trim(),
    personLinkedinDmLastSentAtKey: String(refs.personLinkedinDmLastSentAtKey.value || "").trim(),
    personLinkedinDmEligibleKey: String(refs.personLinkedinDmEligibleKey.value || "").trim(),
    callDispositionFieldKey: String(refs.callDispositionFieldKey.value || "").trim(),
    callDispositionTriggerOptionId: String(refs.callDispositionTriggerOptionId.value || "").trim(),
    emailTemplatesByStage: String(refs.emailTemplatesByStage.value || "").trim(),
    autoOpenPanel: refs.autoOpenPanel.checked,
    showNotes: refs.showNotes.checked,
    showActivities: refs.showActivities.checked
  };
}

function applyValuesToForm(result) {
  refs.apiToken.value = result.apiToken || "";
  refs.backendBaseUrl.value = result.backendBaseUrl || "";
  refs.configSyncSecret.value = result.configSyncSecret || "";
  refs.personLinkedinProfileUrlKey.value = result.personLinkedinProfileUrlKey || "";
  refs.personLinkedinDmSequenceIdKey.value = result.personLinkedinDmSequenceIdKey || "";
  refs.personLinkedinDmStageKey.value = result.personLinkedinDmStageKey || "";
  refs.personLinkedinDmLastSentAtKey.value = result.personLinkedinDmLastSentAtKey || "";
  refs.personLinkedinDmEligibleKey.value = result.personLinkedinDmEligibleKey || "";
  refs.callDispositionFieldKey.value = result.callDispositionFieldKey || "";
  refs.callDispositionTriggerOptionId.value = result.callDispositionTriggerOptionId || "";
  refs.emailTemplatesByStage.value = result.emailTemplatesByStage || "";
  refs.autoOpenPanel.checked = Boolean(result.autoOpenPanel);
  refs.showNotes.checked = Boolean(result.showNotes);
  refs.showActivities.checked = Boolean(result.showActivities);
}

function validateValues(values) {
  if (values.emailTemplatesByStage) {
    try {
      JSON.parse(values.emailTemplatesByStage);
    } catch (_error) {
      return "Template JSON is invalid. Fix formatting before saving.";
    }
  }
  return "";
}

function shouldHydrateFromBackend(result) {
  const backendBaseUrl = String(result.backendBaseUrl || DEFAULTS.backendBaseUrl || "").trim();
  if (!backendBaseUrl || backendBaseUrl === "http://localhost:8787") return false;

  return (!String(result.apiToken || "").trim()
    || !String(result.configSyncSecret || "").trim()
    || !String(result.personLinkedinProfileUrlKey || "").trim()
    || !String(result.personLinkedinDmStageKey || "").trim());
}

async function fetchRemoteConfig(backendBaseUrl) {
  const configSyncSecret = String(refs.configSyncSecret.value || "").trim();
  if (!configSyncSecret) {
    throw new Error("Config sync secret is required.");
  }

  const response = await fetch(`${normalizeBackendBaseUrl(backendBaseUrl)}/extension-config`, {
    method: "GET",
    headers: { "x-peak-access-secret": configSyncSecret }
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || `HTTP ${response.status}`);
  }
  return data.data || null;
}

function normalizeBackendBaseUrl(value) {
  return String(value || "").trim().replace(/\/+$/, "");
}

function setStatus(message, isError = false) {
  refs.status.textContent = message;
  refs.status.style.color = isError ? "#8b2424" : "#2f4841";
}
