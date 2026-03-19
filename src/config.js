export const APP_CONFIG = {
  backendBaseUrl: "https://backend-production-9816.up.railway.app",
  configSyncSecret: "cd58bda4728861c653f8c0749315b438162f9d26233b72ef8edec8fb20105ac4",
  pipedrive: {
    customFields: {
      personLinkedinProfileUrlKey: "person.linkedin_profile_url",
      personLinkedinDmSequenceIdKey: "person.linkedin_dm_sequence_id",
      personLinkedinDmStageKey: "person.linkedin_dm_stage",
      personLinkedinDmLastSentAtKey: "person.linkedin_dm_last_sent_at",
      personLinkedinDmEligibleKey: "person.linkedin_dm_eligible"
    },
    activity: {
      callDispositionFieldKey: "activity.call_disposition",
      // Prefer ID over label because labels can be renamed by admins.
      callDispositionTriggerOptionId: "6",
      callDispositionTriggerOptionLabel: "LinkedIn Outreach next step"
    }
  }
};

export const STORAGE_KEYS = {
  apiToken: "apiToken",
  autoOpenPanel: "autoOpenPanel",
  showNotes: "showNotes",
  showActivities: "showActivities",
  emailTemplatesByStage: "emailTemplatesByStage",
  backendBaseUrl: "backendBaseUrl",
  configSyncSecret: "configSyncSecret",
  personLinkedinProfileUrlKey: "personLinkedinProfileUrlKey",
  personLinkedinDmSequenceIdKey: "personLinkedinDmSequenceIdKey",
  personLinkedinDmStageKey: "personLinkedinDmStageKey",
  personLinkedinDmLastSentAtKey: "personLinkedinDmLastSentAtKey",
  personLinkedinDmEligibleKey: "personLinkedinDmEligibleKey",
  callDispositionFieldKey: "callDispositionFieldKey",
  callDispositionTriggerOptionId: "callDispositionTriggerOptionId",
  callDispositionTriggerOptionLabel: "callDispositionTriggerOptionLabel"
};

export const DEFAULT_STORAGE = {
  apiToken: "",
  autoOpenPanel: true,
  showNotes: true,
  showActivities: true,
  emailTemplatesByStage: "",
  backendBaseUrl: APP_CONFIG.backendBaseUrl,
  configSyncSecret: APP_CONFIG.configSyncSecret,
  personLinkedinProfileUrlKey: APP_CONFIG.pipedrive.customFields.personLinkedinProfileUrlKey,
  personLinkedinDmSequenceIdKey: APP_CONFIG.pipedrive.customFields.personLinkedinDmSequenceIdKey,
  personLinkedinDmStageKey: APP_CONFIG.pipedrive.customFields.personLinkedinDmStageKey,
  personLinkedinDmLastSentAtKey: APP_CONFIG.pipedrive.customFields.personLinkedinDmLastSentAtKey,
  personLinkedinDmEligibleKey: APP_CONFIG.pipedrive.customFields.personLinkedinDmEligibleKey,
  callDispositionFieldKey: APP_CONFIG.pipedrive.activity.callDispositionFieldKey,
  callDispositionTriggerOptionId: APP_CONFIG.pipedrive.activity.callDispositionTriggerOptionId,
  callDispositionTriggerOptionLabel: APP_CONFIG.pipedrive.activity.callDispositionTriggerOptionLabel
};
