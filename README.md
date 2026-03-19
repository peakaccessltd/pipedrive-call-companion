# Peak Access Call Companion (Chrome Extension MV3)

Peak Access is an internal Chrome extension for sales workflows in **Pipedrive** and **LinkedIn**.

It provides:
- Pipedrive deal/person context panel with pre-call brief and talking points
- No-answer follow-up email drafts + Gmail draft creation
- LinkedIn outreach mode with sequence templates
- Human-in-the-loop logging back to Pipedrive (`Log & Advance`)

## Features

### Pipedrive panel
- Detects context on:
  - `/deal/{id}` / `/deals/{id}`
  - `/person/{id}` / `/persons/{id}`
- Fetches context from Pipedrive API
- Builds deterministic pre-call summary + talking points
- No-answer flow with 3 draft options
- Creates Gmail drafts (does not send)
- Paste-and-save missing email to Pipedrive person

### LinkedIn mode
- Sidebar launcher with slide-in right pane
- Match strategy:
  1. LinkedIn URL custom field
  2. Email fallback
  3. Name search + manual confirm
- Loads outreach sequences/templates from backend
- Insert/copy template into LinkedIn composer
- `Log & Advance` writes note/log + advances stage fields

## Repository structure

- `manifest.json` - extension manifest (MV3)
- `background.js` - service worker + API/message orchestration
- `content.js` - Pipedrive content script/panel
- `linkedin-content.js` - LinkedIn content script/fallback panel
- `sidepanel.html|css|js` - LinkedIn side panel UI
- `options.html|js` - extension settings UI
- `ui.css` - Pipedrive panel styles
- `backend/` - shared sequences + webhook backend
- `scripts/pack-extension.sh` - internal CRX packaging script

## Prerequisites

- Google Chrome (current)
- Pipedrive API token
- (Optional) Gmail OAuth client configured for extension draft creation
- Backend deployed (Railway recommended)

## 1) Load extension (dev/unpacked)

1. Open `chrome://extensions`
2. Enable **Developer mode**
3. Click **Load unpacked**
4. Select repo folder:
   - `/Users/larrytoube/Developer/PeakAccess/pipedrive-extension`

## 2) Configure extension options

Open extension options from `chrome://extensions` -> extension -> **Extension options**.

Set at minimum:
- `Pipedrive API token`
- `Shared backend base URL` (example: `https://your-service.up.railway.app`)
- `Config sync secret`
- Person custom field keys:
  - `personLinkedinProfileUrlKey`
  - `personLinkedinDmSequenceIdKey`
  - `personLinkedinDmStageKey`
  - `personLinkedinDmLastSentAtKey`
  - `personLinkedinDmEligibleKey`

Optional:
- `Call disposition field key` and trigger option id
- No-answer email template JSON overrides

Config sync:
- `Save` stores config locally in the extension
- `Pull from backend` restores the centrally managed config from your backend
- if local settings are empty and both `backendBaseUrl` and `configSyncSecret` are already known, the extension will try to auto-hydrate from backend on next load

Save, then reload extension.

## 3) Backend setup (local)

```bash
cd backend
npm install
cp .env.example .env
npm run dev
```

Health check:
- `http://localhost:8787/health`

## 4) Deploy backend to Railway (basic)

1. Create a new Railway project
2. Connect this GitHub repo
3. Set **Root Directory** to `backend`
4. Railway will auto-detect Node and run `npm start`
5. Deploy and copy the generated public URL (`https://<service>.up.railway.app`)

Set environment variables:
- `WEBHOOK_SECRET` (required)
- `DATABASE_URL` (optional but recommended for durable queue)
- `PIPEDRIVE_BASE_URL` (optional)
- `PIPEDRIVE_API_TOKEN` (optional)
- `PERSON_DM_ELIGIBLE_FIELD_KEY` (optional)
- `CALL_DISPOSITION_TRIGGER_OPTION_ID` (optional, default `6`)
- `CALL_DISPOSITION_TRIGGER_LABEL` (optional)

### Backend endpoints

- `GET /health`
- `GET /sequences`
- `GET /sequences/:id`
- `GET /templates?sequence_id=<id>&stage=<n>`
- `GET /eligible/:personId`
- `GET /extension-config`
- `PUT /extension-config`
- `POST /pipedrive/webhook`
- `GET /admin` (basic-auth protected template admin UI)
- `GET /admin/api/sequences`
- `PUT /admin/api/sequences`

### Online template editor

You can manage shared templates/sequences and extension config in-browser:

1. Set backend env vars:
   - `ADMIN_USERNAME`
   - `ADMIN_PASSWORD`
2. Deploy backend
3. Open:
   - `https://<your-backend>/admin`
4. Sign in with basic auth credentials
5. Edit sequences and extension config, then click **Save**

### Config sync auth

Set:
- `CONFIG_SYNC_SECRET`

The extension uses this secret when calling `GET /extension-config` to restore central config.

## 5) Pipedrive webhook setup

If your Pipedrive webhook UI cannot send custom headers, use query-secret auth:

`https://<your-backend>/pipedrive/webhook?secret=<WEBHOOK_SECRET>`

Recommended trigger:
- Object: `activity`
- Action: `change`
- Condition: call completed/updated according to your process

Minimal webhook body (works):
```json
{
  "person_id": "<activity contact person id>"
}
```

## 6) Gmail draft setup (optional)

OAuth scope required in manifest:
- `https://www.googleapis.com/auth/gmail.compose`

Create Google OAuth client for Chrome extension and set your extension ID.

## 7) Internal packaging (.crx)

Run:

```bash
./scripts/pack-extension.sh
```

Outputs:
- `dist/pipedrive-extension.crx`
- `dist/pipedrive-extension.pem`

Important:
- Keep `.pem` private and backed up
- Reuse same `.pem` for all future updates

## 8) Test checklist (quick)

1. Open Pipedrive deal/person page -> panel appears
2. Click refresh -> context + brief render
3. No-answer -> create Gmail draft
4. Open LinkedIn profile -> LinkedIn pane loads
5. Verify sequence/template appears
6. Insert/copy template
7. Send manually on LinkedIn
8. Click `Log & Advance`
9. Confirm Pipedrive person fields/stage update
10. Validate eligibility endpoint:
    - `/eligible/<personId>` returns `eligible: true`

## Troubleshooting

### Service worker shows "Inactive"
Normal for MV3. Service workers sleep when idle and wake on events/messages.

### LinkedIn template list empty
- Verify backend URL in extension options
- Check backend `/sequences` directly
- Reload extension after changes

### Insert template does nothing
- Open LinkedIn message composer first
- Use `Copy` fallback and paste manually

### Pipedrive webhook 401
- Wrong/missing `WEBHOOK_SECRET`
- If using query auth, confirm `?secret=...` matches env var exactly

### No person match in LinkedIn mode
- Confirm person field key for LinkedIn URL is correct
- Use search + `Confirm Match + Save URL`

## Security notes

- Do not hardcode secrets in source
- Keep API tokens and OAuth credentials in env/options only
- Keep `.pem` private
- LinkedIn sending stays manual (human-in-the-loop)

## License / Usage

Internal tool for Peak Access sales operations.
