# KVS Workshop Sessions Feedback

A responsive session-wise feedback web app for KVS teachers attending professional development programmes at ZIET Mumbai.

## Features
- Participant, Vidyalaya, KVS Region, workshop, session date/title and resource-person details
- Seven 1–5 rating criteria plus an overall session rating
- Qualitative reflections, strengths, improvement suggestions and follow-up needs
- Mobile-friendly form and accessible rating controls
- Google Apps Script + Google Sheets backend (configured separately)

## Repository
https://github.com/taprimaryzietmumbai-kjrd/KVS-Workshop-Sessions-Feedback

## Deploy the frontend with GitHub Pages
1. Open **Settings → Pages**.
2. Under **Build and deployment**, choose **Deploy from a branch**.
3. Select branch **main** and folder **/(root)**, then save.
4. After deployment, open the Pages URL shown by GitHub.

## Connect Google Sheets
1. Create a Google Sheet for feedback responses.
2. Open **Extensions → Apps Script**.
3. Copy the contents of `Code.gs` from this repository into the Apps Script editor.
4. Set the spreadsheet ID in the script configuration.
5. Deploy as a **Web app**, execute as yourself, and choose access appropriate to your institution's data-governance requirements.
6. Copy the deployed web app URL ending in `/exec`.
7. In `index.html`, set `const APPS_SCRIPT_URL = "YOUR_DEPLOYED_WEB_APP_URL";` to that URL and commit the change.

**Privacy:** The frontend is public. Do not publish feedback data, spreadsheet links, credentials, or private participant information in this repository. Restrict access to the response spreadsheet. Choose the web-app access policy deliberately and test with non-sensitive sample responses before collecting real data.

## Files
- `index.html` — feedback form
- `Code.gs` — Google Apps Script endpoint
