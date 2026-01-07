# Upwork Activity Reader (Read-Only)

Read-only Chrome extension that extracts job activity data from the current Upwork
job details page. It never clicks buttons, submits forms, or sends data anywhere.

## What It Captures

- Name (job title)
- Link (current page URL without query params)
- Date (capture date)
- Proposals
- Invites sent
- Interviewing
- Last viewed by client
- Unanswered invites

## Usage

1. Load this directory as an unpacked extension in Chrome.
2. Open an Upwork job details page.
3. Click the extension icon, then click "Read current job page".
4. Click "Copy CSV row" to paste into Google Sheets or a tracker.

## Google Sheets (Optional)

Use the Google Sheets API directly from the extension.

1. Create a Google Cloud project and enable the Google Sheets API.
2. Configure the OAuth consent screen.
3. Create an OAuth client ID for **Chrome Extension** and use your extension ID.
4. Add the OAuth client ID to `manifest.json` under `oauth2`.
5. In the popup, paste your spreadsheet ID (or full URL) and sheet tab name.
6. Click "Send to Sheets" to authorize and append a row.

Spreadsheet ID is the string between `/d/` and `/edit` in the sheet URL.

## Safety Notes

- Read-only: only parses visible text in the current tab.
- No automated navigation, scraping, or background collection.
- Network requests only happen when you click "Send to Sheets".
