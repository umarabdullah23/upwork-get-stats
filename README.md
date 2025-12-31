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

## Safety Notes

- Read-only: only parses visible text in the current tab.
- No automated navigation, scraping, or background collection.
- No network requests are made by the extension.
