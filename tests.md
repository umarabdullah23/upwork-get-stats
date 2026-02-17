# Tests

## Playwright MCP (fixture-driven now)

### Job activity extraction (test_data/job_page.html)
- Extracts job name from heading and ignores the "Activity on this job" header.
- Derives job link + job ID from the saved Upwork URL in the fixture header comment.
- Captures posted-since text (jobCreatedSince) from the posted-on-line block.
- Captures client country from the client-location block.
- Detects payment verification status from page text.
- Captures activity metrics: proposals, invites sent, interviewing, last viewed, unanswered invites.
- Handles missing proposal link by returning empty proposalId (or populated if fixture adds one).
- Returns an error when activity items are removed (negative test by stripping the activity section).

### Proposals extraction (test_data/Proposals.html)
- Returns a list of proposal rows with proposalId + title.
- Marks viewed when proposal-views-slot is visible and includes last-seen/mark content.
- Does not mark viewed when proposal-views-slot is hidden (d-none).
- Ignores rows without proposal links.

### Connects history extraction (test_data/ConnectsHistory.html)
- Parses jobId, title, and link from connects history table rows.
- Calculates connects spent vs refund from negative/positive values (including commas).
- Flags boosted rows and maps values to boostedConnectsSpent/Refund.
- Returns an error when the table has no rows.

### URL mode detection (js/popup/utils.js)
- Returns proposals mode for test_data/Proposals.html and /nx/proposals URLs.
- Returns connects mode for test_data/ConnectsHistory.html and /nx/plans/connects/history URLs.
- Defaults to job mode for job pages and unknown URLs.

## Later E2E (Chrome extension + Sheets)

### Popup UI
- Read page updates status + fields and enables copy/send actions.
- Toggle job details expands/collapses and updates aria-expanded + button label.
- Copy CSV writes to clipboard and reports success/error.
- Bidder input persists to storage and restores on reopen.
- Populate job IDs only shows when current tab is a sheet URL.
- Mode switching updates button labels and hides/shows actions.

### Sheets integration
- Auth flow handles timeout and bad client ID error.
- Headers read + write, template prep, and row formatting behaviors.
- Update existing row by job ID, or insert into empty row.
- Proposal view update marks read rows green.
- Connects history updates existing rows and appends new ones (including boosted columns).
