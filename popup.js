const readButton = document.getElementById("read-page");
const copyButton = document.getElementById("copy-csv");
const checkViewedButton = document.getElementById("check-viewed");
const currentJobCard = document.getElementById("current-job-card");
const currentJobStatus = document.getElementById("current-job-status");
const jobDetails = document.getElementById("job-details");
const toggleJobDetailsButton = document.getElementById("toggle-job-details");
const currentJobActions = document.getElementById("current-job-actions");
const spreadsheetInput = document.getElementById("spreadsheet-id");
const sheetNameInput = document.getElementById("sheet-name");
const openSheetButton = document.getElementById("open-sheet");
const prepareSheetButton = document.getElementById("prepare-sheet");
const sendSheetsButton = document.getElementById("send-sheets");
const statusEl = document.getElementById("status");
const bidderInput = document.getElementById("bidder-input");

const fields = {
	name: document.getElementById("field-name"),
	link: document.getElementById("field-link"),
	jobId: document.getElementById("field-job-id"),
	date: document.getElementById("field-date"),
	proposals: document.getElementById("field-proposals"),
	invitesSent: document.getElementById("field-invites"),
	interviewing: document.getElementById("field-interviewing"),
	lastViewed: document.getElementById("field-last-viewed"),
	unansweredInvites: document.getElementById("field-unanswered"),
};

let currentData = null;
let spreadsheetId = "";
let sheetName = "Sheet1";
let bidder = "";
let saveTimer = null;
let uiMode = "job";
let jobDetailsExpanded = false;

const setStatus = (message, tone = "") => {
	statusEl.textContent = message;
	statusEl.className = `status ${tone}`.trim();
};

const setFields = (data) => {
	fields.name.textContent = data?.name || "-";
	fields.link.textContent = data?.link || "-";
	fields.jobId.textContent = data?.jobId || "-";
	fields.date.textContent = data?.date || "-";
	fields.proposals.textContent = data?.proposals || "-";
	fields.invitesSent.textContent = data?.invitesSent || "-";
	fields.interviewing.textContent = data?.interviewing || "-";
	fields.lastViewed.textContent = data?.lastViewed || "-";
	fields.unansweredInvites.textContent = data?.unansweredInvites || "-";

	const fetched = Boolean(data?.name || data?.jobId || data?.link);
	if (currentJobStatus) {
		currentJobStatus.textContent = fetched ? "Fetched" : "Not fetched yet";
		currentJobStatus.classList.toggle("is-fetched", fetched);
		currentJobStatus.classList.toggle("is-empty", !fetched);
		currentJobStatus.classList.remove("is-detected");
	}
	if (toggleJobDetailsButton) {
		toggleJobDetailsButton.disabled = !fetched;
	}
	if (!fetched) {
		jobDetailsExpanded = false;
		if (jobDetails) {
			jobDetails.hidden = true;
		}
		if (toggleJobDetailsButton) {
			toggleJobDetailsButton.textContent = "View more";
			toggleJobDetailsButton.setAttribute("aria-expanded", "false");
		}
	}
};

const getCurrentSpreadsheetId = () =>
	extractSpreadsheetId(spreadsheetInput.value) || spreadsheetId;

const setButtons = () => {
	if (copyButton) {
		copyButton.disabled = !currentData;
	}
	sendSheetsButton.disabled = !currentData || !spreadsheetId;
	openSheetButton.disabled = !getCurrentSpreadsheetId();
	if (checkViewedButton) {
		checkViewedButton.disabled = !spreadsheetId;
	}
};

const setMode = (mode) => {
	uiMode = mode;
	if (!checkViewedButton) {
		return;
	}
	if (mode === "proposals") {
		if (currentJobCard) {
			currentJobCard.hidden = false;
		}
		if (currentJobStatus) {
			currentJobStatus.textContent = "Proposals page detected";
			currentJobStatus.classList.remove("is-fetched");
			currentJobStatus.classList.remove("is-empty");
			currentJobStatus.classList.add("is-detected");
		}
		if (jobDetails) {
			jobDetailsExpanded = false;
			jobDetails.hidden = true;
		}
		if (toggleJobDetailsButton) {
			toggleJobDetailsButton.hidden = true;
		}
		if (readButton) {
			readButton.hidden = true;
		}
		if (currentJobActions) {
			currentJobActions.hidden = true;
		}
		if (sendSheetsButton) {
			sendSheetsButton.hidden = true;
		}
		if (copyButton) {
			copyButton.hidden = true;
		}
		checkViewedButton.hidden = false;
	} else {
		if (currentJobCard) {
			currentJobCard.hidden = false;
		}
		if (currentJobActions) {
			currentJobActions.hidden = false;
		}
		if (toggleJobDetailsButton) {
			toggleJobDetailsButton.hidden = false;
		}
		if (sendSheetsButton) {
			sendSheetsButton.hidden = false;
		}
		readButton.hidden = false;
		if (copyButton) {
			copyButton.hidden = false;
		}
		checkViewedButton.hidden = true;
	}
};

const toggleJobDetails = () => {
	if (!jobDetails || !toggleJobDetailsButton) {
		return;
	}
	if (toggleJobDetailsButton.disabled) {
		return;
	}
	jobDetailsExpanded = !jobDetailsExpanded;
	jobDetails.hidden = !jobDetailsExpanded;
	toggleJobDetailsButton.textContent = jobDetailsExpanded
		? "View less"
		: "View more";
	toggleJobDetailsButton.setAttribute(
		"aria-expanded",
		jobDetailsExpanded ? "true" : "false"
	);
};

if (toggleJobDetailsButton) {
	toggleJobDetailsButton.addEventListener("click", toggleJobDetails);
}

const scheduleSaveSheetsSettings = () => {
	if (saveTimer) {
		clearTimeout(saveTimer);
	}
	saveTimer = setTimeout(async () => {
		const nextId = extractSpreadsheetId(spreadsheetInput.value);
		const nextName = normalizeSheetName(sheetNameInput.value);
		if (!nextId) {
			spreadsheetId = "";
			sheetName = nextName;
			setButtons();
			return;
		}
		const ok = await saveSheetsSettings(nextId, nextName);
		if (ok) {
			spreadsheetId = nextId;
			sheetName = nextName;
			setButtons();
		}
	}, 300);
};

const openSpreadsheet = (id) => {
	const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
		id
	)}/edit`;
	if (chrome?.tabs?.create) {
		chrome.tabs.create({ url });
	} else {
		window.open(url, "_blank", "noopener");
	}
};

readButton.addEventListener("click", async () => {
	setStatus("Reading current tab...", "");
	const tab = await queryActiveTab();

	if (!tab || !tab.id || !tab.url) {
		setStatus("No active tab detected.", "warn");
		return;
	}

	if (!tab.url.includes("upwork.com")) {
		setStatus("Open an Upwork job details page first.", "warn");
		return;
	}

	const result = await executeExtraction(tab.id);
	if (!result.ok) {
		setStatus(result.error || "Unable to read this page.", "error");
		return;
	}

	currentData = result.data;
	currentData.bidder = bidder;
	setFields(currentData);
	setButtons();
	setStatus("Captured job activity.", "success");
});

if (checkViewedButton) {
	checkViewedButton.addEventListener("click", async () => {
		setStatus("Scanning proposals page...", "");
		const tab = await queryActiveTab();

		if (!tab || !tab.id || !tab.url) {
			setStatus("No active tab detected.", "warn");
			return;
		}

		if (detectModeFromUrl(tab.url) !== "proposals") {
			setStatus("Open https://www.upwork.com/nx/proposals/ first.", "warn");
			return;
		}

		if (!spreadsheetId) {
			setStatus("Save your Sheets settings first.", "warn");
			return;
		}

		const scan = await executeProposalScan(tab.id);
		if (!scan.ok) {
			setStatus(scan.error || "Unable to scan proposals on this page.", "error");
			return;
		}

		const proposals = Array.isArray(scan.proposals) ? scan.proposals : [];
		const viewed = proposals.filter((item) => item?.viewed && item?.proposalId);
		if (!viewed.length) {
			setStatus("No 'Viewed by client' proposals found on this page.", "warn");
			return;
		}

		setStatus(`Found ${viewed.length} viewed proposal(s). Updating sheet...`, "");
		const auth = await getAuthToken(true);
		if (!auth.ok) {
			setStatus(auth.error || "Google authorization failed.", "error");
			return;
		}

		const idMap = await getProposalIdRowMap(
			auth.token,
			spreadsheetId,
			sheetName
		);
		if (!idMap) {
			setStatus("Unable to read Proposal ID column from the sheet.", "error");
			return;
		}

		const matchedRows = [];
		for (const item of viewed) {
			const rowIndex = idMap.get(String(item.proposalId).trim());
			if (rowIndex) {
				matchedRows.push(rowIndex);
			}
		}

		const uniqueRows = Array.from(new Set(matchedRows)).sort((a, b) => a - b);
		if (!uniqueRows.length) {
			setStatus(
				`Found ${viewed.length} viewed proposal(s), but none matched the sheet's Proposal ID column.`,
				"warn"
			);
			return;
		}

		const ok = await setReadCellsViewed(
			auth.token,
			spreadsheetId,
			sheetName,
			uniqueRows
		);
		if (!ok) {
			setStatus("Failed to update the sheet formatting.", "error");
			return;
		}
		setStatus(
			`Marked ${uniqueRows.length} row(s) as read (green).`,
			"success"
		);
	});
}

if (copyButton) {
	copyButton.addEventListener("click", async () => {
		if (!currentData) {
			setStatus("Nothing to copy yet.", "warn");
			return;
		}
		currentData.bidder = bidder;
		const csv = toCsvRow(currentData);
		try {
			await navigator.clipboard.writeText(csv);
			setStatus("CSV row copied.", "success");
		} catch (error) {
			setStatus("Clipboard access failed.", "error");
		}
	});
}

openSheetButton.addEventListener("click", () => {
	const nextId = getCurrentSpreadsheetId();
	if (!nextId) {
		setStatus("Enter a spreadsheet ID or URL first.", "warn");
		return;
	}
	openSpreadsheet(nextId);
});

spreadsheetInput.addEventListener("input", () => {
	setButtons();
	scheduleSaveSheetsSettings();
});

sheetNameInput.addEventListener("input", () => {
	scheduleSaveSheetsSettings();
});

prepareSheetButton.addEventListener("click", async () => {
	if (!spreadsheetId) {
		setStatus("Save your Sheets settings first.", "warn");
		return;
	}
	setStatus("Preparing sheet...", "");
	const auth = await getAuthToken(true);
	if (!auth.ok) {
		setStatus(auth.error || "Google authorization failed.", "error");
		return;
	}
	const headerResponse = await setHeaders(auth.token, spreadsheetId, sheetName);
	if (!headerResponse.ok) {
		setStatus(`Sheets API error ${headerResponse.status}.`, "error");
		return;
	}
	await setHeaderBold(auth.token, spreadsheetId, sheetName, 23);
	await clearBodyBold(auth.token, spreadsheetId, sheetName, 23, 1000);
	await setBodyColumnColors(auth.token, spreadsheetId, sheetName, 1000);
	await freezeHeaderRow(auth.token, spreadsheetId, sheetName);
	await ensureBidderDropdown(auth.token, spreadsheetId, sheetName);
	setStatus("Sheet prepared with headers and bidder dropdown.", "success");
});

sendSheetsButton.addEventListener("click", async () => {
	if (!currentData) {
		setStatus("Read a job page first.", "warn");
		return;
	}
	currentData.bidder = bidder;
	if (!spreadsheetId) {
		setStatus("Save your Sheets settings first.", "warn");
		return;
	}

	setStatus("Connecting to Google...", "");
	const auth = await getAuthToken(true);
	if (!auth.ok) {
		setStatus(auth.error || "Google authorization failed.", "error");
		return;
	}
	let activeToken = auth.token;

	let row = null;
	await ensureBidderDropdown(activeToken, spreadsheetId, sheetName);
	let response;
	let existingRow = await findRowByJobId(
		activeToken,
		spreadsheetId,
		sheetName,
		currentData.jobId
	);
	if (existingRow) {
		const [existingDate, existingCreatedSince] = await Promise.all([
			getCellValue(activeToken, spreadsheetId, sheetName, `A${existingRow}`),
			getCellValue(activeToken, spreadsheetId, sheetName, `S${existingRow}`),
		]);
		if (existingDate) {
			currentData.date = existingDate;
		}
		if (existingCreatedSince) {
			currentData.jobCreatedSince = existingCreatedSince;
		}
		row = buildSheetRow(currentData);
		response = await updateRow(
			activeToken,
			spreadsheetId,
			sheetName,
			existingRow,
			row
		);
	} else {
		row = buildSheetRow(currentData);
		response = await appendRow(activeToken, spreadsheetId, sheetName, row);
	}
	if (response.status === 401 || response.status === 403) {
		await removeCachedToken(activeToken);
		const retryAuth = await getAuthToken(true);
		if (!retryAuth.ok) {
			setStatus("Google authorization failed.", "error");
			return;
		}
		activeToken = retryAuth.token;
		if (existingRow) {
			response = await updateRow(
				activeToken,
				spreadsheetId,
				sheetName,
				existingRow,
				row
			);
		} else {
			response = await appendRow(activeToken, spreadsheetId, sheetName, row);
		}
	}

	if (!response.ok) {
		setStatus(`Sheets API error ${response.status}.`, "error");
		return;
	}
	if (!existingRow) {
		let appendedRowIndex = null;
		try {
			const payload = await response.json();
			appendedRowIndex = getRowIndexFromRange(payload?.updates?.updatedRange);
		} catch (error) {
			appendedRowIndex = null;
		}
		const clearUntil = appendedRowIndex ? appendedRowIndex + 20 : 200;
		await clearBodyBold(
			activeToken,
			spreadsheetId,
			sheetName,
			23,
			clearUntil
		);
		setStatus("Row added to Google Sheets.", "success");
		return;
	}
	setStatus("Row updated in Google Sheets.", "success");
});

const loadSheetsSettings = async () => {
	const settings = await getSheetsSettings();
	spreadsheetId = settings.spreadsheetId;
	sheetName = normalizeSheetName(settings.sheetName);
	spreadsheetInput.value = spreadsheetId;
	sheetNameInput.value = sheetName;
	setButtons();
};

loadSheetsSettings();

const setBidder = (nextBidder) => {
	bidder = String(nextBidder || "").trim();
	if (bidderInput) {
		bidderInput.value = bidder;
	}
};

if (bidderInput) {
	const persistBidder = async () => {
		const nextBidder = String(bidderInput.value || "").trim();
		setBidder(nextBidder);
		await saveBidder(nextBidder);
	};

	bidderInput.addEventListener("change", persistBidder);
	bidderInput.addEventListener("blur", persistBidder);
}

loadBidder().then((storedBidder) => {
	setBidder(storedBidder || "Bilal");
});
setFields(null);
setButtons();

const refreshMode = async () => {
	const tab = await queryActiveTab();
	const mode = detectModeFromUrl(tab?.url || "");
	setMode(mode);
	setButtons();
	if (mode === "proposals") {
		if (!tab || !tab.id || !tab.url) {
			setStatus("No active tab detected.", "warn");
			return;
		}

		setStatus("Scanning proposals page...", "");
		const scan = await executeProposalScan(tab.id);
		if (!scan.ok) {
			setStatus(scan.error || "Unable to scan proposals on this page.", "error");
			return;
		}

		const proposals = Array.isArray(scan.proposals) ? scan.proposals : [];
		const viewed = proposals.filter((item) => item?.viewed && item?.proposalId);

		if (!viewed.length) {
			setStatus(
				`Found ${proposals.length} proposal(s), but none viewed by client yet.`,
				"warn"
			);
			return;
		}

		setStatus(
			`Found ${viewed.length} viewed proposal(s). Click "Update proposal views" to mark as read.`,
			"success"
		);
		return;
	}
	if (!tab || !tab.id || !tab.url) {
		setStatus("No active tab detected.", "warn");
		return;
	}
	if (!tab.url.includes("upwork.com")) {
		setStatus("Open an Upwork job details page to begin.", "");
		return;
	}

	setStatus("Reading current tab...", "");
	const result = await executeExtraction(tab.id);
	if (!result.ok) {
		setStatus(result.error || "Unable to read this page.", "error");
		return;
	}

	currentData = result.data;
	currentData.bidder = bidder;
	setFields(currentData);
	setButtons();
	setStatus("Captured job activity.", "success");
};

refreshMode();
