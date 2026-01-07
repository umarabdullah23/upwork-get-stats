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

const detectModeFromUrl = (url) => {
	if (!url) {
		return "job";
	}
	try {
		const parsed = new URL(url);
		const path = parsed.pathname || "";
		if (/\/Proposals\.html$/i.test(path) || /\/test_data\/Proposals\.html$/i.test(path)) {
			return "proposals";
		}
		if (parsed.hostname.endsWith("upwork.com")) {
			if (path.startsWith("/nx/proposals")) {
				return "proposals";
			}
		}
	} catch (error) {
		// ignore invalid URLs
	}
	return "job";
};

const normalizeSheetName = (value) => {
	const trimmed = String(value || "").trim();
	return trimmed || "Sheet1";
};

const extractSpreadsheetId = (value) => {
	const trimmed = String(value || "").trim();
	if (!trimmed) {
		return "";
	}
	const match = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
	return match ? match[1] : trimmed;
};

const getSheetsSettings = () =>
	new Promise((resolve) => {
		chrome.storage.sync.get(["spreadsheetId", "sheetName"], (result) => {
			if (chrome.runtime.lastError) {
				resolve({ spreadsheetId: "", sheetName: "Sheet1" });
				return;
			}
			resolve({
				spreadsheetId: result.spreadsheetId || "",
				sheetName: result.sheetName || "Sheet1",
			});
		});
	});

const saveSheetsSettings = (nextId, nextName) =>
	new Promise((resolve) => {
		chrome.storage.sync.set(
			{ spreadsheetId: nextId, sheetName: nextName },
			() => {
				resolve(!chrome.runtime.lastError);
			}
		);
	});

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

const escapeCsv = (value) => {
	const text = String(value ?? "");
	if (/[",\n]/.test(text)) {
		return `"${text.replace(/"/g, '""')}"`;
	}
	return text;
};

const formatSheetDate = (value) => {
	const text = String(value || "").trim();
	if (!text) {
		return "";
	}
	return `'${text}`;
};

const toCsvRow = (data) =>
	[
		data.date,
		data.name,
		data.bidder,
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		data.invitesSent,
		data.unansweredInvites,
		data.interviewing,
		"",
		data.payment,
		data.country,
		data.jobCreatedSince,
		"",
		data.proposals,
		data.jobId,
		data.proposalId || "--",
	]
		.map(escapeCsv)
		.join(",");

const buildSheetRow = (data) => {
	const safeName = String(data.name || "").replace(/"/g, '""');
	const safeLink = String(data.link || "").replace(/"/g, '""');
	const nameCell =
		safeName && safeLink
			? `=HYPERLINK("${safeLink}","${safeName}")`
			: data.name || "";

	return [
		formatSheetDate(data.date),
		nameCell,
		data.bidder || "",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		data.invitesSent || "",
		data.unansweredInvites || "",
		data.interviewing || "",
		"",
		data.payment || "",
		data.country || "",
		data.jobCreatedSince || "",
		"",
		data.proposals || "",
		data.jobId || "",
		data.proposalId || "--",
	];
};

const formatRange = (name) => {
	const normalized = normalizeSheetName(name);
	const escaped = normalized.replace(/'/g, "''");
	return `'${escaped}'!A1`;
};

const getAuthToken = (interactive) =>
	new Promise((resolve) => {
		if (!chrome.identity) {
			resolve({ ok: false, error: "Missing chrome.identity permission." });
			return;
		}
		chrome.identity.getAuthToken({ interactive }, (token) => {
			if (chrome.runtime.lastError || !token) {
				const rawMessage =
					chrome.runtime.lastError?.message || "Auth failed.";
				const extensionId = chrome?.runtime?.id || "";
				const normalizedMessage = rawMessage.toLowerCase();
				if (normalizedMessage.includes("bad client id")) {
					resolve({
						ok: false,
						error:
							`OAuth2 failed: bad client id. ` +
							`This usually means the OAuth client in manifest.json isn't a Chrome Extension OAuth client, ` +
							`or it's bound to a different extension ID than the one you installed (${extensionId}).`,
					});
					return;
				}
				resolve({
					ok: false,
					error: rawMessage,
				});
				return;
			}
			resolve({ ok: true, token });
		});
	});

const removeCachedToken = (token) =>
	new Promise((resolve) => {
		if (!chrome.identity) {
			resolve();
			return;
		}
		chrome.identity.removeCachedAuthToken({ token }, () => resolve());
	});

const appendRow = async (token, id, name, row) => {
	const range = formatRange(name);
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(range)}:append`
	);
	url.searchParams.set("valueInputOption", "USER_ENTERED");
	url.searchParams.set("insertDataOption", "INSERT_ROWS");

	return fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ values: [row] }),
	});
};

const findRowByJobId = async (token, id, name, jobId) => {
	if (!jobId) {
		return null;
	}
	const range = `'${normalizeSheetName(name).replace(/'/g, "''")}'!V:V`;
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(range)}`
	);
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	const data = await response.json();
	const values = data.values || [];
	for (let i = 1; i < values.length; i += 1) {
		const cell = values[i]?.[0];
		if (String(cell || "").trim() === String(jobId).trim()) {
			return i + 1;
		}
	}
	return null;
};

const getProposalIdRowMap = async (token, id, name) => {
	const range = `'${normalizeSheetName(name).replace(/'/g, "''")}'!W:W`;
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(range)}`
	);
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	const data = await response.json();
	const values = data.values || [];
	const map = new Map();
	for (let i = 1; i < values.length; i += 1) {
		const cell = values[i]?.[0];
		const normalized = String(cell || "").trim();
		if (normalized) {
			map.set(normalized, i + 1);
		}
	}
	return map;
};

const getCellValue = async (token, id, name, cell) => {
	if (!cell) {
		return "";
	}
	const normalized = normalizeSheetName(name);
	const escaped = normalized.replace(/'/g, "''");
	const range = `'${escaped}'!${cell}`;
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(range)}`
	);
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return "";
	}
	const data = await response.json();
	return data.values?.[0]?.[0] || "";
};

const setReadCellsViewed = async (token, id, name, rowIndexes) => {
	if (!rowIndexes || !rowIndexes.length) {
		return false;
	}
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return false;
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
	const green = { red: 0.8, green: 0.95, blue: 0.8 };
	const requests = rowIndexes.map((rowIndex) => ({
		repeatCell: {
			range: {
				sheetId,
				startRowIndex: rowIndex - 1,
				endRowIndex: rowIndex,
				startColumnIndex: 6,
				endColumnIndex: 7,
			},
			cell: {
				userEnteredFormat: {
					backgroundColor: green,
				},
			},
			fields: "userEnteredFormat.backgroundColor",
		},
	}));
	const response = await fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ requests }),
	});
	return response.ok;
};

const updateRow = async (token, id, name, rowIndex, row) => {
	const normalized = normalizeSheetName(name);
	const escaped = normalized.replace(/'/g, "''");
	const range = `'${escaped}'!A${rowIndex}:W${rowIndex}`;
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(range)}`
	);
	url.searchParams.set("valueInputOption", "USER_ENTERED");
	return fetch(url.toString(), {
		method: "PUT",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ values: [row] }),
	});
};

const clearRowBold = async (token, id, name, rowIndex, columnCount) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null || !rowIndex) {
		return false;
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
	const requests = [
		{
			repeatCell: {
				range: {
					sheetId,
					startRowIndex: rowIndex - 1,
					endRowIndex: rowIndex,
					startColumnIndex: 0,
					endColumnIndex: columnCount,
				},
				cell: {
					userEnteredFormat: {
						textFormat: { bold: false },
					},
				},
				fields: "userEnteredFormat.textFormat.bold",
			},
		},
	];
	const response = await fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ requests }),
	});
	return response.ok;
};

const getRowIndexFromRange = (range) => {
	if (!range) {
		return null;
	}
	const match = range.match(/!([A-Z]+)(\\d+)(:([A-Z]+)(\\d+))?/);
	if (!match) {
		return null;
	}
	const start = Number(match[2]);
	const end = match[5] ? Number(match[5]) : start;
	if (!Number.isFinite(start) || !Number.isFinite(end) || start !== end) {
		return null;
	}
	return start;
};

const setHeaders = async (token, id, name) => {
	const range = formatRange(name);
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(range)}`
	);
	url.searchParams.set("valueInputOption", "RAW");
	const headers = [
		"Date",
		"Job Name",
		"Bidder",
		"Boost",
		"Invited",
		"Customer",
		"Read",
		"Reply",
		"Call",
		"Quote",
		"Sale",
		"Job Status",
		"Invites Sent",
		"Unanswered Invites",
		"Interviewing",
		"Remarks",
		"Payment",
		"Country",
		"Job Created Since",
		"Job Status",
		"Proposals",
		"Job ID",
		"Proposal ID",
	];
	return fetch(url.toString(), {
		method: "PUT",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ values: [headers] }),
	});
};

const setHeaderBold = async (token, id, name, columnCount) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return false;
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
	const requests = [
		{
			repeatCell: {
				range: {
					sheetId,
					startRowIndex: 0,
					endRowIndex: 1,
					startColumnIndex: 0,
					endColumnIndex: columnCount,
				},
				cell: {
					userEnteredFormat: {
						textFormat: { bold: true },
					},
				},
				fields: "userEnteredFormat.textFormat.bold",
			},
		},
	];
	const response = await fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ requests }),
	});
	return response.ok;
};

const freezeHeaderRow = async (token, id, name) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return false;
	}
	const requests = [
		{
			updateSheetProperties: {
				properties: {
					sheetId,
					gridProperties: { frozenRowCount: 1 },
				},
				fields: "gridProperties.frozenRowCount",
			},
		},
	];
	const response = await fetch(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`,
		{
			method: "POST",
			headers: {
				Authorization: `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: JSON.stringify({ requests }),
		}
	);
	return response.ok;
};

const clearBodyBold = async (token, id, name, columnCount, endRowIndex) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return false;
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
	const requests = [
		{
			repeatCell: {
				range: {
					sheetId,
					startRowIndex: 1,
					endRowIndex,
					startColumnIndex: 0,
					endColumnIndex: columnCount,
				},
				cell: {
					userEnteredFormat: {
						textFormat: { bold: false },
					},
				},
				fields: "userEnteredFormat.textFormat.bold",
			},
		},
	];
	const response = await fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ requests }),
	});
	return response.ok;
};

const setBodyColumnColors = async (token, id, name, endRowIndex) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return false;
	}
	const purple = { red: 0.91, green: 0.84, blue: 0.89 };
	const red = { red: 0.96, green: 0.8, blue: 0.79 };
	const requests = [
		{
			repeatCell: {
				range: {
					sheetId,
					startRowIndex: 1,
					endRowIndex,
					startColumnIndex: 3,
					endColumnIndex: 4,
				},
				cell: {
					userEnteredFormat: {
						backgroundColor: purple,
					},
				},
				fields: "userEnteredFormat.backgroundColor",
			},
		},
		{
			repeatCell: {
				range: {
					sheetId,
					startRowIndex: 1,
					endRowIndex,
					startColumnIndex: 6,
					endColumnIndex: 11,
				},
				cell: {
					userEnteredFormat: {
						backgroundColor: red,
					},
				},
				fields: "userEnteredFormat.backgroundColor",
			},
		},
	];
	const response = await fetch(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`,
		{
			method: "POST",
			headers: {
				Authorization: `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: JSON.stringify({ requests }),
		}
	);
	return response.ok;
};

const getSheetId = async (token, id, name) => {
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}`
	);
	url.searchParams.set("fields", "sheets.properties");
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	const data = await response.json();
	const match = (data.sheets || []).find(
		(sheet) => sheet?.properties?.title === name
	);
	return match?.properties?.sheetId ?? null;
};

const ensureBidderDropdown = async (token, id, name) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return;
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
	const requests = [
		{
			repeatCell: {
				range: {
					sheetId,
					startRowIndex: 1,
					startColumnIndex: 2,
					endColumnIndex: 3,
				},
				cell: {
					dataValidation: {
						condition: {
							type: "ONE_OF_LIST",
							values: [
								{ userEnteredValue: "Bilal" },
								{ userEnteredValue: "Mehak" },
							],
						},
						strict: true,
						showCustomUi: true,
					},
				},
				fields: "dataValidation",
			},
		},
	];
	await fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ requests }),
	});
};

const queryActiveTab = () =>
	new Promise((resolve) => {
		chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
			resolve(tabs && tabs.length ? tabs[0] : null);
		});
	});

const executeExtraction = (tabId) =>
	new Promise((resolve) => {
		chrome.scripting.executeScript(
			{
				target: { tabId },
				func: extractUpworkJobData,
			},
			(results) => {
				if (chrome.runtime.lastError) {
					resolve({ ok: false, error: chrome.runtime.lastError.message });
					return;
				}
				resolve(results && results[0] ? results[0].result : { ok: false });
			}
		);
	});

const executeProposalScan = (tabId) =>
	new Promise((resolve) => {
		chrome.scripting.executeScript(
			{
				target: { tabId },
				func: extractUpworkViewedProposals,
			},
			(results) => {
				if (chrome.runtime.lastError) {
					resolve({ ok: false, error: chrome.runtime.lastError.message });
					return;
				}
				resolve(results && results[0] ? results[0].result : { ok: false });
			}
		);
	});

function extractUpworkJobData() {
	const normalize = (value) =>
		String(value || "")
			.replace(/\s+/g, " ")
			.trim();
	const getText = (el) => normalize(el ? el.textContent : "");

	const data = {
		name: "",
		link: "",
		jobId: "",
		date: "",
		payment: "",
		country: "",
		jobCreatedSince: "",
		proposals: "",
		invitesSent: "",
		interviewing: "",
		lastViewed: "",
		unansweredInvites: "",
		proposalId: "",
	};

	const url = new URL(window.location.href);
	data.link = `${url.origin}${url.pathname}`;
	const jobMatch = window.location.href.match(/~([0-9]+)/);
	data.jobId = jobMatch ? jobMatch[1] : "";
	data.date = new Date().toLocaleDateString("en-US", {
		month: "short",
		day: "numeric",
		year: "numeric",
	});

	const postedLine =
		document.querySelector(".posted-on-line") ||
		document.querySelector("[data-test='job-details'] .posted-on-line");
	if (postedLine) {
		const postedText = getText(
			postedLine.querySelector(".text-light-on-muted")
		);
		if (postedText) {
			data.jobCreatedSince = postedText;
		}
	}

	const clientLocation =
		document.querySelector("[data-qa='client-location'] strong") ||
		document.querySelector(".cfe-ui-job-about-client [data-qa='client-location'] strong");
	const locationText = getText(clientLocation);
	if (locationText) {
		data.country = locationText;
	}

	const paymentText = normalize(document.body.textContent || "");
	if (paymentText.toLowerCase().includes("payment method verified")) {
		data.payment = "Verified";
	} else if (paymentText.toLowerCase().includes("payment method not verified")) {
		data.payment = "Not verified";
	}

	const jobDetails =
		document.querySelector(".job-details") ||
		document.querySelector("[data-ev-sublocation='jobdetails']") ||
		document.querySelector("[data-test='job-details']");

	const headingCandidates = [];
	if (jobDetails) {
		headingCandidates.push(...jobDetails.querySelectorAll("h1, h2, h3, h4"));
	}
	headingCandidates.push(...document.querySelectorAll("h1, h2, h3, h4"));

	for (const heading of headingCandidates) {
		const text = getText(heading);
		if (text && text.toLowerCase() !== "activity on this job") {
			data.name = text;
			break;
		}
	}

	if (!data.name) {
		const title = normalize(document.title);
		data.name = title.replace(/\s*-\s*Upwork.*$/i, "").trim();
	}

	const headings = Array.from(
		document.querySelectorAll("h1, h2, h3, h4, h5, h6")
	);
	const activityHeading = headings.find(
		(heading) => getText(heading).toLowerCase() === "activity on this job"
	);
	const activitySection = activityHeading
		? activityHeading.closest("section") || activityHeading.parentElement
		: null;

	const items = activitySection
		? activitySection.querySelectorAll(".ca-item, li")
		: document.querySelectorAll(".ca-item");

	items.forEach((item) => {
		let label = getText(item.querySelector(".title"));
		let value = getText(item.querySelector(".value"));

		if (!label || !value) {
			const text = getText(item);
			const parts = text.split(":");
			if (parts.length >= 2) {
				label = label || parts[0];
				value = value || parts.slice(1).join(":").trim();
			}
		}

		if (!label || !value) {
			return;
		}

		const key = label.toLowerCase().replace(/\s+/g, " ").replace(/:$/, "");
		if (key === "proposals") {
			data.proposals = value;
		} else if (key === "last viewed by client") {
			data.lastViewed = value;
		} else if (key === "interviewing") {
			data.interviewing = value;
		} else if (key === "invites sent") {
			data.invitesSent = value;
		} else if (key === "unanswered invites") {
			data.unansweredInvites = value;
		}
	});

	const proposalLink =
		Array.from(document.querySelectorAll("a")).find((anchor) => {
			const href = anchor.getAttribute("href") || "";
			if (!/\/(ab\/proposals|nx\/proposals)\/\d+/.test(href)) {
				return false;
			}
			return getText(anchor).toLowerCase().includes("view proposal");
		}) ||
		document.querySelector("a[href*='/ab/proposals/'], a[href*='/nx/proposals/']");

	const proposalHref = proposalLink?.getAttribute("href") || "";
	const proposalMatch = proposalHref.match(/\/(ab\/proposals|nx\/proposals)\/(\d+)/);
	data.proposalId = proposalMatch ? proposalMatch[2] : "";

	const hasData =
		data.name ||
		data.proposals ||
		data.invitesSent ||
		data.interviewing ||
		data.lastViewed ||
		data.unansweredInvites;

	if (!hasData) {
		return { ok: false, error: "No job activity found on this page." };
	}

	return { ok: true, data };
}

function extractUpworkViewedProposals() {
	const normalize = (value) =>
		String(value || "")
			.replace(/\s+/g, " ")
			.trim();

	const results = [];
	const rows = Array.from(document.querySelectorAll("tr.details-row"));

	for (const row of rows) {
		const link = row.querySelector("td.job-info a[href*='/nx/proposals/']");
		if (!link) {
			continue;
		}
		const href = link.getAttribute("href") || "";
		const match = href.match(/\/nx\/proposals\/(\d+)/);
		if (!match) {
			continue;
		}

		const viewsCell = row.querySelector("td.proposal-views-slot");
		let viewed = false;
		if (viewsCell && !viewsCell.classList.contains("d-none")) {
			const text = normalize(viewsCell.textContent).toLowerCase();
			viewed =
				text.includes("viewed by client") ||
				Boolean(viewsCell.querySelector(".last-seen")) ||
				Boolean(viewsCell.querySelector("[id^='proposal-views-mark']"));
		}

		results.push({
			proposalId: match[1],
			title: normalize(link.textContent),
			viewed,
		});
	}

	return { ok: true, proposals: results };
}

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

const loadBidder = () =>
	new Promise((resolve) => {
		chrome.storage.sync.get(["bidder"], (result) => {
			if (chrome.runtime.lastError) {
				resolve("");
				return;
			}
			resolve(result.bidder || "");
		});
	});

const saveBidder = (nextBidder) =>
	new Promise((resolve) => {
		chrome.storage.sync.set({ bidder: nextBidder }, () => {
			resolve(!chrome.runtime.lastError);
		});
	});

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
		setStatus("Open proposals page and click “Check proposal views”.", "");
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
