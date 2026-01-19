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
const populateJobIdsButton = document.getElementById("populate-job-ids");
const openTemplateButton = document.getElementById("open-template");
const sendSheetsButton = document.getElementById("send-sheets");
const statusEl = document.getElementById("status");
const bidderInput = document.getElementById("bidder-input");
const openDevPanelButton = document.getElementById("open-dev-panel");

const TESTING_PASSWORD = "umar";

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
let currentTabUrl = "";

const setStatus = (message, tone = "") => {
	statusEl.textContent = message;
	statusEl.className = `status ${tone}`.trim();
};

const requestAuthToken = async (interactive, timeoutMs = 15000) => {
	const timeout = new Promise((resolve) => {
		setTimeout(
			() =>
				resolve({
					ok: false,
					error: "Google authorization timed out.",
				}),
			timeoutMs
		);
	});
	return Promise.race([getAuthToken(interactive), timeout]);
};

const withTimeout = (promise, timeoutMs, errorMessage) =>
	Promise.race([
		promise,
		new Promise((_, reject) =>
			setTimeout(() => reject(new Error(errorMessage)), timeoutMs)
		),
	]);

const isSheetUrl = (value) =>
	/https?:\/\/docs\.google\.com\/spreadsheets/i.test(String(value || ""));

const extractHyperlinkUrl = (value) => {
	const text = String(value || "").trim();
	if (!text) {
		return "";
	}
	if (text.startsWith("=")) {
		const match = text.match(/HYPERLINK\(\s*\"([^\"]+)\"/i);
		return match ? match[1] : "";
	}
	return text;
};

const extractJobIdFromLink = (value) => {
	const url = extractHyperlinkUrl(value);
	if (!url) {
		return "";
	}
	const match = url.match(/~([A-Za-z0-9]+)/);
	return match ? match[1] : "";
};

const getValidationOptions = (dataValidation) => {
	const condition = dataValidation?.condition;
	if (!condition) {
		return null;
	}
	const type = String(condition.type || "");
	if (!type.includes("ONE_OF_LIST")) {
		return null;
	}
	return (condition.values || [])
		.map((value) => String(value?.userEnteredValue || "").trim())
		.filter(Boolean);
};

const validationAllowsValue = (dataValidation, value) => {
	const trimmed = String(value || "").trim();
	if (!trimmed) {
		return true;
	}
	const options = getValidationOptions(dataValidation);
	if (!options || !options.length) {
		return true;
	}
	return options.includes(trimmed);
};

const verifyTestingPassword = (actionLabel) => {
	const entered = window.prompt(`Enter password to ${actionLabel}:`);
	if (entered === null) {
		return false;
	}
	return entered === TESTING_PASSWORD;
};

const parseConnectsCellValue = (value) => {
	const match = String(value || "").replace(/,/g, "").match(/[+-]?\d+(\.\d+)?/);
	if (!match) {
		return 0;
	}
	const parsed = Number(match[0]);
	return Number.isFinite(parsed) ? Math.abs(parsed) : 0;
};

const formatConnectsValue = (value, kind) => {
	const numeric = Number(value);
	if (!Number.isFinite(numeric) || numeric <= 0) {
		return "";
	}
	const prefix = kind === "refund" ? "+" : "-";
	return `${prefix}${Math.abs(numeric)}`;
};

const rowNeedsTemplateValidation = async (token, id, name, rowIndex) => {
	if (!rowIndex) {
		return false;
	}
	const bidderValidation = await getCellDataValidation(
		token,
		id,
		name,
		`C${rowIndex}`
	);
	if (!bidderValidation) {
		return true;
	}
	const jobStatusValidation = await getCellDataValidation(
		token,
		id,
		name,
		`L${rowIndex}`
	);
	return !jobStatusValidation;
};

const applyExistingJobStatus = (row, emptyRowInfo, rowIndex) => {
	if (!row || !emptyRowInfo || !rowIndex) {
		return row;
	}
	const jobStatusIndexes = emptyRowInfo.jobStatusIndexes || [];
	const statusValues = emptyRowInfo.jobStatusValuesByRow?.get(rowIndex);
	if (!jobStatusIndexes.length || !statusValues) {
		return row;
	}
	jobStatusIndexes.forEach((columnIndex, index) => {
		const existing = String(statusValues[index] || "").trim();
		if (existing) {
			row[columnIndex] = existing;
		}
	});
	return row;
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
	if (populateJobIdsButton) {
		const showPopulate = isSheetUrl(currentTabUrl);
		populateJobIdsButton.hidden = !showPopulate;
		populateJobIdsButton.disabled = !showPopulate;
	}
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
		if (currentJobActions) {
			currentJobActions.hidden = false;
		}
		if (readButton) {
			readButton.hidden = false;
		}
		if (sendSheetsButton) {
			sendSheetsButton.hidden = true;
		}
		if (copyButton) {
			copyButton.hidden = true;
		}
		checkViewedButton.textContent = "Update proposal views";
		checkViewedButton.hidden = false;
	} else if (mode === "connects") {
		if (currentJobCard) {
			currentJobCard.hidden = false;
		}
		if (currentJobStatus) {
			currentJobStatus.textContent = "Connects history detected";
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
		if (currentJobActions) {
			currentJobActions.hidden = false;
		}
		if (readButton) {
			readButton.hidden = false;
		}
		if (sendSheetsButton) {
			sendSheetsButton.hidden = true;
		}
		if (copyButton) {
			copyButton.hidden = true;
		}
		checkViewedButton.textContent = "Update connects history";
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
	if (uiMode !== "job") {
		await refreshMode();
		return;
	}
	setStatus("Reading current tab...", "");
	const tab = await queryActiveTab();

	if (!tab || !tab.id) {
		setStatus("No active tab detected.", "warn");
		return;
	}

	const result = await executeExtraction(tab.id);
	if (!result.ok) {
		if (result.error === "No job activity found on this page.") {
			setStatus("Open an Upwork job details page first.", "warn");
		} else {
			setStatus(result.error || "Unable to read this page.", "error");
		}
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
		setStatus("Scanning page...", "");
		const tab = await queryActiveTab();

		if (!tab || !tab.id) {
			setStatus("No active tab detected.", "warn");
			return;
		}

		if (!spreadsheetId) {
			setStatus("Save your Sheets settings first.", "warn");
			return;
		}

		const mode = detectModeFromUrl(tab.url || "");
		if (mode === "proposals") {
			const scan = await executeProposalScan(tab.id);
			if (!scan.ok) {
				setStatus(
					scan.error || "Unable to scan proposals on this page.",
					"error"
				);
				return;
			}

			const proposals = Array.isArray(scan.proposals) ? scan.proposals : [];
			const viewed = proposals.filter(
				(item) => item?.viewed && item?.proposalId
			);
			if (!viewed.length) {
				setStatus("No 'Viewed by client' proposals found on this page.", "warn");
				return;
			}

			setStatus(
				`Found ${viewed.length} viewed proposal(s). Updating sheet...`,
				""
			);
			const auth = await requestAuthToken(true);
			if (!auth.ok) {
				setStatus(auth.error || "Google authorization failed.", "error");
				return;
			}

			setStatus("Reading sheet data...", "");
			let headers;
			let idMap;
			try {
				headers = await withTimeout(
					getSheetHeaders(auth.token, spreadsheetId, sheetName),
					20000,
					"Sheet header read timed out."
				);
				idMap = await withTimeout(
					getProposalIdRowMap(
						auth.token,
						spreadsheetId,
						sheetName,
						headers
					),
					20000,
					"Sheet Proposal ID read timed out."
				);
			} catch (error) {
				setStatus(error.message || "Sheet read timed out.", "error");
				return;
			}
			if (!idMap) {
				setStatus("Unable to read Proposal ID column from the sheet.", "error");
				return;
			}

			const matchedRows = [];
			let nameMap;
			let nameMapLoaded = false;
			for (const item of viewed) {
				const rowIndex = idMap.get(String(item.proposalId).trim());
				if (rowIndex) {
					matchedRows.push(rowIndex);
					continue;
				}
				const normalized = normalizeJobName(item.title || "");
				if (!normalized) {
					continue;
				}
				if (!nameMapLoaded) {
					nameMap = await getJobNameRowMap(
						auth.token,
						spreadsheetId,
						sheetName,
						headers
					);
					nameMapLoaded = true;
				}
				const fallbackRow = nameMap?.get(normalized);
				if (fallbackRow) {
					matchedRows.push(fallbackRow);
				}
			}
			const uniqueRows = Array.from(new Set(matchedRows)).sort((a, b) => a - b);
			if (!uniqueRows.length) {
				setStatus(
					`Found ${viewed.length} viewed proposal(s), but none matched the sheet's Proposal ID or Job Name columns.`,
					"warn"
				);
				return;
			}

		const readColumnIndex = getHeaderIndex(headers || [], "Read");
		const ok = await setReadCellsViewed(
			auth.token,
			spreadsheetId,
			sheetName,
			uniqueRows,
			readColumnIndex ? readColumnIndex - 1 : undefined
		);
			if (!ok) {
				setStatus("Failed to update the sheet formatting.", "error");
				return;
			}
			setStatus(
				`Marked ${uniqueRows.length} row(s) as read (green).`,
				"success"
			);
			return;
		}

		if (mode !== "connects") {
			setStatus("Open the connects history page first.", "warn");
			return;
		}

		const scan = await executeConnectsHistoryScan(tab.id);
		if (!scan.ok) {
			setStatus(scan.error || "Unable to scan connects history.", "error");
			return;
		}

		const entries = Array.isArray(scan.entries) ? scan.entries : [];
		if (!entries.length) {
			setStatus("No connects history entries found.", "warn");
			return;
		}

		const totals = new Map();
		for (const entry of entries) {
			const jobIdValue = entry?.jobId ? String(entry.jobId).trim() : "";
			const key = jobIdValue || normalizeJobName(entry?.title || "");
			if (!key) {
				continue;
			}
			if (!totals.has(key)) {
				totals.set(key, {
					jobId: jobIdValue,
					name: entry.title || "",
					link: entry.link || "",
					date: entry.date || "",
					connectsSpent: 0,
					connectsRefund: 0,
					boostedConnectsSpent: 0,
					boostedConnectsRefund: 0,
				});
			}
			const target = totals.get(key);
			target.connectsSpent += Number(entry.connectsSpent || 0);
			target.connectsRefund += Number(entry.connectsRefund || 0);
			target.boostedConnectsSpent += Number(entry.boostedConnectsSpent || 0);
			target.boostedConnectsRefund += Number(
				entry.boostedConnectsRefund || 0
			);
			if (!target.name && entry.title) {
				target.name = entry.title;
			}
			if (!target.link && entry.link) {
				target.link = entry.link;
			}
			if (!target.date && entry.date) {
				target.date = entry.date;
			}
		}

		if (!totals.size) {
			setStatus(
				"No connects history entries with job IDs or job names found.",
				"warn"
			);
			return;
		}

		setStatus(`Found ${totals.size} job(s). Updating sheet...`, "");
		const auth = await requestAuthToken(true);
		if (!auth.ok) {
			setStatus(auth.error || "Google authorization failed.", "error");
			return;
		}

		setStatus("Reading sheet data...", "");
		let connectsMap;
		try {
			connectsMap = await withTimeout(
				getConnectsRowMap(auth.token, spreadsheetId, sheetName),
				20000,
				"Sheet Job ID read timed out."
			);
		} catch (error) {
			setStatus(error.message || "Sheet read timed out.", "error");
			return;
		}
		if (!connectsMap) {
			setStatus("Unable to read Job ID column from the sheet.", "error");
			return;
		}
		const needsBoostedColumns = Array.from(totals.values()).some(
			(entry) =>
				entry.boostedConnectsSpent > 0 || entry.boostedConnectsRefund > 0
		);
		if (
			needsBoostedColumns &&
			(!connectsMap.boostedConnectsSpentColumn ||
				!connectsMap.boostedConnectsRefundColumn)
		) {
			setStatus(
				"Sheet is missing boosted connects columns. Run Prepare sheet first.",
				"warn"
			);
			return;
		}

		const updates = [];
		const newRowUpdates = [];
		const headers = connectsMap.headers || [];
		const activeBidder = String(bidderInput?.value || bidder || "").trim();
		let jobNameMap;
		let jobNameMapLoaded = false;
		let emptyRowInfo;
		try {
			emptyRowInfo = await withTimeout(
				getEmptyRowIndexes(
					auth.token,
					spreadsheetId,
					sheetName,
					headers
				),
				20000,
				"Sheet row scan timed out."
			);
		} catch (error) {
			setStatus(error.message || "Sheet read timed out.", "error");
			return;
		}
		if (!emptyRowInfo) {
			setStatus("Unable to read sheet rows.", "error");
			return;
		}
		setStatus("Updating sheet...", "");
		const emptyRowQueue = [...emptyRowInfo.emptyRows];
		let nextRowIndex = emptyRowInfo.nextRowIndex;
		const newRowIndexes = [];
		const appendedRowIndexes = [];
		for (const entry of totals.values()) {
			let existing = entry.jobId ? connectsMap.map.get(entry.jobId) : null;
			if (!existing) {
				const normalizedName = normalizeJobName(entry.name || "");
				if (normalizedName) {
					if (!jobNameMapLoaded) {
						jobNameMap = await getJobNameRowMap(
							auth.token,
							spreadsheetId,
							sheetName,
							headers
						);
						jobNameMapLoaded = true;
					}
					const fallbackRowIndex = jobNameMap?.get(normalizedName);
					if (fallbackRowIndex) {
						existing = connectsMap.rowsByIndex.get(fallbackRowIndex) || {
							rowIndex: fallbackRowIndex,
							connectsSpent: "",
							connectsRefund: "",
							boostedConnectsSpent: "",
							boostedConnectsRefund: "",
						};
					}
				}
			}
			if (existing) {
				const nextSpentValue = formatConnectsValue(
					entry.connectsSpent,
					"spent"
				);
				const nextRefundValue = formatConnectsValue(
					entry.connectsRefund,
					"refund"
				);
				const nextBoostedSpentValue = formatConnectsValue(
					entry.boostedConnectsSpent,
					"spent"
				);
				const nextBoostedRefundValue = formatConnectsValue(
					entry.boostedConnectsRefund,
					"refund"
				);
				if (activeBidder) {
					updates.push({
						range: `'${normalizeSheetName(sheetName).replace(/'/g, "''")}'!C${existing.rowIndex}:C${existing.rowIndex}`,
						values: [[activeBidder]],
					});
				}
				if (connectsMap.connectsSpentColumn) {
					updates.push({
						range: `'${normalizeSheetName(sheetName).replace(/'/g, "''")}'!${connectsMap.connectsSpentColumn}${existing.rowIndex}`,
						values: [[nextSpentValue]],
					});
				}
				if (connectsMap.connectsRefundColumn) {
					updates.push({
						range: `'${normalizeSheetName(sheetName).replace(/'/g, "''")}'!${connectsMap.connectsRefundColumn}${existing.rowIndex}`,
						values: [[nextRefundValue]],
					});
				}
				if (connectsMap.boostedConnectsSpentColumn) {
					updates.push({
						range: `'${normalizeSheetName(sheetName).replace(/'/g, "''")}'!${connectsMap.boostedConnectsSpentColumn}${existing.rowIndex}`,
						values: [[nextBoostedSpentValue]],
					});
				}
				if (connectsMap.boostedConnectsRefundColumn) {
					updates.push({
						range: `'${normalizeSheetName(sheetName).replace(/'/g, "''")}'!${connectsMap.boostedConnectsRefundColumn}${existing.rowIndex}`,
						values: [[nextBoostedRefundValue]],
					});
				}
				continue;
			}

			const row = buildRowFromHeaders(headers, {
				date: entry.date,
				name: entry.name,
				link: entry.link,
				bidder: activeBidder,
				jobId: entry.jobId,
				connectsSpent: formatConnectsValue(entry.connectsSpent, "spent"),
				connectsRefund: formatConnectsValue(entry.connectsRefund, "refund"),
				boostedConnectsSpent: formatConnectsValue(
					entry.boostedConnectsSpent,
					"spent"
				),
				boostedConnectsRefund: formatConnectsValue(
					entry.boostedConnectsRefund,
					"refund"
				),
				proposalId: "--",
			});
			const useEmptyRow = emptyRowQueue.length > 0;
			const targetRowIndex = useEmptyRow
				? emptyRowQueue.shift()
				: nextRowIndex;
			if (!useEmptyRow) {
				nextRowIndex += 1;
			}
			applyExistingJobStatus(row, emptyRowInfo, targetRowIndex);
			if (!useEmptyRow) {
				appendedRowIndexes.push(targetRowIndex);
			}
			updates.push({
				range: `'${normalizeSheetName(sheetName).replace(/'/g, "''")}'!A${targetRowIndex}:${getColumnLetter(connectsMap.columnCount)}${targetRowIndex}`,
				values: [row],
			});
			newRowUpdates.push({
				range: `'${normalizeSheetName(sheetName).replace(/'/g, "''")}'!A${targetRowIndex}:${getColumnLetter(connectsMap.columnCount)}${targetRowIndex}`,
				values: [row],
			});
			newRowIndexes.push(targetRowIndex);
		}

		const response = await batchUpdateValues(
			auth.token,
			spreadsheetId,
			updates
		);
		if (!response.ok) {
			setStatus(`Sheets API error ${response.status}.`, "error");
			return;
		}
		setStatus("Formatting in progress...", "");
		try {
			const targetSheetId = await getSheetId(
				auth.token,
				spreadsheetId,
				sheetName
			);
			const rowsNeedingTemplate = [];
			for (const rowIndex of newRowIndexes) {
				const needsTemplate = await rowNeedsTemplateValidation(
					auth.token,
					spreadsheetId,
					sheetName,
					rowIndex
				);
				if (needsTemplate) {
					rowsNeedingTemplate.push(rowIndex);
				}
			}
			if (rowsNeedingTemplate.length) {
				let templateSheetId = await ensureTemplateSheet(
					auth.token,
					spreadsheetId
				);
				const bidderValidation = templateSheetId
					? await getCellDataValidation(
							auth.token,
							spreadsheetId,
							"__Upwork Template",
							"C2"
						)
					: null;
				if (
					templateSheetId &&
					activeBidder &&
					!validationAllowsValue(bidderValidation, activeBidder)
				) {
					templateSheetId = await ensureTemplateSheet(
						auth.token,
						spreadsheetId,
						true
					);
				}
				let appliedTemplate = false;
				if (templateSheetId && targetSheetId !== null) {
					appliedTemplate = await applyTemplateRowToRows(
						auth.token,
						spreadsheetId,
						templateSheetId,
						targetSheetId,
						rowsNeedingTemplate,
						DEFAULT_HEADER_COUNT
					);
				}
				if (!appliedTemplate) {
					await applyRowTemplatesInSheet(
						auth.token,
						spreadsheetId,
						sheetName,
						2,
						rowsNeedingTemplate,
						DEFAULT_HEADER_COUNT
					);
				}
				for (const rowIndex of rowsNeedingTemplate) {
					await clearRowsValues(
						auth.token,
						spreadsheetId,
						sheetName,
						rowIndex,
						rowIndex,
						DEFAULT_HEADER_COUNT
					);
				}
				if (newRowUpdates.length) {
					await batchUpdateValues(
						auth.token,
						spreadsheetId,
						newRowUpdates
					);
				}
			}
			setStatus("Connects history updated.", "success");
		} catch (error) {
			setStatus("Connects history updated; formatting skipped.", "warn");
		}
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

if (openTemplateButton) {
	openTemplateButton.addEventListener("click", () => {
		const templateId = window.TEMPLATE_SPREADSHEET_ID;
		if (!templateId) {
			setStatus("Reference sheet not configured.", "warn");
			return;
		}
		openSpreadsheet(templateId);
	});
}

if (openDevPanelButton) {
	openDevPanelButton.addEventListener("click", () => {
		if (!verifyTestingPassword("open the dev panel")) {
			setStatus("Dev panel access canceled or password incorrect.", "warn");
			return;
		}
		const url = chrome?.runtime?.getURL
			? chrome.runtime.getURL("devpanel/devpanel.html")
			: "devpanel/devpanel.html";
		if (chrome?.tabs?.create) {
			chrome.tabs.create({ url });
		} else {
			window.open(url, "_blank", "noopener");
		}
	});
}

if (populateJobIdsButton) {
	populateJobIdsButton.addEventListener("click", async () => {
		if (!spreadsheetId) {
			setStatus("Save your Sheets settings first.", "warn");
			return;
		}
		setStatus("Connecting to Google...", "");
		const auth = await requestAuthToken(true);
		if (!auth.ok) {
			setStatus(auth.error || "Google authorization failed.", "error");
			return;
		}
		setStatus("Reading sheet data...", "");
		let headers;
		try {
			headers = await withTimeout(
				getSheetHeaders(auth.token, spreadsheetId, sheetName),
				20000,
				"Sheet header read timed out."
			);
		} catch (error) {
			setStatus(error.message || "Sheet read timed out.", "error");
			return;
		}
		if (!headers) {
			setStatus("Unable to read sheet headers.", "error");
			return;
		}
		const jobNameIndex = getHeaderIndex(headers, "Job Name");
		const jobIdIndex = getHeaderIndex(headers, "Job ID");
		if (!jobNameIndex || !jobIdIndex) {
			setStatus("Sheet needs Job Name and Job ID columns.", "error");
			return;
		}
		let jobNameValues;
		let jobIdValues;
		try {
			[jobNameValues, jobIdValues] = await Promise.all([
				withTimeout(
					getColumnValues(
						auth.token,
						spreadsheetId,
						sheetName,
						jobNameIndex - 1,
						2,
						{ valueRenderOption: "FORMULA" }
					),
					20000,
					"Sheet Job Name read timed out."
				),
				withTimeout(
					getColumnValues(
						auth.token,
						spreadsheetId,
						sheetName,
						jobIdIndex - 1,
						2
					),
					20000,
					"Sheet Job ID read timed out."
				),
			]);
		} catch (error) {
			setStatus(error.message || "Sheet read timed out.", "error");
			return;
		}
		if (!jobNameValues) {
			setStatus("Unable to read Job Name column.", "error");
			return;
		}
		const updates = [];
		const rowCount = Math.max(
			jobNameValues.length,
			jobIdValues?.length || 0
		);
		const columnLetter = getColumnLetter(jobIdIndex);
		const escapedName = normalizeSheetName(sheetName).replace(/'/g, "''");
		for (let i = 0; i < rowCount; i += 1) {
			const jobId = extractJobIdFromLink(jobNameValues[i]?.[0] || "");
			if (!jobId) {
				continue;
			}
			const existing = String(jobIdValues?.[i]?.[0] || "").trim();
			if (existing === jobId) {
				continue;
			}
			const rowIndex = i + 2;
			updates.push({
				range: `'${escapedName}'!${columnLetter}${rowIndex}`,
				values: [[jobId]],
			});
		}
		if (!updates.length) {
			setStatus("Job IDs already up to date.", "success");
			return;
		}
		setStatus("Populating job IDs...", "");
		const response = await batchUpdateValues(
			auth.token,
			spreadsheetId,
			updates
		);
		if (!response.ok) {
			setStatus(`Sheets API error ${response.status}.`, "error");
			return;
		}
		setStatus(`Job IDs populated (${updates.length}).`, "success");
	});
}

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
	if (!verifyTestingPassword("prepare the sheet")) {
		setStatus("Sheet preparation canceled or password incorrect.", "warn");
		return;
	}
	setStatus("Preparing sheet...", "");
	const auth = await requestAuthToken(true);
	if (!auth.ok) {
		setStatus(auth.error || "Google authorization failed.", "error");
		return;
	}
	await removeBlankHeaderColumn(
		auth.token,
		spreadsheetId,
		sheetName,
		18,
		"Job Status"
	);
	const templateResult = await applyTemplateFormatting(
		auth.token,
		spreadsheetId,
		sheetName
	);
	if (templateResult.ok) {
		const headerResponse = await setHeaders(
			auth.token,
			spreadsheetId,
			sheetName
		);
		if (!headerResponse.ok) {
			setStatus(`Sheets API error ${headerResponse.status}.`, "error");
			return;
		}
		await setHeaderBold(
			auth.token,
			spreadsheetId,
			sheetName,
			DEFAULT_HEADER_COUNT
		);
		await freezeHeaderRow(auth.token, spreadsheetId, sheetName);
		setStatus("Sheet prepared from template.", "success");
		return;
	}
	const headerResponse = await setHeaders(auth.token, spreadsheetId, sheetName);
	if (!headerResponse.ok) {
		setStatus(`Sheets API error ${headerResponse.status}.`, "error");
		return;
	}
	await setHeaderBold(
		auth.token,
		spreadsheetId,
		sheetName,
		DEFAULT_HEADER_COUNT
	);
	await clearBodyBold(
		auth.token,
		spreadsheetId,
		sheetName,
		DEFAULT_HEADER_COUNT,
		1000
	);
	await setBodyColumnColors(auth.token, spreadsheetId, sheetName, 1000);
	await freezeHeaderRow(auth.token, spreadsheetId, sheetName);
	await ensureJobStatusDropdown(auth.token, spreadsheetId, sheetName);
	await ensureJobStatusColors(auth.token, spreadsheetId, sheetName);
	setStatus("Sheet prepared with headers and dropdowns.", "success");
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
	const auth = await requestAuthToken(true);
	if (!auth.ok) {
		setStatus(auth.error || "Google authorization failed.", "error");
		return;
	}
	let activeToken = auth.token;

	let row = null;
	let newRowIndex = null;
	let emptyRowInfo = null;
	setStatus("Reading sheet data...", "");
	let headers;
	try {
		headers = await withTimeout(
			getSheetHeaders(activeToken, spreadsheetId, sheetName),
			20000,
			"Sheet header read timed out."
		);
	} catch (error) {
		setStatus(error.message || "Sheet read timed out.", "error");
		return;
	}
	if (!headers) {
		setStatus("Unable to read sheet headers.", "error");
		return;
	}
	setStatus("Updating sheet...", "");
	let response;
	let existingRow = await findRowByJobId(
		activeToken,
		spreadsheetId,
		sheetName,
		currentData.jobId,
		headers
	);
	if (existingRow) {
		const [existingDate, existingCreatedSince] = await Promise.all([
			getCellValue(activeToken, spreadsheetId, sheetName, `A${existingRow}`),
			getCellValue(activeToken, spreadsheetId, sheetName, `R${existingRow}`),
		]);
		if (existingDate) {
			currentData.date = existingDate;
		}
		if (existingCreatedSince) {
			currentData.jobCreatedSince = existingCreatedSince;
		}
		row = buildRowFromHeaders(headers, currentData);
		response = await updateRow(
			activeToken,
			spreadsheetId,
			sheetName,
			existingRow,
			row
		);
	} else {
		try {
			emptyRowInfo = await withTimeout(
				getEmptyRowIndexes(
					activeToken,
					spreadsheetId,
					sheetName,
					headers
				),
				20000,
				"Sheet row scan timed out."
			);
		} catch (error) {
			setStatus(error.message || "Sheet read timed out.", "error");
			return;
		}
		if (!emptyRowInfo) {
			setStatus("Unable to read sheet rows.", "error");
			return;
		}
		const hasEmptyRow = emptyRowInfo.emptyRows.length > 0;
		newRowIndex = hasEmptyRow
			? emptyRowInfo.emptyRows[0]
			: emptyRowInfo.nextRowIndex;
		row = buildRowFromHeaders(headers, currentData);
		row = applyExistingJobStatus(row, emptyRowInfo, newRowIndex);
		response = await updateRow(
			activeToken,
			spreadsheetId,
			sheetName,
			newRowIndex,
			row
		);
	}
	if (response.status === 401 || response.status === 403) {
		await removeCachedToken(activeToken);
		const retryAuth = await requestAuthToken(true);
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
			response = await updateRow(
				activeToken,
				spreadsheetId,
				sheetName,
				newRowIndex,
				row
			);
		}
	}

	if (!response.ok) {
		setStatus(`Sheets API error ${response.status}.`, "error");
		return;
	}
	if (!existingRow) {
		const targetRowIndex = newRowIndex;
		const reusedRow = emptyRowInfo.emptyRows.includes(targetRowIndex);
		setStatus("Formatting in progress...", "");
		try {
			let templateSheetId = null;
			let targetSheetId = null;
			const needsTemplate =
				!reusedRow ||
				(await rowNeedsTemplateValidation(
					activeToken,
					spreadsheetId,
					sheetName,
					targetRowIndex
				));
			if (targetRowIndex && needsTemplate) {
				templateSheetId = await ensureTemplateSheet(
					activeToken,
					spreadsheetId
				);
				const bidderValidation = templateSheetId
					? await getCellDataValidation(
							activeToken,
							spreadsheetId,
							"__Upwork Template",
							"C2"
						)
					: null;
				if (
					templateSheetId &&
					bidder &&
					!validationAllowsValue(bidderValidation, bidder)
				) {
					templateSheetId = await ensureTemplateSheet(
						activeToken,
						spreadsheetId,
						true
					);
				}
				targetSheetId = await getSheetId(
					activeToken,
					spreadsheetId,
					sheetName
				);
				let appliedTemplate = false;
				if (templateSheetId && targetSheetId !== null) {
					appliedTemplate = await applyTemplateRowToRows(
						activeToken,
						spreadsheetId,
						templateSheetId,
						targetSheetId,
						[targetRowIndex],
						DEFAULT_HEADER_COUNT
					);
				}
				if (!appliedTemplate) {
					await applyRowTemplatesInSheet(
						activeToken,
						spreadsheetId,
						sheetName,
						2,
						[targetRowIndex],
						DEFAULT_HEADER_COUNT
					);
				}
				await clearRowsValues(
					activeToken,
					spreadsheetId,
					sheetName,
					targetRowIndex,
					targetRowIndex,
					DEFAULT_HEADER_COUNT
				);
				await updateRow(
					activeToken,
					spreadsheetId,
					sheetName,
					targetRowIndex,
					row
				);
			}
			const clearUntil = targetRowIndex ? targetRowIndex + 20 : 200;
			await clearBodyBold(
				activeToken,
				spreadsheetId,
				sheetName,
				DEFAULT_HEADER_COUNT,
				clearUntil
			);
			setStatus("Row added to Google Sheets.", "success");
		} catch (error) {
			setStatus("Row added; formatting skipped.", "warn");
		}
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
	setBidder(storedBidder || "");
});
setFields(null);
setButtons();

const refreshMode = async () => {
	const tab = await queryActiveTab();
	currentTabUrl = tab?.url || "";
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
	if (mode === "connects") {
		if (!tab || !tab.id || !tab.url) {
			setStatus("No active tab detected.", "warn");
			return;
		}

		setStatus("Scanning connects history...", "");
		const scan = await executeConnectsHistoryScan(tab.id);
		if (!scan.ok) {
			setStatus(scan.error || "Unable to scan connects history.", "error");
			return;
		}
		const entries = Array.isArray(scan.entries) ? scan.entries : [];
		const jobIds = new Set(
			entries.map((entry) => String(entry?.jobId || "").trim()).filter(Boolean)
		);
		if (!entries.length) {
			setStatus("No connects history entries found.", "warn");
			return;
		}
		setStatus(
			`Found ${jobIds.size || entries.length} job(s). Click "Update connects history" to save.`,
			"success"
		);
		return;
	}
	if (!tab || !tab.id) {
		setStatus("No active tab detected.", "warn");
		return;
	}

	setStatus("Reading current tab...", "");
	const result = await executeExtraction(tab.id);
	if (!result.ok) {
		if (result.error === "No job activity found on this page.") {
			setStatus("Open an Upwork job details page to begin.", "");
		} else {
			setStatus(result.error || "Unable to read this page.", "error");
		}
		return;
	}

	currentData = result.data;
	currentData.bidder = bidder;
	setFields(currentData);
	setButtons();
	setStatus("Captured job activity.", "success");
};

refreshMode();
