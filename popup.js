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
const openTemplateButton = document.getElementById("open-template");
const sendSheetsButton = document.getElementById("send-sheets");
const statusEl = document.getElementById("status");
const bidderInput = document.getElementById("bidder-input");

const PREPARE_SHEET_CONFIRMATION = "umar";

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

const verifyPrepareSheetPassword = () => {
	const entered = window.prompt("Enter password to prepare the sheet:");
	if (entered === null) {
		return false;
	}
	return entered === PREPARE_SHEET_CONFIRMATION;
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
			const auth = await getAuthToken(true);
			if (!auth.ok) {
				setStatus(auth.error || "Google authorization failed.", "error");
				return;
			}

		const headers = await getSheetHeaders(auth.token, spreadsheetId, sheetName);
		const idMap = await getProposalIdRowMap(
			auth.token,
			spreadsheetId,
			sheetName,
			headers
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
			if (!entry?.jobId) {
				continue;
			}
			const key = String(entry.jobId).trim();
			if (!totals.has(key)) {
				totals.set(key, {
					jobId: key,
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
			setStatus("No connects history entries with job IDs found.", "warn");
			return;
		}

		setStatus(`Found ${totals.size} job(s). Updating sheet...`, "");
		const auth = await getAuthToken(true);
		if (!auth.ok) {
			setStatus(auth.error || "Google authorization failed.", "error");
			return;
		}

		const connectsMap = await getConnectsRowMap(
			auth.token,
			spreadsheetId,
			sheetName
		);
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
		const emptyRowInfo = await getEmptyRowIndexes(
			auth.token,
			spreadsheetId,
			sheetName,
			headers
		);
		if (!emptyRowInfo) {
			setStatus("Unable to read sheet rows.", "error");
			return;
		}
		const emptyRowQueue = [...emptyRowInfo.emptyRows];
		let nextRowIndex = emptyRowInfo.nextRowIndex;
		const newRowIndexes = [];
		for (const entry of totals.values()) {
			const existing = connectsMap.map.get(entry.jobId);
			if (existing) {
				const parsedSpent = parseConnectsCellValue(existing.connectsSpent);
				const parsedRefund = parseConnectsCellValue(existing.connectsRefund);
				const parsedBoostedSpent = parseConnectsCellValue(
					existing.boostedConnectsSpent
				);
				const parsedBoostedRefund = parseConnectsCellValue(
					existing.boostedConnectsRefund
				);
				const nextSpent = parsedSpent + entry.connectsSpent;
				const nextRefund = parsedRefund + entry.connectsRefund;
				const nextBoostedSpent =
					parsedBoostedSpent + entry.boostedConnectsSpent;
				const nextBoostedRefund =
					parsedBoostedRefund + entry.boostedConnectsRefund;
				const nextSpentValue = formatConnectsValue(nextSpent, "spent");
				const nextRefundValue = formatConnectsValue(nextRefund, "refund");
				const nextBoostedSpentValue = formatConnectsValue(
					nextBoostedSpent,
					"spent"
				);
				const nextBoostedRefundValue = formatConnectsValue(
					nextBoostedRefund,
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
		if (newRowIndexes.length) {
			const templateSheetId = await ensureTemplateSheet(
				auth.token,
				spreadsheetId
			);
			const targetSheetId = await getSheetId(
				auth.token,
				spreadsheetId,
				sheetName
			);
			let appliedTemplate = false;
			if (templateSheetId && targetSheetId !== null) {
				appliedTemplate = await applyTemplateRowToRows(
					auth.token,
					spreadsheetId,
					templateSheetId,
					targetSheetId,
					newRowIndexes,
					DEFAULT_HEADER_COUNT
				);
			}
			if (!appliedTemplate) {
				await applyRowTemplatesInSheet(
					auth.token,
					spreadsheetId,
					sheetName,
					2,
					newRowIndexes,
					DEFAULT_HEADER_COUNT
				);
			}
			await clearRowsValues(
				auth.token,
				spreadsheetId,
				sheetName,
				newRowIndexes[0],
				newRowIndexes[newRowIndexes.length - 1],
				DEFAULT_HEADER_COUNT
			);
			if (newRowUpdates.length) {
				await batchUpdateValues(
					auth.token,
					spreadsheetId,
					newRowUpdates
				);
			}
		}

		setStatus("Connects history updated.", "success");
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
	if (!verifyPrepareSheetPassword()) {
		setStatus("Sheet preparation canceled or password incorrect.", "warn");
		return;
	}
	setStatus("Preparing sheet...", "");
	const auth = await getAuthToken(true);
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
	await ensureBidderDropdown(auth.token, spreadsheetId, sheetName);
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
	const auth = await getAuthToken(true);
	if (!auth.ok) {
		setStatus(auth.error || "Google authorization failed.", "error");
		return;
	}
	let activeToken = auth.token;

	let row = null;
	let newRowIndex = null;
	await ensureBidderDropdown(activeToken, spreadsheetId, sheetName);
	const headers = await getSheetHeaders(
		activeToken,
		spreadsheetId,
		sheetName
	);
	if (!headers) {
		setStatus("Unable to read sheet headers.", "error");
		return;
	}
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
		const emptyRowInfo = await getEmptyRowIndexes(
			activeToken,
			spreadsheetId,
			sheetName,
			headers
		);
		if (!emptyRowInfo) {
			setStatus("Unable to read sheet rows.", "error");
			return;
		}
		newRowIndex = emptyRowInfo.emptyRows[0] || emptyRowInfo.nextRowIndex;
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
		if (targetRowIndex) {
			const templateSheetId = await ensureTemplateSheet(
				activeToken,
				spreadsheetId
			);
			const targetSheetId = await getSheetId(
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
