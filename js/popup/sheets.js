const formatCsvField = (value) => {
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

const formatSheetId = (value) => {
	const text = String(value ?? "").trim();
	if (!text || text === "--") {
		return text;
	}
	if (!/^\d+$/.test(text)) {
		return text;
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
		data.jobStatus || "",
		data.invitesSent,
		data.interviewing,
		"",
		data.payment,
		data.country,
		data.jobCreatedSince,
		data.jobStatusAlt || "",
		data.jobId,
		data.proposals,
		data.proposalId || "--",
	data.connectsSpent || "",
	data.connectsRefund || "",
	data.boostedConnectsSpent || "",
	data.boostedConnectsRefund || "",
]
		.map(formatCsvField)
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
		data.jobStatus || "",
		data.invitesSent || "",
		data.interviewing || "",
		"",
		data.payment || "",
		data.country || "",
		data.jobCreatedSince || "",
		data.jobStatusAlt || "",
		formatSheetId(data.jobId),
		data.proposals || "",
		formatSheetId(data.proposalId || "--"),
		data.connectsSpent || "",
		data.connectsRefund || "",
		data.boostedConnectsSpent || "",
		data.boostedConnectsRefund || "",
	];
};

const buildConnectsRow = (data) => {
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
		"",
		"",
		"",
		"",
		"",
		"",
		"",
		formatSheetId(data.jobId),
		data.proposals || "",
		formatSheetId(data.proposalId || "--"),
		data.connectsSpent || "",
		data.connectsRefund || "",
		data.boostedConnectsSpent || "",
		data.boostedConnectsRefund || "",
	];
};

const formatRange = (name) => {
	const normalized = normalizeSheetName(name);
	const escaped = normalized.replace(/'/g, "''");
	return `'${escaped}'!A1`;
};

const DEFAULT_HEADERS = [
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
	"Invites",
	"Interview",
	"Remarks",
	"Payment",
	"Country",
	"Job Created Since",
	"Job ID",
	"Proposals",
	"Proposal ID",
	"Connects Spent",
	"Connects Refund",
	"Boosted Connects Spent",
	"Boosted Connects Refund",
	"Total Net Connects",
];
const DEFAULT_HEADER_COUNT = DEFAULT_HEADERS.length;
window.DEFAULT_HEADER_COUNT = DEFAULT_HEADER_COUNT;
const TEMPLATE_SHEET_TITLE = "__Upwork Template";

const TEMPLATE_SPREADSHEET_ID =
	"1sV7RYfXd4cNJdnK0ohTnPbxmSp36_dzndUVFjKE0dzQ";
window.TEMPLATE_SPREADSHEET_ID = TEMPLATE_SPREADSHEET_ID;

const getColumnLetter = (index) => {
	let column = "";
	let value = index;
	while (value > 0) {
		const remainder = (value - 1) % 26;
		column = String.fromCharCode(65 + remainder) + column;
		value = Math.floor((value - 1) / 26);
	}
	return column;
};

const getSheetHeaders = async (token, id, name) => {
	const escaped = normalizeSheetName(name).replace(/'/g, "''");
	const range = `'${escaped}'!1:1`;
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
	return data.values?.[0] || [];
};

const getSheetRows = async (
	token,
	id,
	name,
	columnCount,
	startRow = 2,
	endRow = null
) => {
	const escaped = normalizeSheetName(name).replace(/'/g, "''");
	const endColumn = getColumnLetter(columnCount || DEFAULT_HEADER_COUNT);
	const range = Number.isFinite(endRow)
		? `'${escaped}'!A${startRow}:${endColumn}${endRow}`
		: `'${escaped}'!A${startRow}:${endColumn}`;
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
	return data.values || [];
};

const getEmptyRowIndexes = async (token, id, name, headers) => {
	const resolvedHeaders = headers || (await getSheetHeaders(token, id, name));
	if (!resolvedHeaders) {
		return null;
	}
	const columnCount = resolvedHeaders.length || DEFAULT_HEADER_COUNT;
	const jobStatusIndexes = new Set();
	for (let i = 0; i < resolvedHeaders.length; i += 1) {
		if (resolvedHeaders[i] === "Job Status") {
			jobStatusIndexes.add(i);
		}
	}
	const jobStatusColumns = Array.from(jobStatusIndexes);
	const rows = await getSheetRows(token, id, name, columnCount, 2);
	if (!rows) {
		return null;
	}
	const emptyRows = [];
	const jobStatusValuesByRow = new Map();
	const isEmpty = (row) => {
		for (let i = 0; i < columnCount; i += 1) {
			if (jobStatusIndexes.has(i)) {
				continue;
			}
			const value = row?.[i];
			if (String(value || "").trim()) {
				return false;
			}
		}
		return true;
	};
	for (let i = 0; i < rows.length; i += 1) {
		if (isEmpty(rows[i])) {
			const rowIndex = i + 2;
			emptyRows.push(rowIndex);
			if (jobStatusColumns.length) {
				const values = jobStatusColumns.map(
					(columnIndex) => rows[i]?.[columnIndex] || ""
				);
				jobStatusValuesByRow.set(rowIndex, values);
			}
		}
	}
	return {
		emptyRows,
		nextRowIndex: rows.length ? rows.length + 2 : 2,
		headers: resolvedHeaders,
		columnCount,
		jobStatusIndexes: jobStatusColumns,
		jobStatusValuesByRow,
	};
};

const getColumnValues = async (
	token,
	id,
	name,
	columnIndex,
	startRow = 2,
	options = {}
) => {
	const escaped = normalizeSheetName(name).replace(/'/g, "''");
	const letter = getColumnLetter(columnIndex + 1);
	const range = `'${escaped}'!${letter}${startRow}:${letter}`;
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(range)}`
	);
	if (options.valueRenderOption) {
		url.searchParams.set("valueRenderOption", options.valueRenderOption);
	}
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	const data = await response.json();
	return data.values || [];
};

const applyRowValidationsToRows = async (
	token,
	id,
	name,
	rowIndexes,
	headers,
	sourceSheetId = null
) => {
	if (!rowIndexes || !rowIndexes.length) {
		return false;
	}
	const resolvedHeaders = headers || (await getSheetHeaders(token, id, name));
	if (!resolvedHeaders) {
		return false;
	}
	const destinationSheetId = await getSheetId(token, id, name);
	if (destinationSheetId === null) {
		return false;
	}
	const effectiveSourceSheetId =
		sourceSheetId === null ? destinationSheetId : sourceSheetId;
	const bidderIndex = getHeaderIndex(resolvedHeaders, "Bidder");
	const jobStatusIndexes = [];
	for (let i = 0; i < resolvedHeaders.length; i += 1) {
		if (resolvedHeaders[i] === "Job Status") {
			jobStatusIndexes.push(i + 1);
		}
	}
	const columns = [];
	if (bidderIndex) {
		columns.push(bidderIndex);
	}
	jobStatusIndexes.forEach((index) => columns.push(index));
	if (!columns.length) {
		return true;
	}
	const requests = [];
	rowIndexes.forEach((rowIndex) => {
		columns.forEach((columnIndex) => {
			requests.push({
				copyPaste: {
					source: {
						sheetId: effectiveSourceSheetId,
						startRowIndex: 1,
						endRowIndex: 2,
						startColumnIndex: columnIndex - 1,
						endColumnIndex: columnIndex,
					},
					destination: {
						sheetId: destinationSheetId,
						startRowIndex: rowIndex - 1,
						endRowIndex: rowIndex,
						startColumnIndex: columnIndex - 1,
						endColumnIndex: columnIndex,
					},
					pasteType: "PASTE_DATA_VALIDATION",
					pasteOrientation: "NORMAL",
				},
			});
		});
	});
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
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

const removeBlankHeaderColumn = async (
	token,
	id,
	name,
	columnIndex,
	headerValueToRight
) => {
	const headers = await getSheetHeaders(token, id, name);
	if (!headers) {
		return false;
	}
	const headerValue = String(headers[columnIndex] || "").trim();
	if (headerValue) {
		return false;
	}
	if (headerValueToRight) {
		const rightValue = String(headers[columnIndex + 1] || "").trim();
		if (rightValue !== headerValueToRight) {
			return false;
		}
	}
	const values = await getColumnValues(token, id, name, columnIndex, 2);
	if (!values) {
		return false;
	}
	const hasData = values.some(
		(row) => String(row?.[0] || "").trim() !== ""
	);
	if (hasData) {
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
	const requests = [
		{
			deleteDimension: {
				range: {
					sheetId,
					dimension: "COLUMNS",
					startIndex: columnIndex,
					endIndex: columnIndex + 1,
				},
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

const getHeaderIndex = (headers, name, occurrence = 1) => {
	let count = 0;
	for (let i = 0; i < headers.length; i += 1) {
		if (headers[i] === name) {
			count += 1;
			if (count === occurrence) {
				return i + 1;
			}
		}
	}
	return null;
};

const normalizeJobName = (value) => String(value || "").trim();

const buildRowFromHeaders = (headers, data) => {
	let jobStatusCount = 0;
	const safeName = String(data.name || "").replace(/"/g, '""');
	const safeLink = String(data.link || "").replace(/"/g, '""');
	const nameCell =
		safeName && safeLink
			? `=HYPERLINK("${safeLink}","${safeName}")`
			: data.name || "";

	return headers.map((header) => {
		switch (header) {
			case "Date":
				return formatSheetDate(data.date);
			case "Job Name":
				return nameCell;
			case "Bidder":
				return data.bidder || "";
			case "Job Status":
				jobStatusCount += 1;
				if (jobStatusCount === 2) {
					return data.jobStatusAlt || data.jobStatus || "";
				}
				return data.jobStatus || "";
			case "Invites":
				return data.invitesSent || "";
			case "Interview":
				return data.interviewing || "";
			case "Payment":
				return data.payment || "";
			case "Country":
				return data.country || "";
			case "Job Created Since":
				return data.jobCreatedSince || "";
			case "Job ID":
				return formatSheetId(data.jobId);
			case "Proposals":
				return data.proposals || "";
			case "Proposal ID":
				return formatSheetId(data.proposalId || "--");
			case "Connects Spent":
				return data.connectsSpent || "";
			case "Connects Refund":
				return data.connectsRefund || "";
			case "Boosted Connects Spent":
				return data.boostedConnectsSpent || "";
			case "Boosted Connects Refund":
				return data.boostedConnectsRefund || "";
			default:
				return "";
		}
	});
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
	url.searchParams.set("insertDataOption", "OVERWRITE");

	return fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ values: [row] }),
	});
};

const findRowByJobId = async (token, id, name, jobId, headers) => {
	if (!jobId) {
		return null;
	}
	const resolvedHeaders = headers || (await getSheetHeaders(token, id, name));
	if (!resolvedHeaders) {
		return null;
	}
	const jobIdIndex = getHeaderIndex(resolvedHeaders, "Job ID");
	if (!jobIdIndex) {
		return null;
	}
	const column = getColumnLetter(jobIdIndex);
	const range = `'${normalizeSheetName(name).replace(/'/g, "''")}'!${column}:${column}`;
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

const getProposalIdRowMap = async (token, id, name, headers) => {
	const resolvedHeaders = headers || (await getSheetHeaders(token, id, name));
	if (!resolvedHeaders) {
		return null;
	}
	const proposalIndex = getHeaderIndex(resolvedHeaders, "Proposal ID");
	if (!proposalIndex) {
		return null;
	}
	const column = getColumnLetter(proposalIndex);
	const range = `'${normalizeSheetName(name).replace(/'/g, "''")}'!${column}:${column}`;
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

const getJobNameRowMap = async (token, id, name, headers) => {
	const resolvedHeaders = headers || (await getSheetHeaders(token, id, name));
	if (!resolvedHeaders) {
		return null;
	}
	const jobNameIndex = getHeaderIndex(resolvedHeaders, "Job Name");
	if (!jobNameIndex) {
		return null;
	}
	const values = await getColumnValues(token, id, name, jobNameIndex - 1, 2);
	if (!values) {
		return null;
	}
	const map = new Map();
	for (let i = 0; i < values.length; i += 1) {
		const normalized = normalizeJobName(values[i]?.[0] || "");
		if (normalized && !map.has(normalized)) {
			map.set(normalized, i + 2);
		}
	}
	return map;
};

const getConnectsRowMap = async (token, id, name, headers) => {
	const resolvedHeaders = headers || (await getSheetHeaders(token, id, name));
	if (!resolvedHeaders) {
		return null;
	}
	const jobIdIndex = getHeaderIndex(resolvedHeaders, "Job ID");
	const connectsSpentIndex = getHeaderIndex(resolvedHeaders, "Connects Spent");
	const connectsRefundIndex = getHeaderIndex(resolvedHeaders, "Connects Refund");
	const boostedConnectsSpentIndex = getHeaderIndex(
		resolvedHeaders,
		"Boosted Connects Spent"
	);
	const boostedConnectsRefundIndex = getHeaderIndex(
		resolvedHeaders,
		"Boosted Connects Refund"
	);
	if (!jobIdIndex || !connectsSpentIndex || !connectsRefundIndex) {
		return null;
	}
	const sheet = normalizeSheetName(name).replace(/'/g, "''");
	const jobIdColumn = getColumnLetter(jobIdIndex);
	const columnEntries = [
		{ key: "connectsSpent", index: connectsSpentIndex },
		{ key: "connectsRefund", index: connectsRefundIndex },
		{ key: "boostedConnectsSpent", index: boostedConnectsSpentIndex },
		{ key: "boostedConnectsRefund", index: boostedConnectsRefundIndex },
	];
	const ranges = [`'${sheet}'!${jobIdColumn}:${jobIdColumn}`];
	const columnKeys = [];
	const columnLetters = {};
	for (const entry of columnEntries) {
		if (!entry.index) {
			continue;
		}
		const letter = getColumnLetter(entry.index);
		columnLetters[`${entry.key}Column`] = letter;
		columnKeys.push(entry.key);
		ranges.push(`'${sheet}'!${letter}:${letter}`);
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values:batchGet`
	);
	ranges.forEach((range) => url.searchParams.append("ranges", range));
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	const data = await response.json();
	const valueRanges = data.valueRanges || [];
	const jobIdValues = valueRanges[0]?.values || [];
	const columnValues = {};
	for (let i = 0; i < columnKeys.length; i += 1) {
		columnValues[columnKeys[i]] = valueRanges[i + 1]?.values || [];
	}
	const map = new Map();
	const rowsByIndex = new Map();
	const lengths = [jobIdValues.length];
	for (const values of Object.values(columnValues)) {
		lengths.push(values.length);
	}
	const rowCount = Math.max(...lengths);
	for (let i = 1; i < rowCount; i += 1) {
		const jobId = String(jobIdValues[i]?.[0] || "").trim();
		const connectsSpent = columnValues.connectsSpent?.[i]?.[0] || "";
		const connectsRefund = columnValues.connectsRefund?.[i]?.[0] || "";
		const boostedConnectsSpent =
			columnValues.boostedConnectsSpent?.[i]?.[0] || "";
		const boostedConnectsRefund =
			columnValues.boostedConnectsRefund?.[i]?.[0] || "";
		rowsByIndex.set(i + 1, {
			rowIndex: i + 1,
			connectsSpent,
			connectsRefund,
			boostedConnectsSpent,
			boostedConnectsRefund,
		});
		if (!jobId) {
			continue;
		}
		map.set(jobId, {
			rowIndex: i + 1,
			connectsSpent,
			connectsRefund,
			boostedConnectsSpent,
			boostedConnectsRefund,
		});
	}
	return {
		map,
		rowsByIndex,
		nextRowIndex: rowCount ? rowCount + 1 : 2,
		headers: resolvedHeaders,
		columnCount: resolvedHeaders.length,
		connectsSpentColumn: columnLetters.connectsSpentColumn,
		connectsRefundColumn: columnLetters.connectsRefundColumn,
		boostedConnectsSpentColumn: columnLetters.boostedConnectsSpentColumn,
		boostedConnectsRefundColumn: columnLetters.boostedConnectsRefundColumn,
	};
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

const getCellDataValidation = async (token, id, name, cell) => {
	if (!cell) {
		return null;
	}
	const normalized = normalizeSheetName(name);
	const escaped = normalized.replace(/'/g, "''");
	const range = `'${escaped}'!${cell}`;
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}`
	);
	url.searchParams.set("ranges", range);
	url.searchParams.set(
		"fields",
		"sheets.data.rowData.values.dataValidation"
	);
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	const data = await response.json();
	return (
		data?.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values?.[0]
			?.dataValidation || null
	);
};

const setReadCellsViewed = async (token, id, name, rowIndexes, columnIndex) => {
	if (!rowIndexes || !rowIndexes.length) {
		return false;
	}
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return false;
	}
	const startColumnIndex =
		typeof columnIndex === "number" ? columnIndex : 6;
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
				startColumnIndex,
				endColumnIndex: startColumnIndex + 1,
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
	const endColumn = getColumnLetter(row.length || 1);
	const range = `'${escaped}'!A${rowIndex}:${endColumn}${rowIndex}`;
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

const batchUpdateValues = async (token, id, ranges) => {
	if (!ranges.length) {
		return { ok: true, status: 200 };
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values:batchUpdate`
	);
	url.searchParams.set("valueInputOption", "USER_ENTERED");
	return fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ data: ranges }),
	});
};

const updateConnectsColumns = async (
	token,
	id,
	name,
	rowIndex,
	connectsSpent,
	connectsRefund
) => {
	const normalized = normalizeSheetName(name);
	const escaped = normalized.replace(/'/g, "''");
	const range = `'${escaped}'!X${rowIndex}:Y${rowIndex}`;
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
		body: JSON.stringify({
			values: [[connectsSpent || "", connectsRefund || ""]],
		}),
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
	const match = range.match(/!([A-Z]+)(\d+)(:([A-Z]+)(\d+))?/);
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
	return fetch(url.toString(), {
		method: "PUT",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ values: [DEFAULT_HEADERS] }),
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
	const purple = {
		red: 234 / 255,
		green: 209 / 255,
		blue: 220 / 255,
	};
	const red = {
		red: 244 / 255,
		green: 204 / 255,
		blue: 204 / 255,
	};
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

const getFirstSheetId = async (token, id) => {
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
	return data.sheets?.[0]?.properties?.sheetId ?? null;
};

const clearSheetValues = async (token, id, name, range) => {
	const escaped = normalizeSheetName(name).replace(/'/g, "''");
	const target = `'${escaped}'!${range}`;
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}/values/${encodeURIComponent(target)}:clear`
	);
	return fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({}),
	});
};

const ensureTemplateSheet = async (token, destinationId, forceRefresh = false) => {
	const existingTemplateId = await getSheetId(
		token,
		destinationId,
		TEMPLATE_SHEET_TITLE
	);
	if (existingTemplateId !== null && !forceRefresh) {
		return existingTemplateId;
	}
	if (existingTemplateId !== null && forceRefresh) {
		const deleteUrl = new URL(
			`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
				destinationId
			)}:batchUpdate`
		);
		const deleteResponse = await fetch(deleteUrl.toString(), {
			method: "POST",
			headers: {
				Authorization: `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: JSON.stringify({
				requests: [{ deleteSheet: { sheetId: existingTemplateId } }],
			}),
		});
		if (!deleteResponse.ok) {
			return null;
		}
	}
	const templateSheetId = await getFirstSheetId(
		token,
		TEMPLATE_SPREADSHEET_ID
	);
	if (templateSheetId === null) {
		return null;
	}
	const copyUrl = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			TEMPLATE_SPREADSHEET_ID
		)}/sheets/${encodeURIComponent(templateSheetId)}:copyTo`
	);
	const copyResponse = await fetch(copyUrl.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ destinationSpreadsheetId: destinationId }),
	});
	if (!copyResponse.ok) {
		return null;
	}
	const copyData = await copyResponse.json();
	const newSheetId = copyData.sheetId;
	if (!newSheetId) {
		return null;
	}
	const updateUrl = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			destinationId
		)}:batchUpdate`
	);
	const updateResponse = await fetch(updateUrl.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({
			requests: [
				{
					updateSheetProperties: {
						properties: {
							sheetId: newSheetId,
							title: TEMPLATE_SHEET_TITLE,
							hidden: true,
						},
						fields: "title,hidden",
					},
				},
			],
		}),
	});
	if (!updateResponse.ok) {
		return null;
	}
	return newSheetId;
};

const clearRowsValues = async (
	token,
	id,
	name,
	startRowIndex,
	endRowIndex,
	columnCount
) => {
	const endColumn = getColumnLetter(columnCount || DEFAULT_HEADER_COUNT);
	return clearSheetValues(
		token,
		id,
		name,
		`A${startRowIndex}:${endColumn}${endRowIndex}`
	);
};

const applyTemplateRowToRows = async (
	token,
	id,
	sourceSheetId,
	destinationSheetId,
	targetRowIndexes,
	columnCount
) => {
	if (!targetRowIndexes || !targetRowIndexes.length) {
		return false;
	}
	const endColumnIndex = columnCount || DEFAULT_HEADER_COUNT;
	const pasteTypes = [
		"PASTE_FORMAT",
		"PASTE_DATA_VALIDATION",
		"PASTE_CONDITIONAL_FORMATTING",
	];
	const requests = [];
	targetRowIndexes.forEach((rowIndex) => {
		pasteTypes.forEach((pasteType) => {
			requests.push({
				copyPaste: {
					source: {
						sheetId: sourceSheetId,
						startRowIndex: 1,
						endRowIndex: 2,
						startColumnIndex: 0,
						endColumnIndex,
					},
					destination: {
						sheetId: destinationSheetId,
						startRowIndex: rowIndex - 1,
						endRowIndex: rowIndex,
						startColumnIndex: 0,
						endColumnIndex,
					},
					pasteType,
					pasteOrientation: "NORMAL",
				},
			});
		});
	});
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
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

const applyRowTemplatesInSheet = async (
	token,
	id,
	name,
	sourceRowIndex,
	targetRowIndexes,
	columnCount
) => {
	if (!targetRowIndexes || !targetRowIndexes.length) {
		return false;
	}
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return false;
	}
	return applyTemplateRowToRows(
		token,
		id,
		sheetId,
		sheetId,
		targetRowIndexes,
		columnCount
	);
};

const applyTemplateFormatting = async (token, destinationId, destinationName) => {
	const templateSheetId = await getFirstSheetId(
		token,
		TEMPLATE_SPREADSHEET_ID
	);
	if (templateSheetId === null) {
		return { ok: false, error: "Template sheet not found." };
	}

	const existingSheetId = await getSheetId(
		token,
		destinationId,
		destinationName
	);
	const copyUrl = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			TEMPLATE_SPREADSHEET_ID
		)}/sheets/${encodeURIComponent(templateSheetId)}:copyTo`
	);
	const copyResponse = await fetch(copyUrl.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ destinationSpreadsheetId: destinationId }),
	});
	if (!copyResponse.ok) {
		return { ok: false, status: copyResponse.status };
	}
	const copyData = await copyResponse.json();
	const newSheetId = copyData.sheetId;
	if (!newSheetId) {
		return { ok: false, error: "Template copy failed." };
	}

	if (!existingSheetId) {
		const requests = [
			{
				updateSheetProperties: {
					properties: {
						sheetId: newSheetId,
						title: destinationName,
					},
					fields: "title",
				},
			},
			{
				deleteDimension: {
					range: {
						sheetId: newSheetId,
						dimension: "COLUMNS",
						startIndex: 18,
						endIndex: 19,
					},
				},
			},
		];
		const updateUrl = new URL(
			`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
				destinationId
			)}:batchUpdate`
		);
		const updateResponse = await fetch(updateUrl.toString(), {
			method: "POST",
			headers: {
				Authorization: `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: JSON.stringify({ requests }),
		});
		if (!updateResponse.ok) {
			return { ok: false, status: updateResponse.status };
		}
		const lastColumn = getColumnLetter(DEFAULT_HEADER_COUNT);
		const clearResponse = await clearSheetValues(
			token,
			destinationId,
			destinationName,
			`A2:${lastColumn}`
		);
		if (!clearResponse.ok) {
			return { ok: false, status: clearResponse.status };
		}
		return { ok: true, createdSheet: true };
	}

	const ranges = [
		{
			source: {
				sheetId: newSheetId,
				startRowIndex: 0,
				endRowIndex: 1000,
				startColumnIndex: 0,
				endColumnIndex: 18,
			},
			destination: {
				sheetId: existingSheetId,
				startRowIndex: 0,
				endRowIndex: 1000,
				startColumnIndex: 0,
				endColumnIndex: 18,
			},
		},
		{
			source: {
				sheetId: newSheetId,
				startRowIndex: 0,
				endRowIndex: 1000,
				startColumnIndex: 19,
				endColumnIndex: 25,
			},
			destination: {
				sheetId: existingSheetId,
				startRowIndex: 0,
				endRowIndex: 1000,
				startColumnIndex: 18,
				endColumnIndex: 24,
			},
		},
	];
	const pasteTypes = [
		"PASTE_FORMAT",
		"PASTE_DATA_VALIDATION",
		"PASTE_CONDITIONAL_FORMATTING",
	];
	const requests = [];
	ranges.forEach((rangePair) => {
		pasteTypes.forEach((pasteType) => {
			requests.push({
				copyPaste: {
					source: rangePair.source,
					destination: rangePair.destination,
					pasteType,
					pasteOrientation: "NORMAL",
				},
			});
		});
	});
	requests.push({
		deleteSheet: { sheetId: newSheetId },
	});
	const updateUrl = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			destinationId
		)}:batchUpdate`
	);
	const updateResponse = await fetch(updateUrl.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ requests }),
	});
	if (!updateResponse.ok) {
		const cleanupUrl = new URL(
			`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
				destinationId
			)}:batchUpdate`
		);
		await fetch(cleanupUrl.toString(), {
			method: "POST",
			headers: {
				Authorization: `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: JSON.stringify({
				requests: [{ deleteSheet: { sheetId: newSheetId } }],
			}),
		});
		return { ok: false, status: updateResponse.status };
	}

	return { ok: true };
};

const copyTemplateSheet = async (token, destinationId, destinationName) => {
	const templateSheetId = await getFirstSheetId(
		token,
		TEMPLATE_SPREADSHEET_ID
	);
	if (templateSheetId === null) {
		return { ok: false, error: "Template sheet not found." };
	}

	const existingSheetId = await getSheetId(token, destinationId, destinationName);

	const copyUrl = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			TEMPLATE_SPREADSHEET_ID
		)}/sheets/${encodeURIComponent(templateSheetId)}:copyTo`
	);
	const copyResponse = await fetch(copyUrl.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ destinationSpreadsheetId: destinationId }),
	});
	if (!copyResponse.ok) {
		return { ok: false, status: copyResponse.status };
	}
	const copyData = await copyResponse.json();
	const newSheetId = copyData.sheetId;
	if (!newSheetId) {
		return { ok: false, error: "Template copy failed." };
	}

	const requests = [];
	if (existingSheetId && existingSheetId !== newSheetId) {
		requests.push({
			deleteSheet: { sheetId: existingSheetId },
		});
	}
	requests.push({
		updateSheetProperties: {
			properties: {
				sheetId: newSheetId,
				title: destinationName,
			},
			fields: "title",
		},
	});

	if (requests.length) {
		const updateUrl = new URL(
			`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
				destinationId
			)}:batchUpdate`
		);
		const updateResponse = await fetch(updateUrl.toString(), {
			method: "POST",
			headers: {
				Authorization: `Bearer ${token}`,
				"Content-Type": "application/json",
			},
			body: JSON.stringify({ requests }),
		});
		if (!updateResponse.ok) {
			return { ok: false, status: updateResponse.status };
		}
	}

	const clearResponse = await clearSheetValues(
		token,
		destinationId,
		destinationName,
		"A2:Y"
	);
	if (!clearResponse.ok) {
		return { ok: false, status: clearResponse.status };
	}

	return { ok: true };
};

const ensureBidderDropdown = async (token, id, name) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return;
	}
	const templateSheetId = await ensureTemplateSheet(token, id);
	if (!templateSheetId) {
		return;
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
	const requests = [
		{
			copyPaste: {
				source: {
					sheetId: templateSheetId,
					startRowIndex: 1,
					endRowIndex: 2,
					startColumnIndex: 2,
					endColumnIndex: 3,
				},
				destination: {
					sheetId,
					startRowIndex: 1,
					endRowIndex: 1000,
					startColumnIndex: 2,
					endColumnIndex: 3,
				},
				pasteType: "PASTE_DATA_VALIDATION",
				pasteOrientation: "NORMAL",
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

const ensureJobStatusDropdown = async (token, id, name) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return;
	}
	const existingValidation = await getCellDataValidation(token, id, name, "L2");
	if (existingValidation) {
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
					endRowIndex: 1000,
					startColumnIndex: 11,
					endColumnIndex: 12,
				},
				cell: {
					dataValidation: {
						condition: {
							type: "ONE_OF_LIST",
							values: [
								{ userEnteredValue: "Open(No Hire)" },
								{ userEnteredValue: "Declined" },
								{ userEnteredValue: "Rejected" },
								{ userEnteredValue: "Hired Us" },
								{ userEnteredValue: "Hired Someone" },
								{ userEnteredValue: "Closed with no Hire" },
								{ userEnteredValue: "Open" },
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

const ensureJobStatusColors = async (token, id, name) => {
	const sheetId = await getSheetId(token, id, name);
	if (sheetId === null) {
		return;
	}
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(
			id
		)}:batchUpdate`
	);
	const range = {
		sheetId,
		startRowIndex: 1,
		endRowIndex: 1000,
		startColumnIndex: 11,
		endColumnIndex: 12,
	};
	const rules = [
		{
			text: "Open(No Hire)",
			backgroundColor: { red: 0.99, green: 0.91, blue: 0.7 },
			textColor: { red: 0.36, green: 0.27, blue: 0 },
		},
		{
			text: "Declined",
			backgroundColor: { red: 0.96, green: 0.78, blue: 0.76 },
			textColor: { red: 0.67, green: 0.1, blue: 0.09 },
		},
		{
			text: "Rejected",
			backgroundColor: { red: 0.8, green: 0.0, blue: 0.0 },
			textColor: { red: 1, green: 1, blue: 1 },
		},
		{
			text: "Hired Us",
			backgroundColor: { red: 0.24, green: 0.52, blue: 0.36 },
			textColor: { red: 1, green: 1, blue: 1 },
		},
		{
			text: "Hired Someone",
			backgroundColor: { red: 0.7, green: 0.84, blue: 0.96 },
			textColor: { red: 0.12, green: 0.33, blue: 0.69 },
		},
		{
			text: "Closed with no Hire",
			backgroundColor: { red: 0.84, green: 0.7, blue: 0.96 },
			textColor: { red: 0.36, green: 0.2, blue: 0.6 },
		},
		{
			text: "Open",
			backgroundColor: { red: 0.9, green: 0.9, blue: 0.9 },
			textColor: { red: 0.2, green: 0.2, blue: 0.2 },
		},
	];
	const requests = rules.map((rule) => ({
		addConditionalFormatRule: {
			rule: {
				ranges: [range],
				booleanRule: {
					condition: {
						type: "TEXT_EQ",
						values: [{ userEnteredValue: rule.text }],
					},
					format: {
						backgroundColor: rule.backgroundColor,
						textFormat: {
							foregroundColor: rule.textColor,
							bold: true,
						},
					},
				},
			},
			index: 0,
		},
	}));
	await fetch(url.toString(), {
		method: "POST",
		headers: {
			Authorization: `Bearer ${token}`,
			"Content-Type": "application/json",
		},
		body: JSON.stringify({ requests }),
	});
};
