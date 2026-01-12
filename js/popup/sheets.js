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
