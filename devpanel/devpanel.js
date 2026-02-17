const fixtureOutput = document.getElementById("fixture-output");
const statusEl = document.getElementById("dev-status");
const seedInput = document.getElementById("mock-seed");
const jobsInput = document.getElementById("mock-jobs");
const proposalsInput = document.getElementById("mock-proposals");
const connectsInput = document.getElementById("mock-connects");
const generateButton = document.getElementById("generate-mock");
const copyButton = document.getElementById("copy-mock");
const jsonOutput = document.getElementById("mock-json");
const previewEl = document.getElementById("mock-preview");
const sheetIdInput = document.getElementById("sheet-id");
const sheetNameInput = document.getElementById("sheet-name");
const sheetBidderInput = document.getElementById("sheet-bidder");
const viewedRateInput = document.getElementById("sheet-viewed-rate");
const connectsRefundInput = document.getElementById("sheet-connects-refund");
const connectsBoostedInput = document.getElementById("sheet-connects-boosted");
const saveSheetButton = document.getElementById("sheet-save");
const openSheetButton = document.getElementById("sheet-open");
const prepareSheetButton = document.getElementById("sheet-prepare");
const bulkAddButton = document.getElementById("bulk-add");
const bulkEditButton = document.getElementById("bulk-edit");
const bulkViewedButton = document.getElementById("bulk-viewed");
const bulkConnectsButton = document.getElementById("bulk-connects");
const fullTestButton = document.getElementById("full-test");
const sheetLogEl = document.getElementById("sheet-log");

let mockPayload = null;
let sheetId = "";
let sheetName = "Sheet1";

const setStatus = (message, tone = "") => {
	statusEl.textContent = message;
	statusEl.className = `status ${tone}`.trim();
};

const setSheetLog = (title, lines = []) => {
	if (!sheetLogEl) {
		return;
	}
	const body = lines.length ? `\n\n${lines.join("\n")}` : "";
	sheetLogEl.textContent = `${title}${body}`.trim();
};

const buildResult = (ok, tone, message, extra = {}) => ({
	ok,
	tone,
	message,
	...extra,
});

const parseSavedUrl = (html) => {
	const match = html.match(/saved from url=\(\d+\)([^\s>]+)/i);
	if (!match) {
		return "";
	}
	return match[1].replace(/-->/g, "").trim();
};

const renderFixtureOutput = (title, summary, result) => {
	const header = `${title}\n${summary}`.trim();
	fixtureOutput.textContent = `${header}\n\n${JSON.stringify(result, null, 2)}`;
};

const summarizeConnects = (entries) => {
	return entries.reduce(
		(acc, entry) => {
			acc.count += 1;
			acc.spent += Number(entry.connectsSpent || 0);
			acc.refund += Number(entry.connectsRefund || 0);
			acc.boostedSpent += Number(entry.boostedConnectsSpent || 0);
			acc.boostedRefund += Number(entry.boostedConnectsRefund || 0);
			return acc;
		},
		{ count: 0, spent: 0, refund: 0, boostedSpent: 0, boostedRefund: 0 }
	);
};

const runFixture = async (fixture) => {
	if (!fixtureOutput) {
		return;
	}
	setStatus(`Loading ${fixture} fixture...`, "");
	const fixturePath = `test_data/${fixture}`;
	let fixtureUrl = "";
	let html = "";
	try {
		fixtureUrl = chrome.runtime.getURL(fixturePath);
		const response = await fetch(fixtureUrl);
		html = await response.text();
	} catch (error) {
		renderFixtureOutput(
			"Fixture load failed",
			String(error.message || error),
			{}
		);
		setStatus("Fixture load failed.", "error");
		return;
	}

	const doc = new DOMParser().parseFromString(html, "text/html");
	const savedUrl = parseSavedUrl(html) || fixtureUrl;

	try {
		if (fixture === "job_page.html") {
			const result = window.extractUpworkJobDataFromDocument(doc, savedUrl);
			const data = result?.data || {};
			const summary = [
				`Name: ${data.name || "-"}`,
				`Job ID: ${data.jobId || "-"}`,
				`Proposals: ${data.proposals || "-"}`,
				`Invites: ${data.invitesSent || "-"}`,
				`Interviewing: ${data.interviewing || "-"}`,
				`Last viewed: ${data.lastViewed || "-"}`,
				`Unanswered invites: ${data.unansweredInvites || "-"}`,
				`Payment: ${data.payment || "-"}`,
				`Country: ${data.country || "-"}`,
				`Posted: ${data.jobCreatedSince || "-"}`,
				`Proposal ID: ${data.proposalId || "-"}`,
			].join("\n");
			renderFixtureOutput("Job activity", summary, result);
			setStatus("Job fixture parsed.", "success");
			return;
		}

		if (fixture === "Proposals.html") {
			const result = window.extractUpworkViewedProposalsFromDocument(doc);
			const proposals = result?.proposals || [];
			const viewed = proposals.filter((item) => item.viewed).length;
			const sample = proposals
				.slice(0, 5)
				.map(
					(item) =>
						`${item.proposalId}: ${item.viewed ? "viewed" : "unviewed"}`
							+ ` - ${item.title}`
				)
				.join("\n");
			const summary = [
				`Total proposals: ${proposals.length}`,
				`Viewed: ${viewed}`,
				"Sample:",
				sample || "-",
			].join("\n");
			renderFixtureOutput("Proposals", summary, result);
			setStatus("Proposals fixture parsed.", "success");
			return;
		}

		if (fixture === "ConnectsHistory.html") {
			let origin = "https://www.upwork.com";
			try {
				origin = new URL(savedUrl).origin;
			} catch (error) {
				origin = "https://www.upwork.com";
			}
			const result = window.extractUpworkConnectsHistoryFromDocument(
				doc,
				origin
			);
			const entries = result?.entries || [];
			const totals = summarizeConnects(entries);
			const sample = entries
				.slice(0, 5)
				.map(
					(entry) =>
						`${entry.jobId}: ${entry.isBoosted ? "boosted" : "standard"}`
							+ ` - ${entry.connectsSpent || entry.connectsRefund || 0}`
				)
				.join("\n");
			const summary = [
				`Total rows: ${totals.count}`,
				`Spent: ${totals.spent}`,
				`Refund: ${totals.refund}`,
				`Boosted spent: ${totals.boostedSpent}`,
				`Boosted refund: ${totals.boostedRefund}`,
				"Sample:",
				sample || "-",
			].join("\n");
			renderFixtureOutput("Connects history", summary, result);
			setStatus("Connects fixture parsed.", "success");
			return;
		}

		throw new Error("Unknown fixture.");
	} catch (error) {
		renderFixtureOutput(
			"Fixture parse failed",
			String(error.message || error),
			{}
		);
		setStatus("Fixture parse failed.", "error");
	}
};

const seedFromString = (value) => {
	let hash = 0;
	const text = String(value || "");
	for (let i = 0; i < text.length; i += 1) {
		hash = (hash * 31 + text.charCodeAt(i)) | 0;
	}
	return hash >>> 0;
};

const createRng = (seed) => {
	let state = seedFromString(seed);
	return () => {
		state |= 0;
		state = (state + 0x6d2b79f5) | 0;
		let t = Math.imul(state ^ (state >>> 15), 1 | state);
		t ^= t + Math.imul(t ^ (t >>> 7), 61 | t);
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
};

const randomInt = (rng, min, max) =>
	Math.floor(rng() * (max - min + 1)) + min;

const pick = (rng, list) => list[Math.floor(rng() * list.length)];

const slugify = (value) =>
	String(value || "")
		.toLowerCase()
		.replace(/[^a-z0-9]+/g, "-")
		.replace(/(^-|-$)/g, "");

const randomId = (rng, length) => {
	let output = "";
	for (let i = 0; i < length; i += 1) {
		output += String(randomInt(rng, 0, 9));
	}
	return output;
};

const randomProposalId = (rng) => `200${randomId(rng, 16)}`;

const randomDate = (rng) => {
	const daysBack = randomInt(rng, 0, 30);
	const date = new Date(Date.now() - daysBack * 86400000);
	return date.toLocaleDateString("en-US", {
		month: "short",
		day: "numeric",
		year: "numeric",
	});
};

const MOCK_TITLES = [
	"AI workflow automation",
	"Shopify storefront cleanup",
	"Mobile app MVP build",
	"Data dashboard redesign",
	"Salesforce integration",
	"Marketing analytics pipeline",
	"Node.js API revamp",
	"Customer support chatbot",
	"Fintech reporting upgrade",
	"Risk scoring service",
];

const MOCK_COUNTRIES = ["United States", "Canada", "Saudi Arabia", "UK", "UAE"];

const MOCK_CREATED_SINCE = [
	"1 day ago",
	"2 days ago",
	"1 week ago",
	"2 weeks ago",
	"1 month ago",
];

const generateMockData = () => {
	const rng = createRng(seedInput.value);
	const jobsCount = Math.max(1, Number(jobsInput.value) || 1);
	const proposalsCount = Math.max(1, Number(proposalsInput.value) || 1);
	const connectsCount = Math.max(1, Number(connectsInput.value) || 1);
	const refundRate = getConnectsRefundRate() / 100;
	const boostedRate = getConnectsBoostedRate() / 100;

	const jobs = Array.from({ length: jobsCount }, () => {
		const name = pick(rng, MOCK_TITLES);
		const jobId = randomId(rng, 18);
		const proposalsMin = randomInt(rng, 2, 5);
		const proposalsMax = proposalsMin + randomInt(rng, 4, 12);
		const invitesSent = randomInt(rng, 0, 8);
		const interviewing = randomInt(rng, 0, 3);
		const unansweredInvites = Math.max(0, invitesSent - randomInt(rng, 0, 3));
		return {
			name,
			link: `https://www.upwork.com/jobs/${slugify(name)}_~0${jobId}/`,
			jobId,
			date: randomDate(rng),
			payment: rng() > 0.2 ? "Verified" : "Not verified",
			country: pick(rng, MOCK_COUNTRIES),
			jobCreatedSince: pick(rng, MOCK_CREATED_SINCE),
			proposals: `${proposalsMin} to ${proposalsMax}`,
			invitesSent: String(invitesSent),
			interviewing: String(interviewing),
			lastViewed: `${randomInt(rng, 1, 7)} days ago`,
			unansweredInvites: String(unansweredInvites),
			proposalId: randomProposalId(rng),
		};
	});

	const proposals = Array.from({ length: proposalsCount }, () => {
		const job = pick(rng, jobs);
		const viewed = rng() > 0.6;
		return {
			proposalId: randomProposalId(rng),
			title: job.name,
			viewed,
		};
	});

	const connects = Array.from({ length: connectsCount }, () => {
		const job = pick(rng, jobs);
		const isBoosted = rng() < boostedRate;
		const isRefund = rng() < refundRate;
		const magnitude = randomInt(rng, 2, 24);
		const spent = isRefund ? 0 : magnitude;
		const refund = isRefund ? magnitude : 0;
		return {
			jobId: job.jobId,
			title: job.name,
			link: job.link,
			date: randomDate(rng),
			isBoosted,
			connectsSpent: isBoosted ? 0 : spent,
			connectsRefund: isBoosted ? 0 : refund,
			boostedConnectsSpent: isBoosted ? spent : 0,
			boostedConnectsRefund: isBoosted ? refund : 0,
		};
	});

	const summary = summarizeConnects(connects);
	const payload = {
		summary: {
			jobs: jobs.length,
			proposals: proposals.length,
			proposalsViewed: proposals.filter((item) => item.viewed).length,
			connects: connects.length,
			connectsSpent: summary.spent,
			connectsRefund: summary.refund,
			boostedConnectsSpent: summary.boostedSpent,
			boostedConnectsRefund: summary.boostedRefund,
		},
		jobs,
		proposals,
		connects,
	};

	mockPayload = payload;
	jsonOutput.value = JSON.stringify(payload, null, 2);
	previewEl.innerHTML = "";
	previewEl.appendChild(
		createPreviewCard("Summary", [
			{ label: "Jobs", value: payload.summary.jobs },
			{ label: "Proposals", value: payload.summary.proposals },
			{ label: "Viewed", value: payload.summary.proposalsViewed },
			{ label: "Connects", value: payload.summary.connects },
		])
	);
	previewEl.appendChild(
		createPreviewCard(
			"Jobs",
			jobs.slice(0, 4).map((job) => ({
				label: job.jobId,
				value: job.name,
			}))
		)
	);
	previewEl.appendChild(
		createPreviewCard(
			"Proposals",
			proposals.slice(0, 4).map((proposal) => ({
				label: proposal.proposalId,
				value: proposal.viewed ? "viewed" : "unviewed",
			}))
		)
	);
	previewEl.appendChild(
		createPreviewCard(
			"Connects",
			connects.slice(0, 4).map((entry) => ({
				label: entry.jobId,
				value: entry.isBoosted ? "boosted" : "standard",
			}))
		)
	);
	setStatus("Mock data generated.", "success");
};

const createPreviewCard = (title, items) => {
	const card = document.createElement("div");
	card.className = "preview-card";
	const heading = document.createElement("h3");
	heading.textContent = title;
	card.appendChild(heading);
	const list = document.createElement("ul");
	list.className = "preview-list";
	items.forEach((item) => {
		const row = document.createElement("li");
		const label = document.createElement("span");
		label.className = "label";
		label.textContent = String(item.label);
		const value = document.createElement("span");
		value.textContent = String(item.value);
		row.appendChild(label);
		row.appendChild(value);
		list.appendChild(row);
	});
	card.appendChild(list);
	return card;
};

const copyJson = async () => {
	const text = jsonOutput.value;
	if (!text) {
		setStatus("Nothing to copy yet.", "warn");
		return;
	}
	try {
		await navigator.clipboard.writeText(text);
		setStatus("Mock JSON copied.", "success");
	} catch (error) {
		setStatus("Clipboard access failed.", "error");
	}
};

const getMockPayload = () => {
	if (mockPayload) {
		return mockPayload;
	}
	const raw = jsonOutput?.value || "";
	if (!raw.trim()) {
		return null;
	}
	try {
		return JSON.parse(raw);
	} catch (error) {
		return null;
	}
};

const updateSheetStateFromInputs = () => {
	if (!sheetIdInput || !sheetNameInput) {
		return;
	}
	if (typeof extractSpreadsheetId === "function") {
		sheetId = extractSpreadsheetId(sheetIdInput.value) || "";
	} else {
		sheetId = String(sheetIdInput.value || "").trim();
	}
	if (typeof normalizeSheetName === "function") {
		sheetName = normalizeSheetName(sheetNameInput.value);
	} else {
		sheetName = String(sheetNameInput.value || "").trim() || "Sheet1";
	}
};

const getBidderValue = () => String(sheetBidderInput?.value || "").trim();

const getPercentValue = (input, fallback) => {
	const raw = Number(input?.value);
	if (!Number.isFinite(raw)) {
		return fallback;
	}
	return Math.min(100, Math.max(0, raw));
};

const getViewedRate = () => getPercentValue(viewedRateInput, 60);
const getConnectsRefundRate = () => getPercentValue(connectsRefundInput, 20);
const getConnectsBoostedRate = () =>
	getPercentValue(connectsBoostedInput, 30);

const getSheetRangePrefix = () => {
	const normalized =
		typeof normalizeSheetName === "function"
			? normalizeSheetName(sheetName)
			: sheetName || "Sheet1";
	const escaped = normalized.replace(/'/g, "''");
	return `'${escaped}'!`;
};

const isEmptyValue = (value) =>
	value === null || value === undefined || value === "";

const pushCellUpdate = (updates, rangePrefix, columnIndex, rowIndex, value) => {
	if (!columnIndex || isEmptyValue(value)) {
		return;
	}
	const letter = getColumnLetter(columnIndex);
	updates.push({
		range: `${rangePrefix}${letter}${rowIndex}`,
		values: [[value]],
	});
};

const normalizeCellValue = (value) => String(value ?? "").trim();

const stripLeadingApostrophe = (value) => {
	const text = normalizeCellValue(value);
	return text.startsWith("'") ? text.slice(1) : text;
};

const getRowValue = (row, columnIndex) =>
	row && columnIndex ? row[columnIndex - 1] ?? "" : "";

const readSheetRow = async (token, id, name, columnCount, rowIndex) => {
	const rows = await getSheetRows(token, id, name, columnCount, rowIndex, rowIndex);
	return rows?.[0] || null;
};

const getSpreadsheetMeta = async (token, id, fields) => {
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}`
	);
	if (fields) {
		url.searchParams.set("fields", fields);
	}
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	return response.json();
};

const getSheetGridData = async (token, id, range, fields) => {
	const url = new URL(
		`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(id)}`
	);
	url.searchParams.set("includeGridData", "true");
	if (range) {
		url.searchParams.append("ranges", range);
	}
	if (fields) {
		url.searchParams.set("fields", fields);
	}
	const response = await fetch(url.toString(), {
		headers: { Authorization: `Bearer ${token}` },
	});
	if (!response.ok) {
		return null;
	}
	return response.json();
};

const normalizeColor = (color) => {
	if (!color) {
		return null;
	}
	const round = (value) =>
		Number.isFinite(value) ? Math.round(value * 100) / 100 : null;
	return {
		red: round(color.red),
		green: round(color.green),
		blue: round(color.blue),
	};
};

const normalizeFormat = (format) => {
	if (!format) {
		return null;
	}
	return {
		backgroundColor: normalizeColor(format.backgroundColor),
		horizontalAlignment: format.horizontalAlignment || null,
		verticalAlignment: format.verticalAlignment || null,
		wrapStrategy: format.wrapStrategy || null,
		numberFormat: format.numberFormat
			? {
					type: format.numberFormat.type || null,
					pattern: format.numberFormat.pattern || null,
				}
			: null,
		textFormat: format.textFormat
			? {
					bold:
						typeof format.textFormat.bold === "boolean"
							? format.textFormat.bold
							: null,
					italic:
						typeof format.textFormat.italic === "boolean"
							? format.textFormat.italic
							: null,
					fontFamily: format.textFormat.fontFamily || null,
					fontSize: format.textFormat.fontSize || null,
					foregroundColor: normalizeColor(
						format.textFormat.foregroundColor
					),
				}
			: null,
	};
};

const formatsMatch = (current, template) => {
	if (!template) {
		return true;
	}
	const isObject = (value) =>
		value && typeof value === "object" && !Array.isArray(value);
	const compare = (cur, tmpl) => {
		if (!isObject(tmpl)) {
			return cur === tmpl;
		}
		return Object.keys(tmpl).every((key) => {
			const tmplValue = tmpl[key];
			if (tmplValue === null || tmplValue === undefined) {
				return true;
			}
			const curValue = cur ? cur[key] : null;
			if (isObject(tmplValue)) {
				return compare(curValue || {}, tmplValue);
			}
			return curValue === tmplValue;
		});
	};
	return compare(current || {}, template);
};

const getGridRow = (grid, rowIndex) =>
	grid?.rowData?.[rowIndex - 1] || null;

const getGridCell = (rowData, columnIndex) =>
	rowData?.values?.[columnIndex - 1] || null;

const isGreenish = (color) => {
	if (!color) {
		return false;
	}
	const green = color.green ?? 0;
	const red = color.red ?? 0;
	const blue = color.blue ?? 0;
	return green >= 0.85 && red >= 0.7 && blue >= 0.7;
};

const getHeaderLabel = (headers, columnIndex) =>
	headers?.[columnIndex - 1] || `Column ${getColumnLetter(columnIndex)}`;

const compareRowFormats = (currentRow, templateRow, columnCount, headers) => {
	const mismatches = [];
	for (let i = 1; i <= columnCount; i += 1) {
		const currentFormat = normalizeFormat(
			getGridCell(currentRow, i)?.userEnteredFormat
		);
		const templateFormat = normalizeFormat(
			getGridCell(templateRow, i)?.userEnteredFormat
		);
		if (!formatsMatch(currentFormat, templateFormat)) {
			mismatches.push(getHeaderLabel(headers, i));
		}
	}
	return mismatches;
};

const computeUpdatedJobById = (jobs, seed, jobId) => {
	const rng = createRng(seed);
	let result = null;
	for (const job of jobs) {
		const updated = buildUpdatedJob(job, rng);
		if (String(job?.jobId || "").trim() === jobId) {
			result = updated;
		}
	}
	return result;
};

const normalizeJobLabel = (value) =>
	typeof normalizeJobName === "function"
		? normalizeJobName(value)
		: String(value || "").trim();

const withTimeout = (promise, timeoutMs, errorMessage) =>
	Promise.race([
		promise,
		new Promise((_, reject) =>
			setTimeout(() => reject(new Error(errorMessage)), timeoutMs)
		),
	]);

const requestAuthToken = async (interactive, timeoutMs = 15000) => {
	if (typeof getAuthToken !== "function") {
		return { ok: false, error: "Google auth helper unavailable." };
	}
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

const openSpreadsheet = (id) => {
	const url = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(
		id
	)}/edit`;
	if (chrome?.tabs?.create) {
		chrome.tabs.create({ url });
		return;
	}
	window.open(url, "_blank", "noopener");
};

const loadSheetSettings = async () => {
	if (!sheetIdInput || !sheetNameInput) {
		return;
	}
	if (typeof getSheetsSettings !== "function") {
		return;
	}
	const settings = await getSheetsSettings();
	sheetId = settings.spreadsheetId || "";
	sheetName =
		typeof normalizeSheetName === "function"
			? normalizeSheetName(settings.sheetName)
			: settings.sheetName || "Sheet1";
	sheetIdInput.value = sheetId;
	sheetNameInput.value = sheetName;
};

const saveSheetSettings = async () => {
	updateSheetStateFromInputs();
	if (!sheetId) {
		setStatus("Enter a spreadsheet ID or URL first.", "warn");
		return;
	}
	if (typeof saveSheetsSettings !== "function") {
		setStatus("Sheet settings storage unavailable.", "warn");
		return;
	}
	const ok = await saveSheetsSettings(sheetId, sheetName);
	if (!ok) {
		setStatus("Unable to save sheet settings.", "error");
		return;
	}
	setStatus("Sheet settings saved.", "success");
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

const formatConnectsValue = (value, kind) => {
	const numeric = Number(value);
	if (!Number.isFinite(numeric) || numeric <= 0) {
		return "";
	}
	const prefix = kind === "refund" ? "+" : "-";
	return `${prefix}${Math.abs(numeric)}`;
};

const getDefaultHeaderCount = () => {
	if (typeof DEFAULT_HEADER_COUNT === "number") {
		return DEFAULT_HEADER_COUNT;
	}
	if (typeof window !== "undefined" && typeof window.DEFAULT_HEADER_COUNT === "number") {
		return window.DEFAULT_HEADER_COUNT;
	}
	return 26;
};

const getJobIdRowMap = async (token, id, name, headers) => {
	const resolvedHeaders = headers || (await getSheetHeaders(token, id, name));
	if (!resolvedHeaders) {
		return null;
	}
	const jobIdIndex = getHeaderIndex(resolvedHeaders, "Job ID");
	if (!jobIdIndex) {
		return null;
	}
	const values = await getColumnValues(token, id, name, jobIdIndex - 1, 2);
	if (!values) {
		return null;
	}
	const map = new Map();
	for (let i = 0; i < values.length; i += 1) {
		const jobId = String(values[i]?.[0] || "").trim();
		if (jobId) {
			map.set(jobId, i + 2);
		}
	}
	return { map, headers: resolvedHeaders };
};

const applyTemplateToRows = async (
	token,
	id,
	name,
	rowIndexes,
	columnCount
) => {
	if (!rowIndexes || !rowIndexes.length) {
		return false;
	}
	if (typeof getSheetId !== "function") {
		return false;
	}
	const targetSheetId = await getSheetId(token, id, name);
	if (targetSheetId === null) {
		return false;
	}
	const defaultCount = getDefaultHeaderCount();
	const effectiveCount = Math.min(columnCount || defaultCount, defaultCount);
	if (typeof ensureTemplateSheet === "function" && typeof applyTemplateRowToRows === "function") {
		const templateSheetId = await ensureTemplateSheet(token, id);
		if (templateSheetId) {
			return applyTemplateRowToRows(
				token,
				id,
				templateSheetId,
				targetSheetId,
				rowIndexes,
				effectiveCount
			);
		}
	}
	if (typeof applyRowTemplatesInSheet === "function") {
		return applyRowTemplatesInSheet(
			token,
			id,
			name,
			2,
			rowIndexes,
			effectiveCount
		);
	}
	return false;
};

const buildUpdatedJob = (job, rng) => {
	const proposalsMin = randomInt(rng, 1, 6);
	const proposalsMax = proposalsMin + randomInt(rng, 2, 10);
	const invitesSent = randomInt(rng, 0, 10);
	const interviewing = randomInt(rng, 0, 4);
	const unansweredInvites = Math.max(0, invitesSent - randomInt(rng, 0, 4));
	return {
		...job,
		date: randomDate(rng),
		payment: rng() > 0.2 ? "Verified" : "Not verified",
		country: pick(rng, MOCK_COUNTRIES),
		jobCreatedSince: pick(rng, MOCK_CREATED_SINCE),
		proposals: `${proposalsMin} to ${proposalsMax}`,
		invitesSent: String(invitesSent),
		interviewing: String(interviewing),
		lastViewed: `${randomInt(rng, 1, 10)} days ago`,
		unansweredInvites: String(unansweredInvites),
	};
};

const prepareSheet = async () => {
	updateSheetStateFromInputs();
	if (!sheetId) {
		setStatus("Enter a spreadsheet ID or URL first.", "warn");
		return buildResult(false, "warn", "Missing sheet ID.");
	}
	setStatus("Preparing sheet...", "");
	const auth = await requestAuthToken(true);
	if (!auth.ok) {
		const message = auth.error || "Google authorization failed.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	await removeBlankHeaderColumn(
		auth.token,
		sheetId,
		sheetName,
		18,
		"Job Status"
	);
	const templateResult = await applyTemplateFormatting(
		auth.token,
		sheetId,
		sheetName
	);
	const headerCount = getDefaultHeaderCount();
	if (templateResult?.ok) {
		const headerResponse = await setHeaders(
			auth.token,
			sheetId,
			sheetName
		);
		if (!headerResponse.ok) {
			const message = `Sheets API error ${headerResponse.status}.`;
			setStatus(message, "error");
			return buildResult(false, "error", message);
		}
		await setHeaderBold(auth.token, sheetId, sheetName, headerCount);
		await freezeHeaderRow(auth.token, sheetId, sheetName);
		setStatus("Sheet prepared from template.", "success");
		setSheetLog("Prepare sheet complete.", ["Template formatting applied."]);
		return buildResult(true, "success", "Template formatting applied.");
	}
	const headerResponse = await setHeaders(auth.token, sheetId, sheetName);
	if (!headerResponse.ok) {
		const message = `Sheets API error ${headerResponse.status}.`;
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	await setHeaderBold(auth.token, sheetId, sheetName, headerCount);
	await clearBodyBold(auth.token, sheetId, sheetName, headerCount, 1000);
	await setBodyColumnColors(auth.token, sheetId, sheetName, 1000);
	await freezeHeaderRow(auth.token, sheetId, sheetName);
	await ensureJobStatusDropdown(auth.token, sheetId, sheetName);
	await ensureJobStatusColors(auth.token, sheetId, sheetName);
	setStatus("Sheet prepared with headers and dropdowns.", "success");
	setSheetLog("Prepare sheet complete.", ["Headers and dropdowns applied."]);
	return buildResult(true, "success", "Headers and dropdowns applied.");
};

const bulkAddJobs = async () => {
	const payload = getMockPayload();
	const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
	if (!jobs.length) {
		setStatus("Generate mock data first.", "warn");
		return buildResult(false, "warn", "Mock data missing.");
	}
	updateSheetStateFromInputs();
	if (!sheetId) {
		setStatus("Enter a spreadsheet ID or URL first.", "warn");
		return buildResult(false, "warn", "Missing sheet ID.");
	}
	setStatus("Connecting to Google...", "");
	const auth = await requestAuthToken(true);
	if (!auth.ok) {
		const message = auth.error || "Google authorization failed.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	setStatus("Reading sheet data...", "");
	let headers;
	let jobMap;
	let emptyRowInfo;
	try {
		headers = await withTimeout(
			getSheetHeaders(auth.token, sheetId, sheetName),
			20000,
			"Sheet header read timed out."
		);
		jobMap = await withTimeout(
			getJobIdRowMap(auth.token, sheetId, sheetName, headers),
			20000,
			"Sheet Job ID read timed out."
		);
		emptyRowInfo = await withTimeout(
			getEmptyRowIndexes(auth.token, sheetId, sheetName, headers),
			20000,
			"Sheet row scan timed out."
		);
	} catch (error) {
		const message = error.message || "Sheet read timed out.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	if (!headers) {
		setStatus("Unable to read sheet headers.", "error");
		return buildResult(false, "error", "Unable to read sheet headers.");
	}
	if (!jobMap) {
		setStatus("Unable to read Job ID column from the sheet.", "error");
		return buildResult(false, "error", "Unable to read Job ID column.");
	}
	if (!emptyRowInfo) {
		setStatus("Unable to read sheet rows.", "error");
		return buildResult(false, "error", "Unable to read sheet rows.");
	}
	const columnCount = headers.length || getDefaultHeaderCount();
	const endColumn = getColumnLetter(columnCount);
	const rangePrefix = getSheetRangePrefix();
	const bidder = getBidderValue();
	const updates = [];
	const newRowIndexes = [];
	const emptyRowQueue = [...emptyRowInfo.emptyRows];
	let nextRowIndex = emptyRowInfo.nextRowIndex;
	let skipped = 0;
	for (const job of jobs) {
		const jobIdValue = String(job?.jobId || "").trim();
		if (jobIdValue && jobMap.map.has(jobIdValue)) {
			skipped += 1;
			continue;
		}
		const row = buildRowFromHeaders(headers, { ...job, bidder });
		const useEmptyRow = emptyRowQueue.length > 0;
		const targetRowIndex = useEmptyRow ? emptyRowQueue.shift() : nextRowIndex;
		if (!useEmptyRow) {
			nextRowIndex += 1;
		}
		applyExistingJobStatus(row, emptyRowInfo, targetRowIndex);
		updates.push({
			range: `${rangePrefix}A${targetRowIndex}:${endColumn}${targetRowIndex}`,
			values: [row],
		});
		newRowIndexes.push(targetRowIndex);
	}
	if (!updates.length) {
		setStatus("No new jobs to add.", "warn");
		setSheetLog("Bulk add skipped.", [`Existing rows: ${skipped}`]);
		return buildResult(true, "warn", "No new jobs to add.", {
			added: 0,
			skipped,
		});
	}
	setStatus(`Writing ${updates.length} row(s)...`, "");
	const response = await batchUpdateValues(auth.token, sheetId, updates);
	if (!response.ok) {
		const message = `Sheets API error ${response.status}.`;
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	await applyTemplateToRows(
		auth.token,
		sheetId,
		sheetName,
		newRowIndexes,
		columnCount
	);
	setStatus(`Bulk add complete (${updates.length}).`, "success");
	setSheetLog("Bulk add complete.", [
		`Added rows: ${updates.length}`,
		`Skipped existing: ${skipped}`,
	]);
	return buildResult(true, "success", "Bulk add complete.", {
		added: updates.length,
		skipped,
		newRowIndexes,
	});
};

const bulkEditJobs = async () => {
	const payload = getMockPayload();
	const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
	if (!jobs.length) {
		setStatus("Generate mock data first.", "warn");
		return buildResult(false, "warn", "Mock data missing.");
	}
	updateSheetStateFromInputs();
	if (!sheetId) {
		setStatus("Enter a spreadsheet ID or URL first.", "warn");
		return buildResult(false, "warn", "Missing sheet ID.");
	}
	setStatus("Connecting to Google...", "");
	const auth = await requestAuthToken(true);
	if (!auth.ok) {
		const message = auth.error || "Google authorization failed.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	setStatus("Reading sheet data...", "");
	let headers;
	let jobMap;
	try {
		headers = await withTimeout(
			getSheetHeaders(auth.token, sheetId, sheetName),
			20000,
			"Sheet header read timed out."
		);
		jobMap = await withTimeout(
			getJobIdRowMap(auth.token, sheetId, sheetName, headers),
			20000,
			"Sheet Job ID read timed out."
		);
	} catch (error) {
		const message = error.message || "Sheet read timed out.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	if (!headers) {
		setStatus("Unable to read sheet headers.", "error");
		return buildResult(false, "error", "Unable to read sheet headers.");
	}
	if (!jobMap) {
		setStatus("Unable to read Job ID column from the sheet.", "error");
		return buildResult(false, "error", "Unable to read Job ID column.");
	}
	const rangePrefix = getSheetRangePrefix();
	const bidder = getBidderValue();
	const rng = createRng(`${seedInput.value}-bulk-edit`);
	const columnIndexes = {
		date: getHeaderIndex(headers, "Date"),
		bidder: getHeaderIndex(headers, "Bidder"),
		invites: getHeaderIndex(headers, "Invites"),
		interview: getHeaderIndex(headers, "Interview"),
		payment: getHeaderIndex(headers, "Payment"),
		country: getHeaderIndex(headers, "Country"),
		created: getHeaderIndex(headers, "Job Created Since"),
		proposals: getHeaderIndex(headers, "Proposals"),
	};
	const updates = [];
	let missing = 0;
	let updatedRows = 0;
	for (const job of jobs) {
		const jobIdValue = String(job?.jobId || "").trim();
		const rowIndex = jobIdValue ? jobMap.map.get(jobIdValue) : null;
		if (!rowIndex) {
			missing += 1;
			continue;
		}
		const updated = buildUpdatedJob(job, rng);
		const formattedDate =
			typeof formatSheetDate === "function"
				? formatSheetDate(updated.date)
				: updated.date;
		if (bidder) {
			pushCellUpdate(
				updates,
				rangePrefix,
				columnIndexes.bidder,
				rowIndex,
				bidder
			);
		}
		pushCellUpdate(
			updates,
			rangePrefix,
			columnIndexes.date,
			rowIndex,
			formattedDate
		);
		pushCellUpdate(
			updates,
			rangePrefix,
			columnIndexes.invites,
			rowIndex,
			updated.invitesSent
		);
		pushCellUpdate(
			updates,
			rangePrefix,
			columnIndexes.interview,
			rowIndex,
			updated.interviewing
		);
		pushCellUpdate(
			updates,
			rangePrefix,
			columnIndexes.payment,
			rowIndex,
			updated.payment
		);
		pushCellUpdate(
			updates,
			rangePrefix,
			columnIndexes.country,
			rowIndex,
			updated.country
		);
		pushCellUpdate(
			updates,
			rangePrefix,
			columnIndexes.created,
			rowIndex,
			updated.jobCreatedSince
		);
		pushCellUpdate(
			updates,
			rangePrefix,
			columnIndexes.proposals,
			rowIndex,
			updated.proposals
		);
		updatedRows += 1;
	}
	if (!updates.length) {
		setStatus("No matching jobs found to update.", "warn");
		setSheetLog("Bulk edit skipped.", [`Missing rows: ${missing}`]);
		return buildResult(true, "warn", "No matching jobs to update.", {
			updated: 0,
			missing,
		});
	}
	setStatus(`Updating ${updatedRows} row(s)...`, "");
	const response = await batchUpdateValues(auth.token, sheetId, updates);
	if (!response.ok) {
		const message = `Sheets API error ${response.status}.`;
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	setStatus(`Bulk edit complete (${updatedRows}).`, "success");
	setSheetLog("Bulk edit complete.", [
		`Updated rows: ${updatedRows}`,
		`Missing rows: ${missing}`,
	]);
	return buildResult(true, "success", "Bulk edit complete.", {
		updated: updatedRows,
		missing,
	});
};

const bulkMarkViewed = async () => {
	const payload = getMockPayload();
	const proposals = Array.isArray(payload?.proposals) ? payload.proposals : [];
	if (!proposals.length) {
		setStatus("Generate mock data first.", "warn");
		return buildResult(false, "warn", "Mock data missing.");
	}
	updateSheetStateFromInputs();
	if (!sheetId) {
		setStatus("Enter a spreadsheet ID or URL first.", "warn");
		return buildResult(false, "warn", "Missing sheet ID.");
	}
	setStatus("Connecting to Google...", "");
	const auth = await requestAuthToken(true);
	if (!auth.ok) {
		const message = auth.error || "Google authorization failed.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	setStatus("Reading sheet data...", "");
	let headers;
	let idMap;
	try {
		headers = await withTimeout(
			getSheetHeaders(auth.token, sheetId, sheetName),
			20000,
			"Sheet header read timed out."
		);
		idMap = await withTimeout(
			getProposalIdRowMap(auth.token, sheetId, sheetName, headers),
			20000,
			"Sheet Proposal ID read timed out."
		);
	} catch (error) {
		const message = error.message || "Sheet read timed out.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	if (!headers) {
		setStatus("Unable to read sheet headers.", "error");
		return buildResult(false, "error", "Unable to read sheet headers.");
	}
	if (!idMap) {
		setStatus("Unable to read Proposal ID column from the sheet.", "error");
		return buildResult(false, "error", "Unable to read Proposal ID column.");
	}
	const viewedRate = getViewedRate() / 100;
	const rng = createRng(`${seedInput.value}-viewed`);
	const candidates = proposals.filter(() => rng() < viewedRate);
	if (!candidates.length) {
		setStatus("No proposals selected to mark as viewed.", "warn");
		return buildResult(true, "warn", "No proposals selected.");
	}
	const matchedRows = [];
	let nameMap;
	let nameMapLoaded = false;
	for (const item of candidates) {
		const proposalId = String(item?.proposalId || "").trim();
		let rowIndex = proposalId ? idMap.get(proposalId) : null;
		if (!rowIndex) {
			const normalized = normalizeJobLabel(item?.title || "");
			if (normalized) {
				if (!nameMapLoaded) {
					nameMap = await getJobNameRowMap(
						auth.token,
						sheetId,
						sheetName,
						headers
					);
					nameMapLoaded = true;
				}
				rowIndex = nameMap?.get(normalized);
			}
		}
		if (rowIndex) {
			matchedRows.push(rowIndex);
		}
	}
	const uniqueRows = Array.from(new Set(matchedRows)).sort((a, b) => a - b);
	if (!uniqueRows.length) {
		setStatus("No matching rows found to mark as viewed.", "warn");
		return buildResult(true, "warn", "No matching rows found.");
	}
	const readColumnIndex = getHeaderIndex(headers || [], "Read");
	const ok = await setReadCellsViewed(
		auth.token,
		sheetId,
		sheetName,
		uniqueRows,
		readColumnIndex ? readColumnIndex - 1 : undefined
	);
	if (!ok) {
		setStatus("Failed to update the sheet formatting.", "error");
		return buildResult(false, "error", "Failed to update sheet formatting.");
	}
	setStatus(`Marked ${uniqueRows.length} row(s) as read.`, "success");
	setSheetLog("Bulk viewed update complete.", [
		`Candidates: ${candidates.length}`,
		`Matched rows: ${uniqueRows.length}`,
	]);
	return buildResult(true, "success", "Bulk viewed update complete.", {
		candidates: candidates.length,
		matched: uniqueRows.length,
	});
};

const bulkUpdateConnects = async () => {
	const payload = getMockPayload();
	const entries = Array.isArray(payload?.connects) ? payload.connects : [];
	if (!entries.length) {
		setStatus("Generate mock data first.", "warn");
		return buildResult(false, "warn", "Mock data missing.");
	}
	updateSheetStateFromInputs();
	if (!sheetId) {
		setStatus("Enter a spreadsheet ID or URL first.", "warn");
		return buildResult(false, "warn", "Missing sheet ID.");
	}
	setStatus("Connecting to Google...", "");
	const auth = await requestAuthToken(true);
	if (!auth.ok) {
		const message = auth.error || "Google authorization failed.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	setStatus("Reading sheet data...", "");
	let headers;
	let connectsMap;
	try {
		headers = await withTimeout(
			getSheetHeaders(auth.token, sheetId, sheetName),
			20000,
			"Sheet header read timed out."
		);
		connectsMap = await withTimeout(
			getConnectsRowMap(auth.token, sheetId, sheetName, headers),
			20000,
			"Sheet Job ID read timed out."
		);
	} catch (error) {
		const message = error.message || "Sheet read timed out.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	if (!headers) {
		setStatus("Unable to read sheet headers.", "error");
		return buildResult(false, "error", "Unable to read sheet headers.");
	}
	if (!connectsMap) {
		setStatus("Unable to read Job ID column from the sheet.", "error");
		return buildResult(false, "error", "Unable to read Job ID column.");
	}
	const totals = new Map();
	for (const entry of entries) {
		const jobIdValue = entry?.jobId ? String(entry.jobId).trim() : "";
		const key = jobIdValue || normalizeJobLabel(entry?.title || "");
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
		target.boostedConnectsRefund += Number(entry.boostedConnectsRefund || 0);
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
		setStatus("No connects entries found to update.", "warn");
		return buildResult(true, "warn", "No connects entries found.");
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
		return buildResult(
			true,
			"warn",
			"Missing boosted connects columns."
		);
	}
	let emptyRowInfo;
	try {
		emptyRowInfo = await withTimeout(
			getEmptyRowIndexes(auth.token, sheetId, sheetName, headers),
			20000,
			"Sheet row scan timed out."
		);
	} catch (error) {
		const message = error.message || "Sheet read timed out.";
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	if (!emptyRowInfo) {
		setStatus("Unable to read sheet rows.", "error");
		return buildResult(false, "error", "Unable to read sheet rows.");
	}
	const updates = [];
	const newRowIndexes = [];
	const emptyRowQueue = [...emptyRowInfo.emptyRows];
	let nextRowIndex = emptyRowInfo.nextRowIndex;
	const rangePrefix = getSheetRangePrefix();
	const columnCount = connectsMap.columnCount || headers.length || getDefaultHeaderCount();
	const endColumn = getColumnLetter(columnCount);
	const bidder = getBidderValue();
	let jobNameMap;
	let jobNameMapLoaded = false;
	let updated = 0;
	let added = 0;
	for (const entry of totals.values()) {
		let existing = entry.jobId ? connectsMap.map.get(entry.jobId) : null;
		if (!existing) {
			const normalizedName = normalizeJobLabel(entry.name || "");
			if (normalizedName) {
				if (!jobNameMapLoaded) {
					jobNameMap = await getJobNameRowMap(
						auth.token,
						sheetId,
						sheetName,
						headers
					);
					jobNameMapLoaded = true;
				}
				const fallbackRowIndex = jobNameMap?.get(normalizedName);
				if (fallbackRowIndex) {
					existing = connectsMap.rowsByIndex.get(fallbackRowIndex) || {
						rowIndex: fallbackRowIndex,
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
			if (bidder) {
				updates.push({
					range: `${rangePrefix}C${existing.rowIndex}:C${existing.rowIndex}`,
					values: [[bidder]],
				});
			}
			if (connectsMap.connectsSpentColumn) {
				updates.push({
					range: `${rangePrefix}${connectsMap.connectsSpentColumn}${existing.rowIndex}`,
					values: [[nextSpentValue]],
				});
			}
			if (connectsMap.connectsRefundColumn) {
				updates.push({
					range: `${rangePrefix}${connectsMap.connectsRefundColumn}${existing.rowIndex}`,
					values: [[nextRefundValue]],
				});
			}
			if (connectsMap.boostedConnectsSpentColumn) {
				updates.push({
					range: `${rangePrefix}${connectsMap.boostedConnectsSpentColumn}${existing.rowIndex}`,
					values: [[nextBoostedSpentValue]],
				});
			}
			if (connectsMap.boostedConnectsRefundColumn) {
				updates.push({
					range: `${rangePrefix}${connectsMap.boostedConnectsRefundColumn}${existing.rowIndex}`,
					values: [[nextBoostedRefundValue]],
				});
			}
			updated += 1;
			continue;
		}
		const row = buildRowFromHeaders(headers, {
			date: entry.date,
			name: entry.name,
			link: entry.link,
			bidder,
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
		const targetRowIndex = useEmptyRow ? emptyRowQueue.shift() : nextRowIndex;
		if (!useEmptyRow) {
			nextRowIndex += 1;
		}
		applyExistingJobStatus(row, emptyRowInfo, targetRowIndex);
		updates.push({
			range: `${rangePrefix}A${targetRowIndex}:${endColumn}${targetRowIndex}`,
			values: [row],
		});
		newRowIndexes.push(targetRowIndex);
		added += 1;
	}
	if (!updates.length) {
		setStatus("No connects updates to apply.", "warn");
		return buildResult(true, "warn", "No connects updates to apply.");
	}
	setStatus("Updating sheet...", "");
	const response = await batchUpdateValues(auth.token, sheetId, updates);
	if (!response.ok) {
		const message = `Sheets API error ${response.status}.`;
		setStatus(message, "error");
		return buildResult(false, "error", message);
	}
	await applyTemplateToRows(
		auth.token,
		sheetId,
		sheetName,
		newRowIndexes,
		columnCount
	);
	setStatus("Connects history updated.", "success");
	setSheetLog("Bulk connects complete.", [
		`Updated rows: ${updated}`,
		`Added rows: ${added}`,
	]);
	return buildResult(true, "success", "Connects history updated.", {
		updated,
		added,
	});
};

const formatTestLine = (label, result) => {
	if (result?.skipped) {
		return `SKIP: ${label} - ${result.message || "Skipped."}`;
	}
	const tone = result?.tone || (result?.ok ? "success" : "error");
	const tag = tone === "success" ? "OK" : tone === "warn" ? "WARN" : "FAIL";
	const detail = result?.message ? ` - ${result.message}` : "";
	return `${tag}: ${label}${detail}`;
};

const runFullTest = async () => {
	if (fullTestButton) {
		fullTestButton.disabled = true;
	}
	try {
		setStatus("Running full test...", "");
		const results = [];
		let halted = false;
		let authToken = "";
		let headers = null;
		let columnCount = getDefaultHeaderCount();
		let preJobIds = new Set();
		let preDuplicateJobIds = new Set();
		let bulkAddResult = null;
		let bulkViewedResult = null;
		let bulkConnectsResult = null;
		let sampleJob = null;
		let sampleRowIndex = null;
		let sampleRowBefore = null;
		let viewedMatchedRows = [];

		const runStep = async (label, action, options = {}) => {
			const { critical = true } = options;
			if (halted) {
				results.push({
					label,
					...buildResult(false, "warn", "Skipped after failure.", {
						skipped: true,
					}),
				});
				return null;
			}
			let result;
			try {
				result = await action();
			} catch (error) {
				result = buildResult(
					false,
					"error",
					error?.message || "Unexpected error."
				);
			}
			const normalized = result || buildResult(false, "error", "No result.");
			results.push({ ...normalized, label });
			if (critical && !normalized.ok && normalized.tone === "error") {
				halted = true;
			}
			return normalized;
		};

		const validateHeaders = (currentHeaders, strictTemplate) => {
			if (!Array.isArray(currentHeaders) || !currentHeaders.length) {
				return buildResult(false, "error", "No headers returned.");
			}
			const required = [
				"Job Name",
				"Job ID",
				"Proposal ID",
				"Read",
				"Invites",
				"Interview",
				"Proposals",
				"Connects Spent",
				"Connects Refund",
			];
			const missing = required.filter(
				(name) => !getHeaderIndex(currentHeaders, name)
			);
			if (missing.length) {
				return buildResult(
					false,
					"error",
					`Missing headers: ${missing.join(", ")}.`
				);
			}
			const boostedRequired = getConnectsBoostedRate() > 0;
			const boostedMissing = boostedRequired
				? [
						!getHeaderIndex(currentHeaders, "Boosted Connects Spent")
							? "Boosted Connects Spent"
							: null,
						!getHeaderIndex(currentHeaders, "Boosted Connects Refund")
							? "Boosted Connects Refund"
							: null,
					].filter(Boolean)
				: [];
			if (boostedMissing.length) {
				return buildResult(
					false,
					"error",
					`Missing boosted columns: ${boostedMissing.join(", ")}.`
				);
			}
			const warnings = [];
			if (currentHeaders.length < getDefaultHeaderCount()) {
				warnings.push("Header count is shorter than the template.");
			}
			if (typeof DEFAULT_HEADERS !== "undefined") {
				const matchesTemplate = DEFAULT_HEADERS.every(
					(header, index) => currentHeaders[index] === header
				);
				if (!matchesTemplate) {
					if (strictTemplate) {
						return buildResult(
							false,
							"error",
							"Header row does not match the template."
						);
					}
					warnings.push("Header row does not match the template.");
				}
			}
			const counts = new Map();
			currentHeaders.forEach((header) => {
				const key = String(header || "").trim();
				if (!key) {
					return;
				}
				counts.set(key, (counts.get(key) || 0) + 1);
			});
			const duplicates = Array.from(counts.entries())
				.filter((entry) => entry[1] > 1)
				.map((entry) => entry[0]);
			if (duplicates.length) {
				warnings.push(`Duplicate headers: ${duplicates.join(", ")}.`);
			}
			if (warnings.length) {
				return buildResult(true, "warn", warnings.join(" "));
			}
			return buildResult(true, "success", "Headers look good.");
		};

		await runStep("Validate sheet settings", () => {
			updateSheetStateFromInputs();
			if (!sheetId || !sheetName) {
				return buildResult(false, "error", "Sheet ID or tab is missing.");
			}
			return buildResult(true, "success", "Sheet settings ready.");
		});

		await runStep(
			"Ensure mock data",
			() => {
				if (!getMockPayload()) {
					generateMockData();
				}
				return buildResult(true, "success", "Mock data ready.");
			},
			{ critical: false }
		);

		await runStep("Authorize Google", async () => {
			const auth = await requestAuthToken(true);
			if (!auth.ok) {
				return buildResult(
					false,
					"error",
					auth.error || "Google authorization failed."
				);
			}
			authToken = auth.token;
			return buildResult(true, "success", "Authorized.");
		});

		await runStep("Validate sheet tab", async () => {
			if (!authToken) {
				return buildResult(false, "error", "Missing auth token.");
			}
			const tabId = await getSheetId(authToken, sheetId, sheetName);
			if (tabId === null) {
				return buildResult(false, "error", "Sheet tab not found.");
			}
			return buildResult(true, "success", "Sheet tab found.");
		});

		await runStep("Read headers", async () => {
			if (!authToken) {
				return buildResult(false, "error", "Missing auth token.");
			}
			headers = await getSheetHeaders(authToken, sheetId, sheetName);
			if (!headers) {
				return buildResult(false, "error", "Unable to read headers.");
			}
			columnCount = headers.length || getDefaultHeaderCount();
			return validateHeaders(headers, false);
		});

		await runStep(
			"Check Job Status dropdown",
			async () => {
				if (!authToken || !headers) {
					return buildResult(false, "error", "Missing sheet context.");
				}
				const jobStatusIndex = getHeaderIndex(headers, "Job Status");
				if (!jobStatusIndex) {
					return buildResult(true, "warn", "Job Status header missing.");
				}
				const cell = `${getColumnLetter(jobStatusIndex)}2`;
				const validation = await getCellDataValidation(
					authToken,
					sheetId,
					sheetName,
					cell
				);
				if (!validation) {
					return buildResult(
						true,
						"warn",
						"Job Status dropdown is missing."
					);
				}
				return buildResult(true, "success", "Job Status dropdown found.");
			},
			{ critical: false }
		);

		await runStep(
			"Check bidder dropdown",
			async () => {
				if (!authToken || !headers) {
					return buildResult(false, "error", "Missing sheet context.");
				}
				const bidderIndex = getHeaderIndex(headers, "Bidder");
				if (!bidderIndex) {
					return buildResult(true, "warn", "Bidder header missing.");
				}
				const cell = `${getColumnLetter(bidderIndex)}2`;
				const validation = await getCellDataValidation(
					authToken,
					sheetId,
					sheetName,
					cell
				);
				if (!validation) {
					return buildResult(
						true,
						"warn",
						"Bidder dropdown is missing."
					);
				}
				return buildResult(true, "success", "Bidder dropdown found.");
			},
			{ critical: false }
		);

		await runStep(
			"Check empty rows",
			async () => {
				if (!authToken || !headers) {
					return buildResult(false, "error", "Missing sheet context.");
				}
				const emptyInfo = await getEmptyRowIndexes(
					authToken,
					sheetId,
					sheetName,
					headers
				);
				if (!emptyInfo) {
					return buildResult(true, "warn", "Unable to scan empty rows.");
				}
				if (!emptyInfo.emptyRows.length) {
					return buildResult(
						true,
						"warn",
						"No empty rows found; bulk add will append."
					);
				}
				return buildResult(
					true,
					"success",
					`Empty rows: ${emptyInfo.emptyRows.length}.`
				);
			},
			{ critical: false }
		);

		await runStep("Prepare sheet", prepareSheet);

		await runStep("Re-check headers after prepare", async () => {
			if (!authToken) {
				return buildResult(false, "error", "Missing auth token.");
			}
			headers = await getSheetHeaders(authToken, sheetId, sheetName);
			if (!headers) {
				return buildResult(false, "error", "Unable to read headers.");
			}
			columnCount = headers.length || getDefaultHeaderCount();
			return validateHeaders(headers, true);
		});

		await runStep("Snapshot Job IDs", async () => {
			if (!authToken || !headers) {
				return buildResult(false, "error", "Missing sheet context.");
			}
			const jobIdIndex = getHeaderIndex(headers, "Job ID");
			if (!jobIdIndex) {
				return buildResult(false, "error", "Job ID header missing.");
			}
			const rows = await getColumnValues(
				authToken,
				sheetId,
				sheetName,
				jobIdIndex - 1,
				2
			);
			if (!rows) {
				return buildResult(false, "error", "Unable to read Job ID column.");
			}
			const counts = new Map();
			rows.forEach((row) => {
				const value = normalizeCellValue(row?.[0]);
				if (!value) {
					return;
				}
				counts.set(value, (counts.get(value) || 0) + 1);
			});
			preJobIds = new Set(counts.keys());
			preDuplicateJobIds = new Set(
				Array.from(counts.entries())
					.filter((entry) => entry[1] > 1)
					.map((entry) => entry[0])
			);
			if (preDuplicateJobIds.size) {
				return buildResult(
					true,
					"warn",
					"Duplicate Job IDs already exist in sheet."
				);
			}
			return buildResult(true, "success", "Captured Job ID snapshot.");
		});

		bulkAddResult = await runStep("Bulk add jobs", bulkAddJobs);

		await runStep("Verify bulk add results", async () => {
			if (!authToken || !headers) {
				return buildResult(false, "error", "Missing sheet context.");
			}
			const payload = getMockPayload();
			const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
			if (!jobs.length) {
				return buildResult(false, "error", "Mock data missing.");
			}
			const jobIdIndex = getHeaderIndex(headers, "Job ID");
			const rows = await getColumnValues(
				authToken,
				sheetId,
				sheetName,
				jobIdIndex - 1,
				2
			);
			if (!rows) {
				return buildResult(false, "error", "Unable to read Job IDs.");
			}
			const counts = new Map();
			rows.forEach((row) => {
				const value = normalizeCellValue(row?.[0]);
				if (!value) {
					return;
				}
				counts.set(value, (counts.get(value) || 0) + 1);
			});
			const afterJobIds = new Set(counts.keys());
			const afterDuplicateJobIds = new Set(
				Array.from(counts.entries())
					.filter((entry) => entry[1] > 1)
					.map((entry) => entry[0])
			);
			const newDuplicates = Array.from(afterDuplicateJobIds).filter(
				(id) => !preDuplicateJobIds.has(id)
			);
			if (newDuplicates.length) {
				return buildResult(
					false,
					"error",
					`New duplicate Job IDs detected: ${newDuplicates.join(", ")}.`
				);
			}
			const missingFromSheet = jobs
				.map((job) => normalizeCellValue(job?.jobId))
				.filter((jobId) => jobId && !afterJobIds.has(jobId));
			if (missingFromSheet.length) {
				return buildResult(
					false,
					"error",
					`Missing Job IDs after add: ${missingFromSheet.slice(0, 5).join(", ")}.`
				);
			}
			const missingExisting = Array.from(preJobIds).filter(
				(jobId) => !afterJobIds.has(jobId)
			);
			if (missingExisting.length) {
				return buildResult(
					false,
					"error",
					`Existing Job IDs disappeared: ${missingExisting.slice(0, 5).join(", ")}.`
				);
			}
			const payloadUnique = new Set(
				jobs.map((job) => normalizeCellValue(job?.jobId)).filter(Boolean)
			);
			const expectedAdded = Array.from(payloadUnique).filter(
				(jobId) => !preJobIds.has(jobId)
			).length;
			if (
				typeof bulkAddResult?.added === "number" &&
				bulkAddResult.added !== expectedAdded
			) {
				return buildResult(
					true,
					"warn",
					`Expected ${expectedAdded} new rows, got ${bulkAddResult.added}.`
				);
			}
			return buildResult(true, "success", "Bulk add verified.");
		});

		await runStep(
			"Capture sample row",
			async () => {
				if (!authToken || !headers) {
					return buildResult(false, "error", "Missing sheet context.");
				}
				const payload = getMockPayload();
				const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
				if (!jobs.length) {
					return buildResult(false, "error", "Mock data missing.");
				}
				const jobMap = await getJobIdRowMap(
					authToken,
					sheetId,
					sheetName,
					headers
				);
				if (!jobMap) {
					return buildResult(false, "error", "Job ID map unavailable.");
				}
				for (const job of jobs) {
					const jobIdValue = normalizeCellValue(job?.jobId);
					const rowIndex = jobIdValue
						? jobMap.map.get(jobIdValue)
						: null;
					if (rowIndex) {
						sampleJob = job;
						sampleRowIndex = rowIndex;
						break;
					}
				}
				if (!sampleJob || !sampleRowIndex) {
					return buildResult(
						true,
						"warn",
						"No matching job found for sample checks."
					);
				}
				sampleRowBefore = await readSheetRow(
					authToken,
					sheetId,
					sheetName,
					columnCount,
					sampleRowIndex
				);
				if (!sampleRowBefore) {
					return buildResult(
						true,
						"warn",
						"Unable to capture sample row."
					);
				}
				return buildResult(true, "success", "Sample row captured.");
			},
			{ critical: false }
		);

		await runStep("Bulk edit jobs", bulkEditJobs);

		await runStep(
			"Verify bulk edit results",
			async () => {
				if (
					!authToken ||
					!headers ||
					!sampleJob ||
					!sampleRowIndex ||
					!sampleRowBefore
				) {
					return buildResult(
						true,
						"warn",
						"Sample row unavailable for edit validation."
					);
				}
				const sampleRowAfter = await readSheetRow(
					authToken,
					sheetId,
					sheetName,
					columnCount,
					sampleRowIndex
				);
				if (!sampleRowAfter) {
					return buildResult(
						false,
						"error",
						"Unable to read sample row after edit."
					);
				}
				const allowedHeaders = new Set([
					"Date",
					"Bidder",
					"Invites",
					"Interview",
					"Payment",
					"Country",
					"Job Created Since",
					"Proposals",
				]);
				const unexpectedChanges = [];
				headers.forEach((header, index) => {
					const name = String(header || "").trim();
					if (!name || allowedHeaders.has(name)) {
						return;
					}
					const before = normalizeCellValue(sampleRowBefore[index]);
					const after = normalizeCellValue(sampleRowAfter[index]);
					if (before !== after) {
						unexpectedChanges.push(name);
					}
				});
				if (unexpectedChanges.length) {
					return buildResult(
						false,
						"error",
						`Unexpected changes in: ${unexpectedChanges.join(", ")}.`
					);
				}
				const payload = getMockPayload();
				const jobs = Array.isArray(payload?.jobs) ? payload.jobs : [];
				const expected = computeUpdatedJobById(
					jobs,
					`${seedInput.value}-bulk-edit`,
					normalizeCellValue(sampleJob?.jobId)
				);
				if (!expected) {
					return buildResult(true, "warn", "Expected data not found.");
				}
				const checks = [
					{ header: "Invites", value: expected.invitesSent },
					{ header: "Interview", value: expected.interviewing },
					{ header: "Proposals", value: expected.proposals },
					{ header: "Payment", value: expected.payment },
					{ header: "Country", value: expected.country },
					{ header: "Job Created Since", value: expected.jobCreatedSince },
				];
				const bidderValue = getBidderValue();
				if (bidderValue) {
					checks.push({ header: "Bidder", value: bidderValue });
				}
				const dateIndex = getHeaderIndex(headers, "Date");
				if (dateIndex) {
					checks.push({
						header: "Date",
						value: stripLeadingApostrophe(expected.date),
						strip: true,
					});
				}
				const mismatches = [];
				checks.forEach((check) => {
					const index = getHeaderIndex(headers, check.header);
					if (!index) {
						return;
					}
					const actual = check.strip
						? stripLeadingApostrophe(getRowValue(sampleRowAfter, index))
						: normalizeCellValue(getRowValue(sampleRowAfter, index));
					const expectedValue = check.strip
						? stripLeadingApostrophe(check.value)
						: normalizeCellValue(check.value);
					if (actual !== expectedValue) {
						mismatches.push(check.header);
					}
				});
				if (mismatches.length) {
					return buildResult(
						false,
						"error",
						`Edited values mismatch: ${mismatches.join(", ")}.`
					);
				}
				return buildResult(true, "success", "Bulk edit verified.");
			},
			{ critical: false }
		);

		await runStep(
			"Bulk add idempotency",
			async () => {
				const result = await bulkAddJobs();
				if (!result || typeof result.added !== "number") {
					return buildResult(true, "warn", "Idempotency result unavailable.");
				}
				if (result.added !== 0) {
					return buildResult(
						false,
						"error",
						"Bulk add added rows on second pass."
					);
				}
				return buildResult(true, "success", "Bulk add is idempotent.");
			},
			{ critical: false }
		);

		bulkViewedResult = await runStep("Mark viewed", bulkMarkViewed);

		await runStep(
			"Verify viewed selection",
			async () => {
				if (!authToken || !headers) {
					return buildResult(false, "error", "Missing sheet context.");
				}
				const payload = getMockPayload();
				const proposals = Array.isArray(payload?.proposals)
					? payload.proposals
					: [];
				if (!proposals.length) {
					return buildResult(false, "error", "Mock proposals missing.");
				}
				const viewedRate = getViewedRate() / 100;
				const rng = createRng(`${seedInput.value}-viewed`);
				const candidates = proposals.filter(() => rng() < viewedRate);
				const idMap = await getProposalIdRowMap(
					authToken,
					sheetId,
					sheetName,
					headers
				);
				const nameMap = await getJobNameRowMap(
					authToken,
					sheetId,
					sheetName,
					headers
				);
				const matchedRows = [];
				candidates.forEach((item) => {
					const proposalId = normalizeCellValue(item?.proposalId);
					let rowIndex = proposalId ? idMap?.get(proposalId) : null;
					if (!rowIndex) {
						const normalized = normalizeJobLabel(item?.title || "");
						rowIndex = normalized ? nameMap?.get(normalized) : null;
					}
					if (rowIndex) {
						matchedRows.push(rowIndex);
					}
				});
				viewedMatchedRows = Array.from(new Set(matchedRows)).sort((a, b) => a - b);
				const expectedMatched = viewedMatchedRows.length;
				if (!bulkViewedResult) {
					return buildResult(true, "warn", "Viewed result unavailable.");
				}
				if (
					typeof bulkViewedResult.candidates === "number" &&
					bulkViewedResult.candidates !== candidates.length
				) {
					return buildResult(
						true,
						"warn",
						"Viewed candidate counts differ."
					);
				}
				if (
					typeof bulkViewedResult.matched === "number" &&
					bulkViewedResult.matched !== expectedMatched
				) {
					return buildResult(
						false,
						"error",
						"Viewed matched rows differ from expectation."
					);
				}
				return buildResult(true, "success", "Viewed selection verified.");
			},
			{ critical: false }
		);

		bulkConnectsResult = await runStep("Update connects", bulkUpdateConnects);

		await runStep(
			"Verify connects values",
			async () => {
				if (!authToken || !headers) {
					return buildResult(false, "error", "Missing sheet context.");
				}
				const payload = getMockPayload();
				const entries = Array.isArray(payload?.connects) ? payload.connects : [];
				if (!entries.length) {
					return buildResult(false, "error", "Mock connects missing.");
				}
				const totals = new Map();
				entries.forEach((entry) => {
					const jobIdValue = normalizeCellValue(entry?.jobId);
					const key = jobIdValue || normalizeJobLabel(entry?.title || "");
					if (!key) {
						return;
					}
					if (!totals.has(key)) {
						totals.set(key, {
							jobId: jobIdValue,
							name: entry.title || "",
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
					target.boostedConnectsRefund += Number(entry.boostedConnectsRefund || 0);
				});
				if (!totals.size) {
					return buildResult(true, "warn", "No connects totals found.");
				}
				const connectsMap = await getConnectsRowMap(
					authToken,
					sheetId,
					sheetName,
					headers
				);
				if (!connectsMap) {
					return buildResult(false, "error", "Connects map unavailable.");
				}
				let sampleEntry = null;
				let rowIndex = null;
				for (const entry of totals.values()) {
					if (entry.jobId && connectsMap.map.has(entry.jobId)) {
						sampleEntry = entry;
						rowIndex = connectsMap.map.get(entry.jobId).rowIndex;
						break;
					}
				}
				if (!sampleEntry) {
					for (const entry of totals.values()) {
						const normalized = normalizeJobLabel(entry.name || "");
						if (!normalized) {
							continue;
						}
						const nameMap = await getJobNameRowMap(
							authToken,
							sheetId,
							sheetName,
							headers
						);
						rowIndex = nameMap?.get(normalized) || null;
						if (rowIndex) {
							sampleEntry = entry;
							break;
						}
					}
				}
				if (!sampleEntry || !rowIndex) {
					return buildResult(
						true,
						"warn",
						"No connects row found for validation."
					);
				}
				const row = await readSheetRow(
					authToken,
					sheetId,
					sheetName,
					columnCount,
					rowIndex
				);
				if (!row) {
					return buildResult(false, "error", "Unable to read connects row.");
				}
				const checks = [
					{
						header: "Connects Spent",
						value: formatConnectsValue(sampleEntry.connectsSpent, "spent"),
					},
					{
						header: "Connects Refund",
						value: formatConnectsValue(sampleEntry.connectsRefund, "refund"),
					},
					{
						header: "Boosted Connects Spent",
						value: formatConnectsValue(
							sampleEntry.boostedConnectsSpent,
							"spent"
						),
					},
					{
						header: "Boosted Connects Refund",
						value: formatConnectsValue(
							sampleEntry.boostedConnectsRefund,
							"refund"
						),
					},
				];
				const mismatches = [];
				checks.forEach((check) => {
					const index = getHeaderIndex(headers, check.header);
					if (!index) {
						if (check.value) {
							mismatches.push(check.header);
						}
						return;
					}
					const actual = normalizeCellValue(getRowValue(row, index));
					const expectedValue = normalizeCellValue(check.value);
					if (actual !== expectedValue) {
						mismatches.push(check.header);
					}
				});
				if (mismatches.length) {
					return buildResult(
						false,
						"error",
						`Connects values mismatch: ${mismatches.join(", ")}.`
					);
				}
				if (bulkConnectsResult?.ok === false) {
					return buildResult(
						false,
						"error",
						"Connects update reported failure."
					);
				}
				return buildResult(true, "success", "Connects values verified.");
			},
			{ critical: false }
		);

		await runStep(
			"Fetch full sheet data and styles",
			async () => {
				if (!authToken || !headers) {
					return buildResult(false, "error", "Missing sheet context.");
				}
				const endColumn = getColumnLetter(columnCount);
				const values = await getSheetRows(
					authToken,
					sheetId,
					sheetName,
					columnCount,
					1
				);
				if (!values) {
					return buildResult(false, "error", "Unable to read sheet values.");
				}
				const lastRow = Math.max(values.length, 1);
				const normalizedName =
					typeof normalizeSheetName === "function"
						? normalizeSheetName(sheetName)
						: sheetName;
				const escaped = String(normalizedName || "").replace(/'/g, "''");
				const range = `'${escaped}'!A1:${endColumn}${lastRow}`;
				const fields =
					"sheets(data.rowData.values(userEnteredValue,userEnteredFormat),properties)";
				const gridResponse = await getSheetGridData(
					authToken,
					sheetId,
					range,
					fields
				);
				if (!gridResponse) {
					return buildResult(false, "error", "Unable to read sheet grid data.");
				}
				const gridData = gridResponse?.sheets?.[0]?.data?.[0];
				if (!gridData) {
					return buildResult(
						false,
						"error",
						"Sheet grid data missing from response."
					);
				}
				let nonEmptyCells = 0;
				let rowsWithValues = 0;
				values.forEach((row) => {
					const rowHasValue = row?.some((cell) => normalizeCellValue(cell));
					if (rowHasValue) {
						rowsWithValues += 1;
					}
					row?.forEach((cell) => {
						if (normalizeCellValue(cell)) {
							nonEmptyCells += 1;
						}
					});
				});
				const headerRow = values[0] || [];
				const headerMismatches = headers
					.map((header, index) => ({
						header,
						value: normalizeCellValue(headerRow[index]),
					}))
					.filter(
						(item) =>
							normalizeCellValue(item.header) !== normalizeCellValue(item.value)
					)
					.map((item) => item.header);

				const headerRowData = getGridRow(gridData, 1);
				const headerStyleMissing = [];
				for (let i = 1; i <= columnCount; i += 1) {
					const cell = getGridCell(headerRowData, i);
					const format = cell?.userEnteredFormat;
					const bold = format?.textFormat?.bold;
					if (!bold) {
						headerStyleMissing.push(getHeaderLabel(headers, i));
					}
				}

				let templateHeaderMismatch = [];
				let templateRowMismatch = [];
				let templateWarning = "";
				const templateMeta = await getSpreadsheetMeta(
					authToken,
					TEMPLATE_SPREADSHEET_ID,
					"sheets.properties"
				);
				const templateTitle =
					templateMeta?.sheets?.[0]?.properties?.title || null;
				if (templateTitle) {
					const templateRange = `'${templateTitle.replace(/'/g, "''")}'!A1:${endColumn}2`;
					const templateGridResponse = await getSheetGridData(
						authToken,
						TEMPLATE_SPREADSHEET_ID,
						templateRange,
						fields
					);
					const templateGrid = templateGridResponse?.sheets?.[0]?.data?.[0];
					if (templateGrid) {
						const templateHeaderRow = getGridRow(templateGrid, 1);
						const templateRow = getGridRow(templateGrid, 2);
						templateHeaderMismatch = compareRowFormats(
							headerRowData,
							templateHeaderRow,
							columnCount,
							headers
						);
						const rowsToCheck =
							bulkAddResult?.newRowIndexes?.length
								? bulkAddResult.newRowIndexes
								: lastRow >= 2
									? [2]
									: [];
						rowsToCheck.forEach((rowIndex) => {
							const currentRow = getGridRow(gridData, rowIndex);
							if (!currentRow || !templateRow) {
								return;
							}
							const mismatches = compareRowFormats(
								currentRow,
								templateRow,
								columnCount,
								headers
							);
							if (mismatches.length) {
								templateRowMismatch.push(
									`Row ${rowIndex}: ${mismatches.join(", ")}`
								);
							}
						});
					} else {
						templateWarning = "Template grid data unavailable.";
					}
				} else {
					templateWarning = "Template sheet metadata unavailable.";
				}

				const readColumnIndex = getHeaderIndex(headers, "Read");
				let viewedStyleMismatch = [];
				if (readColumnIndex && viewedMatchedRows.length) {
					viewedMatchedRows.forEach((rowIndex) => {
						const rowData = getGridRow(gridData, rowIndex);
						const cell = getGridCell(rowData, readColumnIndex);
						const color = normalizeColor(
							cell?.userEnteredFormat?.backgroundColor
						);
						if (!isGreenish(color)) {
							viewedStyleMismatch.push(String(rowIndex));
						}
					});
				}

				const warnings = [];
				const errors = [];
				if (headerMismatches.length) {
					errors.push(
						`Header values mismatch: ${headerMismatches.join(", ")}.`
					);
				}
				if (headerStyleMissing.length) {
					errors.push(
						`Header bold missing: ${headerStyleMissing.slice(0, 6).join(", ")}.`
					);
				}
				if (templateHeaderMismatch.length) {
					errors.push(
						`Template header format mismatch: ${templateHeaderMismatch
							.slice(0, 6)
							.join(", ")}.`
					);
				}
				if (templateRowMismatch.length) {
					errors.push(
						`Template row format mismatch: ${templateRowMismatch
							.slice(0, 3)
							.join(" | ")}.`
					);
				}
				if (viewedStyleMismatch.length) {
					errors.push(
						`Viewed highlight mismatch at rows: ${viewedStyleMismatch
							.slice(0, 8)
							.join(", ")}.`
					);
				}
				if (templateWarning) {
					warnings.push(templateWarning);
				}
				if (!errors.length && !warnings.length) {
					return buildResult(
						true,
						"success",
						`Sheet snapshot OK. Rows: ${rowsWithValues}, cells: ${nonEmptyCells}.`
					);
				}
				if (!errors.length) {
					return buildResult(
						true,
						"warn",
						`${warnings.join(" ")} Rows: ${rowsWithValues}, cells: ${nonEmptyCells}.`
					);
				}
				return buildResult(
					false,
					"error",
					`${errors.concat(warnings).join(" ")} Rows: ${rowsWithValues}, cells: ${nonEmptyCells}.`
				);
			},
			{ critical: false }
		);

		const lines = results.map((result) =>
			formatTestLine(result.label, result)
		);
		setSheetLog("Full test complete.", lines);
		setStatus(
			halted ? "Full test finished with errors." : "Full test complete.",
			halted ? "error" : "success"
		);
	} finally {
		if (fullTestButton) {
			fullTestButton.disabled = false;
		}
	}
};

const fixtures = {
	job: "job_page.html",
	proposals: "Proposals.html",
	connects: "ConnectsHistory.html",
};

document.querySelectorAll("[data-fixture]").forEach((button) => {
	button.addEventListener("click", () => {
		const key = button.getAttribute("data-fixture");
		const fixture = fixtures[key];
		if (!fixture) {
			return;
		}
		runFixture(fixture);
	});
});

generateButton.addEventListener("click", generateMockData);
copyButton.addEventListener("click", copyJson);

if (sheetIdInput) {
	sheetIdInput.addEventListener("input", updateSheetStateFromInputs);
}
if (sheetNameInput) {
	sheetNameInput.addEventListener("input", updateSheetStateFromInputs);
}
if (saveSheetButton) {
	saveSheetButton.addEventListener("click", saveSheetSettings);
}
if (openSheetButton) {
	openSheetButton.addEventListener("click", () => {
		updateSheetStateFromInputs();
		if (!sheetId) {
			setStatus("Enter a spreadsheet ID or URL first.", "warn");
			return;
		}
		openSpreadsheet(sheetId);
	});
}
if (prepareSheetButton) {
	prepareSheetButton.addEventListener("click", prepareSheet);
}
if (bulkAddButton) {
	bulkAddButton.addEventListener("click", bulkAddJobs);
}
if (bulkEditButton) {
	bulkEditButton.addEventListener("click", bulkEditJobs);
}
if (bulkViewedButton) {
	bulkViewedButton.addEventListener("click", bulkMarkViewed);
}
if (bulkConnectsButton) {
	bulkConnectsButton.addEventListener("click", bulkUpdateConnects);
}
if (fullTestButton) {
	fullTestButton.addEventListener("click", runFullTest);
}

loadSheetSettings();

if (jsonOutput.value.trim() === "") {
	generateMockData();
}
