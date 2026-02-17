window.extractUpworkJobData = function extractUpworkJobData(docArg, hrefArg) {
	const doc = docArg || document;
	const href = hrefArg || window.location.href;
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

	const sourceHref = String(href || "");
	let url;
	try {
		url = new URL(sourceHref);
	} catch (error) {
		url = new URL("https://www.upwork.com/");
	}
	const origin = url.origin && url.origin !== "null" ? url.origin : "";
	data.link = origin ? `${origin}${url.pathname}` : url.href;
	const jobMatch = sourceHref.match(/~([0-9]+)/);
	data.jobId = jobMatch ? jobMatch[1] : "";
	data.date = new Date().toLocaleDateString("en-US", {
		month: "short",
		day: "numeric",
		year: "numeric",
	});

	const postedLine =
		doc.querySelector(".posted-on-line") ||
		doc.querySelector("[data-test='job-details'] .posted-on-line");
	if (postedLine) {
		const postedText = getText(
			postedLine.querySelector(".text-light-on-muted")
		);
		if (postedText) {
			data.jobCreatedSince = postedText;
		}
	}

	const clientLocation =
		doc.querySelector("[data-qa='client-location'] strong") ||
		doc.querySelector(
			".cfe-ui-job-about-client [data-qa='client-location'] strong"
		);
	const locationText = getText(clientLocation);
	if (locationText) {
		data.country = locationText;
	}

	const paymentText = normalize(doc.body ? doc.body.textContent : "");
	if (paymentText.toLowerCase().includes("payment method verified")) {
		data.payment = "Verified";
	} else if (
		paymentText.toLowerCase().includes("payment method not verified")
	) {
		data.payment = "Not verified";
	}

	const jobDetails =
		doc.querySelector(".job-details") ||
		doc.querySelector("[data-ev-sublocation='jobdetails']") ||
		doc.querySelector("[data-test='job-details']");

	const headingCandidates = [];
	if (jobDetails) {
		headingCandidates.push(...jobDetails.querySelectorAll("h1, h2, h3, h4"));
	}
	headingCandidates.push(...doc.querySelectorAll("h1, h2, h3, h4"));

	for (const heading of headingCandidates) {
		const text = getText(heading);
		if (text && text.toLowerCase() !== "activity on this job") {
			data.name = text;
			break;
		}
	}

	if (!data.name) {
		const title = normalize(doc.title);
		data.name = title.replace(/\s*-\s*Upwork.*$/i, "").trim();
	}

	const headings = Array.from(
		doc.querySelectorAll("h1, h2, h3, h4, h5, h6")
	);
	const activityHeading = headings.find(
		(heading) => getText(heading).toLowerCase() === "activity on this job"
	);
	const activitySection = activityHeading
		? activityHeading.closest("section") || activityHeading.parentElement
		: null;

	const items = activitySection
		? activitySection.querySelectorAll(".ca-item, li")
		: doc.querySelectorAll(".ca-item");

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
		Array.from(doc.querySelectorAll("a")).find((anchor) => {
			const hrefValue = anchor.getAttribute("href") || "";
			if (!/\/(ab\/proposals|nx\/proposals)\/\d+/.test(hrefValue)) {
				return false;
			}
			return getText(anchor).toLowerCase().includes("view proposal");
		}) ||
		doc.querySelector(
			"a[href*='/ab/proposals/'], a[href*='/nx/proposals/']"
		);

	const proposalHref = proposalLink?.getAttribute("href") || "";
	const proposalMatch = proposalHref.match(
		/\/(ab\/proposals|nx\/proposals)\/(\d+)/
	);
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
};

window.extractUpworkJobDataFromDocument = function extractUpworkJobDataFromDocument(
	doc,
	href
) {
	return window.extractUpworkJobData(doc, href);
};

window.extractUpworkViewedProposals = function extractUpworkViewedProposals(
	docArg
) {
	const doc = docArg || document;
	const normalize = (value) =>
		String(value || "")
			.replace(/\s+/g, " ")
			.trim();

	const results = [];
	const rows = Array.from(doc.querySelectorAll("tr.details-row"));

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
};

window.extractUpworkViewedProposalsFromDocument =
	function extractUpworkViewedProposalsFromDocument(doc) {
		return window.extractUpworkViewedProposals(doc);
	};

window.extractUpworkConnectsHistory = function extractUpworkConnectsHistory(
	docArg,
	originArg
) {
	const doc = docArg || document;
	const origin = originArg || window.location.origin;
	const normalize = (value) =>
		String(value || "")
			.replace(/\s+/g, " ")
			.trim();
	const normalizeLink = (href) => {
		if (!href) {
			return "";
		}
		try {
			const base = origin || "https://www.upwork.com";
			return new URL(href, base).toString();
		} catch (error) {
			return href;
		}
	};

	const parseConnectsValue = (value) => {
		const match = String(value || "").replace(/,/g, "").match(/[+-]?\d+/);
		if (!match) {
			return 0;
		}
		const parsed = Number(match[0]);
		return Number.isFinite(parsed) ? parsed : 0;
	};

	const formatDate = (value) => {
		const text = normalize(value);
		if (!text) {
			return "";
		}
		const lower = text.toLowerCase();
		const now = new Date();
		if (lower === "today") {
			return now.toLocaleDateString("en-US", {
				month: "short",
				day: "numeric",
				year: "numeric",
			});
		}
		if (lower === "yesterday") {
			const yesterday = new Date(now);
			yesterday.setDate(now.getDate() - 1);
			return yesterday.toLocaleDateString("en-US", {
				month: "short",
				day: "numeric",
				year: "numeric",
			});
		}
		const parsed = new Date(text);
		if (Number.isNaN(parsed.getTime())) {
			return text;
		}
		return parsed.toLocaleDateString("en-US", {
			month: "short",
			day: "numeric",
			year: "numeric",
		});
	};

	const findConnectsTable = () => {
		const tables = Array.from(doc.querySelectorAll("table")).filter(
			(table) =>
				!table.closest("#fixture-connects-history") &&
				!table.closest("[style*='display: none']")
		);
		const ariaCandidates = tables.filter((table) =>
			String(table.getAttribute("aria-label") || "")
				.toLowerCase()
				.includes("connects")
		);
		if (ariaCandidates.length) {
			return ariaCandidates.reduce((best, table) => {
				if (!best) {
					return table;
				}
				const bestRows = best.querySelectorAll("tbody tr").length;
				const nextRows = table.querySelectorAll("tbody tr").length;
				return nextRows > bestRows ? table : best;
			}, null);
		}
		for (const table of tables) {
			const headerCells = Array.from(
				table.querySelectorAll("thead th, thead td")
			);
			const headerText = headerCells.map((cell) =>
				normalize(cell.textContent || "")
			);
			if (
				headerText.includes("Date") &&
				headerText.includes("Action") &&
				headerText.includes("Connects")
			) {
				return table;
			}
		}
		return null;
	};

	const connectsTable = findConnectsTable();
	let rows = [];
	if (connectsTable) {
		rows = Array.from(connectsTable.querySelectorAll("tbody tr"));
	} else {
		rows = Array.from(
			doc.querySelectorAll("#connects-history-table tbody tr")
		).filter((row) => !row.closest("#fixture-connects-history"));
	}
	const entries = [];

	for (const row of rows) {
		const cells = row.querySelectorAll("td");
		if (!cells.length) {
			continue;
		}
		const dateText = formatDate(cells[0]?.textContent || "");
		const actionCell = cells[1];
		const connectsCell = cells[cells.length - 1];
		const link = actionCell?.querySelector("a[href*='/jobs/']") || null;
		const href = link?.getAttribute("href") || "";
		const linkText = normalize(link?.textContent || "");
		let actionText = normalize(
			actionCell?.querySelector("small")?.textContent ||
				actionCell?.querySelector("span")?.textContent ||
				actionCell?.textContent ||
				""
		);
		if (linkText && actionText.includes(linkText)) {
			actionText = normalize(actionText.replace(linkText, ""));
		}
		const isBoosted = /\bboost/i.test(actionText);
		const jobLink = normalizeLink(href);
		const match = href.match(/~(\d+)/);
		const jobId = match ? match[1] : "";
		const title = normalize(
			actionCell?.querySelector("strong")?.textContent || link?.textContent
		);
		const connectsRaw = normalize(connectsCell?.textContent || "");
		const connectsValue = parseConnectsValue(connectsRaw);
		if (!jobId) {
			continue;
		}
		const spent = connectsValue < 0 ? Math.abs(connectsValue) : 0;
		const refund = connectsValue > 0 ? connectsValue : 0;
		entries.push({
			jobId,
			title,
			link: jobLink,
			date: dateText,
			isBoosted,
			connectsSpent: isBoosted ? 0 : spent,
			connectsRefund: isBoosted ? 0 : refund,
			boostedConnectsSpent: isBoosted ? spent : 0,
			boostedConnectsRefund: isBoosted ? refund : 0,
		});
	}

	if (!entries.length) {
		return { ok: false, error: "No connects history rows found." };
	}

	return { ok: true, entries };
};

window.extractUpworkConnectsHistoryFromDocument =
	function extractUpworkConnectsHistoryFromDocument(doc, origin) {
		return window.extractUpworkConnectsHistory(doc, origin);
	};
