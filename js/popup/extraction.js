window.extractUpworkJobData = function extractUpworkJobData() {
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
		document.querySelector(
			".cfe-ui-job-about-client [data-qa='client-location'] strong"
		);
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
		document.querySelector(
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

window.extractUpworkViewedProposals = function extractUpworkViewedProposals() {
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
};

window.extractUpworkConnectsHistory = function extractUpworkConnectsHistory() {
	const normalize = (value) =>
		String(value || "")
			.replace(/\s+/g, " ")
			.trim();

	const parseConnectsValue = (value) => {
		const match = String(value || "").replace(/,/g, "").match(/[+-]?\d+/);
		if (!match) {
			return 0;
		}
		const parsed = Number(match[0]);
		return Number.isFinite(parsed) ? parsed : 0;
	};

	const currentDate = new Date().toLocaleDateString("en-US", {
		month: "short",
		day: "numeric",
		year: "numeric",
	});

	const rows = Array.from(
		document.querySelectorAll("#connects-history-table tbody tr")
	);
	const entries = [];

	for (const row of rows) {
		const cells = row.querySelectorAll("td");
		if (!cells.length) {
			continue;
		}
		const dateText = currentDate;
		const actionCell = cells[1];
		const connectsCell = cells[cells.length - 1];
		const link = actionCell?.querySelector("a[href*='/jobs/']") || null;
		const href = link?.getAttribute("href") || "";
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
		entries.push({
			jobId,
			title,
			link: href || "",
			date: dateText,
			connectsSpent: connectsValue < 0 ? Math.abs(connectsValue) : 0,
			connectsRefund: connectsValue > 0 ? connectsValue : 0,
		});
	}

	if (!entries.length) {
		return { ok: false, error: "No connects history rows found." };
	}

	return { ok: true, entries };
};
