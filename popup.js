const readButton = document.getElementById("read-page");
const copyButton = document.getElementById("copy-csv");
const statusEl = document.getElementById("status");

const fields = {
	name: document.getElementById("field-name"),
	link: document.getElementById("field-link"),
	date: document.getElementById("field-date"),
	proposals: document.getElementById("field-proposals"),
	invitesSent: document.getElementById("field-invites"),
	interviewing: document.getElementById("field-interviewing"),
	lastViewed: document.getElementById("field-last-viewed"),
	unansweredInvites: document.getElementById("field-unanswered"),
};

let currentData = null;

const setStatus = (message, tone = "") => {
	statusEl.textContent = message;
	statusEl.className = `status ${tone}`.trim();
};

const setFields = (data) => {
	fields.name.textContent = data?.name || "-";
	fields.link.textContent = data?.link || "-";
	fields.date.textContent = data?.date || "-";
	fields.proposals.textContent = data?.proposals || "-";
	fields.invitesSent.textContent = data?.invitesSent || "-";
	fields.interviewing.textContent = data?.interviewing || "-";
	fields.lastViewed.textContent = data?.lastViewed || "-";
	fields.unansweredInvites.textContent = data?.unansweredInvites || "-";
};

const escapeCsv = (value) => {
	const text = String(value ?? "");
	if (/[",\n]/.test(text)) {
		return `"${text.replace(/"/g, '""')}"`;
	}
	return text;
};

const toCsvRow = (data) =>
	[
		data.name,
		data.link,
		data.date,
		data.proposals,
		data.invitesSent,
		data.interviewing,
	].map(escapeCsv).join(",");

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
			},
		);
	});

function extractUpworkJobData() {
	const normalize = (value) => String(value || "").replace(/\s+/g, " ").trim();
	const getText = (el) => normalize(el ? el.textContent : "");

	const data = {
		name: "",
		link: "",
		date: "",
		proposals: "",
		invitesSent: "",
		interviewing: "",
		lastViewed: "",
		unansweredInvites: "",
	};

	const url = new URL(window.location.href);
	data.link = `${url.origin}${url.pathname}`;
	data.date = new Date().toISOString().slice(0, 10);

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
		document.querySelectorAll("h1, h2, h3, h4, h5, h6"),
	);
	const activityHeading = headings.find(
		(heading) => getText(heading).toLowerCase() === "activity on this job",
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
	setFields(currentData);
	copyButton.disabled = false;
	setStatus("Captured job activity.", "success");
});

copyButton.addEventListener("click", async () => {
	if (!currentData) {
		setStatus("Nothing to copy yet.", "warn");
		return;
	}
	const csv = toCsvRow(currentData);
	try {
		await navigator.clipboard.writeText(csv);
		setStatus("CSV row copied.", "success");
	} catch (error) {
		setStatus("Clipboard access failed.", "error");
	}
});

setFields(null);
