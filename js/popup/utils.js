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
		if (
			/\/ConnectsHistory\.html$/i.test(path) ||
			/\/test_data\/ConnectsHistory\.html$/i.test(path)
		) {
			return "connects";
		}
		if (parsed.hostname.endsWith("upwork.com")) {
			if (path.startsWith("/nx/proposals")) {
				return "proposals";
			}
			if (path.startsWith("/nx/plans/connects/history")) {
				return "connects";
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
