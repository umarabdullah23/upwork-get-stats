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

const getTestingVisibility = () =>
	new Promise((resolve) => {
		chrome.storage.sync.get(["showTesting"], (result) => {
			if (chrome.runtime.lastError) {
				resolve(false);
				return;
			}
			resolve(Boolean(result.showTesting));
		});
	});

const saveTestingVisibility = (visible) =>
	new Promise((resolve) => {
		chrome.storage.sync.set({ showTesting: Boolean(visible) }, () => {
			resolve(!chrome.runtime.lastError);
		});
	});
