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
				func: window.extractUpworkJobData,
			},
			(results) => {
				if (chrome.runtime.lastError) {
					resolve({ ok: false, error: chrome.runtime.lastError.message });
					return;
				}
				resolve(results && results[0] ? results[0].result : { ok: false });
			}
		);
	});

const executeProposalScan = (tabId) =>
	new Promise((resolve) => {
		chrome.scripting.executeScript(
			{
				target: { tabId },
				func: window.extractUpworkViewedProposals,
			},
			(results) => {
				if (chrome.runtime.lastError) {
					resolve({ ok: false, error: chrome.runtime.lastError.message });
					return;
				}
				resolve(results && results[0] ? results[0].result : { ok: false });
			}
		);
	});

const executeConnectsHistoryScan = (tabId) =>
	new Promise((resolve) => {
		chrome.scripting.executeScript(
			{
				target: { tabId },
				func: window.extractUpworkConnectsHistory,
			},
			(results) => {
				if (chrome.runtime.lastError) {
					resolve({ ok: false, error: chrome.runtime.lastError.message });
					return;
				}
				resolve(results && results[0] ? results[0].result : { ok: false });
			}
		);
	});
