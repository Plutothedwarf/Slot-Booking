// Paste your deployed Apps Script Web App URL here (see SETUP.md step 4).
// It looks like: https://script.google.com/macros/s/AKfycb.../exec
const API_URL = "https://script.google.com/macros/s/AKfycbyRwBbkfmD0pjrvia88Synr2JGkja2Z9IB6f6jlx1U7EpFsOPm1dVlXPermN4XRXQ0b/exec";

async function api(action, payload) {
  const url = API_URL + '?action=' + encodeURIComponent(action)
    + '&payload=' + encodeURIComponent(JSON.stringify(payload || {}))
    + '&_ts=' + Date.now();
  const res = await fetch(url, { cache: 'no-store' });
  return res.json();
}
