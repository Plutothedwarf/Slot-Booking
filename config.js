// Paste your deployed Apps Script Web App URL here (see SETUP.md step 4).
// It looks like: https://script.google.com/macros/s/AKfycb.../exec
const API_URL = "PASTE_YOUR_APPS_SCRIPT_URL_HERE";

async function api(action, payload) {
  const url = API_URL + '?action=' + encodeURIComponent(action) + '&payload=' + encodeURIComponent(JSON.stringify(payload || {}));
  const res = await fetch(url);
  return res.json();
}
