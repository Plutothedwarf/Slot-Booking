// Paste your deployed Apps Script Web App URL here (see SETUP.md step 4).
// It looks like: https://script.google.com/macros/s/AKfycb.../exec
const API_URL = "https://script.google.com/macros/s/AKfycbyRwBbkfmD0pjrvia88Synr2JGkja2Z9IB6f6jlx1U7EpFsOPm1dVlXPermN4XRXQ0b/exec";

async function api(action, payload) {
  const url = API_URL + '?action=' + encodeURIComponent(action)
    + '&payload=' + encodeURIComponent(JSON.stringify(payload || {}))
    + '&_ts=' + Date.now();
  let res;
  try {
    res = await fetch(url, { cache: 'no-store' });
  } catch (e) {
    throw new Error('Could not reach the server. Check your internet connection and the API_URL in config.js.');
  }
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error('The backend sent back an error page instead of data. In Apps Script: run setup() once and approve permissions, then Deploy > Manage deployments > edit > New version > Deploy.');
  }
}

// Digits only; drops a leading 91 or 0 on a longer number. Valid = exactly 10 digits.
function cleanMobile(s) {
  let d = String(s || '').replace(/\D/g, '');
  if (d.length > 10 && (d.indexOf('91') === 0 || d.indexOf('0') === 0)) d = d.slice(-10);
  return d;
}
