/**
 * C-SDTP External Metrics API  ->  Google Sheet (tab: API_CSDTP)
 *
 * SETUP (one time)
 * 1. Project Settings (gear icon) > Script properties > add:
 *      API_EMAIL     = the API account email
 *      API_PASSWORD  = the API account password
 *    (Never type credentials into this file.)
 * 2. Run installTrigger() once and approve the permissions.
 * 3. Run syncCsdtp() once manually to test. Check the API_CSDTP tab.
 */

const CONFIG = {
  BASE_URL: 'https://lfmes.mlf.go.tz',
  API_PREFIX: '/api/external/csdtp/v1',
  LOGIN_PATH: '/api/login',
  DEVICE_NAME: 'C-SDTP',
  SHEET_NAME: 'API_CSDTP',
  SOURCE: 'CSDTP',
  REFRESH_MINUTES: 15 // allowed values: 1, 5, 10, 15, 30
};

/* ---------- 1. MAIN SYNC ---------- */

function syncCsdtp() {
  const response = apiGet_('/metrics/summary'); // one call returns all 11 metrics
  if (!response || !response.data) {
    throw new Error('Unexpected API response: ' + JSON.stringify(response).slice(0, 300));
  }

  const now = Utilities.formatDate(new Date(), 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'"); // ISO text, read by the dashboard
  const flat = {};
  flatten_(response.data, '', flat);

  // Long format: source | group | metric | value | updated
  const rows = Object.keys(flat).map(function (key) {
    return [CONFIG.SOURCE, key.split('.')[0], key, flat[key], now];
  });

  writeSheet_(rows);
  Logger.log('Synced ' + rows.length + ' metrics at ' + now);
}

/* ---------- 2. AUTH + HTTP ---------- */

function getToken_(forceNew) {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('API_TOKEN');
  if (token && !forceNew) return token;

  const email = props.getProperty('API_EMAIL');
  const password = props.getProperty('API_PASSWORD');
  if (!email || !password) {
    throw new Error('Set API_EMAIL and API_PASSWORD in Script properties.');
  }

  const res = UrlFetchApp.fetch(CONFIG.BASE_URL + CONFIG.LOGIN_PATH, {
    method: 'post',
    contentType: 'application/json',
    headers: { Accept: 'application/json' },
    payload: JSON.stringify({ email: email, password: password, device_name: CONFIG.DEVICE_NAME }),
    muteHttpExceptions: true
  });

  const body = JSON.parse(res.getContentText() || '{}');
  // The documentation does not name the token field, so try the common ones.
  token = body.token || body.access_token ||
          (body.data && (body.data.token || body.data.access_token));

  if (res.getResponseCode() !== 200 || !token) {
    throw new Error('Login failed (' + res.getResponseCode() + '): ' +
                    res.getContentText().slice(0, 200));
  }

  props.setProperty('API_TOKEN', token); // reuse until it stops working
  return token;
}

function apiGet_(path, retried) {
  const res = UrlFetchApp.fetch(CONFIG.BASE_URL + CONFIG.API_PREFIX + path, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + getToken_(false), Accept: 'application/json' },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() === 401 && !retried) {
    getToken_(true); // token expired or revoked: log in again once
    return apiGet_(path, true);
  }
  if (res.getResponseCode() !== 200) {
    throw new Error('API ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  return JSON.parse(res.getContentText());
}

/* ---------- 3. FLATTEN + WRITE ---------- */

// Turns nested JSON into dotted keys, e.g. disbursement.expenditure
// AWPB years become awpb_utilization_by_year.years.2024-2025.budget_amount
function flatten_(obj, prefix, out) {
  Object.keys(obj).forEach(function (k) {
    if (k === 'filters' || k === 'financial_year') return;
    const v = obj[k];
    const key = prefix ? prefix + '.' + k : k;

    if (Array.isArray(v)) {
      v.forEach(function (item, i) {
        const id = item && item.financial_year ? String(item.financial_year).replace('/', '-') : i;
        flatten_(item, key + '.' + id, out);
      });
    } else if (v !== null && typeof v === 'object') {
      flatten_(v, key, out);
    } else {
      out[key] = v;
    }
  });
}

function writeSheet_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  const header = [['source', 'group', 'metric', 'value', 'updated']];

  sheet.clearContents(); // only reached if the API call succeeded, so old data survives failures
  sheet.getRange(1, 1, 1, 5).setValues(header).setFontWeight('bold');
  if (rows.length) {
    sheet.getRange(2, 5, rows.length, 1).setNumberFormat('@'); // keep the timestamp as plain text
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

/* ---------- 4. TRIGGER ---------- */

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncCsdtp') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncCsdtp').timeBased().everyMinutes(CONFIG.REFRESH_MINUTES).create();
}
