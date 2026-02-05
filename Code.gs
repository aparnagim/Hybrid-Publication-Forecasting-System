/******************** CONFIG ********************/
const SPREADSHEET_ID = "1Kj3vDGxZZsjd-zBqMWTSR_xC-eeuTmlvKK6MWIL3pBg";
const SHEET_NAME     = "Raw_Data";

//  Sheet that contains field forecasts / field list 
const ASJC_SHEET_NAME = "Outputs_Field_Forecast";

// Column names 
const COL_YEAR   = "Year";
const COL_PUB    = "Publication type";
const COL_OA     = "Open Access";
const COL_LANG   = "Language";

//  ngrok endpoint 
const AI_API_URL = "https://sizy-merilyn-bombastic.ngrok-free.dev/predict";
/***********************************************/

/**  Web App entry point */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : "dashboard";
  const file = (page === "predict") ? "Prediction" : "Index";

  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle("Scopus Analytics")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** optional include helper */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Read sheet once */
function readSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_NAME}`);
  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) throw new Error("No data found in sheet.");
  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);
  return { headers, rows };
}

/** Build header -> index map */
function headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[String(h).trim()] = i);
  return map;
}

/** Filter rows by dropdown values */
function filterRows_(headers, rows, filters) {
  const map = headerMap_(headers);

  const idxYear = map[COL_YEAR];
  const idxPub  = map[COL_PUB];
  const idxOA   = map[COL_OA];
  const idxLang = map[COL_LANG];

  const year = (filters && filters.year) ? String(filters.year) : "ALL";
  const pub  = (filters && filters.pubType) ? String(filters.pubType) : "ALL";
  const oa   = (filters && filters.oa) ? String(filters.oa) : "ALL";
  const lang = (filters && filters.lang) ? String(filters.lang) : "ALL";

  return rows.filter(r => {
    if (idxYear != null && idxYear > -1 && year !== "ALL" && String(r[idxYear]) !== year) return false;
    if (idxPub  != null && idxPub  > -1 && pub  !== "ALL" && String(r[idxPub]  || "") !== pub)  return false;
    if (idxOA   != null && idxOA   > -1 && oa   !== "ALL" && String(r[idxOA]   || "") !== oa)   return false;
    if (idxLang != null && idxLang > -1 && lang !== "ALL" && String(r[idxLang] || "") !== lang) return false;
    return true;
  });
}

/**  Table API: returns filtered rows with limit (default 10) */
function getRawDataFiltered(filters, limit = 10) {
  const { headers, rows } = readSheet_();
  const filtered = filterRows_(headers, rows, filters);
  const preview = filtered.slice(0, Math.min(limit, filtered.length));
  return { headers, rows: preview, total: filtered.length };
}

/**  Dropdown options API */
function getFilterOptions() {
  const { headers, rows } = readSheet_();
  const map = headerMap_(headers);

  function uniq(colName) {
    const idx = map[colName];
    if (idx == null || idx < 0) return [];
    const set = new Set();
    rows.forEach(r => {
      const v = r[idx];
      if (v !== "" && v !== null && v !== undefined) set.add(String(v).trim());
    });
    return Array.from(set).sort();
  }

  return {
    years: uniq(COL_YEAR),
    pubTypes: uniq(COL_PUB),
    openAccess: uniq(COL_OA),
    languages: uniq(COL_LANG)
  };
}

/**  NEW: ASJC Field dropdown list (includes "All") */
function getAsjcFieldOptions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(ASJC_SHEET_NAME);
  if (!sh) {
    // fallback if sheet not found
    return ["All"];
  }

  const values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return ["All"];

  const headers = values[0].map(h => String(h).trim());
  const idx = headers.indexOf("ASJC_Field");
  if (idx < 0) return ["All"];

  const set = new Set();
  for (let i = 1; i < values.length; i++) {
    const v = values[i][idx];
    if (v !== "" && v !== null && v !== undefined) set.add(String(v).trim());
  }

  const fields = Array.from(set).sort();
  return ["All"].concat(fields);
}

/**  Charts API */
function getChartSummary(filters) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${SHEET_NAME}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { byYear:{}, byPub:{}, byOA:{}, byLang:{}, total:0 };

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const map = {};
  headers.forEach((h, i) => map[h] = i);

  const iYear = map[COL_YEAR];
  const iPub  = map[COL_PUB];
  const iOA   = map[COL_OA];
  const iLang = map[COL_LANG];

  const allRows = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const yearF = filters?.year || "ALL";
  const pubF  = filters?.pubType || "ALL";
  const oaF   = filters?.oa || "ALL";
  const langF = filters?.lang || "ALL";

  const byYear = {}, byPub = {}, byOA = {}, byLang = {};
  let total = 0;

  allRows.forEach(r => {
    const year = iYear >= 0 ? String(r[iYear] ?? "") : "";
    const pub  = iPub  >= 0 ? String(r[iPub]  ?? "Unknown") : "Unknown";
    const oa   = iOA   >= 0 ? String(r[iOA]   ?? "Unknown") : "Unknown";
    const lang = iLang >= 0 ? String(r[iLang] ?? "Unknown") : "Unknown";

    if (yearF !== "ALL" && year !== String(yearF)) return;
    if (pubF  !== "ALL" && pub  !== String(pubF))  return;
    if (oaF   !== "ALL" && oa   !== String(oaF))   return;
    if (langF !== "ALL" && lang !== String(langF)) return;

    total++;
    if (year) byYear[year] = (byYear[year] || 0) + 1;
    byPub[pub]   = (byPub[pub]   || 0) + 1;
    byOA[oa]     = (byOA[oa]     || 0) + 1;
    byLang[lang] = (byLang[lang] || 0) + 1;
  });

  return { byYear, byPub, byOA, byLang, total };
}

/**  CSV download */
function exportCsv(filters) {
  const { headers, rows } = readSheet_();
  const filtered = filterRows_(headers, rows, filters);

  const escapeCsv = (v) => {
    const s = (v === null || v === undefined) ? "" : String(v);
    return `"${s.replace(/"/g, '""')}"`;
  };

  const lines = [];
  lines.push(headers.map(escapeCsv).join(","));
  filtered.forEach(r => lines.push(r.map(escapeCsv).join(",")));

  const csv = lines.join("\n");
  const blob = Utilities.newBlob(csv, "text/csv", "scopus_filtered.csv");

  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getDownloadUrl();
}

/**  Real-time AI Prediction: calls Colab (ngrok)
 *  Python side will handle asjc_field="All"
 */
function getRealtimePrediction(input) {
  const res = UrlFetchApp.fetch(AI_API_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(input),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error("AI API error: " + res.getContentText());
  }

  return JSON.parse(res.getContentText());
}


/**  ASJC fields for Prediction dropdown (from Outputs_Field_Forecast sheet) */
function getAsjcFieldOptions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName("Outputs_Field_Forecast"); // <-- must exist
  if (!sh) throw new Error("Sheet not found: Outputs_Field_Forecast");

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const idx = headers.indexOf("ASJC_Field");
  if (idx < 0) throw new Error("ASJC_Field column not found in Outputs_Field_Forecast");

  const set = new Set();
  for (let i = 1; i < values.length; i++) {
    const v = values[i][idx];
    if (v) set.add(String(v).trim());
  }

  const fields = Array.from(set).sort();
  fields.unshift("All"); //  add All on top
  return fields;
}