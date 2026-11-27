/**
 * CampusFix — Google Apps Script backend
 *
 * Architecture:
 *   CampusFix.html
 *       ↓ google.script.run
 *   Apps Script
 *       ↓ server-side POST
 *   Fastn webhook → "Deduplicate or Create Ticket" workflow
 *       ↓
 *   Google Sheets / Slack / existing Fastn backend
 *
 * IMPORTANT:
 * 1) Put CampusFix.html in this Apps Script project with the exact file name "CampusFix".
 * 2) If this Apps Script project is NOT bound to the CampusFix Google Sheet,
 *    paste that Sheet ID into SPREADSHEET_ID below.
 */

const FASTN_WEBHOOK_URL = "https://webhooks.fastn.dev/prod/triggers/personal_596c0569efaa02b4000c/webhooks/93c946fb-5c59-4836-8c5b-f982ac4a9d61";

// If this Apps Script is standalone, put your existing CampusFix spreadsheet ID here.
// Example: "1AbCdEfGhIjKlMnOpQrStUvWxYz123456789"
// If the script is bound directly to that spreadsheet, leave this blank.
const SPREADSHEET_ID = "YOUR_ID";

/**
 * Serves CampusFix.html as the Web App.
 */
function doGet() {
  // Try the current South Campus page first, then the shorter TriageDesk name,
  // and finally the old CampusFix page as a compatibility fallback.
  const candidates = ["TriageDesk-South-Campus", "TriageDesk", "CampusFix"];
  for (let i = 0; i < candidates.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(candidates[i])
        .setTitle("TriageDesk")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (e) {
      // Try the next file name.
    }
  }
  throw new Error("TriageDesk HTML file not found. Add TriageDesk-South-Campus.html to this Apps Script project.");
}

/**
 * Receives a complaint from the browser-side Apps Script HTML
 * and forwards it to the EXISTING Fastn webhook.
 */
function submitToFastn(payload) {
  if (!payload || typeof payload !== "object") {
    return {
      ok: false,
      status: 400,
      error: "Invalid complaint payload."
    };
  }

  try {
    const response = UrlFetchApp.fetch(FASTN_WEBHOOK_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const text = response.getContentText();

    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch (err) {
      data = text;
    }

    return {
      ok: status >= 200 && status < 300,
      status: status,
      data: data,
      error: status >= 200 && status < 300 ? null : (
        typeof data === "string" ? data : JSON.stringify(data || {})
      )
    };

  } catch (err) {
    return {
      ok: false,
      status: 500,
      error: err && err.message ? err.message : String(err)
    };
  }
}


/**
 * Run this once from the Apps Script editor to verify the existing Fastn webhook.
 * It sends a clearly marked test complaint through the same backend path used by the website.
 */
function testFastnConnection() {
  const testPayload = {
    campus_id: "CAMPUS_01",
    campus_name: "Main Campus",
    location: "Building A",
    category: "IT",
    description: "CampusFix backend connection test — please ignore this ticket.",
    student_email: "campusfix-test@example.com"
  };

  const result = submitToFastn(testPayload);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Reads the existing Tickets sheet for the student list and admin dashboard.
 *
 * Expected header names are matched flexibly, so small header-name differences
 * in the existing sheet do not break the dashboard.
 */
function getTickets(filter) {
  try {
    const sheet = getTicketsSheet_();
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];

    const headers = values[0].map(function (h) {
      return String(h || "").trim();
    });

    const tickets = values.slice(1)
      .filter(function (row) {
        return row.some(function (cell) {
          return String(cell || "").trim() !== "";
        });
      })
      .map(function (row) {
        return rowToTicket_(headers, row);
      });

    // Student view: only return that student's tickets.
    if (filter && filter.email) {
      const email = String(filter.email).trim().toLowerCase();
      return tickets.filter(function (ticket) {
        return String(ticket.student_email || "").trim().toLowerCase() === email;
      });
    }

    // Admin view: all tickets.
    return tickets.reverse();

  } catch (err) {
    throw new Error(
      "Could not read the Tickets sheet. " +
      (err && err.message ? err.message : String(err))
    );
  }
}

/**
 * Finds the existing Tickets sheet.
 */
function getTicketsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheet = ss.getSheetByName("Tickets");
  if (!sheet) {
    const names = ss.getSheets().map(function(s) { return s.getName(); });
    throw new Error(
      'The spreadsheet was opened, but the "Tickets" tab was not found. Available tabs: ' +
      names.join(", ")
    );
  }

  return sheet;
}

/**
 * Diagnostic test for the Admin Dashboard.
 * Run this ONCE from the Apps Script editor.
 * It verifies the exact spreadsheet and Tickets tab used by the dashboard.
 */
function testDashboardConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Tickets");

  if (!sheet) {
    const names = ss.getSheets().map(function(s) { return s.getName(); });
    throw new Error(
      'Dashboard connection failed: "Tickets" tab not found. Available tabs: ' +
      names.join(", ")
    );
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetName: ss.getName(),
    sheetName: sheet.getName(),
    rows: values.length,
    columns: values.length ? values[0].length : 0,
    headers: values.length ? values[0] : []
  };
}

/**
 * Returns a useful error object instead of hiding the real Google Sheets error.
 */
function getDashboardData() {
  try {
    return {
      ok: true,
      tickets: getTickets({})
    };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err)
    };
  }
}

/**
 * Converts a sheet row into the object expected by CampusFix.html.
 */
function rowToTicket_(headers, row) {
  const out = {};

  headers.forEach(function (header, i) {
    const key = normalizeHeader_(header);
    if (key) {
      out[key] = formatCell_(row[i]);
    }
  });

  // Common aliases used by different sheet header styles.
  out.ticket_id = firstValue_(out.ticket_id, out.ticketid, out.id);
  out.campus_id = firstValue_(out.campus_id, out.campusid);
  out.campus_name = firstValue_(out.campus_name, out.campusname, out.campus);
  out.student_email = firstValue_(
    out.student_email,
    out.studentemail,
    out.email
  );
  out.created_at = firstValue_(out.created_at, out.created, out.timestamp);
  out.updated_at = firstValue_(out.updated_at, out.updated);
  out.status = firstValue_(out.status, "Open");
  out.action = firstValue_(out.action, out.result_action);

  return out;
}

/**
 * Turns sheet headers like "Student Email" into "student_email".
 */
function normalizeHeader_(header) {
  return String(header || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function firstValue_() {
  for (let i = 0; i < arguments.length; i++) {
    const value = arguments[i];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return value;
    }
  }
  return "";
}

function formatCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || "Asia/Karachi",
      "yyyy-MM-dd HH:mm:ss"
    );
  }
  return value;
}


/**
 * CampusFix — Google Apps Script backend
 *
 * Architecture:
 *   CampusFix.html
 *       ↓ google.script.run
 *   Apps Script
 *       ↓ server-side POST
 *   Fastn webhook → "Deduplicate or Create Ticket" workflow
 *       ↓
 *   Google Sheets / Slack / existing Fastn backend
 *
 * IMPORTANT:
 * 1) Put CampusFix.html in this Apps Script project with the exact file name "CampusFix".
 * 2) If this Apps Script project is NOT bound to the CampusFix Google Sheet,
 *    paste that Sheet ID into SPREADSHEET_ID below.
 */

const FASTN_WEBHOOK_URL = "https://webhooks.fastn.dev/prod/triggers/personal_596c0569efaa02b4000c/webhooks/93c946fb-5c59-4836-8c5b-f982ac4a9d61";

// If this Apps Script is standalone, put your existing CampusFix spreadsheet ID here.
// Example: "1AbCdEfGhIjKlMnOpQrStUvWxYz123456789"
// If the script is bound directly to that spreadsheet, leave this blank.
const SPREADSHEET_ID = "YOUR_ID";

/**
 * Serves CampusFix.html as the Web App.
 */
function doGet() {
  // Try the current South Campus page first, then the shorter TriageDesk name,
  // and finally the old CampusFix page as a compatibility fallback.
  const candidates = ["TriageDesk-South-Campus", "TriageDesk", "CampusFix"];
  for (let i = 0; i < candidates.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(candidates[i])
        .setTitle("TriageDesk")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (e) {
      // Try the next file name.
    }
  }
  throw new Error("TriageDesk HTML file not found. Add TriageDesk-South-Campus.html to this Apps Script project.");
}

/**
 * Receives a complaint from the browser-side Apps Script HTML
 * and forwards it to the EXISTING Fastn webhook.
 */
function submitToFastn(payload) {
  if (!payload || typeof payload !== "object") {
    return {
      ok: false,
      status: 400,
      error: "Invalid complaint payload."
    };
  }

  try {
    const response = UrlFetchApp.fetch(FASTN_WEBHOOK_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const text = response.getContentText();

    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch (err) {
      data = text;
    }

    return {
      ok: status >= 200 && status < 300,
      status: status,
      data: data,
      error: status >= 200 && status < 300 ? null : (
        typeof data === "string" ? data : JSON.stringify(data || {})
      )
    };

  } catch (err) {
    return {
      ok: false,
      status: 500,
      error: err && err.message ? err.message : String(err)
    };
  }
}


/**
 * Run this once from the Apps Script editor to verify the existing Fastn webhook.
 * It sends a clearly marked test complaint through the same backend path used by the website.
 */
function testFastnConnection() {
  const testPayload = {
    campus_id: "CAMPUS_01",
    campus_name: "Main Campus",
    location: "Building A",
    category: "IT",
    description: "CampusFix backend connection test — please ignore this ticket.",
    student_email: "campusfix-test@example.com"
  };

  const result = submitToFastn(testPayload);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Reads the existing Tickets sheet for the student list and admin dashboard.
 *
 * Expected header names are matched flexibly, so small header-name differences
 * in the existing sheet do not break the dashboard.
 */
function getTickets(filter) {
  try {
    const sheet = getTicketsSheet_();
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];

    const headers = values[0].map(function (h) {
      return String(h || "").trim();
    });

    const tickets = values.slice(1)
      .filter(function (row) {
        return row.some(function (cell) {
          return String(cell || "").trim() !== "";
        });
      })
      .map(function (row) {
        return rowToTicket_(headers, row);
      });

    // Student view: only return that student's tickets.
    if (filter && filter.email) {
      const email = String(filter.email).trim().toLowerCase();
      return tickets.filter(function (ticket) {
        return String(ticket.student_email || "").trim().toLowerCase() === email;
      });
    }

    // Admin view: all tickets.
    return tickets.reverse();

  } catch (err) {
    throw new Error(
      "Could not read the Tickets sheet. " +
      (err && err.message ? err.message : String(err))
    );
  }
}

/**
 * Finds the existing Tickets sheet.
 */
function getTicketsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheet = ss.getSheetByName("Tickets");
  if (!sheet) {
    const names = ss.getSheets().map(function(s) { return s.getName(); });
    throw new Error(
      'The spreadsheet was opened, but the "Tickets" tab was not found. Available tabs: ' +
      names.join(", ")
    );
  }

  return sheet;
}

/**
 * Diagnostic test for the Admin Dashboard.
 * Run this ONCE from the Apps Script editor.
 * It verifies the exact spreadsheet and Tickets tab used by the dashboard.
 */
function testDashboardConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Tickets");

  if (!sheet) {
    const names = ss.getSheets().map(function(s) { return s.getName(); });
    throw new Error(
      'Dashboard connection failed: "Tickets" tab not found. Available tabs: ' +
      names.join(", ")
    );
  }

  const range = sheet.getDataRange();
  const values = range.getValues();

  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    spreadsheetName: ss.getName(),
    sheetName: sheet.getName(),
    rows: values.length,
    columns: values.length ? values[0].length : 0,
    headers: values.length ? values[0] : []
  };
}

/**
 * Returns a useful error object instead of hiding the real Google Sheets error.
 */
function getDashboardData() {
  try {
    return {
      ok: true,
      tickets: getTickets({})
    };
  } catch (err) {
    return {
      ok: false,
      error: err && err.message ? err.message : String(err)
    };
  }
}

/**
 * Converts a sheet row into the object expected by CampusFix.html.
 */
function rowToTicket_(headers, row) {
  const out = {};

  headers.forEach(function (header, i) {
    const key = normalizeHeader_(header);
    if (key) {
      out[key] = formatCell_(row[i]);
    }
  });

  // Common aliases used by different sheet header styles.
  out.ticket_id = firstValue_(out.ticket_id, out.ticketid, out.id);
  out.campus_id = firstValue_(out.campus_id, out.campusid);
  out.campus_name = firstValue_(out.campus_name, out.campusname, out.campus);
  out.student_email = firstValue_(
    out.student_email,
    out.studentemail,
    out.email
  );
  out.created_at = firstValue_(out.created_at, out.created, out.timestamp);
  out.updated_at = firstValue_(out.updated_at, out.updated);
  out.status = firstValue_(out.status, "Open");
  out.action = firstValue_(out.action, out.result_action);

  return out;
}

/**
 * Turns sheet headers like "Student Email" into "student_email".
 */
function normalizeHeader_(header) {
  return String(header || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function firstValue_() {
  for (let i = 0; i < arguments.length; i++) {
    const value = arguments[i];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return value;
    }
  }
  return "";
}

function formatCell_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || "Asia/Karachi",
      "yyyy-MM-dd HH:mm:ss"
    );
  }
  return value;
}
