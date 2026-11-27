/**
 * TriageDesk / CampusFix — Google Apps Script backend (hardened)
 *
 * Architecture:
 *   HTML page -> google.script.run -> Apps Script -> Fastn webhook -> Sheets / Slack
 *
 * PRIVATE VALUES (do NOT hard-code real values if you share this file):
 *   Project Settings -> Script Properties -> add these 3 properties:
 *     FASTN_WEBHOOK_URL  = your real Fastn webhook URL
 *     SPREADSHEET_ID     = your real Google Sheet ID
 *     ADMIN_EMAILS       = comma-separated admin emails (e.g. a@x.com,b@x.com)
 *
 * The placeholders below are only fallbacks. Replace them ONLY if you are not
 * using Script Properties, and never share the file with real values in it.
 */

const PLACEHOLDERS = {
  FASTN_WEBHOOK_URL: "PASTE_YOUR_FASTN_WEBHOOK_URL_HERE",
  SPREADSHEET_ID: "PASTE_YOUR_SPREADSHEET_ID_HERE",
  ADMIN_EMAILS: "PASTE_ADMIN_EMAIL_1_HERE,PASTE_ADMIN_EMAIL_2_HERE"
};

// Only these fields are forwarded to Fastn. Add more here if your form sends them.
const ALLOWED_FIELDS = [
  "campus_id",
  "campus_name",
  "location",
  "category",
  "description",
  "student_email"
];

const MAX_FIELD_LENGTH = 2000;

/* ------------------------------------------------------------------ */
/* Config + auth helpers                                               */
/* ------------------------------------------------------------------ */

function getConfig_(key) {
  const fromProps = PropertiesService.getScriptProperties().getProperty(key);
  const value = (fromProps || PLACEHOLDERS[key] || "").trim();
  if (!value || value.indexOf("PASTE_") === 0) {
    throw new Error(
      "Missing configuration: " + key +
      ". Add it in Project Settings -> Script Properties."
    );
  }
  return value;
}

function getUserEmail_() {
  // Reliable when the web app is deployed with "Execute as: User accessing
  // the web app" (or inside the same Google Workspace domain).
  let email = "";
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) {
    try { email = Session.getEffectiveUser().getEmail(); } catch (e) {}
  }
  return String(email || "").trim().toLowerCase();
}

function getAdminList_() {
  const raw = PropertiesService.getScriptProperties().getProperty("ADMIN_EMAILS") ||
              PLACEHOLDERS.ADMIN_EMAILS;
  return raw.split(",")
    .map(function (e) { return e.trim().toLowerCase(); })
    .filter(function (e) { return e && e.indexOf("paste_") !== 0; });
}

function isAdmin_() {
  const email = getUserEmail_();
  return !!email && getAdminList_().indexOf(email) !== -1;
}

function assertAdmin_() {
  if (!isAdmin_()) {
    throw new Error("Not authorized.");
  }
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || ""));
}

/* ------------------------------------------------------------------ */
/* Web app                                                             */
/* ------------------------------------------------------------------ */

function doGet() {
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
 * Lets the page know who is logged in and whether they are an admin.
 * Optional: call from the HTML to show/hide the dashboard.
 */
function getSession() {
  return { email: getUserEmail_(), isAdmin: isAdmin_() };
}

/* ------------------------------------------------------------------ */
/* Submit a complaint                                                  */
/* ------------------------------------------------------------------ */

function submitToFastn(payload) {
  if (!payload || typeof payload !== "object") {
    return { ok: false, status: 400, error: "Invalid complaint payload." };
  }

  // Keep only allowed fields, as trimmed strings with a length cap.
  const clean = {};
  ALLOWED_FIELDS.forEach(function (key) {
    if (payload[key] !== undefined && payload[key] !== null) {
      clean[key] = String(payload[key]).trim().substring(0, MAX_FIELD_LENGTH);
    }
  });

  if (!clean.description || !clean.category) {
    return { ok: false, status: 400, error: "Category and description are required." };
  }

  // Use the verified login email when available, so students can't submit as someone else.
  const verifiedEmail = getUserEmail_();
  if (verifiedEmail) {
    clean.student_email = verifiedEmail;
  } else if (clean.student_email && !isValidEmail_(clean.student_email)) {
    return { ok: false, status: 400, error: "Invalid email address." };
  }

  try {
    const response = UrlFetchApp.fetch(getConfig_("FASTN_WEBHOOK_URL"), {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(clean),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const text = response.getContentText();

    let data = null;
    try { data = text ? JSON.parse(text) : null; } catch (err) { data = text; }

    const ok = status >= 200 && status < 300;
    if (!ok) {
      // Log details privately; don't send internals back to the browser.
      Logger.log("Fastn error " + status + ": " + text);
    }

    return {
      ok: ok,
      status: status,
      data: ok ? data : null,
      error: ok ? null : "The ticket service rejected the request. Please try again."
    };

  } catch (err) {
    Logger.log("submitToFastn failed: " + (err && err.message ? err.message : err));
    return { ok: false, status: 500, error: "Could not reach the ticket service." };
  }
}

/* ------------------------------------------------------------------ */
/* Read tickets                                                        */
/* ------------------------------------------------------------------ */

/**
 * - Student: always limited to the logged-in user's own email.
 * - Admin: can read all tickets (filter.email optional).
 * - Anyone else asking for all tickets: rejected.
 */
function getTickets(filter) {
  const userEmail = getUserEmail_();
  const admin = isAdmin_();
  const requested = filter && filter.email
    ? String(filter.email).trim().toLowerCase()
    : "";

  let emailToFilter = "";
  if (admin) {
    emailToFilter = requested; // blank = all tickets
  } else {
    if (!userEmail) {
      throw new Error("Please sign in to view your tickets.");
    }
    if (requested && requested !== userEmail) {
      throw new Error("Not authorized.");
    }
    emailToFilter = userEmail; // students only ever see their own
  }

  try {
    const sheet = getTicketsSheet_();
    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) return [];

    const headers = values[0].map(function (h) { return String(h || "").trim(); });

    let tickets = values.slice(1)
      .filter(function (row) {
        return row.some(function (cell) { return String(cell || "").trim() !== ""; });
      })
      .map(function (row) { return rowToTicket_(headers, row); });

    if (emailToFilter) {
      tickets = tickets.filter(function (t) {
        return String(t.student_email || "").trim().toLowerCase() === emailToFilter;
      });
      return tickets;
    }

    return tickets.reverse(); // admin: newest first

  } catch (err) {
    Logger.log("getTickets failed: " + (err && err.message ? err.message : err));
    throw new Error("Could not read tickets. Please contact the administrator.");
  }
}

function getTicketsSheet_() {
  const ss = SpreadsheetApp.openById(getConfig_("SPREADSHEET_ID"));
  const sheet = ss.getSheetByName("Tickets");
  if (!sheet) {
    throw new Error('The "Tickets" tab was not found in the spreadsheet.');
  }
  return sheet;
}

/** Admin dashboard data. Admin only. */
function getDashboardData() {
  try {
    assertAdmin_();
    return { ok: true, tickets: getTickets({}) };
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

/* ------------------------------------------------------------------ */
/* Diagnostics (admin only; run from the Apps Script editor)           */
/* NOTE: public functions can be called from the browser, so these     */
/* are locked behind assertAdmin_(). Delete them before sharing code.  */
/* ------------------------------------------------------------------ */

function testFastnConnection() {
  assertAdmin_();
  const result = submitToFastn({
    campus_id: "CAMPUS_01",
    campus_name: "Main Campus",
    location: "Building A",
    category: "IT",
    description: "Backend connection test, please ignore this ticket.",
    student_email: "test@example.com"
  });
  Logger.log(JSON.stringify(result, null, 2));
  return { ok: result.ok, status: result.status };
}

function testDashboardConnection() {
  assertAdmin_();
  const sheet = getTicketsSheet_();
  const values = sheet.getDataRange().getValues();
  const info = {
    ok: true,
    sheetName: sheet.getName(),
    rows: values.length,
    columns: values.length ? values[0].length : 0
  };
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

/* ------------------------------------------------------------------ */
/* Row helpers                                                         */
/* ------------------------------------------------------------------ */

function rowToTicket_(headers, row) {
  const out = {};

  headers.forEach(function (header, i) {
    const key = normalizeHeader_(header);
    if (key) out[key] = formatCell_(row[i]);
  });

  out.ticket_id = firstValue_(out.ticket_id, out.ticketid, out.id);
  out.campus_id = firstValue_(out.campus_id, out.campusid);
  out.campus_name = firstValue_(out.campus_name, out.campusname, out.campus);
  out.student_email = firstValue_(out.student_email, out.studentemail, out.email);
  out.created_at = firstValue_(out.created_at, out.created, out.timestamp);
  out.updated_at = firstValue_(out.updated_at, out.updated);
  out.status = firstValue_(out.status, "Open");
  out.action = firstValue_(out.action, out.result_action);

  return out;
}

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
