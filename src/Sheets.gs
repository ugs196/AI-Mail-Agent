// ── Column maps per tab ──────────────────────────────────────
const COL = {
  COMMANDS: {
    TIMESTAMP: 1, FROM: 2, COMMAND: 3, STATUS: 4, RESULT: 5
  },
  EMAIL_TASKS: {
    TIMESTAMP: 1, TASK: 2, TARGET: 3, STATUS: 4, RESULT: 5
  },
  EMAIL_REPORTS: {
    TIMESTAMP: 1, FROM: 2, SUBJECT: 3, PRIORITY: 4, SUMMARY: 5, ACTION: 6
  },
  BOSS_REPORTS: {
    TIMESTAMP: 1, REPORT: 2
  },
  MEMORY_LOG: {
    TIMESTAMP: 1, AGENT: 2, NOTE: 3
  },
};


// ── Read pending commands (any tab) ─────────────────────────
/**
 * Get all rows with STATUS = "PENDING" from a tab.
 * @param {string} tabName
 * @returns {Array} Array of row objects
 */
function sheetGetPending(tabName) {
  const ss    = getControlSheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) { Logger.log("[Sheets] Tab not found: " + tabName); return []; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const statusCol = headers.indexOf("STATUS");
  if (statusCol === -1) return [];

  const pending = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][statusCol]).toUpperCase() === "PENDING") {
      const row = { _rowIndex: i + 1 };
      headers.forEach((h, j) => row[h] = data[i][j]);
      pending.push(row);
    }
  }
  Logger.log("[Sheets] " + pending.length + " pending rows in " + tabName);
  return pending;
}


/**
 * Mark a row's STATUS and RESULT.
 * @param {string} tabName
 * @param {number} rowIndex  - 1-based sheet row
 * @param {string} status    - "DONE" | "ERROR" | "IN_PROGRESS"
 * @param {string} result    - Short result message
 */
function sheetMarkRow(tabName, rowIndex, status, result) {
  const ss    = getControlSheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;

  const headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf("STATUS")  + 1;
  const resultCol = headers.indexOf("RESULT")  + 1;

  if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue(status);
  if (resultCol > 0) sheet.getRange(rowIndex, resultCol).setValue(result || "");
  Logger.log("[Sheets] Marked row " + rowIndex + " in " + tabName + " as " + status);
}


/**
 * Append a new row to any tab.
 * @param {string} tabName
 * @param {Array}  values   - Array of cell values in header order
 */
function sheetAppendRow(tabName, values) {
  const ss    = getControlSheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    Logger.log("[Sheets] Cannot append — tab not found: " + tabName);
    return;
  }
  sheet.appendRow(values);
  Logger.log("[Sheets] Appended row to " + tabName);
}


/**
 * Write a boss report.
 * @param {string} report
 */
function sheetBossReport(report) {
  sheetAppendRow("BOSS_REPORTS", [new Date().toISOString(), report]);
}


/**
 * Write an email report row.
 */
function sheetEmailReport(from, subject, priority, summary, action) {
  sheetAppendRow("EMAIL_REPORTS", [
    new Date().toISOString(), from, subject, priority, summary, action
  ]);
}


/**
 * Add a command for an agent to the appropriate tab.
 * @param {string} tabName    - "COMMANDS" | "EMAIL_TASKS"
 * @param {string} from       - Who issued it (e.g. "Boss Agent")
 * @param {string} command    - The instruction text
 * @param {string} target     - For EMAIL_TASKS: target email address
 */
function sheetAddTask(tabName, from, command, target) {
  if (tabName === "EMAIL_TASKS") {
    sheetAppendRow("EMAIL_TASKS", [new Date().toISOString(), command, target || "", "PENDING", ""]);
  } else {
    sheetAppendRow("COMMANDS", [new Date().toISOString(), from, command, "PENDING", ""]);
  }
}


/**
 * Read the full content of a tab as an array of row objects.
 * @param {string} tabName
 * @param {number} limit   - Max rows to return (default 50)
 * @returns {Array}
 */
function sheetReadTab(tabName, limit) {
  limit = limit || 50;
  const ss    = getControlSheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];

  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];

  return data.slice(1, limit + 1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}


/**
 * Read role file from Drive (Boss or Mail Agent role text).
 * @param {string} filename  - "BOSS.txt" | "EMAILAGENT.txt"
 * @returns {string}
 */
function readRoleFile(filename) {
  try {
    const folder = getRolesFolder();
    const iter   = folder.getFilesByName(filename);
    if (!iter.hasNext()) throw new Error("Role file not found: " + filename);
    return iter.next().getBlob().getDataAsString();
  } catch (e) {
    Logger.log("[Sheets] Role file error: " + e.message);
    return "(Role file unavailable)";
  }
}
