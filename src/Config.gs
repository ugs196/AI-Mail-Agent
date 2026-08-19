const CONFIG = {

  // ── OpenRouter ──────────────────────────────────────────
  OPENROUTER_API_KEY: "",                        // ← paste your OpenRouter key here (or set the OPENROUTER_API_KEY script property instead — see README)
  MODEL: "nvidia/nemotron-3-super-120b-a12b:free",  // change model anytime — any OpenRouter model id works

  ROOT_FOLDER_NAME:  "Agent Teams",
  ROLES_FOLDER_NAME: "Roles",

  // These four IDs are filled in automatically by runSetup() the first
  // time it runs, and cached in Script Properties from then on.
  // You do not need to set them by hand.
  ROOT_FOLDER_ID:   "",
  ROLES_FOLDER_ID:  "",
  MEMORY_DOC_ID:    "",
  CONTROL_SHEET_ID: "",

  BOSS_INTERVAL_MINUTES:  5,
  EMAIL_INTERVAL_MINUTES: 5,

  MAX_EMAILS_PER_RUN:  10,
  AUTO_REPLY_ENABLED:  true,
  REPLY_SIGNATURE:     "\n\n— Your Assistant",   // ← customize the sign-off
  SENDER_NAME:         "Your Assistant",         // ← used in the HTML email header

  MAX_TOKENS:   2000,
  TEMPERATURE:  0.4
};

function getProp(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

function setProp(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}

function getPropOrConfig(key) {
  var val = getProp(key);
  return (val && val.length > 0) ? val : (CONFIG[key] || "");
}

function getRootFolder() {
  var id = getPropOrConfig("ROOT_FOLDER_ID");
  if (id && id.length > 0) return DriveApp.getFolderById(id);
  var folders = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  throw new Error("Root folder not found. Run Setup first.");
}

function getRolesFolder() {
  var id = getPropOrConfig("ROLES_FOLDER_ID");
  if (id && id.length > 0) return DriveApp.getFolderById(id);
  var root = getRootFolder();
  var sub = root.getFoldersByName(CONFIG.ROLES_FOLDER_NAME);
  if (sub.hasNext()) return sub.next();
  throw new Error("Roles folder not found. Run Setup first.");
}

function getMemoryDoc() {
  var id = getPropOrConfig("MEMORY_DOC_ID");
  if (!id || id.length === 0) throw new Error("Memory doc not configured. Run Setup first.");
  return DocumentApp.openById(id);
}

function getControlSheet() {
  var id = getPropOrConfig("CONTROL_SHEET_ID");
  if (!id || id.length === 0) throw new Error("Control sheet not configured. Run Setup first.");
  return SpreadsheetApp.openById(id);
}
