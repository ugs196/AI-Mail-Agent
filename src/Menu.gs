function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🤖 Agents")
    .addItem("🚀 Initial Setup (run first)", "runSetup")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("▶ Run Agents")
        .addItem("Run Boss Agent now",       "runBossNow")
        .addItem("Run Mail Agent now",       "runMailAgentNow")
        .addItem("Scan Unread Emails only",  "scanUnreadNow")
        .addItem("Execute Mail Tasks only",  "runMailTasksNow")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("📋 Send Commands")
        .addItem("Send command to Boss...",  "menuSendCommand")
        .addItem("Add Mail task...",         "menuAddMailTask")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("🧠 Memory")
        .addItem("View Memory Doc",          "menuOpenMemoryDoc")
        .addItem("Force Memory Summarize",   "memoryMaybeSummarize")
        .addItem("Append note to Memory...", "menuAppendMemory")
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("⚙ Triggers")
        .addItem("Reinstall triggers",       "_installTriggers")
        .addItem("Remove all triggers",      "removeTriggers")
    )
    .addSeparator()
    .addItem("ℹ System Status",              "menuSystemStatus")
    .addToUi();
}


// ── Menu actions ────────────────────────────────────────────

function menuSendCommand() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "📨 Command to Boss",
    "Type your instruction for the Boss Agent:",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const cmd = res.getResponseText().trim();
  if (!cmd) { ui.alert("No command entered."); return; }
  sheetAddTask("COMMANDS", "Human", cmd);
  ui.alert("✅ Command sent!\n\nThe Boss will process it on the next trigger run (or click Run Boss Agent now).");
}


function menuAddMailTask() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "📧 Mail Task",
    "Describe what the Mail Agent should do\n(e.g. 'Reply to john@example.com and say I will call tomorrow'):",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const task = res.getResponseText().trim();
  if (!task) { ui.alert("No task entered."); return; }
  sheetAddTask("EMAIL_TASKS", "Human", task, "");
  ui.alert("✅ Mail task added!\n\nThe Mail Agent will execute it on the next run.");
}


function menuOpenMemoryDoc() {
  const ui = SpreadsheetApp.getUi();
  try {
    const doc = getMemoryDoc();
    ui.alert("📄 Memory Doc URL:\n\n" + doc.getUrl() +
             "\n\nOpen it in your browser to read or edit memory.");
  } catch (e) {
    ui.alert("❌ Error: " + e.message + "\n\nHave you run Setup?");
  }
}


function menuAppendMemory() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "🧠 Append to Memory",
    "Enter a note to save in shared memory (visible to all agents):",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const note = res.getResponseText().trim();
  if (!note) { ui.alert("No note entered."); return; }
  memoryAppend("Human (via Menu)", note);
  ui.alert("✅ Note added to memory!");
}


function menuSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const triggerInfo = triggers.map(t =>
      t.getHandlerFunction() + " — every " + (t.getTriggerSource() === ScriptApp.TriggerSource.CLOCK ? "~" + t.getEveryMinutes() + "min" : "?")
    ).join("\n  ");

    const memLen  = memoryRead().length;
    const sheetId = getPropOrConfig("CONTROL_SHEET_ID");
    const memId   = getPropOrConfig("MEMORY_DOC_ID");

    ui.alert(
      "📊 Agent System Status",
      "✅ Active triggers:\n  " + (triggerInfo || "(none)") +
      "\n\n📄 Memory Doc ID: "   + (memId   || "Not set") +
      "\n📊 Control Sheet ID: " + (sheetId || "Not set") +
      "\n🧠 Memory size: "      + memLen + " characters" +
      "\n\n🤖 Model: " + CONFIG.MODEL +
      "\n⏱ Boss interval: " + CONFIG.BOSS_INTERVAL_MINUTES + " min" +
      "\n📧 Mail interval: " + CONFIG.EMAIL_INTERVAL_MINUTES + " min",
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert("❌ Status error: " + e.message);
  }
}
