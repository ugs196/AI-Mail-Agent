function runBossAgent() {
  Logger.log("=== BOSS AGENT RUN: " + new Date().toISOString() + " ===");

  try {
    // 1. Load role
    const rolePrompt = readRoleFile("BOSS.txt");

    // 2. Gather context
    const memory   = memoryRead();
    const commands = sheetGetPending("COMMANDS");

    if (commands.length === 0) {
      Logger.log("[Boss] No pending commands. Idle.");
      return;
    }

    Logger.log("[Boss] Processing " + commands.length + " pending command(s)...");

    // 3. Process each pending command
    for (const cmd of commands) {
      _bossProcessCommand(rolePrompt, memory, cmd);
      Utilities.sleep(2000); // Avoid rate limiting
    }

    // 4. Periodically summarize memory
    memoryMaybeSummarize();

  } catch (e) {
    Logger.log("[Boss] FATAL ERROR: " + e.message);
    sheetBossReport("❌ Boss Agent crashed: " + e.message);
  }
}


/**
 * Process a single COMMANDS row.
 */
function _bossProcessCommand(rolePrompt, memory, cmd) {
  const commandText = String(cmd.COMMAND || "").trim();
  const from        = String(cmd.FROM || "Human").trim();
  Logger.log("[Boss] Command from [" + from + "]: " + commandText.substring(0, 80));

  // Build context message
  const contextMsg = [
    "=== CURRENT MEMORY ===",
    memory.substring(0, 3000), // truncate to avoid token overflow
    "",
    "=== NEW COMMAND ===",
    "From: " + from,
    "Command: " + commandText,
    "",
    "=== YOUR TASK ===",
    "Analyze this command and respond with a JSON object:",
    "{",
    '  "action": "reply" | "delegate_mail" | "drive_task" | "memory_update" | "ignore",',
    '  "response": "Your report/reply text for the BOSS_REPORTS tab",',
    '  "mail_task": "If delegating: full instruction for mail agent (or null)",',
    '  "mail_target": "Email address if mail task targets a specific sender (or null)",',
    '  "memory_note": "Any important fact to save to memory (or null)"',
    "}",
    "",
    "Respond ONLY with the JSON object, no other text.",
  ].join("\n");

  let decision;
  try {
    const raw = askLLM(rolePrompt, contextMsg);
    decision  = JSON.parse(_extractJSON(raw));
  } catch (e) {
    Logger.log("[Boss] LLM parse error: " + e.message);
    sheetMarkRow("COMMANDS", cmd._rowIndex, "ERROR", "LLM parse failed: " + e.message);
    return;
  }

  Logger.log("[Boss] Decision action: " + decision.action);

  // Execute decision
  switch (decision.action) {

    case "delegate_mail":
      if (decision.mail_task) {
        sheetAddTask("EMAIL_TASKS", "Boss Agent", decision.mail_task, decision.mail_target || "");
        Logger.log("[Boss] Delegated task to Mail Agent: " + decision.mail_task.substring(0, 60));
      }
      break;

    case "memory_update":
      // Just save the memory note (done below)
      break;

    case "drive_task":
      // Future: Drive agent delegation
      Logger.log("[Boss] Drive task noted (no Drive agent yet): " + decision.response);
      break;

    case "ignore":
      Logger.log("[Boss] Command ignored per decision.");
      break;

    default: // "reply" and anything else
      break;
  }

  // Save memory note if any
  if (decision.memory_note && decision.memory_note !== "null") {
    memoryAppend("Boss Agent", decision.memory_note);
  }

  // Write boss report
  const reportText = [
    "Command from [" + from + "]: " + commandText,
    "Action taken: " + decision.action,
    decision.response || "",
  ].join("\n");
  sheetBossReport(reportText);

  // Mark command done
  sheetMarkRow("COMMANDS", cmd._rowIndex, "DONE", decision.action + ": " + (decision.response || "").substring(0, 100));
  Logger.log("[Boss] Command processed ✅");
}


// ── Manual trigger: run boss immediately ──────────────────────
function runBossNow() {
  runBossAgent();
}


// ── Manual: send a command as the human ────────────────────────
function sendCommandToBoss(commandText) {
  sheetAddTask("COMMANDS", "Human", commandText);
  Logger.log("Command added for Boss: " + commandText);
}


// ── Utility: extract JSON from LLM response ────────────────────
function _extractJSON(text) {
  // Try to find a JSON block even if the LLM added extra text
  const match = text.match(/\{[\s\S]*\}/);
  if (match) return match[0];
  return text; // will fail JSON.parse if malformed — caught by caller
}
