/**
 * Read the full text content of the Memory Doc.
 * @returns {string}
 */
function memoryRead() {
  try {
    const doc  = getMemoryDoc();
    const text = doc.getBody().getText();
    Logger.log("[Memory] Read " + text.length + " chars");
    return text;
  } catch (e) {
    Logger.log("[Memory] Read error: " + e.message);
    return "(Memory unavailable: " + e.message + ")";
  }
}


/**
 * Append a note to the Memory Doc under AGENT NOTES.
 * @param {string} agentName  - e.g. "Boss Agent", "Mail Agent"
 * @param {string} note       - The note text to append
 */
function memoryAppend(agentName, note) {
  try {
    const doc  = getMemoryDoc();
    const body = doc.getBody();
    const ts   = new Date().toISOString();
    const line = "\n[" + ts + "] [" + agentName + "]\n" + note;

    // Find AGENT NOTES heading and insert after it
    const paras = body.getParagraphs();
    let insertAfter = null;
    for (const p of paras) {
      if (p.getText().includes("AGENT NOTES")) {
        insertAfter = p;
      }
    }

    if (insertAfter) {
      body.insertParagraph(paras.indexOf(insertAfter) + 1, line);
    } else {
      body.appendParagraph(line);
    }

    // Update last-updated timestamp in the first paragraph area
    _updateTimestamp(body);
    doc.saveAndClose();
    Logger.log("[Memory] Appended note from " + agentName);
  } catch (e) {
    Logger.log("[Memory] Append error: " + e.message);
  }
}


/**
 * Replace the entire Memory Doc content (used for summarization).
 * @param {string} newContent
 */
function memoryReplace(newContent) {
  try {
    const doc  = getMemoryDoc();
    const body = doc.getBody();
    body.clear();
    body.appendParagraph(newContent);
    doc.saveAndClose();
    Logger.log("[Memory] Replaced full content (" + newContent.length + " chars)");
  } catch (e) {
    Logger.log("[Memory] Replace error: " + e.message);
  }
}


/**
 * Summarize memory if it exceeds a length threshold (called by Boss periodically).
 */
function memoryMaybeSummarize() {
  const text = memoryRead();
  if (text.length < 5000) return; // No need to summarize yet

  Logger.log("[Memory] Memory is long (" + text.length + " chars), requesting summary...");
  const summary = askLLM(
    "You are a memory manager for an AI agent system. Summarize the following agent memory, preserving all important facts, preferences, and context. Be concise but complete. Keep the same section headings.",
    text
  );
  memoryReplace(summary);
  Logger.log("[Memory] Memory summarized to " + summary.length + " chars");
}


// ── Internal helper ───────────────────────────────────────────
function _updateTimestamp(body) {
  const paras = body.getParagraphs();
  for (const p of paras) {
    if (p.getText().startsWith("Last Updated:")) {
      p.replaceText("Last Updated:.*", "Last Updated: " + new Date().toISOString());
      return;
    }
  }
}
