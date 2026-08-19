function runMailAgent() {
  Logger.log("=== MAIL AGENT RUN: " + new Date().toISOString() + " ===");

  try {
    const rolePrompt = readRoleFile("EMAILAGENT.txt");
    const memory     = memoryRead();

    // Run in two phases:
    // Phase 1: Process unread emails (scan + auto-reply)
    _mailPhase1_ScanUnread(rolePrompt, memory);

    // Phase 2: Execute pending tasks from Boss / Human
    _mailPhase2_ExecuteTasks(rolePrompt, memory);

  } catch (e) {
    Logger.log("[Mail] FATAL ERROR: " + e.message);
    sheetEmailReport("SYSTEM", "Error", "ERROR", "Mail Agent crashed", e.message);
  }
}


// ── Phase 1: Scan Unread Emails ─────────────────────────────
function _mailPhase1_ScanUnread(rolePrompt, memory) {
  const threads = GmailApp.search("is:unread in:inbox", 0, CONFIG.MAX_EMAILS_PER_RUN);
  Logger.log("[Mail] Found " + threads.length + " unread thread(s)");

  for (const thread of threads) {
    try {
      _processThread(rolePrompt, memory, thread);
      Utilities.sleep(3000); // Rate limit buffer
    } catch (e) {
      Logger.log("[Mail] Error processing thread: " + e.message);
    }
  }
}


/**
 * Process a single Gmail thread: summarize, prioritize, reply/draft.
 */
function _processThread(rolePrompt, memory, thread) {
  const messages  = thread.getMessages();
  const lastMsg   = messages[messages.length - 1];
  const from      = lastMsg.getFrom();
  const subject   = thread.getFirstMessageSubject();
  const body      = lastMsg.getPlainBody().substring(0, 2000); // truncate long bodies
  const msgId     = lastMsg.getId();

  Logger.log("[Mail] Processing: \"" + subject + "\" from " + from);

  // Build analysis prompt
  const analysisPrompt = [
    "=== OPERATOR MEMORY / CONTEXT ===",
    memory.substring(0, 1500),
    "",
    "=== EMAIL TO ANALYZE ===",
    "From: "    + from,
    "Subject: " + subject,
    "Body:\n"   + body,
    "",
    "=== YOUR TASK ===",
    "Respond ONLY with a JSON object:",
    "{",
    '  "priority": "URGENT" | "IMPORTANT" | "NORMAL" | "SPAM",',
    '  "summary": "2-3 sentence summary of the email",',
    '  "auto_reply": true | false,',
    '  "reply_text": "Full reply text if auto_reply is true, else null",',
    '  "create_draft": true | false,',
    '  "draft_text": "Draft text if create_draft is true, else null",',
    '  "memory_note": "Important fact to remember about this sender/topic, or null",',
    '  "label": "Label to apply: URGENT | IMPORTANT | NORMAL | SPAM | null"',
    "}",
    "",
    "Rules:",
    "- Set auto_reply=true for all emails except SPAM.",
    "- Set create_draft=true as well for IMPORTANT/URGENT so the human has a copy.",
    "- Never auto_reply to SPAM.",
    "- Keep reply_text professional and concise (2-4 sentences).",
    "- Respond ONLY with JSON.",
  ].join("\n");

  let decision;
  try {
    const raw = askLLM(rolePrompt, analysisPrompt);
    decision  = JSON.parse(_extractJSON(raw));
  } catch (e) {
    Logger.log("[Mail] LLM parse error for email \"" + subject + "\": " + e.message);
    sheetEmailReport(from, subject, "ERROR", "LLM parse failed", e.message);
    return;
  }

  Logger.log("[Mail] Priority: " + decision.priority + " | AutoReply: " + decision.auto_reply);

  // Apply Gmail label
  _applyLabel(thread, decision.priority || "NORMAL");

  // Auto-reply
  let actionTaken = "Labeled:" + decision.priority;
  if (decision.auto_reply && decision.reply_text && CONFIG.AUTO_REPLY_ENABLED && decision.priority !== "SPAM") {
    const replyText = decision.reply_text + CONFIG.REPLY_SIGNATURE;
    const replyHtml = _wrapHtmlTemplate({
      heading: "Reply from " + (CONFIG.SENDER_NAME || "the team"),
      bodyText: decision.reply_text,
      signatureHtml: CONFIG.REPLY_SIGNATURE_HTML || _textToHtmlParagraphs(CONFIG.REPLY_SIGNATURE)
    });
    lastMsg.reply(replyText, { htmlBody: replyHtml });
    thread.markRead();
    actionTaken = "AUTO_REPLIED";
    Logger.log("[Mail] Auto-replied to: " + from);
  }

  // Create draft
  if (decision.create_draft && decision.draft_text) {
    const draftText = decision.draft_text + CONFIG.REPLY_SIGNATURE;
    const draftHtml = _wrapHtmlTemplate({
      heading: "Draft — " + subject,
      bodyText: decision.draft_text,
      signatureHtml: CONFIG.REPLY_SIGNATURE_HTML || _textToHtmlParagraphs(CONFIG.REPLY_SIGNATURE)
    });
    GmailApp.createDraft(from, "Re: " + subject, draftText, { htmlBody: draftHtml });
    actionTaken += " | DRAFT_CREATED";
    Logger.log("[Mail] Draft created for: " + from);
  }

  // Mark as read if SPAM
  if (decision.priority === "SPAM") {
    thread.markRead();
    actionTaken = "MARKED_READ (SPAM)";
  }

  // Save memory note
  if (decision.memory_note && decision.memory_note !== "null") {
    memoryAppend("Mail Agent", "Re: \"" + subject + "\" from " + from + " — " + decision.memory_note);
  }

  // Write to report
  sheetEmailReport(from, subject, decision.priority, decision.summary, actionTaken);
}


// ── Phase 2: Execute Boss/Human Tasks ───────────────────────
function _mailPhase2_ExecuteTasks(rolePrompt, memory) {
  const tasks = sheetGetPending("EMAIL_TASKS");
  Logger.log("[Mail] " + tasks.length + " pending task(s) from Boss/Human");

  for (const task of tasks) {
    try {
      _executeMailTask(rolePrompt, memory, task);
      Utilities.sleep(2000);
    } catch (e) {
      Logger.log("[Mail] Task error: " + e.message);
      sheetMarkRow("EMAIL_TASKS", task._rowIndex, "ERROR", e.message);
    }
  }
}


/**
 * Execute a specific task from Boss/Human (e.g. "Reply to X saying Y").
 */
function _executeMailTask(rolePrompt, memory, task) {
  const instruction = String(task.TASK   || "").trim();
  const target      = String(task.TARGET || "").trim();
  Logger.log("[Mail] Executing task: " + instruction.substring(0, 80));

  // Find target thread if target is specified
  let targetThread = null;
  if (target && target.includes("@")) {
    const results = GmailApp.search("from:" + target, 0, 1);
    if (results.length > 0) targetThread = results[0];
  } else if (target) {
    const results = GmailApp.search("subject:(" + target + ")", 0, 1);
    if (results.length > 0) targetThread = results[0];
  }

  // Build execution prompt
  const taskPrompt = [
    "=== OPERATOR MEMORY ===",
    memory.substring(0, 1000),
    "",
    "=== TASK ===",
    "Instruction: " + instruction,
    "Target: "      + (target || "not specified"),
    targetThread ? ("Email found: " + targetThread.getFirstMessageSubject()) : "No matching email found.",
    "",
    "=== YOUR TASK ===",
    "Respond ONLY with JSON:",
    "{",
    '  "action": "send" | "draft" | "search" | "label" | "none",',
    '  "to": "recipient email or null",',
    '  "subject": "email subject or null",',
    '  "body": "email body or null",',
    '  "label": "label name or null",',
    '  "result": "Short description of what you did"',
    "}",
  ].join("\n");

  let decision;
  try {
    const raw = askLLM(rolePrompt, taskPrompt);
    decision  = JSON.parse(_extractJSON(raw));
  } catch (e) {
    sheetMarkRow("EMAIL_TASKS", task._rowIndex, "ERROR", "Parse error: " + e.message);
    return;
  }

  let resultMsg = decision.result || "Task processed";

  switch (decision.action) {
    case "send":
      if (decision.to && decision.body) {
        const html = _wrapHtmlTemplate({
          heading: decision.subject || "Message",
          bodyText: decision.body,
          signatureHtml: CONFIG.REPLY_SIGNATURE_HTML || _textToHtmlParagraphs(CONFIG.REPLY_SIGNATURE)
        });
        GmailApp.sendEmail(
          decision.to,
          decision.subject || "(No Subject)",
          decision.body + CONFIG.REPLY_SIGNATURE, // plain-text fallback
          { htmlBody: html }
        );
        resultMsg = "Email sent to " + decision.to;
        Logger.log("[Mail] Email sent to " + decision.to);
      }
      break;

    case "draft":
      if (decision.to && decision.body) {
        const html = _wrapHtmlTemplate({
          heading: decision.subject || "Message",
          bodyText: decision.body,
          signatureHtml: CONFIG.REPLY_SIGNATURE_HTML || _textToHtmlParagraphs(CONFIG.REPLY_SIGNATURE)
        });
        GmailApp.createDraft(
          decision.to,
          decision.subject || "(No Subject)",
          decision.body + CONFIG.REPLY_SIGNATURE,
          { htmlBody: html }
        );
        resultMsg = "Draft created for " + decision.to;
        Logger.log("[Mail] Draft created for " + decision.to);
      }
      break;

    case "label":
      if (targetThread && decision.label) {
        _applyLabel(targetThread, decision.label);
        resultMsg = "Labeled thread: " + decision.label;
      }
      break;

    case "search":
      resultMsg = "Search performed: " + instruction;
      break;

    case "none":
    default:
      resultMsg = "No action taken: " + (decision.result || "task unclear");
      break;
  }

  sheetMarkRow("EMAIL_TASKS", task._rowIndex, "DONE", resultMsg);
  sheetEmailReport("TASK", instruction, "TASK", "", resultMsg);
}


// ── Apply/create Gmail label ─────────────────────────────────
function _applyLabel(thread, labelName) {
  try {
    let label = GmailApp.getUserLabelByName("AI/" + labelName);
    if (!label) label = GmailApp.createLabel("AI/" + labelName);
    thread.addLabel(label);
  } catch (e) {
    Logger.log("[Mail] Label error: " + e.message);
  }
}


// ── JSON extractor (same as Boss) ────────────────────────────
function _extractJSON(text) {
  const match = text.match(/\{[\s\S]*\}/);
  if (match) return match[0];
  return text;
}


// ============================================================
//  HTML TEMPLATE — black & white theme
// ============================================================

/**
 * Converts plain text (possibly multi-paragraph, \n separated) into
 * safe HTML <p> blocks styled to match the template body copy.
 */
function _textToHtmlParagraphs(text) {
  if (!text) return "";
  return String(text)
    .split(/\n{2,}/)                 // paragraph breaks on blank lines
    .map(function (para) {
      const escaped = para
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\n/g, "<br>");     // single line breaks inside a paragraph
      return '<p style="font-size:15px; color:#4d4d4d; line-height:1.7; margin:0 0 16px 0;">' + escaped + '</p>';
    })
    .join("");
}

/**
 * Wraps content in the black & white responsive card template.
 *
 * opts = {
 *   heading:       string  — text shown in the dark header bar
 *   bodyText:      string  — raw text, will be turned into <p> blocks
 *   bodyHtml:      string  — optional, use instead of bodyText if you
 *                            already have HTML (bodyText is ignored if set)
 *   signatureHtml: string  — optional HTML signature block
 *   footerText:    string  — optional footer line
 * }
 */
function _wrapHtmlTemplate(opts) {
  opts = opts || {};
  const heading    = opts.heading    || "Message";
  const bodyHtml   = opts.bodyHtml   || _textToHtmlParagraphs(opts.bodyText || "");
  const signature  = opts.signatureHtml || "";
  const footerText = opts.footerText || "This message was sent automatically. Reply to this email to reach a human.";

  return [
    '<!DOCTYPE html>',
    '<html lang="en"><head><meta charset="UTF-8">',
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">',
    '<style>',
    '  @media only screen and (max-width: 620px) {',
    '    .container { width: 100% !important; }',
    '    .padded { padding: 24px !important; }',
    '    .header-padded { padding: 26px 24px !important; }',
    '    h1 { font-size: 18px !important; }',
    '  }',
    '</style></head>',
    '<body style="margin:0; padding:0; background-color:#eeeeee; font-family: -apple-system, BlinkMacSystemFont, \'Segoe UI\', Arial, Helvetica, sans-serif;">',
    '  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#eeeeee; padding:24px 12px;">',
    '    <tr><td align="center">',
    '      <table role="presentation" class="container" width="600" cellpadding="0" cellspacing="0" style="width:600px; max-width:100%; background-color:#ffffff; border-radius:10px; overflow:hidden; box-shadow:0 4px 14px rgba(0,0,0,0.12);">',

    // Header
    '        <tr><td class="header-padded" style="background-color:#1a1a1a; padding:32px 36px;">',
    '          <table role="presentation" cellpadding="0" cellspacing="0" width="100%"><tr>',
    '            <td style="width:44px; vertical-align:middle;">',
    '              <table role="presentation" cellpadding="0" cellspacing="0" width="36" height="36" style="background-color:#ffffff; border-radius:50%;"><tr>',
    '                <td align="center" valign="middle" style="font-size:18px; color:#1a1a1a;">•</td>',
    '              </tr></table>',
    '            </td>',
    '            <td style="vertical-align:middle;">',
    '              <h1 style="margin:0; color:#ffffff; font-size:20px; font-weight:600; letter-spacing:0.2px;">' + heading + '</h1>',
    '            </td>',
    '          </tr></table>',
    '        </td></tr>',

    // Body
    '        <tr><td class="padded" style="padding:36px;">',
    bodyHtml,
    signature ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:28px 0;"><tr><td style="border-top:1px solid #e0e0e0;"></td></tr></table>' + signature : '',
    '        </td></tr>',

    // Footer
    '        <tr><td style="background-color:#f5f5f5; padding:18px 36px; font-size:12px; color:#999999; text-align:center; border-top:1px solid #e8e8e8;">',
    footerText,
    '        </td></tr>',

    '      </table>',
    '    </td></tr>',
    '  </table>',
    '</body></html>'
  ].join("\n");
}


// ── Manual triggers ─────────────────────────────────────────
function runMailAgentNow()      { runMailAgent(); }
function scanUnreadNow()        { _mailPhase1_ScanUnread(readRoleFile("EMAILAGENT.txt"), memoryRead()); }
function runMailTasksNow()      { _mailPhase2_ExecuteTasks(readRoleFile("EMAILAGENT.txt"), memoryRead()); }

// Quick manual test of the template itself, without touching Gmail search:
function testHtmlTemplate() {
  const html = _wrapHtmlTemplate({
    heading: "Test Email",
    bodyText: "Hi,\n\nThis is a test of the black & white responsive template.\n\nSecond paragraph here.",
    signatureHtml: '<p style="font-size:15px; color:#1a1a1a; line-height:1.7; margin:0;">Best regards,<br><strong>' + (CONFIG.SENDER_NAME || "The Team") + '</strong></p>'
  });
  GmailApp.sendEmail(Session.getActiveUser().getEmail(), "Template Test", "Plain text fallback", { htmlBody: html });
}
