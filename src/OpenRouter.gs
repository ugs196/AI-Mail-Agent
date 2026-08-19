/**
 * Core LLM call via OpenRouter.
 * @param {string} systemPrompt  - The agent's role/system instructions
 * @param {Array}  messages      - Array of {role, content} message objects
 * @param {number} maxTokens     - Max response tokens
 * @returns {string} The assistant's reply text
 */
function callLLM(systemPrompt, messages, maxTokens) {
  maxTokens = maxTokens || CONFIG.MAX_TOKENS;

  // Prefer a Script Property over the hardcoded CONFIG value so the key
  // never has to live in source control. See README "API key" section.
  var apiKey = getPropOrConfig("OPENROUTER_API_KEY");

  if (!apiKey || apiKey.length < 10) {
    throw new Error("OpenRouter API key missing. Set the OPENROUTER_API_KEY script property, or paste it into CONFIG.OPENROUTER_API_KEY in Config.gs.");
  }

  var payload = {
    model: CONFIG.MODEL,
    max_tokens: maxTokens,
    messages: [{ role: "user", content: systemPrompt + "\n\n" + messages[0].content }]
  };

  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + apiKey,
      "HTTP-Referer": "https://script.google.com",
      "X-Title": "GAS Agent Team"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  Logger.log("[OpenRouter] Calling: " + CONFIG.MODEL);

  var response = UrlFetchApp.fetch("https://openrouter.ai/api/v1/chat/completions", options);
  var code = response.getResponseCode();
  var text = response.getContentText();

  Logger.log("[OpenRouter] Response code: " + code);
  Logger.log("[OpenRouter] Response: " + text.substring(0, 300));

  if (code !== 200) {
    throw new Error("OpenRouter error " + code + ": " + text.substring(0, 300));
  }

  var json = JSON.parse(text);

  if (!json.choices || !json.choices[0]) {
    throw new Error("No choices in response: " + text.substring(0, 300));
  }

  var reply = json.choices[0].message.content;
  Logger.log("[OpenRouter] Reply: " + reply.substring(0, 150));
  return reply.trim();
}

function askLLM(systemPrompt, userMessage) {
  return callLLM(systemPrompt, [{ role: "user", content: userMessage }]);
}

function chatLLM(systemPrompt, history, newUserMsg) {
  var messages = [];
  for (var i = 0; i < history.length; i++) {
    messages.push(history[i]);
  }
  messages.push({ role: "user", content: newUserMsg });
  return callLLM(systemPrompt, messages);
}

function testOpenRouter() {
  var result = askLLM("You are a helpful assistant.", "Say hello in one sentence.");
  Logger.log("TEST RESULT: " + result);
}
