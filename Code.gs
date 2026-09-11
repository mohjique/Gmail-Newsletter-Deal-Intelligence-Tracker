// =============================================================================
//  VC Newsletter Intelligence Pipeline
//  Automatically extracts deals from newsletter emails into Google Sheets
//  using Google Gemini (free tier — no cost under normal use)
//
//  Setup instructions: see README.md
// =============================================================================


// ---------------------------------------------------------------------------
// CONFIG
// ---------------------------------------------------------------------------

const CONFIG = {
  // Retrieved from Script Properties — never hardcode keys here
  // Set these in: Project Settings > Script Properties
  get GEMINI_API_KEY()     { return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY'); },
  get SHEET_ID()           { return PropertiesService.getScriptProperties().getProperty('SHEET_ID'); },
  // Optional — leave unset in Script Properties to disable Telegram notifications
  get TELEGRAM_BOT_TOKEN() { return PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN'); },
  get TELEGRAM_CHAT_ID()   { return PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID'); },

  SHEET_NAME:      'Deals',
  PROCESSED_SHEET: 'Processed',   // hidden tab that caches processed thread IDs
  MODEL:           'gemini-3.6-flash',

  BATCH_SIZE: 10,   // emails per Gemini API call

  // Update this list to match your newsletter senders
  NEWSLETTER_QUERY: [
    'from:connie@strictlyvc.com',
    'from:lucinda.shen@axios.com',
    'from:newsletters@techcrunch.com',
    'from:newsletters@cheddar.com',
    'from:vcdeals@startupcareeradvice.com',
    'from:opening-bell@mail.beehiiv.com',
    'from:newcomer@substack.com',
    'from:blog@tomtunguz.com',
  ].join(' OR '),

  // Sheet columns in order
  COLUMNS: [
    'Date',
    'Category',
    'Company',
    'Deal Type',
    'Amount (USD)',
    'Investors / Acquirer / Parties',
    'Stage',
    'Sector',
    'Summary',
    'Source Newsletter',
    'Dedup Hash',   // hidden helper column — do not delete
  ],

  CATEGORIES: ['vc_funding', 'ma', 'ipo', 'notable', 'layoffs', 'leadership'],

  // Keywords used to pre-filter emails before sending to Gemini
  // Emails with none of these are skipped entirely
  DEAL_KEYWORDS: [
    'funding', 'raises', 'raised', 'acquisition', 'acquires', 'acquired',
    'ipo', 'layoffs', 'layoff', 'series', 'million', 'billion',
    'merger', 'stake', 'valuation', 'round', 'invest',
  ],
};


// ---------------------------------------------------------------------------
// DAILY SCAN — called by time-driven trigger each morning
// ---------------------------------------------------------------------------

function dailyScan() {
  Logger.log('Running daily scan...');

  const processedHashes  = getExistingHashes();
  const processedThreads = getProcessedThreadIds();

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const dateStr = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const query = '(' + CONFIG.NEWSLETTER_QUERY + ') after:' + dateStr;

  const response = Gmail.Users.Threads.list('me', {
    q: query,
    maxResults: 50,
  });

  if (!response.threads || response.threads.length === 0) {
    Logger.log('No new newsletter emails found.');
    return;
  }

  var batchBuffer = [];
  var totalDeals  = 0;

  for (var i = 0; i < response.threads.length; i++) {
    var threadMeta = response.threads[i];

    if (processedThreads.has(threadMeta.id)) continue;

    var thread   = GmailApp.getThreadById(threadMeta.id);
    var messages = thread.getMessages();

    for (var j = 0; j < messages.length; j++) {
      var emailData = extractEmailData(messages[j]);
      if (!emailData) continue;
      batchBuffer.push(emailData);

      if (batchBuffer.length >= CONFIG.BATCH_SIZE) {
        var deals    = callGeminiForDeals(batchBuffer);
        var newDeals = filterAndWrite(deals, processedHashes);
        totalDeals += newDeals;
        batchBuffer = [];
        Utilities.sleep(1000);
      }
    }

    markThreadProcessed(threadMeta.id);
    processedThreads.add(threadMeta.id);
  }

  // Process any remaining emails in the buffer
  if (batchBuffer.length > 0) {
    var deals    = callGeminiForDeals(batchBuffer);
    var newDeals = filterAndWrite(deals, processedHashes);
    totalDeals += newDeals;
  }

  Logger.log('Daily scan complete. New deals written: ' + totalDeals);
}


// ---------------------------------------------------------------------------
// INSTALL / REMOVE TRIGGER
// ---------------------------------------------------------------------------

function installDailyTrigger() {
  // Remove any existing dailyScan triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'dailyScan'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger('dailyScan')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  Logger.log('Daily trigger installed. dailyScan() will run every day at 7am.');
  try { SpreadsheetApp.getUi().alert('Daily automation is active. It will run every morning at 7am.'); } catch(e) {}
}

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(function(t) { return t.getHandlerFunction() === 'dailyScan'; })
    .forEach(function(t) { ScriptApp.deleteTrigger(t); });
  Logger.log('Daily trigger removed.');
}


// ---------------------------------------------------------------------------
// GEMINI API CALL
// ---------------------------------------------------------------------------

function callGeminiForDeals(emails) {
  var emailsText = emails.map(function(e, i) {
    return '---EMAIL ' + (i + 1) + '---\nFROM:' + e.sender + '\nDATE:' + e.date + '\nSUBJECT:' + e.subject + '\n' + e.body;
  }).join('\n\n');

  var prompt = 'Extract all deals from these emails and return a JSON array. Each object must have:\n' +
    'date(YYYY-MM-DD),\n' +
    'category(must be exactly one of: vc_funding | ma | ipo | notable | layoffs | leadership),\n' +
    'company(primary company name as a string),\n' +
    'deal_type(must be exactly one of: Seed | Series A | Series B | Series C | Series D | Series E+ | Growth Round | Acquisition | Merger | IPO | Funding Round | Layoff | Leadership Change | Other),\n' +
    'amount(plain number in millions, no $ sign or suffix — e.g. 500 for $500M, 1300 for $1.3B, 0.5 for $500K, 45.5 for $45.5M; if a range use the midpoint e.g. "$40M-$60M" becomes 50; if approximate e.g. "~$500M" use the stated number; null if not mentioned),\n' +
    'parties(external investors, acquirers, or counterparties only — never the company itself; comma-separated string; null if none named),\n' +
    'stage(must be exactly one of: Seed | Series A | Series B | Series C | Series D | Series E+ | Growth | Public | null if not applicable),\n' +
    'sector(must be exactly one of: AI/ML | Fintech | Biotech | Healthcare | Cybersecurity | Defense | Robotics | Semiconductors | Energy & Climate | Space & Aerospace | SaaS / Enterprise Software | Crypto / Web3 | Logistics & Supply Chain | Media & Entertainment | Quantum Computing | AgTech | Deep Tech / Hardware | Real Estate & PropTech | EdTech | Legal Tech | Data Infrastructure | Automotive & Mobility | Consumer & Retail | GovTech | Venture Capital | Telecommunications | Technology | Other),\n' +
    'summary(one concise sentence, max 20 words),\n' +
    'source(must be exactly one of: StrictlyVC | TechCrunch | Newcomer | Cheddar | Opening Bell | Axios | Tom Tunguz | VC Deals — match to the closest regardless of exact sender name in the email header)\n\n' +
    'Rules:\n' +
    '- All field values must strictly match the allowed lists above. Do not invent new values.\n' +
    '- If uncertain which sector fits, use Other.\n' +
    '- If uncertain which deal_type fits, use Other.\n' +
    '- If stage is not applicable or unclear, use null.\n' +
    '- Amount must be a plain number in millions or null. Never output a string, $ sign, or suffix.\n' +
    '- If money changes hands for equity always use vc_funding or ma over notable.\n' +
    '- Only extract explicitly stated deals — do not infer or hallucinate.\n' +
    '- Return [] if no relevant deals found.\n' +
    '- No markdown, no explanation. Pure JSON only.\n\n' +
    emailsText;

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + CONFIG.MODEL + ':generateContent?key=' + CONFIG.GEMINI_API_KEY;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0,
      maxOutputTokens: 8192,
    },
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  // [Change 6] Retry schedule for transient 5xx errors (e.g. 503 "high demand").
  // Up to 4 attempts total (1 initial + 3 retries), waits of 2s / 5s / 10s between them.
  var retryDelays  = [2000, 5000, 10000];
  var responseCode = null;
  var responseText = null;
  var response      = null;
  var rawText = null;
  var cleaned = null;

  try {
    for (var attempt = 0; attempt <= retryDelays.length; attempt++) {
      response     = UrlFetchApp.fetch(url, options);
      responseCode = response.getResponseCode();
      responseText = response.getContentText();

      if (responseCode === 200) break;

      var isRetryable = responseCode >= 500 && responseCode < 600;
      var isLastAttempt = attempt === retryDelays.length;

      if (!isRetryable || isLastAttempt) {
        Logger.log('Gemini API error: ' + responseCode + ' — ' + responseText);
        return [];
      }

      Logger.log('Gemini API error: ' + responseCode + ' — retrying in ' + (retryDelays[attempt] / 1000) + 's (attempt ' + (attempt + 1) + ' of ' + retryDelays.length + ')');
      Utilities.sleep(retryDelays[attempt]);
    }

    var json = JSON.parse(responseText);
    rawText  = json.candidates[0].content.parts[0].text.trim();
    cleaned  = rawText.replace(/^```(?:json)?\n?/, '').replace(/\n?```$/, '').trim();

    return JSON.parse(cleaned);
  } catch(e) {
    Logger.log('Error calling Gemini or parsing response: ' + e.message);
    // [Change 3] Surface what Gemini actually returned so a parse failure
    // is diagnosable instead of a dead end. Truncated to keep the log readable.
    if (cleaned) {
      Logger.log('Raw Gemini output (first 2000 chars): ' + cleaned.substring(0, 2000));
    }
    return [];
  }
}


// ---------------------------------------------------------------------------
// SHEET WRITE + DEDUPLICATION
// ---------------------------------------------------------------------------

function getExistingHashes() {
  var sheet   = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return new Set();

  // Hash lives in column 11
  var hashes = sheet.getRange(2, 11, lastRow - 1, 1).getValues();
  return new Set(hashes.map(function(r) { return r[0]; }).filter(Boolean));
}

function makeHash(deal) {
  var raw = [
    (deal.date      || '').toString().toLowerCase().trim(),
    (deal.company   || '').toLowerCase().trim(),
    (deal.deal_type || '').toLowerCase().trim(),
  ].join('|');
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw)
    .map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); })
    .join('');
}

function filterAndWrite(deals, existingHashes) {
  if (!deals || deals.length === 0) return 0;

  var sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  var rows  = [];

  for (var i = 0; i < deals.length; i++) {
    var deal = deals[i];
    if (!CONFIG.CATEGORIES.includes(deal.category)) continue;

    var hash = makeHash(deal);
    if (existingHashes.has(hash)) continue;

    rows.push([
      deal.date      || '',
      deal.category  || '',
      deal.company   || '',
      deal.deal_type || '',
      deal.amount    || '',
      deal.parties   || '',
      deal.stage     || '',
      deal.sector    || '',
      deal.summary   || '',
      deal.source    || '',
      hash,
    ]);

    existingHashes.add(hash);
    sendTelegramNotification(deal);
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, CONFIG.COLUMNS.length)
      .setValues(rows);
  }

  return rows.length;
}


// ---------------------------------------------------------------------------
// TELEGRAM NOTIFICATIONS
// ---------------------------------------------------------------------------

var TELEGRAM_CATEGORY_LABELS = {
  vc_funding: '💰 VC Funding',
  ma:         '🤝 M&A',
  ipo:        '📈 IPO',
  notable:    '📰 Notable',
  layoffs:    '📉 Layoffs',
  leadership: '👤 Leadership Change',
};

function sendTelegramNotification(deal) {
  if (!CONFIG.TELEGRAM_BOT_TOKEN || !CONFIG.TELEGRAM_CHAT_ID) return;

  var url = 'https://api.telegram.org/bot' + CONFIG.TELEGRAM_BOT_TOKEN + '/sendMessage';

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      chat_id: CONFIG.TELEGRAM_CHAT_ID,
      text: formatTelegramMessage(deal),
      parse_mode: 'HTML',
      disable_web_page_preview: true,
    }),
    muteHttpExceptions: true,
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('Telegram notification failed: ' + response.getContentText());
    }
  } catch(e) {
    Logger.log('Error sending Telegram notification: ' + e.message);
  }
}

function formatTelegramMessage(deal) {
  var lines = [];

  lines.push('<b>' + (TELEGRAM_CATEGORY_LABELS[deal.category] || deal.category) + '</b>');
  lines.push('<b>' + escapeHtml(deal.company || 'Unknown company') + '</b>' +
    (deal.deal_type ? ' — ' + escapeHtml(deal.deal_type) : ''));

  if (deal.amount != null && deal.amount !== '') lines.push('💵 $' + deal.amount + 'M');
  if (deal.parties)                              lines.push('🔗 ' + escapeHtml(deal.parties));
  if (deal.sector)                               lines.push('🏷 ' + escapeHtml(deal.sector));
  if (deal.summary)                              lines.push(escapeHtml(deal.summary));

  lines.push('📅 ' + (deal.date || '') + '  •  ' + (deal.source || ''));

  return lines.join('\n');
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function testTelegramNotification() {
  sendTelegramNotification({
    date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    category: 'vc_funding',
    company: 'Test Co',
    deal_type: 'Series B',
    amount: 50,
    parties: 'Example Ventures',
    stage: 'Series B',
    sector: 'AI/ML',
    summary: 'This is a test notification to confirm Telegram is configured correctly.',
    source: 'TechCrunch',
  });
  Logger.log('Test Telegram notification sent (check your chat, and the Execution log for errors).');
}


// ---------------------------------------------------------------------------
// THREAD ID CACHE — prevents re-processing seen threads
// ---------------------------------------------------------------------------

function getProcessedThreadIds() {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PROCESSED_SHEET);
    sheet.hideSheet();
    sheet.appendRow(['ThreadId']);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return new Set();

  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return new Set(ids.map(function(r) { return r[0]; }).filter(Boolean));
}

function markThreadProcessed(threadId) {
  var ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.PROCESSED_SHEET);
  sheet.appendRow([threadId]);
}


// ---------------------------------------------------------------------------
// EMAIL HELPERS
// ---------------------------------------------------------------------------

function extractEmailData(message) {
  try {
    var sender  = message.getFrom();
    var subject = message.getSubject();
    var date    = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var body = message.getPlainBody();
    if (!body || body.trim().length < 100) {
      body = stripHtml(message.getBody());
    }

    // Keyword pre-filter — skip emails with no deal signals
    var lowerBody  = (body || '').toLowerCase();
    var hasSignal  = CONFIG.DEAL_KEYWORDS.some(function(kw) { return lowerBody.includes(kw); });
    if (!hasSignal) return null;

    body = preprocessBody(body);

    // Long-form newsletters get a higher char limit to capture all deals
    var isLongNewsletter = sender.toLowerCase().includes('strictlyvc.com');
    var charLimit = isLongNewsletter ? 8000 : 3500;
    if (body.length > charLimit) {
      body = body.substring(0, charLimit) + '\n[truncated]';
    }

    if (!body || body.trim().length < 50) return null;

    return { sender: sender, subject: subject, date: date, body: body };
  } catch(e) {
    Logger.log('Error extracting email data: ' + e.message);
    return null;
  }
}

function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<img[^>]*>/gi, '')
    .replace(/<figure[^>]*>[\s\S]*?<\/figure>/gi, '')
    .replace(/<picture[^>]*>[\s\S]*?<\/picture>/gi, '')
    .replace(/<svg[^>]*>[\s\S]*?<\/svg>/gi, '')
    .replace(/style="[^"]*background-image[^"]*"/gi, '')
    .replace(/<span[^>]*>\s*<\/span>/gi, '')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/\s{3,}/g, '\n\n')
    .trim();
}

function preprocessBody(text) {
  if (!text) return '';
  return text
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\u2026/g, '...')
    .replace(/[•·●◆▶►]\s*/g, '')
    .replace(/(-{3,}|={3,}|\*{3,})/g, '')
    .split('\n')
    .filter(function(line, i, arr) {
      var t = line.trim();
      if (t.length === 0) return false;
      if (/[\d$%]/.test(t)) return true;
      var lower = t.toLowerCase();
      if (CONFIG.DEAL_KEYWORDS.some(function(kw) { return lower.includes(kw); })) return true;
      var prev = (arr[i - 1] || '').toLowerCase();
      var next = (arr[i + 1] || '').toLowerCase();
      return t.length < 80 && CONFIG.DEAL_KEYWORDS.some(function(kw) { return prev.includes(kw) || next.includes(kw); });
    })
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}


// ---------------------------------------------------------------------------
// UTILITIES
// ---------------------------------------------------------------------------

function showConfig() {
  var senders = CONFIG.NEWSLETTER_QUERY.split(' OR ').map(function(s) { return s.replace('from:', '').trim(); });
  Logger.log('Configured newsletter senders:\n' + senders.join('\n'));
  try { SpreadsheetApp.getUi().alert('Monitoring these senders:\n\n' + senders.join('\n')); } catch(e) {}
}

function reprocessDate(dateStr) {
  // [Change 2] One-off recovery: re-scan a single day's threads without
  // consulting or updating the Processed cache, so a failed Gemini call
  // earlier that day can't permanently block recovery. Safe to re-run —
  // filterAndWrite() still dedups against existing hashes.

  // Parse 'YYYY-MM-DD' as local (script-timezone) components rather than
  // via new Date(dateStr), which parses as UTC midnight and can land on
  // the wrong calendar day once reformatted into the script's timezone.
  var targetDate;
  if (dateStr) {
    var parts = dateStr.split(/[-\/]/).map(Number);
    targetDate = new Date(parts[0], parts[1] - 1, parts[2]);
  } else {
    targetDate = new Date();
  }
  var dayStart = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  var nextDay = new Date(targetDate);
  nextDay.setDate(nextDay.getDate() + 1);
  var dayEnd = Utilities.formatDate(nextDay, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  var query = '(' + CONFIG.NEWSLETTER_QUERY + ') after:' + dayStart + ' before:' + dayEnd;

  var processedHashes = getExistingHashes();

  var response = Gmail.Users.Threads.list('me', {
    q: query,
    maxResults: 50,
  });

  if (!response.threads || response.threads.length === 0) {
    Logger.log('reprocessDate: no matching threads found for ' + dayStart);
    return;
  }

  var batchBuffer = [];
  var totalDeals  = 0;

  for (var i = 0; i < response.threads.length; i++) {
    var thread   = GmailApp.getThreadById(response.threads[i].id);
    var messages = thread.getMessages();

    for (var j = 0; j < messages.length; j++) {
      var emailData = extractEmailData(messages[j]);
      if (!emailData) continue;
      batchBuffer.push(emailData);

      if (batchBuffer.length >= CONFIG.BATCH_SIZE) {
        var deals = callGeminiForDeals(batchBuffer);
        totalDeals += filterAndWrite(deals, processedHashes);
        batchBuffer = [];
        Utilities.sleep(1000);
      }
    }
    // Intentionally no markThreadProcessed() call here — this function
    // never reads or writes the Processed cache.
  }

  if (batchBuffer.length > 0) {
    var deals = callGeminiForDeals(batchBuffer);
    totalDeals += filterAndWrite(deals, processedHashes);
  }

  Logger.log('reprocessDate complete for ' + dayStart + '. New deals written: ' + totalDeals);
}
