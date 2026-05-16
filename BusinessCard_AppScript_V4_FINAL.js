// ==========================================
// CONFIGURATION
// ==========================================
const OPENAI_API_KEY  = "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";
const DRIVE_FOLDER_ID = "1WPNIJ2kQ2xj-RYZmxDP5uaeAyMnFxaaG";

// ── Models ────────────────────────────────────────────────────────────────────
// EXTRACT  : gpt-4o  — vision, card image OCR  (Chat Completions, JSON mode)
// ENRICH   : gpt-5.5 — Responses API + web_search tool, agentic search
//            Docs ref: https://developers.openai.com/api/docs/guides/tools-web-search
const EXTRACT_MODEL = "gpt-4o";
const ENRICH_MODEL  = "gpt-5.5";

// ── Endpoints ─────────────────────────────────────────────────────────────────
const CHAT_ENDPOINT      = "https://api.openai.com/v1/chat/completions"; // for vision extraction
const RESPONSES_ENDPOINT = "https://api.openai.com/v1/responses";        // for web search enrichment

const SHEET_NAME        = "Business card";
const GITHUB_HOSTED_URL = "https://laserpowe.github.io/visiting-card/";

const SHEET_HEADERS = [
  "CREATED_AT", "DRIVE_FILE_URL", "FULL_NAME", "DESIGNATION",
  "COMPANY", "DEPARTMENT", "MOBILE", "OFFICE_PHONES", "FAX", "EMAILS",
  "WEBSITE", "ADDRESS", "CITY", "PIN_CODE", "COUNTRY", "CONFIDENCE",
  "NOTES", "RAW_SUMMARY",
  "LINKEDIN_URL", "LINKEDIN_PHOTO_URL", "LINKEDIN_POSITION",
  "OTHER_WEB_PROFILES", "COMPANY_CORE_BUSINESS", "ENRICH_SOURCES"
];

// ==========================================
// 1. MENU
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📷 Card Desk")
    .addItem("📷 Upload Business Cards",               "openExternalScanner")
    .addSeparator()
    .addItem("🔍 Enrich Selected Row",                 "enrichSelectedRow")
    .addItem("🔁 Enrich ALL Rows",                     "enrichAllRows")
    .addToUi();
}

function openExternalScanner() {
  const html = `<script>window.open('${GITHUB_HOSTED_URL}','_blank');google.script.host.close();</script>
    <div style="font-family:sans-serif;text-align:center;padding:20px">
      <p>Opening scanner...</p>
      <p style="font-size:12px;color:gray">If not, <a href="${GITHUB_HOSTED_URL}" target="_blank">click here</a></p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(300).setHeight(150), "Redirecting..."
  );
}

// ==========================================
// 2. WEB APP — doGet (Profile Directory)
// ==========================================
function doGet(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  // Detect request format: ?format=json for static frontend, else HTML
  const wantsJson = e && e.parameter && e.parameter.format === "json";
  
  if (!sheet) {
    if (wantsJson) {
      return ContentService.createTextOutput(JSON.stringify({error: "Sheet not found", cards: []}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput("Sheet not found.");
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    if (wantsJson) {
      return ContentService.createTextOutput(JSON.stringify({cards: [], count: 0}))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput("No cards found.");
  }
  
  const headers = data[0];
  const cards   = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse();
  
  // ─── JSON response for static GitHub Pages frontend ────────────────────────
  if (wantsJson) {
    return ContentService
      .createTextOutput(JSON.stringify({
        cards: cards,
        count: cards.length,
        generated_at: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // ─── Default: HTML view (existing behavior, kept as fallback) ──────────────
  return HtmlService.createHtmlOutput(renderProfileView_(cards))
    .setTitle("Business Card Directory")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function renderProfileView_(cards) {
  // Helper: extract slug from LinkedIn URL → derive unavatar.io photo URL
  // Example: https://in.linkedin.com/in/shaurin-shah-608b6b15 → shaurin-shah-608b6b15
  function deriveLinkedInPhoto(liUrl) {
    if (!liUrl) return "";
    const m = String(liUrl).match(/linkedin\.com\/in\/([a-zA-Z0-9\-_%\.]+)/i);
    if (!m) return "";
    return "https://unavatar.io/linkedin/" + m[1];
  }

  const cardHtml = cards.map(c => {
    // Priority: 1) explicit LINKEDIN_PHOTO_URL  2) unavatar.io derived from slug  3) initials placeholder
    const derivedPhoto = deriveLinkedInPhoto(c.LINKEDIN_URL);
    const finalPhoto   = c.LINKEDIN_PHOTO_URL || derivedPhoto;
    
    // Initials fallback (e.g. "Shaurin Shah" → "SS")
    const initials = (c.FULL_NAME || "?")
      .split(/\s+/).filter(Boolean).slice(0, 2)
      .map(w => w[0].toUpperCase()).join("");
    
    const photo = finalPhoto
      ? `<img src="${finalPhoto}" class="avatar" onerror="this.outerHTML='<div class=\'avatar-ph\'>${initials}</div>'">`
      : `<div class="avatar-ph">${initials}</div>`;
    const others = (c.OTHER_WEB_PROFILES || "").split("\n")
      .map(u => u.trim()).filter(Boolean)
      .map(u => `<a href="${u}" target="_blank" class="ol">🔗 ${u}</a>`).join("");
    const liLink = c.LINKEDIN_URL
      ? `<a href="${c.LINKEDIN_URL}" target="_blank">🔍 View Profile</a>`
      : "Not Found";
    const siteUrl = c.WEBSITE
      ? (c.WEBSITE.startsWith("http") ? c.WEBSITE : "https://" + c.WEBSITE)
      : "";
    return `
    <div class="card">
      <div class="hdr">${photo}
        <div><h3>${c.FULL_NAME || "Unknown"}</h3>
        <p class="desig">${c.DESIGNATION || ""}${c.COMPANY ? " @ " + c.COMPANY : ""}</p></div>
      </div>
      <div class="body">
        <p>📧 ${c.EMAILS || "N/A"}</p>
        <p>📱 ${c.MOBILE || "N/A"}</p>
        <p>🌐 ${siteUrl ? `<a href="${siteUrl}" target="_blank">${c.WEBSITE}</a>` : "N/A"}</p>
        <div class="eb">
          <p><b>LinkedIn:</b> ${liLink}</p>
          <p class="sm">${c.LINKEDIN_POSITION || "No summary"}</p>
          ${others ? `<div class="op"><b>Other:</b><br>${others}</div>` : ""}
        </div>
        <p class="biz">${c.COMPANY_CORE_BUSINESS || "N/A"}</p>
        ${c.ENRICH_SOURCES ? `<p class="src">Sources: ${c.ENRICH_SOURCES}</p>` : ""}
      </div>
      <a class="btn" href="${c.DRIVE_FILE_URL}" target="_blank">View Original Card</a>
    </div>`;
  }).join("");

  return `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <style>
    body{font-family:"Segoe UI",sans-serif;background:#f4f7f9;padding:20px;color:#333;margin:0}
    .wrap{display:flex;flex-wrap:wrap;gap:20px;justify-content:center;max-width:1400px;margin:0 auto}
    .card{background:#fff;border-radius:12px;box-shadow:0 4px 15px rgba(0,0,0,.06);width:340px;padding:20px;display:flex;flex-direction:column;transition:.3s}
    .card:hover{transform:translateY(-5px);box-shadow:0 8px 25px rgba(0,0,0,.12)}
    .hdr{display:flex;align-items:center;gap:12px;border-bottom:2px solid #f0f2f5;padding-bottom:12px;margin-bottom:12px}
    .avatar{width:60px;height:60px;border-radius:50%;object-fit:cover;border:2px solid #1a73e8;flex-shrink:0}
    .avatar-ph{width:60px;height:60px;border-radius:50%;background:linear-gradient(135deg,#1a73e8,#4285f4);color:#fff;display:flex;align-items:center;justify-content:center;font-size:20px;font-weight:bold;flex-shrink:0;letter-spacing:1px}
    h3{margin:0;color:#1a73e8;font-size:1.15em}.desig{margin:3px 0 0;color:#5f6368;font-size:.84em;font-weight:600}
    .body p{margin:7px 0;font-size:.87em;word-break:break-word}
    .eb{background:#e8f0fe;padding:12px;border-radius:8px;margin:12px 0;border-left:4px solid #1a73e8}
    .sm{font-size:.82em;color:#444;line-height:1.5;margin-top:6px!important}
    .op{margin-top:8px;font-size:.8em}.ol{display:block;color:#1a73e8;word-break:break-all;margin-bottom:3px}
    .biz{font-size:.83em;color:#555;background:#f8f9fa;padding:10px;border-radius:8px;line-height:1.5}
    .src{font-size:.72em;color:#999;font-style:italic}
    .btn{display:block;text-align:center;background:#1a73e8;color:#fff;text-decoration:none;padding:11px;border-radius:6px;margin-top:15px;font-weight:bold}
    .btn:hover{background:#1557b0}h1{text-align:center;color:#1a73e8;margin-bottom:28px;font-size:2em}
  </style></head>
  <body><h1>📇 Business Card Directory</h1><div class="wrap">${cardHtml}</div></body></html>`;
}

// ==========================================
// 3. doPost — Receive card from web scanner
// ==========================================
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const result = mainProcessingLogic_(data.base64, data.fileName, data.mimeType);
    return ContentService.createTextOutput(JSON.stringify({ success: true, message: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 4. MAIN PROCESSING PIPELINE
// ==========================================
function mainProcessingLogic_(base64Data, fileName, mimeType) {
  // Step 1: Upload image to Drive
  const folder   = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const blob     = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file     = folder.createFile(blob);
  const driveUrl = file.getUrl();

  // Step 2: Extract card text via GPT-4o vision (Chat Completions)
  const extraction = extractWithVision_(base64Data, mimeType);

  // Step 3: Prepare sheet
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  ensureHeaders_(sheet);

  const safe = v => {
    if (!v) return "";
    const s = String(v).trim();
    return /^[+=\-@]/.test(s) ? "'" + s : s;
  };

  const rowData = {
    CREATED_AT:    new Date().toISOString(),
    DRIVE_FILE_URL: driveUrl,
    FULL_NAME:     extraction.full_name     || "",
    DESIGNATION:   extraction.designation   || "",
    COMPANY:       extraction.company       || "",
    DEPARTMENT:    extraction.department    || "",
    MOBILE:        extraction.mobile        || "",
    OFFICE_PHONES: Array.isArray(extraction.office_phones)
                     ? extraction.office_phones.join(", ")
                     : String(extraction.office_phones || ""),
    FAX:           extraction.fax           || "",
    EMAILS:        Array.isArray(extraction.emails)
                     ? extraction.emails.join(", ")
                     : String(extraction.emails || ""),
    WEBSITE:       extraction.website       || "",
    ADDRESS:       extraction.address       || "",
    CITY:          extraction.city          || "",
    PIN_CODE:      extraction.pin_code      || "",
    COUNTRY:       extraction.country       || "",
    CONFIDENCE:    extraction.confidence    || "",
    NOTES:         extraction.notes         || "",
    RAW_SUMMARY:   extraction.raw_summary   || "",
    LINKEDIN_URL: "", LINKEDIN_PHOTO_URL: "", LINKEDIN_POSITION: "",
    OTHER_WEB_PROFILES: "", COMPANY_CORE_BUSINESS: "", ENRICH_SOURCES: ""
  };

  sheet.appendRow(SHEET_HEADERS.map(h => safe(rowData[h])));
  const newRow = sheet.getLastRow();

  // Step 4: Enrich via Responses API + web search
  enrichRowByNumber_(sheet, newRow,
    extraction.full_name, extraction.designation, extraction.company);

  return "Card processed and enriched!";
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length)
      .setFontWeight("bold").setBackground("#4a86e8").setFontColor("#ffffff");
    return;
  }
  const cur = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  ["LINKEDIN_URL","LINKEDIN_PHOTO_URL","LINKEDIN_POSITION",
   "OTHER_WEB_PROFILES","COMPANY_CORE_BUSINESS","ENRICH_SOURCES"].forEach(col => {
    if (!cur.includes(col)) {
      const nc = sheet.getLastColumn() + 1;
      sheet.getRange(1, nc).setValue(col)
        .setFontWeight("bold").setBackground("#4a86e8").setFontColor("#ffffff");
    }
  });
}

// ==========================================
// 5. MENU — Enrich Selected Row
// ==========================================
function enrichSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Sheet not found."); return; }
  const row = sheet.getActiveRange().getRow();
  if (row <= 1) { SpreadsheetApp.getUi().alert("Select a data row (not header)."); return; }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const vals    = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const g = col => { const i = headers.indexOf(col); return i >= 0 ? String(vals[i]).trim() : ""; };
  const name = g("FULL_NAME"), desig = g("DESIGNATION"), co = g("COMPANY");
  if (!name && !co) { SpreadsheetApp.getUi().alert("FULL_NAME and COMPANY both empty."); return; }
  enrichRowByNumber_(sheet, row, name, desig, co);
  SpreadsheetApp.getUi().alert("✅ Row " + row + " enriched!");
}

// ==========================================
// 6. MENU — Enrich All Rows
// ==========================================
function enrichAllRows() {
  const ui  = SpreadsheetApp.getUi();
  const res = ui.alert("Enrich ALL Rows",
    "Unenriched rows will be processed. This uses gpt-5.5 with live web search.\nContinue?",
    ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { ui.alert("No data rows."); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let done = 0, skip = 0;

  for (let r = 2; r <= lastRow; r++) {
    const vals = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
    const g    = col => { const i = headers.indexOf(col); return i >= 0 ? String(vals[i]).trim() : ""; };
    const name = g("FULL_NAME"), desig = g("DESIGNATION"), co = g("COMPANY");
    if (!name && !co) { skip++; continue; }
    // Skip only if all three enrichment fields are already filled
    if (g("LINKEDIN_URL") && g("COMPANY_CORE_BUSINESS") && g("LINKEDIN_POSITION")) { skip++; continue; }
    enrichRowByNumber_(sheet, r, name, desig, co);
    done++;
    Utilities.sleep(4000); // gpt-5.5 agentic search needs buffer time
  }
  ui.alert("✅ Done!\nProcessed: " + done + "\nSkipped: " + skip);
}

// ==========================================
// 7. ENRICHMENT CORE — Responses API + web_search
// ==========================================
function enrichRowByNumber_(sheet, rowNum, fullName, designation, company) {
  try {
    const data    = enrichWithResponsesAPI_(fullName, designation, company);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const set     = (col, val) => {
      const i = headers.indexOf(col);
      if (i >= 0) sheet.getRange(rowNum, i + 1).setValue(val || "");
    };
    set("LINKEDIN_URL",          data.linkedin_url);
    set("LINKEDIN_PHOTO_URL",    data.linkedin_photo_url);
    set("LINKEDIN_POSITION",     data.linkedin_position);
    set("OTHER_WEB_PROFILES",    data.other_web_profiles);
    set("COMPANY_CORE_BUSINESS", data.company_core_business);
    set("ENRICH_SOURCES",        data.enrich_sources);
    Logger.log("✅ Enriched row " + rowNum + " — LinkedIn: " + data.linkedin_url);
  } catch (err) {
    Logger.log("❌ Enrich error row " + rowNum + ": " + err.message);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const i = headers.indexOf("ENRICH_SOURCES");
    if (i >= 0) sheet.getRange(rowNum, i + 1).setValue("ERROR: " + err.message.substring(0, 300));
  }
}

/**
 * enrichWithResponsesAPI_
 * ────────────────────────
 * Uses the NEW OpenAI Responses API (v1/responses) with:
 *   - model: gpt-5.5 (best reasoning + web search model per docs)
 *   - tools: [{ type: "web_search" }]  ← correct per official docs
 *   - reasoning: { effort: "high" }    ← agentic search mode
 *
 * Docs: https://developers.openai.com/api/docs/guides/tools-web-search
 *
 * The model will:
 *   1. Search Google for the person's LinkedIn profile
 *   2. Open the LinkedIn page if found
 *   3. Search for the company's website
 *   4. Return structured data with citations
 *
 * NO HALLUCINATION: if URL can't be verified via live web, it returns
 * a LinkedIn search URL instead of a fake profile link.
 */
function enrichWithResponsesAPI_(fullName, designation, company) {
  // ── Normalize company name — strip legal suffixes for better search ──
  // "Enam Holdings Pvt. Ltd." → "Enam Holdings"
  // "ABC Industries Private Limited" → "ABC Industries"
  function stripCompanySuffix(name) {
    if (!name) return "";
    return String(name)
      .replace(/\s*(Pvt\.?|Private)\s+(Ltd\.?|Limited)\.?$/i, "")
      .replace(/\s+(Ltd\.?|Limited|LLC|LLP|Inc\.?|Corp\.?|Corporation|Co\.?)\.?$/i, "")
      .replace(/[\.,]+$/, "")
      .trim();
  }
  var companyShort = stripCompanySuffix(company);  // "ENAM Holdings"
  
  // Pre-build fallback Google search URL (use SHORTENED company name, no quotes around company)
  var nameQ       = fullName ? '"' + fullName + '"' : '';
  var googleQuery = encodeURIComponent('site:linkedin.com/in ' + [nameQ, companyShort].filter(Boolean).join(' '));
  var liSearchUrl = "https://www.google.com/search?q=" + googleQuery;
  
  Logger.log("Company normalized: \"" + company + "\" → \"" + companyShort + "\"");

  // Build prompt
  var inputText =
    "You are a professional researcher. Use your web search tool to find REAL, VERIFIED information.\n\n" +
    "PERSON TO RESEARCH:\n" +
    "- Full Name: " + (fullName    || "Unknown") + "\n" +
    "- Designation: " + (designation || "Unknown") + "\n" +
    "- Company: " + (company      || "Unknown") + "\n\n" +

    "═══════════════════════════════════════════════════════════════════════\n" +
    "STEP 1 — FIND THE EXACT LINKEDIN PROFILE URL (MANDATORY)\n" +
    "═══════════════════════════════════════════════════════════════════════\n" +
    "Your #1 priority: find the URL matching pattern linkedin.com/in/<slug>\n" +
    "Example of what you MUST return: https://in.linkedin.com/in/shaurin-shah-608b6b15\n\n" +
    "Company on LinkedIn often differs from business card:\n" +
    "   Card: \"" + (company || "") + "\"  →  LinkedIn: \"" + companyShort + "\"\n\n" +
    "Run these searches in order until you find a linkedin.com/in/<slug> URL:\n" +
    "  1. site:linkedin.com/in \"" + fullName + "\" " + companyShort + "\n" +
    "  2. site:linkedin.com/in \"" + fullName + "\"\n" +
    "  3. \"" + fullName + "\" " + companyShort + " linkedin\n" +
    "  4. \"" + fullName + "\" linkedin\n\n" +
    "⛔ DO NOT use exact-quote on \"Pvt. Ltd.\" / \"Private Limited\" — searches fail.\n" +
    "⛔ DO NOT return a google.com/search URL — it will be REJECTED by validation.\n" +
    "⛔ DO NOT return linkedin.com/search/... — it will be REJECTED.\n" +
    "✅ ONLY return https://[xx.]linkedin.com/in/<slug> format.\n" +
    "✅ If you genuinely cannot find the profile after all 4 searches, return empty string \"\".\n\n" +

    "═══════════════════════════════════════════════════════════════════════\n" +
    "STEP 2 — DEEP PERSONAL DETAILS (POINT BY POINT)\n" +
    "═══════════════════════════════════════════════════════════════════════\n" +
    "From LinkedIn snippet + any other web sources (Crunchbase, news, company bio, MarketScreener, SignalHire):\n" +
    "Compile detailed bullet points about the PERSON. Format EXACTLY like this:\n\n" +
    "• Current Role: <title> at <company> (<location if known>)\n" +
    "• Career Trajectory:\n" +
    "    - <Role> at <Company> (<duration / dates>)\n" +
    "    - <Role> at <Company> (<duration / dates>)\n" +
    "    - <Role> at <Company> (<duration / dates>)\n" +
    "• Education:\n" +
    "    - <Degree>, <Institution>, <Year if known>\n" +
    "• Expertise / Focus Areas: <skills, domains, specializations>\n" +
    "• Notable: <publications, talks, awards, recognition, certifications>\n" +
    "• Network: <e.g. 500+ connections, board memberships>\n\n" +
    "If a bullet has no data, write \"• <bullet name>: Not publicly available\".\n" +
    "Use ONLY information confirmed from your live web searches — never invent.\n\n" +

    "═══════════════════════════════════════════════════════════════════════\n" +
    "STEP 3 — DEEP COMPANY PROFILE (POINT BY POINT)\n" +
    "═══════════════════════════════════════════════════════════════════════\n" +
    "Visit the official website of \"" + companyShort + "\" (also \"" + (company || "") + "\").\n" +
    "Read About / Products / Services / Leadership pages. Compile EXACTLY this format:\n\n" +
    "• Overview: <one-line summary of what they do>\n" +
    "• Founded: <year> by <founders if known>\n" +
    "• Headquarters: <city, country>\n" +
    "• Core Business: <main activities, business model>\n" +
    "• Products / Services:\n" +
    "    - <Product/Service 1>\n" +
    "    - <Product/Service 2>\n" +
    "    - <Product/Service 3>\n" +
    "• Industry / Sector: <industry classification>\n" +
    "• Key Leadership:\n" +
    "    - <Name> (<Role>)\n" +
    "    - <Name> (<Role>)\n" +
    "• Market Position: <size, ranking, key competitors if known>\n" +
    "• Recent Highlights: <news, expansions, milestones, fund activity>\n" +
    "• Subsidiaries / Divisions: <if any>\n" +
    "• Official Website: <URL>\n\n" +
    "If a bullet has no data, write \"• <bullet name>: Not publicly available\".\n\n" +

    "═══════════════════════════════════════════════════════════════════════\n" +
    "STEP 4 — OTHER WEB PROFILES\n" +
    "═══════════════════════════════════════════════════════════════════════\n" +
    "Search for other profiles: Twitter/X, Crunchbase, MarketScreener, SignalHire, company bio page, news mentions.\n" +
    "Return ONLY URLs that actually appeared in search results.\n\n" +

    "═══════════════════════════════════════════════════════════════════════\n" +
    "CRITICAL RULES (NON-NEGOTIABLE)\n" +
    "═══════════════════════════════════════════════════════════════════════\n" +
    "1. linkedin_url MUST match: https://[xx.]linkedin.com/in/<slug>  — anything else = REJECTED\n" +
    "2. NEVER invent a slug. If unsure, return empty string \"\".\n" +
    "3. linkedin_photo_url ONLY from media.licdn.com or similar real CDN.\n" +
    "4. linkedin_position and company_core_business MUST be in bullet-point format above.\n" +
    "5. All data must come from REAL web search results — no hallucination.\n\n" +

    "Respond with ONLY this JSON (no markdown fences, no extra text):\n" +
    "{\n" +
    "  \"linkedin_url\": \"https://...linkedin.com/in/<slug> OR empty string\",\n" +
    "  \"linkedin_photo_url\": \"<CDN photo URL or empty>\",\n" +
    "  \"linkedin_position\": \"• Current Role: ...\\n• Career Trajectory:\\n    - ...\\n• Education: ...\\n• Expertise: ...\\n• Notable: ...\\n• Network: ...\",\n" +
    "  \"other_web_profiles\": \"<verified URLs, one per line>\",\n" +
    "  \"company_core_business\": \"• Overview: ...\\n• Founded: ...\\n• Headquarters: ...\\n• Core Business: ...\\n• Products/Services:\\n    - ...\\n• Industry: ...\\n• Key Leadership:\\n    - ...\\n• Market Position: ...\\n• Recent Highlights: ...\\n• Subsidiaries: ...\\n• Official Website: ...\",\n" +
    "  \"enrich_sources\": \"<all URLs you actually visited, comma-separated>\"\n" +
    "}";

  // Responses API call
  var payload = {
    model:     ENRICH_MODEL,
    input:     inputText,
    reasoning: { effort: "high" },
    tools: [{ type: "web_search", search_context_size: "high" }],
    tool_choice:    "required",
    max_output_tokens: 15000
  };

  var options = {
    method:             "post",
    headers:            { "Authorization": "Bearer " + OPENAI_API_KEY },
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  Logger.log("Calling Responses API for: " + fullName + " @ " + company);
  var response     = UrlFetchApp.fetch(RESPONSES_ENDPOINT, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();

  Logger.log("Responses API code: " + responseCode);

  if (responseCode !== 200) {
    Logger.log("Responses API error: " + responseText.substring(0, 500));
    return buildFallbackResult_(liSearchUrl, "Responses API error: " + responseCode);
  }

  // PARSING — BULLETPROOF
  var result;
  try {
    result = JSON.parse(responseText);
  } catch (parseErr) {
    Logger.log("Failed to parse response JSON: " + parseErr);
    return buildFallbackResult_(liSearchUrl, "Response parse error");
  }

  // Extract text from output array
  var rawText = "";
  
  if (result.output && Array.isArray(result.output)) {
    for (var i = 0; i < result.output.length; i++) {
      var item = result.output[i];
      
      // Skip non-message items
      if (item.type !== "message") {
        continue;
      }
      
      // Get content array
      if (!item.content || !Array.isArray(item.content)) {
        continue;
      }
      
      // Extract text from content blocks
      for (var j = 0; j < item.content.length; j++) {
        var block = item.content[j];
        if (block && block.type === "output_text" && block.text) {
          rawText += block.text;
        }
      }
    }
  }

  Logger.log("Extracted text length: " + rawText.length);
  if (rawText.length > 0) {
    Logger.log("Extracted text (first 300 chars): " + rawText.substring(0, 300));
  }

  // Check if response was incomplete (token limit hit)
  if (result.status === "incomplete") {
    var reason = result.incomplete_details ? result.incomplete_details.reason : "unknown";
    Logger.log("⚠️  Response INCOMPLETE - reason: " + reason);
    if (reason === "max_output_tokens") {
      Logger.log("→ Model ran out of tokens. Increase max_output_tokens or reduce reasoning effort.");
    }
  }

  // If no text found, log full response and return fallback
  if (rawText.length === 0) {
    Logger.log("WARNING: No text extracted from output array");
    Logger.log("Full response (first 1500): " + responseText.substring(0, 1500));
    return buildFallbackResult_(liSearchUrl, "No text in Responses API output");
  }

  // Parse and validate
  return parseAndValidate_(rawText, liSearchUrl);
}



// ==========================================
// 8. PARSE + VALIDATE MODEL RESPONSE
// ==========================================
function parseAndValidate_(rawText, liSearchUrl) {
  // Remove markdown fences if present
  const cleaned   = rawText.replace(/```json[\s\S]*?```|```[\s\S]*?```/g, "").trim();
  const jsonMatch = cleaned.match(/\{[\s\S]*\}/);

  if (!jsonMatch) {
    Logger.log("No JSON found. Raw text: " + rawText.substring(0, 300));
    return buildFallbackResult_(liSearchUrl, "No JSON in response");
  }

  let parsed;
  try {
    parsed = JSON.parse(jsonMatch[0]);
  } catch (e) {
    Logger.log("JSON parse error: " + e.message);
    return buildFallbackResult_(liSearchUrl, "JSON parse failed");
  }

  // ── Safe string converter (handles string / array / object / null) ──────────
  const str = v => {
    if (v === null || v === undefined) return "";
    if (typeof v === "string")         return v.trim();
    if (Array.isArray(v))              return v.map(x => String(x).trim()).join(", ");
    return String(v).trim();
  };

  // ── Validate LinkedIn URL — STRICT: must be linkedin.com/in/<slug> ──────────
  let liUrl = str(parsed.linkedin_url);
  if (liUrl && !isValidLinkedInUrl_(liUrl)) {
    Logger.log("⚠️ Invalid LinkedIn URL rejected (not linkedin.com/in/...): " + liUrl);
    liUrl = "";  // Leave empty rather than putting a search URL
  }

  // ── Validate photo URL ───────────────────────────────────────────────────────
  let photoUrl = str(parsed.linkedin_photo_url);
  if (photoUrl && !isValidPhotoUrl_(photoUrl)) {
    Logger.log("⚠️ Invalid photo URL rejected: " + photoUrl);
    photoUrl = "";
  }

  // ── Clean other profiles (remove invalid URLs) ────────────────────────────────
  const rawOthers  = str(parsed.other_web_profiles);
  const otherClean = rawOthers.split(/[\n,]/)
    .map(u => u.trim()).filter(u => u && isValidHttpUrl_(u)).join("\n");

  return {
    linkedin_url:          liUrl,
    linkedin_photo_url:    photoUrl,
    linkedin_position:     str(parsed.linkedin_position),
    other_web_profiles:    otherClean,
    company_core_business: str(parsed.company_core_business),
    enrich_sources:        str(parsed.enrich_sources)
  };
}

function buildFallbackResult_(liSearchUrl, reason) {
  Logger.log("Using fallback. Reason: " + reason);
  return {
    linkedin_url:          liSearchUrl || "",
    linkedin_photo_url:    "",
    linkedin_position:     "Enrichment issue: " + reason + ". Check LinkedIn manually.",
    other_web_profiles:    "",
    company_core_business: "Enrichment issue. Please visit company website directly.",
    enrich_sources:        "Fallback: " + reason
  };
}

// ==========================================
// 9. URL VALIDATORS
// ==========================================
function isValidLinkedInUrl_(url) {
  if (!url) return false;
  // STRICT: ONLY accept real LinkedIn profile URLs: linkedin.com/in/<slug>
  // Handles: www.linkedin.com, in.linkedin.com, uk.linkedin.com, etc.
  // REJECTS: google.com/search, linkedin.com/search, fake URLs
  return /^https?:\/\/([a-z]{2,}\.)?linkedin\.com\/in\/[a-zA-Z0-9\-_%\.]+\/?(\?.*)?$/i.test(url.trim());
}

function isValidPhotoUrl_(url) {
  if (!url) return false;
  const allowed = [
    "media.licdn.com", "media-exp1.licdn.com", "media-exp2.licdn.com",
    "static.licdn.com", "pbs.twimg.com", "unavatar.io",
    "gravatar.com", "githubusercontent.com", "lh3.googleusercontent.com"
  ];
  try {
    const domain = url.replace(/^https?:\/\//, "").split("/")[0].toLowerCase();
    return allowed.some(d => domain.includes(d));
  } catch (e) { return false; }
}

function isValidHttpUrl_(url) {
  try { return /^https?:\/\/.+\..+/.test(url.trim()); }
  catch (e) { return false; }
}

// ==========================================
// 10. VISION EXTRACTION — Business Card OCR
//     Uses Chat Completions (gpt-4o + vision)
// ==========================================
function extractWithVision_(base64Data, mimeType) {
  const prompt =
    "You are an expert at reading business cards. Extract ALL visible information precisely.\n\n" +
    "Return JSON with these exact keys:\n" +
    "- full_name: Complete person name exactly as printed\n" +
    "- designation: Job title/role exactly as printed\n" +
    "- company: Company/organization name\n" +
    "- department: Department if mentioned, else empty string\n" +
    "- mobile: Mobile/cell numbers (keep + prefix), comma-separate if multiple\n" +
    "- office_phones: Array of office/landline numbers\n" +
    "- fax: Fax number or empty string\n" +
    "- emails: Array of all email addresses\n" +
    "- website: Website URL(s)\n" +
    "- address: Complete street address\n" +
    "- city: City name\n" +
    "- pin_code: ZIP/postal code\n" +
    "- country: Country or empty string\n" +
    "- confidence: 'High' | 'Medium' | 'Low' based on card readability\n" +
    "- notes: Taglines, social handles, QR info, anything not fitting above fields\n" +
    "- raw_summary: Complete verbatim text on card, line by line\n\n" +
    "Use empty string for missing text fields. Empty array for missing arrays.";

  const payload = {
    model: EXTRACT_MODEL,  // gpt-4o (vision)
    messages: [{
      role: "user",
      content: [
        { type: "text",      text: prompt },
        { type: "image_url", image_url: { url: "data:" + mimeType + ";base64," + base64Data } }
      ]
    }],
    response_format: { type: "json_object" },
    max_tokens: 1200
  };

  const options = {
    method:             "post",
    headers:            { "Authorization": "Bearer " + OPENAI_API_KEY },
    contentType:        "application/json",
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(CHAT_ENDPOINT, options);
  if (resp.getResponseCode() !== 200) {
    throw new Error("Vision extraction failed (" + resp.getResponseCode() + "): " + resp.getContentText().substring(0, 200));
  }
  return JSON.parse(JSON.parse(resp.getContentText()).choices[0].message.content);
}
