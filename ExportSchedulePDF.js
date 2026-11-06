// =============================================================================
// EXPORT & EMAIL MONTHLY SCHEDULE (PDF)
// =============================================================================
// Exports ONLY next month's generated tab as a PDF, archives it in the same
// Drive folder as this spreadsheet, and emails it to a distribution list.
// Sent manually from the menu, after council review. The email is sent from the
// account of the coordinator who runs the menu action (MailApp default).
//
// The email wording is NOT in code. It is read from four tabs that are created
// automatically (never overwritten) the first time they are needed:
//   EmailSettings  - key/value: names, reply email, languages, subject
//   EmailTemplate  - one row per block of text (English + Chinese columns)
//   EmailLinks     - bulleted "Note the following" items (English + Chinese)
//   Coordinators   - team coordinators shown at the bottom of the email
//   PublishedLinks - one "Publish to web" link per month tab (see resolvePublishedLink_)
// Text supports {tokens} and Markdown-style links: [label](https://...).
// =============================================================================

var PDF_DISTRIBUTION_PROPERTY = 'PDF_DISTRIBUTION_LIST';
var PDF_MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
                       "July", "August", "September", "October", "November", "December"];

var EMAIL_TAB_SETTINGS     = "EmailSettings";
var EMAIL_TAB_TEMPLATE     = "EmailTemplate";
var EMAIL_TAB_LINKS        = "EmailLinks";
var EMAIL_TAB_COORDINATORS = "Coordinators";
var EMAIL_TAB_PUBLISHED    = "PublishedLinks";

// -----------------------------------------------------------------------------
// MENU ENTRY POINTS
// -----------------------------------------------------------------------------

/** Menu: set/update the distribution list (comma, semicolon or newline separated). */
function promptSetPdfDistributionList() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();
  var current = props.getProperty(PDF_DISTRIBUTION_PROPERTY) || "";
  var resp = ui.prompt(
    "PDF Distribution List",
    "Enter the distribution list email address(es), separated by commas." +
      (current ? "\n\nCurrent: " + current : ""),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var emails = parsePdfRecipients_(resp.getResponseText());
  if (emails.length === 0) {
    ui.alert("No valid email addresses were entered. The list was not changed.");
    return;
  }
  props.setProperty(PDF_DISTRIBUTION_PROPERTY, emails.join(","));
  ui.alert("Saved", "PDF distribution list set to:\n" + emails.join("\n"), ui.ButtonSet.OK);
}

/** Menu: create any missing email config tabs and jump to the template. */
function promptEditEmailTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = ensureEmailConfigTabs_(ss);
  var templateSheet = ss.getSheetByName(EMAIL_TAB_TEMPLATE);
  if (templateSheet && templateSheet.isSheetHidden()) {
    templateSheet.showSheet();
  }
  ss.setActiveSheet(templateSheet);
  SpreadsheetApp.getUi().alert(
    "Email Template",
    (created.length ? "Created tabs: " + created.join(", ") + ".\n\n" : "") +
    "Edit these tabs to change the email wording:\n" +
    "• " + EMAIL_TAB_SETTINGS + " – names, languages, subject\n" +
    "• " + EMAIL_TAB_TEMPLATE + " – paragraphs (English / Chinese columns)\n" +
    "• " + EMAIL_TAB_LINKS + " – 'Note the following' bullets\n" +
    "• " + EMAIL_TAB_COORDINATORS + " – team coordinators (fill in their emails)\n\n" +
    "Use {tokens} and [label](https://link) for links. Use 'Preview Email (send to me)' to test.",
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Menu: send the rendered email (with PDF) to the running user only. Nothing archived. */
function promptPreviewScheduleEmail() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = getNextMonthTarget_(ss);
  var sheet = ss.getSheetByName(target.tabName);
  if (!sheet) {
    ui.alert("Schedule Not Found", "No tab named '" + target.tabName + "' exists yet. Generate it first.", ui.ButtonSet.OK);
    return;
  }
  var me = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
  if (!me) { ui.alert("Could not determine your email address."); return; }

  try {
    var scheduleUrl = resolvePublishedLink_(ss, sheet, target.tabName);
    if (!scheduleUrl) return;
    var email = buildScheduleEmail_(ss, sheet, target, scheduleUrl);
    var pdfBlob = buildMonthPdfBlob_(ss, sheet, email.fileBaseName);
    sendScheduleEmail_(me, "[PREVIEW] " + email.subject, email, pdfBlob, me);
    var msg = "A preview was sent to " + me + ".\nNothing was sent to the distribution list or archived.";
    if (email.warnings.length) msg += "\n\n⚠️ " + email.warnings.join("\n⚠️ ");
    ui.alert("Preview Sent", msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Error", "Could not build the preview:\n" + e.message, ui.ButtonSet.OK);
  }
}

/** Menu: export next month's tab to PDF, archive it, and email it. */
function promptExportAndEmailNextMonthPdf() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = getNextMonthTarget_(ss);

  var sheet = ss.getSheetByName(target.tabName);
  if (!sheet) {
    ui.alert("Schedule Not Found",
      "No tab named '" + target.tabName + "' exists yet.\n\nPlease generate it first with " +
      "'Generate Monthly Schedule' and have the council review it.", ui.ButtonSet.OK);
    return;
  }

  var recipients = parsePdfRecipients_(
    PropertiesService.getScriptProperties().getProperty(PDF_DISTRIBUTION_PROPERTY) || "");
  if (recipients.length === 0) {
    ui.alert("No Distribution List",
      "No distribution list is configured yet.\n\nUse '📅 Schedule Assistant > Set PDF Distribution List' first.",
      ui.ButtonSet.OK);
    return;
  }

  // Resolve this month's published link (prompts once if it isn't saved yet).
  var scheduleUrl = resolvePublishedLink_(ss, sheet, target.tabName);
  if (!scheduleUrl) return; // cancelled or invalid; the user was told why

  // Render first so template problems surface BEFORE anything is sent.
  var email;
  try {
    email = buildScheduleEmail_(ss, sheet, target, scheduleUrl);
  } catch (e) {
    ui.alert("Email Template Problem", e.message, ui.ButtonSet.OK);
    return;
  }

  var sender = Session.getActiveUser().getEmail() || "(your account)";
  var confirm = ui.alert(
    "Export & Email Schedule PDF",
    "Tab: " + target.tabName + "\n" +
    "Subject: " + email.subject + "\n" +
    "To: " + recipients.join(", ") + "\n" +
    "From: " + sender + "\n" +
    "Schedule link: " + email.scheduleUrl + "\n" +
    (email.warnings.length ? "\n⚠️ " + email.warnings.join("\n⚠️ ") + "\n" : "") +
    "\nHas the council reviewed this month's schedule? The PDF will be saved to the " +
    "spreadsheet's Drive folder and emailed now.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    var result = exportAndEmailMonthPdf_(ss, sheet, email, recipients);
    ui.alert("Sent ✅",
      "Emailed '" + result.fileName + "' to:\n" + recipients.join("\n") +
      "\n\nArchived in Drive folder: " + result.folderName + "\n" + result.fileUrl,
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Error", "Could not export/send the PDF:\n" + e.message, ui.ButtonSet.OK);
  }
}

// -----------------------------------------------------------------------------
// CORE: PDF + SEND
// -----------------------------------------------------------------------------

/** Next month relative to today, in the spreadsheet's time zone. */
function getNextMonthTarget_(ss) {
  var tz = ss.getSpreadsheetTimeZone();
  var now = new Date();
  var year = parseInt(Utilities.formatDate(now, tz, "yyyy"), 10);
  var month = parseInt(Utilities.formatDate(now, tz, "M"), 10) + 1;
  if (month > 12) { month = 1; year += 1; }
  return { year: year, month: month, tabName: PDF_MONTH_NAMES[month - 1] + " " + year, tz: tz };
}

/** Archive first (so a copy always exists), then email the list. Throws on failure. */
function exportAndEmailMonthPdf_(ss, sheet, email, recipients) {
  var pdfBlob = buildMonthPdfBlob_(ss, sheet, email.fileBaseName);

  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HHmm");
  var archived = folder.createFile(pdfBlob.copyBlob().setName(email.fileBaseName + " (sent " + stamp + ").pdf"));

  sendScheduleEmail_(recipients.join(","), email.subject, email, pdfBlob,
                     Session.getActiveUser().getEmail());

  if (typeof emailLog === "function") {
    try { emailLog("INFO", "Schedule PDF emailed: " + email.subject, recipients.join(",")); } catch (e) {}
  }
  return { fileName: archived.getName(), fileUrl: archived.getUrl(), folderName: folder.getName() };
}

function sendScheduleEmail_(to, subject, email, pdfBlob, replyTo) {
  var options = {
    to: to,
    subject: subject,
    htmlBody: email.html,
    body: email.text,
    attachments: [pdfBlob],
    name: email.senderName
  };
  if (replyTo) options.replyTo = replyTo;
  MailApp.sendEmail(options);
}

/** Exports just this tab as a landscape, fit-to-width letter PDF. */
function buildMonthPdfBlob_(ss, sheet, fileBaseName) {
  SpreadsheetApp.flush(); // make sure pending edits are in the export
  var params = {
    format: "pdf",
    gid: sheet.getSheetId(),
    size: "letter",
    portrait: "false",
    fitw: "true",
    gridlines: "false",
    sheetnames: "false",
    printtitle: "false",
    pagenumbers: "false",
    top_margin: "0.4", bottom_margin: "0.4", left_margin: "0.4", right_margin: "0.4"
  };
  var query = Object.keys(params).map(function (k) { return k + "=" + params[k]; }).join("&");
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?" + query;

  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error("PDF export failed (HTTP " + response.getResponseCode() + ").");
  }
  return response.getBlob().setName(fileBaseName + ".pdf");
}

// -----------------------------------------------------------------------------
// EMAIL RENDERING (template-driven)
// -----------------------------------------------------------------------------

/**
 * Renders subject + HTML + plain text from the config tabs.
 * Throws a readable Error for unknown {tokens} or missing settings so a broken
 * template can never be emailed to the whole congregation.
 */
function buildScheduleEmail_(ss, sheet, target, scheduleUrl) {
  ensureEmailConfigTabs_(ss);
  var warnings = [];

  var settings = readEmailSettings_(ss);
  if (!scheduleUrl) throw new Error("No published schedule link was provided for " + target.tabName + ".");

  var langs = String(settings.languages || "en,zh").toLowerCase().split(/[\s,;]+/)
    .filter(function (l) { return l === "en" || l === "zh"; });
  if (langs.length === 0) throw new Error("'languages' on " + EMAIL_TAB_SETTINGS + " must include en and/or zh.");

  // Tokens: every EmailSettings key, plus computed values (computed values win).
  var tokens = {};
  Object.keys(settings).forEach(function (k) { tokens[k] = String(settings[k]); });
  tokens.month = PDF_MONTH_NAMES[target.month - 1];
  tokens.month_num = String(target.month);
  tokens.year = String(target.year);
  tokens.month_year = tokens.month + " " + tokens.year;
  tokens.schedule_url = scheduleUrl;

  var subject = fillTokens_(settings.subject || "{month_year} {church_short} Worship Schedule", tokens, "subject");
  var fileBaseName = fillTokens_(settings.pdf_filename || "{month_year} {church_short} Worship Schedule",
                                 tokens, "pdf_filename").replace(/[\\\/:*?"<>|]/g, "-");

  var blocks = readTable_(ss.getSheetByName(EMAIL_TAB_TEMPLATE))
    .filter(function (r) { return isEnabled_(r.enabled) && String(r.block_id || "").trim() !== ""; })
    .sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });
  if (blocks.length === 0) throw new Error("The " + EMAIL_TAB_TEMPLATE + " tab has no enabled blocks.");

  var html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.5;color:#222;">';

  blocks.forEach(function (b) {
    var id = String(b.block_id).trim().toLowerCase();
    var blockHtml = "";

    if (id === "links") {
      blockHtml += renderLinksList_(ss, langs, tokens);
    } else {
      langs.forEach(function (lang) {
        var cell = String(b[lang] === undefined ? "" : b[lang]).trim();
        if (!cell) return;
        var rendered = renderRichText_(fillTokens_(cell, tokens, EMAIL_TAB_TEMPLATE + " / " + id + " / " + lang));
        var style = (id === "coordinators") ? "margin:0 0 4px;font-weight:bold;" : "margin:0 0 6px;";
        blockHtml += '<p style="' + style + '">' + rendered + "</p>";
      });
      if (id === "coordinators") blockHtml += renderCoordinators_(ss, langs, warnings);
    }
    if (blockHtml) html += '<div style="margin:0 0 14px;">' + blockHtml + "</div>";
  });
  html += "</div>";

  return {
    subject: subject,
    html: html,
    text: htmlToText_(html),
    fileBaseName: fileBaseName,
    senderName: settings.sender_name || settings.church_short || "",
    scheduleUrl: scheduleUrl,
    warnings: warnings
  };
}

/** "Note the following" bullets, one <li> per EmailLinks row (English line, then Chinese). */
function renderLinksList_(ss, langs, tokens) {
  var rows = readTable_(ss.getSheetByName(EMAIL_TAB_LINKS))
    .filter(function (r) { return isEnabled_(r.enabled); })
    .sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });

  var items = rows.map(function (r, i) {
    var lines = [];
    langs.forEach(function (lang) {
      var cell = String(r[lang] === undefined ? "" : r[lang]).trim();
      if (cell) lines.push(renderRichText_(fillTokens_(cell, tokens, EMAIL_TAB_LINKS + " row " + (i + 2) + " / " + lang)));
    });
    return lines.length ? '<li style="margin:0 0 6px;">' + lines.join("<br>") + "</li>" : "";
  }).filter(function (s) { return s; });

  return items.length ? '<ul style="margin:0 0 0 18px;padding:0;">' + items.join("") + "</ul>" : "";
}

/** One line per role, e.g. "Usher 接待: Darrus Lee or Naomi Lee". Rows sharing a role are grouped. */
function renderCoordinators_(ss, langs, warnings) {
  var rows = readTable_(ss.getSheetByName(EMAIL_TAB_COORDINATORS))
    .filter(function (r) {
      if (String(r.name || "").trim() === "") return false;
      if (!isEnabled_(r.enabled)) return false;
      if (r.include_in_email !== undefined && !isEnabled_(r.include_in_email)) return false;
      return true;
    })
    .sort(function (a, b) { return (Number(a.order) || 0) - (Number(b.order) || 0); });

  var order = [];
  var groups = {};
  rows.forEach(function (r) {
    var key = String(r.role_en || r.role_zh).trim();
    if (!groups[key]) { groups[key] = { en: r.role_en, zh: r.role_zh, people: [] }; order.push(key); }
    groups[key].people.push(r);
  });

  var useEn = langs.indexOf("en") !== -1;
  var html = "";
  order.forEach(function (key) {
    var g = groups[key];
    var label = [];
    if (useEn && g.en) label.push(escapeHtml_(g.en));
    if (langs.indexOf("zh") !== -1 && g.zh) label.push(escapeHtml_(g.zh));

    var people = g.people.map(function (p) {
      var name = String(p.name).trim();
      var mail = String(p.email || "").trim();
      if (mail && /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mail)) {
        return '<a href="mailto:' + escapeHtml_(mail) + '">' + escapeHtml_(name) + "</a>";
      }
      warnings.push("Coordinator '" + name + "' (" + key + ") has no valid email on the " + EMAIL_TAB_COORDINATORS + " tab.");
      return escapeHtml_(name);
    });
    html += '<p style="margin:0 0 4px;">' + label.join(" ") + ": " + people.join(useEn ? " or " : " 或 ") + "</p>";
  });

  if (!html) warnings.push("The " + EMAIL_TAB_COORDINATORS + " tab has no coordinators listed.");
  return html;
}

/** Replaces {token}. Unknown tokens throw so a typo never reaches the congregation. */
function fillTokens_(text, tokens, where) {
  return String(text).replace(/\{([a-z_0-9]+)\}/gi, function (m, key) {
    var k = key.toLowerCase();
    if (!Object.prototype.hasOwnProperty.call(tokens, k)) {
      throw new Error("Unknown token {" + key + "} in " + where + ". Valid tokens: " +
        Object.keys(tokens).sort().map(function (t) { return "{" + t + "}"; }).join(", "));
    }
    return tokens[k];
  });
}

/** Escapes HTML, converts [label](https://… | mailto:…) to links and newlines to <br>. */
function renderRichText_(text) {
  var escaped = escapeHtml_(text);
  escaped = escaped.replace(/\[([^\]]+)\]\(((?:https?:\/\/|mailto:)[^)\s]+)\)/g, function (m, label, url) {
    return '<a href="' + url + '">' + label + "</a>";
  });
  return escaped.replace(/\r?\n/g, "<br>");
}

function escapeHtml_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function htmlToText_(html) {
  return html.replace(/<\/(p|li|div)>/gi, "\n").replace(/<br\s*\/?>/gi, "\n")
    .replace(/<a href="([^"]+)">([^<]*)<\/a>/gi, "$2 ($1)")
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
    .replace(/\n{3,}/g, "\n\n").trim();
}

// -----------------------------------------------------------------------------
// CONFIG TABS (auto-created, never overwritten)
// -----------------------------------------------------------------------------

function readEmailSettings_(ss) {
  var out = {};
  readTable_(ss.getSheetByName(EMAIL_TAB_SETTINGS)).forEach(function (r) {
    var k = String(r.key || "").trim().toLowerCase();
    if (k) out[k] = (r.value === undefined || r.value === null) ? "" : String(r.value).trim();
  });
  return out;
}

/** Sheet -> array of objects keyed by lower-cased header. Skips fully blank rows. */
function readTable_(sheet) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var obj = {}, any = false;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      obj[headers[c]] = values[i][c];
      if (String(values[i][c]).trim() !== "") any = true;
    }
    if (any) rows.push(obj);
  }
  return rows;
}

/** Blank counts as enabled; N / NO / FALSE disables. */
function isEnabled_(v) {
  if (v === false) return false;
  var s = String(v).trim().toUpperCase();
  return !(s === "N" || s === "NO" || s === "FALSE");
}

/** Creates any missing config tab with Garden Grove defaults. Returns names created. */
function ensureEmailConfigTabs_(ss) {
  var created = [];
  var active = ss.getActiveSheet();
  var defs = getEmailConfigDefaults_();

  Object.keys(defs).forEach(function (name) {
    if (ss.getSheetByName(name)) return;
    var def = defs[name];
    var sh = ss.insertSheet(name);
    var data = [def.headers].concat(def.rows);
    sh.getRange(1, 1, data.length, def.headers.length).setValues(data);
    sh.getRange(1, 1, 1, def.headers.length).setFontWeight("bold").setBackground(typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.ACCENT_SUCCESS : "#d9ead3");
    sh.setFrozenRows(1);
    def.widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.getRange(2, 1, Math.max(def.rows.length, 1), def.headers.length)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top");
    if (def.note) sh.getRange(1, def.headers.length + 2).setValue(def.note).setFontStyle("italic");
    sh.hideSheet();
    created.push(name);
  });

  if (created.length) { try { ss.setActiveSheet(active); } catch (e) {} }
  return created;
}

/**
 * Seed content = the Garden Grove bilingual email. Coordinator EMAILS are left
 * blank on purpose so personal addresses are not stored in source code; fill
 * them in on the Coordinators tab.
 */
function getEmailConfigDefaults_() {
  var defs = {};

  defs[EMAIL_TAB_SETTINGS] = {
    headers: ["key", "value", "notes"],
    widths: [160, 360, 420],
    note: "Every key here is also usable as a {token} in the other tabs.",
    rows: [
      ["church_short",    typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.CHURCH_NAME : "Garden Grove", "Short church name, e.g. used in the schedule link text."],
      ["sender_name",     "True Jesus Church in " + (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.CHURCH_NAME : "Garden Grove") + " 真耶穌教會園林教會", "Display name shown as the email sender."],
      ["council_name_en", (typeof APP_CONFIG !== 'undefined' ? APP_CONFIG.CHURCH_NAME : "Garden Grove") + " Church Council", "English sign-off."],
      ["council_name_zh", "園林教會職務會", "Chinese sign-off."],
      ["reply_email",     "gg.info@tjc.org", "Where schedule-switch notices go."],
      ["languages",       "en,zh", "en, zh, or en,zh (English first)."],
      ["subject",         "{month_year} {church_short} Worship Schedule 聚會安排表", "Email subject."],
      ["pdf_filename",    "{month_year} {church_short} Worship Schedule", "PDF file name (without .pdf)."]
    ]
  };

  defs[EMAIL_TAB_TEMPLATE] = {
    headers: ["order", "block_id", "enabled", "en", "zh"],
    widths: [55, 120, 70, 480, 480],
    note: "Special block_ids: 'links' (renders EmailLinks) and 'coordinators' (en/zh = heading, then Coordinators list). Tokens: {month_year} {month} {month_num} {year} {schedule_url} + any EmailSettings key.",
    rows: [
      [1, "greeting", "Y", "Dear Brothers and Sisters,", "親愛的弟兄姊妹們平安,"],
      [2, "intro", "Y",
        "The [{month_year} {church_short} worship schedule]({schedule_url}) is now available. Please review any assignments you may have. If you have any schedule conflicts, please switch with other team members and then inform the team coordinator(s) noted below, along with [{reply_email}](mailto:{reply_email}) of the switch.",
        "園林教會UPDATE:{year}年{month_num}月份的聚會安排表已經推出。請查閱您是否被安排任何的事工。如果您有任何安排上的衝突，請您自行與其他的組員調換，並電郵至 [{reply_email}](mailto:{reply_email}) 通知組負責。"],
      [3, "notes_header", "Y", "Note the following:", "請注意以下事項:"],
      [4, "links", "Y", "", ""],
      [5, "coordinators", "Y", "Team Coordinators:", "組負責:"],
      [6, "closing", "Y", "May the Lord remember your labor and love!", "願神祝福您的辛勞與愛心！"],
      [7, "signoff", "Y", "In Christ,\n{council_name_en}", "在基督裏,\n{council_name_zh}"]
    ]
  };

  defs[EMAIL_TAB_LINKS] = {
    headers: ["order", "enabled", "en", "zh"],
    widths: [55, 70, 520, 520],
    note: "Update the year-specific URLs (Master Calendar, Bible Reading) each year. Leave 'zh' blank for an English-only line.",
    rows: [
      [1, "Y",
        "Upcoming events are on the [{year} Master Calendar](https://docs.google.com/spreadsheets/d/e/2PACX-1vT_S9RFiFW0nx1maU4YhHMuSCzEFyN79xf9y9Qiig0QvSlvIb_2vr6kOF4GqYoXTIz30_GjawTjkldk/pubhtml?gid=0&single=true)",
        "{year}年行事曆中每一個將舉行的活動 [{year} Master Calendar](https://docs.google.com/spreadsheets/d/e/2PACX-1vT_S9RFiFW0nx1maU4YhHMuSCzEFyN79xf9y9Qiig0QvSlvIb_2vr6kOF4GqYoXTIz30_GjawTjkldk/pubhtml?gid=0&single=true)"],
      [2, "Y",
        "Subscribe to the [TJC Garden Grove Google Calendar](https://calendar.google.com/calendar/embed?src=tjc.org_i3qr1o6r016g587gv2gbh9eqv0%40group.calendar.google.com&ctz=America%2FLos_Angeles)",
        "訂閱真耶穌教會園林教會的谷歌行事曆。"],
      [3, "Y",
        "[{year} Bible Reading Schedule](https://drive.google.com/file/d/1FMgI4IulWOpVmq3jPwbRvhj9NVeS9lTN/view?usp=sharing) is available, and we encourage you to participate.",
        "{year}年的讀經安排表已經推出。我們鼓勵您踴躍參與。"],
      [4, "Y",
        "[Prayer Requests 代禱事項](https://docs.google.com/document/d/e/2PACX-1vRLVLL32sptiRHN6NbW3dFUztEl9hXHs7qiIrXzdVxCZFPPSdNOPQjRHzmVa27eol3_9mrkRxKNF8PL/pub)",
        ""]
    ]
  };

  defs[EMAIL_TAB_COORDINATORS] = {
    headers: ["order", "enabled", "role_en", "role_zh", "name", "email", "include_in_email", "q_columns"],
    widths: [55, 65, 180, 120, 160, 240, 110, 140],
    note: "Coordinators directory. Set 'include_in_email' to Y for email reminders (N to hide). 'q_columns' maps to Q1-Q4 master columns.",
    rows: [
      [1, "Y", "Speaker", "講道", "", "", "Y", "B"],
      [2, "Y", "Interpreter", "翻譯", "", "", "Y", "C"],
      [3, "Y", "Praise", "詩頌", "", "", "Y", "D"],
      [4, "Y", "Accompanist", "司琴", "", "", "Y", "E"],
      [5, "Y", "Usher", "接待", "", "", "Y", "F, G"],
      [6, "Y", "A/V", "視訊", "", "", "Y", "H, I"],
      [7, "Y", "Grace / Prayer", "謝飯 / 禱告", "", "", "Y", "J, K, X"],
      [8, "Y", "Religious Education (K)", "幼兒班 (K)", "", "", "N", "L, M"],
      [9, "Y", "Religious Education (E1)", "幼年班 (E1)", "", "", "N", "N, O"],
      [10, "Y", "Religious Education (E2)", "少年班 (E2)", "", "", "N", "P, Q"],
      [11, "Y", "Religious Education (J1)", "初級班 (J1)", "", "", "N", "R, S"],
      [12, "Y", "Religious Education (J2)", "中級班 (J2)", "", "", "N", "T, U"],
      [13, "Y", "Religious Education (Senior)", "高級班 (Senior)", "", "", "N", "V"],
      [14, "Y", "Cooking", "炊事", "", "", "N", "W"],
      [15, "Y", "Truth Seekers", "慕道班", "", "", "N", "Y"],
      [16, "Y", "New Believers", "初信班", "", "", "N", "Z"]
    ]
  };

  defs[EMAIL_TAB_PUBLISHED] = {
    headers: ["tab_name", "url"],
    widths: [160, 760],
    note: "One row per month tab. Paste the link from File > Share > Publish to web for THAT tab only. The script asks for it at send time if the row is missing.",
    rows: [
      ["October 2026", "https://docs.google.com/spreadsheets/d/e/2PACX-1vRbbRFpc3cX6Bco8DWq2q0faK5NzDr7MUiKtQxtimbHHv7gB8hRBMLBr9QM_6bQUWx1yIP8acodQWRp/pubhtml?gid=72197419&single=true"]
    ]
  };

  return defs;
}

// -----------------------------------------------------------------------------
// HELPERS
// -----------------------------------------------------------------------------

/** Splits on commas/semicolons/whitespace, de-dupes, drops invalid entries. */
function parsePdfRecipients_(text) {
  var seen = {};
  return String(text || "").split(/[\s,;]+/).map(function (s) { return s.trim(); })
    .filter(function (s) {
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s)) return false;
      var k = s.toLowerCase();
      if (seen[k]) return false;
      seen[k] = true;
      return true;
    });
}

// -----------------------------------------------------------------------------
// PUBLISHED LINK (one 'Publish to web' link per month tab)
// -----------------------------------------------------------------------------
// Google does not offer an API to publish a single tab to the web, so the Info
// Department still does that step manually (File > Share > Publish to web >
// choose the month's tab). The script stores that link per tab, validates it,
// and uses it in the email, so the whole workbook never has to be published.

/**
 * Returns a validated published link for this month's tab, or "" if the user
 * cancelled / the link was unusable (in which case they have been told why).
 * Looks in the PublishedLinks tab first; otherwise asks the user to paste it
 * and saves it for next time.
 */
function resolvePublishedLink_(ss, sheet, tabName) {
  var ui = SpreadsheetApp.getUi();
  ensureEmailConfigTabs_(ss);
  var gid = sheet.getSheetId();

  var saved = findPublishedLink_(ss, tabName);
  if (saved) {
    var err = validatePublishedLink_(saved, gid);
    if (!err) return saved;
    ui.alert("Saved Link Needs Attention",
      "The link saved for '" + tabName + "' can't be used:\n" + err + "\n\nYou'll be asked for a new one.",
      ui.ButtonSet.OK);
  }

  var resp = ui.prompt(
    "Published Link for " + tabName,
    "Paste the 'Publish to web' link for the '" + tabName + "' tab only.\n\n" +
    "To get it: open the '" + tabName + "' tab > File > Share > Publish to web > Link tab > " +
    "choose the '" + tabName + "' sheet (not 'Entire document') > Web page > Publish, then copy the link.",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return "";

  var url = resp.getResponseText().trim().replace(/&amp;/g, "&");
  var problem = validatePublishedLink_(url, gid);
  if (problem) {
    ui.alert("Link Not Accepted", problem, ui.ButtonSet.OK);
    return "";
  }
  savePublishedLink_(ss, tabName, url);
  return url;
}

function findPublishedLink_(ss, tabName) {
  var rows = readTable_(ss.getSheetByName(EMAIL_TAB_PUBLISHED));
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].tab_name).trim() === tabName) return String(rows[i].url || "").trim();
  }
  return "";
}

function savePublishedLink_(ss, tabName, url) {
  var sh = ss.getSheetByName(EMAIL_TAB_PUBLISHED);
  if (!sh) return;
  var values = sh.getDataRange().getDisplayValues();
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][0]).trim().toLowerCase() === String(tabName).trim().toLowerCase()) {
      sh.getRange(r + 1, 2).setValue(url);
      return;
    }
  }
  sh.appendRow([tabName, url]);
}

/** Returns "" if the link is acceptable, otherwise a human-readable reason. */
function validatePublishedLink_(url, expectedGid) {
  if (!/^https:\/\/docs\.google\.com\/spreadsheets\/d\/e\/2PACX-[\w-]+\/pubhtml\?\S+$/.test(url)) {
    return "That doesn't look like a 'Publish to web' web-page link " +
           "(expected https://docs.google.com/spreadsheets/d/e/2PACX-…/pubhtml?gid=…&single=true).";
  }
  var m = url.match(/[?&]gid=(\d+)/);
  if (!m) return "The link has no gid, so it isn't tied to a single tab.";
  if (String(m[1]) !== String(expectedGid)) {
    return "This link is for a different tab (gid " + m[1] + "), not this month's tab (gid " + expectedGid + ").";
  }
  if (!/[?&]single=true/.test(url)) {
    return "This link would show the whole workbook. Publish just this month's tab (the link should end with single=true).";
  }
  // Make sure it actually opens for the public.
  try {
    var code = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }).getResponseCode();
    if (code !== 200) {
      return "The link didn't open publicly (HTTP " + code + "). Make sure the tab is still published " +
             "and 'Published outside your organization' is allowed.";
    }
  } catch (e) {
    return "Couldn't check the link: " + e.message;
  }
  return "";
}

// -----------------------------------------------------------------------------
// COORDINATOR SYNC TO Q1-Q4
// -----------------------------------------------------------------------------

/**
 * Reads the Workers tab and builds a dictionary mapping lowercase worker names to their email addresses.
 */
function getWorkerEmailMap_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var workerSheet = ss.getSheetByName(typeof CONFIG !== 'undefined' && CONFIG.WORKERS_SHEET ? CONFIG.WORKERS_SHEET : 'Workers');
  if (!workerSheet) return {};
  
  var data = workerSheet.getDataRange().getValues();
  if (data.length < 2) return {};
  
  var nameCols = [];
  var emailCol = -1;
  
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h.indexOf('email') !== -1 && emailCol === -1) {
      emailCol = i;
    } else if (h.indexOf('name') !== -1 || h.indexOf('worker') !== -1 || h.indexOf('member') !== -1 || h === 'english' || h === 'chinese') {
      nameCols.push(i);
    }
  }
  
  if (emailCol === -1) emailCol = 5;
  if (nameCols.length === 0) nameCols = [0, 1, 4, 7, 15, 16];
  
  var emailMap = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row) continue;
    
    var email = String(row[emailCol] || "").trim();
    if (!email || email.indexOf("@") === -1) continue;
    
    nameCols.forEach(function(cIdx) {
      if (cIdx < row.length) {
        var n = String(row[cIdx] || "").trim().toLowerCase();
        if (n && n !== 'true' && n !== 'false') {
          emailMap[n] = email;
        }
      }
    });
  }
  
  return emailMap;
}

/**
 * Upgrades or populates the Coordinators tab to ensure it has the required Option 2 columns:
 * [order, enabled, role_en, role_zh, name, email, include_in_email, q_columns]
 * Preserves all existing entries and automatically pulls emails from the Workers tab!
 */
function upgradeCoordinatorsTabStructure_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureEmailConfigTabs_(ss);
  
  var sh = ss.getSheetByName(EMAIL_TAB_COORDINATORS);
  if (!sh) return;
  
  var existingRows = readTable_(sh);
  var roleSeenMap = {};
  
  existingRows.forEach(function(r) {
    var roleKey = String(r.role_en || "").trim().toLowerCase();
    if (roleKey) roleSeenMap[roleKey] = r;
  });
  
  var workerEmailMap = getWorkerEmailMap_(ss);
  var fullHeaders = ["order", "enabled", "role_en", "role_zh", "name", "email", "include_in_email", "q_columns"];
  
  var standardRoles = [
    { order: 1, enabled: "Y", role_en: "Speaker", role_zh: "講道", defaultName: "Anthony Chang", defaultEmail: "anthony.chang@tjc.org", inc: "Y", qCols: "B" },
    { order: 2, enabled: "Y", role_en: "Interpreter", role_zh: "翻譯", defaultName: "Jeremy Chang", defaultEmail: "jjchang588@gmail.com", inc: "Y", qCols: "C" },
    { order: 3, enabled: "Y", role_en: "Praise", role_zh: "詩頌", defaultName: "Ephraim Liu", defaultEmail: "ephraim.liu@tjc.org", inc: "Y", qCols: "D" },
    { order: 4, enabled: "Y", role_en: "Accompanist", role_zh: "司琴", defaultName: "Cindy Lin", defaultEmail: "cingoo2015@gmail.com", inc: "Y", qCols: "E" },
    { order: 5, enabled: "Y", role_en: "Usher", role_zh: "接待", defaultName: "Darrus Lee", defaultEmail: "darruslee@gmail.com", inc: "Y", qCols: "F, G" },
    { order: 5, enabled: "Y", role_en: "Usher", role_zh: "接待", defaultName: "Naomi Lee", defaultEmail: "naomiwlee21@gmail.com", inc: "Y", qCols: "F, G" },
    { order: 6, enabled: "Y", role_en: "A/V", role_zh: "視訊", defaultName: "Charlie Shen", defaultEmail: "mttbox@hotmail.com", inc: "Y", qCols: "H, I" },
    { order: 7, enabled: "Y", role_en: "Grace / Prayer", role_zh: "謝飯 / 禱告", defaultName: "Anthony Chang", defaultEmail: "anthony.chang@tjc.org", inc: "Y", qCols: "J, K, X" },
    { order: 8, enabled: "Y", role_en: "Religious Education (K)", role_zh: "幼兒班 (K)", defaultName: "Philip Chen", defaultEmail: "", inc: "N", qCols: "L, M" },
    { order: 9, enabled: "Y", role_en: "Religious Education (E1)", role_zh: "幼年班 (E1)", defaultName: "Matt Shih", defaultEmail: "", inc: "N", qCols: "N, O" },
    { order: 10, enabled: "Y", role_en: "Religious Education (E2)", role_zh: "少年班 (E2)", defaultName: "Justina Lai", defaultEmail: "", inc: "N", qCols: "P, Q" },
    { order: 11, enabled: "Y", role_en: "Religious Education (J1)", role_zh: "初級班 (J1)", defaultName: "Elaine Pang", defaultEmail: "", inc: "N", qCols: "R, S" },
    { order: 12, enabled: "Y", role_en: "Religious Education (J2)", role_zh: "中級班 (J2)", defaultName: "Matai Shih", defaultEmail: "", inc: "N", qCols: "T, U" },
    { order: 13, enabled: "Y", role_en: "Religious Education (Senior)", role_zh: "高級班 (Senior)", defaultName: "Sam Chang", defaultEmail: "", inc: "N", qCols: "V" },
    { order: 14, enabled: "Y", role_en: "Cooking", role_zh: "炊事", defaultName: "Jon Chen", defaultEmail: "", inc: "N", qCols: "W" },
    { order: 15, enabled: "Y", role_en: "Truth Seekers", role_zh: "慕道班", defaultName: "Joanne Lou", defaultEmail: "", inc: "N", qCols: "Y" },
    { order: 16, enabled: "Y", role_en: "New Believers", role_zh: "初信班", defaultName: "Anthony Chang", defaultEmail: "", inc: "N", qCols: "Z" },
    { order: 17, enabled: "Y", role_en: "Cleaning", role_zh: "清潔", defaultName: "Lawrence Lee", defaultEmail: "lawrence.lee@tjc.org", inc: "N", qCols: "" }
  ];

  var newTableData = [fullHeaders];

  standardRoles.forEach(function(st) {
    var key = st.role_en.toLowerCase();
    var match = roleSeenMap[key];
    
    // Synonym matching (e.g. Pianist -> Accompanist, Hymn Leader -> Praise)
    if (!match && key === "praise") match = roleSeenMap["hymn leader"];
    if (!match && key === "accompanist") match = roleSeenMap["pianist"];
    
    var name = (match && String(match.name || "").trim()) ? String(match.name).trim() : st.defaultName;
    
    // Look up email: First check match, then Workers sheet lookup, then default
    var email = (match && String(match.email || "").trim()) ? String(match.email).trim() : "";
    if (name && workerEmailMap[name.toLowerCase()]) {
      email = workerEmailMap[name.toLowerCase()];
    }
    if (!email) email = st.defaultEmail;

    var enabled = (match && match.enabled !== undefined) ? (isEnabled_(match.enabled) ? "Y" : "N") : st.enabled;
    var inc = (match && match.include_in_email !== undefined) ? String(match.include_in_email).toUpperCase() : st.inc;
    var qCols = (match && match.q_columns !== undefined && String(match.q_columns).trim()) ? String(match.q_columns).trim() : st.qCols;

    newTableData.push([
      st.order,
      enabled,
      st.role_en,
      st.role_zh,
      name,
      email,
      inc,
      qCols
    ]);
  });

  // Retain any custom non-standard rows user added
  existingRows.forEach(function(r) {
    var key = String(r.role_en || "").trim().toLowerCase();
    var isStd = standardRoles.some(function(st) { 
      return st.role_en.toLowerCase() === key || (key === "pianist" && st.role_en === "Accompanist") || (key === "hymn leader" && st.role_en === "Praise"); 
    });
    if (!isStd && key) {
      var name = String(r.name || "").trim();
      var email = String(r.email || "").trim();
      if (name && workerEmailMap[name.toLowerCase()]) {
        email = workerEmailMap[name.toLowerCase()];
      }
      var enabled = isEnabled_(r.enabled) ? "Y" : "N";
      var inc = (r.include_in_email !== undefined) ? String(r.include_in_email).toUpperCase() : "Y";
      var qCols = String(r.q_columns || "").trim();
      var orderNum = Number(r.order) || 99;
      
      newTableData.push([
        orderNum,
        enabled,
        String(r.role_en || "").trim(),
        String(r.role_zh || "").trim(),
        name,
        email,
        inc,
        qCols
      ]);
    }
  });

  sh.clear();
  sh.getRange(1, 1, newTableData.length, fullHeaders.length).setValues(newTableData);
  sh.getRange(1, 1, 1, fullHeaders.length).setFontWeight("bold").setBackground(typeof APP_CONFIG !== 'undefined' && APP_CONFIG.THEME ? APP_CONFIG.THEME.ACCENT_SUCCESS : "#d9ead3");
  sh.setFrozenRows(1);
  
  var widths = [55, 65, 180, 120, 160, 240, 110, 140];
  widths.forEach(function(w, i) { sh.setColumnWidth(i + 1, w); });
  
  sh.getRange(1, fullHeaders.length + 2).setValue("Coordinators directory. Set 'include_in_email' to Y for email reminders (N to hide). 'q_columns' maps to Q1-Q4 master columns.").setFontStyle("italic");
}

/**
 * Synchronizes Coordinator names from the master Coordinators config tab into Row 3 (COORDINATOR:) across Q1-Q4.
 */
function syncCoordinatorsToQuarterSheets(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  upgradeCoordinatorsTabStructure_(ss);
  var coordSheet = ss.getSheetByName(EMAIL_TAB_COORDINATORS);
  if (!coordSheet) return { success: false, message: "Coordinators tab not found." };
  
  var rows = readTable_(coordSheet);
  if (rows.length === 0) return { success: false, message: "No data in Coordinators tab." };
  
  var colToNameMap = {};
  
  rows.forEach(function(r) {
    if (!isEnabled_(r.enabled)) return;
    var name = String(r.name || "").trim();
    if (!name) return;
    
    var qColsStr = String(r.q_columns || "").trim();
    if (qColsStr) {
      var colTokens = qColsStr.split(/[\s,;]+/);
      colTokens.forEach(function(colLetter) {
        var cleanLetter = colLetter.trim().toUpperCase();
        if (cleanLetter.match(/^[A-Z]+$/)) {
          var colIdx = letterToColumnIndex_(cleanLetter) + 1;
          colToNameMap[colIdx] = name;
        }
      });
    }
  });
  
  var quarters = (typeof CONFIG !== 'undefined' && CONFIG.QUARTERS) ? CONFIG.QUARTERS : ['Q1', 'Q2', 'Q3', 'Q4'];
  var updatedColsCount = 0;
  
  quarters.forEach(function(qName) {
    var qSheet = ss.getSheetByName(qName);
    if (!qSheet) return;
    
    var maxCols = qSheet.getLastColumn();
    if (maxCols < 2) return;
    
    var row2Vals = qSheet.getRange(2, 1, 1, maxCols).getValues()[0];
    var row3Range = qSheet.getRange(3, 1, 1, maxCols);
    var row3Vals = row3Range.getValues()[0];
    
    var modified = false;
    
    for (var c = 2; c <= maxCols; c++) {
      var targetName = colToNameMap[c];
      
      // Fallback matching by role_en if column is not explicitly in q_columns
      if (!targetName) {
        var headerText = String(row2Vals[c - 1] || "").trim().toUpperCase();
        if (headerText) {
          rows.forEach(function(r) {
            if (!isEnabled_(r.enabled)) return;
            var roleEn = String(r.role_en || "").trim().toUpperCase();
            if (roleEn && headerText.indexOf(roleEn) !== -1 && String(r.name || "").trim()) {
              targetName = String(r.name).trim();
            }
          });
        }
      }
      
      if (targetName && String(row3Vals[c - 1] || "").trim() !== targetName) {
        row3Vals[c - 1] = targetName;
        modified = true;
        updatedColsCount++;
      }
    }
    
    if (modified) {
      row3Range.setValues([row3Vals]);
    }
  });
  
  return {
    success: true,
    updatedCount: updatedColsCount
  };
}

function letterToColumnIndex_(letter) {
  var col = 0;
  var len = letter.length;
  for (var i = 0; i < len; i++) {
    col += (letter.charCodeAt(i) - 64) * Math.pow(26, len - i - 1);
  }
  return col - 1;
}

/**
 * UI Prompt for syncing Coordinators to Q1-Q4.
 */
function promptSyncCoordinatorsToQuarters() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ss.toast('Synchronizing coordinator names from Coordinators tab to Q1-Q4 Row 3...', 'Please wait', 15);
  try {
    var res = syncCoordinatorsToQuarterSheets(ss);
    if (res.success) {
      ui.alert(
        'Sync Complete! 🎉',
        'Successfully synchronized coordinator names across Q1-Q4.\n\nUpdated slots: ' + res.updatedCount,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('Sync Notice', res.message, ui.ButtonSet.OK);
    }
  } catch(e) {
    ui.alert('Error', 'Could not sync coordinators: ' + e.message, ui.ButtonSet.OK);
  }
}
