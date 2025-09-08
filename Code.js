// Code.gs

// Style for email cards
var EMAIL_CARD_STYLE = 'max-width:600px;margin:40px auto;padding:20px;border:1px solid #ddd;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1);background:#fff;font-family:Arial,sans-serif;';

// One-time: set a random salt into Script Properties.
// Run initAuthSalt() once from the editor, then leave it alone.
function initAuthSalt() {
  var rand = Utilities.getUuid() + ':' + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('AUTH_SALT', rand);
  return 'AUTH_SALT set.';
}

function getAuthSalt_() {
  var salt = PropertiesService.getScriptProperties().getProperty('AUTH_SALT');
  if (!salt) throw new Error('AUTH_SALT not set. Run initAuthSalt() once.');
  return salt;
}

// Canonical hash for admin passwords (salted SHA-256 → base64)
function hashAdminPassword_(plain) {
  var salt = getAuthSalt_();
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + String(plain))
  );
}

// Main entry point. Determines which page to serve.
function doGet(e) {
  var page = e.parameter.page || 'index';
  try {
    if (page === 'approve' && e.parameter.ticket && e.parameter.approver) {
      var tpl = HtmlService.createTemplateFromFile('approve');
      tpl.ticket = e.parameter.ticket;
      tpl.approver = e.parameter.approver;
      tpl.action = e.parameter.action; // pass action (approve/reject) to template
      return tpl.evaluate().setTitle('Approve Ticket ' + tpl.ticket);
    }
    return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle("YNGEN GROUP IT TICKETING SYSTEM");
  } catch (error) {
    return HtmlService.createHtmlOutput("ERROR LOADING PAGE: " + error);
  }
}

// Utility to include HTML snippets if needed.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Open the spreadsheet using its ID.
function getSpreadsheet() {
  var ssId = "1mUwIXdyHfqnf36iQp4ENXSOAIAmC1-JQrfihuVqF0q8";
  return SpreadsheetApp.openById(ssId);
}

  function getEmployeeDetails() {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName("Employee Details");
    var rows  = sheet.getDataRange().getValues();
    rows.shift(); // drop header

    return rows.map(function(r) {
      return {
        id:       r[0].toString().trim(),  // col A
        name:     r[1].toString().trim(),  // col B
        lob:      r[2].toString().trim(),  // col C
        codeHash: r[5].toString().trim()   // ← now col F, your Unique Code
      };
    });
  }


// Auto-generate the next ticket number.
function getNextTicketNumber() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return "ITID000001";
  }
  var lastTicket = sheet.getRange(lastRow, 1).getValue();
  var number = parseInt(lastTicket.replace("ITID", ""), 10);
  return "ITID" + String(number + 1).padStart(6, "0");
}

/**
 * Process attachments (passed as a JSON string) and save files to the Drive folder.
 * Returns a comma-separated string of file URLs.
 */
function processAttachments(attachmentsData) {
  var links = "";
  if (attachmentsData) {
    try {
      var files = JSON.parse(attachmentsData);
      var folder = DriveApp.getFolderById("10nFOPM2PYah8fZzEqhTOEMDYHSvni9o8");
      var urls = [];
      for (var i = 0; i < files.length; i++) {
        var fileObj = files[i];
        var decoded = Utilities.base64Decode(fileObj.data);
        var blob = Utilities.newBlob(decoded, fileObj.type, fileObj.name);
        var file = folder.createFile(blob);
        urls.push(file.getUrl());
      }
      links = urls.join(", ");
    } catch (e) {
      Logger.log("Error processing attachments: " + e);
    }
  }
  return links;
}

/**
 * Create a new ticket based on form data.
 * Now also snapshots the approvers list into column 22 (index 21).
 */
function createTicket(ticketData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  var ticketNumber = getNextTicketNumber();
  var timestamp = new Date();
  var row = [];

  // 1) Build the fixed fields
  row.push(
    ticketNumber,
    timestamp,
    ticketData.employeeId,
    ticketData.employeeName,
    ticketData.lineOfBusiness,
    ticketData.officeSite,
    ticketData.email,
    ticketData.ticketCategory,
    ticketData.priorityLevel
  );

  // 2) Category-specific fields
  if (ticketData.ticketCategory === "Incident") {
    row.push(
      ticketData.detailedLocation || "",
      ticketData.issueClassification || "",
      ticketData.issueDescription || "",
      processAttachments(ticketData.attachments),
      "",
      ""
    );
  } else if (ticketData.ticketCategory === "Service Request") {
    row.push(
      "",
      "",
      "",
      processAttachments(ticketData.attachments),
      ticketData.requestType || "",
      ticketData.additionalDetails || ""
    );
  } else {
    row.push("", "", "", "", "", "");
  }

  // 3) Status, IT-In-charge, Solution/Remarks
  row.push("Open", "", "");

  // 4) Approval status & timestamp
  if (ticketData.ticketCategory === "Service Request") {
    row.push("Pending Approval", timestamp);
  } else {
    row.push("N/A", "");
  }

  // 5) Remarks
  row.push(ticketData.remarks || "");

  // ─── NEW: snapshot the approvers list for this LOB ────────────────
  var approvers = getApprovers(ticketData.lineOfBusiness);
  row.push(JSON.stringify(approvers));

  // 6) Append and log
  sheet.appendRow(row);
  logAction(ticketNumber, "Ticket Created", ticketData.employeeId,
            "Ticket created.", ticketData.remarks || "", "");

  // Build detail list
  var labels = [
    'Ticket Number','Timestamp','Employee ID','Employee Name','Line Of Business',
    'Office Site','Email Address','Ticket Category','Priority Level',
    'Detailed Location','Issue Classification','Issue Description','Attachment',
    'Request Type','Additional Details','Status','IT-Incharge','Solution/Remarks',
    'Approval Status','Approval Timestamp','Remarks'
  ];
  var details = labels.map(function(l, i) { return [l, row[i]]; });

  // Notify IT
  var itHtml = '<div style="' + EMAIL_CARD_STYLE + '"><h2 style="text-align:center;">New Ticket Submitted</h2><table style="width:100%;border-collapse:collapse;">';
  details.forEach(function(f) {
    if (f[1]) {
      itHtml += '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">'+f[0]+'</th>';
      if (f[0] === 'Attachment') {
        itHtml += '<td style="padding:8px;border-bottom:1px solid #eee;">';
        f[1].split(',').forEach(function(url) {
          url = url.trim();
          if (url) {
            itHtml += '<a href="'+url+'" target="_blank" style="text-decoration:none;font-size:1.2em;margin-right:8px;">📎</a>';
          }
        });
        itHtml += '</td>';
      } else if (f[0] === 'Additional Details') {
        itHtml += '<td style="padding:8px;border-bottom:1px solid #eee;">';
        try {
          var obj = JSON.parse(f[1]);
          for (var key in obj) {
            if (obj.hasOwnProperty(key) && obj[key]) {
              var label = key.replace(/([A-Z])/g, ' $1');
              label = label.charAt(0).toUpperCase() + label.slice(1);
              itHtml += '<strong>' + label + ':</strong> ' + obj[key] + '<br>';
            }
          }
        } catch (e) {
          itHtml += f[1];
        }
        itHtml += '</td>';
      } else {
        itHtml += '<td style="padding:8px;border-bottom:1px solid #eee;">'+f[1]+'</td>';
      }
      itHtml += '</tr>';
    }
  });
  itHtml += '</table></div>';
  MailApp.sendEmail({
    to: 'itsupport@ydc.com.ph',
    subject: '[New Ticket] ' + ticketNumber,
    htmlBody: itHtml
  });

  // Notify Requestor
  var searchUrl = ScriptApp.getService().getUrl() + '?page=searchStatus';
  var userHtml = '<div style="' + EMAIL_CARD_STYLE + '"><h2 style="text-align:center;">Ticket Submitted Successfully</h2><table style="width:100%;border-collapse:collapse;">';
  details.forEach(function(f) {
    if (f[1]) {
      userHtml += '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">'+f[0]+'</th>';
      if (f[0] === 'Attachment') {
        userHtml += '<td style="padding:8px;border-bottom:1px solid #eee;">';
        f[1].split(',').forEach(function(url) {
          url = url.trim();
          if (url) {
            userHtml += '<a href="'+url+'" target="_blank" style="text-decoration:none;font-size:1.2em;margin-right:8px;">📎</a>';
          }
        });
        userHtml += '</td>';
      } else if (f[0] === 'Additional Details') {
        userHtml += '<td style="padding:8px;border-bottom:1px solid #eee;">';
        try {
          var obj2 = JSON.parse(f[1]);
          for (var k in obj2) {
            if (obj2.hasOwnProperty(k) && obj2[k]) {
              var lbl = k.replace(/([A-Z])/g, ' $1');
              lbl = lbl.charAt(0).toUpperCase() + lbl.slice(1);
              userHtml += '<strong>' + lbl + ':</strong> ' + obj2[k] + '<br>';
            }
          }
        } catch (e) {
          userHtml += f[1];
        }
        userHtml += '</td>';
      } else {
        userHtml += '<td style="padding:8px;border-bottom:1px solid #eee;">'+f[1]+'</td>';
      }
      userHtml += '</tr>';
    }
  });
  userHtml += '</table><div style="text-align:center;margin-top:20px;">' +
    '<a href="' + searchUrl + '" style="padding:10px 20px;background:#007bff;color:#fff;text-decoration:none;border-radius:4px;">View Your Ticket</a>' +
    '</div></div>';
  MailApp.sendEmail({
    to: ticketData.email,
    subject: 'Your Ticket ' + ticketNumber + ' Has Been Submitted',
    htmlBody: userHtml
  });

   if (ticketData.ticketCategory === "Service Request") {
    sendNextApprovalEmail(ticketNumber);
    notifyAllApprovers(ticketNumber, row, ticketData.lineOfBusiness);
  }

  return ticketNumber;
}

// ─────────────────────────────────────────────────────────────────────────────
// Dynamic Questions Schema — read from "Dynamic Questions" sheet
// Columns: Category | RequestType | Subtype | Key | Label | Type | Required | Options | Placeholder | Help | Order
// ─────────────────────────────────────────────────────────────────────────────

function getDQMeta(category) {
  // Returns all request types and their subtypes for a given category
  category = String(category || '').trim();
  var rows = _dqLoad_();
  var map = {}; // { RequestType: Set(subtypes) }
  rows.forEach(function(r) {
    if (category && r.Category !== category) return;
    if (!map[r.RequestType]) map[r.RequestType] = new Set();
    if (r.Subtype) map[r.RequestType].add(r.Subtype);
  });

  var requestTypes = Object.keys(map).sort().map(function(rt) {
    return {
      name: rt,
      subtypes: Array.from(map[rt]).sort() // may be empty
    };
  });
  return { requestTypes: requestTypes };
}

function getDQFields(category, requestType, subtype) {
  // Returns field objects to render for the selected combo
  var rows = _dqLoad_().filter(function(r) {
    if (category && r.Category !== category) return false;
    if (requestType && r.RequestType !== requestType) return false;
    // subtype match: if schema row has Subtype, it must equal the chosen subtype.
    // If schema Subtype is blank, show it for all subtypes of that RequestType.
    if (r.Subtype && subtype && r.Subtype !== subtype) return false;
    if (r.Subtype && !subtype) return false; // row expects a subtype, user hasn't chosen yet
    return true;
  });

  // sort by Order (numeric; missing -> 9999)
  rows.sort(function(a, b) {
    var oa = isNaN(a.Order) ? 9999 : Number(a.Order);
    var ob = isNaN(b.Order) ? 9999 : Number(b.Order);
    return oa - ob;
  });

  // expand Options (LIST:/SHEET:)
  rows.forEach(function(r) {
    r.Options = _dqParseOptions_(r.Options);
    r.Required = _dqToBool_(r.Required);
  });

  return rows;
}

function _dqLoad_() {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName('Dynamic Questions');
  if (!sh) return [];
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var h = values[0];
  var COLS = {
    Category:     h.indexOf('Category'),
    RequestType:  h.indexOf('RequestType'),
    Subtype:      h.indexOf('Subtype'),
    Key:          h.indexOf('Key'),
    Label:        h.indexOf('Label'),
    Type:         h.indexOf('Type'),
    Required:     h.indexOf('Required'),
    Options:      h.indexOf('Options'),
    Placeholder:  h.indexOf('Placeholder'),
    Help:         h.indexOf('Help'),
    Order:        h.indexOf('Order')
  };
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[COLS.Category] || !r[COLS.RequestType] || !r[COLS.Key]) continue;
    out.push({
      Category:    String(r[COLS.Category]).trim(),
      RequestType: String(r[COLS.RequestType]).trim(),
      Subtype:     String(r[COLS.Subtype] || '').trim(),
      Key:         String(r[COLS.Key]).trim(),
      Label:       String(r[COLS.Label] || '').trim(),
      Type:        String(r[COLS.Type] || 'text').trim().toLowerCase(),
      Required:    r[COLS.Required],
      Options:     String(r[COLS.Options] || '').trim(),
      Placeholder: String(r[COLS.Placeholder] || '').trim(),
      Help:        String(r[COLS.Help] || '').trim(),
      Order:       r[COLS.Order]
    });
  }
  return out;
}

function _dqParseOptions_(optStr) {
  if (!optStr) return [];
  if (/^LIST:/i.test(optStr)) {
    // LIST:Opt1|Opt2|Opt3
    return optStr.replace(/^LIST:/i, '').split('|').map(function(s){ return s.trim(); }).filter(Boolean);
  }
  if (/^SHEET:/i.test(optStr)) {
    // SHEET:SheetName!A2:A
    var spec = optStr.replace(/^SHEET:/i, '').trim();
    var m = /^([^!]+)!(.+)$/.exec(spec);
    if (!m) return [];
    var sh = getSpreadsheet().getSheetByName(m[1].trim());
    if (!sh) return [];
    var rng = sh.getRange(m[2].trim());
    var vals = rng.getValues().flat().map(function(v){ return String(v||'').trim(); }).filter(Boolean);
    // de-dupe while preserving order
    var seen = {};
    return vals.filter(function(v){ if (seen[v]) return false; seen[v]=1; return true; });
  }
  // raw comma/pipe fallback
  if (optStr.indexOf('|') >= 0) {
    return optStr.split('|').map(function(s){ return s.trim(); }).filter(Boolean);
  }
  return optStr.split(',').map(function(s){ return s.trim(); }).filter(Boolean);
}

function _dqToBool_(v) {
  if (typeof v === 'boolean') return v;
  var s = String(v || '').toLowerCase().trim();
  return (s === 'true' || s === 'y' || s === 'yes' || s === '1');
}


/**
 * Append a log entry in the "Logs" sheet.
 */
function logAction(ticketNumber, action, performedBy, details, remarks, editReason, validationResult) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Logs");
  // ensure 8‐column header, adding “Validation Result” if missing
  if (!sheet) {
    sheet = ss.insertSheet("Logs");
    sheet.appendRow([
      "Timestamp",
      "Request Number",
      "Action",
      "Performed By",
      "Details",
      "Remarks",
      "Edit Reason",
      "Validation Result"
    ]);
  } else {
    var header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    if (header.length < 8 || header[7] !== "Validation Result") {
      sheet.getRange(1,8).setValue("Validation Result");
    }
  }
  // now append 8 values
  sheet.appendRow([
    new Date(),
    ticketNumber,
    action,
    performedBy,
    details         || "",
    remarks         || "",
    editReason      || "",
    validationResult|| ""
  ]);
}


// —————————————————————————————————————————————————
// Approval workflow helpers
// —————————————————————————————————————————————————

/**
 * Returns an array of approver emails for a given line of business.
 */
function getApprovers(lineOfBusiness) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Approvers");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === lineOfBusiness) {
      return data[i].slice(1).filter(function(email) {
        return email;
      });
    }
  }
  return [];
}

/**
 * Sends the next approval email in sequence, including previous approvals, and updates the approval timestamp.
 */
/**
 * Create a new ticket based on form data.
 * Now also snapshots the approvers list into column 22 (index 21).
 */
function createTicket(ticketData) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  var ticketNumber = getNextTicketNumber();
  var timestamp = new Date();
  var row = [];

  // 1) Build the fixed fields
  row.push(
    ticketNumber,
    timestamp,
    ticketData.employeeId,
    ticketData.employeeName,
    ticketData.lineOfBusiness,
    ticketData.officeSite,
    ticketData.email,
    ticketData.ticketCategory,
    ticketData.priorityLevel
  );

  // 2) Category-specific fields
  if (ticketData.ticketCategory === "Incident") {
    row.push(
      ticketData.detailedLocation || "",
      ticketData.issueClassification || "",
      ticketData.issueDescription || "",
      processAttachments(ticketData.attachments),
      "",
      ""
    );
  } else if (ticketData.ticketCategory === "Service Request") {
    row.push(
      "",
      "",
      "",
      processAttachments(ticketData.attachments),
      ticketData.requestType || "",
      ticketData.additionalDetails || ""
    );
  } else {
    row.push("", "", "", "", "", "");
  }

  // 3) Status, IT-In-charge, Solution/Remarks
  row.push("Open", "", "");

  // 4) Approval status & timestamp
  if (ticketData.ticketCategory === "Service Request") {
    row.push("Pending Approval", timestamp);
  } else {
    row.push("N/A", "");
  }

  // 5) Remarks
  row.push(ticketData.remarks || "");

  // ─── NEW: snapshot the approvers list for this LOB ────────────────
  var approvers = getApprovers(ticketData.lineOfBusiness);
  row.push(JSON.stringify(approvers));

  // 6) Append and log
  sheet.appendRow(row);
  logAction(ticketNumber, "Ticket Created", ticketData.employeeId,
            "Ticket created.", ticketData.remarks || "", "");

  // … then your existing email-notification logic here …
  // (unchanged)

  if (ticketData.ticketCategory === "Service Request") {
    sendNextApprovalEmail(ticketNumber);
    notifyAllApprovers(ticketNumber, row, ticketData.lineOfBusiness);
  }

  return ticketNumber;
}


/**
 * Sends the next approval email in sequence, reading the
 * *stored* approvers list instead of re-looking up the sheet.
 */
/**
 * Sends the next approval email in sequence, reading the
 * *stored* approvers list instead of re-looking up the sheet.
 */
function sendNextApprovalEmail(ticketNumber) {
  var ss      = getSpreadsheet();
  var raw     = ss.getSheetByName("Raw");
  var allRows = raw.getDataRange().getValues();
  var ticketRow, rawRowIndex;
  for (var i = 1; i < allRows.length; i++) {
    if (allRows[i][0] === ticketNumber) {
      ticketRow   = allRows[i];
      rawRowIndex = i + 1;
      break;
    }
  }
  if (!ticketRow) return;

  // 1) Load the snapshot of approvers from column V (index 21)
  var approvers = JSON.parse(ticketRow[21] || "[]");

  // 2) Load past approvals
  var appSh = ss.getSheetByName("Approvals") || ss.insertSheet("Approvals");
  if (appSh.getLastRow() === 0) {
    appSh.appendRow(["Request Number","Approver Email","Approval Timestamp","Action","Comment"]);
  }
  var past = appSh.getDataRange().getValues().slice(1)
               .filter(r => r[0] === ticketNumber);

  // 3) If we’re done, finalize
  if (past.filter(r => r[3]==='approve').length >= approvers.length) {
    raw.getRange(rawRowIndex, 19).setValue("Approved");
    raw.getRange(rawRowIndex, 20).setValue(new Date());
    sendCompletionEmail(ticketNumber,"Approved");
    return;
  }

  // 4) Otherwise, email the next approver
  var next = approvers[past.length];
  raw.getRange(rawRowIndex, 19).setValue("Pending Approval – " + next);
  raw.getRange(rawRowIndex, 20).setValue(new Date());

  // 1) Build the card header & details table
  var html = '<div style="' + EMAIL_CARD_STYLE + '">'
           +  '<h2 style="text-align:center;margin-bottom:20px;">'
           +    'Ticket ' + ticketNumber + ' – Pending Approval'
           +  '</h2>'
           +  '<h3 style="margin-bottom:10px;font-weight:600;">Ticket Details</h3>'
           +  '<table style="width:100%;border-collapse:collapse;">'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Request Number</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[0] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Timestamp</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[1] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Employee ID</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[2] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Employee Name</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[3] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Line Of Business</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[4] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Office Site</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[5] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Email Address</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[6] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Attachment</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">'
           +        ticketRow[12].split(',').map(function(url){
                        url = url.trim();
                        return '<a href="'+url+'" target="_blank" style="text-decoration:none;font-size:1.2em;margin-right:8px;">📎</a>';
                      }).join('')
           +      '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Request Type</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketRow[13] + '</td></tr>'
           +    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Additional Details</th>'
           +      '<td style="padding:8px;border-bottom:1px solid #eee;">';
  try {
    var obj = JSON.parse(ticketRow[14]);
    for (var key in obj) {
      if (obj[key]) {
        var label = key.replace(/([A-Z])/g,' $1')
                       .replace(/^./,c=>c.toUpperCase());
        html += '<strong>' + label + ':</strong> ' + obj[key] + '<br>';
      }
    }
  } catch(e) {
    html += ticketRow[14];
  }
  // ** insert Remarks row just before closing the table **
  html +=   '</td></tr>'
        +   '<tr>'
        +     '<th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Remarks</th>'
        +     '<td style="padding:8px;border-bottom:1px solid #eee;">'
        +       (ticketRow[20] || '')
        +     '</td>'
        +   '</tr>'
        + '</table>';

   // 6) Approval History
  if (past.length) {
    html += '<h3 style="margin-top:20px;margin-bottom:10px;font-weight:600;">Approval History</h3>'
         +  '<table style="width:100%;border-collapse:collapse;">'
         +    '<tr>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Approver</th>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Action</th>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Timestamp</th>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Comment</th>'
         +    '</tr>';
    past.forEach(function(r){
      html += '<tr>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + r[1] + '</td>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + r[3] + '</td>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + r[2] + '</td>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + (r[4]||'') + '</td>'
           +  '</tr>';
    });
    html += '</table>';
  }

  // 7) Buttons
  var base = ScriptApp.getService().getUrl()
           + '?page=approve&ticket='  + encodeURIComponent(ticketNumber)
           + '&approver=' + encodeURIComponent(next);
  html += '<div style="text-align:center;margin-top:30px;">'
       +   '<a href="' + base + '&action=approve" '
       +     'style="margin-right:10px;padding:10px 20px;background:#28a745;color:#fff;text-decoration:none;border-radius:4px;">Approve</a>'
       +   '<a href="' + base + '&action=reject" '
       +     'style="padding:10px 20px;background:#dc3545;color:#fff;text-decoration:none;border-radius:4px;">Decline</a>'
       +  '</div></div>';

  MailApp.sendEmail({
    to:       next,
    subject:  "[Service Request Approval] Ticket " + ticketNumber,
    htmlBody: html
  });
}



/**
 * Called by approve.html; records approval/rejection then continues sequence.
 */
function recordApproval(ticketNumber, approverEmail, action, comment) {
  // ── 0) Authorization check ─────────────────────────────────────────
  var currentUser = Session.getActiveUser().getEmail(); 
  if (!currentUser || currentUser.toLowerCase().trim() !== approverEmail.toLowerCase().trim()) {
    throw new Error("You are not authorized to respond to this approval.");
  }

  // ── 1) require a comment on both Approve and Decline ──────────────
  comment = (comment || "").trim();
  if (!comment) {
    throw new Error(
      "Please provide a comment when " +
      (action === 'approve' ? "approving" : "declining") +
      " this request."
    );
  }

  // ── 2) append to the Approvals sheet ──────────────────────────────
  var ss    = getSpreadsheet();
  var appSh = ss.getSheetByName("Approvals") || ss.insertSheet("Approvals");
  if (appSh.getLastRow() === 0) {
    appSh.appendRow([
      "Request Number",
      "Approver Email",
      "Approval Timestamp",
      "Action",
      "Comment"
    ]);
  }
  // prevent double‐submission
  var already = appSh.getDataRange().getValues().slice(1).some(function(r) {
    return r[0] === ticketNumber
        && r[1].toString().toLowerCase().trim() === approverEmail.toString().toLowerCase().trim();
  });
  if (already) {
    throw new Error("You have already responded to this request.");
  }
  var now = new Date();
  appSh.appendRow([ ticketNumber, approverEmail.trim(), now, action, comment ]);

  // ─── NEW: record this approval decision in Logs ─────────────
  logAction(
    ticketNumber,
    action === 'approve' ? 'Ticket Approved' : 'Ticket Declined',
    approverEmail.trim(),
    'Approval ' + action,
    comment,
    '',    // no edit‐reason
    ''     // no validationResult here
  );

  // ── 3) update the Raw sheet’s Approval Status & timestamp ─────────
  var rawSh   = ss.getSheetByName("Raw");
  var rawData = rawSh.getDataRange().getValues();
  var rowIdx  = rawData.findIndex(function(r){ return r[0] === ticketNumber; });
  if (rowIdx < 1) return;  // not found

  // col 19 = Approval Status, col 20 = Approval Timestamp
  rawSh.getRange(rowIdx+1, 19)
       .setValue(action === 'approve' ? "Approved" : "Declined");
  rawSh.getRange(rowIdx+1, 20)
       .setValue(now);

  // ── 4) decide next step ────────────────────────────────────────────
  if (action === 'reject') {
    sendCompletionEmail(ticketNumber, 'Declined');
    return;
  }

  // count how many have already APPROVED
  var approvalsDone = appSh.getDataRange().getValues().slice(1)
    .filter(function(r){ return r[0] === ticketNumber && r[3] === 'approve'; })
    .length;

  // read the snapshot of approvers from column V (22nd column)
  var snapshotJson = rawSh.getRange(rowIdx+1, 22).getValue() || "[]";
  var approvers    = JSON.parse(snapshotJson);

  if (approvalsDone < approvers.length) {
    // not done yet → ping the next approver
    sendNextApprovalEmail(ticketNumber);
  } else {
    // all done!
    sendCompletionEmail(ticketNumber, 'Approved');
  }
}




/**
 * Send completion notification to requestor and IT support.
 */
function sendCompletionEmail(ticketNumber, finalStatus) {
  var ss = getSpreadsheet();
  var raw = ss.getSheetByName("Raw");
  var data = raw.getDataRange().getValues();
  var rawRow;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === ticketNumber) {
      rawRow = data[i];
      break;
    }
  }
  if (!rawRow) return;
  var requestorEmail = rawRow[6];
  var category = rawRow[7];

  // Gather decline comment if applicable
  var comment = '';
  if (finalStatus === 'Declined') {
    var apps = ss.getSheetByName("Approvals").getDataRange().getValues().slice(1)
      .filter(function(r) { return r[0] === ticketNumber && r[3] === 'reject'; });
    if (apps.length) comment = apps[apps.length - 1][4];
  }

  // Build card-style HTML
  var html = '<div style="' + EMAIL_CARD_STYLE + '">';
  html += '<h2 style="text-align:center;margin-bottom:20px;">Ticket ' + ticketNumber + ' - ' + finalStatus + '</h2>';

  // Ticket details
  html += '<h3>Ticket Details</h3><table style="width:100%;border-collapse:collapse;">' +
    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;width:40%;">Employee Name</th><td style="padding:8px;border-bottom:1px solid #eee;">' + rawRow[3] + '</td></tr>' +
    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Line Of Business</th><td style="padding:8px;border-bottom:1px solid #eee;">' + rawRow[4] + '</td></tr>' +
    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Ticket Category</th><td style="padding:8px;border-bottom:1px solid #eee;">' + rawRow[7] + '</td></tr>' +
    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Priority Level</th><td style="padding:8px;border-bottom:1px solid #eee;">' + rawRow[8] + '</td></tr>';
  if (comment) {
    html += '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Decline Comment</th>' +
      '<td style="padding:8px;border-bottom:1px solid #eee;">' + comment + '</td></tr>';
  }
  html += '</table>';

  // Service Request final approval: history, no reopen
  if (finalStatus === 'Approved' && category === 'Service Request') {
    var approvals = ss.getSheetByName("Approvals").getDataRange().getValues().slice(1)
      .filter(function(r){ return r[0] === ticketNumber; });
    if (approvals.length) {
      html += '<h3 style="margin-top:20px;">Approval History</h3><table style="width:100%;border-collapse:collapse;">' +
        '<tr><th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Approver</th>' +
        '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Action</th>' +
        '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Timestamp</th></tr>';
      approvals.forEach(function(a) {
        html += '<tr><td style="padding:8px;border-bottom:1px solid #eee;">' + a[1] + '</td>' +
          '<td style="padding:8px;border-bottom:1px solid #eee;">' + a[3] + '</td>' +
          '<td style="padding:8px;border-bottom:1px solid #eee;">' + a[2] + '</td></tr>';
      });
      html += '</table>';
    }
  }
  // Other completions: show reopen link to search page
  else if (finalStatus === 'Approved' || finalStatus === 'Completed') {
    html += '<p style="margin-top:20px;font-size:0.9em;color:#555;">If this issue persists, you may reopen your ticket within 30 days by visiting the search page.</p>';
    var reopenUrl = ScriptApp.getService().getUrl() + '?page=searchStatus';
    html += '<div style="text-align:center;margin-top:10px;">' +
      '<a href="' + reopenUrl + '" style="padding:10px 20px;background:#007bff;color:#fff;text-decoration:none;border-radius:4px;">Reopen Ticket</a>' +
      '</div>';
  }

  html += '</div>';

  var subject = '[Ticket ' + ticketNumber + '] ' + finalStatus;
  var recipients = requestorEmail + ', itsupport@ydc.com.ph';
  MailApp.sendEmail({ to: recipients, subject: subject, htmlBody: html });
}

/**
 * Search tickets by employee ID.
 */
function searchTickets(employeeId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Raw");
    if (!sheet) {
      Logger.log("Sheet 'Raw' not found.");
      return { headers: [], data: [] };
    }
    var lastRow = sheet.getLastRow();
    var dataRange = sheet.getRange(1, 1, lastRow, 21);
    var data = dataRange.getValues();
    if (data.length < 2) {
      return { headers: [], data: [] };
    }
    var headers = data.shift().map(function(cell) {
      return cell instanceof Date ? cell.toISOString() : cell;
    });
    var searchId = employeeId.toString().toUpperCase().trim();
    var results = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i].map(function(cell) {
        return cell instanceof Date ? cell.toISOString() : cell;
      });
      var rowId = row[2].toString().toUpperCase().trim();
      if (rowId === searchId) {
        results.push(row);
      }
    }
    return { headers: headers, data: results };
  } catch (e) {
    Logger.log("Error in searchTickets: " + e);
    return { headers: [], data: [] };
  }
}

// Returns only the signed-in requester's tickets (by Raw col 7 = Email Address)
function requesterGetMyTickets() {
  var me = Session.getActiveUser().getEmail();
  if (!me) throw new Error('Must be signed in with Google Workspace.');

  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  if (!sheet) return { headers: [], data: [] };

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { headers: [], data: [] };

  var rng = sheet.getRange(1, 1, lastRow, 21).getValues();
  var headers = rng.shift().map(function(c){ return c instanceof Date ? c.toISOString() : c; });

  var mine = rng.filter(function(row){
    return String(row[6]).toLowerCase().trim() === String(me).toLowerCase().trim();
  }).map(function(row){
    return row.map(function(c){ return c instanceof Date ? c.toISOString() : (c == null ? "" : String(c)); });
  });

  return { headers: headers, data: mine };
}

/**
 * Get all tickets for admin view.
 */
function getAllTickets() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  if (!sheet) {
    Logger.log("Sheet 'Raw' not found.");
    return { headers: [], data: [] };
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No data found in 'Raw' sheet beyond the header row.");
    return { headers: [], data: [] };
  }
  var dataRange = sheet.getRange(1, 1, lastRow, 21);
  var data = dataRange.getValues();
  data = data.map(function(row) {
    return row.map(function(cell) {
      if (cell instanceof Date) {
        return cell.toLocaleString();
      } else if (cell == null) {
        return "";
      } else {
        return cell.toString();
      }
    });
  });
  var headers = data.shift();
  return { headers: headers, data: data };
}

/**
 * Updates a ticket’s status and sends the proper notifications.
 */
/**
 * Updates a ticket’s status and sends the proper notifications.
 */
function updateTicket(ticketNumber, newStatus, itIncharge, resolutionRemarks, changeDetails) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  var data = sheet.getDataRange().getValues();
  var oldStatus = "", requestorEmail = "";
  var rawRowIndex;

  // 1) find the row, grab old status + requestor email, update the row
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === ticketNumber.toString().trim()) {
      oldStatus       = data[i][15];
      requestorEmail  = data[i][6];
      rawRowIndex     = i + 1;
      sheet.getRange(rawRowIndex, 16).setValue(newStatus);        // Status
      sheet.getRange(rawRowIndex, 17).setValue(itIncharge);       // IT In-charge
      sheet.getRange(rawRowIndex, 18).setValue(resolutionRemarks); // Solution/Remarks
      break;
    }
  }

  // 2) log the update
  logAction(
    ticketNumber,
    "Ticket Updated",
    itIncharge,
    "Status update: " + oldStatus + " => " + newStatus,
    resolutionRemarks,
    changeDetails
  );

  // 3) send the usual notifications for “Completed” or “Reopened”
  if (newStatus.toLowerCase().indexOf('comp') === 0) {
    // Completed → full completion card
    sendCompletionEmail(ticketNumber, 'Completed');

  } else if (newStatus === 'Reopened') {
    // Reopen by user → log & notify both requestor & IT
    logAction(
      ticketNumber,
      "Ticket Reopen",
      itIncharge,
      new Date().toLocaleString(),
      resolutionRemarks,
      ""
    );
    // Requestor
    MailApp.sendEmail({
      to: requestorEmail,
      subject: "Ticket " + ticketNumber + " Reopened",
      htmlBody:
        '<div style="' + EMAIL_CARD_STYLE + '">' +
          '<h2 style="text-align:center;">Ticket Reopened</h2>' +
          '<p>Your ticket <strong>' + ticketNumber + '</strong> has been reopened.</p>' +
        '</div>'
    });
    // IT support
    MailApp.sendEmail({
      to: 'itsupport@ydc.com.ph',
      subject: "[Reopened] Ticket " + ticketNumber,
      htmlBody:
        '<div style="' + EMAIL_CARD_STYLE + '">' +
          '<h2 style="text-align:center;">Ticket Reopened</h2>' +
          '<p>Ticket <strong>' + ticketNumber + '</strong> has been reopened by the user.</p>' +
        '</div>'
    });
  }

  // ────────────────────────────────────────────────────────────────
  // NEW: notify the requestor when you validate their reopen request
  // ────────────────────────────────────────────────────────────────

  if (changeDetails === 'Valid Reopen') {
    MailApp.sendEmail({
      to: requestorEmail,
      subject: "Ticket " + ticketNumber + " Reopen Approved",
      htmlBody:
        '<div style="' + EMAIL_CARD_STYLE + '">' +
          '<h2 style="text-align:center;">Reopen Request Approved</h2>' +
          '<p>Your request to reopen ticket <strong>' + ticketNumber + '</strong> has been approved.</p>' +
          '<p>The ticket status is now <strong>' + newStatus + '</strong>.</p>' +
        '</div>'
    });
  }

  if (changeDetails === 'Invalid Reopen') {
    MailApp.sendEmail({
      to: requestorEmail,
      subject: "Ticket " + ticketNumber + " Reopen Declined",
      htmlBody:
        '<div style="' + EMAIL_CARD_STYLE + '">' +
          '<h2 style="text-align:center;">Reopen Request Declined</h2>' +
          '<p>We\'re sorry, but your request to reopen ticket <strong>' + ticketNumber + '</strong> has been declined.</p>' +
          '<p><strong>Reason provided:</strong> ' + resolutionRemarks + '</p>' +
        '</div>'
    });
  }

  return true;
}


/**
 * Edit ticket and notify of field changes.
 */
function editTicket(ticketNumber, newData, editReason) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  var data = sheet.getDataRange().getValues();
  var oldRow = null;
  var rowIndex;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === ticketNumber.toString().trim()) {
      oldRow = data[i];
      rowIndex = i+1;
      sheet.getRange(rowIndex, 1, 1, newData.length).setValues([newData]);
      break;
    }
  }
  if (!oldRow) return false;
  var diffs = [];
  if (oldRow[7]  !== newData[7])  diffs.push("Category: " + oldRow[7] + " => " + newData[7]);
  if (oldRow[8]  !== newData[8])  diffs.push("Priority: " + oldRow[8] + " => " + newData[8]);
  if (oldRow[10] !== newData[10]) diffs.push("Class: "    + oldRow[10] + " => " + newData[10]);
  if (oldRow[13] !== newData[13]) diffs.push("ReqType changed");
  if (oldRow[14] !== newData[14]) diffs.push("Additional details changed");
  if (oldRow[15] !== newData[15]) diffs.push("Status: "   + oldRow[15] + " => " + newData[15]);

  logAction(ticketNumber, "Ticket Edited", newData[16], "Ticket Edited: " + diffs.join(", "), newData[17], editReason);

  // Notify requestor of edits
  var requestorEmail = oldRow[6];
  var editHtml = '<div style="' + EMAIL_CARD_STYLE + '"><h2 style="text-align:center;">Ticket Edited</h2>' +
    '<p>The following fields were updated on ticket <strong>' + ticketNumber + '</strong>:</p><ul>';
  diffs.forEach(function(d) {
    editHtml += '<li>' + d + '</li>';
  });
  editHtml += '</ul></div>';
  MailApp.sendEmail({ to: requestorEmail, subject: 'Your Ticket ' + ticketNumber + ' Was Edited', htmlBody: editHtml });

  // If status changed to Completed via edit, send completion notification
  if (newData[15].toLowerCase().indexOf('comp') === 0) {
    sendCompletionEmail(ticketNumber, 'Completed');
  }

  return true;
}

function authenticateAdmin({ username, password }) {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Admin Credential");
  var rows  = sheet.getDataRange().getValues();
  var headers = rows.shift();
  var COL_USER = headers.indexOf("Username");
  var COL_PW   = headers.indexOf("Password");
  var COL_NAME = headers.indexOf("Full Name");
  var COL_ROLE = headers.indexOf("Role");

  // hash the incoming password the same way you stored it
  var hashedInput = hashItForYou(password);

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row[COL_USER] === username) {
      if (row[COL_PW] === hashedInput) {
        return {
          success: true,
          fullName: row[COL_NAME],
          role:     row[COL_ROLE]
        };
      }
      break;
    }
  }
  return { success: false };
}



function getAdminList() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Admin Credential");
  var data = sheet.getDataRange().getValues();
  data.shift();
  return data.map(function(r) {
    return r[0];
  });
}

function getLogs() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Logs");
  if (!sheet) {
    Logger.log("Sheet 'Logs' not found.");
    return { headers: [], data: [] };
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No logs found in 'Logs' sheet.");
    return { headers: [], data: [] };
  }
  var data = sheet.getRange(1, 1, lastRow, 7).getValues();
  data = data.map(function(r) {
    return r.map(function(c) {
      return c instanceof Date ? c.toLocaleString() : (c || "").toString();
    });
  });
  var headers = data.shift();
  return { headers: headers, data: data };
}

function exportLogs(itid) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Logs");
  var ssId = ss.getId();
  if (itid && itid.trim() !== "") {
    var tempSheetName = "TempExportLogs_" + new Date().getTime();
    var tempSheet = ss.insertSheet(tempSheetName);
    var logsData = sheet.getDataRange().getValues();
    var filtered = logsData.filter(function(r, i) {
      if (i === 0) return true;
      return (
        r[1] && r[1].toString().toUpperCase().trim() === itid.toUpperCase().trim()
      );
    });
    tempSheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
    var url = exportSheetAsPdf(ssId, tempSheetName);
    ss.deleteSheet(tempSheet);
    return url;
  } else {
    return exportSheetAsPdf(ssId, "Logs");
  }
}

function exportSheetAsPdf(spreadsheetId, sheetName) {
  var sheetGid = getSheetGid(spreadsheetId, sheetName);
  return (
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export?format=pdf&portrait=true&gid=" +
    sheetGid +
    "&size=A4&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false"
  );
}

function getSheetGid(spreadsheetId, sheetName) {
  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error("Sheet not found: " + sheetName);
  }
  return sheet.getSheetId();
}

/**
 * Makes a timestamped copy of the active spreadsheet into the specified Drive folder.
 */
function backupTicketingSystem() {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var file     = DriveApp.getFileById(ss.getId());
    var baseName = ss.getName();

    var backupFolderId = '128iUeH46ZC9fjsbjsMVMT7GvRq5lwm3z';
    var backupFolder   = DriveApp.getFolderById(backupFolderId);

    var tz       = Session.getScriptTimeZone();
    var now      = new Date();
    var stamp    = Utilities.formatDate(now, tz, 'yyyy-MM-dd_HHmmss');
    var copyName = baseName + '*Backup*' + stamp;

    file.makeCopy(copyName, backupFolder);
    Logger.log('✅ Backup successful: ' + copyName);

  } catch (err) {
    Logger.log('❌ Backup failed: ' + err.toString());
  }
}

/**
 * Send reminder for all pending approvals and update timestamp.
 */
function resendPendingApprovals() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    var status = data[i][18];
    if (status && status.toString().startsWith("Pending Approval -")) {
      var parts = status.split(" - ");
      var ticketNumber = data[i][0];
      var approverEmail = parts[1];
      // build and send reminder email
      var html = '<div style="' + EMAIL_CARD_STYLE + '"><h2 style="text-align:center;margin-bottom:20px;">Reminder: Pending Approval</h2>' +
        '<p>You have a pending approval for ticket <strong>' + ticketNumber + '</strong>.</p></div>';
      MailApp.sendEmail({
        to: approverEmail,
        subject: '[Reminder] Pending Approval for Ticket ' + ticketNumber,
        htmlBody: html
      });
      // update timestamp
      sheet.getRange(i + 1, 20).setValue(now);
    }
  }
}

/**
 * Dashboard helper: totals, averages, status counts (Open+Reopened → Open),
 * breakdowns, and enriched critical‐ticket details.
 */
function getDashboardMetrics(granularity, fromMonth, toMonth) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  if (!sheet) {
    return {
      totalTickets: 0,
      avgPerMonth:  0,
      avgPerYear:   0,
      statusCounts: {
        Open:0, Reopened:0, Completed:0, "In Progress":0, "On Hold":0, Canceled:0
      },
      breakdowns: {
        lineOfBusiness:{}, officeSite:{}, ticketCategory:{},
        priorityLevel:{}, issueClassification:{}, requestType:{}
      }
    };
  }

  // Read and filter rows by date
  var data = sheet.getDataRange().getValues();
  data.shift(); // header

  var start = fromMonth ? new Date(fromMonth) : null;
  var end   = toMonth   ? new Date(toMonth)   : null;
  if (end) end.setHours(23,59,59,999);

  var totalTickets = 0;
  var statusCounts = {};
  var breakdowns = {
    lineOfBusiness:     {},
    officeSite:         {},
    ticketCategory:     {},
    priorityLevel:      {},
    issueClassification:{},
    requestType:        {}
  };

  data.forEach(function(row) {
    var ts = row[1]; // Timestamp
    if (!(ts instanceof Date)) return;
    if (start && ts < start) return;
    if (end   && ts > end)   return;

    totalTickets++;
    var status = (row[15]||"").toString().trim();
    if (status) statusCounts[status] = (statusCounts[status]||0) + 1;

    // breakdown keys
    var lob = (row[4]  ||"").toString().trim();
    var site= (row[5]  ||"").toString().trim();
    var cat = (row[7]  ||"").toString().trim();
    var pri = (row[8]  ||"").toString().trim();
    var cls = (row[10] ||"").toString().trim();
    var req = (row[13] ||"").toString().trim();

    if (lob) breakdowns.lineOfBusiness[lob]     = (breakdowns.lineOfBusiness[lob]     ||0) + 1;
    if (site)breakdowns.officeSite[site]         = (breakdowns.officeSite[site]         ||0) + 1;
    if (cat) breakdowns.ticketCategory[cat]      = (breakdowns.ticketCategory[cat]      ||0) + 1;
    if (pri) breakdowns.priorityLevel[pri]       = (breakdowns.priorityLevel[pri]       ||0) + 1;
    if (cls) breakdowns.issueClassification[cls] = (breakdowns.issueClassification[cls] ||0) + 1;
    if (req) breakdowns.requestType[req]         = (breakdowns.requestType[req]         ||0) + 1;
  });

  // ensure all status keys exist
  var sc = {
    Open:           statusCounts.Open           ||0,
    Reopened:       statusCounts.Reopened       ||0,
    Completed:      statusCounts.Completed      ||0,
    "In Progress":  statusCounts["In Progress"]||0,
    "On Hold":      statusCounts["On Hold"]     ||0,
    Canceled:       statusCounts.Canceled       ||0
  };

  // compute months/years span
  var months = 1, years = 1;
  if (start && end) {
    var y1=start.getFullYear(), m1=start.getMonth();
    var y2=end.getFullYear(),   m2=end.getMonth();
    months = (y2-y1)*12 + (m2-m1) + 1;
    years  = (y2-y1) + 1;
  }
  var avgPerMonth = months>0 ? Math.round(totalTickets/months) : 0;
  var avgPerYear  = years>0  ? Math.round(totalTickets/years)  : 0;

  return {
    totalTickets: totalTickets,
    avgPerMonth:  avgPerMonth,
    avgPerYear:   avgPerYear,
    statusCounts: sc,
    breakdowns:   breakdowns
  };
}

function getPendingApprovalsCount() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  if (!sheet) return 0;
  var data  = sheet.getRange(2, 19, sheet.getLastRow()-1, 1).getValues();
  // column 19 = Approval Status
  return data.reduce(function(count, row) {
    return count + ((row[0]||'').toString().startsWith("Pending Approval") ? 1 : 0);
  }, 0);
}

/**
 * Dashboard data endpoint: returns everything the front end needs,
 * including time series for total and completed tickets.
 */
function getDashboardData(granularity, fromMonth, toMonth) {
  // existing metrics
  var d = getDashboardMetrics(granularity, fromMonth, toMonth);

  // compute yesterday’s metrics
  var yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  var yString = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // reuse metrics function but filter to only yesterday
  var yd = getDashboardMetrics(granularity, yString, yString);

  // build deltas
  var deltas = {
    Open: ((d.statusCounts.Open + d.statusCounts.Reopened) - (yd.statusCounts.Open + yd.statusCounts.Reopened)),
    Critical: (d.breakdowns.priorityLevel['Critical'] || 0) - (yd.breakdowns.priorityLevel['Critical'] || 0),
    Pending: getPendingApprovalsCount() - getPendingApprovalsCountForDate(yesterday),
  };

  // Build time series arrays (date => counts)
  var raw = getSpreadsheet().getSheetByName("Raw").getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var totalByDate = {};
  var completedByDate = {};
  // Skip header row at index 0
  for (var i = 1; i < raw.length; i++) {
    var row = raw[i];
    var ts = row[1];
    if (!(ts instanceof Date)) continue;
    var dateKey = Utilities.formatDate(ts, tz, 'yyyy-MM-dd');
    totalByDate[dateKey] = (totalByDate[dateKey] || 0) + 1;
    var status = (row[15] || '').toString().trim();
    if (status.toLowerCase().indexOf('comp') === 0) {
      completedByDate[dateKey] = (completedByDate[dateKey] || 0) + 1;
    }
  }
  // Convert to sorted arrays
  var dates = Object.keys(totalByDate).sort();
  var timeSeries = dates.map(function(date) {
    return { date: date, count: totalByDate[date] };
  });
  var timeSeriesCompleted = dates.map(function(date) {
    return { date: date, count: completedByDate[date] || 0 };
  });

  return {
    totalTickets:     d.totalTickets,
    openTickets:      d.statusCounts.Open + d.statusCounts.Reopened,
    criticalTickets:  d.breakdowns.priorityLevel['Critical'] || 0,
    pendingApprovals: getPendingApprovalsCount(),
    avgPerMonth:      d.avgPerMonth,
    avgPerYear:       d.avgPerYear,
    statusCounts:     d.statusCounts,
    breakdowns:       d.breakdowns,
    deltas:           deltas,
    timeSeries:       timeSeries,
    timeSeriesCompleted: timeSeriesCompleted
  };
}

// helper to count pending approvals on a given date
function getPendingApprovalsCountForDate(date) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Raw");
  if (!sheet) return 0;
  var rows  = sheet.getDataRange().getValues();
  var count = 0;
  var tz    = Session.getScriptTimeZone();
  rows.slice(1).forEach(function(r) {
    var ts = r[1];
    var ap = r[18];
    if (ts instanceof Date &&
        Utilities.formatDate(ts, tz, 'yyyy-MM-dd') === Utilities.formatDate(date, tz, 'yyyy-MM-dd') &&
        ap && ap.toString().startsWith("Pending Approval")) {
      count++;
    }
  });
  return count;
}

/**
 * Re-send the same Service Request approval email to one approver.
 */
function resendApprovalFor(ticketNumber, approverEmail) {
  var ss    = getSpreadsheet();
  var raw   = ss.getSheetByName("Raw");
  var all   = raw.getDataRange().getValues();
  var ticketRow, rawRowIndex;
  for (var i = 1; i < all.length; i++) {
    if (all[i][0] === ticketNumber) {
      ticketRow    = all[i];
      rawRowIndex  = i + 1;
      break;
    }
  }
  if (!ticketRow) return;

  // Load current approvals so we can show "Previous Approvals"
  var approvalsSheet = ss.getSheetByName("Approvals") || ss.insertSheet("Approvals");
  var approvalsData  = approvalsSheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] === ticketNumber);

  // Build the uniform approval email
  var html = '<div style="' + EMAIL_CARD_STYLE + '">'
           +  '<h2 style="text-align:center;margin-bottom:20px;">'
           +    'Ticket ' + ticketNumber + ' – Pending Approval'
           +  '</h2>'
           +  '<h3 style="margin-bottom:10px;font-weight:600;">Ticket Details</h3>'
           +  '<table style="width:100%;border-collapse:collapse;">';

  // List of fields in the same order
  var fieldSpecs = [
    { label: "Request Number",     value: ticketRow[0] },
    { label: "Timestamp",          value: ticketRow[1] },
    { label: "Employee ID",        value: ticketRow[2] },
    { label: "Employee Name",      value: ticketRow[3] },
    { label: "Line Of Business",   value: ticketRow[4] },
    { label: "Office Site",        value: ticketRow[5] },
    { label: "Email Address",      value: ticketRow[6] },
    { label: "Attachment",         value: ticketRow[12] },
    { label: "Request Type",       value: ticketRow[13] },
    { label: "Additional Details", value: ticketRow[14] },
    { label: "Remarks",            value: ticketRow[20] }
  ];

  fieldSpecs.forEach(function(f) {
    if (!f.value) return;
    html += '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;width:40%;">'
         +    f.label + '</th><td style="padding:8px;border-bottom:1px solid #eee;">';

    if (f.label === "Attachment") {
      f.value.toString().split(',').forEach(function(url) {
        url = url.trim();
        if (url) {
          html += '<a href="' + url + '" target="_blank" '
               +  'style="text-decoration:none;font-size:1.2em;margin-right:8px;">📎</a>';
        }
      });
    } else if (f.label === "Additional Details") {
      try {
        var obj = JSON.parse(f.value);
        for (var key in obj) {
          if (obj[key]) {
            var label = key.replace(/([A-Z])/g,' $1');
            label = label.charAt(0).toUpperCase() + label.slice(1);
            html += '<strong>' + label + ':</strong> ' + obj[key] + '<br>';
          }
        }
      } catch (e) {
        html += f.value;
      }
    } else {
      html += f.value;
    }

    html += '</td></tr>';
  });

  html += '</table>';

   // Approval History
  if (approvalsData.length) {
    html += '<h3 style="margin-top:20px;margin-bottom:10px;font-weight:600;">Approval History</h3>'
         +  '<table style="width:100%;border-collapse:collapse;">'
         +    '<tr>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Approver</th>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Action</th>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Timestamp</th>'
         +      '<th style="padding:8px;border-bottom:1px solid #eee;text-align:left;">Comment</th>'
         +    '</tr>';
    approvalsData.forEach(function(r) {
      html += '<tr>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + r[1] + '</td>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + r[3] + '</td>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + r[2] + '</td>'
           +   '<td style="padding:8px;border-bottom:1px solid #eee;">' + (r[4]||'') + '</td>'
           +  '</tr>';
    });
    html += '</table>';
  }

  // Approve / Decline buttons for this single approver
  var baseUrl   = ScriptApp.getService().getUrl();
  var approveUrl = baseUrl
    + '?page=approve'
    + '&ticket='   + encodeURIComponent(ticketNumber)
    + '&approver=' + encodeURIComponent(approverEmail)
    + '&action=approve';
  var rejectUrl  = baseUrl
    + '?page=approve'
    + '&ticket='   + encodeURIComponent(ticketNumber)
    + '&approver=' + encodeURIComponent(approverEmail)
    + '&action=reject';

  html += '<div style="text-align:center;margin-top:30px;">'
       +    '<a href="' + approveUrl + '" '
       +      'style="display:inline-block;margin:0 10px;padding:10px 20px;'
       +            'background:#28a745;color:#fff;text-decoration:none;'
       +            'border-radius:4px;">Approve</a>'
       +    '<a href="' + rejectUrl + '" '
       +      'style="display:inline-block;margin:0 10px;padding:10px 20px;'
       +            'background:#dc3545;color:#fff;text-decoration:none;'
       +            'border-radius:4px;">Decline</a>'
       +  '</div></div>';

  // Finally, send it
  MailApp.sendEmail({
    to:       approverEmail,
    subject:  "[Service Request Approval] Ticket " + ticketNumber,
    htmlBody: html
  });
}


/**
 * Helper: build the same HTML you send in sendNextApprovalEmail,
 * but for a single approver.
 */
function buildServiceRequestApprovalHtml(ticketNumber, ticketRow, approverEmail) {
  // You can basically copy-paste the HTML construction from sendNextApprovalEmail(),
  // but only output the buttons for this one approver rather than the “nextApprover”.
  // For brevity, assume you refactored that logic into this helper.
  // e.g.:
  var now = new Date();
  var html =
    '<div style="' + EMAIL_CARD_STYLE + '">' +
    '<h2 style="text-align:center;margin-bottom:20px;">Service Request Approval</h2>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr><th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">Request Number</th>' +
      '<td style="padding:8px;border-bottom:1px solid #eee;">' + ticketNumber + '</td></tr>' +
    // … include all your fields …
    '</table>' +
    '<div style="text-align:center;margin-top:30px;">' +
    '<a href="' + ScriptApp.getService().getUrl()
      + '?page=approve&ticket=' + encodeURIComponent(ticketNumber)
      + '&approver=' + encodeURIComponent(approverEmail)
      + '&action=approve" '
      + 'style="display:inline-block;margin:0 10px;padding:10px 20px;'
      + 'background:#28a745;color:#fff;text-decoration:none;border-radius:4px;">'
      + 'Approve</a>' +
    '<a href="' + ScriptApp.getService().getUrl()
      + '?page=approve&ticket=' + encodeURIComponent(ticketNumber)
      + '&approver=' + encodeURIComponent(approverEmail)
      + '&action=reject" '
      + 'style="display:inline-block;margin:0 10px;padding:10px 20px;'
      + 'background:#dc3545;color:#fff;text-decoration:none;border-radius:4px;">'
      + 'Decline</a>' +
    '</div></div>';

  return html;
}

/**
 * Returns for each Service-Request ticket:
 *  - ticket: 'ITID000123'
 *  - approvers: [email1, email2, …]  ← from the snapshot in Raw!V
 *  - approvals: [ { approver, ts, action, comment }, … ]
 */
function getApprovalOverview() {
  var ss     = getSpreadsheet();
  var raw    = ss.getSheetByName("Raw").getDataRange().getValues();
  var apps   = ss.getSheetByName("Approvals").getDataRange().getValues().slice(1);
  var overview = [];

  // skip header
  for (var i = 1; i < raw.length; i++) {
    var row = raw[i];
    if (row[7] !== "Service Request") continue;

    var ticket = row[0];
    // read *your* snapshot JSON from column V (zero-based index 21)
    var snapshotJson = row[21] || "[]";
    var approvers    = JSON.parse(snapshotJson);

    // collect any recorded approvals
    var recs = apps
      .filter(function(r){ return r[0] === ticket; })
      .map(function(r){
        return {
          approver: r[1],
          ts:       (r[2] instanceof Date ? r[2].toLocaleString() : r[2]),
          action:   r[3],
          comment:  r[4]
        };
      });

    overview.push({
      ticket:    ticket,
      approvers: approvers,
      approvals: recs
    });
  }

  return overview;
}

// RUN ONCE ONLY if your Admin Credential sheet currently stores PLAINTEXT passwords.
// It will convert each Password to a salted base64 SHA-256 hash.
function upgradeAdminPasswords() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Admin Credential");
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return 'No rows to upgrade.';

  var headers = data[0];
  var COL_PW  = headers.indexOf("Password");
  if (COL_PW < 0) throw new Error('Column "Password" not found in Admin Credential.');

  for (var i = 1; i < data.length; i++) {
    var plain = data[i][COL_PW];
    if (!plain) continue;
    var str = String(plain);
    // crude check: if it already looks base64-ish, skip
    var looksHashed = /^[A-Za-z0-9+/=]{40,}$/.test(str);
    if (looksHashed) continue;

    var hash = hashAdminPassword_(str);
    sheet.getRange(i+1, COL_PW+1).setValue(hash);
  }
  return 'Upgrade complete.';
}


/**
 * Creates a new admin user if empId is allowed.
 * Automatically sets the new user’s role to “Admin”,
 * and hashes the password the same way authenticateAdmin expects.
 */
function createAdminUser(o) {
  var ss      = getSpreadsheet();
  var allowSh = ss.getSheetByName("AllowedAdmins");
  var allowed = allowSh
    .getRange(2, 1, allowSh.getLastRow() - 1, 1)
    .getValues()
    .flat();

  // Only employees in AllowedAdmins!
  if (allowed.indexOf(o.empId) < 0) {
    return { success: false, error: "Employee ID not allowed." };
  }

  var sheet = ss.getSheetByName("Admin Credential");
  var data  = sheet.getDataRange().getValues();
  // data[0] is header row.

  // Check duplicates
  for (var i = 1; i < data.length; i++) {
    var row           = data[i];
    var existingFull  = (row[0] || "").toString().trim();
    var existingUser  = (row[1] || "").toString().trim();
    var existingEmail = (row[3] || "").toString().toLowerCase().trim();
    var existingEmpId = (row[5] || "").toString().trim();

    if (existingFull === o.fullName) {
      return { success: false, error: "Full Name already exists." };
    }
    if (existingUser === o.username) {
      return { success: false, error: "Username already exists." };
    }
    if (existingEmail === o.email) {
      return { success: false, error: "Email address already in use." };
    }
    if (existingEmpId === o.empId) {
      return { success: false, error: "Employee ID already registered." };
    }
  }

  // All clear → append new row with default role = “Admin”
  // Use the same hashItForYou() salt‐based hashing that authenticateAdmin expects
  var hash = hashItForYou(o.password);
  sheet.appendRow([
    o.fullName,     // Full Name
    o.username,     // Username
    hash,           // PasswordHash (salted, via hashItForYou)
    o.email,        // Email Address
    o.department,   // Department
    o.empId,        // Employee ID
    "Admin"         // Role (default)
  ]);

  return { success: true };
}




/**
 * Sends a temporary password to the given email if it exists in Admin Credential!D.
 * Returns { success: true } on success, or { success: false, error: "…" }.
 */
function sendPasswordResetByEmail(email) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Admin Credential");
  if (!sheet) {
    return { success: false, error: "Admin Credential sheet not found." };
  }

  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // [ FullName, Username, Password, Email Address, Department, Employee ID, Role ]
  var COL_USER = headers.indexOf("Username");
  var COL_PW   = headers.indexOf("Password");
  var COL_MAIL = headers.indexOf("Email Address");

  for (var i = 0; i < data.length; i++) {
    var rowEmail = (data[i][COL_MAIL] || '').toString().toLowerCase().trim();
    if (rowEmail && rowEmail === email.toLowerCase()) {
      var username = data[i][COL_USER];

      // generate temp password & store canonical hash
      var tempPass = Math.random().toString(36).slice(-8);
      var tempHash = hashItForYou(tempPass);
      sheet.getRange(i + 2, COL_PW + 1).setValue(tempHash);

      MailApp.sendEmail({
        to: email,
        subject: "Your temporary password",
        body:
          "Hello " + username + ",\n\n" +
          "Your password has been reset. Your temporary password is:\n\n" +
          tempPass + "\n\n" +
          "Please log in and change it immediately."
      });
      return { success: true };
    }
  }
  return { success: false, error: "Email not found." };
}

function changeAdminPassword(o) {
  // verify old password
  var auth = authenticateAdmin({ username: o.username, password: o.oldPassword });
  if (!auth.success) {
    return { success: false, error: "Old password is incorrect." };
  }
  // hash & store the new one
  return updateAdminPassword({ username: o.username, password: o.newPassword });
}


/**
 * Update an existing admin user's password in the "Admin Credential" sheet.
 * Expects o = { username: string, password: string }.
 * Returns { success: true } or { success: false, error: string }.
 */
function updateAdminPassword(o) {
  // enforce 8+ chars, letters + numbers + symbol
  var pwPattern = /^(?=.*[A-Za-z])(?=.*\d)(?=.*[^A-Za-z0-9]).{8,}$/;
  if (!pwPattern.test(o.password)) {
    return {
      success: false,
      error: "Password must be at least 8 characters and include letters, numbers & symbols."
    };
  }

  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Admin Credential");
  if (!sheet) {
    return { success: false, error: "Admin Credential sheet not found." };
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data.shift(); // [ FullName, Username, PasswordHash, Email, Department, Employee ID, Role ]
  var COL_USER = headers.indexOf("Username");
  var COL_PW   = headers.indexOf("Password");

  for (var i = 0; i < data.length; i++) {
    var storedUsername = data[i][COL_USER];
    if (storedUsername === o.username) {
      var newHash = hashItForYou(o.password);   // ← canonical
      sheet.getRange(i + 2, COL_PW + 1).setValue(newHash);
      return { success: true };
    }
  }
  return { success: false, error: "Username not found." };
}


/**
 * onOpen installable trigger will fire this.
 */
/**
 * Installable onOpen trigger: re‐lock everything, then show the login prompt.
 */
function onOpen(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sh => {
    if (sh.getName() === "Login") {
      sh.showSheet();      // ensure Login is visible
    } else {
      sh.hideSheet();      // re‐hide all data sheets
    }
  });

  // now show the password dialog
  SpreadsheetApp.getUi()
    .showModalDialog(
      HtmlService
        .createHtmlOutputFromFile("Login")
        .setWidth(300)
        .setHeight(150),
      "Please enter password"
    );
}

const PASSWORD = "TAOitservices16TAO";  // your secret

function checkPassword(pw) {
  if (pw !== PASSWORD) {
    throw new Error("Invalid password");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sh => {
    if (sh.getName() !== "Login") sh.showSheet();  // unhide data
  });
  ss.getSheetByName("Login").hideSheet();          // hide the login sheet
}

/**
 * Once‐a‐month: archive tickets >6 months and logs >3 months
 * from THIS spreadsheet into another one.
 */
function archiveOldData() {
  var SRC_ID  = "1mUwIXdyHfqnf36iQp4ENXSOAIAmC1-JQrfihuVqF0q8";  
  var DEST_ID = "106U0Vn51Chxcf8d-zYW_0SwySpBFSJV549VeZdCGT_c";

  var srcSs  = SpreadsheetApp.openById(SRC_ID);
  var rawS   = srcSs.getSheetByName("Raw");
  var logS   = srcSs.getSheetByName("Logs");
  if (!rawS || !logS) return;

  var destSs = SpreadsheetApp.openById(DEST_ID);
  var rawA   = destSs.getSheetByName("Raw Archive") || destSs.insertSheet("Raw Archive");
  var logA   = destSs.getSheetByName("Logs Archive")|| destSs.insertSheet("Logs Archive");

  // copy headers if empty
  if (rawA.getLastRow() === 0) {
    rawA.appendRow(rawS.getRange(1,1,1, rawS.getLastColumn()).getValues()[0]);
  }
  if (logA.getLastRow() === 0) {
    logA.appendRow(logS.getRange(1,1,1, logS.getLastColumn()).getValues()[0]);
  }

  var rawData = rawS.getRange(1,1, rawS.getLastRow(), rawS.getLastColumn()).getValues();
  var logData = logS.getRange(1,1, logS.getLastRow(), logS.getLastColumn()).getValues();

  var now    = new Date();
  var cutRaw = new Date(now); cutRaw.setMonth(now.getMonth() - 6);
  var cutLog = new Date(now); cutLog.setMonth(now.getMonth() - 3);

  // archive Raw bottom-up
  for (var i = rawData.length; i >= 2; i--) {
    var row   = rawData[i-1];
    var ticket= row[0];
    var ts    = row[1];
    if (ticket && ts instanceof Date && ts < cutRaw) {
      Logger.log("Archiving RAW row " + i + " (ticket "+ticket+", date "+ts+")");
      rawA.appendRow(rawS.getRange(i,1,1,rawS.getLastColumn()).getValues()[0]);
      rawS.deleteRow(i);
    }
  }

  // archive Logs bottom-up
  for (var j = logData.length; j >= 2; j--) {
    var row = logData[j-1];
    var lt  = row[0];
    if (lt instanceof Date && lt < cutLog) {
      Logger.log("Archiving LOGS row " + j + " (date "+lt+")");
      logA.appendRow(logS.getRange(j,1,1,logS.getLastColumn()).getValues()[0]);
      logS.deleteRow(j);
    }
  }
}

function logReopenValidation(ticketNumber, validationResult, performedBy) {
  logAction(
    ticketNumber,
    "Reopen Validation",
    performedBy,
    validationResult,  // ← now just “Valid Reopened” or “Invalid Reopened”
    "",                 // Remarks blank
    ""                  // Edit Reason blank
  );
}


/**
 * Returns all rows from Admin Credential as an array of objects.
 */
function getAdminUsers() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Admin Credential");
  var data  = sheet.getDataRange().getValues();
  var headers = data.shift();
  return data.map(function(r) {
    return {
      fullName:   r[0],
      username:   r[1],
      email:      r[3],
      department: r[4],
      empId:      r[5],
      role:       r[6]
    };
  });
}

function updateAdminUsers(adminUsers) {
  var ss    = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Admin Credential");
  var data  = sheet.getDataRange().getValues();
  
  // 1) pull off the header row and find column‐indexes
  var headers       = data.shift();
  var COL_FULLNAME  = headers.indexOf("Full Name");
  var COL_USERNAME  = headers.indexOf("Username");
  var COL_PASSWORD  = headers.indexOf("Password");
  var COL_EMAIL     = headers.indexOf("Email Address");
  var COL_DEPT      = headers.indexOf("Department");
  var COL_EMPID     = headers.indexOf("Employee ID");
  var COL_ROLE      = headers.indexOf("Role");
  
  // 2) build a map of existing password‐hashes by username
  var oldPasswordMap = {};
  data.forEach(function(row){
    var uname = row[COL_USERNAME];
    var pw    = row[COL_PASSWORD];
    oldPasswordMap[uname] = pw;
  });
  
  // 3) clear out old rows (keeps the header intact)
  if (data.length>0) {
    sheet.getRange(2,1,data.length, headers.length).clearContent();
  }
  
  // 4) build your new 2D array, preserving or hashing
  var output = adminUsers.map(function(u){
    var pw = u.password;
    if (!pw) {
      // user didn’t type a new password → keep the old hash
      pw = oldPasswordMap[u.username] || "";
    } else {
      // they did change it → hash it however you do
      pw = hashItForYou(u.password);
    }
    return [
      u.fullName,
      u.username,
      pw,
      u.email,
      u.department,
      u.empId,
      u.role
    ];
  });
  
  // 5) write it all back
  if (output.length) {
    sheet.getRange(2,1,output.length, output[0].length)
         .setValues(output);
  }
}

function hashItForYou(plain) {
  return hashAdminPassword_(plain);
}

// Create a session token and bind identity (8h TTL)
function issueSessionToken_(identityObj) {
  var token = Utilities.getUuid() + "." + Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  cache.put('sess:' + token, JSON.stringify(identityObj), 8 * 60 * 60); // 8 hours
  return token;
}

// Read identity from a token (or null if missing/expired)
function readSession_(token) {
  if (!token) return null;
  var raw = CacheService.getScriptCache().get('sess:' + token);
  return raw ? JSON.parse(raw) : null;
}

// Destroy a session token
function destroySession(token) {
  if (!token) return false;
  CacheService.getScriptCache().remove('sess:' + token);
  return true;
}

// Minimal guard for Step 1.2 (token only; NO role checks yet)
function requireSession_(token) {
  var ident = readSession_(token);
  if (!ident) throw new Error('Unauthorized: missing/expired session.');
  return ident; // { fullName, username, email, empId, role }
}

// (Optional) expose identity to client after login
function getSessionIdentity(token) {
  var ident = requireSession_(token);
  return ident;
}

// === SECURITY LAYER · 1.3 · Role guards (PLACE RIGHT BELOW 1.2 session helpers) ===

// True if Employee ID exists in AllowedAdmins!A
function isAllowedAdminEmpId_(empId) {
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName("AllowedAdmins");
  if (!sh) return false;
  var last = sh.getLastRow();
  if (last < 2) return false;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues().flat();
  return vals.some(function(v){ return String(v).trim() === String(empId).trim(); });
}

// Throws if no valid session OR empId is not in AllowedAdmins
function requireIT_(token) {
  var ident = readSession_(token);
  if (!ident) throw new Error('Unauthorized: missing/expired session.');
  if (!ident.empId) throw new Error('Forbidden: admin identity missing Employee ID.');
  if (!isAllowedAdminEmpId_(ident.empId)) throw new Error('Forbidden: not in AllowedAdmins.');
  return ident; // { fullName, username, email, empId, role }
}

// Requester guard (use for endpoints where a user should only see their own data)
function requireRequester_(targetEmail) {
  var me = Session.getActiveUser().getEmail();
  if (!me) throw new Error('Must be signed in with Google Workspace.');
  if (String(me).toLowerCase().trim() !== String(targetEmail).toLowerCase().trim()) {
    throw new Error('You may only access your own tickets.');
  }
  return me;
}


// === SECURITY LAYER · 1.2 · Login that returns a token ===
function loginAdmin(username, password) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Admin Credential");
  if (!sheet) return { success: false, error: "Admin Credential sheet not found." };

  var rows    = sheet.getDataRange().getValues();
  var headers = rows.shift(); // [ Full Name, Username, Password, Email Address, Department, Employee ID, Role ]
  var COL_FULL = headers.indexOf("Full Name");
  var COL_USER = headers.indexOf("Username");
  var COL_PW   = headers.indexOf("Password");       // hashed
  var COL_MAIL = headers.indexOf("Email Address");
  var COL_DEPT = headers.indexOf("Department");
  var COL_EID  = headers.indexOf("Employee ID");
  var COL_ROLE = headers.indexOf("Role");

  var hashedInput = hashItForYou(password);

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[COL_USER]) === String(username)) {
      if (String(r[COL_PW]) === String(hashedInput)) {
        var ident = {
          fullName: r[COL_FULL],
          username: r[COL_USER],
          email:    r[COL_MAIL],
          empId:    r[COL_EID],
          role:     r[COL_ROLE] || "Admin"
        };
        var token = issueSessionToken_(ident);
        return {
          success: true,
          token: token,
          fullName: ident.fullName,
          username: ident.username,
          email:    ident.email,
          role:     ident.role
        };
      }
      break;
    }
  }
  return { success: false, error: "Invalid username or password." };
}

// === SECURITY LAYER · 1.2 · Secure wrappers (requireSession_ only) ===

// Admin: get full ticket table
function adminGetAllTickets(token) {
  requireIT_(token);
  return getAllTickets();
}

function adminGetLogs(token) {
  requireIT_(token);
  return getLogs();
}

function adminExportLogs(token, itid) {
  requireIT_(token);
  return exportLogs(itid);
}

function adminResendPendingApprovals(token) {
  requireIT_(token);
  resendPendingApprovals();
  return { success: true };
}

function adminUpdateTicket(token, ticketNumber, newStatus, itIncharge, resolutionRemarks, changeDetails) {
  var ident = requireIT_(token);
  if (!itIncharge) itIncharge = ident.email || ident.username; // your 1.2 behavior
  return updateTicket(ticketNumber, newStatus, itIncharge, resolutionRemarks, changeDetails);
}

function adminGetDashboardData(token, granularity, fromMonth, toMonth) {
  requireIT_(token);
  return getDashboardData(granularity, fromMonth, toMonth);
}


/**
 * Send a single notification to *all* approvers
 * letting them know “Ticket X is now in FirstApprover’s hands”
 * plus a nice details‐table.
 */
function notifyAllApprovers(ticketNumber, row, lineOfBusiness) {
  var approvers = getApprovers(lineOfBusiness);
  if (!approvers.length) return;
  var first = approvers[0];

  var labels = [
    'Ticket Number','Timestamp','Employee ID','Employee Name','Line Of Business',
    'Office Site','Email Address','Ticket Category','Priority Level',
    'Detailed Location','Issue Classification','Issue Description','Attachment',
    'Request Type','Additional Details','Status','IT-Incharge','Solution/Remarks',
    'Approval Status','Approval Timestamp','Remarks'
  ];

  var html = '<div style="' + EMAIL_CARD_STYLE + '">'
           +  '<h2 style="text-align:center;">New Service Request Submitted</h2>'
           +  '<p>Ticket <strong>' + ticketNumber +
              '</strong> is now pending approval by <strong>' + first + '</strong>.</p>'
           +  '<table style="width:100%;border-collapse:collapse;">';

  labels.forEach(function(label, i) {
    var v = row[i] || '';
    if (!v) return;

    var cellHtml = '';
    if (label === 'Attachment') {
      // show as 📎 icons
      var links = v.split(',').map(u=>u.trim()).filter(u=>u);
      cellHtml = '<td style="padding:8px;border-bottom:1px solid #eee;">'
               +  links.map(function(url) {
                   return '<a href="' + url + '" target="_blank" style="margin-right:6px;font-size:1.2em;text-decoration:none;">📎</a>';
                 }).join('')
               +  '</td>';
    }
    else if (label === 'Additional Details') {
      // pretty-print only non-empty entries
      try {
        var detailsObj = JSON.parse(v);
        // filter out keys with no value
        var entries = Object.entries(detailsObj)
                          .filter(([key,val]) => val && val.toString().trim() !== '');
        if (entries.length === 0) return;   // skip the whole row if nothing to show

        cellHtml = '<td style="padding:8px;border-bottom:1px solid #eee;">';
        entries.forEach(function([key, val]) {
          var labelFormatted = key
            .replace(/([A-Z])/g,' $1')
            .replace(/^./,c=>c.toUpperCase());
          cellHtml += '<strong>' + labelFormatted + ':</strong> ' +
                      val + '<br>';
        });
        cellHtml += '</td>';
      } catch (e) {
        // fallback to raw
        cellHtml = '<td style="padding:8px;border-bottom:1px solid #eee;">' + v + '</td>';
      }
    }
    else {
      cellHtml = '<td style="padding:8px;border-bottom:1px solid #eee;">' + v + '</td>';
    }

    html += '<tr>'
         +   '<th style="text-align:left;padding:8px;border-bottom:1px solid #eee;">'
               + label + '</th>'
         +   cellHtml
         +  '</tr>';
  });

  html += '</table></div>';

  MailApp.sendEmail({
    to: approvers.join(','),
    subject: "[Service Request] " + ticketNumber +
             " → Ongoing Approval Process",
    htmlBody: html
  });
}

/**
 * Compute SHA-256 hex digest of a string.
 */
function HEX256(str) {
  if (str == null) return '';
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    str.toString(),
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

/**
 * Server-side: SHA-256 the raw code and look it up in your imported hashes.
 */
function verifyEmployeeByCode(rawCode) {
  if (!rawCode) return null;
  // either call your local HEX256 or library version:
  const targetHash = HEX256(rawCode.trim());
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName("Employee Details");
  const data  = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][5].toString().trim() === targetHash) {
      return {
        id:   data[i][0],
        name: data[i][1],
        lob:  data[i][2]
      };
    }
  }
  return null;
}

/**
 * Returns an array like:
 * [
 *   { ticket: 'ITID000123', approvals: [
 *       { approver: 'a@x.com', ts, action, comment },
 *       { approver: 'b@x.com', … },
 *       …
 *     ]
 *   },
 *   …
 * ]
 */
function getApprovalHistory() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName("Approvals");
  if (!sheet) return [];
  const rows  = sheet.getDataRange().getValues().slice(1);
  const map   = {};
  rows.forEach(r => {
    const [ ticket, approver, ts, action, comment ] = r;
    if (!map[ticket]) map[ticket] = [];
    map[ticket].push({
      approver,
      ts:      (ts instanceof Date ? ts.toLocaleString() : ts),
      action,
      comment
    });
  });
  return Object.keys(map).map(ticket=>({
    ticket,
    approvals: map[ticket]
  }));
}

function getApproversData() {
  const ss    = getSpreadsheet();
  const sheet = ss.getSheetByName("Approvers");
  if (!sheet) return [];
  const rows  = sheet.getDataRange().getValues().slice(1);
  return rows.map(r => ({
    lob:       r[0],
    approvers: [r[1], r[2], r[3], r[4]]
  }));
}

function saveApproversData(data) {
  const ss    = getSpreadsheet();
  let sheet   = ss.getSheetByName("Approvers");
  if (!sheet) sheet = ss.insertSheet("Approvers");
  sheet.clearContents();
  sheet.appendRow(["Line Of Business","Approver 1","Approver 2","Approver 3","Approver 4"]);
  const out = data.map(r => [
    r.lob,
    r.approvers[0]||"",
    r.approvers[1]||"",
    r.approvers[2]||"",
    r.approvers[3]||""
  ]);
  if (out.length) sheet.getRange(2,1,out.length,out[0].length).setValues(out);
  return true;
}
/**
 * Returns a deduped list of all Line-of-Business values
 * from column C of the “Employee Details” sheet.
 */
function getBusinessUnits() {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName("Employee Details");
  if (!sheet) return [];
  var last = sheet.getLastRow();
  if (last < 2) return [];
  // column C is the 3rd column
  var values = sheet.getRange(2, 3, last - 1, 1).getValues()
    .map(function(r){ return r[0].toString().trim(); })
    .filter(function(v){ return v; });
  // dedupe
  var unique = Array.from(new Set(values));
  return unique.sort();
}

