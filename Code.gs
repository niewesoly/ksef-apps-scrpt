// ============================================================
// zhr-ksef-sheets — sync faktur KSeF do Google Sheets
// ============================================================

var SHEET_PULA = "Pula";
var SHEET_CONFIG = "Konfiguracja";
var SHEET_UPRAWNIENIA = "Uprawnienia";

var PROP_DRIVE_FOLDER_ID = "ZHR_KSEF_DRIVE_FOLDER_ID";

var COL = {
  NR_FAKTURY: 1,
  NUMER_KSEF: 2,
  DATA: 3,
  SPRZEDAWCA: 4,
  KWOTA_BRUTTO: 5,
  ODBIORCA: 6,
  PDF: 7,
  WYDARZENIE: 8,
  INVOICE_ID: 9
};

var HEADERS = [
  "Nr faktury",
  "Numer KSeF",
  "Data wystawienia",
  "Sprzedawca",
  "Kwota brutto",
  "Odbiorca",
  "PDF",
  "Wydarzenie / Jednostka",
  "ID (nie usuwać)"
];

// ============================================================
// Sync — pobiera faktury z zhr-ksef i dopisuje nowe do PULI
// ============================================================

function syncInvoices() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pula = getOrCreatePula(ss);
  var config = getConfig();
  var folder = getDriveFolder();

  var existingIds = getExistingInvoiceIds(pula);
  var invoices = fetchAllInvoices(config);

  var newRows = [];
  var quotaHit = false;
  for (var i = 0; i < invoices.length; i++) {
    var inv = invoices[i];
    if (existingIds[inv.id]) continue;

    var sellerName = inv.sellerName || "";
    var buyerName = inv.buyerName || "";
    var grossAmount = inv.grossAmount !== undefined ? inv.grossAmount : "";

    var pdfDriveUrl;
    try {
      pdfDriveUrl = fetchAndSavePdf(config, folder, inv.id, inv.invoiceNumber || inv.ksefNumber || inv.id);
    } catch (err) {
      var msg = String(err);
      if (msg.indexOf("Bandwidth quota") !== -1 || msg.indexOf("quota") !== -1) {
        Logger.log("syncInvoices: limit pasma wyczerpany, przerywam. Przetworzone: " + newRows.length + "/" + invoices.length);
        quotaHit = true;
        break;
      }
      throw err;
    }

    newRows.push([
      inv.invoiceNumber || "",
      inv.ksefNumber || "",
      inv.issueDate || "",
      sellerName,
      grossAmount,
      buyerName,
      pdfDriveUrl,
      "",
      inv.id
    ]);
  }

  if (newRows.length > 0) {
    var startRow = pula.getLastRow() + 1;
    pula.getRange(startRow, 1, newRows.length, HEADERS.length).setValues(newRows);

    formatNewRows(pula, startRow, newRows.length);
    updateDropdowns(pula);
  }

  if (quotaHit) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Limit pasma Apps Script wyczerpany. Dodano " + newRows.length + " faktur. Reszta — po reset limitu (co 24h, o północy PT).",
      "Sync KSeF — limit",
      10
    );
  } else if (newRows.length > 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      newRows.length + " nowych faktur dodanych do Puli.",
      "Sync KSeF"
    );
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Brak nowych faktur.",
      "Sync KSeF"
    );
  }
}

// ============================================================
// Setup — jednorazowa konfiguracja
// ============================================================

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  getOrCreatePula(ss);
  getOrCreateConfigSheet(ss);
  getOrCreateUprawnieniaSheet(ss);

  createTimeTrigger();
  installOnEditTrigger_();

  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "Setup zakończony",
    "1. Uzupełnij zakładkę 'Konfiguracja' — nazwy wydarzeń / jednostek.\n" +
    "2. Uzupełnij zakładkę 'Uprawnienia' — jednostka + email (jeden email na wiersz).\n" +
    "3. W: Rozszerzenia → Apps Script → Ustawienia projektu → Właściwości skryptu — ustaw:\n" +
    "   • ZHR_KSEF_API_KEY — klucz API\n" +
    "   • ZHR_KSEF_TENANT_ID — ID tenanta\n" +
    "   • ZHR_KSEF_DRIVE_FOLDER_ID — ID folderu na shared drive\n" +
    "4. W: Rozszerzenia → Apps Script → Services → kliknij '+' → Drive API → v3 → Add.\n" +
    "5. Uruchom 'Synchronizuj faktury' z menu KSeF.",
    ui.ButtonSet.OK
  );
}

// ============================================================
// Menu
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("KSeF")
    .addItem("Synchronizuj faktury", "syncInvoices")
    .addItem("Udostępnij oznaczone faktury", "shareAllAssigned")
    .addSeparator()
    .addItem("Zainstaluj trigger onEdit", "reinstallOnEditTrigger")
    .addItem("Diagnostyka", "diagnostics")
    .addSeparator()
    .addItem("Jednorazowa konfiguracja", "setup")
    .addToUi();
}

function reinstallOnEditTrigger() {
  installOnEditTrigger_();
  SpreadsheetApp.getUi().alert(
    "Trigger zainstalowany",
    "Trigger onEdit dla funkcji 'onInvoiceEdit' został zainstalowany.\n\n" +
    "Sprawdź w: Rozszerzenia → Apps Script → Triggers (ikona zegara).",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function diagnostics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var report = [];

  var triggers = ScriptApp.getProjectTriggers();
  report.push("Triggerów w projekcie: " + triggers.length);
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    var eventType = "";
    try { eventType = String(t.getEventType()); } catch (e) { eventType = "?"; }
    var source = "";
    try { source = String(t.getTriggerSource()); } catch (e) { source = "?"; }
    var srcId = "";
    try { srcId = String(t.getTriggerSourceId()); } catch (e) { srcId = "?"; }
    report.push("  [" + (i+1) + "] fn=" + t.getHandlerFunction() +
                " | event=" + eventType +
                " | source=" + source +
                " | srcId=" + srcId);
  }
  report.push("Aktywny arkusz ID: " + ss.getId());

  report.push("Advanced Drive Service ('Drive' obiekt): " + (typeof Drive !== "undefined" ? "TAK" : "NIE — włącz w Services"));

  var folderProp = PropertiesService.getScriptProperties().getProperty(PROP_DRIVE_FOLDER_ID);
  report.push("Script Property " + PROP_DRIVE_FOLDER_ID + ": " + (folderProp ? "ustawione" : "BRAK"));

  var uprSheet = ss.getSheetByName(SHEET_UPRAWNIENIA);
  if (!uprSheet) {
    report.push("Zakładka 'Uprawnienia': BRAK");
  } else {
    var lastRow = uprSheet.getLastRow();
    report.push("Zakładka 'Uprawnienia': istnieje, wierszy danych: " + Math.max(0, lastRow - 1));
  }

  var konfSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!konfSheet) {
    report.push("Zakładka 'Konfiguracja': BRAK");
  } else {
    var lastRowK = konfSheet.getLastRow();
    report.push("Zakładka 'Konfiguracja': istnieje, wierszy danych: " + Math.max(0, lastRowK - 1));
  }

  ui.alert("Diagnostyka KSeF", report.join("\n"), ui.ButtonSet.OK);
}

// ============================================================
// Helpers — API
// ============================================================

function getConfig() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty("ZHR_KSEF_API_KEY");
  var tenantId = props.getProperty("ZHR_KSEF_TENANT_ID");

  if (!apiKey || !tenantId) {
    throw new Error(
      "Brak konfiguracji API. Ustaw ZHR_KSEF_API_KEY i ZHR_KSEF_TENANT_ID w właściwościach skryptu."
    );
  }

  return {
    baseUrl: "https://zhrksef.bieda.it",
    apiKey: apiKey,
    tenantId: tenantId
  };
}

function fetchAllInvoices(config) {
  var all = [];
  var page = 1;
  var pageSize = 100;

  while (true) {
    var url = config.baseUrl + "/api/v1/tenants/" + config.tenantId +
              "/invoices?page=" + page + "&pageSize=" + pageSize;

    var response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "X-API-Key": config.apiKey },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Błąd API: " + response.getResponseCode() + " — " + response.getContentText());
    }

    var data = JSON.parse(response.getContentText());
    var invoices = data.items || [];

    if (!Array.isArray(invoices) || invoices.length === 0) break;

    all = all.concat(invoices);

    if (invoices.length < pageSize) break;
    page++;
  }

  return all;
}

function fetchAndSavePdf(config, folder, invoiceId, fileName) {
  var safeName = fileName.replace(/[\/\\:*?"<>|]/g, "_") + ".pdf";

  var existingUrl = retryOnTransient_(function() {
    var existing = folder.getFilesByName(safeName);
    return existing.hasNext() ? existing.next().getUrl() : null;
  });
  if (existingUrl) return existingUrl;

  var url = config.baseUrl + "/api/v1/tenants/" + config.tenantId +
            "/invoices/" + invoiceId + "/pdf";

  var response;
  try {
    response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "X-API-Key": config.apiKey },
      muteHttpExceptions: true
    });
  } catch (err) {
    var msg = String(err);
    if (msg.indexOf("Bandwidth quota") !== -1 || msg.indexOf("quota") !== -1) {
      throw err;
    }
    Logger.log("fetchAndSavePdf network error (invoiceId=" + invoiceId + "): " + err);
    return "";
  }

  if (response.getResponseCode() !== 200) {
    return "";
  }

  var blob = response.getBlob().setName(safeName);
  var file = retryOnTransient_(function() {
    return folder.createFile(blob);
  });
  return file.getUrl();
}

// ============================================================
// Helpers — Sheets
// ============================================================

function getOrCreatePula(ss) {
  var sheet = ss.getSheetByName(SHEET_PULA);
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName(SHEET_PULA);
  }

  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== HEADERS[0]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff");

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(COL.NR_FAKTURY, 180);
    sheet.setColumnWidth(COL.NUMER_KSEF, 200);
    sheet.setColumnWidth(COL.DATA, 120);
    sheet.setColumnWidth(COL.SPRZEDAWCA, 250);
    sheet.setColumnWidth(COL.KWOTA_BRUTTO, 120);
    sheet.setColumnWidth(COL.ODBIORCA, 200);
    sheet.setColumnWidth(COL.PDF, 80);
    sheet.setColumnWidth(COL.WYDARZENIE, 200);
    sheet.setColumnWidth(COL.INVOICE_ID, 300);

    sheet.hideColumns(COL.INVOICE_ID);
  }

  return sheet;
}

function getOrCreateConfigSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG);
    sheet.getRange(1, 1, 1, 1).setValues([["Wydarzenie / Jednostka"]]);
    sheet.getRange(1, 1, 1, 1)
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("#ffffff");

    sheet.setColumnWidth(1, 300);

    sheet.getRange(2, 1).setValue("Przykład: Obóz Alfa");
    sheet.getRange(2, 1).setFontColor("#999999");
  }
  return sheet;
}

function getExistingInvoiceIds(sheet) {
  var ids = {};
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return ids;

  var values = sheet.getRange(2, COL.INVOICE_ID, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0]) ids[values[i][0]] = true;
  }
  return ids;
}


function updateDropdowns(pula) {
  var lastRow = pula.getLastRow();
  if (lastRow <= 1) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) return;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(configSheet.getRange("A2:A"), true)
    .setAllowInvalid(false)
    .build();

  pula.getRange(2, COL.WYDARZENIE, lastRow - 1, 1).setDataValidation(rule);
}

function formatNewRows(sheet, startRow, count) {
  var pdfRange = sheet.getRange(startRow, COL.PDF, count, 1);
  var values = pdfRange.getValues();
  var richValues = [];

  for (var i = 0; i < values.length; i++) {
    var url = values[i][0];
    if (url) {
      richValues.push([SpreadsheetApp.newRichTextValue().setText("PDF").setLinkUrl(url).build()]);
    } else {
      richValues.push([SpreadsheetApp.newRichTextValue().setText("—").build()]);
    }
  }

  pdfRange.setRichTextValues(richValues);
}

function getDriveFolder() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty(PROP_DRIVE_FOLDER_ID);
  if (!folderId) {
    throw new Error(
      "Brak konfiguracji folderu docelowego. Ustaw Script Property '" + PROP_DRIVE_FOLDER_ID +
      "' na ID folderu na shared drive (Project Settings → Script Properties)."
    );
  }
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    throw new Error(
      "Nie mogę otworzyć folderu o ID '" + folderId + "'. " +
      "Sprawdź czy ID jest poprawne i czy masz dostęp do folderu. (" + err + ")"
    );
  }
}

function createTimeTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "syncInvoices") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("syncInvoices")
    .timeBased()
    .everyMinutes(30)
    .create();
}

// ============================================================
// Helpers — Udostępnianie
// ============================================================

function retryOnTransient_(fn, attempts) {
  attempts = attempts || 3;
  var lastErr = null;
  for (var a = 0; a < attempts; a++) {
    try {
      return fn();
    } catch (err) {
      var msg = String(err);
      if (msg.indexOf("Bandwidth quota") !== -1) throw err;
      var isTransient =
        msg.indexOf("Service error") !== -1 ||
        msg.indexOf("Błąd usługi") !== -1 ||
        msg.indexOf("internal error") !== -1 ||
        msg.indexOf("timeout") !== -1 ||
        msg.indexOf("timed out") !== -1 ||
        msg.indexOf("rate limit") !== -1;
      if (!isTransient) throw err;
      lastErr = err;
      Logger.log("retryOnTransient_: próba " + (a + 1) + "/" + attempts + " nieudana: " + err);
      if (a < attempts - 1) Utilities.sleep(Math.pow(2, a) * 1000);
    }
  }
  throw lastErr;
}

function extractDriveFileId(url) {
  if (!url || typeof url !== "string") return null;
  var m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return null;
}

function getOrCreateUprawnieniaSheet(ss) {
  var sheet = ss.getSheetByName(SHEET_UPRAWNIENIA);
  if (sheet) return sheet;

  sheet = ss.insertSheet(SHEET_UPRAWNIENIA);
  sheet.getRange(1, 1, 1, 2).setValues([["Jednostka", "Email"]]);
  sheet.getRange(1, 1, 1, 2)
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 300);

  // Dropdown w kolumnie A z listy jednostek z Konfiguracji
  var configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (configSheet) {
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(configSheet.getRange("A2:A"), true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(2, 1, 1000, 1).setDataValidation(rule);
  }

  return sheet;
}

function getEmailsForUnit(unitName, uprawnieniaSheet) {
  if (!unitName) return [];
  var lastRow = uprawnieniaSheet.getLastRow();
  if (lastRow <= 1) return [];

  var values = uprawnieniaSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var seen = {};
  var result = [];
  for (var i = 0; i < values.length; i++) {
    var unit = values[i][0];
    var email = values[i][1];
    if (unit !== unitName) continue;
    if (!email || typeof email !== "string") continue;
    email = email.trim();
    if (!email) continue;
    if (seen[email]) continue;
    seen[email] = true;
    result.push(email);
  }
  return result;
}

function shareFileWithUser(fileId, email) {
  if (typeof Drive === "undefined") {
    Logger.log("shareFileWithUser: Advanced Drive Service nie jest włączone. Włącz w Apps Script: Services → Drive API → v3 → Add.");
    return "error";
  }
  try {
    return retryOnTransient_(function() {
      var file = DriveApp.getFileById(fileId);
      var existingEmails = {};
      var viewers = file.getViewers();
      for (var i = 0; i < viewers.length; i++) {
        existingEmails[viewers[i].getEmail()] = true;
      }
      var editors = file.getEditors();
      for (var j = 0; j < editors.length; j++) {
        existingEmails[editors[j].getEmail()] = true;
      }
      if (existingEmails[email]) {
        return "skipped";
      }

      Drive.Permissions.create(
        {
          role: "reader",
          type: "user",
          emailAddress: email
        },
        fileId,
        {
          sendNotificationEmail: true,
          supportsAllDrives: true
        }
      );
      return "added";
    });
  } catch (err) {
    Logger.log("shareFileWithUser error (fileId=" + fileId + ", email=" + email + "): " + err);
    return "error";
  }
}

function shareRowByUnit(unitName, fileId) {
  var stats = { added: 0, skipped: 0, errors: 0 };
  if (!unitName || !fileId) return stats;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uprawnieniaSheet = getOrCreateUprawnieniaSheet(ss);
  var emails = getEmailsForUnit(unitName, uprawnieniaSheet);
  if (emails.length === 0) return stats;

  for (var i = 0; i < emails.length; i++) {
    var result = shareFileWithUser(fileId, emails[i]);
    if (result === "added") stats.added++;
    else if (result === "skipped") stats.skipped++;
    else stats.errors++;
  }
  return stats;
}

function onInvoiceEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_PULA) return;

  var editedCol = e.range.getColumn();
  var editedNumCols = e.range.getNumColumns();
  // Interesuje nas tylko kolumna WYDARZENIE — albo sam ten edytowany zasięg ją zawiera
  if (editedCol > COL.WYDARZENIE || editedCol + editedNumCols - 1 < COL.WYDARZENIE) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Trigger odpalony, przetwarzam...", "KSeF — udostępnienie", 3);

  var startRow = e.range.getRow();
  var numRows = e.range.getNumRows();

  var totals = { added: 0, skipped: 0, errors: 0, rowsProcessed: 0, rowsSkippedEmptyUnit: 0, rowsSkippedNoFile: 0, rowsSkippedNoEmails: 0 };

  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row < 2) continue; // nagłówek

    var unitName = sheet.getRange(row, COL.WYDARZENIE).getValue();
    if (!unitName) { totals.rowsSkippedEmptyUnit++; continue; } // pusta wartość — nic nie robimy (no auto-revoke)

    var pdfCell = sheet.getRange(row, COL.PDF);
    var richText = pdfCell.getRichTextValue();
    var url = richText ? richText.getLinkUrl() : pdfCell.getValue();
    var fileId = extractDriveFileId(url);
    if (!fileId) {
      Logger.log("onInvoiceEdit: row " + row + " — brak/błędny PDF URL: " + url);
      totals.rowsSkippedNoFile++;
      continue;
    }

    var stats = shareRowByUnit(unitName, fileId);
    if (stats.added === 0 && stats.skipped === 0 && stats.errors === 0) {
      totals.rowsSkippedNoEmails++;
      Logger.log("onInvoiceEdit: row " + row + " (unit='" + unitName + "') — brak emaili w Uprawnieniach");
    }
    totals.added += stats.added;
    totals.skipped += stats.skipped;
    totals.errors += stats.errors;
    totals.rowsProcessed++;
  }

  var msg = "Wierszy: " + totals.rowsProcessed +
            " | Dodano: " + totals.added +
            " | Pominięto: " + totals.skipped +
            " | Błędy: " + totals.errors;
  if (totals.rowsSkippedEmptyUnit) msg += " | Pusta jednostka: " + totals.rowsSkippedEmptyUnit;
  if (totals.rowsSkippedNoFile) msg += " | Bez PDF: " + totals.rowsSkippedNoFile;
  if (totals.rowsSkippedNoEmails) msg += " | Bez emaili: " + totals.rowsSkippedNoEmails;

  ss.toast(msg, "KSeF — udostępnienie", 8);
}

function installOnEditTrigger_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onInvoiceEdit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("onInvoiceEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

function shareAllAssigned() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pula = ss.getSheetByName(SHEET_PULA);
  var ui = SpreadsheetApp.getUi();

  if (!pula) {
    ui.alert("Brak zakładki 'Pula'. Uruchom najpierw setup.");
    return;
  }

  var lastRow = pula.getLastRow();
  if (lastRow <= 1) {
    ui.alert("Brak faktur w Puli.");
    return;
  }

  var totals = { added: 0, skipped: 0, errors: 0, rowsProcessed: 0, rowsWithoutUnit: 0, rowsWithoutFile: 0 };

  for (var row = 2; row <= lastRow; row++) {
    var unitName = pula.getRange(row, COL.WYDARZENIE).getValue();
    if (!unitName) { totals.rowsWithoutUnit++; continue; }

    var pdfCell = pula.getRange(row, COL.PDF);
    var richText = pdfCell.getRichTextValue();
    var url = richText ? richText.getLinkUrl() : pdfCell.getValue();
    var fileId = extractDriveFileId(url);
    if (!fileId) { totals.rowsWithoutFile++; continue; }

    var stats = shareRowByUnit(unitName, fileId);
    totals.added += stats.added;
    totals.skipped += stats.skipped;
    totals.errors += stats.errors;
    totals.rowsProcessed++;
  }

  ui.alert(
    "Udostępnianie zakończone",
    "Przetworzono wierszy: " + totals.rowsProcessed + "\n" +
    "Dodano uprawnień: " + totals.added + "\n" +
    "Pominięto (już miał dostęp): " + totals.skipped + "\n" +
    "Błędów: " + totals.errors + "\n\n" +
    "Wierszy bez jednostki: " + totals.rowsWithoutUnit + "\n" +
    "Wierszy bez pliku PDF: " + totals.rowsWithoutFile,
    ui.ButtonSet.OK
  );
}

// ============================================================
// Testy — manualne (uruchamiaj z edytora Apps Script)
// ============================================================

function test_extractDriveFileId() {
  var assertEq = function(actual, expected, label) {
    if (actual !== expected) {
      throw new Error("FAIL [" + label + "]: expected " + JSON.stringify(expected) + " got " + JSON.stringify(actual));
    }
  };

  assertEq(
    extractDriveFileId("https://drive.google.com/file/d/1aBcD_-xyz123/view?usp=drivesdk"),
    "1aBcD_-xyz123",
    "standard /d/ URL"
  );
  assertEq(
    extractDriveFileId("https://drive.google.com/open?id=1aBcD_-xyz123"),
    "1aBcD_-xyz123",
    "?id= URL"
  );
  assertEq(
    extractDriveFileId("https://docs.google.com/document/d/1aBcD_-xyz123/edit"),
    "1aBcD_-xyz123",
    "docs URL"
  );
  assertEq(extractDriveFileId(""), null, "empty");
  assertEq(extractDriveFileId("not a url"), null, "garbage");
  assertEq(extractDriveFileId(null), null, "null");

  Logger.log("test_extractDriveFileId: PASS");
}

function test_getEmailsForUnit() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateUprawnieniaSheet(ss);

  // Wyczyść istniejące dane (zostaw nagłówek)
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();

  // Seed: 3 wiersze, 2 jednostki, 1 duplikat
  sheet.getRange(2, 1, 4, 2).setValues([
    ["Obóz Alfa", "jan@example.com"],
    ["Obóz Alfa", "anna@example.com"],
    ["Obóz Beta", "piotr@example.com"],
    ["Obóz Alfa", "jan@example.com"]  // duplikat
  ]);
  SpreadsheetApp.flush();

  var emailsAlfa = getEmailsForUnit("Obóz Alfa", sheet);
  if (emailsAlfa.length !== 2) {
    throw new Error("FAIL: expected 2 emails for Alfa, got " + emailsAlfa.length + ": " + JSON.stringify(emailsAlfa));
  }
  if (emailsAlfa.indexOf("jan@example.com") === -1 || emailsAlfa.indexOf("anna@example.com") === -1) {
    throw new Error("FAIL: Alfa missing expected emails: " + JSON.stringify(emailsAlfa));
  }

  var emailsBeta = getEmailsForUnit("Obóz Beta", sheet);
  if (emailsBeta.length !== 1 || emailsBeta[0] !== "piotr@example.com") {
    throw new Error("FAIL: expected [piotr@example.com] for Beta, got " + JSON.stringify(emailsBeta));
  }

  var emailsGamma = getEmailsForUnit("Obóz Gamma", sheet);
  if (emailsGamma.length !== 0) {
    throw new Error("FAIL: expected [] for Gamma, got " + JSON.stringify(emailsGamma));
  }

  Logger.log("test_getEmailsForUnit: PASS");
}

function test_onInvoiceEdit_singleCell() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pula = ss.getSheetByName(SHEET_PULA);
  if (!pula) throw new Error("Brak zakładki Pula — uruchom setup() najpierw.");

  // Znajdź wiersz testowy: potrzebujemy wiersza z wypełnionym PDF URL (kol 7)
  // Dla testu — user ręcznie wskazuje wiersz w stałej poniżej
  var TEST_ROW = 2; // PODMIEŃ na realny wiersz z PDF URL

  var pdfCell = pula.getRange(TEST_ROW, COL.PDF);
  var richText = pdfCell.getRichTextValue();
  var url = richText ? richText.getLinkUrl() : pdfCell.getValue();
  if (!url) throw new Error("Wiersz " + TEST_ROW + " nie ma PDF URL");

  // Symulujemy event onEdit: range = komórka jednostki w TEST_ROW, value = "Test Unit"
  var fakeEvent = {
    range: pula.getRange(TEST_ROW, COL.WYDARZENIE),
    value: "Test Unit",
    source: ss
  };

  // Przed wywołaniem — ustaw wartość w komórce (event.range.getValue() musi to zwrócić)
  fakeEvent.range.setValue("Test Unit");
  SpreadsheetApp.flush();

  onInvoiceEdit(fakeEvent);
  Logger.log("test_onInvoiceEdit_singleCell: sprawdź toast i czy email dotarł");
}
