function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("History");
  const trialSheet = ss.getSheetByName("Trials");

  // --- SAVE LOGIC ---
  if (e.parameter.method === "SAVE") {
    const headers = historySheet.getDataRange().getValues()[0];
    const newRow = headers.map(h => e.parameter[h] || "");
    historySheet.appendRow(newRow);

    if (String(e.parameter.Status).trim().toLowerCase() === "trial") {
      const trialHeaders = trialSheet.getDataRange().getValues()[0];
      const newTrialRow = trialHeaders.map(h => e.parameter[h] || "");
      trialSheet.appendRow(newTrialRow);
    }
    return ContentService.createTextOutput("Done").setMimeType(ContentService.MimeType.TEXT);
  }

  // --- UPDATE STATUS LOGIC ---
  if (e.parameter.method === "UPDATE_STATUS") {
    const productName = e.parameter.Product_Name;
    const newStatus = e.parameter.NewStatus;
    const data = historySheet.getDataRange().getValues();
    const headers = data[0];
    const statusColIndex = headers.indexOf("Status") + 1;

    for (let i = 1; i < data.length; i++) {
      // Find the row matching the product name where status is NOT Produced
      if (data[i][2] === productName && data[i][statusColIndex-1] !== "Produced") {
        historySheet.getRange(i + 1, statusColIndex).setValue(newStatus);
        break;
      }
    }
    return ContentService.createTextOutput("Updated").setMimeType(ContentService.MimeType.TEXT);
  }

  // --- UPDATE_ROW LOGIC ---
  if (e.parameter.method === "UPDATE_ROW") {
    var originalName = e.parameter.Original_Name || e.parameter.Product_Name;
    originalName = decodeURIComponent(originalName);
    var data = historySheet.getDataRange().getValues();
    var headers = data[0];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][headers.indexOf("Product_Name")]).trim() === originalName.trim()) {
        var updatedRow = headers.map(function(h, colIdx) {
          if (e.parameter.hasOwnProperty(h) && h !== "method" && h !== "Original_Name") {
            return e.parameter[h];
          }
          return data[i][colIdx];
        });
        historySheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
        break;
      }
    }
    return ContentService.createTextOutput("Row updated").setMimeType(ContentService.MimeType.TEXT);
  }

  // --- DELETE_ROW LOGIC ---
  if (e.parameter.method === "DELETE_ROW") {
    var productName = decodeURIComponent(e.parameter.Product_Name);

    // Delete from History sheet (reverse iteration for safe index handling)
    var hData = historySheet.getDataRange().getValues();
    var hHeaders = hData[0];
    var hNameCol = hHeaders.indexOf("Product_Name");
    for (var i = hData.length - 1; i >= 1; i--) {
      if (String(hData[i][hNameCol]).trim() === productName.trim()) {
        historySheet.deleteRow(i + 1);
      }
    }

    // Delete from Trials sheet
    var tData = trialSheet.getDataRange().getValues();
    var tHeaders = tData[0];
    var tNameCol = tHeaders.indexOf("Product_Name");
    for (var j = tData.length - 1; j >= 1; j--) {
      if (String(tData[j][tNameCol]).trim() === productName.trim()) {
        trialSheet.deleteRow(j + 1);
      }
    }

    return ContentService.createTextOutput("Row deleted").setMimeType(ContentService.MimeType.TEXT);
  }

  // --- CLEANUP_NAMES LOGIC ---
  if (e.parameter.method === "CLEANUP_NAMES") {
    var cleanupPattern = /\b(full|xxsl|20mm|antislip|organic)\b/gi;

    // Clean History sheet
    var chData = historySheet.getDataRange().getValues();
    var chHeaders = chData[0];
    var chNameCol = chHeaders.indexOf("Product_Name");
    if (chNameCol >= 0 && chData.length > 1) {
      var cleanedNames = [];
      for (var i = 1; i < chData.length; i++) {
        var name = String(chData[i][chNameCol]);
        var cleaned = name.replace(cleanupPattern, "").replace(/\s{2,}/g, " ").trim();
        cleanedNames.push([cleaned]);
      }
      historySheet.getRange(2, chNameCol + 1, cleanedNames.length, 1).setValues(cleanedNames);
    }

    // Clean Trials sheet
    var ctData = trialSheet.getDataRange().getValues();
    var ctHeaders = ctData[0];
    var ctNameCol = ctHeaders.indexOf("Product_Name");
    if (ctNameCol >= 0 && ctData.length > 1) {
      var cleanedTrialNames = [];
      for (var j = 1; j < ctData.length; j++) {
        var tName = String(ctData[j][ctNameCol]);
        var tCleaned = tName.replace(cleanupPattern, "").replace(/\s{2,}/g, " ").trim();
        cleanedTrialNames.push([tCleaned]);
      }
      trialSheet.getRange(2, ctNameCol + 1, cleanedTrialNames.length, 1).setValues(cleanedTrialNames);
    }

    return ContentService.createTextOutput("Names cleaned").setMimeType(ContentService.MimeType.TEXT);
  }

  // --- FETCH LOGIC (JSONP) ---
  const historyData = getSheetData(historySheet);
  const trialData = getSheetData(trialSheet);
  const callback = e.parameter.callback;
  const output = callback + "(" + JSON.stringify({history: historyData, trials: trialData}) + ")";
  return ContentService.createTextOutput(output).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSheetData(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => obj[header] = row[i]);
    return obj;
  });
}
