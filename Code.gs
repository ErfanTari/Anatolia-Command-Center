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

  // --- FETCH LOGIC ---
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
