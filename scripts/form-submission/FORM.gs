/**
 * PORTFOLIO NEXUS: Armor-Plated Sweeper Engine
 * Immune to trailing spaces, case-sensitivity, and ghost rows.
 */
function processTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  const startRow = Math.max(2, lastRow - 50);
  const numRows = (lastRow - startRow) + 1;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dataRange = sheet.getRange(startRow, 1, numRows, lastCol);
  const rows = dataRange.getValues();

  let colIndex = {};
  for (let j = 0; j < headers.length; j++) {
    let h = headers[j].toString().toLowerCase().trim();
    colIndex[h] = j + 1; 
  }

  for (let i = 0; i < rows.length; i++) {
    let rowData = rows[i];
    let actualRowNumber = startRow + i;

    let data = {};
    for (let j = 0; j < headers.length; j++) {
      let h = headers[j].toString().toLowerCase().trim();
      data[h] = rowData[j];
    }

    if (!data["account"] && !data["amount"]) continue;

    let category = (data["category"] || "").toString().toLowerCase().trim();
    let personRaw = (data["person"] || "").toString().trim();
    let personLower = personRaw.toLowerCase();

    // --- MODULE 1: TRANSFER & INVESTMENT LOGIC ---
    if ((category === "internal transfers" || category === "investments") && personLower !== "me" && personRaw !== "") {
      sheet.getRange(actualRowNumber, colIndex["person"]).setValue("Me");
      
      let newRow = [...rowData]; 
      if (colIndex["account"]) newRow[colIndex["account"] - 1] = personRaw; 
      if (colIndex["amount"]) newRow[colIndex["amount"] - 1] = Math.abs(parseFloat(data["amount"] || 0));
      if (colIndex["type"]) newRow[colIndex["type"] - 1] = "IN";
      if (colIndex["person"]) newRow[colIndex["person"] - 1] = "Me"; 
      if (colIndex["notes"]) newRow[colIndex["notes"] - 1] = "Transfer IN from " + (data["account"] || "") + ". " + (data["notes"] || "");
      
      sheet.appendRow(newRow);
    }

    // --- MODULE 2: SPLIT TRANSACTION LOGIC ---
    if (personRaw.includes(",")) {
      const people = personRaw.split(",").map(p => p.trim());
      const n = people.length;
      
      const originalAmt = parseFloat(data["amount"] || 0);
      const splitAmt = Math.round((originalAmt / n) * 100) / 100; 

      const oldNotes = data["notes"] || "";
      const auditString = `[Split: ${originalAmt} / ${n} = ${splitAmt}]`;
      const finalNotes = oldNotes ? `${oldNotes} | ${auditString}` : auditString;

      if (colIndex["amount"]) sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(splitAmt);
      if (colIndex["person"]) sheet.getRange(actualRowNumber, colIndex["person"]).setValue("Me"); 
      if (colIndex["notes"]) sheet.getRange(actualRowNumber, colIndex["notes"]).setValue(finalNotes);

      for (let k = 0; k < n; k++) {
        if (people[k].toLowerCase() === "me") continue; 

        let friendRow = [...rowData]; 
        if (colIndex["amount"]) friendRow[colIndex["amount"] - 1] = splitAmt;
        if (colIndex["person"]) friendRow[colIndex["person"] - 1] = people[k];
        if (colIndex["notes"]) friendRow[colIndex["notes"] - 1] = finalNotes;
        if (colIndex["status"]) friendRow[colIndex["status"] - 1] = "Pending";
        
        // FIX 1: Erase cloned settlement date for split debts
        if (colIndex["settlement date"]) friendRow[colIndex["settlement date"] - 1] = ""; 
        
        sheet.appendRow(friendRow);
      }
    }

    // --- MODULE 3: SINGLE PERSON DEBT LOGIC ---
    // Catch-all for when you assign an entire transaction to one friend
    if (!personRaw.includes(",") && personLower !== "me" && personRaw !== "" && category !== "internal transfers" && category !== "investments") {
      
      // FIX 2: Force the Status to Pending and wipe the Settlement Date
      if (colIndex["status"]) sheet.getRange(actualRowNumber, colIndex["status"]).setValue("Pending");
      if (colIndex["settlement date"]) sheet.getRange(actualRowNumber, colIndex["settlement date"]).setValue("");
    }
  }
}