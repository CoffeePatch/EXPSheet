/**
 * PORTFOLIO NEXUS: Armor-Plated Sweeper Engine
 * Immune to trailing spaces, case-sensitivity, and ghost rows.
 */
function processTransaction() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  // 1. Expand scan to 50 rows to bypass any invisible "ghost" rows at the bottom
  const startRow = Math.max(2, lastRow - 50);
  const numRows = (lastRow - startRow) + 1;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dataRange = sheet.getRange(startRow, 1, numRows, lastCol);
  const rows = dataRange.getValues();

  // 2. Build a bulletproof header dictionary (forces lowercase, strips spaces)
  let colIndex = {};
  for (let j = 0; j < headers.length; j++) {
    let h = headers[j].toString().toLowerCase().trim();
    colIndex[h] = j + 1; // 1-based index for getRange()
  }

  // Iterate through the rows
  for (let i = 0; i < rows.length; i++) {
    let rowData = rows[i];
    let actualRowNumber = startRow + i;

    // 3. Map row data safely
    let data = {};
    for (let j = 0; j < headers.length; j++) {
      let h = headers[j].toString().toLowerCase().trim();
      data[h] = rowData[j];
    }

    // Skip entirely empty rows
    if (!data["account"] && !data["amount"]) continue;

    // 4. Sanitize variables before testing conditions
    let category = (data["category"] || "").toString().toLowerCase().trim();
    let personRaw = (data["person"] || "").toString().trim();
    let personLower = personRaw.toLowerCase();

    // --- MODULE 1: TRANSFER & INVESTMENT LOGIC ---
    if ((category === "internal transfers" || category === "investments") && personLower !== "me" && personRaw !== "") {
      
      // Overwrite Original Row to "Me"
      sheet.getRange(actualRowNumber, colIndex["person"]).setValue("Me");
      
      // Append IN Leg Safely (checks if column exists before writing)
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

      // Overwrite N=1
      if (colIndex["amount"]) sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(splitAmt);
      if (colIndex["person"]) sheet.getRange(actualRowNumber, colIndex["person"]).setValue("Me"); 
      if (colIndex["notes"]) sheet.getRange(actualRowNumber, colIndex["notes"]).setValue(finalNotes);

      // Append N-1
      for (let k = 0; k < n; k++) {
        if (people[k].toLowerCase() === "me") continue; 

        let friendRow = [...rowData]; 
        if (colIndex["amount"]) friendRow[colIndex["amount"] - 1] = splitAmt;
        if (colIndex["person"]) friendRow[colIndex["person"] - 1] = people[k];
        if (colIndex["notes"]) friendRow[colIndex["notes"] - 1] = finalNotes;
        if (colIndex["status"]) friendRow[colIndex["status"] - 1] = "Pending";
        
        // CRITICAL UPDATE: Erase the cloned settlement date for pending debts
        if (colIndex["settlement date"]) friendRow[colIndex["settlement date"] - 1] = ""; 
        
        sheet.appendRow(friendRow);
      }
    }
  }
}