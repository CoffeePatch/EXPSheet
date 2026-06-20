/**
 * PORTFOLIO NEXUS: Armor-Plated Sweeper Engine
 * Immune to trailing spaces, case-sensitivity, and ghost rows.
 */
function processTransaction() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List");
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) return;

  // 1. Scan ALL rows to ensure we don't miss rows inserted by AppSheet above empty/formatted ghost rows
  const startRow = 2;
  const numRows = Math.max(1, lastRow - 1);
  console.log(`Starting scan: checking ${numRows} rows (from row ${startRow} to ${lastRow}).`);

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const dataRange = sheet.getRange(startRow, 1, numRows, lastCol);
  const rows = dataRange.getValues();

  // 2. Build a bulletproof header dictionary (forces lowercase, strips spaces)
  let colIndex = {};
  for (let j = 0; j < headers.length; j++) {
    let h = headers[j].toString().toLowerCase().trim();
    colIndex[h] = j + 1; // 1-based index for getRange()
  }

  // Column helpers mapping
  const debtEntityCol = colIndex["debt entity"] || colIndex["debtentity"] || colIndex["person"];

  // Helper function to modify ID by replacing its last character
  function modifyId(originalId, replacement) {
    if (!originalId) return "";
    return originalId.toString().slice(0, -1) + replacement;
  }

  let rowsInserted = 0;

  // Iterate through the rows
  for (let i = 0; i < rows.length; i++) {
    let rowData = rows[i];
    let actualRowNumber = startRow + i + rowsInserted;

    // 3. Map row data safely
    let data = {};
    for (let j = 0; j < headers.length; j++) {
      let h = headers[j].toString().toLowerCase().trim();
      data[h] = rowData[j];
    }

    // Skip entirely empty rows
    if (!data["account"] && !data["amount"]) continue;

    // 4. Sanitize variables before testing conditions
    let categoryRaw = (data["category"] || "").toString().trim();
    let categoryLower = categoryRaw.toLowerCase();
    
    // Check for Debt Entity with fallback to Person
    let debtEntityRaw = "";
    if (debtEntityCol) {
      debtEntityRaw = (rowData[debtEntityCol - 1] || "").toString().trim();
    }
    let debtEntityLower = debtEntityRaw.toLowerCase();

    // Check if transfer category
    const isTransferCategory = 
      categoryLower === "internal transfer" || 
      categoryLower === "internal transfers" || 
      categoryLower === "investments" || 
      categoryLower === "credit card payment" || 
      categoryLower === "credit card payments" ||
      categoryLower === "credit card" ||
      categoryLower === "credit cards";

    const originalId = colIndex["id"] ? (data["id"] || "").toString().trim() : "";

    // --- MODULE 1: TRANSFER & INVESTMENT ENGINE ---
    if (isTransferCategory && debtEntityRaw !== "" && debtEntityLower !== "me") {
      const outAccount = (data["account"] || "").toString().trim();
      const inAccount = debtEntityRaw;
      const originalAmt = Math.abs(parseFloat(data["amount"] || 0));
      
      const isInternalTransfer = (categoryLower === "internal transfer" || categoryLower === "internal transfers");
      
      // Notes Logic: If Category is "Internal Transfer", Notes = Original Notes + " | Transfer to [In Account]".
      // If Category is "Investments", do not append transfer notes.
      let originalNotes = (data["notes"] || "").toString().trim();
      let transferNote = originalNotes;
      if (isInternalTransfer) {
        const transferSuffix = "Transfer to " + inAccount;
        transferNote = originalNotes ? (originalNotes + " | " + transferSuffix) : transferSuffix;
      }

      // Update Row 1 (Out) in place
      if (colIndex["id"]) {
        sheet.getRange(actualRowNumber, colIndex["id"]).setValue(modifyId(originalId, "A"));
      }
      if (colIndex["amount"]) {
        sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(-originalAmt);
      }
      if (debtEntityCol) {
        sheet.getRange(actualRowNumber, debtEntityCol).setValue(""); // BLANK
      }
      if (colIndex["notes"]) {
        sheet.getRange(actualRowNumber, colIndex["notes"]).setValue(transferNote);
      }

      // Insert Row 2 (In leg) directly underneath
      let newRow = [...rowData];
      if (colIndex["id"]) newRow[colIndex["id"] - 1] = modifyId(originalId, "B");
      if (colIndex["account"]) newRow[colIndex["account"] - 1] = inAccount;
      if (colIndex["amount"]) newRow[colIndex["amount"] - 1] = originalAmt;
      if (debtEntityCol) newRow[debtEntityCol - 1] = ""; // BLANK
      if (colIndex["notes"]) newRow[colIndex["notes"] - 1] = transferNote;
      if (colIndex["type"]) newRow[colIndex["type"] - 1] = "IN";

      sheet.insertRowAfter(actualRowNumber);
      sheet.getRange(actualRowNumber + 1, 1, 1, newRow.length).setValues([newRow]);
      rowsInserted++;
    }

    // --- MODULE 2: SPLIT BILL / LENDING ENGINE ---
    else if (!isTransferCategory && debtEntityRaw.includes(",")) {
      // Split the text by commas into an array
      const names = debtEntityRaw.split(",").map(p => p.trim()).filter(p => p.length > 0);
      const totalShares = names.length;
      
      if (totalShares > 0) {
        const originalAmt = parseFloat(data["amount"] || 0);
        const splitAmt = Math.round((originalAmt / totalShares) * 100) / 100; 
        const finalSplitAmt = splitAmt; // Keep original sign

        const oldNotes = (data["notes"] || "").toString().trim();
        // Append Split audit info if multiple people
        const auditString = totalShares > 1 ? `[Split: ${Math.abs(originalAmt)} / ${totalShares} = ${Math.abs(splitAmt)}]` : "";
        const finalNotes = auditString ? (oldNotes ? `${oldNotes} | ${auditString}` : auditString) : oldNotes;

        let currentInsertIndex = actualRowNumber;
        
        // Loop through names and update/insert rows
        for (let k = 0; k < names.length; k++) {
          const currentName = names[k];
          const isMe = (currentName.toLowerCase() === "me");
          const nameToWrite = isMe ? "" : currentName;
          const suffix = (k + 1).toString();
          
          if (k === 0) {
            // Update Row 1 in place
            if (colIndex["id"]) {
              sheet.getRange(actualRowNumber, colIndex["id"]).setValue(modifyId(originalId, suffix));
            }
            if (colIndex["amount"]) {
              sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(finalSplitAmt);
            }
            if (debtEntityCol) {
              sheet.getRange(actualRowNumber, debtEntityCol).setValue(nameToWrite);
            }
            if (colIndex["notes"]) {
              sheet.getRange(actualRowNumber, colIndex["notes"]).setValue(finalNotes);
            }
            if (!isMe) {
              if (colIndex["status"]) {
                sheet.getRange(actualRowNumber, colIndex["status"]).setValue("Pending");
              }
              if (colIndex["settlement date"]) {
                sheet.getRange(actualRowNumber, colIndex["settlement date"]).setValue("");
              }
            } else {
              if (colIndex["status"]) {
                sheet.getRange(actualRowNumber, colIndex["status"]).setValue("Settled");
              }
            }
          } else {
            // Insert Row 2, 3, etc. directly underneath
            let friendRow = [...rowData]; 
            if (colIndex["id"]) friendRow[colIndex["id"] - 1] = modifyId(originalId, suffix);
            if (colIndex["amount"]) friendRow[colIndex["amount"] - 1] = finalSplitAmt;
            if (debtEntityCol) friendRow[debtEntityCol - 1] = nameToWrite;
            if (colIndex["notes"]) friendRow[colIndex["notes"] - 1] = finalNotes;
            if (colIndex["status"]) friendRow[colIndex["status"] - 1] = isMe ? "Settled" : "Pending";
            
            // Erase the cloned settlement date for pending debts
            if (!isMe && colIndex["settlement date"]) {
              friendRow[colIndex["settlement date"] - 1] = ""; 
            }
            
            sheet.insertRowAfter(currentInsertIndex);
            currentInsertIndex++;
            sheet.getRange(currentInsertIndex, 1, 1, friendRow.length).setValues([friendRow]);
            rowsInserted++;
          }
        }
      }
    }

    // --- MODULE 3: STANDARD EXPENSE ENGINE (FALLBACK) ---
    else if (!isTransferCategory && (debtEntityRaw === "" || debtEntityLower === "me")) {
      const originalAmt = parseFloat(data["amount"] || 0);
      const targetAmt = originalAmt; // Keep original sign
      
      const needsAmtChange = (originalAmt !== targetAmt);
      const needsDebtChange = (debtEntityRaw === "me" || debtEntityLower === "me");
      
      if (needsAmtChange || needsDebtChange) {
        if (colIndex["amount"] && needsAmtChange) {
          sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(targetAmt);
        }
        if (debtEntityCol && needsDebtChange) {
          sheet.getRange(actualRowNumber, debtEntityCol).setValue(""); // BLANK
        }
        // ID and Notes remain as original ID and Original Notes
      }
    }
  }
}