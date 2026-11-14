/**
 * PORTFOLIO NEXUS: Armor-Plated Sweeper Engine
 * Immune to trailing spaces, case-sensitivity, and ghost rows.
 */
function processTransaction(e) {
  // Ignore filter changes, file/tab renames, and formatting tweaks when triggered by On Change
  if (e && e.changeType && e.changeType !== "EDIT" && e.changeType !== "INSERT_ROW") {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("List");
  if (!sheet) return;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    console.error('Could not obtain lock after 30 seconds.');
    return;
  }

  try {
    // 1. Read current spreadsheet boundaries AFTER acquiring lock
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) return;

    // Scan rows from row 2 to lastRow
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

    // Collect all existing transaction IDs to detect collisions
    const existingIds = new Set();
    if (colIndex["id"]) {
      for (let r = 0; r < rows.length; r++) {
        const rowId = (rows[r][colIndex["id"] - 1] || "").toString().trim();
        if (rowId) existingIds.add(rowId);
      }
    }

    // Helper function to append note segments idempotently using exact segment comparison
    function appendNoteSegment(existingNotes, segmentToAppend) {
      const cleanExisting = String(existingNotes || "").trim();
      const cleanSegment = String(segmentToAppend || "").trim();
      if (!cleanSegment) return cleanExisting;
      if (!cleanExisting) return cleanSegment;

      const segments = cleanExisting.split("|").map(s => s.trim());
      const alreadyPresent = segments.some(s => s.toLowerCase() === cleanSegment.toLowerCase());
      if (alreadyPresent) {
        return cleanExisting;
      }
      return cleanExisting + " | " + cleanSegment;
    }

    // Helper function to modify ID by appending a suffix, preserving complete original ID
    function modifyId(originalId, replacement) {
      if (!originalId) return "";
      const base = String(originalId).trim();
      // Avoid accidental repeated suffixing if already suffixed
      if (base.length > 8 && base.endsWith(replacement)) {
        return base;
      }
      const newId = base + replacement;
      if (existingIds.has(newId) && newId !== base) {
        console.warn(`ID collision detected: "${newId}" already exists in sheet.`);
        throw new Error(`ID collision detected: generated ID "${newId}" already exists.`);
      }
      existingIds.add(newId);
      return newId;
    }

    // Helper to format ID cell as text to preserve leading zeros and numeric-looking identifiers
    function writeIdCell(targetRow, colNum, idValue) {
      const range = sheet.getRange(targetRow, colNum);
      range.setNumberFormat("@");
      range.setValue(idValue);
    }

    // Helper to resolve transaction type and amount without fragile keyword guessing
    function resolveTransaction(typeRaw, rawAmount) {
      const enteredType = (typeRaw || "").toString().trim().toUpperCase();
      const enteredAmt = parseFloat(rawAmount) || 0;
      
      if (enteredType === "IN") {
        return { finalType: "IN", finalAmt: Math.abs(enteredAmt) };
      } else if (enteredType === "OUT") {
        return { finalType: "OUT", finalAmt: -Math.abs(enteredAmt) };
      } else {
        // If Type is missing, trust the user's entered sign rather than overriding it
        return { finalType: enteredAmt >= 0 ? "IN" : "OUT", finalAmt: enteredAmt };
      }
    }

    let rowsInserted = 0;

    // Iterate through the rows
    for (let i = 0; i < rows.length; i++) {
      let rowData = rows[i];
      let actualRowNumber = startRow + i + rowsInserted;
      try {
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
          const sourceAccount = (data["account"] || "").toString().trim();
          const targetAccount = debtEntityRaw;
          const enteredAmt = parseFloat(data["amount"] || 0);
          const originalAmt = Math.abs(enteredAmt);
          const sourceType = "OUT";
          const targetType = "IN";
          
          const isInternalTransfer = (categoryLower === "internal transfer" || categoryLower === "internal transfers");
          
          let originalNotes = (data["notes"] || "").toString().trim();
          let transferNote = originalNotes;
          if (isInternalTransfer) {
            const transferSuffix = "Transfer to " + targetAccount;
            transferNote = appendNoteSegment(originalNotes, transferSuffix);
          }

          // Verify row still contains expected transaction ID before writing (row-shift guard)
          if (colIndex["id"] && originalId) {
            const currentCellId = String(sheet.getRange(actualRowNumber, colIndex["id"]).getValue() || "").trim();
            if (currentCellId !== originalId) {
              console.warn(`Row shift detected at row ${actualRowNumber}: expected ID "${originalId}", found "${currentCellId}". Skipping.`);
              continue;
            }
          }

          const newOutId = modifyId(originalId, "A");
          const newInId = modifyId(originalId, "B");

          // Update Row 1 (Out) - write only changed cells to preserve untouched columns and formulas
          if (colIndex["id"]) writeIdCell(actualRowNumber, colIndex["id"], newOutId);
          if (colIndex["amount"]) sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(-originalAmt);
          if (colIndex["type"]) sheet.getRange(actualRowNumber, colIndex["type"]).setValue(sourceType);
          if (debtEntityCol) sheet.getRange(actualRowNumber, debtEntityCol).setValue(""); // BLANK
          if (colIndex["notes"]) sheet.getRange(actualRowNumber, colIndex["notes"]).setValue(transferNote);

          // Insert Row 2 (In leg) directly underneath (legitimate new-row write)
          let newRow = [...rowData];
          if (colIndex["id"]) newRow[colIndex["id"] - 1] = newInId;
          if (colIndex["account"]) newRow[colIndex["account"] - 1] = targetAccount;
          if (colIndex["amount"]) newRow[colIndex["amount"] - 1] = originalAmt;
          if (debtEntityCol) newRow[debtEntityCol - 1] = ""; // BLANK
          if (colIndex["notes"]) newRow[colIndex["notes"] - 1] = transferNote;
          if (colIndex["type"]) newRow[colIndex["type"] - 1] = targetType;

          sheet.insertRowAfter(actualRowNumber);
          sheet.getRange(actualRowNumber + 1, 1, 1, newRow.length).setValues([newRow]);
          if (colIndex["id"]) writeIdCell(actualRowNumber + 1, colIndex["id"], newInId);
          rowsInserted++;
        }

        // --- MODULE 2: SPLIT BILL / LENDING ENGINE ---
        else if (!isTransferCategory && debtEntityRaw !== "" && debtEntityLower !== "me") {
          const names = debtEntityRaw.split(",").map(p => p.trim()).filter(p => p.length > 0);
          const totalShares = names.length;
          
          if (totalShares > 0) {
            const originalAmt = parseFloat(data["amount"] || 0);
            const { finalType, finalAmt: resolvedAmt } = resolveTransaction(data["type"], originalAmt);
            const splitAmt = Math.round((Math.abs(resolvedAmt) / totalShares) * 100) / 100; 
            const finalSplitAmt = finalType === "IN" ? splitAmt : -splitAmt;

            // FIXED IDEMPOTENCY GUARD:
            // Only require statusRaw !== "" if a "status" column actually exists in the sheet!
            const currentType = (data["type"] || "").toString().trim().toUpperCase();
            const hasStatusCol = Boolean(colIndex["status"]);
            const statusRaw = hasStatusCol ? (data["status"] || "").toString().trim() : "NO_STATUS_COL";

            if (totalShares === 1 && originalAmt === finalSplitAmt && currentType === finalType && statusRaw !== "") {
              continue; // Already processed 1-person debt row — skip immediately!
            }

            const oldNotes = (data["notes"] || "").toString().trim();
            const auditString = totalShares > 1 ? `[Split: ${Math.abs(originalAmt)} / ${totalShares} = ${Math.abs(splitAmt)}]` : "";
            const finalNotes = auditString ? appendNoteSegment(oldNotes, auditString) : oldNotes;

            // Verify row still contains expected transaction ID before writing (row-shift guard)
            if (colIndex["id"] && originalId) {
              const currentCellId = String(sheet.getRange(actualRowNumber, colIndex["id"]).getValue() || "").trim();
              if (currentCellId !== originalId) {
                console.warn(`Row shift detected at row ${actualRowNumber}: expected ID "${originalId}", found "${currentCellId}". Skipping.`);
                continue;
              }
            }

            let currentInsertIndex = actualRowNumber;
            
            for (let k = 0; k < names.length; k++) {
              const currentName = names[k];
              const isMe = (currentName.toLowerCase() === "me");
              const nameToWrite = isMe ? "" : currentName;
              const suffix = (k + 1).toString();
              
              if (k === 0) {
                // Update Row 1 - write only changed cells to preserve untouched columns and formulas
                if (colIndex["id"] && totalShares > 1) {
                  const newId = modifyId(originalId, suffix);
                  writeIdCell(actualRowNumber, colIndex["id"], newId);
                }
                if (colIndex["amount"]) sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(finalSplitAmt);
                if (colIndex["type"]) sheet.getRange(actualRowNumber, colIndex["type"]).setValue(finalType);
                if (debtEntityCol) sheet.getRange(actualRowNumber, debtEntityCol).setValue(nameToWrite);
                if (colIndex["notes"]) sheet.getRange(actualRowNumber, colIndex["notes"]).setValue(finalNotes);
                if (colIndex["status"]) {
                  sheet.getRange(actualRowNumber, colIndex["status"]).setValue(isMe ? "Settled" : "Pending");
                }
                if (!isMe && colIndex["settlement date"]) {
                  sheet.getRange(actualRowNumber, colIndex["settlement date"]).setValue("");
                }
              } else {
                // Insert Row 2, 3, etc. directly underneath (legitimate new-row write)
                const newFriendId = colIndex["id"] ? modifyId(originalId, suffix) : "";
                let friendRow = [...rowData]; 
                if (colIndex["id"]) friendRow[colIndex["id"] - 1] = newFriendId;
                if (colIndex["amount"]) friendRow[colIndex["amount"] - 1] = finalSplitAmt;
                if (colIndex["type"]) friendRow[colIndex["type"] - 1] = finalType;
                if (debtEntityCol) friendRow[debtEntityCol - 1] = nameToWrite;
                if (colIndex["notes"]) friendRow[colIndex["notes"] - 1] = finalNotes;
                if (colIndex["status"]) friendRow[colIndex["status"] - 1] = isMe ? "Settled" : "Pending";
                
                if (!isMe && colIndex["settlement date"]) {
                  friendRow[colIndex["settlement date"] - 1] = ""; 
                }
                
                sheet.insertRowAfter(currentInsertIndex);
                currentInsertIndex++;
                sheet.getRange(currentInsertIndex, 1, 1, friendRow.length).setValues([friendRow]);
                if (colIndex["id"]) writeIdCell(currentInsertIndex, colIndex["id"], newFriendId);
                rowsInserted++;
              }
            }
          }
        }

        // --- MODULE 3: STANDARD EXPENSE ENGINE (FALLBACK) ---
        else if (!isTransferCategory && (debtEntityRaw === "" || debtEntityLower === "me")) {
          const originalAmt = parseFloat(data["amount"] || 0);
          const { finalType, finalAmt: targetAmt } = resolveTransaction(data["type"], originalAmt);
          
          const needsAmtChange = (originalAmt !== targetAmt);
          const needsDebtChange = (debtEntityLower === "me");
          const currentType = (data["type"] || "").toString().trim().toUpperCase();
          const needsTypeChange = (currentType !== finalType);
          
          if (needsTypeChange || needsAmtChange || needsDebtChange) {
            // Verify row still contains expected transaction ID before writing (row-shift guard)
            if (colIndex["id"] && originalId) {
              const currentCellId = String(sheet.getRange(actualRowNumber, colIndex["id"]).getValue() || "").trim();
              if (currentCellId !== originalId) {
                console.warn(`Row shift detected at row ${actualRowNumber}: expected ID "${originalId}", found "${currentCellId}". Skipping.`);
                continue;
              }
            }

            // Write only changed cells to preserve untouched columns and formulas
            if (colIndex["type"] && needsTypeChange) sheet.getRange(actualRowNumber, colIndex["type"]).setValue(finalType);
            if (colIndex["amount"] && needsAmtChange) sheet.getRange(actualRowNumber, colIndex["amount"]).setValue(targetAmt);
            if (debtEntityCol && needsDebtChange) sheet.getRange(actualRowNumber, debtEntityCol).setValue(""); // BLANK
          }
        }
      } catch (err) {
        console.error(`Error processing row ${actualRowNumber}: ${err.message}`);
        if (colIndex["notes"]) {
          const oldNotes = (rowData[colIndex["notes"] - 1] || "").toString().trim();
          const errorNote = `[SCRIPT ERROR: ${err.message}]`;
          const newNotes = oldNotes ? `${oldNotes} | ${errorNote}` : errorNote;
          sheet.getRange(actualRowNumber, colIndex["notes"]).setValue(newNotes);
        }
      }
    }
  } finally {
    lock.releaseLock();
  }
}
