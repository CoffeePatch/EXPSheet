/**
 * Processes Google Form responses for expense splitting and transfers.
 * Optimized for rigorous data mapping and independent transaction splitting.
 */
function processFormResponses() {
  const ss = SpreadsheetApp.getActive();
  const formSheet = ss.getSheetByName("Form");
  const listSheet = ss.getSheetByName("List");

  if (!formSheet || !listSheet) {
    throw new Error("Missing sheet named 'Form' or 'List'. Please check sheet names.");
  }

  const data = formSheet.getDataRange().getValues();
  if (data.length < 2) return; 

  // =======================================================================
  // ⚙️ CONFIGURATION: UPDATE THESE STRINGS TO MATCH YOUR NEW HEADERS EXACTLY
  // =======================================================================
  const COL_SPLIT_PERSON = "Split Person";       // <-- Change to your Expense Person header
  const COL_TRANSFER_PERSON = "Transfer Person"; // <-- Change to your Transfer Person header
  // =======================================================================

  const headers = data[0].map(h => String(h || '').trim());
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  // Validate that the script can find your newly named columns
  if (idx[COL_SPLIT_PERSON] === undefined && idx[COL_TRANSFER_PERSON] === undefined) {
      throw new Error(`CRITICAL STOP: Cannot find '${COL_SPLIT_PERSON}' or '${COL_TRANSFER_PERSON}' in Row 1 headers.`);
  }

  const outRows = [];
  const processedRowNumbers = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const sheetRowNumber = r + 1;

    // 1. Skip already processed rows
    if (String(row[idx['Processed']]).toUpperCase() === "YES") continue;

    try {
      const timestamp = row[idx['Timestamp']];
      const dateInput = row[idx['Date']] || timestamp;
      const finalDate = formatDate(dateInput);
      
      const timeInput = String(row[idx['Time']] || '').trim();
      const finalTime = resolveTime(timeInput, timestamp);

      const title = String(row[idx['Title']] || '').trim();
      const rawAmount = Math.abs(Number(row[idx['Amount']])); 
      const trxTypeRaw = String(row[idx['Transaction Type']] || '').toLowerCase();
      
      const isTransfer = trxTypeRaw.includes('transfer');
      const isTransaction = trxTypeRaw.includes('transac');

      // -----------------------------------------------------------------
      // LOGIC: TRANSFERS
      // -----------------------------------------------------------------
      if (isTransfer) {
        const outAcc = String(row[idx['Out Account']] || '').trim();
        const inAcc = String(row[idx['In Account']] || '').trim();
        
        // Target the specific Transfer Person column securely
        const personColIndex = idx[COL_TRANSFER_PERSON];
        const person = (personColIndex !== undefined && String(row[personColIndex]).trim() !== "") 
                        ? String(row[personColIndex]).trim() 
                        : 'me';
        
        outRows.push([finalDate, finalTime, outAcc, title, -rawAmount, person, "Transfer", "Transfer Out", "Completed", "OUT", finalDate]);
        outRows.push([finalDate, finalTime, inAcc, title, rawAmount, person, "Transfer", "Transfer In", "Completed", "IN", finalDate]);
      }

      // -----------------------------------------------------------------
      // LOGIC: TRANSACTIONS (Splits & Expenses)
      // -----------------------------------------------------------------
      else if (isTransaction) {
        const account = String(row[idx['Account']] || '').trim();
        const category = String(row[idx['Category']] || '').trim();
        const direction = String(row[idx['Type']] || 'OUT').toUpperCase().trim();
        const splitNotes = String(row[idx['Transaction Notes']] || '').trim();

        // Target the specific Split Person column securely
        const splitPersonColIndex = idx[COL_SPLIT_PERSON];
        const personsRaw = (splitPersonColIndex !== undefined && String(row[splitPersonColIndex]).trim() !== "") 
                            ? String(row[splitPersonColIndex]) 
                            : 'me';
        
        // Robust regex split handles standard commas, semicolons, and newlines
        const personList = personsRaw.split(/[,;\n]+/)
          .map(p => p.trim())
          .filter(Boolean);

        const count = personList.length || 1;
        const sign = (direction === 'OUT') ? -1 : 1;
        const perHeadAmount = (rawAmount / count) * sign;

        personList.forEach(person => {
          const isMe = (person.toLowerCase() === 'me');
          const isCreditorWallet = account.toLowerCase().includes('credit') || account.toLowerCase().includes('pending');

          let rowStatus = "Pending";
          let rowSettlementDate = "";

          // Auto-settlement rule
          if (isMe && !isCreditorWallet) {
            rowStatus = "Settled";
            rowSettlementDate = finalDate;
          }

          const note = (count > 1) ? `${splitNotes} (Split: ${rawAmount}/${count})`.trim() : splitNotes;

          outRows.push([
            finalDate, finalTime, account, title, 
            perHeadAmount, person, category, note, 
            rowStatus, direction, rowSettlementDate
          ]);
        });
      }

      processedRowNumbers.push(sheetRowNumber);

    } catch (err) {
      console.error(`Row ${sheetRowNumber} Error: ${err.message}`);
    }
  }

  // Append Data
  if (outRows.length > 0) {
    const nextRow = listSheet.getLastRow() + 1;
    listSheet.getRange(nextRow, 1, outRows.length, outRows[0].length).setValues(outRows);
    console.log(`Successfully Appended ${outRows.length} distinct transaction rows.`);
  }

  // Mark Processed
  if (processedRowNumbers.length > 0) {
    const procColIdx = idx['Processed'] + 1;
    processedRowNumbers.forEach(r => {
      formSheet.getRange(r, procColIdx).setValue("YES");
    });
  }
}

/** HELPER FUNCTIONS **/
function formatDate(d) {
  try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
  catch (e) { return ""; }
}

function resolveTime(input, timestamp) {
  if (!input) return Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "HH:mm");
  return input; 
}
