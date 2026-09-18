// ============================================
// RBC BANK STATEMENT EXTRACTOR
// Extracts transactions from RBC PDF statements
// Includes recursive folder scanning
// ============================================

// CONFIGURATION
var CONFIG = {
  FOLDER_ID: 'FOLDER_ID_XXX',  // Top-level folder (e.g., "Bank Statements")
  SPREADSHEET_ID: 'SPREADSHEET_ID_XXX',
  SHEET_NAME: 'Extracted Data',
  LOG_SHEET_NAME: 'Log'
};

// ============================================
// MAIN FUNCTION - WITH RECURSIVE SCANNING
// ============================================
function extractBankTransactions() {
  try {
    var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // --- SETUP MAIN DATA SHEET ---
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
      sheet.getRange(1, 1, 1, 9).setValues([
        ['Date', 'Description', 'Withdrawal', 'Deposit', 'Balance', 'Type', 'Match Keyword', 'Category', 'Source File']
      ]);
    }
    
    // --- SETUP LOG SHEET ---
    var logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = spreadsheet.insertSheet(CONFIG.LOG_SHEET_NAME);
      logSheet.getRange(1, 1, 1, 4).setValues([
        ['File Name', 'Date Processed', 'Transactions Found', 'Status']
      ]);
    }
    
    // --- GET EXISTING LOG ---
    var processedFiles = new Set();
    if (logSheet.getLastRow() > 1) {
      var logData = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < logData.length; i++) {
        if (logData[i][0]) {
          processedFiles.add(logData[i][0].toString().trim());
        }
      }
    }
    
    // --- RECURSIVELY COLLECT PDF FILES ---
    var rootFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var pdfFiles = [];
    collectPDFsRecursive(rootFolder, pdfFiles);
    
    // Sort by file path
    pdfFiles.sort(function(a, b) {
      return a.path.localeCompare(b.path);
    });
    
    Logger.log(`📁 Found ${pdfFiles.length} PDF files in all subfolders`);
    
    var allTransactions = [];
    var logEntries = [];
    var skipped = 0;
    var errors = 0;
    var processed = 0;
    
    // --- PROCESS EACH FILE ---
    for (var i = 0; i < pdfFiles.length; i++) {
      var file = pdfFiles[i].file;
      var fileName = file.getName();
      var filePath = pdfFiles[i].path;
      
      // Use full path for duplicate checking
      if (processedFiles.has(filePath)) {
        Logger.log(`⏭️ Skipping ${filePath} - already processed`);
        skipped++;
        continue;
      }
      
      try {
        Logger.log(`📄 Processing: ${filePath}`);
        var text = extractTextFromPDF(file);
        var transactions = parseRBCStatement(text);
        
        if (transactions.length > 0) {
          // Add source file to each transaction
          for (var t = 0; t < transactions.length; t++) {
            transactions[t].push(filePath);  // Add source file column
          }
          allTransactions = allTransactions.concat(transactions);
          logEntries.push([
            filePath,
            new Date().toLocaleString(),
            transactions.length,
            'Success'
          ]);
          processed++;
          Logger.log(`✅ Found ${transactions.length} transactions in ${filePath}`);
        } else {
          Logger.log(`⚠️ No transactions found in ${filePath}`);
          logEntries.push([
            filePath,
            new Date().toLocaleString(),
            0,
            'No transactions found'
          ]);
          errors++;
        }
        
      } catch (e) {
        Logger.log(`❌ Error with ${filePath}: ${e}`);
        logEntries.push([
          filePath,
          new Date().toLocaleString(),
          0,
          'Error: ' + e.message
        ]);
        errors++;
      }
    }
    
    // --- WRITE NEW DATA ---
    if (allTransactions.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, allTransactions.length, allTransactions[0].length).setValues(allTransactions);
    }
    
    // --- WRITE TO LOG ---
    if (logEntries.length > 0) {
      var logStartRow = logSheet.getLastRow() + 1;
      logSheet.getRange(logStartRow, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
    }
    
    // --- SUMMARY ---
    var message = '✅ Extraction Complete!\n\n';
    message += `📁 Folders scanned: All subfolders\n`;
    message += `📄 Files processed: ${processed}\n`;
    message += `⏭️ Skipped (already logged): ${skipped}\n`;
    message += `📊 Transactions added: ${allTransactions.length}\n`;
    message += `❌ Errors: ${errors}\n`;
    message += `📋 Log entries added: ${logEntries.length}`;
    
    Logger.log(message);
    
    try {
      SpreadsheetApp.getUi().alert(message);
    } catch (e) {
      Logger.log('UI not available, check your spreadsheet for results!');
    }
    
  } catch (error) {
    Logger.log(`❌ CRITICAL ERROR: ${error}`);
    try {
      SpreadsheetApp.getUi().alert(`❌ Error: ${error}`);
    } catch (e) {
      Logger.log('Error occurred, check logs: ' + error);
    }
  }
}

// ============================================
// RECURSIVELY COLLECT PDF FILES
// ============================================
function collectPDFsRecursive(folder, pdfFiles, currentPath) {
  // If no current path, start with folder name
  if (!currentPath) {
    currentPath = folder.getName();
  }
  
  // Get all files in this folder
  var files = folder.getFilesByType(MimeType.PDF);
  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    // Match bank statement files (adjust pattern as needed)
    if (fileName.match(/Saving|Statement|RBC|bank|statement/i)) {
      pdfFiles.push({
        file: file,
        path: currentPath + '/' + fileName
      });
    }
  }
  
  // Get all subfolders and recurse
  var subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();
    var subPath = currentPath + '/' + subFolder.getName();
    collectPDFsRecursive(subFolder, pdfFiles, subPath);
  }
}

// ============================================
// PARSE RBC STATEMENT TEXT (same as before)
// ============================================
function parseRBCStatement(text) {
  var transactions = [];
  var lines = text.split('\n');
  
  var datePattern = /(\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{2,4})/i;
  var datePatternAlt = /(\d{1,2}\/\d{1,2}\/\d{2,4})/;
  var currencyPattern = /[\d,]+\.\d{2}/g;
  
  var tableStart = false;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    
    // Look for table header
    if (line.match(/Date\s*Description\s*Withdrawals/i)) {
      tableStart = true;
      continue;
    }
    
    // Skip summary lines
    if (line.match(/Opening Balance|Closing Balance|Total deposits|Total withdrawals|Summary|Page|Statement/i)) {
      continue;
    }
    
    if (!tableStart) continue;
    
    var dateMatch = line.match(datePattern) || line.match(datePatternAlt);
    if (!dateMatch) continue;
    
    var date = dateMatch[0].trim();
    var remaining = line.substring(line.indexOf(date) + date.length).trim();
    
    var amounts = remaining.match(currencyPattern);
    if (!amounts || amounts.length < 2) continue;
    
    var balance = amounts[amounts.length - 1];
    var transactionAmount = amounts[0];
    
    var firstAmountIndex = remaining.indexOf(amounts[0]);
    var description = remaining.substring(0, firstAmountIndex).trim();
    description = description.replace(/\s+/g, ' ').trim();
    
    // --- DETECTION LOGIC ---
    var deposit = '';
    var withdrawal = '';
    var type = '';
    var matchKeyword = '';
    var isDeposit = false;
    var isWithdrawal = false;
    
    // STEP 1: Check for explicit keywords
    var depositKeywords = {
      'Payroll Deposit': /Payroll Deposit/i,
      'e-Transfer received': /e-Transfer received/i,
      'Transfer received': /Transfer received/i,
      'Deposit': /Deposit(?!.*Withdrawal)/i,
      'Credit': /Credit(?!.*Debit)/i,
      'Refund': /Refund/i,
      'Interest': /Interest/i,
      'Payroll': /Payroll(?!.*Withdrawal)/i
    };
    
    var withdrawalKeywords = {
      'Interac purchase': /Interac purchase|Interac\s*\-/i,
      'e-Transfer sent': /e-Transfer sent/i,
      'Withdrawal': /Withdrawal|withdrawal/i,
      'Debit': /Debit|debit/i,
      'Purchase': /Purchase/i,
      'Online Transfer': /Online Transfer/i,
      'Payment': /Payment(?!.*Deposit)/i,
      'Transfer sent': /Transfer sent/i
    };
    
    // Check deposit keywords
    for (var key in depositKeywords) {
      if (depositKeywords[key].test(description)) {
        isDeposit = true;
        matchKeyword = key;
        break;
      }
    }
    
    // Check withdrawal keywords (only if not already identified)
    if (!isDeposit) {
      for (var key in withdrawalKeywords) {
        if (withdrawalKeywords[key].test(description)) {
          isWithdrawal = true;
          matchKeyword = key;
          break;
        }
      }
    }
    
    // STEP 2: If still undetermined, check column headers in the line
    if (!isDeposit && !isWithdrawal) {
      if (line.match(/Withdrawals\s*\(/i)) {
        isWithdrawal = true;
        matchKeyword = 'Column: Withdrawals ($)';
      } else if (line.match(/Deposits\s*\(/i)) {
        isDeposit = true;
        matchKeyword = 'Column: Deposits ($)';
      } else {
        // STEP 3: If still undetermined, check pattern
        if (line.match(/Withdrawal\s*:\s*\$?[\d,]+\.\d{2}/i)) {
          isWithdrawal = true;
          matchKeyword = 'Pattern: Withdrawal: $XX.XX';
        } else {
          // STEP 4: Default to withdrawal
          isWithdrawal = true;
          matchKeyword = 'Default (undetermined)';
        }
      }
    }
    
    // Assign the transaction
    if (isDeposit) {
      deposit = transactionAmount;
      type = 'Deposit';
    } else if (isWithdrawal) {
      withdrawal = transactionAmount;
      type = 'Withdrawal';
    } else {
      // Safety fallback
      withdrawal = transactionAmount;
      type = 'Withdrawal';
      matchKeyword = 'Safety fallback';
    }
    
    transactions.push([
      date,
      description,
      withdrawal || '',
      deposit || '',
      balance,
      type,
      matchKeyword,
      '' // Category
    ]);
  }
  
  // Fallback if no transactions found
  if (transactions.length === 0) {
    return parseRBCStatementFallback(text);
  }
  
  return transactions;
}

// ============================================
// FALLBACK PARSER
// ============================================
function parseRBCStatementFallback(text) {
  var transactions = [];
  var lines = text.split('\n');
  var dayMonthPattern = /^(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{2,4})?/i;
  
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    
    if (line.match(/Page|Statement|Opening|Closing|Summary|Balance|Deposit|Withdrawal/i)) continue;
    if (line.match(/^\s*$/) || line.match(/^[0-9,]+\./)) continue;
    
    var dateMatch = line.match(dayMonthPattern);
    if (!dateMatch) continue;
    
    var date = dateMatch[0].trim();
    var remaining = line.substring(line.indexOf(date) + date.length).trim();
    var amounts = remaining.match(/\d{1,3}(?:,\d{3})*\.\d{2}/g);
    
    if (!amounts || amounts.length < 2) continue;
    
    var balance = amounts[amounts.length - 1];
    var transactionAmount = amounts[0];
    var firstAmountIndex = remaining.indexOf(amounts[0]);
    var description = remaining.substring(0, firstAmountIndex).trim();
    description = description.replace(/\s+/g, ' ').trim();
    
    var isDeposit = false;
    var isWithdrawal = false;
    var matchKeyword = '';
    
    if (description.match(/Payroll|e-Transfer received|Deposit|Credit|Refund|Interest/i)) {
      isDeposit = true;
      matchKeyword = 'Fallback: Deposit keyword';
    } else if (description.match(/Interac|e-Transfer sent|Purchase|Withdrawal|Debit|Online Transfer|Payment/i)) {
      isWithdrawal = true;
      matchKeyword = 'Fallback: Withdrawal keyword';
    } else {
      isWithdrawal = true;
      matchKeyword = 'Fallback: Default';
    }
    
    var deposit = isDeposit ? transactionAmount : '';
    var withdrawal = isWithdrawal ? transactionAmount : '';
    var type = isDeposit ? 'Deposit' : 'Withdrawal';
    
    transactions.push([
      date,
      description,
      withdrawal,
      deposit,
      balance,
      type,
      matchKeyword,
      '' // Category
    ]);
  }
  
  return transactions;
}

// ============================================
// EXTRACT TEXT FROM PDF - DRIVE API v3
// ============================================
function extractTextFromPDF(file) {
  try {
    var blob = file.getBlob();
    
    var fileMetadata = {
      'name': 'temp_doc_' + new Date().getTime(),
      'mimeType': 'application/vnd.google-apps.document'
    };
    
    var convertedFile = Drive.Files.create(fileMetadata, blob, {
      'fields': 'id',
      'convert': true
    });
    
    var doc = DocumentApp.openById(convertedFile.id);
    var text = doc.getBody().getText();
    
    Drive.Files.remove(convertedFile.id);
    
    return text;
    
  } catch (e) {
    Logger.log(`❌ Drive API error: ${e}`);
    return '';
  }
}

// ============================================
// TEST FUNCTIONS
// ============================================
function testFolderScanning() {
  try {
    var rootFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var pdfFiles = [];
    collectPDFsRecursive(rootFolder, pdfFiles);
    
    var results = pdfFiles.map(function(item) {
      return item.path;
    });
    
    SpreadsheetApp.getUi().alert(
      '📁 Folder Scan Complete!\n\n' +
      'Found ' + pdfFiles.length + ' PDF files.\n\n' +
      'First 20 files:\n' +
      results.slice(0, 20).join('\n') +
      (results.length > 20 ? '\n... and ' + (results.length - 20) + ' more' : '')
    );
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

function testClassification() {
  try {
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var pdfFiles = [];
    collectPDFsRecursive(folder, pdfFiles);
    
    if (pdfFiles.length === 0) {
      SpreadsheetApp.getUi().alert('❌ No PDFs found.');
      return;
    }
    
    var firstFile = pdfFiles[0];
    var text = extractTextFromPDF(firstFile.file);
    var transactions = parseRBCStatement(text);
    
    Logger.log('=== TRANSACTIONS FROM: ' + firstFile.path + ' ===');
    transactions.forEach(function(t) {
      Logger.log(t[0] + ' | ' + t[1].substring(0, 30) + '... | ' + t[5] + ' | ' + t[6]);
    });
    
    SpreadsheetApp.getUi().alert(
      '✅ Test complete!\n\n' +
      'File: ' + firstFile.path + '\n' +
      'Found ' + transactions.length + ' transactions.\n' +
      'Check View → Logs for details with keywords.'
    );
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================
// SHOW LOG DATA
// ============================================
function showLogData() {
  try {
    var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    if (!logSheet) {
      SpreadsheetApp.getUi().alert('❌ Log sheet not found.');
      return;
    }
    
    var lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      SpreadsheetApp.getUi().alert('No entries in log yet.');
      return;
    }
    
    var data = logSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    var files = data.map(function(row) { 
      return row[0] + ' (' + row[1] + ') - ' + row[2] + ' transactions'; 
    });
    
    SpreadsheetApp.getUi().alert(
      '📋 Files logged:\n\n' +
      files.join('\n') +
      '\n\nTotal: ' + files.length
    );
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================
// RESET LOG
// ============================================
function resetLog() {
  try {
    var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    if (!logSheet) {
      SpreadsheetApp.getUi().alert('❌ Log sheet not found.');
      return;
    }
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '⚠️ Reset Log?',
      'This will clear the log sheet but keep your extracted data.\n\n' +
      'Use this if you want to reprocess all bank statements.',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      if (logSheet.getLastRow() > 1) {
        logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).clearContent();
      }
      SpreadsheetApp.getUi().alert('✅ Log cleared! Run "Extract Transactions" to reprocess all files.');
    }
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================
// CLEAR DATA
// ============================================
function clearData() {
  try {
    var spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
    var logSheet = spreadsheet.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    if (!sheet && !logSheet) {
      SpreadsheetApp.getUi().alert('❌ No sheets found.');
      return;
    }
    
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      '⚠️ Clear all data?',
      'This will remove ALL data from:\n' +
      '- ' + CONFIG.SHEET_NAME + '\n' +
      '- ' + CONFIG.LOG_SHEET_NAME + '\n\n' +
      'This cannot be undone!',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      if (sheet && sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).clearContent();
      }
      if (logSheet && logSheet.getLastRow() > 1) {
        logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).clearContent();
      }
      SpreadsheetApp.getUi().alert('✅ Data cleared from both sheets!');
    }
    
  } catch (e) {
    Logger.log('❌ Error: ' + e);
    SpreadsheetApp.getUi().alert('❌ Error: ' + e);
  }
}

// ============================================
// CREATE MENU
// ============================================
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📊 Bank Statement Extractor')
      .addItem('💰 Extract Transactions', 'extractBankTransactions')
      .addSeparator()
      .addItem('📁 Test Folder Scanning', 'testFolderScanning')
      .addItem('🧪 Test Classification', 'testClassification')
      .addItem('📋 Show Log Data', 'showLogData')
      .addSeparator()
      .addItem('🔄 Reset Log (Reprocess All)', 'resetLog')
      .addSeparator()
      .addItem('🗑️ Clear All Data', 'clearData')
      .addToUi();
  } catch (e) {
    Logger.log('Menu creation failed: ' + e);
  }
}
