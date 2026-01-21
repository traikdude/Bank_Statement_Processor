/**
 * 🏦 Bank Statement Processor - Enterprise Edition (FIXED & COMPLETE)
 * ====================================================================
 * 
 * ✅ FIXED: Standalone script uses openById() instead of getActiveSpreadsheet()
 * ✅ FIXED: All 7 IDs hardcoded in CONFIG
 * ✅ FIXED: Complete parser implementations (no placeholders!)
 * ✅ FIXED: Enhanced logging throughout
 * ✅ FIXED: Comprehensive error handling with try-catch blocks
 * 
 * @author GAS Master + Elite Architect
 * @version 2.1.0 - PRODUCTION READY
 * @date 2026-01-19
 */

// =============================================================================
// 🎯 CRITICAL FIX: HARDCODED IDs FOR STANDALONE SCRIPT
// =============================================================================

const SPREADSHEET_ID = '1XuvPyWNhB3WAOMHDO9wXkZ5AO36Cce13PH8PAojG9Eo';
const SCRIPT_ID = '1Y40DccCVEpn29uA3P0gvQpyWmINnM_9CVZ7YzLqSQPieFGUBd3s83oa9';
const COLAB_FILE_ID = '179m_x7ezIi5MGy1o23MJd4kwfKawN7QD';

// =============================================================================
// ⚙️ CONFIGURATION CONSTANTS
// =============================================================================

const CONFIG = {
  // Spreadsheet Settings
  SPREADSHEET: {
    ID: SPREADSHEET_ID, // 🔥 CRITICAL: Explicit ID for standalone script
    NAME: 'Bank Statement Processor',
    TRANSACTIONS_SHEET: 'Transactions',
    SUMMARY_SHEET: 'Summary',
    SETTINGS_SHEET: 'Settings',
    LOG_SHEET: 'Processing Log'
  },

  // 📁 Folder Settings - ALL 4 FOLDERS CONFIGURED
  FOLDERS: {
    INPUT_PDF_FOLDER_ID: '1vVQC5F8ZrKZnqv5QAWcye5VKGxCQ1c3q',
    PROCESSED_FOLDER_ID: '198Xe7BBn3ibRgXoUnOOz2NMWRcPgNtQj',
    OUTPUT_FOLDER_ID: '1GQxr8YiFnm23YM78k77DYzT0tF1CIQMp',
    ARCHIVE_FOLDER_ID: '1KQ31p9k1QSgh2dgteEWuDhrs0x3la-4n'
  },

  // Bank Statement Patterns
  BANKS: {
    CAPITAL_ONE: {
      IDENTIFIER: 'capitalone',
      NAME: 'Capital One',
      DATE_FORMAT: 'MMM d',
      PATTERNS: {
        TRANSACTION: /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+(.+?)(?:\s+(Credit|Debit))?\s*([+-]?\s*\$?[\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})$/i,
        OPENING_BALANCE: /Opening Balance\s+\$?([\d,]+\.\d{2})/i,
        CLOSING_BALANCE: /Closing Balance\s+\$?([\d,]+\.\d{2})/i,
        ACCOUNT_NUMBER: /360 Checking.*?(\d{11})/i,
        STATEMENT_PERIOD: /Statement Period\s*([\w\s,]+)/i
      }
    },
    CHASE: {
      IDENTIFIER: 'chase',
      NAME: 'Chase Bank',
      DATE_FORMAT: 'MM/dd',
      PATTERNS: {
        TRANSACTION: /^(\d{1,2})\/(\d{1,2})\s+(.+?)\s+\$?([\d,]+\.\d{2})$/i,
        OPENING_BALANCE: /Beginning Balance\s+\$?([\d,]+\.\d{2})/i,
        CLOSING_BALANCE: /Ending Balance\s+\$?([\d,]+\.\d{2})/i,
        ACCOUNT_NUMBER: /Account Number:?\s*(\d+)/i,
        STATEMENT_PERIOD: /(\w+\s+\d{1,2},?\s*\d{4})\s*through\s*(\w+\s+\d{1,2},?\s*\d{4})/i,
        DEPOSIT_SECTION: /DEPOSITS AND ADDITIONS/i,
        WITHDRAWAL_SECTION: /ELECTRONIC WITHDRAWALS/i
      }
    }
  },

  // Transaction Categories
  CATEGORIES: {
    INCOME: ['deposit', 'payroll', 'transfer received', 'credit', 'ssa', 'social security', 'refund'],
    TRANSFERS: ['zelle', 'transfer', 'venmo', 'paypal', 'cash app'],
    SUBSCRIPTIONS: ['rocket money', 'netflix', 'spotify', 'amazon prime', 'subscription'],
    SHOPPING: ['amazon', 'walmart', 'target', 'purchase'],
    BILLS: ['utilities', 'electric', 'water', 'internet', 'phone'],
    FOOD: ['restaurant', 'uber eats', 'doordash', 'grubhub'],
    OTHER: []
  },

  // Formatting
  FORMAT: {
    HEADER_COLOR: '#1a73e8',
    HEADER_TEXT_COLOR: '#ffffff',
    INCOME_COLOR: '#e6f4ea',
    EXPENSE_COLOR: '#fce8e6',
    NEUTRAL_COLOR: '#f8f9fa',
    BORDER_COLOR: '#dadce0',
    CURRENCY_FORMAT: '$#,##0.00',
    DATE_FORMAT: 'yyyy-MM-dd'
  },

  // Processing Settings
  PROCESSING: {
    MAX_FILES_PER_BATCH: 10,
    OCR_TIMEOUT_MS: 30000,
    RETRY_ATTEMPTS: 3,
    CACHE_TTL_SECONDS: 600
  }
};

// =============================================================================
// 🔧 CRITICAL FIX: SPREADSHEET ACCESS FUNCTION
// =============================================================================

/**
 * Gets the spreadsheet by explicit ID (FIXED for standalone scripts)
 * @returns {Spreadsheet} The target spreadsheet
 */
function getSpreadsheet() {
  try {
    console.log('📊 Opening spreadsheet by ID:', CONFIG.SPREADSHEET.ID);
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
    console.log('✅ Spreadsheet opened successfully:', ss.getName());
    return ss;
  } catch (error) {
    console.error('❌ Failed to open spreadsheet:', error);
    throw new Error('Cannot open spreadsheet. Check SPREADSHEET_ID: ' + error.message);
  }
}

// =============================================================================
// 📋 MENU & UI SETUP
// =============================================================================

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();

    ui.createMenu('🏦 Bank Statement Processor')
      .addItem('📁 Process New Statements', 'showProcessingDialog')
      .addItem('📤 Upload PDFs', 'showUploadDialog')
      .addSeparator()
      .addSubMenu(ui.createMenu('📊 Reports')
        .addItem('Generate Summary Report', 'generateSummaryReport')
        .addItem('Monthly Analysis', 'generateMonthlyAnalysis')
        .addItem('Category Breakdown', 'generateCategoryBreakdown'))
      .addSubMenu(ui.createMenu('📥 Export')
        .addItem('Export to CSV', 'exportToCSV')
        .addItem('Export to PDF Report', 'exportToPDFReport')
        .addItem('Export Selected Range', 'exportSelectedRange'))
      .addSeparator()
      .addSubMenu(ui.createMenu('🔗 Quick Links')
        .addItem('📁 Input Folder (PDFs)', 'openInputFolder')
        .addItem('✅ Processed Folder', 'openProcessedFolder')
        .addItem('📤 Output Folder (CSVs)', 'openOutputFolder')
        .addItem('📚 Archive Folder', 'openArchiveFolder'))
      .addSeparator()
      .addItem('⚙️ Settings', 'showSettingsDialog')
      .addItem('🧪 Test Extraction', 'testDataExtraction')
      .addItem('📖 Help & Documentation', 'showHelpDialog')
      .addItem('🔄 Refresh Data', 'refreshAllData')
      .addToUi();

    // 🐍 Python Tools Menu
    ui.createMenu('🐍 Python Tools')
      .addItem('📊 Open Colab Notebook', 'openColabNotebook')
      //.addItem('🔄 Sync Data to Colab', 'syncToColab')
      .addSeparator()
      .addItem('📂 View GitHub Repository', 'openGitHubRepo')
      .addToUi();

    console.log('✅ Menu created successfully');
  } catch (error) {
    console.error('❌ Menu creation error:', error);
  }
}

function onInstall(e) {
  onOpen(e);
  initializeSpreadsheet();
}

// =============================================================================
// 📄 PDF PROCESSING & EXTRACTION (COMPLETE WITH LOGGING)
// =============================================================================

/**
 * Extracts text from PDF using Drive OCR
 * @param {File} pdfFile - The PDF file
 * @returns {string} Extracted text
 */
function extractTextFromPDF(pdfFile) {
  try {
    console.log('📄 Extracting text from:', pdfFile.getName());
    const startTime = new Date().getTime();

    // Create temporary Google Doc from PDF using OCR
    const resource = {
      title: pdfFile.getName().replace('.pdf', '_temp_ocr'),
      mimeType: MimeType.GOOGLE_DOCS
    };

    const blob = pdfFile.getBlob();
    const tempDoc = Drive.Files.insert(resource, blob, { ocr: true, ocrLanguage: 'en' });

    console.log('🔄 OCR conversion complete, reading text...');

    // Get document content
    const doc = DocumentApp.openById(tempDoc.id);
    const text = doc.getBody().getText();

    console.log('📝 PDF Text Length:', text.length, 'characters');

    // Clean up temporary document
    DriveApp.getFileById(tempDoc.id).setTrashed(true);

    const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
    console.log('✅ Text extraction completed in', duration, 'seconds');

    return text;

  } catch (error) {
    console.error('❌ PDF extraction error:', error);

    // Fallback attempt for text-based PDFs
    try {
      console.log('🔄 Attempting fallback extraction...');
      const blob = pdfFile.getBlob();
      const text = blob.getDataAsString();
      console.log('✅ Fallback successful, text length:', text.length);
      return text;
    } catch (fallbackError) {
      console.error('❌ Fallback also failed:', fallbackError);
      throw new Error('Could not extract text from PDF: ' + error.message);
    }
  }
}

/**
 * Detects bank type from statement text
 * @param {string} text - Statement text
 * @returns {string} Bank identifier
 */
function detectBankType(text) {
  const lowerText = text.toLowerCase();

  console.log('🔍 Detecting bank type from text...');

  if (lowerText.includes('capitalone') || lowerText.includes('capital one') || lowerText.includes('360 checking')) {
    console.log('🏦 Detected: Capital One');
    return 'CAPITAL_ONE';
  }

  if (lowerText.includes('chase') || lowerText.includes('jpmorgan')) {
    console.log('🏦 Detected: Chase');
    return 'CHASE';
  }

  if (lowerText.includes('bank of america') || lowerText.includes('bofa')) {
    console.log('🏦 Detected: Bank of America');
    return 'BANK_OF_AMERICA';
  }

  if (lowerText.includes('wells fargo')) {
    console.log('🏦 Detected: Wells Fargo');
    return 'WELLS_FARGO';
  }

  console.log('⚠️ Unknown bank format');
  return 'UNKNOWN';
}

// =============================================================================
// 🏦 CAPITAL ONE PARSER (COMPLETE IMPLEMENTATION)
// =============================================================================

/**
 * Parses Capital One bank statement
 * @param {string} text - Statement text
 * @param {string} sourceFile - Source filename
 * @returns {Array} Transaction objects
 */
function parseCapitalOneStatement(text, sourceFile) {
  console.log('🔧 Parsing Capital One statement...');

  const transactions = [];
  const lines = text.split('\n').map(line => line.trim()).filter(line => line);

  // Extract statement period
  let statementPeriod = '';
  const periodPatterns = [
    /Sep(?:tember)?\s+\d{1,2}\s*[-–]\s*Sep(?:tember)?\s+\d{1,2},?\s*\d{4}/i,
    /Statement Period\s*([\w\s\d,-]+)/i,
    /(\w+\s+\d{1,2})\s*[-–]\s*(\w+\s+\d{1,2},?\s*\d{4})/i
  ];

  for (const pattern of periodPatterns) {
    const match = text.match(pattern);
    if (match) {
      statementPeriod = match[0].replace('Statement Period', '').trim();
      console.log('📅 Statement period:', statementPeriod);
      break;
    }
  }

  // Extract account number
  let accountNumber = '';
  const accountPatterns = [
    /360 Checking.*?(\d{11})/i,
    /Account.*?(\d{11})/i,
    /(\d{11})/
  ];

  for (const pattern of accountPatterns) {
    const match = text.match(pattern);
    if (match) {
      accountNumber = match[1];
      console.log('💳 Account:', '...' + accountNumber.slice(-4));
      break;
    }
  }

  // Parse transactions
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Skip non-transaction lines
    if (line.includes('Opening Balance') || line.includes('Closing Balance')) continue;
    if (line.includes('Page') || line.includes('capitalone.com')) continue;

    // Transaction pattern: "Sep 3 Description Credit +$XX.XX $XX.XX"
    const transactionRegex = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+(.+?)(?:\s+(Credit|Debit))?\s*([+-]?\s*\$?[\d,]+\.\d{2})\s+\$?([\d,]+\.\d{2})$/i;
    const match = line.match(transactionRegex);

    if (match) {
      const [, month, day, description, type, amountStr, balanceStr] = match;

      // Parse amount
      let amount = parseFloat(amountStr.replace(/[$,\s+]/g, ''));
      if (amountStr.includes('-') || type === 'Debit') {
        amount = -Math.abs(amount);
      }

      // Determine year
      const monthIndex = months.indexOf(month);
      const currentDate = new Date();
      let year = currentDate.getFullYear();
      if (monthIndex > currentDate.getMonth()) {
        year--;
      }

      const transactionDate = new Date(year, monthIndex, parseInt(day));

      transactions.push({
        id: Utilities.getUuid(),
        date: transactionDate,
        description: description.trim(),
        category: '',
        amount: amount,
        type: amount >= 0 ? 'Income' : 'Expense',
        balance: parseFloat(balanceStr.replace(/[$,]/g, '')),
        bank: 'Capital One',
        account: accountNumber,
        statementPeriod: statementPeriod,
        processedDate: new Date(),
        sourceFile: sourceFile
      });

      console.log('✅ Transaction:', transactionDate.toLocaleDateString(), description.substring(0, 30), amount);
    }
  }

  console.log('📊 Parsed', transactions.length, 'Capital One transactions');
  return transactions;
}

// =============================================================================
// 🏦 CHASE PARSER (COMPLETE IMPLEMENTATION)
// =============================================================================

/**
 * Parses Chase bank statement
 * @param {string} text - Statement text
 * @param {string} sourceFile - Source filename
 * @returns {Array} Transaction objects
 */
function parseChaseStatement(text, sourceFile) {
  console.log('🔧 Parsing Chase statement...');

  const transactions = [];
  const lines = text.split('\n').map(line => line.trim()).filter(line => line);

  // Extract statement period
  let statementPeriod = '';
  const periodMatch = text.match(/(\w+\s+\d{1,2},?\s*\d{4})\s*through\s*(\w+\s+\d{1,2},?\s*\d{4})/i);
  if (periodMatch) {
    statementPeriod = `${periodMatch[1]} - ${periodMatch[2]}`;
    console.log('📅 Statement period:', statementPeriod);
  }

  // Extract account number
  let accountNumber = '';
  const accountMatch = text.match(/Account\s*(?:Number)?:?\s*(\d{12,15})/i);
  if (accountMatch) {
    accountNumber = accountMatch[1];
    console.log('💳 Account:', '...' + accountNumber.slice(-4));
  }

  // Determine year
  let year = new Date().getFullYear();
  const yearMatch = text.match(/\d{4}/);
  if (yearMatch) {
    year = parseInt(yearMatch[0]);
  }

  // Track sections
  let isInDepositsSection = false;
  let isInWithdrawalsSection = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Section markers
    if (line.toUpperCase().includes('DEPOSITS AND ADDITIONS')) {
      isInDepositsSection = true;
      isInWithdrawalsSection = false;
      console.log('📥 Entered deposits section');
      continue;
    }

    if (line.toUpperCase().includes('ELECTRONIC WITHDRAWALS') || line.toUpperCase().includes('WITHDRAWALS')) {
      isInDepositsSection = false;
      isInWithdrawalsSection = true;
      console.log('📤 Entered withdrawals section');
      continue;
    }

    if (line.toUpperCase().includes('CHASE SAVINGS') ||
      line.includes('Beginning Balance') ||
      line.includes('Ending Balance')) {
      isInDepositsSection = false;
      isInWithdrawalsSection = false;
    }

    // Skip if not in a transaction section
    if (!isInDepositsSection && !isInWithdrawalsSection) continue;

    // Skip header/footer lines
    if (line.includes('Page') || line.includes('JPMorgan') || line.includes('Total')) continue;

    // Transaction pattern: "12/05 Description $XXX.XX"
    const transMatch = line.match(/^(\d{1,2})\/(\d{1,2})\s+(.+?)\s+\$?([\d,]+\.\d{2})$/i);

    if (transMatch) {
      const [, month, day, description, amountStr] = transMatch;

      let amount = parseFloat(amountStr.replace(/[$,]/g, ''));

      // Withdrawals are negative
      if (isInWithdrawalsSection) {
        amount = -Math.abs(amount);
      }

      const transactionDate = new Date(year, parseInt(month) - 1, parseInt(day));

      transactions.push({
        id: Utilities.getUuid(),
        date: transactionDate,
        description: description.trim(),
        category: '',
        amount: amount,
        type: amount >= 0 ? 'Income' : 'Expense',
        balance: null,
        bank: 'Chase',
        account: accountNumber,
        statementPeriod: statementPeriod,
        processedDate: new Date(),
        sourceFile: sourceFile
      });

      console.log('✅ Transaction:', transactionDate.toLocaleDateString(), description.substring(0, 30), amount);
    }
  }

  console.log('📊 Parsed', transactions.length, 'Chase transactions');
  return transactions;
}

// =============================================================================
// 🔧 GENERIC PARSER (FALLBACK)
// =============================================================================

/**
 * Generic parser for unknown bank formats
 * @param {string} text - Statement text
 * @param {string} sourceFile - Source filename
 * @returns {Array} Transaction objects
 */
function parseGenericStatement(text, sourceFile) {
  console.log('🔧 Using generic parser...');

  const transactions = [];
  const lines = text.split('\n').map(line => line.trim()).filter(line => line);

  // Generic patterns
  const patterns = [
    /^(\d{1,2})[\/ \-](\d{1,2})[\/ \-]?(\d{2,4})?\s+(.+?)\s+([+-]?\$?[\d,]+\.\d{2})$/,
    /^(\d{1,2})[\/ \-](\d{1,2})\s+(.+?)\s+([+-]?\$?[\d,]+\.\d{2})$/
  ];

  for (const line of lines) {
    for (const pattern of patterns) {
      const match = line.match(pattern);
      if (match) {
        const groups = match.slice(1);

        let month, day, year, description, amountStr;

        if (groups.length === 5) {
          [month, day, year, description, amountStr] = groups;
        } else {
          [month, day, description, amountStr] = groups;
          year = String(new Date().getFullYear());
        }

        let amount = parseFloat(amountStr.replace(/[$,\s+]/g, ''));
        if (amountStr.includes('-')) {
          amount = -Math.abs(amount);
        }

        const fullYear = year && year.length === 4 ? parseInt(year) : (year ? 2000 + parseInt(year) : new Date().getFullYear());
        const transactionDate = new Date(fullYear, parseInt(month) - 1, parseInt(day));

        transactions.push({
          id: Utilities.getUuid(),
          date: transactionDate,
          description: (description || '').trim(),
          category: '',
          amount: amount,
          type: amount >= 0 ? 'Income' : 'Expense',
          balance: null,
          bank: 'Unknown',
          account: '',
          statementPeriod: '',
          processedDate: new Date(),
          sourceFile: sourceFile
        });

        console.log('✅ Transaction:', transactionDate.toLocaleDateString(), (description || '').substring(0, 30), amount);
        break;
      }
    }
  }

  console.log('📊 Parsed', transactions.length, 'generic transactions');
  return transactions;
}

// =============================================================================
// 🎯 TRANSACTION CATEGORIZATION
// =============================================================================

/**
 * Categorizes a transaction based on description
 * @param {string} description - Transaction description
 * @returns {string} Category name
 */
function categorizeTransaction(description) {
  const lowerDesc = description.toLowerCase();

  for (const [category, keywords] of Object.entries(CONFIG.CATEGORIES)) {
    for (const keyword of keywords) {
      if (lowerDesc.includes(keyword)) {
        return category;
      }
    }
  }

  return 'OTHER';
}

// =============================================================================
// 📝 WRITE TRANSACTIONS TO SHEET (FIXED FOR STANDALONE SCRIPT)
// =============================================================================

/**
 * Writes transactions to the Transactions sheet
 * @param {Array} transactions - Transaction objects
 * @param {boolean} appendMode - Append or replace
 * @param {boolean} duplicateCheck - Check for duplicates
 */
function writeTransactionsToSheet(transactions, appendMode, duplicateCheck) {
  try {
    console.log('📝 Writing', transactions.length, 'transactions to sheet...');
    console.log('⚙️ Append mode:', appendMode, '| Duplicate check:', duplicateCheck);

    const ss = getSpreadsheet(); // 🔥 CRITICAL FIX: Use openById()
    let sheet = ss.getSheetByName(CONFIG.SPREADSHEET.TRANSACTIONS_SHEET);

    if (!sheet) {
      console.log('🔧 Transactions sheet not found, creating...');
      sheet = ss.insertSheet(CONFIG.SPREADSHEET.TRANSACTIONS_SHEET);
      setupTransactionsSheet(sheet);
    }

    // Get existing transactions for duplicate check
    let existingTransactions = [];
    if (appendMode && duplicateCheck) {
      console.log('🔍 Checking for duplicates...');
      const existingData = sheet.getDataRange().getValues();
      existingTransactions = existingData.slice(1).map(row => ({
        date: new Date(row[1]),
        description: row[2],
        amount: row[4]
      }));
      console.log('📊 Found', existingTransactions.length, 'existing transactions');
    }

    // Filter duplicates
    const newTransactions = duplicateCheck ?
      transactions.filter(t => !isDuplicate(t, existingTransactions)) :
      transactions;

    console.log('✅ New unique transactions:', newTransactions.length);

    if (newTransactions.length === 0) {
      console.log('⚠️ No new transactions to write');
      return;
    }

    // Convert to 2D array
    const dataRows = newTransactions.map(t => [
      t.id,
      t.date,
      t.description,
      t.category,
      t.amount,
      t.type,
      t.balance || '',
      t.bank,
      t.account,
      t.statementPeriod,
      t.processedDate,
      t.sourceFile
    ]);

    // Write to sheet
    if (appendMode) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, dataRows.length, 12).setValues(dataRows);
      console.log('✅ Appended', dataRows.length, 'rows starting at row', lastRow + 1);
    } else {
      sheet.getRange(2, 1, sheet.getMaxRows() - 1, 12).clearContent();
      sheet.getRange(2, 1, dataRows.length, 12).setValues(dataRows);
      console.log('✅ Replaced all data with', dataRows.length, 'rows');
    }

    // Apply conditional formatting for income/expenses
    applyConditionalFormatting(sheet);

    console.log('✅ Transaction write complete!');

  } catch (error) {
    console.error('❌ Error writing transactions:', error);
    throw error;
  }
}

/**
 * Checks if transaction is a duplicate
 * @param {Object} transaction - Transaction to check
 * @param {Array} existingTransactions - Existing transactions
 * @returns {boolean} True if duplicate
 */
function isDuplicate(transaction, existingTransactions) {
  return existingTransactions.some(existing =>
    existing.date.getTime() === transaction.date.getTime() &&
    existing.description === transaction.description &&
    existing.amount === transaction.amount
  );
}

/**
 * Applies conditional formatting to highlight income/expenses
 * @param {Sheet} sheet - The sheet to format
 */
function applyConditionalFormatting(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    // Income - green background
    const incomeRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2="Income"')
      .setBackground(CONFIG.FORMAT.INCOME_COLOR)
      .setRanges([sheet.getRange(2, 1, lastRow - 1, 12)])
      .build();

    // Expense - red background
    const expenseRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$F2="Expense"')
      .setBackground(CONFIG.FORMAT.EXPENSE_COLOR)
      .setRanges([sheet.getRange(2, 1, lastRow - 1, 12)])
      .build();

    sheet.setConditionalFormatRules([incomeRule, expenseRule]);

    console.log('✅ Conditional formatting applied');
  } catch (error) {
    console.error('⚠️ Conditional formatting error:', error);
  }
}

// =============================================================================
// 🚀 MAIN PROCESSING FUNCTION (WITH COMPREHENSIVE ERROR HANDLING)
// =============================================================================

/**
 * Processes all PDF statements in input folder
 * @param {Object} options - Processing options
 * @returns {string} Result message
 */
function processAllStatements(options) {
  const startTime = new Date();
  console.log('🚀 Starting batch processing at', startTime.toISOString());
  console.log('⚙️ Options:', JSON.stringify(options));

  let processedCount = 0;
  let errorCount = 0;
  let totalTransactions = 0;
  const errors = [];

  try {
    loadSettings();

    const inputFolder = DriveApp.getFolderById(CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID);
    const processedFolder = DriveApp.getFolderById(CONFIG.FOLDERS.PROCESSED_FOLDER_ID);
    const pdfFiles = inputFolder.getFilesByType(MimeType.PDF);

    const allTransactions = [];

    console.log('📁 Scanning input folder...');

    while (pdfFiles.hasNext()) {
      const pdfFile = pdfFiles.next();

      try {
        console.log('\n📄 Processing:', pdfFile.getName());
        logProcessing(pdfFile.getName(), 'Starting extraction...', 'processing');

        // Extract text
        const extractedText = extractTextFromPDF(pdfFile);

        if (!extractedText || extractedText.trim().length < 100) {
          throw new Error('Insufficient text extracted (less than 100 chars)');
        }

        console.log('✅ Text extracted:', extractedText.length, 'characters');

        // Detect bank
        const bankType = detectBankType(extractedText);

        // Parse transactions
        let transactions;
        if (bankType === 'CAPITAL_ONE') {
          transactions = parseCapitalOneStatement(extractedText, pdfFile.getName());
        } else if (bankType === 'CHASE') {
          transactions = parseChaseStatement(extractedText, pdfFile.getName());
        } else {
          transactions = parseGenericStatement(extractedText, pdfFile.getName());
        }

        if (transactions.length === 0) {
          logProcessing(pdfFile.getName(), 'No transactions found', 'warning');
        } else {
          // Auto-categorize
          if (options.autoCategory) {
            transactions.forEach(t => {
              t.category = categorizeTransaction(t.description);
            });
          }

          allTransactions.push(...transactions);
          totalTransactions += transactions.length;

          logProcessing(pdfFile.getName(), `Extracted ${transactions.length} transactions`, 'success');
        }

        // Move file
        if (options.moveProcessed) {
          pdfFile.moveTo(processedFolder);
          console.log('📦 Moved to processed folder');
        }

        processedCount++;

      } catch (fileError) {
        errorCount++;
        const errorMsg = `${pdfFile.getName()}: ${fileError.message}`;
        errors.push(errorMsg);
        logProcessing(pdfFile.getName(), `Error: ${fileError.message}`, 'error');
        console.error('❌ File processing error:', errorMsg);
      }
    }

    console.log('\n📊 Processing summary:');
    console.log('  - Files processed:', processedCount);
    console.log('  - Total transactions:', totalTransactions);



    console.log('  - Errors:', errorCount);

    // Write transactions
    if (allTransactions.length > 0) {
      console.log('\n💾 Writing transactions to sheet...');
      writeTransactionsToSheet(allTransactions, options.appendMode, options.duplicateCheck);
    }

    // Update summary
    updateSummarySheet();

    const duration = ((new Date() - startTime) / 1000).toFixed(1);
    const resultMessage = `✅ Processed ${processedCount} files with ${totalTransactions} transactions in ${duration}s. Errors: ${errorCount}`;

    logProcessing('System', resultMessage, 'success');
    console.log('\n' + resultMessage);

    if (errors.length > 0) {
      console.log('\n⚠️ Errors encountered:');
      errors.forEach(err => console.log('  - ' + err));
    }

    return resultMessage;

  } catch (error) {
    const errorMsg = `Critical error: ${error.message}`;
    logProcessing('System', errorMsg, 'error');
    console.error('❌ ' + errorMsg);
    console.error('Stack:', error.stack);
    throw error;
  }
}

/**
 * Logs processing events to the log sheet
 * @param {string} source - Source identifier
 * @param {string} message - Log message
 * @param {string} status - Status (success/error/warning/processing)
 * @param {string} details - Additional details
 */
function logProcessing(source, message, status, details = '') {
  try {
    const ss = getSpreadsheet(); // 🔥 FIXED
    const logSheet = ss.getSheetByName(CONFIG.SPREADSHEET.LOG_SHEET);

    if (!logSheet) return;

    const timestamp = new Date();
    const logRow = [timestamp, source, message, status.toUpperCase(), details];

    logSheet.appendRow(logRow);

    const lastRow = logSheet.getLastRow();
    const statusCell = logSheet.getRange(lastRow, 4);

    // Color-code status
    switch (status.toLowerCase()) {
      case 'success':
        statusCell.setBackground('#e6f4ea').setFontColor('#137333');
        break;
      case 'error':
        statusCell.setBackground('#fce8e6').setFontColor('#c5221f');
        break;
      case 'warning':
        statusCell.setBackground('#fef7e0').setFontColor('#b06000');
        break;
      case 'processing':
        statusCell.setBackground('#e8f0fe').setFontColor('#1967d2');
        break;
    }
  } catch (e) {
    console.log('Failed to log to sheet:', message);
  }
}

// =============================================================================
// 🧪 TEST VALIDATION FUNCTION
// =============================================================================

/**
 * Test function to validate data extraction
 * Run this from the menu to test with a single PDF
 */
function testDataExtraction() {
  console.log('\n🧪 ========== TEST DATA EXTRACTION ==========');
  console.log('📅 Test started:', new Date().toISOString());

  try {
    // Load settings
    loadSettings();
    console.log('✅ Settings loaded');
    console.log('📁 Input folder ID:', CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID);

    // Get spreadsheet
    const ss = getSpreadsheet();
    console.log('✅ Spreadsheet opened:', ss.getName());

    // Get input folder
    const inputFolder = DriveApp.getFolderById(CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID);
    console.log('✅ Input folder accessed');

    // Get first PDF
    const pdfFiles = inputFolder.getFilesByType(MimeType.PDF);

    if (!pdfFiles.hasNext()) {
      throw new Error('No PDF files found in input folder');
    }

    const testFile = pdfFiles.next();
    console.log('📄 Test file:', testFile.getName());

    // Extract text
    console.log('\n🔄 Extracting text from PDF...');
    const extractedText = extractTextFromPDF(testFile);
    console.log('✅ Text extracted');
    console.log('📏 Text length:', extractedText.length, 'characters');
    console.log('📝 First 200 characters:', extractedText.substring(0, 200));

    // Detect bank
    console.log('\n🏦 Detecting bank type...');
    const bankType = detectBankType(extractedText);
    console.log('✅ Bank detected:', bankType);

    // Parse transactions
    console.log('\n🔧 Parsing transactions...');
    let transactions;
    if (bankType === 'CAPITAL_ONE') {
      transactions = parseCapitalOneStatement(extractedText, testFile.getName());
    } else if (bankType === 'CHASE') {
      transactions = parseChaseStatement(extractedText, testFile.getName());
    } else {
      transactions = parseGenericStatement(extractedText, testFile.getName());
    }

    console.log('✅ Transactions parsed:', transactions.length);

    // Display first few transactions
    console.log('\n📊 Sample transactions:');
    transactions.slice(0, 5).forEach((t, i) => {
      console.log(`  ${i + 1}. ${t.date.toLocaleDateString()} | ${t.description.substring(0, 40)} | $${t.amount.toFixed(2)}`);
    });

    // Test write to sheet
    console.log('\n💾 Test writing to sheet...');
    writeTransactionsToSheet(transactions, false, false);
    console.log('✅ Write successful!');

    // Show results
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ Test Successful!',
      `Test completed successfully!\n\n` +
      `📄 File: ${testFile.getName()}\n` +
      `🏦 Bank: ${bankType}\n` +
      `📊 Transactions found: ${transactions.length}\n` +
      `📝 Text extracted: ${extractedText.length} characters\n\n` +
      `Check the Execution Log for detailed output.`,
      ui.ButtonSet.OK
    );

    console.log('\n✅ ========== TEST COMPLETE ==========');

  } catch (error) {
    console.error('\n❌ ========== TEST FAILED ==========');
    console.error('Error:', error.message);
    console.error('Stack:', error.stack);

    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '❌ Test Failed',
      `Test failed with error:\n\n${error.message}\n\nCheck Execution Log for details.`,
      ui.ButtonSet.OK
    );

    throw error;
  }
}

/**
 * Gets count of PDFs in input folder
 * @returns {number} Count of PDFs
 */
function getPDFCount() {
  try {
    loadSettings();
    const folder = DriveApp.getFolderById(CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID);
    const files = folder.getFilesByType(MimeType.PDF);
    let count = 0;
    while (files.hasNext()) {
      files.next();
      count++;
    }
    return count;
  } catch (error) {
    console.error('Error getting PDF count:', error);
    return 0;
  }
}

// =============================================================================
// 📊 SUMMARY SHEET UPDATE (FIXED FOR STANDALONE SCRIPT)
// =============================================================================

/**
 * Updates the Summary sheet with current statistics
 */
function updateSummarySheet() {
  try {
    console.log('📊 Updating summary sheet...');

    const ss = getSpreadsheet(); // 🔥 FIXED
    const summarySheet = ss.getSheetByName(CONFIG.SPREADSHEET.SUMMARY_SHEET);
    const transactionsSheet = ss.getSheetByName(CONFIG.SPREADSHEET.TRANSACTIONS_SHEET);

    if (!summarySheet || !transactionsSheet) {
      console.log('⚠️ Required sheets not found');
      return;
    }

    const data = transactionsSheet.getDataRange().getValues();
    const transactions = data.slice(1); // Skip header

    if (transactions.length === 0) {
      console.log('⚠️ No transactions to summarize');
      return;
    }

    // Calculate category breakdown
    const categoryStats = {};
    let totalAmount = 0;

    transactions.forEach(row => {
      const category = row[3] || 'OTHER';
      const amount = parseFloat(row[4]) || 0;

      if (!categoryStats[category]) {
        categoryStats[category] = { count: 0, total: 0 };
      }

      categoryStats[category].count++;
      categoryStats[category].total += amount;
      totalAmount += Math.abs(amount);
    });

    // Write category breakdown
    const categories = Object.keys(categoryStats);
    const categoryData = categories.map(cat => [
      cat,
      categoryStats[cat].count,
      categoryStats[cat].total,
      totalAmount > 0 ? ((Math.abs(categoryStats[cat].total) / totalAmount) * 100).toFixed(1) + '%' : '0%'
    ]);

    if (categoryData.length > 0) {
      // Clear previous data
      summarySheet.getRange(16, 1, summarySheet.getMaxRows() - 15, 4).clearContent();

      // Write new data
      summarySheet.getRange(16, 1, categoryData.length, 4).setValues(categoryData);

      // Format currency column
      summarySheet.getRange(16, 3, categoryData.length, 1).setNumberFormat(CONFIG.FORMAT.CURRENCY_FORMAT);
    }

    console.log('✅ Summary sheet updated');

  } catch (error) {
    console.error('❌ Error updating summary:', error);
  }
}

// =============================================================================
// 📤 CSV EXPORT FUNCTION (FIXED FOR STANDALONE SCRIPT)
// =============================================================================

/**
 * Exports transactions to CSV file
 */
function exportToCSV() {
  try {
    console.log('📤 Starting CSV export...');

    const ss = getSpreadsheet(); // 🔥 FIXED
    const sheet = ss.getSheetByName(CONFIG.SPREADSHEET.TRANSACTIONS_SHEET);

    if (!sheet) {
      throw new Error('Transactions sheet not found');
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('No data to export');
      return;
    }

    // Convert to CSV
    const csvContent = data.map(row =>
      row.map(cell => {
        if (cell instanceof Date) {
          return cell.toISOString().split('T')[0];
        }
        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"'))) {
          return `"${cell.replace(/"/g, '""')}"`;
        }
        return cell;
      }).join(',')
    ).join('\n');

    // Create CSV file
    const fileName = `Bank_Transactions_${new Date().toISOString().split('T')[0]}.csv`;
    const outputFolder = DriveApp.getFolderById(CONFIG.FOLDERS.OUTPUT_FOLDER_ID);
    const csvFile = outputFolder.createFile(fileName, csvContent, MimeType.CSV);

    console.log('✅ CSV exported:', fileName);

    // Show success dialog
    const url = csvFile.getUrl();
    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head><base target="_blank"></head>
        <body style="font-family: Arial; padding: 20px; text-align: center;">
          <h2>✅ Export Successful!</h2>
          <p>File: <strong>${fileName}</strong></p>
          <p>Rows: <strong>${data.length - 1}</strong></p>
          <br>
          <a href="${url}" style="display: inline-block; padding: 12px 24px; background: #1a73e8; color: white; text-decoration: none; border-radius: 4px;">
            📥 Download CSV
          </a>
        </body>
      </html>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(400)
      .setHeight(250);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📤 CSV Export');

  } catch (error) {
    console.error('❌ Export error:', error);
    SpreadsheetApp.getUi().alert('Export failed: ' + error.message);
  }
}

// =============================================================================
// 🔗 FOLDER NAVIGATION FUNCTIONS
// =============================================================================

function openInputFolder() { openFolderLink('INPUT_PDF_FOLDER_ID', 'Input PDF Folder'); }
function openProcessedFolder() { openFolderLink('PROCESSED_FOLDER_ID', 'Processed Folder'); }
function openOutputFolder() { openFolderLink('OUTPUT_FOLDER_ID', 'Output Folder'); }
function openArchiveFolder() { openFolderLink('ARCHIVE_FOLDER_ID', 'Archive Folder'); }

/**
 * Opens a folder link in the UI
 * @param {string} folderKey - Config key for folder ID
 * @param {string} folderName - Display name
 */
function openFolderLink(folderKey, folderName) {
  const folderId = CONFIG.FOLDERS[folderKey];

  if (!folderId) {
    SpreadsheetApp.getUi().alert('Folder ID not configured. Please run initialization.');
    return;
  }

  try {
    DriveApp.getFolderById(folderId);
    const url = `https://drive.google.com/drive/folders/${folderId}`;
    showLinkDialog(url, folderName);
  } catch (error) {
    SpreadsheetApp.getUi().alert('Folder not found: ' + folderName);
  }
}

/**
 * Shows a link dialog
 * @param {string} url - The URL
 * @param {string} title - Dialog title
 */
function showLinkDialog(url, title) {
  const html = `
    <!DOCTYPE html>
    <html>
      <head><base target="_blank">
        <style>
          body { font-family: Arial; padding: 20px; text-align: center; }
          a { display: inline-block; padding: 12px 24px; background: #1a73e8; color: white; text-decoration: none; border-radius: 4px; }
        </style>
      </head>
      <body>
        <h2>📁 ${title}</h2>
        <br>
        <a href="${url}" onclick="google.script.host.close();">Open in Google Drive</a>
        <p style="font-size: 11px; color: #666; margin-top: 20px;">${url}</p>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

// =============================================================================
// ⚙️ INITIALIZATION & SETUP FUNCTIONS
// =============================================================================

function initializeSpreadsheet() {
  const ss = getSpreadsheet();

  // Create sheets if needed
  ['TRANSACTIONS_SHEET', 'SUMMARY_SHEET', 'SETTINGS_SHEET', 'LOG_SHEET'].forEach(sheetKey => {
    const sheetName = CONFIG.SPREADSHEET[sheetKey];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
  });

  setupTransactionsSheet(ss.getSheetByName(CONFIG.SPREADSHEET.TRANSACTIONS_SHEET));
  setupSummarySheet(ss.getSheetByName(CONFIG.SPREADSHEET.SUMMARY_SHEET));
  setupSettingsSheet(ss.getSheetByName(CONFIG.SPREADSHEET.SETTINGS_SHEET));
  setupLogSheet(ss.getSheetByName(CONFIG.SPREADSHEET.LOG_SHEET));

  SpreadsheetApp.getUi().alert('✅ Setup Complete!', 'Bank Statement Processor initialized.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setupTransactionsSheet(sheet) {
  const headers = ['ID', 'Date', 'Description', 'Category', 'Amount', 'Type', 'Balance', 'Bank', 'Account', 'Statement Period', 'Processed Date', 'Source File'];
  sheet.getRange(1, 1, 1, 12).setValues([headers]);
  sheet.getRange(1, 1, 1, 12)
    .setBackground(CONFIG.FORMAT.HEADER_COLOR)
    .setFontColor(CONFIG.FORMAT.HEADER_TEXT_COLOR)
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange(2, 5, 1000, 1).setNumberFormat(CONFIG.FORMAT.CURRENCY_FORMAT);
  sheet.getRange(2, 7, 1000, 1).setNumberFormat(CONFIG.FORMAT.CURRENCY_FORMAT);
}

function setupSummarySheet(sheet) {
  sheet.getRange('A1').setValue('📊 Bank Statement Summary Dashboard').setFontSize(18).setFontWeight('bold');
  const overviewHeaders = ['Metric', 'Value'];
  sheet.getRange('A3:B3').setValues([overviewHeaders]);
  sheet.getRange('A3:B3').setBackground(CONFIG.FORMAT.HEADER_COLOR).setFontColor(CONFIG.FORMAT.HEADER_TEXT_COLOR).setFontWeight('bold');

  const overviewData = [
    ['Total Transactions', '=COUNTA(Transactions!A:A)-1'],
    ['Total Income', '=SUMIF(Transactions!F:F,"Income",Transactions!E:E)'],
    ['Total Expenses', '=ABS(SUMIF(Transactions!F:F,"Expense",Transactions!E:E))'],
    ['Net Cash Flow', '=SUM(Transactions!E:E)']
  ];

  sheet.getRange(4, 1, overviewData.length, 2).setValues(overviewData);
  sheet.getRange('B5:B7').setNumberFormat(CONFIG.FORMAT.CURRENCY_FORMAT);

  sheet.getRange('A14').setValue('📁 Category Breakdown').setFontSize(14).setFontWeight('bold');
  const categoryHeaders = ['Category', 'Count', 'Total Amount', '% of Total'];
  sheet.getRange('A15:D15').setValues([categoryHeaders]);
  sheet.getRange('A15:D15').setBackground(CONFIG.FORMAT.HEADER_COLOR).setFontColor(CONFIG.FORMAT.HEADER_TEXT_COLOR).setFontWeight('bold');
}

function setupSettingsSheet(sheet) {
  sheet.getRange('A1').setValue('⚙️ Settings').setFontSize(18).setFontWeight('bold');
  const settingsData = [
    ['Setting', 'Value', 'Description'],
    ['Input Folder ID', CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID, 'PDF upload folder'],
    ['Processed Folder ID', CONFIG.FOLDERS.PROCESSED_FOLDER_ID, 'Processed files'],
    ['Output Folder ID', CONFIG.FOLDERS.OUTPUT_FOLDER_ID, 'CSV exports'],
    ['Archive Folder ID', CONFIG.FOLDERS.ARCHIVE_FOLDER_ID, 'Archive storage']
  ];
  sheet.getRange(3, 1, settingsData.length, 3).setValues(settingsData);
  sheet.getRange(3, 1, 1, 3).setBackground(CONFIG.FORMAT.HEADER_COLOR).setFontColor(CONFIG.FORMAT.HEADER_TEXT_COLOR).setFontWeight('bold');
}

function setupLogSheet(sheet) {
  const headers = ['Timestamp', 'Source', 'Message', 'Status', 'Details'];
  sheet.getRange(1, 1, 1, 5).setValues([headers]);
  sheet.getRange(1, 1, 1, 5).setBackground(CONFIG.FORMAT.HEADER_COLOR).setFontColor(CONFIG.FORMAT.HEADER_TEXT_COLOR).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function loadSettings() {
  const ss = getSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SPREADSHEET.SETTINGS_SHEET);
  if (settingsSheet) {
    const values = settingsSheet.getRange('B4:B7').getValues();
    if (values[0][0]) CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID = values[0][0];
    if (values[1][0]) CONFIG.FOLDERS.PROCESSED_FOLDER_ID = values[1][0];
    if (values[2][0]) CONFIG.FOLDERS.OUTPUT_FOLDER_ID = values[2][0];
    if (values[3][0]) CONFIG.FOLDERS.ARCHIVE_FOLDER_ID = values[3][0];
  }
}

function saveSettings() {
  const ss = getSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SPREADSHEET.SETTINGS_SHEET);
  if (settingsSheet) {
    settingsSheet.getRange('B4').setValue(CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID);
    settingsSheet.getRange('B5').setValue(CONFIG.FOLDERS.PROCESSED_FOLDER_ID);
    settingsSheet.getRange('B6').setValue(CONFIG.FOLDERS.OUTPUT_FOLDER_ID);
    settingsSheet.getRange('B7').setValue(CONFIG.FOLDERS.ARCHIVE_FOLDER_ID);
  }
}


// =============================================================================
// 🎨 COMPLETE UI DIALOGS (REPLACE PLACEHOLDERS IN PART 4)
// =============================================================================

/**
 * Shows the main processing dialog with options
 */
function showProcessingDialog() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
          }

          body {
            font-family: 'Google Sans', 'Segoe UI', Arial, sans-serif;
            padding: 24px;
            background: #f8f9fa;
          }

          .container {
            max-width: 500px;
            background: white;
            border-radius: 8px;
            padding: 24px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }

          h2 {
            color: #202124;
            font-size: 20px;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 8px;
          }

          .subtitle {
            color: #5f6368;
            font-size: 13px;
            margin-bottom: 24px;
          }

          .stat-box {
            background: #e8f0fe;
            border-left: 4px solid #1a73e8;
            padding: 12px 16px;
            margin-bottom: 20px;
            border-radius: 4px;
          }

          .stat-box strong {
            color: #1967d2;
            font-size: 24px;
          }

          .option-group {
            margin-bottom: 20px;
          }

          .option-label {
            font-weight: 500;
            color: #202124;
            margin-bottom: 12px;
            font-size: 14px;
          }

          .checkbox-item {
            display: flex;
            align-items: center;
            padding: 10px;
            margin-bottom: 8px;
            border-radius: 4px;
            cursor: pointer;
            transition: background 0.2s;
          }

          .checkbox-item:hover {
            background: #f1f3f4;
          }

          .checkbox-item input[type="checkbox"] {
            width: 18px;
            height: 18px;
            margin-right: 12px;
            cursor: pointer;
          }

          .checkbox-item label {
            cursor: pointer;
            font-size: 14px;
            color: #202124;
            flex: 1;
          }

          .checkbox-item .description {
            display: block;
            font-size: 12px;
            color: #5f6368;
            margin-top: 4px;
          }

          .button-group {
            display: flex;
            gap: 12px;
            margin-top: 24px;
          }

          .btn {
            flex: 1;
            padding: 12px 24px;
            border: none;
            border-radius: 4px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
          }

          .btn-primary {
            background: #1a73e8;
            color: white;
          }

          .btn-primary:hover {
            background: #1557b0;
            box-shadow: 0 2px 4px rgba(26,115,232,0.3);
          }

          .btn-primary:disabled {
            background: #dadce0;
            cursor: not-allowed;
          }

          .btn-secondary {
            background: white;
            color: #5f6368;
            border: 1px solid #dadce0;
          }

          .btn-secondary:hover {
            background: #f8f9fa;
          }

          .loading {
            display: none;
            text-align: center;
            padding: 20px;
          }

          .loading.active {
            display: block;
          }

          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #1a73e8;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 16px;
          }

          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }

          .result {
            display: none;
            padding: 16px;
            border-radius: 4px;
            margin-top: 20px;
          }

          .result.success {
            background: #e6f4ea;
            color: #137333;
            border-left: 4px solid #34a853;
          }

          .result.error {
            background: #fce8e6;
            color: #c5221f;
            border-left: 4px solid #ea4335;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div id="mainForm">
            <h2>📄 Process Bank Statements</h2>
            <p class="subtitle">Configure processing options and run batch extraction</p>

            <div class="stat-box" id="pdfCount">
              <div style="color: #5f6368; font-size: 12px; margin-bottom: 4px;">PDFs in Input Folder</div>
              <strong>Loading...</strong>
            </div>

            <div class="option-group">
              <div class="option-label">📋 Processing Mode</div>

              <div class="checkbox-item">
                <input type="checkbox" id="appendMode" checked>
                <label for="appendMode">
                  Append Mode
                  <span class="description">Add new transactions to existing data (recommended)</span>
                </label>
              </div>

              <div class="checkbox-item">
                <input type="checkbox" id="duplicateCheck" checked>
                <label for="duplicateCheck">
                  Duplicate Check
                  <span class="description">Skip transactions already in the sheet</span>
                </label>
              </div>

              <div class="checkbox-item">
                <input type="checkbox" id="autoCategory" checked>
                <label for="autoCategory">
                  Auto-Categorize
                  <span class="description">Automatically assign categories to transactions</span>
                </label>
              </div>

              <div class="checkbox-item">
                <input type="checkbox" id="moveProcessed" checked>
                <label for="moveProcessed">
                  Move Processed Files
                  <span class="description">Move PDFs to Processed folder after extraction</span>
                </label>
              </div>
            </div>

            <div class="button-group">
              <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
              <button class="btn btn-primary" id="processBtn" onclick="startProcessing()">
                🚀 Start Processing
              </button>
            </div>
          </div>

          <div class="loading" id="loading">
            <div class="spinner"></div>
            <div style="color: #5f6368; font-size: 14px;">Processing bank statements...</div>
            <div style="color: #80868b; font-size: 12px; margin-top: 8px;">This may take a few minutes</div>
          </div>

          <div class="result" id="result"></div>
        </div>

        <script>
          // Load PDF count on page load
          window.onload = function() {
            google.script.run
              .withSuccessHandler(function(count) {
                document.getElementById('pdfCount').querySelector('strong').textContent = count + ' PDF' + (count !== 1 ? 's' : '');
                if (count === 0) {
                  document.getElementById('processBtn').disabled = true;
                  document.getElementById('processBtn').textContent = '📭 No PDFs Found';
                }
              })
              .withFailureHandler(function(error) {
                document.getElementById('pdfCount').querySelector('strong').textContent = 'Error loading count';
              })
              .getPDFCount();
          };

          function startProcessing() {
            // Get options
            const options = {
              appendMode: document.getElementById('appendMode').checked,
              duplicateCheck: document.getElementById('duplicateCheck').checked,
              autoCategory: document.getElementById('autoCategory').checked,
              moveProcessed: document.getElementById('moveProcessed').checked
            };

            // Show loading
            document.getElementById('mainForm').style.display = 'none';
            document.getElementById('loading').classList.add('active');

            // Call server-side function
            google.script.run
              .withSuccessHandler(function(message) {
                document.getElementById('loading').classList.remove('active');
                const resultDiv = document.getElementById('result');
                resultDiv.className = 'result success';
                resultDiv.textContent = '✅ ' + message;
                resultDiv.style.display = 'block';

                // Auto-close after 3 seconds
                setTimeout(function() {
                  google.script.host.close();
                }, 3000);
              })
              .withFailureHandler(function(error) {
                document.getElementById('loading').classList.remove('active');
                const resultDiv = document.getElementById('result');
                resultDiv.className = 'result error';
                resultDiv.textContent = '❌ Error: ' + error.message;
                resultDiv.style.display = 'block';

                // Show form again
                document.getElementById('mainForm').style.display = 'block';
              })
              .processAllStatements(options);
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(550)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🏦 Bank Statement Processor');
}

/**
 * Shows upload instructions dialog
 */
function showUploadDialog() {
  const folderId = CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID;
  const folderUrl = `https://drive.google.com/drive/folders/${folderId}`;

  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_blank">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 24px;
            text-align: center;
          }
          h2 { color: #202124; margin-bottom: 16px; }
          .step {
            background: #f8f9fa;
            padding: 16px;
            margin: 12px 0;
            border-radius: 8px;
            text-align: left;
          }
          .step-number {
            display: inline-block;
            background: #1a73e8;
            color: white;
            width: 28px;
            height: 28px;
            border-radius: 50%;
            text-align: center;
            line-height: 28px;
            margin-right: 12px;
            font-weight: bold;
          }
          .btn {
            display: inline-block;
            padding: 12px 24px;
            background: #1a73e8;
            color: white;
            text-decoration: none;
            border-radius: 4px;
            margin-top: 20px;
            font-weight: 500;
          }
          .btn:hover {
            background: #1557b0;
          }
        </style>
      </head>
      <body>
        <h2>📤 Upload Bank Statement PDFs</h2>

        <div class="step">
          <span class="step-number">1</span>
          <strong>Open your Input Folder</strong> by clicking the button below
        </div>

        <div class="step">
          <span class="step-number">2</span>
          <strong>Drag & drop or upload</strong> your bank statement PDF files
        </div>

        <div class="step">
          <span class="step-number">3</span>
          <strong>Return to this sheet</strong> and run "Process New Statements"
        </div>

        <a href="${folderUrl}" class="btn" onclick="setTimeout(() => google.script.host.close(), 500)">
          📁 Open Input Folder
        </a>

        <p style="color: #5f6368; font-size: 12px; margin-top: 20px;">
          Supported formats: Capital One, Chase (more banks coming soon)
        </p>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(450)
    .setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📤 Upload PDFs');
}

/**
 * Generate summary report
 */
function generateSummaryReport() {
  try {
    updateSummarySheet();

    const ss = getSpreadsheet();
    const summarySheet = ss.getSheetByName(CONFIG.SPREADSHEET.SUMMARY_SHEET);

    if (summarySheet) {
      ss.setActiveSheet(summarySheet);
      SpreadsheetApp.getUi().alert(
        '✅ Summary Updated!',
        'The Summary sheet has been refreshed with latest statistics.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

/**
 * Generate monthly analysis
 */
function generateMonthlyAnalysis() {
  try {
    const ss = getSpreadsheet();
    const transSheet = ss.getSheetByName(CONFIG.SPREADSHEET.TRANSACTIONS_SHEET);

    if (!transSheet) {
      throw new Error('Transactions sheet not found');
    }

    const data = transSheet.getDataRange().getValues();

    if (data.length <= 1) {
      SpreadsheetApp.getUi().alert('No data to analyze');
      return;
    }

    // Create or get Monthly Analysis sheet
    let monthlySheet = ss.getSheetByName('Monthly Analysis');
    if (!monthlySheet) {
      monthlySheet = ss.insertSheet('Monthly Analysis');
    } else {
      monthlySheet.clear();
    }

    // Headers
    monthlySheet.getRange('A1:D1').setValues([['Month', 'Income', 'Expenses', 'Net']]);
    monthlySheet.getRange('A1:D1')
      .setBackground(CONFIG.FORMAT.HEADER_COLOR)
      .setFontColor(CONFIG.FORMAT.HEADER_TEXT_COLOR)
      .setFontWeight('bold');

    // Group by month
    const monthlyData = {};

    for (let i = 1; i < data.length; i++) {
      const date = new Date(data[i][1]);
      const amount = parseFloat(data[i][4]) || 0;

      if (!date || isNaN(date.getTime())) continue;

      const monthKey = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM');

      if (!monthlyData[monthKey]) {
        monthlyData[monthKey] = { income: 0, expenses: 0 };
      }

      if (amount > 0) {
        monthlyData[monthKey].income += amount;
      } else {
        monthlyData[monthKey].expenses += Math.abs(amount);
      }
    }

    // Write data
    const months = Object.keys(monthlyData).sort();
    const outputData = months.map(month => [
      month,
      monthlyData[month].income,
      monthlyData[month].expenses,
      monthlyData[month].income - monthlyData[month].expenses
    ]);

    if (outputData.length > 0) {
      monthlySheet.getRange(2, 1, outputData.length, 4).setValues(outputData);
      monthlySheet.getRange(2, 2, outputData.length, 3).setNumberFormat(CONFIG.FORMAT.CURRENCY_FORMAT);
    }

    ss.setActiveSheet(monthlySheet);

    SpreadsheetApp.getUi().alert(
      '✅ Monthly Analysis Complete!',
      `Generated analysis for ${months.length} months.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

/**
 * Generate category breakdown
 */
function generateCategoryBreakdown() {
  generateSummaryReport(); // Summary already has category breakdown
}

/**
 * Export to PDF report
 */
function exportToPDFReport() {
  try {
    const ss = getSpreadsheet();
    const summarySheet = ss.getSheetByName(CONFIG.SPREADSHEET.SUMMARY_SHEET);

    if (!summarySheet) {
      throw new Error('Summary sheet not found. Generate report first.');
    }

    // Create PDF blob
    const pdfBlob = ss.getAs(MimeType.PDF);

    // Save to output folder
    const outputFolder = DriveApp.getFolderById(CONFIG.FOLDERS.OUTPUT_FOLDER_ID);
    const fileName = `Bank_Statement_Report_${new Date().toISOString().split('T')[0]}.pdf`;
    const pdfFile = outputFolder.createFile(pdfBlob.setName(fileName));

    const url = pdfFile.getUrl();

    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head><base target="_blank"></head>
        <body style="font-family: Arial; padding: 20px; text-align: center;">
          <h2>✅ PDF Report Generated!</h2>
          <p>File: <strong>${fileName}</strong></p>
          <br>
          <a href="${url}" style="display: inline-block; padding: 12px 24px; background: #1a73e8; color: white; text-decoration: none; border-radius: 4px;">
            📥 Open PDF Report
          </a>
        </body>
      </html>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(250);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📄 PDF Export');

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Export failed: ' + error.message);
  }
}

/**
 * Export selected range to CSV
 */
function exportSelectedRange() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getActiveSheet();
    const range = sheet.getActiveRange();

    if (!range) {
      SpreadsheetApp.getUi().alert('Please select a range to export');
      return;
    }

    const data = range.getValues();

    // Convert to CSV
    const csvContent = data.map(row =>
      row.map(cell => {
        if (cell instanceof Date) {
          return cell.toISOString().split('T')[0];
        }
        if (typeof cell === 'string' && (cell.includes(',') || cell.includes('"'))) {
          return `"${cell.replace(/"/g, '""')}"`;
        }
        return cell;
      }).join(',')
    ).join('\n');

    // Create file
    const fileName = `${sheet.getName()}_Export_${new Date().toISOString().split('T')[0]}.csv`;
    const outputFolder = DriveApp.getFolderById(CONFIG.FOLDERS.OUTPUT_FOLDER_ID);
    const csvFile = outputFolder.createFile(fileName, csvContent, MimeType.CSV);

    const url = csvFile.getUrl();

    const htmlContent = `
      <!DOCTYPE html>
      <html>
        <head><base target="_blank"></head>
        <body style="font-family: Arial; padding: 20px; text-align: center;">
          <h2>✅ Range Exported!</h2>
          <p>File: <strong>${fileName}</strong></p>
          <p>Rows: <strong>${data.length}</strong></p>
          <br>
          <a href="${url}" style="display: inline-block; padding: 12px 24px; background: #1a73e8; color: white; text-decoration: none; border-radius: 4px;">
            📥 Download CSV
          </a>
        </body>
      </html>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(250);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📤 CSV Export');

  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Export failed: ' + error.message);
  }
}

/**
 * Show settings dialog
 */
function showSettingsDialog() {
  const ss = getSpreadsheet();
  const settingsSheet = ss.getSheetByName(CONFIG.SPREADSHEET.SETTINGS_SHEET);

  if (settingsSheet) {
    ss.setActiveSheet(settingsSheet);
  }

  SpreadsheetApp.getUi().alert(
    '⚙️ Settings',
    'Update folder IDs in the Settings sheet as needed.\n\nChanges are saved automatically.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Show help dialog
 */
function showHelpDialog() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            line-height: 1.6;
          }
          h2 { color: #1a73e8; margin-bottom: 16px; }
          .section {
            margin-bottom: 20px;
            padding: 16px;
            background: #f8f9fa;
            border-radius: 8px;
          }
          .section h3 {
            margin-top: 0;
            color: #202124;
            font-size: 16px;
          }
          ol, ul {
            margin: 8px 0;
            padding-left: 24px;
          }
          li {
            margin: 6px 0;
          }
          .tip {
            background: #e8f0fe;
            border-left: 4px solid #1a73e8;
            padding: 12px;
            margin: 12px 0;
            font-size: 13px;
          }
        </style>
      </head>
      <body>
        <h2>📖 Bank Statement Processor - Quick Guide</h2>

        <div class="section">
          <h3>🚀 Quick Start</h3>
          <ol>
            <li>Click <strong>📤 Upload PDFs</strong> to add statement files</li>
            <li>Click <strong>📁 Process New Statements</strong> to extract data</li>
            <li>View results in the <strong>Transactions</strong> sheet</li>
          </ol>
        </div>

        <div class="section">
          <h3>🏦 Supported Banks</h3>
          <ul>
            <li>✅ <strong>Capital One</strong> - 360 Checking statements</li>
            <li>✅ <strong>Chase</strong> - Checking and savings statements</li>
            <li>⚠️ Other banks use generic parser (may vary)</li>
          </ul>
        </div>

        <div class="section">
          <h3>📊 Reports</h3>
          <ul>
            <li><strong>Summary Report</strong> - Overview of all transactions</li>
            <li><strong>Monthly Analysis</strong> - Income/expenses by month</li>
            <li><strong>Category Breakdown</strong> - Spending by category</li>
          </ul>
        </div>

        <div class="tip">
          <strong>💡 Tip:</strong> Use <strong>🧪 Test Extraction</strong> to validate setup with a single PDF before batch processing.
        </div>

        <div class="section">
          <h3>🔗 Quick Links</h3>
          <ul>
            <li><strong>Input Folder</strong> - Where to upload PDFs</li>
            <li><strong>Processed Folder</strong> - Completed files move here</li>
            <li><strong>Output Folder</strong> - CSV/PDF exports saved here</li>
          </ul>
        </div>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(500)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📖 Help & Documentation');
}

/**
 * Refresh all data
 */
function refreshAllData() {
  try {
    updateSummarySheet();

    SpreadsheetApp.getUi().alert(
      '✅ Data Refreshed!',
      'All summaries and calculations have been updated.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Error: ' + error.message);
  }
}

console.log('✅ Complete UI dialogs loaded!');

console.log('✅ Bank Statement Processor loaded successfully!');

// =============================================================================
// 🐍 PYTHON INTEGRATION TOOLS & HELPERS
// =============================================================================

function openColabNotebook() {
  const url = 'https://colab.research.google.com/drive/15azpbyehCWpjAySUW7SfZDdWPujul00B?usp=sharing';
  const htmlOutput = HtmlService
    .createHtmlOutput('<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>')
    .setWidth(100)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening Colab...');
}

function openGitHubRepo() {
  const url = 'https://github.com/traikdude/Bank_Statement_Processor';
  const htmlOutput = HtmlService
    .createHtmlOutput('<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>')
    .setWidth(100)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening GitHub...');
}

// =============================================================================
// 🌐 WEB APP DASHBOARD HANDLERS (doGet / doPost)
// =============================================================================

/**
 * Handles GET requests to the web app - Returns Dashboard HTML
 * @param {Object} e - Event object with query parameters
 * @returns {HtmlOutput} Dashboard HTML page
 */
function doGet(e) {
  try {
    console.log('🌐 Web App GET request received');

    // Get dashboard data
    const dashboardData = getDashboardData();

    // Create HTML template
    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.data = dashboardData;

    return template.evaluate()
      .setTitle('Bank Statement Processor - Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');

  } catch (error) {
    console.error('❌ doGet error:', error);

    // Fallback: Return simple status page
    return HtmlService.createHtmlOutput(generateFallbackDashboard())
      .setTitle('Bank Statement Processor')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * Handles POST requests to the web app - API endpoint
 * @param {Object} e - Event object with postData
 * @returns {ContentService.TextOutput} JSON response
 */
function doPost(e) {
  try {
    console.log('🌐 Web App POST request received');

    const params = JSON.parse(e.postData.contents);
    const action = params.action || 'status';

    let response = { success: true, timestamp: new Date().toISOString() };

    switch (action) {
      case 'status':
        response.data = getDashboardData();
        break;

      case 'process':
        // Trigger processing (async)
        const result = processAllStatements({
          appendMode: true,
          duplicateCheck: true,
          autoCategory: true,
          moveProcessed: false
        });
        response.data = { message: result };
        break;

      case 'health':
        response.data = runHealthCheck();
        break;

      default:
        response.success = false;
        response.error = 'Unknown action: ' + action;
    }

    return ContentService.createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('❌ doPost error:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Generates dashboard data for the web app
 * @returns {Object} Dashboard statistics
 */
function getDashboardData() {
  try {
    const ss = getSpreadsheet();
    const transSheet = ss.getSheetByName(CONFIG.SPREADSHEET.TRANSACTIONS_SHEET);
    const logSheet = ss.getSheetByName(CONFIG.SPREADSHEET.LOG_SHEET);

    let stats = {
      projectName: 'Bank Statement Processor',
      scriptId: SCRIPT_ID,
      spreadsheetId: SPREADSHEET_ID,
      lastUpdated: new Date().toISOString(),
      transactions: { total: 0, income: 0, expenses: 0 },
      processingLog: [],
      status: 'operational'
    };

    // Get transaction stats
    if (transSheet && transSheet.getLastRow() > 1) {
      const data = transSheet.getDataRange().getValues();
      stats.transactions.total = data.length - 1;

      for (let i = 1; i < data.length; i++) {
        const amount = parseFloat(data[i][4]) || 0;
        if (amount >= 0) {
          stats.transactions.income += amount;
        } else {
          stats.transactions.expenses += Math.abs(amount);
        }
      }
    }

    // Get recent log entries
    if (logSheet && logSheet.getLastRow() > 1) {
      const logData = logSheet.getRange(
        Math.max(2, logSheet.getLastRow() - 9),
        1,
        Math.min(10, logSheet.getLastRow() - 1),
        4
      ).getValues();

      stats.processingLog = logData.reverse().map(row => ({
        timestamp: row[0],
        source: row[1],
        message: row[2],
        status: row[3]
      }));
    }

    return stats;

  } catch (error) {
    console.error('❌ getDashboardData error:', error);
    return {
      projectName: 'Bank Statement Processor',
      status: 'error',
      error: error.message,
      lastUpdated: new Date().toISOString()
    };
  }
}

/**
 * Runs a health check on the system
 * @returns {Object} Health check results
 */
function runHealthCheck() {
  const results = {
    timestamp: new Date().toISOString(),
    checks: []
  };

  // Check 1: Spreadsheet access
  try {
    const ss = getSpreadsheet();
    results.checks.push({ name: 'Spreadsheet Access', status: 'PASS', message: ss.getName() });
  } catch (e) {
    results.checks.push({ name: 'Spreadsheet Access', status: 'FAIL', message: e.message });
  }

  // Check 2: Input folder access
  try {
    const folder = DriveApp.getFolderById(CONFIG.FOLDERS.INPUT_PDF_FOLDER_ID);
    results.checks.push({ name: 'Input Folder', status: 'PASS', message: folder.getName() });
  } catch (e) {
    results.checks.push({ name: 'Input Folder', status: 'FAIL', message: e.message });
  }

  // Check 3: Network connectivity
  try {
    UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
    results.checks.push({ name: 'Network', status: 'PASS', message: 'Connected' });
  } catch (e) {
    results.checks.push({ name: 'Network', status: 'FAIL', message: e.message });
  }

  results.overall = results.checks.every(c => c.status === 'PASS') ? 'HEALTHY' : 'DEGRADED';
  return results;
}

/**
 * Generates fallback dashboard HTML when template is missing
 * @returns {string} HTML string
 */
function generateFallbackDashboard() {
  const data = getDashboardData();

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Bank Statement Processor - Dashboard</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
      min-height: 100vh;
      color: #e4e6eb;
      padding: 20px;
    }
    .container { max-width: 1200px; margin: 0 auto; }
    h1 {
      font-size: 2.5rem;
      background: linear-gradient(90deg, #00d4ff, #7c4dff);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      margin-bottom: 30px;
      text-align: center;
    }
    .status-badge {
      display: inline-block;
      padding: 8px 16px;
      border-radius: 20px;
      font-weight: 600;
      text-transform: uppercase;
      font-size: 0.8rem;
    }
    .status-operational { background: #00c853; color: #fff; }
    .status-error { background: #ff5252; color: #fff; }
    .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px; margin-bottom: 30px; }
    .card {
      background: rgba(255,255,255,0.05);
      backdrop-filter: blur(10px);
      border-radius: 16px;
      padding: 24px;
      border: 1px solid rgba(255,255,255,0.1);
      transition: transform 0.3s, box-shadow 0.3s;
    }
    .card:hover {
      transform: translateY(-5px);
      box-shadow: 0 10px 40px rgba(0,212,255,0.2);
    }
    .card h3 { color: #00d4ff; margin-bottom: 15px; font-size: 1.1rem; }
    .card .value { font-size: 2rem; font-weight: 700; color: #fff; }
    .card .label { font-size: 0.85rem; color: #888; margin-top: 5px; }
    .income { color: #00e676 !important; }
    .expense { color: #ff5252 !important; }
    .links { margin-top: 30px; text-align: center; }
    .links a {
      display: inline-block;
      margin: 10px;
      padding: 12px 24px;
      background: linear-gradient(135deg, #7c4dff, #00d4ff);
      color: #fff;
      text-decoration: none;
      border-radius: 8px;
      font-weight: 600;
      transition: opacity 0.3s;
    }
    .links a:hover { opacity: 0.85; }
    footer { text-align: center; margin-top: 40px; color: #666; font-size: 0.85rem; }
  </style>
</head>
<body>
  <div class="container">
    <h1>🏦 Bank Statement Processor</h1>
    
    <div style="text-align: center; margin-bottom: 30px;">
      <span class="status-badge status-${data.status === 'operational' ? 'operational' : 'error'}">
        ${data.status === 'operational' ? '✅ Operational' : '❌ Error'}
      </span>
    </div>
    
    <div class="cards">
      <div class="card">
        <h3>📊 Total Transactions</h3>
        <div class="value">${data.transactions.total.toLocaleString()}</div>
        <div class="label">Processed statements</div>
      </div>
      
      <div class="card">
        <h3>💰 Total Income</h3>
        <div class="value income">$${data.transactions.income.toLocaleString(undefined, { minimumFractionDigits: 2 })}</div>
        <div class="label">Credits & deposits</div>
      </div>
      
      <div class="card">
        <h3>💸 Total Expenses</h3>
        <div class="value expense">$${data.transactions.expenses.toLocaleString(undefined, { minimumFractionDigits: 2 })}</div>
        <div class="label">Debits & withdrawals</div>
      </div>
      
      <div class="card">
        <h3>🕐 Last Updated</h3>
        <div class="value" style="font-size: 1.2rem;">${new Date(data.lastUpdated).toLocaleString()}</div>
        <div class="label">System timestamp</div>
      </div>
    </div>
    
    <div class="links">
      <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}" target="_blank">📊 Open Spreadsheet</a>
      <a href="https://github.com/traikdude/Bank_Statement_Processor" target="_blank">📂 GitHub Repository</a>
      <a href="https://colab.research.google.com/drive/15azpbyehCWpjAySUW7SfZDdWPujul00B" target="_blank">🐍 Colab Notebook</a>
    </div>
    
    <footer>
      <p>Script ID: ${SCRIPT_ID}</p>
      <p>© 2026 Bank Statement Processor - Powered by Google Apps Script</p>
    </footer>
  </div>
</body>
</html>
  `;
}

console.log('✅ Web App Dashboard handlers loaded!');