// ===== CONFIGURATION =====
const RAW_SHEET_NAME = 'Raw Data';
const CLEANED_SHEET_NAME = 'Cleaned Data';

// Test order identification (customize these)
const TEST_EMAILS = ['test@example.com', 'demo@myco.pet']; // Add test emails
const TEST_CUSTOMER_NAMES = ['Test Customer', 'Demo User']; // Add test names

// ===== MAIN CLEANING FUNCTION =====
function cleanShopifyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_SHEET_NAME);
  const cleanedSheet = ss.getSheetByName(CLEANED_SHEET_NAME);
  
  // Get all data from raw sheet
  const rawData = rawSheet.getDataRange().getValues();
  
  if (rawData.length <= 1) {
    Logger.log('No data to clean');
    return;
  }
  
  // Get headers (first row)
  const headers = rawData[0];
  
  // Find column indices
  const colIndices = {
    id: headers.indexOf('ID'),
    email: headers.indexOf('Email'),
    customerName: headers.indexOf('Customer Name'),
    total: headers.indexOf('Order total'),
    date: headers.indexOf('Date')
  };
  
  // Clean data array (start with headers)
  const cleanedData = [headers];
  const seenOrderIds = new Set();
  
  // Process each row (skip header)
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    
    // Skip if row is completely empty
    if (row.every(cell => cell === '' || cell === null)) {
      continue;
    }
    
    // CLEANING RULE 1: Remove duplicate orders
    const orderId = row[colIndices.id];
    if (seenOrderIds.has(orderId)) {
      Logger.log(`Skipping duplicate order: ${orderId}`);
      continue;
    }
    seenOrderIds.add(orderId);
    
    // CLEANING RULE 2: Remove test orders
    const email = row[colIndices.email];
    const customerName = row[colIndices.customerName];
    
    if (TEST_EMAILS.includes(email) || TEST_CUSTOMER_NAMES.includes(customerName)) {
      Logger.log(`Skipping test order: ${orderId}`);
      continue;
    }
    
    // CLEANING RULE 3: Remove $0 orders
    const total = parseFloat(row[colIndices.total]) || 0;
    if (total === 0) {
      Logger.log(`Skipping $0 order: ${orderId}`);
      continue;
    }
    
    // CLEANING RULE 4: Clean each cell in the row
    const cleanedRow = row.map((cell, index) => {
      // Trim whitespace from text
      if (typeof cell === 'string') {
        cell = cell.trim();
        
        // Proper case for customer names
        if (index === colIndices.customerName && cell) {
          cell = toProperCase(cell);
        }
      }
      
      // Format dates consistently
      if (index === colIndices.date && cell) {
        if (cell instanceof Date) {
          cell = Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        }
      }
      
      return cell;
    });
    
    cleanedData.push(cleanedRow);
  }
  
  // Clear cleaned sheet and write new data
  cleanedSheet.clear();
  
  if (cleanedData.length > 0) {
    cleanedSheet.getRange(1, 1, cleanedData.length, cleanedData[0].length)
      .setValues(cleanedData);
    
    // Format header row
    cleanedSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4CAF50')
      .setFontColor('#FFFFFF');
    
    // Freeze header row
    cleanedSheet.setFrozenRows(1);
    
    Logger.log(`Cleaned ${cleanedData.length - 1} orders (removed ${rawData.length - cleanedData.length} rows)`);
  }
}

// ===== HELPER FUNCTIONS =====
function toProperCase(str) {
  return str.toLowerCase().replace(/\b\w/g, char => char.toUpperCase());
}

// ===== AUTO-RUN ON EDIT (Optional) =====
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  
  // Only run if edit was in Raw Data sheet
  if (sheet.getName() === RAW_SHEET_NAME) {
    // Add small delay to let OSync finish writing
    Utilities.sleep(2000);
    cleanShopifyData();
  }
}

// ===== MANUAL RUN FUNCTION =====
function runCleaning() {
  cleanShopifyData();
}
