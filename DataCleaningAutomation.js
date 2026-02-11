// ===== CONFIGURATION =====
const RAW_SHEET_NAME = 'Raw Data';
const CLEANED_SHEET_NAME = 'Cleaned Data';

// Test order identification
const TEST_EMAILS = ['marketing@myco.pet', 'test@', 'demo@'];
const TEST_CUSTOMER_NAMES = ['Manthan Pandey', 'Test Customer'];

// ===== MAIN CLEANING FUNCTION =====
function cleanShopifyData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName(RAW_SHEET_NAME);
    const cleanedSheet = ss.getSheetByName(CLEANED_SHEET_NAME);
    
    if (!rawSheet || !cleanedSheet) {
      Logger.log('Error: Required sheets not found');
      return;
    }
    
    const rawData = rawSheet.getDataRange().getValues();
    
    if (rawData.length <= 1) {
      Logger.log('No data to clean');
      return;
    }
    
    // Headers
    const headers = rawData[0];
    
    // Find column indices
    const colIndices = {
      id: headers.indexOf('ID'),
      items: headers.indexOf('Items'),
      customerName: headers.indexOf('Customer Name'),
      shippingAddress: headers.indexOf('Shipping address'),
      orderTotal: headers.indexOf('Order total'),
      date: headers.indexOf('Date'),
      email: headers.indexOf('Email'),
      subtotal: headers.indexOf('Subtotal'),
      discount: headers.indexOf('Discount')
    };
    
    // New headers with split columns
    const newHeaders = [
      'Order ID',
      'Product Name',
      'Product Price',
      'Quantity',
      'Customer Name',
      'City',
      'Province',
      'Country',
      'Order Total',
      'Order Date',
      'Email',
      'Subtotal',
      'Discount Amount',
      'Discount Type'
    ];
    
    const cleanedData = [newHeaders];
    const seenOrderIds = new Set();
    let stats = {
      duplicates: 0,
      testOrders: 0,
      zeroOrders: 0,
      emptyRows: 0,
      processed: 0
    };
    
    // Track current order details for multi-line items
    let currentOrder = null;
    
    // Process each row
    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];
      
      // Check if this is a new order (has Order ID)
      const orderId = String(row[colIndices.id] || '').trim();
      const items = String(row[colIndices.items] || '').trim();
      const customerName = String(row[colIndices.customerName] || '').trim();
      const email = String(row[colIndices.email] || '').trim().toLowerCase();
      
      // Skip completely empty rows
      if (!orderId && !items && !customerName) {
        stats.emptyRows++;
        continue;
      }
      
      // If this row has an Order ID, it's a new order
      if (orderId) {
        // Check for test orders
        let isTestOrder = false;
        for (let testEmail of TEST_EMAILS) {
          if (email.includes(testEmail.toLowerCase())) {
            isTestOrder = true;
            break;
          }
        }
        for (let testName of TEST_CUSTOMER_NAMES) {
          if (customerName.toLowerCase().includes(testName.toLowerCase())) {
            isTestOrder = true;
            break;
          }
        }
        
        if (isTestOrder) {
          stats.testOrders++;
          currentOrder = null;
          continue;
        }
        
        // Check for $0 orders
        const totalStr = String(row[colIndices.orderTotal] || '').replace(/[^0-9.-]/g, '');
        const total = parseFloat(totalStr) || 0;
        
        if (total === 0) {
          stats.zeroOrders++;
          currentOrder = null;
          continue;
        }
        
        // Check for duplicates
        if (seenOrderIds.has(orderId)) {
          stats.duplicates++;
          currentOrder = null;
          continue;
        }
        seenOrderIds.add(orderId);
        
        // Store current order details
        currentOrder = {
          id: orderId,
          customerName: cleanCustomerName(customerName),
          address: parseAddress(String(row[colIndices.shippingAddress] || '')),
          total: cleanCurrency(String(row[colIndices.orderTotal] || '')),
          date: cleanDate(row[colIndices.date]),
          email: email,
          subtotal: cleanCurrency(String(row[colIndices.subtotal] || '')),
          discount: parseDiscount(String(row[colIndices.discount] || ''))
        };
      }
      
      // Process items (works for both first row and continuation rows)
      if (items && currentOrder) {
        const itemDetails = parseItem(items);
        
        const cleanedRow = [
          currentOrder.id,
          itemDetails.productName,
          itemDetails.price,
          itemDetails.quantity,
          currentOrder.customerName,
          currentOrder.address.city,
          currentOrder.address.province,
          currentOrder.address.country,
          currentOrder.total,
          currentOrder.date,
          currentOrder.email,
          currentOrder.subtotal,
          currentOrder.discount.amount,
          currentOrder.discount.type
        ];
        
        cleanedData.push(cleanedRow);
        stats.processed++;
      }
    }
    
    // Write to cleaned sheet
    cleanedSheet.clear();
    
    if (cleanedData.length > 0) {
      cleanedSheet.getRange(1, 1, cleanedData.length, cleanedData[0].length)
        .setValues(cleanedData);
      
      // Format header
      cleanedSheet.getRange(1, 1, 1, newHeaders.length)
        .setFontWeight('bold')
        .setBackground('#4CAF50')
        .setFontColor('#FFFFFF');
      
      cleanedSheet.setFrozenRows(1);
      
      // Auto-resize columns
      for (let i = 1; i <= newHeaders.length; i++) {
        cleanedSheet.autoResizeColumn(i);
      }
      
      const summary = `
✅ CLEANING COMPLETE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 Total rows processed: ${stats.processed}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🗑️ Removed:
   • Duplicates: ${stats.duplicates}
   • Test orders: ${stats.testOrders}
   • Zero-value: ${stats.zeroOrders}
   • Empty rows: ${stats.emptyRows}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
      `;
      
      Logger.log(summary);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Processed ${stats.processed} line items from ${seenOrderIds.size} orders`,
        '✅ Data Cleaned',
        5
      );
    }
    
  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error: ' + error.toString(),
      '❌ Cleaning Failed',
      10
    );
  }
}

// ===== HELPER FUNCTIONS =====

// Parse item string: "Joint & Mobility+ 54.99 1" → {productName, price, quantity}
function parseItem(itemStr) {
  // Remove extra spaces
  itemStr = itemStr.trim().replace(/\s+/g, ' ');
  
  // Split by space and work backwards
  const parts = itemStr.split(' ');
  
  // Last part is quantity
  const quantity = parts[parts.length - 1] || '1';
  
  // Second to last is price
  const price = parts[parts.length - 2] || '0.00';
  
  // Everything else is product name
  const productName = parts.slice(0, -2).join(' ');
  
  return {
    productName: productName || itemStr,
    price: price,
    quantity: quantity
  };
}

// Parse address: "Middleton Nova Scotia   Canada" → {city, province, country}
function parseAddress(addressStr) {
  // Clean multiple spaces
  addressStr = addressStr.trim().replace(/\s+/g, ' ');
  
  const parts = addressStr.split(' ');
  
  if (parts.length >= 3) {
    return {
      country: parts[parts.length - 1] || '',
      province: parts[parts.length - 2] || '',
      city: parts.slice(0, -2).join(' ') || ''
    };
  }
  
  return {
    city: addressStr,
    province: '',
    country: ''
  };
}

// Parse discount: "100.0 manual" → {amount, type}
function parseDiscount(discountStr) {
  discountStr = discountStr.trim();
  
  if (!discountStr) {
    return { amount: '', type: '' };
  }
  
  const parts = discountStr.split(' ');
  
  return {
    amount: parts[0] || '',
    type: parts.slice(1).join(' ') || ''
  };
}

// Clean customer name: "  Linda Powers" → "Linda Powers"
function cleanCustomerName(name) {
  return name.trim().replace(/\s+/g, ' ');
}

// Clean currency: "CAD86.17" → "86.17" or keep as is
function cleanCurrency(currencyStr) {
  return currencyStr.trim();
}

// Clean date
function cleanDate(dateValue) {
  if (!dateValue) return '';
  
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  
  if (typeof dateValue === 'string') {
    // Handle ISO format: "2026-02-06T12:34:04-05:00"
    try {
      const date = new Date(dateValue);
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    } catch (e) {
      return dateValue;
    }
  }
  
  return String(dateValue);
}

// ===== SCHEDULED TRIGGER =====
function createHourlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  ScriptApp.newTrigger('cleanShopifyData')
    .timeBased()
    .everyHours(1)
    .create();
  
  Logger.log('Hourly trigger created');
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Auto-cleaning enabled (runs every hour)',
    '✅ Trigger Created',
    5
  );
}

// ===== MANUAL RUN =====
function runCleaningNow() {
  cleanShopifyData();
}
