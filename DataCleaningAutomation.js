// ===== CONFIGURATION =====
const RAW_SHEET_NAME = 'Raw Data';
const CLEANED_SHEET_NAME = 'Cleaned Data';

// Test order identification
const TEST_EMAILS = ['marketing@myco.pet', 'test@', 'demo@'];
const TEST_CUSTOMER_NAMES = ['Manthan Pandey', 'Test Customer'];

// ===== PRODUCT CODE MAPPING =====
const PRODUCT_CODE_MAP = {
  'MPGI3000': 'Gut & Immunity+30gm',
  'MPGI100': 'Gut & Immunity+100gm',
  'MPPR1200': 'Protect',
  'MPLM3000': 'Lean mass+',
  'MPJD3000': 'Joint & Mobility+30gm',
  'MPFC3000': 'Focus & calm 30gm',
  'MPPK3000': 'Puppy/Kitten',
  'MPNBO240': "Nature's BugOff"
  // Add more codes here after running analyzeProductData()
};

// ===== MAIN CLEANING FUNCTION WITH FALLBACK =====
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
      discount: headers.indexOf('Discount'),
      productCode: headers.indexOf('Product code')
    };
    
    // New headers with Product Code column
    const newHeaders = [
      'Order ID',
      'Product Code',
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
      mappedProducts: 0,
      unmappedProducts: 0,
      noCodeProducts: 0,
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
      const productCode = String(row[colIndices.productCode] || '').trim();
      
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
        
        let finalProductCode = '';
        let finalProductName = '';
        
        // ===== STRATEGY: Use product code if available, fallback to item name =====
        if (productCode && PRODUCT_CODE_MAP[productCode]) {
          // Product code exists and is mapped
          finalProductCode = productCode;
          finalProductName = PRODUCT_CODE_MAP[productCode];
          stats.mappedProducts++;
        } else if (productCode && !PRODUCT_CODE_MAP[productCode]) {
          // Product code exists but not mapped - use item name
          const itemDetails = parseItem(items);
          finalProductCode = productCode;
          finalProductName = itemDetails.productName;
          stats.unmappedProducts++;
        } else {
          // No product code - use item name
          const itemDetails = parseItem(items);
          finalProductCode = 'NO_CODE';
          finalProductName = itemDetails.productName;
          stats.noCodeProducts++;
        }
        
        const itemDetails = parseItem(items);
        
        const cleanedRow = [
          currentOrder.id,
          finalProductCode,
          finalProductName,
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
📦 Products:
   • Mapped (standardized): ${stats.mappedProducts}
   • Unmapped codes (using item name): ${stats.unmappedProducts}
   • No code (using item name): ${stats.noCodeProducts}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️ Run analyzeProductData() to see unmapped products
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
      `;
      
      Logger.log(summary);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Processed ${stats.processed} items. ${stats.unmappedProducts + stats.noCodeProducts} need mapping.`,
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

// ===== DIAGNOSTIC FUNCTION - RUN THIS TO SEE WHAT'S MISSING =====
function analyzeProductData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_SHEET_NAME);
  const rawData = rawSheet.getDataRange().getValues();
  const headers = rawData[0];
  
  const productCodeIndex = headers.indexOf('Product code');
  const itemsIndex = headers.indexOf('Items');
  
  if (productCodeIndex === -1 || itemsIndex === -1) {
    Logger.log('Required columns not found');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error: Required columns not found',
      '❌ Analysis Failed',
      5
    );
    return;
  }
  
  // Track all unique combinations
  const productData = {};
  const unmappedCodes = new Set();
  const itemsWithoutCodes = [];
  
  for (let i = 1; i < rawData.length; i++) {
    const productCode = String(rawData[i][productCodeIndex] || '').trim();
    const items = String(rawData[i][itemsIndex] || '').trim();
    
    if (!items) continue;
    
    // Parse item name (everything before last 2 parts which are price and quantity)
    const itemParts = items.split(' ');
    const itemName = itemParts.slice(0, -2).join(' ');
    
    if (productCode) {
      // Track code → item name mapping
      if (!productData[productCode]) {
        productData[productCode] = {
          itemNames: new Set(),
          count: 0,
          mapped: !!PRODUCT_CODE_MAP[productCode]
        };
      }
      productData[productCode].itemNames.add(itemName);
      productData[productCode].count++;
      
      // Track unmapped codes
      if (!PRODUCT_CODE_MAP[productCode]) {
        unmappedCodes.add(productCode);
      }
    } else {
      // Items without product codes
      itemsWithoutCodes.push(itemName);
    }
  }
  
  // ===== REPORT =====
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('📊 PRODUCT DATA ANALYSIS');
  Logger.log('═══════════════════════════════════════════════════════\n');
  
  // 1. Mapped Products (Currently Included)
  Logger.log('✅ MAPPED PRODUCTS (Standardized Names):');
  Logger.log('─────────────────────────────────────────────────────');
  let mappedCount = 0;
  Object.keys(productData).sort().forEach(code => {
    if (productData[code].mapped) {
      Logger.log(`${code} → ${PRODUCT_CODE_MAP[code]}`);
      Logger.log(`   Orders: ${productData[code].count}`);
      Logger.log(`   Item variations: ${Array.from(productData[code].itemNames).join(', ')}`);
      Logger.log('');
      mappedCount += productData[code].count;
    }
  });
  Logger.log(`Total mapped orders: ${mappedCount}\n`);
  
  // 2. Unmapped Products (Using Item Name as Fallback)
  if (unmappedCodes.size > 0) {
    Logger.log('⚠️ UNMAPPED PRODUCT CODES (Using Item Name):');
    Logger.log('─────────────────────────────────────────────────────');
    let unmappedCount = 0;
    Array.from(unmappedCodes).sort().forEach(code => {
      Logger.log(`${code} - NOT IN MAPPING`);
      Logger.log(`   Orders: ${productData[code].count}`);
      Logger.log(`   Item names: ${Array.from(productData[code].itemNames).join(', ')}`);
      Logger.log('');
      unmappedCount += productData[code].count;
    });
    Logger.log(`Total unmapped orders: ${unmappedCount}\n`);
  }
  
  // 3. Items without product codes
  if (itemsWithoutCodes.length > 0) {
    Logger.log('⚠️ ITEMS WITHOUT PRODUCT CODES:');
    Logger.log('─────────────────────────────────────────────────────');
    const uniqueItems = [...new Set(itemsWithoutCodes)];
    uniqueItems.forEach(item => {
      const count = itemsWithoutCodes.filter(i => i === item).length;
      Logger.log(`${item} - ${count} orders`);
    });
    Logger.log(`Total: ${itemsWithoutCodes.length}\n`);
  }
  
  // 4. Summary
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('📈 SUMMARY:');
  Logger.log('─────────────────────────────────────────────────────');
  Logger.log(`Total unique product codes: ${Object.keys(productData).length}`);
  Logger.log(`Mapped codes: ${Object.keys(PRODUCT_CODE_MAP).length}`);
  Logger.log(`Unmapped codes: ${unmappedCodes.size}`);
  Logger.log(`Orders with mapped products: ${mappedCount}`);
  Logger.log(`Orders with unmapped products: ${unmappedCodes.size > 0 ? Array.from(unmappedCodes).reduce((sum, code) => sum + productData[code].count, 0) : 0}`);
  Logger.log(`Items without codes: ${itemsWithoutCodes.length}`);
  Logger.log('═══════════════════════════════════════════════════════\n');
  
  // 5. Suggested additions to PRODUCT_CODE_MAP
  if (unmappedCodes.size > 0) {
    Logger.log('💡 SUGGESTED CODE TO ADD TO PRODUCT_CODE_MAP:');
    Logger.log('─────────────────────────────────────────────────────');
    Logger.log('Copy and paste this into your PRODUCT_CODE_MAP:\n');
    Array.from(unmappedCodes).sort().forEach(code => {
      const itemNames = Array.from(productData[code].itemNames);
      const suggestedName = itemNames[0]; // Use first item name as suggestion
      Logger.log(`  '${code}': '${suggestedName}',`);
    });
    Logger.log('\n');
  }
  
  // Show toast notification
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Found ${unmappedCodes.size} unmapped codes, ${itemsWithoutCodes.length} items without codes. Check logs.`,
    '📊 Analysis Complete',
    10
  );
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
