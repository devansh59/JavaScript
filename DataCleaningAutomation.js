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
  'MPLM3000': 'Lean Mass+',
  'MPJD3000': 'Joint & Mobility+30gm',
  'MPFC3000': 'Focus & Calm 30gm',
  'MPPK3000': 'Puppy/Kitten',
  'MPNBO240': "Nature's BugOff"
};

// ===== ITEM NAME MAPPING (for items without product codes) =====
const ITEM_NAME_MAP = {
  // Gut & Immunity variations
  'Gut (Allergies) & Immunity +': 'Gut & Immunity+',
  'Gut (Allergies) & Immunity': 'Gut & Immunity+',
  'Gut & Immunity+': 'Gut & Immunity+',
  'Gut & Immunity +': 'Gut & Immunity+',
  'Gut & Immunity': 'Gut & Immunity+',
  'Allergies & Immunity': 'Gut & Immunity+',
  'Allergies & Immunity - Canine': 'Gut & Immunity+',
  'Aller-G & Immunity': 'Gut & Immunity+',
  'Gut & Immunity+ 500g': 'Gut & Immunity+100gm',
  'Gut & Immunity + 500g': 'Gut & Immunity+100gm',
  'Gut & Immunity - 120g': 'Gut & Immunity+100gm',
  'G&I 100g': 'Gut & Immunity+100gm',
  
  // Joint & Mobility variations
  'Joint & Mobility+': 'Joint & Mobility+',
  'Joint & Mobility +': 'Joint & Mobility+',
  'Joint & Mobililty+ (Recovery)': 'Joint & Mobility+',
  'Joints & Repair': 'Joint & Mobility+',
  'Joints & Recovery': 'Joint & Mobility+',
  'Joints & Recovery - Canine': 'Joint & Mobility+',
  'Joint & Detox (Recovery)': 'Joint & Mobility+',
  'Joint & Detox': 'Joint & Mobility+',
  'Joint & Mobility+ 500g': 'Joint & Mobility+500gm',
  'Joint & Mobility+ 600g': 'Joint & Mobility+600gm',
  
  // Protect variations
  'PROTECT': 'Protect',
  'PROTECT - Disinfectant': 'Protect',
  'Free PROTECT': 'Protect',
  
  // Lean Mass variations (CANINE ONLY - Equine excluded below)
  'Lean Mass+': 'Lean Mass+',
  'Lean Mass +': 'Lean Mass+',
  'Lean Mass': 'Lean Mass+',
  'Lean Mass - Canine': 'Lean Mass+',
  'Lean Mass+ 100g': 'Lean Mass+100gm',
  'Lean Mass - 100g': 'Lean Mass+100gm',
  
  // Focus & Calm variations (CANINE ONLY - Equine excluded below)
  'Focus & Calm': 'Focus & Calm',
  'Focus & Calm (replace GI)': 'Focus & Calm',
  'Calm & Focused': 'Focus & Calm',
  'Calming+': 'Focus & Calm',
  'Calming': 'Focus & Calm',
  'Calming - Canine': 'Focus & Calm',
  
  // Puppy/Kitten variations
  'Puppy/Kitten Formula': 'Puppy/Kitten',
  'Puppy Formula': 'Puppy/Kitten',
  'Puppy/Kitten 100g': 'Puppy/Kitten 100gm',
  
  // Nature's BugOff variations
  "Nature's BugOff": "Nature's BugOff",
  "Nature's BugOff 1 Gal": "Nature's BugOff 1 Gal",
  
  // Other products (non-core)
  'Myco Pet Gift Card': 'Gift Card',
  
  // ===== EXCLUDE: BLACK SOLDIER FLY PRODUCTS =====
  'Black Soldier Fly - Oil': 'EXCLUDE',
  'Black Soldier Fly Larvae - Powder Topper': 'EXCLUDE',
  'Black Soldier Fly Larvae (BSFL) - Dried': 'EXCLUDE',
  
  // ===== EXCLUDE: HEMP & PET BEDS =====
  'Hemp Animal Bedding': 'EXCLUDE',
  'Pet Beds': 'EXCLUDE',
  
  // ===== EXCLUDE: EQUINE PRODUCTS =====
  'Fore & Hind Gut': 'EXCLUDE',
  '3Nerve': 'EXCLUDE',
  'Nerve': 'EXCLUDE',
  'FOAL - Equine': 'EXCLUDE',
  'Lean Mass - Equine': 'EXCLUDE',
  'Power & Endurance - Equine': 'EXCLUDE',
  'Focus & Calm - Equine': 'EXCLUDE',
  'Joint & Mobility + (Recovery) - Equine': 'EXCLUDE',
  
  // ===== EXCLUDE: SAMPLES & MARKETING MATERIALS =====
  'Puppy/Kitten Samples': 'EXCLUDE',
  'Flyers + Magnets': 'EXCLUDE',
  'Flyers': 'EXCLUDE',
  'Magnets': 'EXCLUDE',
  'WBS Show': 'EXCLUDE',
  'Samples': 'EXCLUDE',
  'PK Samples - 2g': 'EXCLUDE',
  '2g samples - G&I': 'EXCLUDE',
  'G&I Samples': 'EXCLUDE',
  'Jars': 'EXCLUDE',
  'Labels': 'EXCLUDE',
  'Massager': 'EXCLUDE',
  'Stack': 'EXCLUDE',
  'Custom Capsules': 'EXCLUDE',
  'Test product': 'EXCLUDE',
  'FREE J&R 30g': 'EXCLUDE',
  
  // ===== EXCLUDE: FEES =====
  'CC Processing Fee': 'EXCLUDE',
  'CC Fee with wholesale discount applied': 'EXCLUDE',
  '4% CC Processing Fee': 'EXCLUDE',
  'CC Fee': 'EXCLUDE',
  'Shipping Weight for Additional Items': 'EXCLUDE'
};

// ===== MAIN CLEANING FUNCTION WITH INTELLIGENT MAPPING =====
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
    
    // New headers
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
      mappedByCode: 0,
      mappedByName: 0,
      excluded: 0,
      unmapped: 0,
      processed: 0
    };
    
    // Track unmapped items and excluded items for reporting
    const unmappedItems = new Set();
    const excludedItems = {};
    
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
      
      // Process items
      if (items && currentOrder) {
        
        const itemDetails = parseItem(items);
        let finalProductCode = '';
        let finalProductName = '';
        let shouldInclude = true;
        
        // ===== INTELLIGENT MAPPING STRATEGY =====
        
        // 1. Try product code first
        if (productCode && PRODUCT_CODE_MAP[productCode]) {
          finalProductCode = productCode;
          finalProductName = PRODUCT_CODE_MAP[productCode];
          stats.mappedByCode++;
        }
        // 2. Try item name mapping
        else if (ITEM_NAME_MAP[itemDetails.productName]) {
          const mappedName = ITEM_NAME_MAP[itemDetails.productName];
          
          // Check if this should be excluded
          if (mappedName === 'EXCLUDE') {
            stats.excluded++;
            shouldInclude = false;
            
            // Track excluded items
            if (!excludedItems[itemDetails.productName]) {
              excludedItems[itemDetails.productName] = 0;
            }
            excludedItems[itemDetails.productName]++;
          } else {
            finalProductCode = productCode || 'MAPPED_BY_NAME';
            finalProductName = mappedName;
            stats.mappedByName++;
          }
        }
        // 3. No mapping found - use original name
        else {
          finalProductCode = productCode || 'UNMAPPED';
          finalProductName = itemDetails.productName;
          stats.unmapped++;
          unmappedItems.add(itemDetails.productName);
        }
        
        // Only add if not excluded
        if (shouldInclude) {
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
   • Excluded items: ${stats.excluded}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📦 Products:
   • Mapped by code: ${stats.mappedByCode}
   • Mapped by name: ${stats.mappedByName}
   • Unmapped (kept original): ${stats.unmapped}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${unmappedItems.size > 0 ? '⚠️ ' + unmappedItems.size + ' unmapped items - check logs' : '✅ All items mapped!'}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
      `;
      
      Logger.log(summary);
      
      // Show excluded items
      if (Object.keys(excludedItems).length > 0) {
        Logger.log('\n🗑️ EXCLUDED ITEMS:');
        Logger.log('─────────────────────────────────────────');
        Object.keys(excludedItems).sort().forEach(item => {
          Logger.log(`  • ${item}: ${excludedItems[item]} orders`);
        });
        Logger.log('');
      }
      
      // Show unmapped items
      if (unmappedItems.size > 0) {
        Logger.log('\n⚠️ UNMAPPED ITEMS (using original names):');
        Logger.log('─────────────────────────────────────────');
        Array.from(unmappedItems).sort().forEach(item => {
          Logger.log(`  • ${item}`);
        });
        Logger.log('\nAdd these to ITEM_NAME_MAP if needed.\n');
      }
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Processed ${stats.processed} items. ${stats.unmapped} unmapped, ${stats.excluded} excluded.`,
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

// ===== DIAGNOSTIC FUNCTION =====
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
    
    // Parse item name
    const itemParts = items.split(' ');
    const itemName = itemParts.slice(0, -2).join(' ');
    
    if (productCode) {
      if (!productData[productCode]) {
        productData[productCode] = {
          itemNames: new Set(),
          count: 0,
          mapped: !!PRODUCT_CODE_MAP[productCode]
        };
      }
      productData[productCode].itemNames.add(itemName);
      productData[productCode].count++;
      
      if (!PRODUCT_CODE_MAP[productCode]) {
        unmappedCodes.add(productCode);
      }
    } else {
      itemsWithoutCodes.push(itemName);
    }
  }
  
  // ===== REPORT =====
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('📊 PRODUCT DATA ANALYSIS');
  Logger.log('═══════════════════════════════════════════════════════\n');
  
  // 1. Mapped Products
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
  
  // 2. Items without codes - categorized
  if (itemsWithoutCodes.length > 0) {
    const uniqueItems = [...new Set(itemsWithoutCodes)];
    
    // Categorize items
    const willBeIncluded = [];
    const willBeExcluded = [];
    const willBeUnmapped = [];
    
    uniqueItems.forEach(item => {
      const count = itemsWithoutCodes.filter(i => i === item).length;
      const itemInfo = `${item} - ${count} orders`;
      
      if (ITEM_NAME_MAP[item]) {
        if (ITEM_NAME_MAP[item] === 'EXCLUDE') {
          willBeExcluded.push(itemInfo);
        } else {
          willBeIncluded.push(itemInfo);
        }
      } else {
        willBeUnmapped.push(itemInfo);
      }
    });
    
    // Show included items
    if (willBeIncluded.length > 0) {
      Logger.log('✅ ITEMS WITHOUT CODES (Will be mapped):');
      Logger.log('─────────────────────────────────────────────────────');
      willBeIncluded.forEach(item => Logger.log(item));
      Logger.log('');
    }
    
    // Show excluded items
    if (willBeExcluded.length > 0) {
      Logger.log('🗑️ ITEMS WITHOUT CODES (Will be excluded):');
      Logger.log('─────────────────────────────────────────────────────');
      willBeExcluded.forEach(item => Logger.log(item));
      Logger.log('');
    }
    
    // Show unmapped items
    if (willBeUnmapped.length > 0) {
      Logger.log('⚠️ ITEMS WITHOUT CODES (Unmapped - will use original name):');
      Logger.log('─────────────────────────────────────────────────────');
      willBeUnmapped.forEach(item => Logger.log(item));
      Logger.log('');
    }
    
    Logger.log(`Total items without codes: ${itemsWithoutCodes.length}\n`);
  }
  
  // 3. Summary
  Logger.log('═══════════════════════════════════════════════════════');
  Logger.log('📈 SUMMARY:');
  Logger.log('─────────────────────────────────────────────────────');
  Logger.log(`Total unique product codes: ${Object.keys(productData).length}`);
  Logger.log(`Mapped codes: ${Object.keys(PRODUCT_CODE_MAP).length}`);
  Logger.log(`Unmapped codes: ${unmappedCodes.size}`);
  Logger.log(`Orders with mapped products: ${mappedCount}`);
  Logger.log(`Items without codes: ${itemsWithoutCodes.length}`);
  Logger.log('═══════════════════════════════════════════════════════\n');
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    `Analysis complete. Check logs for details.`,
    '📊 Analysis Complete',
    10
  );
}

// ===== HELPER FUNCTIONS =====

function parseItem(itemStr) {
  itemStr = itemStr.trim().replace(/\s+/g, ' ');
  const parts = itemStr.split(' ');
  const quantity = parts[parts.length - 1] || '1';
  const price = parts[parts.length - 2] || '0.00';
  const productName = parts.slice(0, -2).join(' ');
  
  return {
    productName: productName || itemStr,
    price: price,
    quantity: quantity
  };
}

function parseAddress(addressStr) {
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

function cleanCustomerName(name) {
  return name.trim().replace(/\s+/g, ' ');
}

function cleanCurrency(currencyStr) {
  return currencyStr.trim();
}

function cleanDate(dateValue) {
  if (!dateValue) return '';
  
  if (dateValue instanceof Date) {
    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  
  if (typeof dateValue === 'string') {
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
