// ===== UPDATED MAIN CLEANING FUNCTION WITH FALLBACK =====
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
      invalidProducts: 0,
      usedFallback: 0,
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
        
        let finalProductCode = productCode;
        let finalProductName = '';
        
        // ===== STRATEGY: Use product code if available, fallback to item name =====
        if (productCode && PRODUCT_CODE_MAP[productCode]) {
          // Product code exists and is mapped
          finalProductName = PRODUCT_CODE_MAP[productCode];
        } else {
          // No product code OR unmapped code - use item name as fallback
          const itemDetails = parseItem(items);
          finalProductCode = productCode || 'NO_CODE';
          finalProductName = itemDetails.productName;
          stats.usedFallback++;
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
✅ Mapped products: ${stats.processed - stats.usedFallback}
⚠️ Used item name fallback: ${stats.usedFallback}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
      `;
      
      Logger.log(summary);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Processed ${stats.processed} items (${stats.usedFallback} used fallback names)`,
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
