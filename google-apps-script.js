// =====================================================
// HOME KITCHEN ORDER - Google Apps Script
// =====================================================
// This script receives orders from your Home Kitchen app
// and writes them to this Google Sheet automatically.
//
// SETUP INSTRUCTIONS:
// 1. Open Google Sheets -> create a new spreadsheet
// 2. Name it something like "Home Kitchen Orders"
// 3. Click Extensions -> Apps Script
// 4. Delete any code in the editor
// 5. Paste this ENTIRE script
// 6. IMPORTANT: Set your secret token on line 36 below
//    (use the same token you set in the app's Admin settings)
// 7. Click the disk icon (Save), name it "Home Kitchen Script"
// 8. Click Deploy -> New deployment
// 9. Select type: "Web app"
// 10. Set "Execute as": Me
// 11. Set "Who has access": Anyone
// 12. Click Deploy -> Authorize access -> Allow
// 13. Copy the Web app URL
// 14. Paste that URL into your Home Kitchen app:
//     Admin -> Settings -> Google Apps Script URL
// 15. Set the SAME secret token in the app's Admin settings
//
// CUSTOMER WHITELIST:
// 16. Create a sheet named "Customers" in this same spreadsheet
// 17. Add column headers: "Phone" in A1, "PIN" in B1
// 18. List approved phone numbers below (just digits, e.g. 9876543210)
// 19. Leave PIN blank for new customers — default PIN is 1234
// 20. Only orders with a matching phone + PIN will be accepted
//
// That's it! Orders will now flow into your spreadsheet.
// =====================================================

// *** SET YOUR SECRET TOKEN HERE ***
// Pick any word or phrase. Must match what you enter in the app.
var SECRET_TOKEN = 'mango2026';
// Example: var SECRET_TOKEN = 'mango2026';

var DEFAULT_PIN = '1234';

// ============ PHONE WHITELIST ============

function getApprovedPhones(spreadsheet) {
  var custSheet = spreadsheet.getSheetByName('Customers');
  if (!custSheet) return null; // No Customers sheet = allow all (whitelist disabled)

  var data = custSheet.getDataRange().getValues();
  if (data.length < 2) return []; // Only header row, no phones = block all

  // Find the "Phone" column (case-insensitive)
  var headerRow = data[0];
  var phoneCol = -1;
  for (var c = 0; c < headerRow.length; c++) {
    if (String(headerRow[c]).trim().toLowerCase() === 'phone') {
      phoneCol = c;
      break;
    }
  }
  if (phoneCol === -1) return null; // No "Phone" column found = allow all

  // Collect all phone numbers, normalized to digits only
  var phones = [];
  for (var r = 1; r < data.length; r++) {
    var raw = String(data[r][phoneCol]).trim();
    var digits = raw.replace(/\D/g, '');
    if (digits.length >= 10) {
      // Store last 10 digits for matching (strips country code)
      phones.push(digits.slice(-10));
    }
  }
  return phones;
}

function isPhoneApproved(phone, approvedList) {
  if (approvedList === null) return true; // Whitelist disabled
  var digits = String(phone).replace(/\D/g, '');
  var last10 = digits.slice(-10);
  for (var i = 0; i < approvedList.length; i++) {
    if (approvedList[i] === last10) return true;
  }
  return false;
}

// ============ PIN HELPERS ============

function getCustomerRowAndPin(spreadsheet, phone) {
  var custSheet = spreadsheet.getSheetByName('Customers');
  if (!custSheet) return { row: -1, pin: null, pinCol: -1 };

  var data = custSheet.getDataRange().getValues();
  var headerRow = data[0];

  // Find Phone and PIN columns
  var phoneCol = -1;
  var pinCol = -1;
  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c]).trim().toLowerCase();
    if (h === 'phone') phoneCol = c;
    if (h === 'pin') pinCol = c;
  }

  if (phoneCol === -1) return { row: -1, pin: null, pinCol: -1 };

  // If no PIN column exists, create it
  if (pinCol === -1) {
    pinCol = headerRow.length;
    custSheet.getRange(1, pinCol + 1).setValue('PIN');
    custSheet.getRange(1, pinCol + 1).setFontWeight('bold');
  }

  var digits = String(phone).replace(/\D/g, '').slice(-10);

  for (var r = 1; r < data.length; r++) {
    var rowPhone = String(data[r][phoneCol]).replace(/\D/g, '');
    if (rowPhone.slice(-10) === digits) {
      var storedPin = String(data[r][pinCol] || '').trim();
      // Empty PIN means customer hasn't set one yet — use default
      if (!storedPin) storedPin = DEFAULT_PIN;
      return { row: r + 1, pin: storedPin, pinCol: pinCol + 1, sheet: custSheet };
    }
  }

  return { row: -1, pin: null, pinCol: pinCol + 1 };
}

function verifyPin(spreadsheet, phone, suppliedPin) {
  var info = getCustomerRowAndPin(spreadsheet, phone);
  if (info.row === -1) {
    // Phone not found — let the phone whitelist handle rejection
    return { valid: true, info: info };
  }
  var ok = String(suppliedPin).trim() === info.pin;
  return { valid: ok, info: info };
}

// ============ GET HANDLER ============
// action=menu  -> returns the current menu (public, token only)
// action=orders -> returns orders for a phone (requires phone+pin)
// (no action + phone param -> legacy orders request)

function doGet(e) {
  try {
    var token = e.parameter.token || '';
    var action = e.parameter.action || '';

    if (token !== SECRET_TOKEN) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'error', message: 'Invalid token' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet();

    // ---- FETCH MENU ----
    if (action === 'menu') {
      return doGetMenu(sheet);
    }

    // ---- FETCH SETTINGS ----
    if (action === 'settings') {
      return doGetSettings(sheet);
    }

    // ---- FETCH ALL ORDERS (admin view) ----
    if (action === 'admin-orders') {
      return doGetAdminOrders(sheet);
    }

    // ---- VERIFY PHONE + PIN (pre-validation before order/pin-change) ----
    if (action === 'verify') {
      var vPhone = e.parameter.phone || '';
      var vPin = e.parameter.pin || '';

      if (!vPhone || vPhone.replace(/\D/g, '').length < 10) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Invalid phone number' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // Check phone whitelist
      var approvedPhones = getApprovedPhones(sheet);
      if (!isPhoneApproved(vPhone, approvedPhones)) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Phone not registered. Please contact the vendor to register.' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // Check PIN
      var vPinResult = verifyPin(sheet, vPhone, vPin);
      if (!vPinResult.valid) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Incorrect PIN. Please check your PIN and try again.' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      return ContentService.createTextOutput(
        JSON.stringify({ status: 'success', message: 'Verified' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ---- FETCH ORDERS (default) ----
    var phone = e.parameter.phone || '';
    var pin = e.parameter.pin || '';

    if (!phone || phone.length < 10) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'error', message: 'Invalid phone number' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var pinResult = verifyPin(sheet, phone, pin);
    if (!pinResult.valid) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'error', message: 'Incorrect PIN' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return doGetOrders(sheet, phone);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGetMenu(spreadsheet) {
  var menuSheet = spreadsheet.getSheetByName('Menu');
  if (!menuSheet) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', menu: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var data = menuSheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', menu: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  // Headers: ID, Name, Price, Unit, Available, Category, Image URL
  var menuItems = [];
  for (var r = 1; r < data.length; r++) {
    var name = String(data[r][1] || '').trim();
    if (!name) continue;
    menuItems.push({
      id: Number(data[r][0]) || r,
      name: name,
      price: Number(data[r][2]) || 0,
      unit: String(data[r][3] || 'plate').trim(),
      available: Number(data[r][4]) || 0,
      category: String(data[r][5] || 'Others').trim(),
      imageUrl: String(data[r][6] || '').trim()
    });
  }

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'success', menu: menuItems })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doGetOrders(spreadsheet, phone) {
  var ordersSheet = spreadsheet.getSheetByName('Orders');
  if (!ordersSheet) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', orders: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var data = ordersSheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', orders: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var phoneLast10 = String(phone).replace(/\D/g, '').slice(-10);
  var matchingOrders = {};

  for (var r = 1; r < data.length; r++) {
    var rowPhone = String(data[r][4]).replace(/\D/g, '').slice(-10);
    if (rowPhone !== phoneLast10) continue;

    var orderId = data[r][0];
    if (!matchingOrders[orderId]) {
      matchingOrders[orderId] = {
        orderId: orderId,
        date: data[r][1],
        time: data[r][2],
        customerName: data[r][3],
        total: data[r][10],
        paymentStatus: data[r][11],
        notified: String(data[r][13] || '').toLowerCase() === 'yes',
        items: []
      };
    }
    matchingOrders[orderId].items.push({
      name: data[r][5],
      qty: data[r][6],
      unit: data[r][7],
      price: data[r][8],
      amount: data[r][9]
    });
  }

  var orderList = [];
  for (var id in matchingOrders) {
    orderList.push(matchingOrders[id]);
  }
  orderList.reverse();
  if (orderList.length > 20) orderList = orderList.slice(0, 20);

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'success', orders: orderList })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============ ADMIN ORDERS (all orders, no phone filter) ============

function doGetAdminOrders(spreadsheet) {
  var ordersSheet = spreadsheet.getSheetByName('Orders');
  if (!ordersSheet) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', orders: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var data = ordersSheet.getDataRange().getValues();
  if (data.length < 2) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', orders: [] })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var allOrders = {};
  for (var r = 1; r < data.length; r++) {
    var orderId = data[r][0];
    if (!orderId) continue;
    if (!allOrders[orderId]) {
      allOrders[orderId] = {
        id: orderId,
        customerName: String(data[r][3] || ''),
        customerPhone: String(data[r][4] || ''),
        total: Number(data[r][10]) || 0,
        paymentStatus: String(data[r][11] || 'unpaid'),
        notified: String(data[r][13] || '').toLowerCase() === 'yes',
        createdAt: String(data[r][1] || '') + ' ' + String(data[r][2] || ''),
        items: []
      };
    }
    allOrders[orderId].items.push({
      name: String(data[r][5] || ''),
      qty: Number(data[r][6]) || 0,
      unit: String(data[r][7] || ''),
      price: Number(data[r][8]) || 0,
      amount: Number(data[r][9]) || 0
    });
  }

  var orderList = [];
  for (var id in allOrders) {
    orderList.push(allOrders[id]);
  }
  orderList.reverse(); // newest first
  if (orderList.length > 100) orderList = orderList.slice(0, 100);

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'success', orders: orderList })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============ SETTINGS HELPERS ============

function saveSettingsToSheet(spreadsheet, settingsObj) {
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet('Settings');
  } else {
    settingsSheet.clear();
  }
  settingsSheet.appendRow(['Key', 'Value']);
  settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  settingsSheet.setFrozenRows(1);

  var keys = ['vendorName', 'vendorPhone', 'upiId', 'welcomeMessage', 'adminPin'];
  for (var i = 0; i < keys.length; i++) {
    var val = settingsObj[keys[i]];
    if (val !== undefined && val !== null) {
      settingsSheet.appendRow([keys[i], String(val).substring(0, 200)]);
    }
  }
  return true;
}

function doGetSettings(spreadsheet) {
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  if (!settingsSheet) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', settings: {} })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var data = settingsSheet.getDataRange().getValues();
  var result = {};
  for (var r = 1; r < data.length; r++) {
    var key = String(data[r][0] || '').trim();
    var val = String(data[r][1] || '').trim();
    if (key) result[key] = val;
  }

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'success', settings: result })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============ MENU SAVE HELPER ============

function saveMenuToSheet(spreadsheet, menuItems) {
  var menuSheet = spreadsheet.getSheetByName('Menu');

  if (!menuSheet) {
    menuSheet = spreadsheet.insertSheet('Menu');
  } else {
    menuSheet.clear();
  }

  // Write headers
  menuSheet.appendRow(['ID', 'Name', 'Price', 'Unit', 'Available', 'Category', 'Image URL']);
  menuSheet.getRange(1, 1, 1, 7).setFontWeight('bold');
  menuSheet.setFrozenRows(1);

  // Write each item
  for (var i = 0; i < menuItems.length; i++) {
    var item = menuItems[i];
    menuSheet.appendRow([
      Number(item.id) || (i + 1),
      String(item.name || '').substring(0, 100),
      Number(item.price) || 0,
      String(item.unit || 'plate').substring(0, 10),
      Number(item.available) || 0,
      String(item.category || 'Others').substring(0, 30),
      String(item.imageUrl || '').substring(0, 500)
    ]);
  }

  return true;
}

// ============ MAIN HANDLER (POST) ============

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // --- TOKEN CHECK ---
    if (!data.token || data.token !== SECRET_TOKEN) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'error', message: 'Invalid token' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet();

    // ============ PIN CHANGE REQUEST ============
    if (data.pinChange) {
      var pc = data.pinChange;
      var phone = String(pc.phone || '').trim();
      var oldPin = String(pc.oldPin || '').trim();
      var newPin = String(pc.newPin || '').trim();

      if (!phone || phone.length < 10) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Invalid phone number' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      if (!newPin || newPin.length < 4 || newPin.length > 8) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'New PIN must be 4-8 digits' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      var check = verifyPin(sheet, phone, oldPin);
      if (check.info.row === -1) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Phone not registered' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      if (!check.valid) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Current PIN is incorrect' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // Write new PIN
      check.info.sheet.getRange(check.info.row, check.info.pinCol).setValue(newPin);

      return ContentService.createTextOutput(
        JSON.stringify({ status: 'success', message: 'PIN changed successfully' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ============ SAVE MENU (admin pushes menu to sheet) ============
    if (data.saveMenu) {
      if (!data.saveMenu.items || !Array.isArray(data.saveMenu.items)) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Invalid menu data' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      if (data.saveMenu.items.length > 200) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Too many menu items' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      saveMenuToSheet(sheet, data.saveMenu.items);

      return ContentService.createTextOutput(
        JSON.stringify({ status: 'success', message: 'Menu saved (' + data.saveMenu.items.length + ' items)' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ============ SAVE SETTINGS (admin pushes settings to sheet) ============
    if (data.saveSettings) {
      saveSettingsToSheet(sheet, data.saveSettings);
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'success', message: 'Settings saved' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ============ NEW ORDERS ============
    if (data.orders) {
      // Basic validation: reject if too many rows (prevent flooding)
      if (data.orders.length > 50) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Too many items in one order' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // --- PHONE WHITELIST CHECK ---
      var approvedPhones = getApprovedPhones(sheet);
      var customerPhone = data.orders[0] ? data.orders[0].customerPhone : '';
      if (!isPhoneApproved(customerPhone, approvedPhones)) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Phone not registered. Please contact the vendor to register.' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      // --- PIN CHECK ---
      var suppliedPin = String(data.pin || '').trim();
      var pinResult = verifyPin(sheet, customerPhone, suppliedPin);
      if (!pinResult.valid) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Incorrect PIN. Please check your PIN and try again.' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      var ordersSheet = sheet.getSheetByName('Orders');

      // Create Orders sheet with headers if it doesn't exist
      if (!ordersSheet) {
        ordersSheet = sheet.insertSheet('Orders');
        ordersSheet.appendRow([
          'Order ID', 'Date', 'Time', 'Customer Name', 'Customer Phone',
          'Item', 'Qty', 'Unit', 'Price/Unit', 'Amount',
          'Order Total', 'Payment Status', 'Notes', 'Notified'
        ]);
        ordersSheet.getRange(1, 1, 1, 14).setFontWeight('bold');
        ordersSheet.setFrozenRows(1);
      } else {
        // Ensure Notified header exists in column 14 (for existing sheets)
        var hdrValues = ordersSheet.getRange(1, 1, 1, 14).getValues()[0];
        if (String(hdrValues[13] || '').trim().toLowerCase() !== 'notified') {
          ordersSheet.getRange(1, 14).setValue('Notified');
          ordersSheet.getRange(1, 14).setFontWeight('bold');
        }
      }

      // Add each item row (with basic sanitization)
      data.orders.forEach(function(row) {
        ordersSheet.appendRow([
          String(row.orderId || '').substring(0, 20),
          String(row.date || '').substring(0, 20),
          String(row.time || '').substring(0, 10),
          String(row.customerName || '').substring(0, 100),
          String(row.customerPhone || '').substring(0, 15),
          String(row.itemName || '').substring(0, 100),
          Number(row.quantity) || 0,
          String(row.unit || '').substring(0, 10),
          Number(row.pricePerUnit) || 0,
          Number(row.amount) || 0,
          Number(row.orderTotal) || 0,
          String(row.paymentStatus || 'unpaid').substring(0, 10),
          '',
          'no'
        ]);
      });

      // Also update the Summary sheet
      updateSummary(sheet, data.orders[0]);
    }

    // ============ PAYMENT CLAIM (customer says "I paid") ============
    if (data.paymentClaim) {
      var claim = data.paymentClaim;
      var claimOrderId = String(claim.orderId || '').substring(0, 20);
      var claimPhone = String(claim.phone || '').trim();
      var claimPin = String(claim.pin || '').trim();
      var claimNote = String(claim.note || '').substring(0, 200);

      // Verify PIN
      var claimPinResult = verifyPin(sheet, claimPhone, claimPin);
      if (!claimPinResult.valid) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Incorrect PIN' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      var ordersSheet = sheet.getSheetByName('Orders');
      if (ordersSheet) {
        var values = ordersSheet.getDataRange().getValues();
        var phoneLast10 = String(claimPhone).replace(/\D/g, '').slice(-10);

        // Ensure "Notes" header exists in column 13
        if (values[0].length < 13 || String(values[0][12]).trim().toLowerCase() !== 'notes') {
          ordersSheet.getRange(1, 13).setValue('Notes');
          ordersSheet.getRange(1, 13).setFontWeight('bold');
        }

        var found = false;
        for (var i = 1; i < values.length; i++) {
          var rowPhone = String(values[i][4]).replace(/\D/g, '').slice(-10);
          if (values[i][0] === claimOrderId && rowPhone === phoneLast10) {
            ordersSheet.getRange(i + 1, 12).setValue('claimed');
            if (claimNote) ordersSheet.getRange(i + 1, 13).setValue(claimNote);
            found = true;
          }
        }

        // Update summary too
        if (found) {
          var summarySheet = sheet.getSheetByName('Summary');
          if (summarySheet) {
            var sData = summarySheet.getDataRange().getValues();
            for (var i = 1; i < sData.length; i++) {
              if (sData[i][0] === claimOrderId) {
                summarySheet.getRange(i + 1, 6).setValue('claimed');
              }
            }
          }
        }
      }

      return ContentService.createTextOutput(
        JSON.stringify({ status: 'success', message: found ? 'Payment claim recorded' : 'Order not found' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ============ CANCEL ORDER ============
    if (data.cancelOrder) {
      var cancel = data.cancelOrder;
      var cancelOrderId = String(cancel.orderId || '').substring(0, 20);
      var cancelPhone = String(cancel.phone || '').trim();
      var cancelPin = String(cancel.pin || '').trim();

      // Verify PIN
      var cancelPinResult = verifyPin(sheet, cancelPhone, cancelPin);
      if (!cancelPinResult.valid) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'Incorrect PIN' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      var ordersSheet = sheet.getSheetByName('Orders');
      if (ordersSheet) {
        var values = ordersSheet.getDataRange().getValues();
        var phoneLast10 = String(cancelPhone).replace(/\D/g, '').slice(-10);

        var found = false;
        for (var i = 1; i < values.length; i++) {
          var rowPhone = String(values[i][4]).replace(/\D/g, '').slice(-10);
          var currentStatus = String(values[i][11]).toLowerCase();
          if (values[i][0] === cancelOrderId && rowPhone === phoneLast10) {
            // Only allow cancelling unpaid/claimed orders, not already paid
            if (currentStatus === 'paid') {
              return ContentService.createTextOutput(
                JSON.stringify({ status: 'error', message: 'Cannot cancel a paid order. Contact vendor.' })
              ).setMimeType(ContentService.MimeType.JSON);
            }
            ordersSheet.getRange(i + 1, 12).setValue('cancelled');
            found = true;
          }
        }

        if (found) {
          var summarySheet = sheet.getSheetByName('Summary');
          if (summarySheet) {
            var sData = summarySheet.getDataRange().getValues();
            for (var i = 1; i < sData.length; i++) {
              if (sData[i][0] === cancelOrderId) {
                summarySheet.getRange(i + 1, 6).setValue('cancelled');
              }
            }
          }
        }
      }

      return ContentService.createTextOutput(
        JSON.stringify({ status: 'success', message: found ? 'Order cancelled' : 'Order not found' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // ============ NOTIFY UPDATE (admin marks orders as notified) ============
    if (data.notifyUpdate) {
      var notifyIds = data.notifyUpdate.orderIds;
      if (notifyIds && Array.isArray(notifyIds) && notifyIds.length > 0) {
        var ordersSheet = sheet.getSheetByName('Orders');
        if (ordersSheet) {
          var values = ordersSheet.getDataRange().getValues();

          // Ensure Notified header exists in column 14
          if (values[0].length < 14 || String(values[0][13]).trim().toLowerCase() !== 'notified') {
            ordersSheet.getRange(1, 14).setValue('Notified');
            ordersSheet.getRange(1, 14).setFontWeight('bold');
          }

          // Build a set of order IDs for fast lookup
          var idSet = {};
          for (var n = 0; n < notifyIds.length; n++) {
            idSet[String(notifyIds[n]).substring(0, 20)] = true;
          }

          for (var i = 1; i < values.length; i++) {
            if (idSet[values[i][0]]) {
              ordersSheet.getRange(i + 1, 14).setValue('yes');
            }
          }

          // Update Summary sheet too
          var summarySheet = sheet.getSheetByName('Summary');
          if (summarySheet) {
            var sData = summarySheet.getDataRange().getValues();
            // Ensure Notified column in Summary (column 7)
            if (sData[0].length < 7 || String(sData[0][6] || '').trim().toLowerCase() !== 'notified') {
              summarySheet.getRange(1, 7).setValue('Notified');
              summarySheet.getRange(1, 7).setFontWeight('bold');
            }
            for (var i = 1; i < sData.length; i++) {
              if (idSet[sData[i][0]]) {
                summarySheet.getRange(i + 1, 7).setValue('yes');
              }
            }
          }
        }

        return ContentService.createTextOutput(
          JSON.stringify({ status: 'success', message: notifyIds.length + ' orders marked as notified' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Handle payment updates (admin)
    if (data.paymentUpdate) {
      var ordersSheet = sheet.getSheetByName('Orders');
      if (ordersSheet) {
        var dataRange = ordersSheet.getDataRange();
        var values = dataRange.getValues();
        var sanitizedId = String(data.paymentUpdate.orderId || '').substring(0, 20);
        var sanitizedStatus = String(data.paymentUpdate.paymentStatus || '').substring(0, 10);
        for (var i = 1; i < values.length; i++) {
          if (values[i][0] === sanitizedId) {
            ordersSheet.getRange(i + 1, 12).setValue(sanitizedStatus);
          }
        }
      }

      // Update summary too
      var summarySheet = sheet.getSheetByName('Summary');
      if (summarySheet) {
        var sData = summarySheet.getDataRange().getValues();
        var sanitizedId = String(data.paymentUpdate.orderId || '').substring(0, 20);
        var sanitizedStatus = String(data.paymentUpdate.paymentStatus || '').substring(0, 10);
        for (var i = 1; i < sData.length; i++) {
          if (sData[i][0] === sanitizedId) {
            summarySheet.getRange(i + 1, 6).setValue(sanitizedStatus);
          }
        }
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function updateSummary(spreadsheet, orderInfo) {
  var summarySheet = spreadsheet.getSheetByName('Summary');

  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet('Summary');
    summarySheet.appendRow([
      'Order ID', 'Date', 'Customer', 'Phone', 'Total', 'Payment'
    ]);
    summarySheet.getRange(1, 1, 1, 6).setFontWeight('bold');
    summarySheet.setFrozenRows(1);
  }

  // Check if this order already exists in summary
  var data = summarySheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === orderInfo.orderId) return;
  }

  summarySheet.appendRow([
    String(orderInfo.orderId || '').substring(0, 20),
    (orderInfo.date || '') + ' ' + (orderInfo.time || ''),
    String(orderInfo.customerName || '').substring(0, 100),
    String(orderInfo.customerPhone || '').substring(0, 15),
    Number(orderInfo.orderTotal) || 0,
    String(orderInfo.paymentStatus || 'unpaid').substring(0, 10)
  ]);
}

// Test function - run this to verify the sheet is set up correctly
function testSetup() {
  if (SECRET_TOKEN === 'CHANGE_ME') {
    Logger.log('WARNING: You need to change SECRET_TOKEN on line 36!');
    return;
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet();

  // Check Customers sheet
  var custSheet = sheet.getSheetByName('Customers');
  if (!custSheet) {
    Logger.log('NOTE: No "Customers" sheet found. Phone whitelist is DISABLED -- all orders accepted.');
    Logger.log('To enable: create a sheet named "Customers" with "Phone" and "PIN" column headers.');
  } else {
    var phones = getApprovedPhones(sheet);
    if (phones === null) {
      Logger.log('NOTE: Customers sheet exists but no "Phone" column found. Whitelist DISABLED.');
    } else {
      Logger.log('Phone whitelist ACTIVE. ' + phones.length + ' approved numbers found.');
    }

    // Check PIN column
    var data = custSheet.getDataRange().getValues();
    var hasPinCol = false;
    for (var c = 0; c < data[0].length; c++) {
      if (String(data[0][c]).trim().toLowerCase() === 'pin') { hasPinCol = true; break; }
    }
    if (hasPinCol) {
      Logger.log('PIN column found. PIN verification is ACTIVE.');
    } else {
      Logger.log('NOTE: No "PIN" column. It will be auto-created when the first order or PIN change arrives.');
      Logger.log('Default PIN for all customers: ' + DEFAULT_PIN);
    }
  }

  // Check Orders sheet
  var testSheet = sheet.getSheetByName('Orders');
  if (!testSheet) {
    testSheet = sheet.insertSheet('Orders');
    testSheet.appendRow([
      'Order ID', 'Date', 'Time', 'Customer Name', 'Customer Phone',
      'Item', 'Qty', 'Unit', 'Price/Unit', 'Amount',
      'Order Total', 'Payment Status', 'Notes', 'Notified'
    ]);
    testSheet.getRange(1, 1, 1, 14).setFontWeight('bold');
    testSheet.setFrozenRows(1);
  }
  Logger.log('Setup complete! Token is set and Orders sheet is ready.');
}
