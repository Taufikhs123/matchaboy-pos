/**
 * ============================================================
 * MATCHABOY POS — Google Apps Script Backend
 * ============================================================
 * This file handles all data storage in Google Sheets.
 * Deploy as Web App: Execute as Me, Access: Anyone.
 * All requests come as GET with ?payload=JSON to avoid CORS.
 * ============================================================
 */

// ─── BRANCH CONFIGURATION ───────────────────────────────────
// Add your branches here. Each branch has its own spreadsheet.
// To add a new branch, just add another entry below.
const BRANCHES = {
  "main": {
    password: "1234",
    spreadsheetId: "1I1REHGLXtcGLZo9ZucwSAXlmUZZxW9s3MAyWzszq7tc"
  }
  // Example: add more branches like this:
  // "branch2": {
  //   password: "5678",
  //   spreadsheetId: "ANOTHER_SPREADSHEET_ID"
  // }
};

// ─── SHEET DEFINITIONS ──────────────────────────────────────
// Each sheet tab with its column headers (created automatically)
const SHEET_DEFS = {
  Menu:        ["id","name","category","icon","image","price_cash","price_dana","price_gofood","price_shopee","price_grab","active"],
  Ingredients: ["id","name","unit","current_stock","min_stock","pack_size","pack_price","cost_per_unit"],
  Recipes:     ["menu_id","ingredient_id","ingredient_name","amount","unit"],
  Orders:      ["id","date","time","payment","total","branch","status","note","cancel_reason","cancelled_at"],
  OrderItems:  ["order_id","menu_id","menu_name","qty","price","payment_type"],
  Ledger:      ["id","date","time","type","category","description","amount","payment","order_id"],
  StockLog:    ["id","date","ingredient_id","ingredient_name","change","reason","order_id"]
};

// ─── MAIN ENTRY POINT ───────────────────────────────────────
// All requests come through doGet because we use GET to avoid CORS
function doGet(e) {
  try {
    // Parse the payload from the URL parameter
    var raw = e.parameter.payload;
    if (!raw) {
      return jsonResponse({ ok: false, error: "No payload parameter" });
    }
    var req = JSON.parse(raw);
    var action = req.action;
    var branch = req.branch;
    var password = req.password;
    var data = req.data || {};

    // ── Ping action (no auth needed) ──
    if (action === "ping") {
      return jsonResponse({ ok: true, message: "Connected" });
    }

    // ── Authenticate ──
    var branchConfig = BRANCHES[branch];
    if (!branchConfig) {
      return jsonResponse({ ok: false, error: "Unknown branch: " + branch });
    }
    if (branchConfig.password !== password) {
      return jsonResponse({ ok: false, error: "Invalid password" });
    }

    // Get the spreadsheet for this branch
    var ss = SpreadsheetApp.openById(branchConfig.spreadsheetId);

    // Make sure all sheet tabs exist
    ensureSheets(ss);

    // ── Route to the correct handler ──
    switch (action) {
      // --- Menu ---
      case "getMenu":         return jsonResponse(handleGetMenu(ss));
      case "saveMenu":        return jsonResponse(handleSaveMenu(ss, data));
      case "deleteMenu":      return jsonResponse(handleDeleteMenu(ss, data));

      // --- Ingredients ---
      case "getIngredients":  return jsonResponse(handleGetIngredients(ss));
      case "saveIngredient":  return jsonResponse(handleSaveIngredient(ss, data));
      case "deleteIngredient":return jsonResponse(handleDeleteIngredient(ss, data));
      case "restockIngredient":return jsonResponse(handleRestock(ss, data));

      // --- Recipes ---
      case "getRecipes":      return jsonResponse(handleGetRecipes(ss));
      case "saveRecipes":     return jsonResponse(handleSaveRecipes(ss, data));

      // --- Orders ---
      case "placeOrder":      return jsonResponse(handlePlaceOrder(ss, data, branch));
      case "getOrders":       return jsonResponse(handleGetOrders(ss));
      case "cancelOrder":     return jsonResponse(handleCancelOrder(ss, data));

      // --- Ledger ---
      case "getLedger":       return jsonResponse(handleGetLedger(ss));
      case "addLedgerEntry":  return jsonResponse(handleAddLedgerEntry(ss, data));

      // --- Stock Log ---
      case "getStockLog":     return jsonResponse(handleGetStockLog(ss));

      // --- Settings ---
      case "saveSettings":    return jsonResponse(handleSaveSettings(ss, data));

      // --- Sync All (load everything at once) ---
      case "syncAll":         return jsonResponse(handleSyncAll(ss));

      default:
        return jsonResponse({ ok: false, error: "Unknown action: " + action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ─── HELPERS ─────────────────────────────────────────────────

/** Return a JSON response that works with CORS */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Create all sheet tabs if they don't exist yet */
function ensureSheets(ss) {
  for (var name in SHEET_DEFS) {
    var sheet = ss.getSheetByName(name);
    var expected = SHEET_DEFS[name];
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
      sheet.getRange(1, 1, 1, expected.length).setFontWeight("bold");
    } else {
      // Check for missing columns and insert them at the correct position
      var current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
      for (var i = 0; i < expected.length; i++) {
        if (current.indexOf(expected[i]) === -1) {
          // Insert column at position i+1 and set header
          sheet.insertColumnBefore(i + 1);
          sheet.getRange(1, i + 1).setValue(expected[i]).setFontWeight("bold");
          // Refresh current array after insert
          current.splice(i, 0, expected[i]);
        }
      }
    }
  }
}

/** Convert a sheet's data to an array of objects using the header row as keys */
function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // Only headers or empty
  var headers = data[0];
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var val = data[i][j];
      // Google Sheets auto-converts date strings to Date objects — convert back to yyyy-MM-dd
      if (val instanceof Date) {
        var h = val.getHours() + val.getMinutes() + val.getSeconds();
        if (h === 0) {
          // Date only (no time component) → yyyy-MM-dd
          val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
        } else {
          // Date+time → yyyy-MM-dd HH:mm:ss
          val = Utilities.formatDate(val, tz, "yyyy-MM-dd HH:mm:ss");
        }
      }
      obj[headers[j]] = val;
    }
    result.push(obj);
  }
  return result;
}

/** Find a row index by matching a column value. Returns row number (1-based) or -1 */
function findRowByKey(sheet, keyCol, keyVal) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = headers.indexOf(keyCol);
  if (colIdx === -1) return -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIdx]) === String(keyVal)) {
      return i + 1; // 1-based row number
    }
  }
  return -1;
}

/** Insert or update a row. If a row with key matches, update it; otherwise append. */
function upsertRow(sheet, keyCol, obj) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
  if (headers[0] === "") {
    // Sheet has no headers yet — this shouldn't happen after ensureSheets
    return;
  }
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push(obj[headers[i]] !== undefined ? obj[headers[i]] : "");
  }

  var existingRow = findRowByKey(sheet, keyCol, obj[keyCol]);
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

/** Append a row from an object (always adds new, never updates) */
function appendRow(sheet, obj) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    row.push(obj[headers[i]] !== undefined ? obj[headers[i]] : "");
  }
  sheet.appendRow(row);
}

/** Delete a row by key column match */
function deleteRowByKey(sheet, keyCol, keyVal) {
  var rowNum = findRowByKey(sheet, keyCol, keyVal);
  if (rowNum > 0) {
    sheet.deleteRow(rowNum);
    return true;
  }
  return false;
}

/** Delete all rows matching a key value */
function deleteRowsByKey(sheet, keyCol, keyVal) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIdx = headers.indexOf(keyCol);
  if (colIdx === -1) return;
  // Delete from bottom to top to preserve row indices
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][colIdx]) === String(keyVal)) {
      sheet.deleteRow(i + 1);
    }
  }
}

/** Generate a unique ID like "ORD-20240315-001" */
function generateId(prefix) {
  var now = new Date();
  var date = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  var rand = Math.floor(Math.random() * 900) + 100;
  return prefix + "-" + date + "-" + rand;
}

// ─── MENU HANDLERS ──────────────────────────────────────────

function handleGetMenu(ss) {
  var sheet = ss.getSheetByName("Menu");
  return { ok: true, data: sheetToObjects(sheet) };
}

function handleSaveMenu(ss, data) {
  var sheet = ss.getSheetByName("Menu");
  if (!data.id) {
    data.id = generateId("MNU");
  }
  if (data.active === undefined) data.active = true;
  upsertRow(sheet, "id", data);
  return { ok: true, data: data };
}

function handleDeleteMenu(ss, data) {
  var sheet = ss.getSheetByName("Menu");
  deleteRowByKey(sheet, "id", data.id);
  // Also delete associated recipes
  var recipeSheet = ss.getSheetByName("Recipes");
  deleteRowsByKey(recipeSheet, "menu_id", data.id);
  return { ok: true };
}

// ─── INGREDIENT HANDLERS ────────────────────────────────────

function handleGetIngredients(ss) {
  var sheet = ss.getSheetByName("Ingredients");
  return { ok: true, data: sheetToObjects(sheet) };
}

function handleSaveIngredient(ss, data) {
  var sheet = ss.getSheetByName("Ingredients");
  if (!data.id) {
    data.id = generateId("ING");
  }
  // Auto-calculate cost per unit from pack size and price
  if (data.pack_size && data.pack_price && Number(data.pack_size) > 0) {
    data.cost_per_unit = Number(data.pack_price) / Number(data.pack_size);
  }
  upsertRow(sheet, "id", data);
  return { ok: true, data: data };
}

function handleDeleteIngredient(ss, data) {
  var sheet = ss.getSheetByName("Ingredients");
  deleteRowByKey(sheet, "id", data.id);
  // Also remove from recipes
  var recipeSheet = ss.getSheetByName("Recipes");
  deleteRowsByKey(recipeSheet, "ingredient_id", data.id);
  return { ok: true };
}

function handleRestock(ss, data) {
  var sheet = ss.getSheetByName("Ingredients");
  var rowNum = findRowByKey(sheet, "id", data.id);
  if (rowNum < 0) return { ok: false, error: "Ingredient not found" };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var stockIdx = headers.indexOf("current_stock");
  var currentStock = Number(row[stockIdx]) || 0;
  var addAmount = Number(data.amount) || 0;
  row[stockIdx] = currentStock + addAmount;
  sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);

  // Log the restock
  var logSheet = ss.getSheetByName("StockLog");
  appendRow(logSheet, {
    id: generateId("STK"),
    date: new Date().toISOString(),
    ingredient_id: data.id,
    ingredient_name: data.name || "",
    change: addAmount,
    reason: "Restock",
    order_id: ""
  });

  return { ok: true, newStock: currentStock + addAmount };
}

// ─── RECIPE HANDLERS ────────────────────────────────────────

function handleGetRecipes(ss) {
  var sheet = ss.getSheetByName("Recipes");
  return { ok: true, data: sheetToObjects(sheet) };
}

function handleSaveRecipes(ss, data) {
  var sheet = ss.getSheetByName("Recipes");
  // Delete existing recipes for this menu item, then re-insert all
  deleteRowsByKey(sheet, "menu_id", data.menu_id);
  var recipes = data.recipes || [];
  for (var i = 0; i < recipes.length; i++) {
    var r = recipes[i];
    r.menu_id = data.menu_id;
    appendRow(sheet, r);
  }
  return { ok: true };
}

// ─── ORDER HANDLERS ─────────────────────────────────────────

function handlePlaceOrder(ss, data, branch) {
  // Use the frontend-generated ID so local and cloud stay in sync
  var orderId = data.orderId || generateId("ORD");
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var dateStr = data.date || Utilities.formatDate(now, tz, "yyyy-MM-dd");
  var timeStr = data.time || Utilities.formatDate(now, tz, "HH:mm:ss");

  // 1. Save the order
  var orderSheet = ss.getSheetByName("Orders");
  appendRow(orderSheet, {
    id: orderId,
    date: dateStr,
    time: timeStr,
    payment: data.payment,
    total: data.total,
    branch: branch,
    status: "Active",
    note: data.note || "",
    cancel_reason: "",
    cancelled_at: ""
  });

  // 2. Save order items
  var itemSheet = ss.getSheetByName("OrderItems");
  var items = data.items || [];
  for (var i = 0; i < items.length; i++) {
    appendRow(itemSheet, {
      order_id: orderId,
      menu_id: items[i].menu_id,
      menu_name: items[i].menu_name,
      qty: items[i].qty,
      price: items[i].price,
      payment_type: data.payment
    });
  }

  // 3. Deduct ingredient stock based on recipes
  var recipeSheet = ss.getSheetByName("Recipes");
  var recipes = sheetToObjects(recipeSheet);
  var ingSheet = ss.getSheetByName("Ingredients");
  var ingredients = sheetToObjects(ingSheet);
  var logSheet = ss.getSheetByName("StockLog");

  for (var i = 0; i < items.length; i++) {
    var menuId = items[i].menu_id;
    var qty = Number(items[i].qty);
    // Find recipes for this menu item
    for (var r = 0; r < recipes.length; r++) {
      if (String(recipes[r].menu_id) === String(menuId)) {
        var ingId = recipes[r].ingredient_id;
        var usagePerOrder = Number(recipes[r].amount) * qty;
        // Deduct from ingredient
        var ingRowNum = findRowByKey(ingSheet, "id", ingId);
        if (ingRowNum > 0) {
          var headers = ingSheet.getRange(1, 1, 1, ingSheet.getLastColumn()).getValues()[0];
          var ingRow = ingSheet.getRange(ingRowNum, 1, 1, headers.length).getValues()[0];
          var stockIdx = headers.indexOf("current_stock");
          ingRow[stockIdx] = Number(ingRow[stockIdx]) - usagePerOrder;
          ingSheet.getRange(ingRowNum, 1, 1, ingRow.length).setValues([ingRow]);

          // Log stock change
          appendRow(logSheet, {
            id: generateId("STK"),
            date: now.toISOString(),
            ingredient_id: ingId,
            ingredient_name: recipes[r].ingredient_name,
            change: -usagePerOrder,
            reason: "Order " + orderId,
            order_id: orderId
          });
        }
      }
    }
  }

  // 4. Record income in Ledger
  var ledgerSheet = ss.getSheetByName("Ledger");
  appendRow(ledgerSheet, {
    id: generateId("LED"),
    date: dateStr,
    time: timeStr,
    type: "Income",
    category: "Sales",
    description: "Order " + orderId + " (" + data.payment + ")",
    amount: data.total,
    payment: data.payment,
    order_id: orderId
  });

  return { ok: true, orderId: orderId, date: dateStr, time: timeStr };
}

function handleGetOrders(ss) {
  var orderSheet = ss.getSheetByName("Orders");
  var itemSheet = ss.getSheetByName("OrderItems");
  return {
    ok: true,
    orders: sheetToObjects(orderSheet),
    orderItems: sheetToObjects(itemSheet)
  };
}

function handleCancelOrder(ss, data) {
  var orderSheet = ss.getSheetByName("Orders");
  var rowNum = findRowByKey(orderSheet, "id", data.id);
  if (rowNum < 0) return { ok: false, error: "Order not found" };

  var headers = orderSheet.getRange(1, 1, 1, orderSheet.getLastColumn()).getValues()[0];
  var row = orderSheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

  // Check if already cancelled
  var statusIdx = headers.indexOf("status");
  if (row[statusIdx] === "Cancelled") {
    return { ok: false, error: "Order already cancelled" };
  }

  var now = new Date();
  var tz = Session.getScriptTimeZone();

  // Mark as cancelled
  row[statusIdx] = "Cancelled";
  row[headers.indexOf("cancel_reason")] = data.reason || "Cancelled by cashier";
  row[headers.indexOf("cancelled_at")] = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");
  orderSheet.getRange(rowNum, 1, 1, row.length).setValues([row]);

  // Get order total and payment
  var total = Number(row[headers.indexOf("total")]);
  var payment = row[headers.indexOf("payment")];
  var dateStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  // Restore ingredient stock
  var itemSheet = ss.getSheetByName("OrderItems");
  var allItems = sheetToObjects(itemSheet);
  var recipeSheet = ss.getSheetByName("Recipes");
  var recipes = sheetToObjects(recipeSheet);
  var ingSheet = ss.getSheetByName("Ingredients");
  var logSheet = ss.getSheetByName("StockLog");

  var orderItems = allItems.filter(function(item) { return String(item.order_id) === String(data.id); });

  for (var i = 0; i < orderItems.length; i++) {
    var menuId = orderItems[i].menu_id;
    var qty = Number(orderItems[i].qty);
    for (var r = 0; r < recipes.length; r++) {
      if (String(recipes[r].menu_id) === String(menuId)) {
        var ingId = recipes[r].ingredient_id;
        var restoreAmount = Number(recipes[r].amount) * qty;
        var ingRowNum = findRowByKey(ingSheet, "id", ingId);
        if (ingRowNum > 0) {
          var ingHeaders = ingSheet.getRange(1, 1, 1, ingSheet.getLastColumn()).getValues()[0];
          var ingRow = ingSheet.getRange(ingRowNum, 1, 1, ingHeaders.length).getValues()[0];
          var stockIdx = ingHeaders.indexOf("current_stock");
          ingRow[stockIdx] = Number(ingRow[stockIdx]) + restoreAmount;
          ingSheet.getRange(ingRowNum, 1, 1, ingRow.length).setValues([ingRow]);

          appendRow(logSheet, {
            id: generateId("STK"),
            date: now.toISOString(),
            ingredient_id: ingId,
            ingredient_name: recipes[r].ingredient_name,
            change: restoreAmount,
            reason: "Cancel " + data.id,
            order_id: data.id
          });
        }
      }
    }
  }

  // Add reversal entry in Ledger
  var ledgerSheet = ss.getSheetByName("Ledger");
  var cancelTimeStr = Utilities.formatDate(now, tz, "HH:mm:ss");
  appendRow(ledgerSheet, {
    id: generateId("LED"),
    date: dateStr,
    time: cancelTimeStr,
    type: "Expense",
    category: "Sales Reversal",
    description: "Cancelled Order " + data.id,
    amount: total,
    payment: payment,
    order_id: data.id
  });

  return { ok: true };
}

// ─── LEDGER HANDLERS ────────────────────────────────────────

function handleGetLedger(ss) {
  var sheet = ss.getSheetByName("Ledger");
  return { ok: true, data: sheetToObjects(sheet) };
}

function handleAddLedgerEntry(ss, data) {
  var sheet = ss.getSheetByName("Ledger");
  if (!data.id) data.id = generateId("LED");
  appendRow(sheet, data);
  return { ok: true, data: data };
}

// ─── STOCK LOG HANDLER ──────────────────────────────────────

function handleGetStockLog(ss) {
  var sheet = ss.getSheetByName("StockLog");
  return { ok: true, data: sheetToObjects(sheet) };
}

// ─── SETTINGS HANDLER ──────────────────────────────────────

function handleSaveSettings(ss, data) {
  var sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.getRange(1, 1, 1, 2).setValues([["key", "value"]]).setFontWeight("bold");
  }
  // Save each key-value pair
  var keys = Object.keys(data);
  for (var k = 0; k < keys.length; k++) {
    var key = keys[k];
    var rowNum = -1;
    var allData = sheet.getDataRange().getValues();
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][0] === key) { rowNum = i + 1; break; }
    }
    if (rowNum > 0) {
      sheet.getRange(rowNum, 2).setValue(data[key]);
    } else {
      sheet.appendRow([key, data[key]]);
    }
  }
  return { ok: true };
}

// ─── SYNC ALL (bulk load) ───────────────────────────────────

function handleSyncAll(ss) {
  // Get settings from Settings sheet if it exists
  var settings = {};
  var settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet) {
    var sData = settingsSheet.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      if (sData[i][0]) settings[sData[i][0]] = sData[i][1];
    }
  }
  return {
    ok: true,
    settings: settings,
    menu: sheetToObjects(ss.getSheetByName("Menu")),
    ingredients: sheetToObjects(ss.getSheetByName("Ingredients")),
    recipes: sheetToObjects(ss.getSheetByName("Recipes")),
    orders: sheetToObjects(ss.getSheetByName("Orders")),
    orderItems: sheetToObjects(ss.getSheetByName("OrderItems")),
    ledger: sheetToObjects(ss.getSheetByName("Ledger")),
    stockLog: sheetToObjects(ss.getSheetByName("StockLog"))
  };
}
