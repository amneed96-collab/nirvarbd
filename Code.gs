/**
 * ProyojonBD — Apps Script Backend (v6)
 * শীট ও হেডার অটো তৈরি হবে। আগের শীটে নতুন কলাম missing থাকলে অটো যুক্ত হবে।
 */

var SHEET_SCHEMAS = {
  'Products': ['id','name','category','desc','price','discount','stock','image','status'],
  'Orders':   ['orderId','timestamp','customerName','phone','productName','category','deliveryArea','address','quantity','unitPrice','subtotal','deliveryCharge','total','status'],
  'Expenses': ['expenseId','timestamp','category','description','amount'],
  'About':    ['key','value','updatedAt'],
  'Settings': ['key','value','updatedAt']
};

var DEFAULT_SETTINGS = { deliveryInside: 60, deliveryOutside: 150 };
var PRODUCTS_CACHE_KEY = 'products_v1';

/* ======== SHEET AUTO-SETUP ======== */

function ensureSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  var schema = SHEET_SCHEMAS[name];

  if (!sheet) {
    // শীট নেই — নতুন তৈরি করা
    sheet = ss.insertSheet(name);
    sheet.appendRow(schema);
    styleHeader(sheet, schema.length);
    return sheet;
  }

  // শীট আছে — হেডার চেক করা
  var lastCol = sheet.getLastColumn();
  var existingHeaders = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim().toLowerCase(); })
    : [];

  if (existingHeaders.length === 0 || !existingHeaders[0]) {
    // হেডার একদম নেই
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, schema.length).setValues([schema]);
    styleHeader(sheet, schema.length);
    return sheet;
  }

  // Missing কলাম খুঁজে যুক্ত করা
  schema.forEach(function(col) {
    var lc = col.toLowerCase();
    if (existingHeaders.indexOf(lc) === -1) {
      var newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(col);
      sheet.getRange(1, newCol).setFontWeight('bold').setBackground('#1A1F36').setFontColor('#FFFFFF');
    }
  });

  return sheet;
}

function styleHeader(sheet, colCount) {
  sheet.getRange(1, 1, 1, colCount)
       .setFontWeight('bold')
       .setBackground('#1A1F36')
       .setFontColor('#FFFFFF');
}

function ensureAllSheets() {
  Object.keys(SHEET_SCHEMAS).forEach(function(name){ ensureSheet(name); });
}

/* ======== UTILITIES ======== */

function getHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){
    return String(h).trim().toLowerCase().replace(/\s+/g,'');
  });
}

function rowsToObjects(rows) {
  if (!rows || rows.length < 2) return [];
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,''); });
  var out = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0] && !row[1]) continue;
    var obj = { _row: i+1 };
    headers.forEach(function(h, idx){ obj[h] = row[idx]; });
    out.push(obj);
  }
  return out;
}

function buildRow(headers, map) {
  var row = new Array(headers.length).fill('');
  headers.forEach(function(h, i){
    if (map.hasOwnProperty(h)) row[i] = map[h];
  });
  return row;
}

/* ======== HTTP HANDLERS ======== */

function doGet(e) {
  // প্রতিটি হ্যান্ডলার নিজেই ensureSheet() কল করে — তাই এখানে সব শীট চেক করা হয় না (দ্রুত রেসপন্সের জন্য)
  var action = e.parameter.action;
  if (action === 'products') return jsonResponse(getProducts());
  if (action === 'settings') return jsonResponse({ settings: getSettings() });
  if (action === 'orders')   return jsonResponse(getOrders(e.parameter.status));
  if (action === 'expenses') return jsonResponse(getExpenses());
  if (action === 'about')    return jsonResponse(getAbout());
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ error: 'Invalid JSON' }); }
  switch (body.action) {
    case 'order':             return jsonResponse(saveOrder(body.order));
    case 'updateOrderStatus': return jsonResponse(updateOrderStatus(body.orderId, body.status));
    case 'addProduct':        return jsonResponse(addProduct(body.product));
    case 'updateProduct':     return jsonResponse(updateProduct(body.product));
    case 'addExpense':        return jsonResponse(addExpense(body.expense));
    case 'saveAbout':         return jsonResponse(saveAbout(body.about));
    case 'saveSettings':      return jsonResponse(saveSettings(body.settings));
    default:                  return jsonResponse({ error: 'Unknown action' });
  }
}

/* ======== PRODUCTS ======== */

function clearProductsCache() {
  try { CacheService.getScriptCache().remove(PRODUCTS_CACHE_KEY); } catch (e) {}
}

function getProducts() {
  // ক্যাশ থাকলে শীট না পড়েই সাথে সাথে রিটার্ন (৫ মিনিট; পণ্য যোগ/এডিট করলে ক্যাশ মুছে যায়)
  try {
    var cached = CacheService.getScriptCache().get(PRODUCTS_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  var sheet = ensureSheet('Products');
  var rows  = sheet.getDataRange().getValues();
  var result = {
    products: rowsToObjects(rows).map(function(o){
      return {
        id:       String(o.id),
        name:     o.name       || '',
        category: o.category   || 'Others',
        desc:     o.desc       || '',
        price:    Number(o.price)    || 0,
        discount: Number(o.discount) || 0,
        stock:    (o.stock===''||o.stock===undefined) ? null : Number(o.stock),
        image:    o.image      || '',
        status:   o.status ? String(o.status).trim() : 'Active'
      };
    })
  };
  try { CacheService.getScriptCache().put(PRODUCTS_CACHE_KEY, JSON.stringify(result), 300); } catch (e) {}
  return result;
}

function addProduct(p) {
  if (!p || !p.name || p.price===undefined) return { error: 'নাম ও দাম আবশ্যক' };
  var sheet   = ensureSheet('Products');
  var headers = getHeaders(sheet);
  var id      = 'P' + new Date().getTime();
  sheet.appendRow(buildRow(headers, {
    'id':       id,
    'name':     p.name,
    'category': p.category || 'Others',
    'desc':     p.desc     || '',
    'price':    Number(p.price)    || 0,
    'discount': Number(p.discount) || 0,
    'stock':    (p.stock===undefined||p.stock==='') ? '' : Number(p.stock),
    'image':    p.image    || '',
    'status':   'Active'
  }));
  clearProductsCache();
  return { success:true, id:id };
}

function updateProduct(p) {
  if (!p || !p.id) return { error: 'id প্রয়োজন' };
  var sheet   = ensureSheet('Products');
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var idCol   = headers.indexOf('id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) !== String(p.id)) continue;
    var r = i + 1;
    var set = function(col, val){
      if (val === undefined) return;
      var c = headers.indexOf(col);
      if (c >= 0) sheet.getRange(r, c+1).setValue(val);
    };
    set('name',     p.name);
    set('category', p.category);
    set('desc',     p.desc);
    if (p.price    !== undefined) set('price',    Number(p.price));
    if (p.discount !== undefined) set('discount', Number(p.discount));
    if (p.stock    !== undefined) set('stock',    p.stock==='' ? '' : Number(p.stock));
    set('image',  p.image);
    set('status', p.status);
    clearProductsCache();
    return { success:true };
  }
  return { error: 'পণ্য পাওয়া যায়নি' };
}

/* ======== ORDERS ======== */

function getOrders(status) {
  var sheet  = ensureSheet('Orders');
  var rows   = sheet.getDataRange().getValues();
  var orders = rowsToObjects(rows)
    .filter(function(o){
      if (!status) return true;
      return (o.status||'Pending').toString().trim().toLowerCase() === status.toLowerCase();
    })
    .map(function(o){
      return {
        orderId:        o.orderid,
        timestamp:      o.timestamp,
        customerName:   o.customername,
        phone:          o.phone,
        productName:    o.productname,
        category:       o.category || '',
        deliveryArea:   o.deliveryarea,
        address:        o.address,
        quantity:       o.quantity,
        unitPrice:      o.unitprice,
        subtotal:       o.subtotal,
        deliveryCharge: o.deliverycharge,
        total:          o.total,
        status:         (o.status||'Pending').toString().trim()
      };
    });
  return { orders: orders.reverse() };
}

function saveOrder(order) {
  if (!order || !order.productName || !order.customerName) return { error: 'তথ্য অসম্পূর্ণ' };
  var sheet   = ensureSheet('Orders');
  var headers = getHeaders(sheet);
  var orderId = 'PRY-' + new Date().getTime();
  var ts      = new Date();
  var qty     = Number(order.quantity)       || 1;
  var unit    = Number(order.unitPrice)      || 0;
  var sub     = Number(order.subtotal)       || (unit * qty);
  var area    = order.deliveryArea           || 'Inside Feni';
  // ডেলিভারি চার্জ সবসময় Settings শীট থেকে (ক্লায়েন্টের পাঠানো মান বিশ্বাস করা হয় না)
  var st      = getSettings();
  var dc      = (area === 'Outside Feni') ? st.deliveryOutside : st.deliveryInside;
  var total   = sub + dc;

  sheet.appendRow(buildRow(headers, {
    'orderid':        orderId,
    'timestamp':      ts,
    'customername':   order.customerName || '',
    'phone':          order.phone        || '',
    'productname':    order.productName  || '',
    'category':       order.category     || '',
    'deliveryarea':   area,
    'address':        order.address      || '',
    'quantity':       qty,
    'unitprice':      unit,
    'subtotal':       sub,
    'deliverycharge': dc,
    'total':          total,
    'status':         'Pending'
  }));

  return { success:true, orderId:orderId, timestamp:ts.toISOString(),
           subtotal:sub, deliveryArea:area, deliveryCharge:dc, total:total };
}

function updateOrderStatus(orderId, status) {
  if (!orderId || !status) return { error: 'orderId ও status দরকার' };
  var sheet   = ensureSheet('Orders');
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,''); });
  var idCol   = headers.indexOf('orderid');
  var stCol   = headers.indexOf('status');
  if (idCol<0||stCol<0) return { error: 'কলাম পাওয়া যায়নি' };
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(orderId)) {
      sheet.getRange(i+1, stCol+1).setValue(status);
      return { success:true };
    }
  }
  return { error: 'অর্ডার পাওয়া যায়নি' };
}

/* ======== EXPENSES ======== */

function getExpenses() {
  var sheet = ensureSheet('Expenses');
  var rows  = sheet.getDataRange().getValues();
  return {
    expenses: rowsToObjects(rows).map(function(o){
      return { expenseId:o.expenseid, timestamp:o.timestamp,
               category:o.category||'', description:o.description||'', amount:Number(o.amount)||0 };
    }).reverse()
  };
}

function addExpense(exp) {
  if (!exp || exp.amount===undefined) return { error: 'amount দরকার' };
  var sheet   = ensureSheet('Expenses');
  var headers = getHeaders(sheet);
  var id      = 'EXP-' + new Date().getTime();
  sheet.appendRow(buildRow(headers, {
    'expenseid':   id,
    'timestamp':   new Date(),
    'category':    exp.category    || 'অন্যান্য',
    'description': exp.description || '',
    'amount':      Number(exp.amount) || 0
  }));
  return { success:true, expenseId:id };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ======== ABOUT ======== */

function getAbout() {
  var sheet = ensureSheet('About');
  var rows  = sheet.getDataRange().getValues();
  var data  = { title:'', desc:'', points:'', contact:'' };
  for (var i = 1; i < rows.length; i++) {
    var key = String(rows[i][0]).trim();
    var val = String(rows[i][1]);
    if (data.hasOwnProperty(key)) data[key] = val;
  }
  return { about: data };
}

function saveAbout(about) {
  if (!about) return { error: 'about data missing' };
  var sheet   = ensureSheet('About');
  var headers = getHeaders(sheet);
  var keys    = ['title','desc','points','contact'];
  var ts      = new Date();

  // পুরোনো rows পড়া
  var rows = sheet.getDataRange().getValues();
  var existing = {};
  var rowMap   = {};
  for (var i = 1; i < rows.length; i++) {
    var k = String(rows[i][0]).trim();
    existing[k] = true;
    rowMap[k]   = i + 1;
  }

  keys.forEach(function(key) {
    var val = about[key] || '';
    if (existing[key]) {
      // আপডেট করা
      var r = rowMap[key];
      sheet.getRange(r, 2).setValue(val);
      sheet.getRange(r, 3).setValue(ts);
    } else {
      // নতুন row যুক্ত করা
      sheet.appendRow(buildRow(headers, { 'key': key, 'value': val, 'updatedat': ts }));
    }
  });

  return { success: true };
}

/* ======== SETTINGS (ডেলিভারি চার্জ ইত্যাদি) ======== */

function getSettings() {
  var sheet = ensureSheet('Settings');
  var rows  = sheet.getDataRange().getValues();
  var out   = { deliveryInside: DEFAULT_SETTINGS.deliveryInside, deliveryOutside: DEFAULT_SETTINGS.deliveryOutside };
  for (var i = 1; i < rows.length; i++) {
    var k = String(rows[i][0]).trim();
    var v = rows[i][1];
    if ((k === 'deliveryInside' || k === 'deliveryOutside') && v !== '' && !isNaN(Number(v))) {
      out[k] = Number(v);
    }
  }
  return out;
}

function saveSettings(s) {
  if (!s) return { error: 'settings data missing' };
  var sheet   = ensureSheet('Settings');
  var headers = getHeaders(sheet);
  var ts      = new Date();
  var rows    = sheet.getDataRange().getValues();
  var rowMap  = {};
  for (var i = 1; i < rows.length; i++) rowMap[String(rows[i][0]).trim()] = i + 1;

  ['deliveryInside', 'deliveryOutside'].forEach(function(key) {
    if (s[key] === undefined || s[key] === '' || isNaN(Number(s[key])) || Number(s[key]) < 0) return;
    var val = Number(s[key]);
    if (rowMap[key]) {
      sheet.getRange(rowMap[key], 2).setValue(val);
      sheet.getRange(rowMap[key], 3).setValue(ts);
    } else {
      sheet.appendRow(buildRow(headers, { 'key': key, 'value': val, 'updatedat': ts }));
    }
  });
  return { success: true, settings: getSettings() };
}
