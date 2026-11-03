/**
 * ProyojonBD — Apps Script Backend (v6)
 * শীট ও হেডার অটো তৈরি হবে। আগের শীটে নতুন কলাম missing থাকলে অটো যুক্ত হবে।
 */

var SHEET_SCHEMAS = {
  'Products': ['id','name','category','unit','desc','price','discount','stock','image','status'],
  'Orders':   ['orderId','timestamp','customerName','phone','productName','category','deliveryArea','address','quantity','unit','color','size','unitPrice','subtotal','deliveryCharge','total','status'],
  'Expenses': ['expenseId','timestamp','category','description','amount'],
  'About':    ['key','value','updatedAt'],
  'Settings': ['key','value','updatedAt'],
  'Items':    ['itemId','name','brand','createdAt'],
  'Purchases':['invoiceNo','date','company','mobile','address','itemsJson','total','discount','payable','paid','due','authority','buyer','createdAt'],
  'Payments': ['paymentId','invoiceNo','date','amount','note','createdAt'],
  'Admins':   ['adminId','name','designation','mobile','email','passwordHash','salt','createdAt']
};

var DEFAULT_SETTINGS = { deliveryInside: 60, deliveryOutside: 150 };
var PRODUCTS_CACHE_KEY = 'products_v2';
var SETTINGS_CACHE_KEY = 'settings_v2';
var TZ = 'Asia/Dhaka';

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
  if (action === 'items')     return jsonResponse(getItems());
  if (action === 'purchases') return jsonResponse(getPurchases());
  if (action === 'stock')     return jsonResponse(getStock());
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
    case 'addItem':           return jsonResponse(addItem(body.item));
    case 'updateItem':        return jsonResponse(updateItem(body.item));
    case 'deleteItem':        return jsonResponse(deleteItem(body.itemId));
    case 'savePurchase':      return jsonResponse(savePurchase(body.purchase));
    case 'payDue':            return jsonResponse(payDue(body.invoiceNo, body.amount, body.date, body.note));
    case 'listAdmins':        return jsonResponse(listAdmins());
    case 'addAdmin':          return jsonResponse(addAdmin(body.admin));
    case 'updateAdmin':       return jsonResponse(updateAdmin(body.admin));
    case 'deleteAdmin':       return jsonResponse(deleteAdmin(body.adminId));
    case 'adminLogin':        return jsonResponse(adminLogin(body.identifier, body.password));
    default:                  return jsonResponse({ error: 'Unknown action' });
  }
}

/* ======== PRODUCTS ======== */

function normalizeUnit(u) { return String(u || '').trim().toUpperCase() === 'KG' ? 'KG' : 'PCS'; }

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
        unit:     normalizeUnit(o.unit),
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
    'unit':     normalizeUnit(p.unit),
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
    if (p.unit !== undefined) set('unit', normalizeUnit(p.unit));
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
        unit:           o.unit ? normalizeUnit(o.unit) : 'PCS',
        color:          o.color || '',
        size:           o.size  || '',
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
  var uom     = normalizeUnit(order.unit);
  var qty     = Number(order.quantity) || 0;
  qty = (uom === 'KG') ? r2(qty) : Math.round(qty);
  if (!(qty > 0)) return { error: 'পরিমাণ/ওজন সঠিক নয়' };
  var unit    = Number(order.unitPrice)      || 0;
  var sub     = r2(unit * qty);
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
    'unit':           uom,
    'color':          uom === 'PCS' ? String(order.color || '').trim().slice(0, 60) : '',
    'size':           uom === 'PCS' ? String(order.size  || '').trim().slice(0, 60) : '',
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

/* ======== HELPERS (নতুন) ======== */

function r2(n) { return Math.round((Number(n) || 0) * 100) / 100; }

function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '');
}

function todayStr() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

// নির্দিষ্ট কলাম টেক্সট ফরম্যাটে রেখে (০১৭… এর শূন্য ও তারিখ ঠিক রাখতে) নতুন row যোগ করে
function appendRowText(sheet, headers, map, textKeys) {
  var row = buildRow(headers, map);
  var r = sheet.getLastRow() + 1;
  headers.forEach(function(h, i) {
    if (textKeys.indexOf(h) >= 0) sheet.getRange(r, i + 1).setNumberFormat('@');
  });
  sheet.getRange(r, 1, 1, row.length).setValues([row]);
  return r;
}

function colIndex(sheet, name) {
  return getHeaders(sheet).indexOf(String(name).toLowerCase());
}

/* ======== SETTINGS (ডেলিভারি চার্জ, ওয়েব ইনফো, সোশ্যাল) ======== */

var NUM_SETTING_KEYS  = ['deliveryInside', 'deliveryOutside'];
var TEXT_SETTING_KEYS = ['webName','webAddress','webLogo','webMobile','webEmail',
                         'facebook','youtube','instagram','twitter','whatsapp'];

function clearSettingsCache() {
  try { CacheService.getScriptCache().remove(SETTINGS_CACHE_KEY); } catch (e) {}
}

function getSettings() {
  try {
    var cached = CacheService.getScriptCache().get(SETTINGS_CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch (e) {}

  var sheet = ensureSheet('Settings');
  var rows  = sheet.getDataRange().getValues();
  var out   = { deliveryInside: DEFAULT_SETTINGS.deliveryInside, deliveryOutside: DEFAULT_SETTINGS.deliveryOutside };
  for (var i = 1; i < rows.length; i++) {
    var k = String(rows[i][0]).trim();
    var v = rows[i][1];
    if (NUM_SETTING_KEYS.indexOf(k) >= 0) {
      if (v !== '' && !isNaN(Number(v))) out[k] = Number(v);
    } else if (TEXT_SETTING_KEYS.indexOf(k) >= 0) {
      out[k] = String(v === null || v === undefined ? '' : v);
    }
  }
  try { CacheService.getScriptCache().put(SETTINGS_CACHE_KEY, JSON.stringify(out), 300); } catch (e) {}
  return out;
}

// শুধু যে key গুলো পাঠানো হবে সেগুলোই আপডেট হবে
function saveSettings(s) {
  if (!s) return { error: 'settings data missing' };
  var sheet   = ensureSheet('Settings');
  var headers = getHeaders(sheet);
  var ts      = new Date();
  var rows    = sheet.getDataRange().getValues();
  var rowMap  = {};
  for (var i = 1; i < rows.length; i++) rowMap[String(rows[i][0]).trim()] = i + 1;

  function put(key, val, isText) {
    if (rowMap[key]) {
      var cell = sheet.getRange(rowMap[key], 2);
      cell.setNumberFormat(isText ? '@' : 'General');
      cell.setValue(val);
      sheet.getRange(rowMap[key], 3).setValue(ts);
    } else {
      appendRowText(sheet, headers, { 'key': key, 'value': val, 'updatedat': ts }, isText ? ['value'] : []);
    }
  }

  NUM_SETTING_KEYS.forEach(function(key) {
    if (s[key] === undefined || s[key] === '' || isNaN(Number(s[key])) || Number(s[key]) < 0) return;
    put(key, Number(s[key]), false);
  });
  TEXT_SETTING_KEYS.forEach(function(key) {
    if (s[key] === undefined || s[key] === null) return;
    put(key, String(s[key]).trim().slice(0, 500), true);
  });

  clearSettingsCache();
  return { success: true, settings: getSettings() };
}

/* ======== PRODUCT LIST (Items: Product Name + Brand) ======== */

function getItems() {
  var sheet = ensureSheet('Items');
  var rows  = sheet.getDataRange().getValues();
  return {
    items: rowsToObjects(rows).map(function(o) {
      return { itemId: String(o.itemid), name: String(o.name || ''), brand: String(o.brand || '') };
    })
  };
}

function addItem(it) {
  if (!it || !String(it.name || '').trim()) return { error: 'Product Name দিন' };
  var sheet   = ensureSheet('Items');
  var headers = getHeaders(sheet);
  var id      = 'ITM-' + new Date().getTime();
  appendRowText(sheet, headers, {
    'itemid': id, 'name': String(it.name).trim(), 'brand': String(it.brand || '').trim(), 'createdat': new Date()
  }, []);
  return { success: true, itemId: id };
}

function updateItem(it) {
  if (!it || !it.itemId) return { error: 'itemId দরকার' };
  if (!String(it.name || '').trim()) return { error: 'Product Name দিন' };
  var sheet = ensureSheet('Items');
  var rows  = sheet.getDataRange().getValues();
  var idC = colIndex(sheet, 'itemId'), nC = colIndex(sheet, 'name'), bC = colIndex(sheet, 'brand');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) !== String(it.itemId)) continue;
    sheet.getRange(i + 1, nC + 1).setValue(String(it.name).trim());
    sheet.getRange(i + 1, bC + 1).setValue(String(it.brand || '').trim());
    return { success: true };
  }
  return { error: 'পণ্য পাওয়া যায়নি' };
}

function deleteItem(itemId) {
  if (!itemId) return { error: 'itemId দরকার' };
  var sheet = ensureSheet('Items');
  var rows  = sheet.getDataRange().getValues();
  var idC = colIndex(sheet, 'itemId');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) === String(itemId)) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { error: 'পণ্য পাওয়া যায়নি' };
}

/* ======== PURCHASES ======== */

function purchaseFromObj(o) {
  var items = [];
  try { items = JSON.parse(o.itemsjson || '[]'); } catch (e) {}
  return {
    invoiceNo: String(o.invoiceno), date: fmtDate(o.date), company: String(o.company || ''),
    mobile: String(o.mobile || ''), address: String(o.address || ''), items: items,
    total: Number(o.total) || 0, discount: Number(o.discount) || 0, payable: Number(o.payable) || 0,
    paid: Number(o.paid) || 0, due: Number(o.due) || 0,
    authority: String(o.authority || ''), buyer: String(o.buyer || '')
  };
}

function getPurchases() {
  var sheet = ensureSheet('Purchases');
  var rows  = sheet.getDataRange().getValues();
  return { purchases: rowsToObjects(rows).map(purchaseFromObj).reverse() };
}

function savePurchase(p) {
  if (!p) return { error: 'তথ্য অসম্পূর্ণ' };
  var company = String(p.company || '').trim();
  if (!company) return { error: 'Company Name দিন' };

  var items = (p.items || []).map(function(it) {
    var qty = Number(it.qty) || 0, price = Number(it.price) || 0;
    return {
      name: String(it.name || '').trim(), brand: String(it.brand || '').trim(),
      colour: String(it.colour || '').trim(), size: String(it.size || '').trim(),
      price: r2(price), qty: qty, amount: r2(price * qty)
    };
  }).filter(function(it) { return it.name && it.qty > 0; });
  if (!items.length) return { error: 'কমপক্ষে একটি পণ্য যোগ করুন' };

  var total    = r2(items.reduce(function(a, it) { return a + it.amount; }, 0));
  var discount = r2(Math.min(Math.max(Number(p.discount) || 0, 0), total));
  var payable  = r2(total - discount);
  var paid     = r2(Math.max(Number(p.paid) || 0, 0));
  if (paid > payable) return { error: 'Paid, Payable এর চেয়ে বেশি হতে পারবে না' };
  var due = r2(payable - paid);

  var date = /^\d{4}-\d{2}-\d{2}$/.test(String(p.date || '')) ? String(p.date) : todayStr();
  var invoiceNo = 'PUR-' + new Date().getTime();

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet   = ensureSheet('Purchases');
    var headers = getHeaders(sheet);
    appendRowText(sheet, headers, {
      'invoiceno': invoiceNo, 'date': date, 'company': company,
      'mobile': String(p.mobile || '').trim(), 'address': String(p.address || '').trim(),
      'itemsjson': JSON.stringify(items),
      'total': total, 'discount': discount, 'payable': payable, 'paid': paid, 'due': due,
      'authority': String(p.authority || '').trim(), 'buyer': String(p.buyer || '').trim(),
      'createdat': new Date()
    }, ['date', 'mobile']);
  } finally { lock.releaseLock(); }

  return { success: true, purchase: {
    invoiceNo: invoiceNo, date: date, company: company,
    mobile: String(p.mobile || '').trim(), address: String(p.address || '').trim(),
    items: items, total: total, discount: discount, payable: payable, paid: paid, due: due,
    authority: String(p.authority || '').trim(), buyer: String(p.buyer || '').trim()
  } };
}

/* ======== DUE PAY ======== */

function payDue(invoiceNo, amount, date, note) {
  var amt = r2(amount);
  if (!invoiceNo) return { error: 'Invoice দরকার' };
  if (!(amt > 0)) return { error: 'সঠিক টাকার পরিমাণ দিন' };
  var pdate = /^\d{4}-\d{2}-\d{2}$/.test(String(date || '')) ? String(date) : todayStr();

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Purchases');
    var rows  = sheet.getDataRange().getValues();
    var iC = colIndex(sheet, 'invoiceNo'), paidC = colIndex(sheet, 'paid'),
        dueC = colIndex(sheet, 'due'), payC = colIndex(sheet, 'payable');
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][iC]) !== String(invoiceNo)) continue;
      var due = r2(Number(rows[i][dueC]) || 0);
      if (amt > due + 0.001) return { error: 'বাকি আছে মাত্র ' + due + ' টাকা' };
      var newPaid = r2((Number(rows[i][paidC]) || 0) + amt);
      var newDue  = r2((Number(rows[i][payC]) || 0) - newPaid);
      if (newDue < 0) newDue = 0;
      sheet.getRange(i + 1, paidC + 1).setValue(newPaid);
      sheet.getRange(i + 1, dueC + 1).setValue(newDue);

      var pSheet = ensureSheet('Payments');
      appendRowText(pSheet, getHeaders(pSheet), {
        'paymentid': 'PAY-' + new Date().getTime(), 'invoiceno': String(invoiceNo),
        'date': pdate, 'amount': amt, 'note': String(note || '').trim().slice(0, 200), 'createdat': new Date()
      }, ['date']);
      return { success: true, paid: newPaid, due: newDue };
    }
    return { error: 'Invoice পাওয়া যায়নি' };
  } finally { lock.releaseLock(); }
}

/* ======== STOCK (ক্রয় − বিক্রয়, পণ্যের নাম মিলিয়ে) ======== */

function getStock() {
  var pSheet = ensureSheet('Purchases');
  var purchases = rowsToObjects(pSheet.getDataRange().getValues()).map(purchaseFromObj);

  var oSheet = ensureSheet('Orders');
  var orders = rowsToObjects(oSheet.getDataRange().getValues());
  var soldByName = {};
  orders.forEach(function(o) {
    var st = String(o.status || 'Pending').trim().toLowerCase();
    if (st === 'cancel') return;
    var k = String(o.productname || '').trim().toLowerCase();
    if (!k) return;
    soldByName[k] = (soldByName[k] || 0) + (Number(o.quantity) || 0);
  });

  var details = {}, products = {};
  purchases.forEach(function(p) {
    (p.items || []).forEach(function(it) {
      var nk = it.name.toLowerCase();
      var dk = [nk, it.brand.toLowerCase(), (it.colour || '').toLowerCase(), (it.size || '').toLowerCase()].join('|');
      if (!details[dk]) details[dk] = { product: it.name, brand: it.brand, colour: it.colour || '', size: it.size || '', purchased: 0 };
      details[dk].purchased += Number(it.qty) || 0;

      if (!products[nk]) products[nk] = { product: it.name, brands: {}, purchased: 0 };
      products[nk].purchased += Number(it.qty) || 0;
      if (it.brand) products[nk].brands[it.brand] = true;
    });
  });

  var prodList = Object.keys(products).map(function(nk) {
    var x = products[nk], sold = soldByName[nk] || 0;
    return { product: x.product, brand: Object.keys(x.brands).join(', '),
             purchased: x.purchased, sold: sold, stock: x.purchased - sold };
  }).sort(function(a, b) { return a.product.localeCompare(b.product); });

  var detList = Object.keys(details).map(function(k) { return details[k]; })
    .sort(function(a, b) { return a.product.localeCompare(b.product); });

  return { products: prodList, details: detList };
}

/* ======== ADMINISTRATION (ব্যবস্থাপক) ======== */

function hashPassword(salt, pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function(b) { var h = (b < 0 ? b + 256 : b).toString(16); return h.length === 1 ? '0' + h : h; }).join('');
}

function normMobile(m) { return String(m || '').replace(/[\s\-]/g, ''); }

function adminPublic(o) {
  return { adminId: String(o.adminid), name: String(o.name || ''), designation: String(o.designation || ''),
           mobile: String(o.mobile || ''), email: String(o.email || '') };
}

function listAdmins() {
  var sheet = ensureSheet('Admins');
  return { admins: rowsToObjects(sheet.getDataRange().getValues()).map(adminPublic) };
}

function adminDuplicate(rows, mobile, email, exceptId) {
  return rows.some(function(o) {
    if (exceptId && String(o.adminid) === String(exceptId)) return false;
    if (mobile && normMobile(o.mobile) === mobile) return true;
    if (email && String(o.email || '').toLowerCase() === email) return true;
    return false;
  });
}

function addAdmin(a) {
  if (!a || !String(a.name || '').trim()) return { error: 'নাম দিন' };
  var mobile = normMobile(a.mobile), email = String(a.email || '').trim().toLowerCase();
  if (!mobile && !email) return { error: 'মোবাইল বা ইমেইল দিন (লগইনের জন্য)' };
  if (String(a.password || '').length < 6) return { error: 'পাসওয়ার্ড কমপক্ষে ৬ অক্ষরের হতে হবে' };

  var sheet = ensureSheet('Admins');
  var rows  = rowsToObjects(sheet.getDataRange().getValues());
  if (adminDuplicate(rows, mobile, email, null)) return { error: 'এই মোবাইল/ইমেইল আগে থেকেই আছে' };

  var salt = Utilities.getUuid();
  var id   = 'ADM-' + new Date().getTime();
  appendRowText(sheet, getHeaders(sheet), {
    'adminid': id, 'name': String(a.name).trim(), 'designation': String(a.designation || '').trim(),
    'mobile': mobile, 'email': email, 'passwordhash': hashPassword(salt, String(a.password)),
    'salt': salt, 'createdat': new Date()
  }, ['mobile']);
  return { success: true, adminId: id };
}

function updateAdmin(a) {
  if (!a || !a.adminId) return { error: 'adminId দরকার' };
  if (!String(a.name || '').trim()) return { error: 'নাম দিন' };
  var mobile = normMobile(a.mobile), email = String(a.email || '').trim().toLowerCase();
  if (!mobile && !email) return { error: 'মোবাইল বা ইমেইল দিন (লগইনের জন্য)' };
  if (a.password && String(a.password).length < 6) return { error: 'পাসওয়ার্ড কমপক্ষে ৬ অক্ষরের হতে হবে' };

  var sheet = ensureSheet('Admins');
  var data  = sheet.getDataRange().getValues();
  if (adminDuplicate(rowsToObjects(data), mobile, email, a.adminId)) return { error: 'এই মোবাইল/ইমেইল আগে থেকেই আছে' };

  var idC = colIndex(sheet, 'adminId');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idC]) !== String(a.adminId)) continue;
    var r = i + 1;
    function setCol(name, v, text) {
      var c = sheet.getRange(r, colIndex(sheet, name) + 1);
      if (text) c.setNumberFormat('@');
      c.setValue(v);
    }
    setCol('name', String(a.name).trim());
    setCol('designation', String(a.designation || '').trim());
    setCol('mobile', mobile, true);
    setCol('email', email);
    if (a.password) {
      var salt = Utilities.getUuid();
      setCol('salt', salt);
      setCol('passwordHash', hashPassword(salt, String(a.password)));
    }
    return { success: true };
  }
  return { error: 'ব্যবস্থাপক পাওয়া যায়নি' };
}

function deleteAdmin(adminId) {
  if (!adminId) return { error: 'adminId দরকার' };
  var sheet = ensureSheet('Admins');
  var rows  = sheet.getDataRange().getValues();
  var idC = colIndex(sheet, 'adminId');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) === String(adminId)) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { error: 'ব্যবস্থাপক পাওয়া যায়নি' };
}

// মোবাইল/ইমেইল + পাসওয়ার্ড যাচাই। ৫ বার ভুল হলে ১৫ মিনিট লক।
function adminLogin(identifier, password) {
  var ident = String(identifier || '').trim().toLowerCase();
  if (!ident || !password) return { error: 'মোবাইল/ইমেইল ও পাসওয়ার্ড দিন' };
  var cache = CacheService.getScriptCache();
  var key   = 'lf_' + ident.slice(0, 100);
  var fails = Number(cache.get(key)) || 0;
  if (fails >= 5) return { error: 'অনেকবার ভুল হয়েছে। ১৫ মিনিট পর আবার চেষ্টা করুন' };

  var sheet = ensureSheet('Admins');
  var rows  = rowsToObjects(sheet.getDataRange().getValues());
  var identMobile = normMobile(ident);
  for (var i = 0; i < rows.length; i++) {
    var o = rows[i];
    var match = (normMobile(o.mobile) && normMobile(o.mobile) === identMobile) ||
                (String(o.email || '').toLowerCase() === ident);
    if (!match) continue;
    if (hashPassword(String(o.salt), String(password)) === String(o.passwordhash)) {
      cache.remove(key);
      return { success: true, admin: { name: String(o.name || ''), designation: String(o.designation || '') } };
    }
  }
  cache.put(key, String(fails + 1), 900);
  return { error: 'মোবাইল/ইমেইল বা পাসওয়ার্ড ভুল' };
}
