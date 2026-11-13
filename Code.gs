/**
 * NirvarBD — Apps Script Backend (v8)
 * শীট ও হেডার অটো তৈরি হবে। আগের শীটে নতুন কলাম missing থাকলে অটো যুক্ত হবে।
 */

var SHEET_SCHEMAS = {
  'Products': ['id','name','category','unit','desc','price','discount','stock','image','status','colors','sizes'],
  'Orders':   ['orderId','timestamp','customerName','phone','productName','category','deliveryArea','address','quantity','unit','color','size','unitPrice','subtotal','deliveryCharge','total','status','salesProcess','paymentMethod','paymentFrom','trxId'],
  'Slides': ['slideId','image','sort','createdAt'],
  'Expenses': ['expenseId','timestamp','category','description','amount','voucherNo','date'],
  'About':    ['key','value','updatedAt'],
  'Settings': ['key','value','updatedAt'],
  'Items':    ['itemId','name','brand','createdAt','unit','colors','sizes'],
  'Purchases':['invoiceNo','date','company','mobile','address','itemsJson','total','discount','payable','paid','due','authority','buyer','createdAt'],
  'Payments': ['paymentId','invoiceNo','date','amount','note','createdAt'],
  'Admins':   ['adminId','name','designation','mobile','email','passwordHash','salt','createdAt','password'],
  'Invests':  ['id','date','name','mobile','address','amount','createdAt'],
  'Deposits': ['id','date','description','amount','person','createdAt'],
  'Withdrawals': ['id','date','description','amount','person','createdAt'],
  'Incomes':  ['invoiceNo','date','itemsJson','total','createdAt']
};

var DEFAULT_SETTINGS = { deliveryInside: 60, deliveryOutside: 150 };
var PRODUCTS_CACHE_KEY = 'products_v3';
var SETTINGS_CACHE_KEY = 'settings_v3';
var DEFAULT_PRODUCT_CATS = ['Natural','Computer','Hardware','Mobile','Software','Website','Electronic','Cloth','Honey','Oil','Foods','Others'];
var DEFAULT_EXPENSE_CATS = ['পণ্য ক্রয়','ডেলিভারি খরচ','বিজ্ঞাপন','বেতন','অন্যান্য'];
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
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim().toLowerCase().replace(/\s+/g,''); })
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
    var lc = col.toLowerCase().replace(/\s+/g,'');
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

var PUBLIC_GET  = { products:1, slides:1, settings:1, about:1 };
var PUBLIC_POST = { order:1, adminLogin:1, listAdminNames:1, sendRecoveryPin:1, verifyRecoveryPin:1 };
var SESSION_TTL = 21600;   // ৬ ঘণ্টা (CacheService-এর সর্বোচ্চ)

function newSession(name) {
  var t = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  CacheService.getScriptCache().put('sess_' + t, JSON.stringify({ name: name || '' }), SESSION_TTL);
  return t;
}
function checkSession(t) {
  t = String(t || '');
  if (!/^[a-f0-9]{64}$/.test(t)) return false;
  return !!CacheService.getScriptCache().get('sess_' + t);
}

function doGet(e) {
  var action = e.parameter.action;
  if (!PUBLIC_GET[action] && !checkSession(e.parameter.token)) return jsonResponse({ error: 'UNAUTHORIZED' });
  // প্রতিটি হ্যান্ডলার নিজেই ensureSheet() কল করে — তাই এখানে সব শীট চেক করা হয় না (দ্রুত রেসপন্সের জন্য)
  if (action === 'products') return jsonResponse(getProducts());
  if (action === 'slides')   return jsonResponse(getSlides());
  if (action === 'settings') return jsonResponse({ settings: getSettings() });
  if (action === 'items')     return jsonResponse(getItems());
  if (action === 'purchases') return jsonResponse(getPurchases());
  if (action === 'stock')     return jsonResponse(getStock());
  if (action === 'orders')   return jsonResponse(getOrders(e.parameter.status));
  if (action === 'expenses') return jsonResponse(getExpenses());
  if (action === 'about')    return jsonResponse(getAbout());
  if (action === 'invests')     return jsonResponse(getInvests());
  if (action === 'deposits')    return jsonResponse(getDeposits());
  if (action === 'withdrawals') return jsonResponse(getWithdrawals());
  if (action === 'incomes')     return jsonResponse(getIncomes());
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch (err) { return jsonResponse({ error: 'Invalid JSON' }); }
  try { return handlePost(body); }
  catch (err) { return jsonResponse({ error: 'সার্ভার ত্রুটি: ' + (err && err.message ? err.message : err) }); }
}

function handlePost(body) {
  if (!PUBLIC_POST[body.action] && !checkSession(body.token)) return jsonResponse({ error: 'UNAUTHORIZED' });
  switch (body.action) {
    case 'order':             return jsonResponse(saveOrder(body.order, checkSession(body.token)));
    case 'adminLogout':       CacheService.getScriptCache().remove('sess_' + String(body.token || '')); return jsonResponse({ success: true });
    case 'updateOrderStatus': return jsonResponse(updateOrderStatus(body.orderId, body.status));
    case 'deleteOrder':       return jsonResponse(deleteOrder(body.orderId));
    case 'deleteCancelledOrders': return jsonResponse(deleteCancelledOrders());
    case 'addProduct':        return jsonResponse(addProduct(body.product));
    case 'updateProduct':     return jsonResponse(updateProduct(body.product));
    case 'deleteProduct':     return jsonResponse(deleteProduct(body.id));
    case 'addExpense':        return jsonResponse(addExpense(body.expense));
    case 'updateExpense':     return jsonResponse(updateExpense(body.expense));
    case 'deleteExpense':     return jsonResponse(deleteExpense(body.id));
    case 'saveAbout':         return jsonResponse(saveAbout(body.about));
    case 'saveSettings':      return jsonResponse(saveSettings(body.settings));
    case 'addItem':           return jsonResponse(addItem(body.item));
    case 'updateItem':        return jsonResponse(updateItem(body.item));
    case 'deleteItem':        return jsonResponse(deleteItem(body.itemId));
    case 'savePurchase':      return jsonResponse(savePurchase(body.purchase));
    case 'updatePurchase':    return jsonResponse(updatePurchase(body.invoiceNo, body.purchase));
    case 'deletePurchase':    return jsonResponse(deletePurchase(body.invoiceNo));
    case 'payDue':            return jsonResponse(payDue(body.invoiceNo, body.amount, body.date, body.note));
    case 'payAllDue':         return jsonResponse(payAllDue(body.date, body.note));
    case 'addSlide':          return jsonResponse(addSlide(body.image));
    case 'deleteSlide':       return jsonResponse(deleteSlide(body.slideId));
    case 'moveSlide':         return jsonResponse(moveSlide(body.slideId, body.dir));
    case 'listAdmins':        return jsonResponse(listAdmins());
    case 'addAdmin':          return jsonResponse(addAdmin(body.admin));
    case 'updateAdmin':       return jsonResponse(updateAdmin(body.admin));
    case 'deleteAdmin':       return jsonResponse(deleteAdmin(body.adminId));
    case 'adminLogin':        return jsonResponse(adminLogin(body.identifier, body.password));
    case 'listAdminNames':    return jsonResponse(listAdminNames());
    case 'sendRecoveryPin':   return jsonResponse(sendRecoveryPin(body.adminId, body.mobile));
    case 'verifyRecoveryPin': return jsonResponse(verifyRecoveryPin(body.adminId, body.pin));
    case 'addInvest':         return jsonResponse(addInvest(body.record));
    case 'updateInvest':      return jsonResponse(updateInvest(body.record));
    case 'deleteInvest':      return jsonResponse(deleteRecord('Invests', body.id));
    case 'addDeposit':        return jsonResponse(addCashRecord('Deposits', 'DEP-', body.record));
    case 'updateDeposit':     return jsonResponse(updateCashRecord('Deposits', body.record));
    case 'deleteDeposit':     return jsonResponse(deleteRecord('Deposits', body.id));
    case 'addWithdrawal':     return jsonResponse(addCashRecord('Withdrawals', 'WDR-', body.record));
    case 'updateWithdrawal':  return jsonResponse(updateCashRecord('Withdrawals', body.record));
    case 'deleteWithdrawal':  return jsonResponse(deleteRecord('Withdrawals', body.id));
    case 'addIncome':         return jsonResponse(addIncome(body.income));
    case 'updateIncome':      return jsonResponse(updateIncome(body.invoiceNo, body.income));
    case 'deleteIncome':      return jsonResponse(deleteIncome(body.invoiceNo));
    default:                  return jsonResponse({ error: 'Unknown action' });
  }
}

/* ======== PRODUCTS ======== */

var VALID_UNITS = ['PCS', 'KG', 'LITER', 'GRAM'];
function normalizeUnit(u) {
  var x = String(u || '').trim().toUpperCase();
  if (x === 'LITRE' || x === 'L') x = 'LITER';
  if (x === 'G' || x === 'GM') x = 'GRAM';
  return VALID_UNITS.indexOf(x) >= 0 ? x : 'PCS';
}

// "Red, Blue ,Black" → "Red, Blue, Black" (ফাঁকা/ডুপ্লিকেট বাদ)
function cleanList(v) {
  var seen = {}, out = [];
  String(v || '').split(/[,\n]/).forEach(function(s) {
    s = s.trim();
    var k = s.toLowerCase();
    if (s && !seen[k]) { seen[k] = true; out.push(s.slice(0, 40)); }
  });
  return out.join(', ').slice(0, 500);
}

// পণ্যের Color/Size: পণ্যে নিজে দেওয়া থাকলে সেটা, নাহলে Items (Products List) থেকে নামের মিলে
function productOpts(o, itemOpts) {
  if (normalizeUnit(o.unit) !== 'PCS') return { colors: '', sizes: '' };
  var fb = itemOpts[String(o.name || '').trim().toLowerCase()] || {};
  return { colors: cleanList(o.colors) || fb.colors || '', sizes: cleanList(o.sizes) || fb.sizes || '' };
}

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
  var itemOpts = {};
  rowsToObjects(ensureSheet('Items').getDataRange().getValues()).forEach(function(it) {
    itemOpts[String(it.name || '').trim().toLowerCase()] = { colors: cleanList(it.colors), sizes: cleanList(it.sizes) };
  });
  var stockMap = {};
  try { getStock().products.forEach(function(sp){ stockMap[String(sp.product).trim().toLowerCase()] = sp.stock; }); } catch (e) {}
  var result = {
    products: rowsToObjects(rows).map(function(o){
      var opt = productOpts(o, itemOpts);
      var sk = String(o.name || '').trim().toLowerCase();
      return {
        // স্টক ট্র্যাক করা পণ্য (Purchase আছে) এর স্টক ০ বা কম হলে Stock Out
        soldOut:  stockMap.hasOwnProperty(sk) && Number(stockMap[sk]) <= 0,
        id:       String(o.id),
        name:     o.name       || '',
        category: o.category   || 'Others',
        unit:     normalizeUnit(o.unit),
        desc:     o.desc       || '',
        price:    Number(o.price)    || 0,
        discount: Number(o.discount) || 0,
        stock:    (o.stock===''||o.stock===undefined) ? null : Number(o.stock),
        image:    o.image      || '',
        status:   o.status ? String(o.status).trim() : 'Active',
        colors:   opt.colors,
        sizes:    opt.sizes
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
    'status':   (String(p.status || '').trim() === 'Deactive') ? 'Deactive' : 'Active',
    'colors':   normalizeUnit(p.unit) === 'PCS' ? cleanList(p.colors) : '',
    'sizes':    normalizeUnit(p.unit) === 'PCS' ? cleanList(p.sizes)  : ''
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
    if (p.unit !== undefined && normalizeUnit(p.unit) !== 'PCS') { set('colors', ''); set('sizes', ''); }
    else {
      if (p.colors !== undefined) set('colors', cleanList(p.colors));
      if (p.sizes  !== undefined) set('sizes',  cleanList(p.sizes));
    }
    clearProductsCache();
    return { success:true };
  }
  return { error: 'পণ্য পাওয়া যায়নি' };
}

function deleteProduct(id) {
  if (!id) return { error: 'id প্রয়োজন' };
  var sheet = ensureSheet('Products');
  var rows  = sheet.getDataRange().getValues();
  var idC   = colIndex(sheet, 'id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) === String(id)) {
      sheet.deleteRow(i + 1);
      clearProductsCache();
      return { success: true };
    }
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
        phone:          fixPhone(o.phone),
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
        status:         (o.status||'Pending').toString().trim(),
        salesProcess:   (String(o.salesprocess || '').trim().toLowerCase() === 'offline') ? 'Offline' : 'Online',
        paymentMethod:  o.paymentmethod || '',
        paymentFrom:    fixPhone(o.paymentfrom),
        trxId:          o.trxid ? String(o.trxid) : ''
      };
    });
  return { orders: orders.reverse() };
}

// ছোট ক্রমিক আইডি: PREFIX-0001 (পুরোনো লম্বা আইডি হিসাবে ধরা হয় না)। LockService এর ভেতরে ডাকতে হবে
function nextSeqId(sheet, colName, prefix) {
  var c = colIndex(sheet, colName), max = 0;
  if (c >= 0 && sheet.getLastRow() > 1) {
    var re = new RegExp('^' + prefix + '-(\\d{1,6})$');
    sheet.getRange(2, c + 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
      var m = String(r[0] || '').match(re);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  var n = String(max + 1);
  while (n.length < 4) n = '0' + n;
  return prefix + '-' + n;
}

function saveOrder(order, isAdmin) {
  var offline = !!order && order.salesProcess === 'Offline';   // Sales in Shop — সরাসরি ডেলিভারি সম্পন্ন
  if (offline && !isAdmin) return { error: 'UNAUTHORIZED' };   // শপ বিক্রি শুধু লগইন করা অ্যাডমিন করতে পারবে
  if (offline && !String(order.customerName || '').trim()) order.customerName = 'Shop Customer';
  if (!order || !order.productName || !order.customerName) return { error: 'তথ্য অসম্পূর্ণ' };
  var sheet   = ensureSheet('Orders');
  var headers = getHeaders(sheet);
  var orderId = '';
  var ts      = new Date();
  var uom     = normalizeUnit(order.unit);
  var qty     = Number(order.quantity) || 0;
  qty = (uom === 'PCS') ? Math.round(qty) : r2(qty);
  if (!(qty > 0)) return { error: 'পরিমাণ/ওজন সঠিক নয়' };
  if (!offline) {
    var avail = null;
    try { getStock().products.forEach(function(sp){ if (String(sp.product).trim().toLowerCase() === String(order.productName).trim().toLowerCase()) avail = Number(sp.stock); }); } catch (e) {}
    if (avail !== null && avail <= 0) return { error: 'দুঃখিত, "' + order.productName + '" এর স্টক শেষ' };
    if (avail !== null && qty > avail) return { error: 'দুঃখিত, "' + order.productName + '" এর স্টকে মাত্র ' + avail + ' আছে' };
  }
  var unit    = Number(order.unitPrice)      || 0;
  if (!offline) {
    // অনলাইন অর্ডারে দাম সবসময় Products শীট থেকে (ক্লায়েন্টের পাঠানো দাম বিশ্বাস করা হয় না)
    var pf = null, pn = String(order.productName).trim().toLowerCase();
    getProducts().products.forEach(function(p) { if (String(p.name).trim().toLowerCase() === pn) pf = p; });
    if (!pf) return { error: 'পণ্য পাওয়া যায়নি, পেজ রিফ্রেশ করে আবার চেষ্টা করুন' };
    unit = (Number(pf.discount) > 0) ? Number(pf.price) - Number(pf.discount) : Number(pf.price);
    if (!(unit >= 0)) return { error: 'পণ্যের দাম সঠিক নয়' };
  }
  var sub     = r2(unit * qty);
  // ডেলিভারি চার্জ সবসময় Settings শীট থেকে (ক্লায়েন্টের পাঠানো মান বিশ্বাস করা হয় না)
  var st      = getSettings();
  var area    = offline ? 'In Shop' : String(order.deliveryArea || st.deliveryAreas[0].name);
  var dc      = 0;
  if (!offline) {
    var found = null;
    st.deliveryAreas.forEach(function(a) { if (a.name === area) found = a; });
    if (!found) return { error: 'ডেলিভারি এরিয়া সঠিক নয়, পেজ রিফ্রেশ করে আবার চেষ্টা করুন' };
    dc = found.charge;
  }
  if (offline) dc = 0;                     // শপ থেকে বিক্রিতে ডেলিভারি চার্জ নেই
  if (order.noDelivery === true) dc = 0;   // কার্টের একাধিক পণ্যে ডেলিভারি চার্জ একবারই
  var total   = sub + dc;

  // পেমেন্ট: ক্যাশ অন ডেলিভারি অথবা সেটিংসে থাকা ব্যাংক/মোবাইল ব্যাংকিং (ট্রানজেকশন আইডি আবশ্যক)
  var payMethod = 'Cash on Delivery', payFrom = '', trx = '';
  if (offline) {
    payMethod = 'Cash';
  } else if (order.paymentMethod && String(order.paymentMethod) !== 'Cash on Delivery') {
    var pmFound = null;
    st.paymentMethods.forEach(function(m) { if (m.name === String(order.paymentMethod)) pmFound = m; });
    if (!pmFound) return { error: 'পেমেন্ট মাধ্যম সঠিক নয়, পেজ রিফ্রেশ করে আবার চেষ্টা করুন' };
    trx     = String(order.trxId || '').trim().slice(0, 60);
    payFrom = String(order.paymentFrom || '').trim().slice(0, 60);
    if (!trx)     return { error: 'ট্রানজেকশন আইডি (TrxID) দিন' };
    if (!payFrom) return { error: 'যে নম্বর/অ্যাকাউন্ট থেকে টাকা পাঠিয়েছেন তা দিন' };
    payMethod = pmFound.name;
  }

  // মোবাইল নম্বর টেক্সট ফরম্যাটে (শুরুর ০ যাতে হারিয়ে না যায়)
  var oLock = LockService.getScriptLock();
  oLock.waitLock(15000);
  try {
  orderId = nextSeqId(sheet, 'orderId', 'PRY');
  appendRowText(sheet, headers, {
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
    'status':         offline ? 'Delivery' : 'New',   // অনলাইন: নতুন অর্ডার → কনফার্মে Pending (বুকিং); শপ: সরাসরি Delivery
    'salesprocess':   offline ? 'Offline' : 'Online',
    'paymentmethod':  payMethod,
    'paymentfrom':    payFrom,
    'trxid':          trx
  }, ['phone', 'paymentfrom', 'trxid']);
  } finally { oLock.releaseLock(); }

  clearProductsCache();
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
      clearProductsCache();
      return { success:true };
    }
  }
  return { error: 'অর্ডার পাওয়া যায়নি' };
}

// একটি অর্ডার একদম মুছে ফেলা (নতুন অর্ডার ডিলেট / বাতিল তালিকা থেকে ডিলেট)
function deleteOrder(orderId) {
  if (!orderId) return { error: 'orderId দরকার' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Orders');
    var rows  = sheet.getDataRange().getValues();
    var idC   = colIndex(sheet, 'orderId');
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][idC]) === String(orderId)) { sheet.deleteRow(i + 1); clearProductsCache(); return { success: true }; }
    }
    return { error: 'অর্ডার পাওয়া যায়নি' };
  } finally { lock.releaseLock(); }
}

// বাতিল (Cancel) তালিকার সব অর্ডার একসাথে মুছে ফেলা
function deleteCancelledOrders() {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Orders');
    var rows  = sheet.getDataRange().getValues();
    var stC   = colIndex(sheet, 'status');
    if (stC < 0) return { error: 'কলাম পাওয়া যায়নি' };
    var n = 0;
    for (var i = rows.length - 1; i >= 1; i--) {
      if (String(rows[i][stC] || '').trim().toLowerCase() === 'cancel') { sheet.deleteRow(i + 1); n++; }
    }
    clearProductsCache();
    return { success: true, deleted: n };
  } finally { lock.releaseLock(); }
}

/* ======== EXPENSES ======== */

function getExpenses() {
  var sheet = ensureSheet('Expenses');
  var rows  = sheet.getDataRange().getValues();
  return {
    expenses: rowsToObjects(rows).map(function(o){
      return { expenseId:o.expenseid, timestamp:o.timestamp,
               category:o.category||'', description:o.description||'', amount:Number(o.amount)||0,
               voucherNo:String(o.voucherno||''), date:o.date ? fmtDate(o.date) : '' };
    }).reverse()
  };
}

function updateExpense(exp) {
  if (!exp || !exp.expenseId) return { error: 'id দরকার' };
  var amt = r2(exp.amount);
  if (!(amt > 0)) return { error: 'সঠিক পরিমাণ দিন' };
  var sheet = ensureSheet('Expenses');
  var rows  = sheet.getDataRange().getValues();
  var idC   = colIndex(sheet, 'expenseId');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) !== String(exp.expenseId)) continue;
    var map = { 'category': exp.category || 'অন্যান্য', 'description': exp.description || '', 'amount': amt, 'date': validYmd(exp.date) };
    Object.keys(map).forEach(function(k) {
      var c = colIndex(sheet, k);
      if (c < 0) return;
      var cell = sheet.getRange(i + 1, c + 1);
      if (k === 'date') cell.setNumberFormat('@');
      cell.setValue(map[k]);
    });
    return { success: true };
  }
  return { error: 'রেকর্ড পাওয়া যায়নি' };
}

function deleteExpense(id) {
  if (!id) return { error: 'id দরকার' };
  var sheet = ensureSheet('Expenses');
  var rows  = sheet.getDataRange().getValues();
  var idC   = colIndex(sheet, 'expenseId');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) === String(id)) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { error: 'রেকর্ড পাওয়া যায়নি' };
}

function addExpense(exp) {
  if (!exp || exp.amount===undefined) return { error: 'amount দরকার' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet   = ensureSheet('Expenses');
    var headers = getHeaders(sheet);
    // অটো Voucher No: সর্বোচ্চ নম্বর + ১
    var vc = colIndex(sheet, 'voucherNo'), max = 0;
    if (vc >= 0 && sheet.getLastRow() > 1) {
      sheet.getRange(2, vc + 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
        var m = String(r[0] || '').match(/(\d+)\s*$/);
        if (m) max = Math.max(max, parseInt(m[1], 10));
      });
    }
    var voucher = 'V-' + ('0000' + (max + 1)).slice(-4);
    var id      = 'EXP-' + new Date().getTime();
    appendRowText(sheet, headers, {
      'expenseid':   id,
      'timestamp':   new Date(),
      'category':    exp.category    || 'অন্যান্য',
      'description': exp.description || '',
      'amount':      Number(exp.amount) || 0,
      'voucherno':   voucher,
      'date':        validYmd(exp.date)
    }, ['voucherno', 'date']);
    return { success:true, expenseId:id, voucherNo:voucher };
  } finally {
    lock.releaseLock();
  }
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
var TEXT_SETTING_KEYS = ['webName','webAddress','webLogo','webLogoEn','webMobile','webMobile2','webMobile3','webEmail',
                         'facebook','youtube','instagram','twitter','tiktok','whatsapp'];

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
  var areasRaw = null, catsRaw = null, expCatsRaw = null, payRaw = null;
  for (var i = 1; i < rows.length; i++) {
    var k = String(rows[i][0]).trim();
    var v = rows[i][1];
    if (k === 'deliveryAreas') {
      try { var pa = JSON.parse(String(v)); if (Array.isArray(pa)) areasRaw = pa; } catch (e) {}
    } else if (k === 'paymentMethods') {
      try { var pm = JSON.parse(String(v)); if (Array.isArray(pm)) payRaw = pm; } catch (e) {}
    } else if (k === 'categories') {
      try { var pc = JSON.parse(String(v)); if (Array.isArray(pc)) catsRaw = pc; } catch (e) {}
    } else if (k === 'expenseCategories') {
      try { var pe = JSON.parse(String(v)); if (Array.isArray(pe)) expCatsRaw = pe; } catch (e) {}
    } else if (NUM_SETTING_KEYS.indexOf(k) >= 0) {
      if (v !== '' && !isNaN(Number(v))) out[k] = Number(v);
    } else if (TEXT_SETTING_KEYS.indexOf(k) >= 0) {
      out[k] = String(v === null || v === undefined ? '' : v);
    }
  }
  var areas = cleanAreas(areasRaw);
  if (!areas.length) areas = [ { name: 'Inside Feni', charge: out.deliveryInside }, { name: 'Outside Feni', charge: out.deliveryOutside } ];
  out.deliveryAreas = areas;
  out.paymentMethods = cleanPayMethods(payRaw);
  var cc = cleanNames(catsRaw), ec = cleanNames(expCatsRaw);
  out.categories = cc.length ? cc : DEFAULT_PRODUCT_CATS.slice();
  out.expenseCategories = ec.length ? ec : DEFAULT_EXPENSE_CATS.slice();
  try { CacheService.getScriptCache().put(SETTINGS_CACHE_KEY, JSON.stringify(out), 300); } catch (e) {}
  return out;
}

// পেমেন্ট মাধ্যম (ব্যাংক / মোবাইল ব্যাংকিং) পরিষ্কার
function cleanPayMethods(list) {
  var seen = {}, out = [];
  (list || []).forEach(function(m) {
    var name = String((m && m.name) || '').trim().slice(0, 40);
    var number = String((m && m.number) || '').trim().slice(0, 60);
    if (!name || !number) return;
    var k = name.toLowerCase();
    if (seen[k] || k === 'cash on delivery') return;
    seen[k] = true;
    out.push({
      type: (m.type === 'bank') ? 'bank' : 'mobile',
      name: name, number: number,
      holder: String(m.holder || '').trim().slice(0, 60),
      note: String(m.note || '').trim().slice(0, 150)
    });
  });
  return out;
}

// নামের তালিকা পরিষ্কার: ফাঁকা বাদ, ডুপ্লিকেট বাদ
function cleanNames(list) {
  var seen = {}, out = [];
  (list || []).forEach(function(n) {
    var name = String(n === null || n === undefined ? '' : n).trim().slice(0, 40);
    if (!name || seen[name.toLowerCase()]) return;
    seen[name.toLowerCase()] = true;
    out.push(name);
  });
  return out;
}

// ডেলিভারি এরিয়া তালিকা পরিষ্কার: নাম আবশ্যক, চার্জ ≥ ০, ডুপ্লিকেট নাম বাদ
function cleanAreas(list) {
  var seen = {}, out = [];
  (list || []).forEach(function(a) {
    var name = String((a && a.name) || '').trim().slice(0, 60);
    var ch = Number(a && a.charge);
    if (!name || isNaN(ch) || ch < 0) return;
    var k = name.toLowerCase();
    if (seen[k]) return;
    seen[k] = true;
    out.push({ name: name, charge: r2(ch) });
  });
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
  ['categories', 'expenseCategories'].forEach(function(key) {
    if (!Array.isArray(s[key])) return;
    var lst = cleanNames(s[key]);
    if (lst.length) put(key, JSON.stringify(lst), true);
  });
  if (Array.isArray(s.paymentMethods)) put('paymentMethods', JSON.stringify(cleanPayMethods(s.paymentMethods)), true);
  if (Array.isArray(s.deliveryAreas)) {
    var ar = cleanAreas(s.deliveryAreas);
    if (!ar.length) return { error: 'কমপক্ষে একটি ডেলিভারি এরিয়া দিন' };
    put('deliveryAreas', JSON.stringify(ar), true);
  }
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
      return { itemId: String(o.itemid), name: String(o.name || ''), brand: String(o.brand || ''),
             unit: normalizeUnit(o.unit), colors: cleanList(o.colors), sizes: cleanList(o.sizes) };
    })
  };
}

function addItem(it) {
  if (!it || !String(it.name || '').trim()) return { error: 'Product Name দিন' };
  var sheet   = ensureSheet('Items');
  var headers = getHeaders(sheet);
  var id      = 'ITM-' + new Date().getTime();
  appendRowText(sheet, headers, {
    'itemid': id, 'name': String(it.name).trim(), 'brand': String(it.brand || '').trim(), 'createdat': new Date(),
    'unit': normalizeUnit(it.unit),
    'colors': normalizeUnit(it.unit) === 'PCS' ? cleanList(it.colors) : '',
    'sizes':  normalizeUnit(it.unit) === 'PCS' ? cleanList(it.sizes)  : ''
  }, []);
  clearProductsCache();
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
    var uom = normalizeUnit(it.unit);
    sheet.getRange(i + 1, colIndex(sheet, 'unit') + 1).setValue(uom);
    sheet.getRange(i + 1, colIndex(sheet, 'colors') + 1).setValue(uom === 'PCS' ? cleanList(it.colors) : '');
    sheet.getRange(i + 1, colIndex(sheet, 'sizes') + 1).setValue(uom === 'PCS' ? cleanList(it.sizes) : '');
    clearProductsCache();
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
    if (String(rows[i][idC]) === String(itemId)) { sheet.deleteRow(i + 1); clearProductsCache(); return { success: true }; }
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
      unit: normalizeUnit(it.unit),
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
  var invoiceNo = '';

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet   = ensureSheet('Purchases');
    var headers = getHeaders(sheet);
    invoiceNo = nextSeqId(sheet, 'invoiceNo', 'PUR');
    appendRowText(sheet, headers, {
      'invoiceno': invoiceNo, 'date': date, 'company': company,
      'mobile': String(p.mobile || '').trim(), 'address': String(p.address || '').trim(),
      'itemsjson': JSON.stringify(items),
      'total': total, 'discount': discount, 'payable': payable, 'paid': paid, 'due': due,
      'authority': String(p.authority || '').trim(), 'buyer': String(p.buyer || '').trim(),
      'createdat': new Date()
    }, ['date', 'mobile']);
    clearProductsCache();
  } finally { lock.releaseLock(); }

  return { success: true, purchase: {
    invoiceNo: invoiceNo, date: date, company: company,
    mobile: String(p.mobile || '').trim(), address: String(p.address || '').trim(),
    items: items, total: total, discount: discount, payable: payable, paid: paid, due: due,
    authority: String(p.authority || '').trim(), buyer: String(p.buyer || '').trim()
  } };
}

// Purchase এডিট: Invoice নম্বর ঠিক থাকে, বাকি তথ্য নতুন করে হিসাব হয়ে শীটে বসে
function updatePurchase(invoiceNo, p) {
  if (!invoiceNo || !p) return { error: 'তথ্য অসম্পূর্ণ' };
  var company = String(p.company || '').trim();
  if (!company) return { error: 'Company Name দিন' };

  var items = (p.items || []).map(function(it) {
    var qty = Number(it.qty) || 0, price = Number(it.price) || 0;
    return {
      name: String(it.name || '').trim(), brand: String(it.brand || '').trim(),
      unit: normalizeUnit(it.unit),
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
  var due  = r2(payable - paid);
  var date = /^\d{4}-\d{2}-\d{2}$/.test(String(p.date || '')) ? String(p.date) : todayStr();

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Purchases');
    var rows  = sheet.getDataRange().getValues();
    var iC = colIndex(sheet, 'invoiceNo');
    var vals = {
      'date': date, 'company': company,
      'mobile': String(p.mobile || '').trim(), 'address': String(p.address || '').trim(),
      'itemsjson': JSON.stringify(items),
      'total': total, 'discount': discount, 'payable': payable, 'paid': paid, 'due': due,
      'authority': String(p.authority || '').trim(), 'buyer': String(p.buyer || '').trim()
    };
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][iC]) !== String(invoiceNo)) continue;
      Object.keys(vals).forEach(function(k) {
        var c = colIndex(sheet, k);
        if (c < 0) return;
        var cell = sheet.getRange(i + 1, c + 1);
        if (k === 'date' || k === 'mobile') cell.setNumberFormat('@');
        cell.setValue(vals[k]);
      });
      clearProductsCache();
      return { success: true, purchase: {
        invoiceNo: String(invoiceNo), date: date, company: company,
        mobile: vals.mobile, address: vals.address, items: items,
        total: total, discount: discount, payable: payable, paid: paid, due: due,
        authority: vals.authority, buyer: vals.buyer
      } };
    }
    return { error: 'Invoice পাওয়া যায়নি' };
  } finally { lock.releaseLock(); }
}

// Purchase ডিলেট: Invoice এর সব Payment রেকর্ডসহ মুছে যায়
function deletePurchase(invoiceNo) {
  if (!invoiceNo) return { error: 'Invoice দরকার' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Purchases');
    var rows  = sheet.getDataRange().getValues();
    var iC = colIndex(sheet, 'invoiceNo');
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][iC]) !== String(invoiceNo)) continue;
      sheet.deleteRow(i + 1);
      clearProductsCache();
      var pSheet = ensureSheet('Payments');
      var pRows  = pSheet.getDataRange().getValues();
      var pC = colIndex(pSheet, 'invoiceNo');
      for (var j = pRows.length - 1; j >= 1; j--) {
        if (String(pRows[j][pC]) === String(invoiceNo)) pSheet.deleteRow(j + 1);
      }
      return { success: true };
    }
    return { error: 'Invoice পাওয়া যায়নি' };
  } finally { lock.releaseLock(); }
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

// সব Due Invoice একসাথে সম্পূর্ণ পরিশোধ (প্রতিটির জন্য আলাদা Payment রেকর্ড)
function payAllDue(date, note) {
  var pdate = /^\d{4}-\d{2}-\d{2}$/.test(String(date || '')) ? String(date) : todayStr();
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = ensureSheet('Purchases');
    var rows  = sheet.getDataRange().getValues();
    var iC = colIndex(sheet, 'invoiceNo'), paidC = colIndex(sheet, 'paid'),
        dueC = colIndex(sheet, 'due'), payC = colIndex(sheet, 'payable');
    var pSheet = ensureSheet('Payments'), pHeaders = getHeaders(pSheet);
    var count = 0, total = 0, stamp = new Date().getTime();
    for (var i = 1; i < rows.length; i++) {
      var due = r2(Number(rows[i][dueC]) || 0);
      if (due <= 0.001) continue;
      var newPaid = r2((Number(rows[i][paidC]) || 0) + due);
      sheet.getRange(i + 1, paidC + 1).setValue(newPaid);
      sheet.getRange(i + 1, dueC + 1).setValue(0);
      appendRowText(pSheet, pHeaders, {
        'paymentid': 'PAY-' + stamp + '-' + (count + 1), 'invoiceno': String(rows[i][iC]),
        'date': pdate, 'amount': due, 'note': String(note || 'All Due Pay').trim().slice(0, 200), 'createdat': new Date()
      }, ['date']);
      count++; total = r2(total + due);
    }
    if (!count) return { error: 'কোনো Due নেই' };
    return { success: true, count: count, total: total };
  } finally { lock.releaseLock(); }
}

/* ======== SLIDES (হোমপেজ স্লাইডের ছবি) ======== */

function readSlides() {
  var sheet = ensureSheet('Slides');
  var list = rowsToObjects(sheet.getDataRange().getValues()).map(function(o) {
    return { slideId: String(o.slideId), image: String(o.image || ''), sort: Number(o.sort) || 0 };
  }).filter(function(x) { return x.slideId && x.image; });
  list.sort(function(a, b) { return a.sort - b.sort; });
  return list;
}

function getSlides() {
  return { slides: readSlides() };
}

function addSlide(image) {
  image = String(image || '').trim();
  if (!/^https?:\/\//i.test(image) || image.length > 800) return { error: 'ছবির লিংক সঠিক নয়' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Slides'), headers = getHeaders(sheet);
    var max = 0; readSlides().forEach(function(x) { if (x.sort > max) max = x.sort; });
    var id = 'SL-' + new Date().getTime() + '-' + Math.floor(Math.random() * 100000);
    appendRowText(sheet, headers, { 'slideid': id, 'image': image, 'sort': max + 1, 'createdat': new Date() }, []);
    return { success: true, slideId: id };
  } finally { lock.releaseLock(); }
}

function deleteSlide(slideId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Slides'), c = colIndex(sheet, 'slideId');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][c]) === String(slideId)) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { error: 'স্লাইড পাওয়া যায়নি' };
  } finally { lock.releaseLock(); }
}

// সিরিয়াল বদল: পাশের স্লাইডের সাথে sort মান অদলবদল
function moveSlide(slideId, dir) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Slides');
    var idC = colIndex(sheet, 'slideId'), sC = colIndex(sheet, 'sort');
    var rows = sheet.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][idC]) list.push({ row: i + 1, id: String(rows[i][idC]), sort: Number(rows[i][sC]) || 0 });
    }
    list.sort(function(a, b) { return a.sort - b.sort; });
    var k = -1;
    list.forEach(function(x, n) { if (x.id === String(slideId)) k = n; });
    if (k < 0) return { error: 'স্লাইড পাওয়া যায়নি' };
    var j = (dir === 'up') ? k - 1 : k + 1;
    if (j < 0 || j >= list.length) return { success: true };
    // সবাইকে ১..N ক্রমে সাজিয়ে দুটি অদলবদল (ডুপ্লিকেট sort মান থাকলেও ঠিক থাকে)
    var tmp = list[k]; list[k] = list[j]; list[j] = tmp;
    list.forEach(function(x, n) { sheet.getRange(x.row, sC + 1).setValue(n + 1); });
    return { success: true };
  } finally { lock.releaseLock(); }
}

/* ======== STOCK (Purchase − Delivery সম্পন্ন অর্ডার, পণ্যের নাম মিলিয়ে) ======== */

function getStock() {
  var pSheet = ensureSheet('Purchases');
  var purchases = rowsToObjects(pSheet.getDataRange().getValues()).map(purchaseFromObj);

  var oSheet = ensureSheet('Orders');
  var orders = rowsToObjects(oSheet.getDataRange().getValues());
  var soldByName = {};
  orders.forEach(function(o) {
    // বুকিং (Pending) ও ডেলিভারি সম্পন্ন অর্ডারে স্টক কমে। নতুন অর্ডার (New) ও বাতিল (Cancel) ধরা হয় না —
    // তাই বুকিং থেকে বাতিল হলে পণ্য অটো স্টকে ফেরত আসে।
    var st = String(o.status || 'Pending').trim().toLowerCase();
    if (st !== 'pending' && st !== 'delivery' && st !== 'delivered') return;
    var k = String(o.productname || '').trim().toLowerCase();
    if (!k) return;
    soldByName[k] = (soldByName[k] || 0) + (Number(o.quantity) || 0);
  });

  var details = {}, products = {};
  purchases.forEach(function(p) {
    (p.items || []).forEach(function(it) {
      var nk = it.name.toLowerCase();
      var dk = [nk, it.brand.toLowerCase(), (it.colour || '').toLowerCase(), (it.size || '').toLowerCase()].join('|');
      if (!details[dk]) details[dk] = { product: it.name, brand: it.brand, unit: normalizeUnit(it.unit), colour: it.colour || '', size: it.size || '', purchased: 0 };
      details[dk].purchased += Number(it.qty) || 0;

      if (!products[nk]) products[nk] = { product: it.name, unit: normalizeUnit(it.unit), brands: {}, purchased: 0 };
      products[nk].purchased += Number(it.qty) || 0;
      if (it.brand) products[nk].brands[it.brand] = true;
    });
  });

  var prodList = Object.keys(products).map(function(nk) {
    var x = products[nk], sold = soldByName[nk] || 0;
    return { product: x.product, unit: x.unit, brand: Object.keys(x.brands).join(', '),
             purchased: x.purchased, sold: sold, stock: x.purchased - sold };
  }).sort(function(a, b) { return a.product.localeCompare(b.product); });

  var detList = Object.keys(details).map(function(k) { return details[k]; })
    .sort(function(a, b) { return a.product.localeCompare(b.product); });

  return { products: prodList, details: detList, sold: soldByName };
}

/* ======== ADMINISTRATION (ব্যবস্থাপক) ======== */

function hashPassword(salt, pw) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return bytes.map(function(b) { var h = (b < 0 ? b + 256 : b).toString(16); return h.length === 1 ? '0' + h : h; }).join('');
}

// শীটে আগে নম্বর হিসেবে সেভ হওয়া ১০ সংখ্যার (১XXXXXXXXX) নম্বরে শুরুতে ০ বসানো
function fixPhone(v) {
  var t = String(v == null ? '' : v).trim();
  return /^1\d{9}$/.test(t) ? '0' + t : t;
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

function storedPw(o) { return String(o.password === undefined || o.password === null ? '' : o.password).trim(); }

function passwordTaken(rows, pw, exceptId) {
  return rows.some(function(o) {
    if (exceptId && String(o.adminid) === String(exceptId)) return false;
    return storedPw(o) !== '' && storedPw(o) === pw;
  });
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
  if (!mobile) return { error: 'মোবাইল নং দিন (পাসওয়ার্ড উদ্ধারের জন্য)' };
  if (!/^\d{5}$/.test(String(a.password || ''))) return { error: 'পাসওয়ার্ড অবশ্যই ৫ সংখ্যার হতে হবে' };

  var sheet = ensureSheet('Admins');
  var rows  = rowsToObjects(sheet.getDataRange().getValues());
  if (adminDuplicate(rows, mobile, email, null)) return { error: 'এই মোবাইল/ইমেইল আগে থেকেই আছে' };
  if (passwordTaken(rows, String(a.password), null)) return { error: 'এই পাসওয়ার্ড অন্য একজন ব্যবহার করছেন, আলাদা ৫ সংখ্যার পাসওয়ার্ড দিন' };

  var salt = Utilities.getUuid();
  var id   = 'ADM-' + new Date().getTime();
  appendRowText(sheet, getHeaders(sheet), {
    'adminid': id, 'name': String(a.name).trim(), 'designation': String(a.designation || '').trim(),
    'mobile': mobile, 'email': email, 'passwordhash': hashPassword(salt, String(a.password)),
    'salt': salt, 'createdat': new Date(), 'password': String(a.password)
  }, ['mobile', 'password']);
  return { success: true, adminId: id };
}

function updateAdmin(a) {
  if (!a || !a.adminId) return { error: 'adminId দরকার' };
  if (!String(a.name || '').trim()) return { error: 'নাম দিন' };
  var mobile = normMobile(a.mobile), email = String(a.email || '').trim().toLowerCase();
  if (!mobile) return { error: 'মোবাইল নং দিন (পাসওয়ার্ড উদ্ধারের জন্য)' };
  if (a.password && !/^\d{5}$/.test(String(a.password))) return { error: 'পাসওয়ার্ড অবশ্যই ৫ সংখ্যার হতে হবে' };

  var sheet = ensureSheet('Admins');
  var data  = sheet.getDataRange().getValues();
  if (adminDuplicate(rowsToObjects(data), mobile, email, a.adminId)) return { error: 'এই মোবাইল/ইমেইল আগে থেকেই আছে' };
  if (a.password && passwordTaken(rowsToObjects(data), String(a.password), a.adminId)) return { error: 'এই পাসওয়ার্ড অন্য একজন ব্যবহার করছেন, আলাদা ৫ সংখ্যার পাসওয়ার্ড দিন' };

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
      setCol('password', String(a.password), true);
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

// শুধু ৫ সংখ্যার পাসওয়ার্ড দিয়ে লগইন (প্রত্যেক ব্যবস্থাপকের পাসওয়ার্ড আলাদা)।
// ১০ বার ভুল হলে ১৫ মিনিট লক (৫ সংখ্যার পাসওয়ার্ড ব্রুট-ফোর্স ঠেকাতে)।
function adminLogin(identifier, password) {
  var pw = String(password || '').trim();
  if (!/^\d{5}$/.test(pw)) return { error: 'ভুল পাসওয়ার্ড, ৫ সংখ্যার পাসওয়ার্ড দিন' };
  var cache = CacheService.getScriptCache();
  var key   = 'lf_pw';
  var fails = Number(cache.get(key)) || 0;
  if (fails >= 10) return { error: 'অনেকবার ভুল হয়েছে। ১৫ মিনিট পর আবার চেষ্টা করুন' };

  // মালিকের মাস্টার পাসওয়ার্ড: Apps Script → Project Settings → Script properties → MASTER_PIN (কোডে বা সাইটে লেখা থাকে না)
  var master = String(PropertiesService.getScriptProperties().getProperty('MASTER_PIN') || '').trim();
  if (master && pw === master) {
    cache.remove(key);
    return { success: true, token: newSession('Owner'), admin: { name: '', designation: '' } };
  }

  var sheet = ensureSheet('Admins');
  var rows  = rowsToObjects(sheet.getDataRange().getValues());
  for (var i = 0; i < rows.length; i++) {
    var o = rows[i], sp = storedPw(o);
    var ok = sp ? (sp === pw) : (o.salt && o.passwordhash && hashPassword(String(o.salt), pw) === String(o.passwordhash));
    if (ok) {
      cache.remove(key);
      return { success: true, token: newSession(String(o.name || '')), admin: { name: String(o.name || ''), designation: String(o.designation || '') } };
    }
  }
  cache.put(key, String(fails + 1), 900);
  return { error: 'ভুল পাসওয়ার্ড, আবার চেষ্টা করুন' };
}

/* ---------- পাসওয়ার্ড উদ্ধার (ওয়েবসাইটের ইমেইলে Pin Code) ---------- */

function listAdminNames() {
  var sheet = ensureSheet('Admins');
  return { admins: rowsToObjects(sheet.getDataRange().getValues()).map(function(o) {
    return { adminId: String(o.adminid), name: String(o.name || '') };
  }) };
}

function maskEmail(e) {
  var p = String(e).split('@');
  return (p[0].slice(0, 2) + '***@' + (p[1] || ''));
}

function sendRecoveryPin(adminId, mobile) {
  mobile = normMobile(mobile);
  if (!adminId || !mobile) return { error: 'নাম সিলেক্ট করুন ও মোবাইল নং দিন' };
  var cache = CacheService.getScriptCache();
  var sk = 'rs_' + adminId;
  var sends = Number(cache.get(sk)) || 0;
  if (sends >= 3) return { error: 'অনেকবার চেষ্টা হয়েছে। ১৫ মিনিট পর আবার চেষ্টা করুন' };
  cache.put(sk, String(sends + 1), 900);

  var rows = rowsToObjects(ensureSheet('Admins').getDataRange().getValues());
  var me = null;
  rows.forEach(function(o) { if (String(o.adminid) === String(adminId)) me = o; });
  if (!me || normMobile(me.mobile) !== mobile) return { error: 'নাম ও মোবাইল নং মিলছে না' };
  if (!storedPw(me)) return { error: 'এই ব্যবস্থাপকের পাসওয়ার্ড উদ্ধার করা যাবে না। Management থেকে নতুন পাসওয়ার্ড সেট করুন' };

  var email = String(getSettings().webEmail || '').trim();
  if (!/^\S+@\S+\.\S+$/.test(email)) return { error: 'ওয়েবসাইটের ইমেইল সেট করা নেই (Settings → Web Info)' };

  var pin = ('000000' + (parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000)).slice(-6);
  cache.put('rp_' + adminId, hashPassword(String(adminId), pin), 600);
  cache.remove('rf_' + adminId);
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'পাসওয়ার্ড উদ্ধারের Pin Code',
      body: String(me.name || '') + ' এর পাসওয়ার্ড উদ্ধারের Pin Code: ' + pin + '\n\nএই কোডটি ১০ মিনিট পর্যন্ত কার্যকর থাকবে। আপনি অনুরোধ না করে থাকলে এই ইমেইল উপেক্ষা করুন।'
    });
  } catch (err) {
    return { error: 'ইমেইল পাঠানো যায়নি। Apps Script-এ ইমেইল পাঠানোর অনুমতি (authorize) দিন' };
  }
  return { success: true, emailHint: maskEmail(email) };
}

function verifyRecoveryPin(adminId, pin) {
  if (!adminId || !pin) return { error: 'Pin Code দিন' };
  var cache = CacheService.getScriptCache();
  var fk = 'rf_' + adminId;
  var fails = Number(cache.get(fk)) || 0;
  if (fails >= 5) { cache.remove('rp_' + adminId); return { error: 'অনেকবার ভুল হয়েছে। নতুন Pin Code নিন' }; }
  var stored = cache.get('rp_' + adminId);
  if (!stored) return { error: 'Pin Code এর মেয়াদ শেষ বা পাঠানো হয়নি। আবার Pin Code নিন' };
  if (hashPassword(String(adminId), String(pin).trim()) !== stored) {
    cache.put(fk, String(fails + 1), 900);
    return { error: 'ভুল Pin Code' };
  }
  var rows = rowsToObjects(ensureSheet('Admins').getDataRange().getValues());
  var me = null;
  rows.forEach(function(o) { if (String(o.adminid) === String(adminId)) me = o; });
  if (!me || !storedPw(me)) return { error: 'পাসওয়ার্ড পাওয়া যায়নি' };
  cache.remove('rp_' + adminId); cache.remove(fk);
  return { success: true, name: String(me.name || ''), password: storedPw(me) };
}


/* ======== INVEST / CASH (Deposit, Withdrawal) ======== */

function validYmd(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d || '')) ? String(d) : todayStr(); }

function getInvests() {
  var sheet = ensureSheet('Invests');
  return { invests: rowsToObjects(sheet.getDataRange().getValues()).map(function(o) {
    return { id: String(o.id), date: fmtDate(o.date), name: String(o.name || ''), mobile: String(o.mobile || ''),
             address: String(o.address || ''), amount: Number(o.amount) || 0 };
  }).reverse() };
}

function addInvest(r) {
  if (!r || !String(r.name || '').trim()) return { error: 'Name দিন' };
  var amt = r2(r.amount);
  if (!(amt > 0)) return { error: 'সঠিক Amount দিন' };
  var sheet = ensureSheet('Invests');
  var id = 'INV-' + new Date().getTime();
  appendRowText(sheet, getHeaders(sheet), {
    'id': id, 'date': validYmd(r.date), 'name': String(r.name).trim(), 'mobile': String(r.mobile || '').trim(),
    'address': String(r.address || '').trim(), 'amount': amt, 'createdat': new Date()
  }, ['date', 'mobile']);
  return { success: true, id: id };
}

function updateInvest(r) {
  if (!r || !r.id) return { error: 'id দরকার' };
  if (!String(r.name || '').trim()) return { error: 'Name দিন' };
  var amt = r2(r.amount);
  if (!(amt > 0)) return { error: 'সঠিক Amount দিন' };
  return updateRecordRow('Invests', r.id, {
    'date': validYmd(r.date), 'name': String(r.name).trim(), 'mobile': String(r.mobile || '').trim(),
    'address': String(r.address || '').trim(), 'amount': amt
  }, ['date', 'mobile']);
}

function cashList(sheetName, key) {
  var sheet = ensureSheet(sheetName);
  var out = {};
  out[key] = rowsToObjects(sheet.getDataRange().getValues()).map(function(o) {
    return { id: String(o.id), date: fmtDate(o.date), description: String(o.description || ''),
             amount: Number(o.amount) || 0, person: String(o.person || '') };
  }).reverse();
  return out;
}
function getDeposits()    { return cashList('Deposits', 'deposits'); }
function getWithdrawals() { return cashList('Withdrawals', 'withdrawals'); }

function addCashRecord(sheetName, prefix, r) {
  if (!r) return { error: 'তথ্য অসম্পূর্ণ' };
  var amt = r2(r.amount);
  if (!(amt > 0)) return { error: 'সঠিক Amount দিন' };
  if (!String(r.person || '').trim()) return { error: 'নাম দিন' };
  var sheet = ensureSheet(sheetName);
  var id = prefix + new Date().getTime();
  appendRowText(sheet, getHeaders(sheet), {
    'id': id, 'date': validYmd(r.date), 'description': String(r.description || '').trim().slice(0, 300),
    'amount': amt, 'person': String(r.person).trim(), 'createdat': new Date()
  }, ['date']);
  return { success: true, id: id };
}

function updateCashRecord(sheetName, r) {
  if (!r || !r.id) return { error: 'id দরকার' };
  var amt = r2(r.amount);
  if (!(amt > 0)) return { error: 'সঠিক Amount দিন' };
  if (!String(r.person || '').trim()) return { error: 'নাম দিন' };
  return updateRecordRow(sheetName, r.id, {
    'date': validYmd(r.date), 'description': String(r.description || '').trim().slice(0, 300),
    'amount': amt, 'person': String(r.person).trim()
  }, ['date']);
}

// সাধারণ আপডেট: id মিলিয়ে নির্দিষ্ট কলামগুলো বদলায়
function updateRecordRow(sheetName, id, map, textKeys) {
  var sheet = ensureSheet(sheetName);
  var rows  = sheet.getDataRange().getValues();
  var idC   = colIndex(sheet, 'id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) !== String(id)) continue;
    Object.keys(map).forEach(function(k) {
      var c = colIndex(sheet, k);
      if (c < 0) return;
      var cell = sheet.getRange(i + 1, c + 1);
      if (textKeys.indexOf(k) >= 0) cell.setNumberFormat('@');
      cell.setValue(map[k]);
    });
    return { success: true };
  }
  return { error: 'রেকর্ড পাওয়া যায়নি' };
}

function deleteRecord(sheetName, id) {
  if (!id) return { error: 'id দরকার' };
  var sheet = ensureSheet(sheetName);
  var rows  = sheet.getDataRange().getValues();
  var idC   = colIndex(sheet, 'id');
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][idC]) === String(id)) { sheet.deleteRow(i + 1); return { success: true }; }
  }
  return { error: 'রেকর্ড পাওয়া যায়নি' };
}


/* ======== INCOME (আয় এন্ট্রি; একটি Invoice-এ একাধিক লাইন) ======== */

function cleanIncomeItems(items) {
  return (items || []).map(function(it) {
    return { description: String(it.description || '').trim().slice(0, 300), amount: r2(it.amount) };
  }).filter(function(it) { return it.amount > 0; });
}

function incomeFromObj(o) {
  var items = [];
  try { items = JSON.parse(o.itemsjson || '[]'); } catch (e) {}
  return { invoiceNo: String(o.invoiceno), date: fmtDate(o.date), items: items, total: Number(o.total) || 0 };
}

function getIncomes() {
  var sheet = ensureSheet('Incomes');
  return { incomes: rowsToObjects(sheet.getDataRange().getValues()).map(incomeFromObj).reverse() };
}

function addIncome(inc) {
  if (!inc) return { error: 'তথ্য অসম্পূর্ণ' };
  var items = cleanIncomeItems(inc.items);
  if (!items.length) return { error: 'কমপক্ষে একটি সঠিক Amount দিন' };
  var total = r2(items.reduce(function(a, it) { return a + it.amount; }, 0));
  var date = validYmd(inc.date);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Incomes');
    var headers = getHeaders(sheet);
    var vc = colIndex(sheet, 'invoiceNo'), max = 0;
    if (vc >= 0 && sheet.getLastRow() > 1) {
      sheet.getRange(2, vc + 1, sheet.getLastRow() - 1, 1).getValues().forEach(function(r) {
        var m = String(r[0] || '').match(/(\d+)\s*$/);
        if (m) max = Math.max(max, parseInt(m[1], 10));
      });
    }
    var invoiceNo = 'INC-' + ('0000' + (max + 1)).slice(-4);
    appendRowText(sheet, headers, {
      'invoiceno': invoiceNo, 'date': date, 'itemsjson': JSON.stringify(items), 'total': total, 'createdat': new Date()
    }, ['invoiceno', 'date']);
    return { success: true, income: { invoiceNo: invoiceNo, date: date, items: items, total: total } };
  } finally { lock.releaseLock(); }
}

function updateIncome(invoiceNo, inc) {
  if (!invoiceNo || !inc) return { error: 'তথ্য অসম্পূর্ণ' };
  var items = cleanIncomeItems(inc.items);
  if (!items.length) return { error: 'কমপক্ষে একটি সঠিক Amount দিন' };
  var total = r2(items.reduce(function(a, it) { return a + it.amount; }, 0));
  var date = validYmd(inc.date);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Incomes');
    var rows  = sheet.getDataRange().getValues();
    var iC = colIndex(sheet, 'invoiceNo');
    var vals = { 'date': date, 'itemsjson': JSON.stringify(items), 'total': total };
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][iC]) !== String(invoiceNo)) continue;
      Object.keys(vals).forEach(function(k) {
        var c = colIndex(sheet, k);
        if (c < 0) return;
        var cell = sheet.getRange(i + 1, c + 1);
        if (k === 'date') cell.setNumberFormat('@');
        cell.setValue(vals[k]);
      });
      return { success: true, income: { invoiceNo: String(invoiceNo), date: date, items: items, total: total } };
    }
    return { error: 'Invoice পাওয়া যায়নি' };
  } finally { lock.releaseLock(); }
}

function deleteIncome(invoiceNo) {
  if (!invoiceNo) return { error: 'Invoice দরকার' };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var sheet = ensureSheet('Incomes');
    var rows  = sheet.getDataRange().getValues();
    var iC = colIndex(sheet, 'invoiceNo');
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][iC]) === String(invoiceNo)) { sheet.deleteRow(i + 1); return { success: true }; }
    }
    return { error: 'Invoice পাওয়া যায়নি' };
  } finally { lock.releaseLock(); }
}
