// ===== CONFIGURATION =====
var MASTER_SHEET_ID = 'YOUR_MASTER_SHEET_ID_HERE';
var MASTER_FOLDER_ID = 'YOUR_MASTER_FOLDER_ID_HERE';
var SLIP_FOLDER_ID = 'YOUR_SLIP_FOLDER_ID_HERE';
var WEBAPP_URL = 'YOUR_WEBAPP_URL_HERE';
var ADMIN_EMAIL = 'admin@kpcrm.net';
var PROMPTPAY_NUMBER = '0995588665';
var BANK_ACCOUNT = '0271813236';
var BANK_NAME = 'ธนาคารกสิกรไทย';
var ACCOUNT_NAME = 'นายคฑาวุธ มีกุญชร';

// ===== MAIN ENTRY POINT =====
function doGet(e) {
  var page = e.parameter.page || 'index';
  var shopId = e.parameter.shopId || '';

  if (page === 'register') {
    return HtmlService.createHtmlOutputFromFile('register')
      .setTitle('ลงทะเบียนร้านใหม่')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'display') {
    var template = HtmlService.createTemplateFromFile('display');
    template.shopId = shopId;
    return template.evaluate()
      .setTitle('หน้าจอลูกค้า')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'queue') {
    var template = HtmlService.createTemplateFromFile('queue');
    template.shopId = shopId;
    return template.evaluate()
      .setTitle('ติดตามคิว')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'superadmin') {
    return HtmlService.createHtmlOutputFromFile('superadmin')
      .setTitle('Super Admin')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Coffee Shop POS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

// ===== CACHE SERVICE =====
function getCacheService() {
  return CacheService.getScriptCache();
}

function setCache(key, value, expirationInSeconds) {
  var cache = getCacheService();
  cache.put(key, JSON.stringify(value), expirationInSeconds || 300);
}

function getCache(key) {
  var cache = getCacheService();
  var cached = cache.get(key);
  return cached ? JSON.parse(cached) : null;
}

function clearCache(key) {
  var cache = getCacheService();
  cache.remove(key);
}

// ===== PRE-WARM CACHE TRIGGER =====
function createPreWarmTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'preWarmCache') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('preWarmCache')
    .timeBased()
    .everyMinutes(5)
    .create();
}

function preWarmCache() {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  if (!shopsSheet) return;

  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];
  var shopIdCol = headers.indexOf('รหัสร้าน');
  var sheetIdCol = headers.indexOf('SheetID');

  for (var i = 1; i < data.length; i++) {
    var shopId = data[i][shopIdCol];
    var sheetId = data[i][sheetIdCol];
    if (shopId && sheetId) {
      try {
        var shopSheet = SpreadsheetApp.openById(sheetId);
        var products = getProductsFromSheet(shopSheet);
        setCache('products_' + shopId, products, 600);
      } catch (e) {
        Logger.log('Error pre-warming cache for shop ' + shopId + ': ' + e.message);
      }
    }
  }
}

// ===== AUTHENTICATION =====
function login(email, password) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');

  if (!shopsSheet) {
    return { success: false, message: 'ไม่พบข้อมูลร้านค้า' };
  }

  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var emailCol = headers.indexOf('อีเมล');
  var passwordCol = headers.indexOf('รหัสผ่าน');
  var shopIdCol = headers.indexOf('รหัสร้าน');
  var shopNameCol = headers.indexOf('ชื่อร้าน');
  var sheetIdCol = headers.indexOf('SheetID');
  var expiryCol = headers.indexOf('วันหมดอายุ');
  var statusCol = headers.indexOf('สถานะ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][emailCol] === email && data[i][passwordCol] === password) {
      var expiryDate = new Date(data[i][expiryCol]);
      var today = new Date();
      var daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

      if (data[i][statusCol] !== 'ใช้งาน') {
        return { success: false, message: 'บัญชีถูกระงับการใช้งาน' };
      }

      if (daysLeft <= 0) {
        return { success: false, message: 'ไลเซ่นส์หมดอายุแล้ว' };
      }

      return {
        success: true,
        shopId: data[i][shopIdCol],
        shopName: data[i][shopNameCol],
        sheetId: data[i][sheetIdCol],
        expiryDate: expiryDate.toISOString(),
        daysLeft: daysLeft,
        warning: daysLeft <= 30 ? 'ไลเซ่นส์จะหมดอายุในอีก ' + daysLeft + ' วัน' : null
      };
    }
  }

  return { success: false, message: 'อีเมลหรือรหัสผ่านไม่ถูกต้อง' };
}

function getLicenseStatus(shopId) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var shopIdCol = headers.indexOf('รหัสร้าน');
  var shopNameCol = headers.indexOf('ชื่อร้าน');
  var emailCol = headers.indexOf('อีเมล');
  var expiryCol = headers.indexOf('วันหมดอายุ');
  var packageCol = headers.indexOf('แพ็คเกจ');
  var statusCol = headers.indexOf('สถานะ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][shopIdCol] === shopId) {
      var expiryDate = new Date(data[i][expiryCol]);
      var today = new Date();
      var daysLeft = Math.ceil((expiryDate - today) / (1000 * 60 * 60 * 24));

      return {
        shopName: data[i][shopNameCol],
        email: data[i][emailCol],
        package: data[i][packageCol],
        expiryDate: expiryDate.toISOString(),
        daysLeft: daysLeft,
        status: data[i][statusCol]
      };
    }
  }

  return null;
}

function changePassword(shopId, oldPassword, newPassword) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var shopIdCol = headers.indexOf('รหัสร้าน');
  var passwordCol = headers.indexOf('รหัสผ่าน');

  for (var i = 1; i < data.length; i++) {
    if (data[i][shopIdCol] === shopId) {
      if (data[i][passwordCol] !== oldPassword) {
        return { success: false, message: 'รหัสผ่านเดิมไม่ถูกต้อง' };
      }
      shopsSheet.getRange(i + 1, passwordCol + 1).setValue(newPassword);
      return { success: true, message: 'เปลี่ยนรหัสผ่านสำเร็จ' };
    }
  }

  return { success: false, message: 'ไม่พบข้อมูลร้านค้า' };
}

// ===== MASTER SHEET SETUP =====
function createMasterSheet() {
  var ss = SpreadsheetApp.create('Coffee POS Master Database');
  var masterSheetId = ss.getId();

  // สร้างชีทร้านค้า
  var shopsSheet = ss.getActiveSheet();
  shopsSheet.setName('ร้านค้า');
  shopsSheet.appendRow([
    'รหัสร้าน', 'ชื่อร้าน', 'อีเมล', 'รหัสผ่าน', 'SheetID', 'FolderID',
    'วันหมดอายุ', 'แพ็คเกจ', 'สถานะ', 'วันที่สร้าง'
  ]);

  // สร้างชีทการลงทะเบียน
  var registrationSheet = ss.insertSheet('การลงทะเบียน');
  registrationSheet.appendRow([
    'รหัสลงทะเบียน', 'ชื่อร้าน', 'อีเมล', 'รหัสผ่าน', 'แพ็คเกจ', 'จำนวนเงิน',
    'วันที่ลงทะเบียน', 'สถานะ', 'รูปสลิป', 'หมายเหตุ'
  ]);

  Logger.log('Master Sheet ID: ' + masterSheetId);
  return masterSheetId;
}

function createMasterFolder() {
  var folder = DriveApp.createFolder('Coffee POS Data');
  var folderId = folder.getId();
  Logger.log('Master Folder ID: ' + folderId);
  return folderId;
}

// ===== SHOP MANAGEMENT =====
function createNewShop(shopName, email, password, packageType, expiryDays) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');

  // สร้างรหัสร้าน
  var shopId = 'SHOP' + new Date().getTime();

  // สร้างโฟลเดอร์ย่อยสำหรับร้าน
  var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  var shopFolder = masterFolder.createFolder(shopName + '_' + shopId);
  var shopFolderId = shopFolder.getId();

  // สร้าง Spreadsheet ใหม่สำหรับร้าน
  var newShopSS = SpreadsheetApp.create(shopName + ' - POS Data');
  var newSheetId = newShopSS.getId();

  // ย้ายไฟล์ไปยังโฟลเดอร์ร้าน
  var file = DriveApp.getFileById(newSheetId);
  shopFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  // สร้างชีทต่างๆ ในไฟล์ร้าน
  createShopSheets(newShopSS);

  // คำนวณวันหมดอายุ
  var expiryDate = new Date();
  expiryDate.setDate(expiryDate.getDate() + expiryDays);

  // บันทึกข้อมูลร้านใหม่ใน Master Sheet
  shopsSheet.appendRow([
    shopId, shopName, email, password, newSheetId, shopFolderId,
    expiryDate, packageType, 'ใช้งาน', new Date()
  ]);

  // ส่งอีเมลแจ้งเตือน
  sendWelcomeEmail(email, shopName, password);

  return {
    success: true,
    shopId: shopId,
    sheetId: newSheetId,
    folderId: shopFolderId
  };
}

function createShopSheets(ss) {
  // ลบชีทเริ่มต้น
  var sheets = ss.getSheets();

  // 1. สินค้า
  var productsSheet = ss.insertSheet('สินค้า');
  productsSheet.appendRow([
    'รหัสสินค้า', 'ชื่อสินค้า', 'หมวดหมู่', 'ราคาขาย', 'ราคาทุน',
    'วัตถุดิบ', 'ตัวเลือก', 'เวลาทำ(นาที)', 'สถานะ', 'รูปภาพ'
  ]);

  // 2. หมวดหมู่สินค้า
  var categoriesSheet = ss.insertSheet('หมวดหมู่');
  categoriesSheet.appendRow(['รหัสหมวดหมู่', 'ชื่อหมวดหมู่', 'ลำดับ', 'สถานะ']);

  // 3. ตัวเลือกสินค้า
  var optionsSheet = ss.insertSheet('ตัวเลือกสินค้า');
  optionsSheet.appendRow([
    'รหัสตัวเลือก', 'กลุ่มตัวเลือก', 'ชื่อตัวเลือก', 'ราคาเพิ่ม/ลด',
    'วัตถุดิบ', 'เป็นค่าเริ่มต้น', 'ลำดับ', 'สถานะ'
  ]);

  // 4. วัตถุดิบ
  var materialsSheet = ss.insertSheet('วัตถุดิบ');
  materialsSheet.appendRow([
    'รหัสวัตถุดิบ', 'ชื่อวัตถุดิบ', 'หน่วย', 'ราคาต่อหน่วย',
    'จำนวนคงเหลือ', 'จำนวนขั้นต่ำ', 'สถานะ'
  ]);

  // 5. ล็อตวัตถุดิบ
  var lotsSheet = ss.insertSheet('ล็อตวัตถุดิบ');
  lotsSheet.appendRow([
    'รหัสล็อต', 'รหัสวัตถุดิบ', 'จำนวน', 'ราคารวม',
    'วันที่ซื้อ', 'วันหมดอายุ', 'รูปสลิป', 'จำนวนคงเหลือ'
  ]);

  // 6. การขาย
  var salesSheet = ss.insertSheet('การขาย');
  salesSheet.appendRow([
    'รหัสออเดอร์', 'วันที่', 'เวลา', 'ช่องทาง', 'รายการสินค้า',
    'ยอดรวม', 'ส่วนลด', 'ภาษี', 'ยอดสุทธิ', 'การชำระเงิน',
    'รูปสลิป', 'สถานะ', 'หมายเหตุ'
  ]);

  // 7. คิว
  var queueSheet = ss.insertSheet('คิว');
  queueSheet.appendRow([
    'รหัสออเดอร์', 'หมายเลขคิว', 'รายการ', 'สถานะ',
    'เวลาสั่ง', 'เวลาเสร็จ', 'เวลาที่ใช้(นาที)'
  ]);

  // 8. ต้นทุนรายวัน
  var dailyCostsSheet = ss.insertSheet('ต้นทุนรายวัน');
  dailyCostsSheet.appendRow([
    'วันที่', 'รายการ', 'จำนวนเงิน', 'หมวดหมู่', 'หมายเหตุ'
  ]);

  // 9. ต้นทุนรายเดือน
  var monthlyCostsSheet = ss.insertSheet('ต้นทุนรายเดือน');
  monthlyCostsSheet.appendRow([
    'เดือน/ปี', 'รายการ', 'จำนวนเงิน', 'หมวดหมู่', 'หมายเหตุ'
  ]);

  // 10. ของเสีย
  var wasteSheet = ss.insertSheet('ของเสีย');
  wasteSheet.appendRow([
    'วันที่', 'รหัสวัตถุดิบ', 'จำนวน', 'สาเหตุ', 'หมายเหตุ'
  ]);

  // 11. การตั้งค่า
  var settingsSheet = ss.insertSheet('การตั้งค่า');
  settingsSheet.appendRow(['คีย์', 'ค่า']);
  settingsSheet.appendRow(['ชื่อร้าน', '']);
  settingsSheet.appendRow(['ที่อยู่', '']);
  settingsSheet.appendRow(['เบอร์โทร', '']);
  settingsSheet.appendRow(['หมายเลขพร้อมเพย์', '']);
  settingsSheet.appendRow(['เปิดภาษี', 'false']);
  settingsSheet.appendRow(['ประเภทภาษี', 'รวมในราคา']);
  settingsSheet.appendRow(['ข้อความใบเสร็จ', 'ขอบคุณที่ใช้บริการ']);
  settingsSheet.appendRow(['ขนาดใบเสร็จ', '58mm']);
  settingsSheet.appendRow(['รูปแบบเลขออเดอร์', 'PREFIX-YYYYMMDD-XXX']);
  settingsSheet.appendRow(['Prefix เลขออเดอร์', 'ORD']);
  settingsSheet.appendRow(['รูปโลโก้', '']);
  settingsSheet.appendRow(['สไลด์โฆษณา', '']);

  // 12. ช่องทางการสั่งซื้อ
  var channelsSheet = ss.insertSheet('ช่องทางสั่งซื้อ');
  channelsSheet.appendRow([
    'รหัสช่องทาง', 'ชื่อช่องทาง', 'ประเภทเลขออเดอร์', 'สถานะ'
  ]);
  channelsSheet.appendRow(['CH001', 'หน้าร้าน', 'อัตโนมัติ', 'ใช้งาน']);
  channelsSheet.appendRow(['CH002', 'Grab', 'manual', 'ใช้งาน']);
  channelsSheet.appendRow(['CH003', 'Lineman', 'manual', 'ใช้งาน']);

  // ลบชีทเริ่มต้น
  if (sheets.length > 0) {
    ss.deleteSheet(sheets[0]);
  }
}

function sendWelcomeEmail(email, shopName, password) {
  var subject = 'ยินดีต้อนรับสู่ Coffee POS - ' + shopName;
  var body = 'สวัสดีครับ/ค่ะ\n\n' +
    'ขอบคุณที่สมัครใช้งาน Coffee POS System\n\n' +
    'ข้อมูลการเข้าสู่ระบบ:\n' +
    'ลิงค์: ' + WEBAPP_URL + '\n' +
    'อีเมล: ' + email + '\n' +
    'รหัสผ่าน: ' + password + '\n\n' +
    'กรุณาเปลี่ยนรหัสผ่านหลังจากเข้าสู่ระบบครั้งแรก\n\n' +
    'ขอบคุณครับ/ค่ะ\n' +
    'Coffee POS Team';

  MailApp.sendEmail(email, subject, body);
}

// ===== REGISTRATION =====
function registerNewShop(data) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var registrationSheet = masterSheet.getSheetByName('การลงทะเบียน');

  var regId = 'REG' + new Date().getTime();
  var amount = data.packageType === 'รายเดือน' ? 299 : 2999;

  registrationSheet.appendRow([
    regId,
    data.shopName,
    data.email,
    data.password,
    data.packageType,
    amount,
    new Date(),
    'รอชำระเงิน',
    '',
    ''
  ]);

  return { success: true, regId: regId, amount: amount };
}

function uploadRegistrationSlip(regId, base64Data) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var registrationSheet = masterSheet.getSheetByName('การลงทะเบียน');
  var data = registrationSheet.getDataRange().getValues();
  var headers = data[0];

  var regIdCol = headers.indexOf('รหัสลงทะเบียน');
  var slipCol = headers.indexOf('รูปสลิป');
  var statusCol = headers.indexOf('สถานะ');
  var emailCol = headers.indexOf('อีเมล');
  var shopNameCol = headers.indexOf('ชื่อร้าน');

  for (var i = 1; i < data.length; i++) {
    if (data[i][regIdCol] === regId) {
      // บันทึกรูปสลิป
      var folder = DriveApp.getFolderById(SLIP_FOLDER_ID);
      var blob = Utilities.newBlob(
        Utilities.base64Decode(base64Data.split(',')[1]),
        'image/png',
        'slip_' + regId + '.png'
      );
      var file = folder.createFile(blob);
      var fileUrl = file.getUrl();

      registrationSheet.getRange(i + 1, slipCol + 1).setValue(fileUrl);
      registrationSheet.getRange(i + 1, statusCol + 1).setValue('รอตรวจสอบ');

      // ส่งอีเมลแจ้ง admin
      var subject = 'การลงทะเบียนใหม่ - ' + data[i][shopNameCol];
      var body = 'มีการลงทะเบียนใหม่\n' +
        'รหัส: ' + regId + '\n' +
        'ร้าน: ' + data[i][shopNameCol] + '\n' +
        'อีเมล: ' + data[i][emailCol] + '\n' +
        'รูปสลิป: ' + fileUrl;

      MailApp.sendEmail(ADMIN_EMAIL, subject, body);

      return { success: true, message: 'อัพโหลดสลิปสำเร็จ รอการตรวจสอบ' };
    }
  }

  return { success: false, message: 'ไม่พบรหัสลงทะเบียน' };
}

function approveRegistration(regId) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var registrationSheet = masterSheet.getSheetByName('การลงทะเบียน');
  var data = registrationSheet.getDataRange().getValues();
  var headers = data[0];

  var regIdCol = headers.indexOf('รหัสลงทะเบียน');
  var shopNameCol = headers.indexOf('ชื่อร้าน');
  var emailCol = headers.indexOf('อีเมล');
  var passwordCol = headers.indexOf('รหัสผ่าน');
  var packageCol = headers.indexOf('แพ็คเกจ');
  var statusCol = headers.indexOf('สถานะ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][regIdCol] === regId) {
      var expiryDays = data[i][packageCol] === 'รายเดือน' ? 30 : 365;

      var result = createNewShop(
        data[i][shopNameCol],
        data[i][emailCol],
        data[i][passwordCol],
        data[i][packageCol],
        expiryDays
      );

      if (result.success) {
        registrationSheet.getRange(i + 1, statusCol + 1).setValue('เปิดใช้งานแล้ว');
        return { success: true, message: 'เปิดใช้งานร้านสำเร็จ' };
      }
    }
  }

  return { success: false, message: 'ไม่พบรหัสลงทะเบียน' };
}

function getPendingRegistrations() {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var registrationSheet = masterSheet.getSheetByName('การลงทะเบียน');
  var data = registrationSheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var value = data[i][j];
      if (value instanceof Date) {
        value = value.toISOString();
      }
      row[headers[j]] = value;
    }
    if (row['สถานะ'] === 'รอตรวจสอบ') {
      result.push(row);
    }
  }

  return result;
}

// ===== PRODUCTS =====
function getProducts(sheetId) {
  var cacheKey = 'products_' + sheetId;
  var cached = getCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.openById(sheetId);
  var products = getProductsFromSheet(ss);
  setCache(cacheKey, products, 300);

  return products;
}

function getProductsFromSheet(ss) {
  var sheet = ss.getSheetByName('สินค้า');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var product = {};
    for (var j = 0; j < headers.length; j++) {
      var value = data[i][j];
      if (headers[j] === 'วัตถุดิบ' || headers[j] === 'ตัวเลือก') {
        try {
          product[headers[j]] = value ? JSON.parse(value) : [];
        } catch (e) {
          product[headers[j]] = [];
        }
      } else {
        product[headers[j]] = value;
      }
    }
    if (product['รหัสสินค้า']) {
      result.push(product);
    }
  }

  return result;
}

function addProduct(sheetId, productData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('สินค้า');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var productId = 'PROD' + new Date().getTime();
  var row = [];

  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];
    if (key === 'รหัสสินค้า') {
      row.push(productId);
    } else if (key === 'วัตถุดิบ' || key === 'ตัวเลือก') {
      row.push(JSON.stringify(productData[key] || []));
    } else {
      row.push(productData[key] || '');
    }
  }

  sheet.appendRow(row);
  clearCache('products_' + sheetId);

  return { success: true, productId: productId };
}

function updateProduct(sheetId, productId, productData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('สินค้า');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสสินค้า');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === productId) {
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        if (key !== 'รหัสสินค้า' && productData.hasOwnProperty(key)) {
          var value = productData[key];
          if (key === 'วัตถุดิบ' || key === 'ตัวเลือก') {
            value = JSON.stringify(value);
          }
          sheet.getRange(i + 1, j + 1).setValue(value);
        }
      }
      clearCache('products_' + sheetId);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบสินค้า' };
}

function deleteProduct(sheetId, productId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('สินค้า');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสสินค้า');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === productId) {
      sheet.deleteRow(i + 1);
      clearCache('products_' + sheetId);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบสินค้า' };
}

// ===== CATEGORIES =====
function getCategories(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('หมวดหมู่');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var category = {};
    for (var j = 0; j < headers.length; j++) {
      category[headers[j]] = data[i][j];
    }
    if (category['รหัสหมวดหมู่']) {
      result.push(category);
    }
  }

  return result;
}

function addCategory(sheetId, name) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('หมวดหมู่');
  var categoryId = 'CAT' + new Date().getTime();
  var lastRow = sheet.getLastRow();

  sheet.appendRow([categoryId, name, lastRow, 'ใช้งาน']);
  return { success: true, categoryId: categoryId };
}

function updateCategory(sheetId, categoryId, name) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('หมวดหมู่');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสหมวดหมู่');
  var nameCol = headers.indexOf('ชื่อหมวดหมู่');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === categoryId) {
      sheet.getRange(i + 1, nameCol + 1).setValue(name);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบหมวดหมู่' };
}

function deleteCategory(sheetId, categoryId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('หมวดหมู่');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสหมวดหมู่');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === categoryId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบหมวดหมู่' };
}

// ===== PRODUCT OPTIONS =====
function getProductOptions(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var option = {};
    for (var j = 0; j < headers.length; j++) {
      var value = data[i][j];
      if (headers[j] === 'วัตถุดิบ') {
        try {
          option[headers[j]] = value ? JSON.parse(value) : [];
        } catch (e) {
          option[headers[j]] = [];
        }
      } else {
        option[headers[j]] = value;
      }
    }
    if (option['รหัสตัวเลือก']) {
      result.push(option);
    }
  }

  return result;
}

function addProductOption(sheetId, optionData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  var optionId = 'OPT' + new Date().getTime();
  var row = [];

  for (var i = 0; i < headers.length; i++) {
    var key = headers[i];
    if (key === 'รหัสตัวเลือก') {
      row.push(optionId);
    } else if (key === 'วัตถุดิบ') {
      row.push(JSON.stringify(optionData[key] || []));
    } else {
      row.push(optionData[key] || '');
    }
  }

  sheet.appendRow(row);
  return { success: true, optionId: optionId };
}

function updateProductOption(sheetId, optionId, optionData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสตัวเลือก');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === optionId) {
      for (var j = 0; j < headers.length; j++) {
        var key = headers[j];
        if (key !== 'รหัสตัวเลือก' && optionData.hasOwnProperty(key)) {
          var value = optionData[key];
          if (key === 'วัตถุดิบ') {
            value = JSON.stringify(value);
          }
          sheet.getRange(i + 1, j + 1).setValue(value);
        }
      }
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบตัวเลือก' };
}

function deleteProductOption(sheetId, optionId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ตัวเลือกสินค้า');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสตัวเลือก');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === optionId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบตัวเลือก' };
}

// ===== MATERIALS =====
function getMaterials(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('วัตถุดิบ');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var material = {};
    for (var j = 0; j < headers.length; j++) {
      material[headers[j]] = data[i][j];
    }
    if (material['รหัสวัตถุดิบ']) {
      result.push(material);
    }
  }

  return result;
}

function addMaterial(sheetId, materialData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('วัตถุดิบ');

  var materialId = 'MAT' + new Date().getTime();

  sheet.appendRow([
    materialId,
    materialData.name,
    materialData.unit,
    materialData.pricePerUnit,
    materialData.quantity || 0,
    materialData.minQuantity || 0,
    'ใช้งาน'
  ]);

  return { success: true, materialId: materialId };
}

function updateMaterial(sheetId, materialId, materialData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('วัตถุดิบ');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสวัตถุดิบ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === materialId) {
      for (var key in materialData) {
        var col = headers.indexOf(key);
        if (col !== -1) {
          sheet.getRange(i + 1, col + 1).setValue(materialData[key]);
        }
      }
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบวัตถุดิบ' };
}

function deleteMaterial(sheetId, materialId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('วัตถุดิบ');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสวัตถุดิบ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === materialId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบวัตถุดิบ' };
}

function purchaseMaterial(sheetId, purchaseData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var lotsSheet = ss.getSheetByName('ล็อตวัตถุดิบ');
  var materialsSheet = ss.getSheetByName('วัตถุดิบ');

  var lotId = 'LOT' + new Date().getTime();

  // บันทึกล็อตใหม่
  lotsSheet.appendRow([
    lotId,
    purchaseData.materialId,
    purchaseData.quantity,
    purchaseData.totalPrice,
    new Date(),
    purchaseData.expiryDate ? new Date(purchaseData.expiryDate) : '',
    purchaseData.slipUrl || '',
    purchaseData.quantity
  ]);

  // อัพเดทจำนวนคงเหลือในวัตถุดิบ
  var data = materialsSheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf('รหัสวัตถุดิบ');
  var qtyCol = headers.indexOf('จำนวนคงเหลือ');
  var priceCol = headers.indexOf('ราคาต่อหน่วย');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === purchaseData.materialId) {
      var newQty = data[i][qtyCol] + purchaseData.quantity;
      var newPrice = purchaseData.totalPrice / purchaseData.quantity;
      materialsSheet.getRange(i + 1, qtyCol + 1).setValue(newQty);
      materialsSheet.getRange(i + 1, priceCol + 1).setValue(newPrice);
      break;
    }
  }

  return { success: true, lotId: lotId };
}

function addWaste(sheetId, wasteData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var wasteSheet = ss.getSheetByName('ของเสีย');
  var materialsSheet = ss.getSheetByName('วัตถุดิบ');

  // บันทึกของเสีย
  wasteSheet.appendRow([
    new Date(),
    wasteData.materialId,
    wasteData.quantity,
    wasteData.reason,
    wasteData.note || ''
  ]);

  // ลดจำนวนคงเหลือ
  var data = materialsSheet.getDataRange().getValues();
  var headers = data[0];
  var idCol = headers.indexOf('รหัสวัตถุดิบ');
  var qtyCol = headers.indexOf('จำนวนคงเหลือ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === wasteData.materialId) {
      var newQty = data[i][qtyCol] - wasteData.quantity;
      materialsSheet.getRange(i + 1, qtyCol + 1).setValue(Math.max(0, newQty));
      break;
    }
  }

  return { success: true };
}

function getExpiringMaterials(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var lotsSheet = ss.getSheetByName('ล็อตวัตถุดิบ');
  var materialsSheet = ss.getSheetByName('วัตถุดิบ');
  var lotsData = lotsSheet.getDataRange().getValues();
  var materialsData = materialsSheet.getDataRange().getValues();

  var lotsHeaders = lotsData[0];
  var materialsHeaders = materialsData[0];

  var materialIdCol = lotsHeaders.indexOf('รหัสวัตถุดิบ');
  var expiryCol = lotsHeaders.indexOf('วันหมดอายุ');
  var remainingCol = lotsHeaders.indexOf('จำนวนคงเหลือ');

  var materialNameCol = materialsHeaders.indexOf('ชื่อวัตถุดิบ');
  var materialIdMatCol = materialsHeaders.indexOf('รหัสวัตถุดิบ');

  var result = [];
  var today = new Date();
  var warningDate = new Date();
  warningDate.setDate(today.getDate() + 7);

  for (var i = 1; i < lotsData.length; i++) {
    var expiryDate = lotsData[i][expiryCol];
    if (expiryDate && expiryDate <= warningDate && lotsData[i][remainingCol] > 0) {
      var materialName = '';
      for (var j = 1; j < materialsData.length; j++) {
        if (materialsData[j][materialIdMatCol] === lotsData[i][materialIdCol]) {
          materialName = materialsData[j][materialNameCol];
          break;
        }
      }
      result.push({
        materialId: lotsData[i][materialIdCol],
        materialName: materialName,
        expiryDate: expiryDate.toISOString(),
        remaining: lotsData[i][remainingCol]
      });
    }
  }

  return result;
}

// ===== SALES =====
function createOrder(sheetId, orderData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var queueSheet = ss.getSheetByName('คิว');
  var settingsSheet = ss.getSheetByName('การตั้งค่า');

  // สร้างรหัสออเดอร์
  var orderId = generateOrderId(sheetId, orderData.channel);

  // คำนวณภาษี
  var settings = getSettings(sheetId);
  var subtotal = orderData.subtotal;
  var tax = 0;
  var total = subtotal;

  if (settings['เปิดภาษี'] === 'true') {
    if (settings['ประเภทภาษี'] === 'รวมในราคา') {
      tax = subtotal - (subtotal / 1.07);
    } else {
      tax = subtotal * 0.07;
      total = subtotal + tax;
    }
  }

  total = total - (orderData.discount || 0);

  // บันทึกการขาย
  salesSheet.appendRow([
    orderId,
    new Date(),
    Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm:ss'),
    orderData.channel,
    JSON.stringify(orderData.items),
    subtotal,
    orderData.discount || 0,
    tax,
    total,
    JSON.stringify(orderData.payments),
    orderData.slipUrl || '',
    'รอทำ',
    orderData.note || ''
  ]);

  // คำนวณเวลาทำ
  var totalTime = 0;
  for (var i = 0; i < orderData.items.length; i++) {
    totalTime += (orderData.items[i].time || 3) * orderData.items[i].quantity;
  }

  // เพิ่มคิว
  var queueNumber = getNextQueueNumber(sheetId);
  queueSheet.appendRow([
    orderId,
    queueNumber,
    JSON.stringify(orderData.items),
    'รอทำ',
    new Date(),
    '',
    totalTime
  ]);

  // หักวัตถุดิบ
  deductMaterials(sheetId, orderData.items);

  return {
    success: true,
    orderId: orderId,
    queueNumber: queueNumber,
    total: total,
    tax: tax,
    estimatedTime: totalTime
  };
}

function generateOrderId(sheetId, channel) {
  var settings = getSettings(sheetId);
  var prefix = settings['Prefix เลขออเดอร์'] || 'ORD';
  var format = settings['รูปแบบเลขออเดอร์'] || 'PREFIX-YYYYMMDD-XXX';

  var today = new Date();
  var dateStr = Utilities.formatDate(today, 'Asia/Bangkok', 'yyyyMMdd');

  // นับจำนวนออเดอร์วันนี้
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var data = salesSheet.getDataRange().getValues();
  var headers = data[0];
  var dateCol = headers.indexOf('วันที่');

  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var orderDate = data[i][dateCol];
    if (orderDate instanceof Date) {
      var orderDateStr = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'yyyyMMdd');
      if (orderDateStr === dateStr) {
        count++;
      }
    }
  }

  var sequence = String(count + 1).padStart(3, '0');

  return prefix + '-' + dateStr + '-' + sequence;
}

function getNextQueueNumber(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var queueSheet = ss.getSheetByName('คิว');
  var data = queueSheet.getDataRange().getValues();
  var headers = data[0];
  var queueNumCol = headers.indexOf('หมายเลขคิว');
  var timeCol = headers.indexOf('เวลาสั่ง');

  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Bangkok', 'yyyyMMdd');

  var maxQueue = 0;
  for (var i = 1; i < data.length; i++) {
    var orderTime = data[i][timeCol];
    if (orderTime instanceof Date) {
      var orderDateStr = Utilities.formatDate(orderTime, 'Asia/Bangkok', 'yyyyMMdd');
      if (orderDateStr === todayStr) {
        var queueNum = parseInt(data[i][queueNumCol]) || 0;
        if (queueNum > maxQueue) {
          maxQueue = queueNum;
        }
      }
    }
  }

  return maxQueue + 1;
}

function deductMaterials(sheetId, items) {
  var ss = SpreadsheetApp.openById(sheetId);
  var materialsSheet = ss.getSheetByName('วัตถุดิบ');
  var lotsSheet = ss.getSheetByName('ล็อตวัตถุดิบ');

  var materialsData = materialsSheet.getDataRange().getValues();
  var materialsHeaders = materialsData[0];
  var matIdCol = materialsHeaders.indexOf('รหัสวัตถุดิบ');
  var matQtyCol = materialsHeaders.indexOf('จำนวนคงเหลือ');

  var lotsData = lotsSheet.getDataRange().getValues();
  var lotsHeaders = lotsData[0];
  var lotMatIdCol = lotsHeaders.indexOf('รหัสวัตถุดิบ');
  var lotExpiryCol = lotsHeaders.indexOf('วันหมดอายุ');
  var lotRemainingCol = lotsHeaders.indexOf('จำนวนคงเหลือ');

  // รวมวัตถุดิบที่ต้องใช้
  var materialsNeeded = {};

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var materials = item.materials || [];

    for (var j = 0; j < materials.length; j++) {
      var mat = materials[j];
      var totalQty = mat.quantity * item.quantity;

      if (!materialsNeeded[mat.id]) {
        materialsNeeded[mat.id] = 0;
      }
      materialsNeeded[mat.id] += totalQty;
    }
  }

  // หักจากล็อต FEFO
  for (var materialId in materialsNeeded) {
    var needed = materialsNeeded[materialId];

    // หาล็อตที่ใกล้หมดอายุที่สุด
    var lots = [];
    for (var i = 1; i < lotsData.length; i++) {
      if (lotsData[i][lotMatIdCol] === materialId && lotsData[i][lotRemainingCol] > 0) {
        lots.push({
          row: i + 1,
          expiry: lotsData[i][lotExpiryCol],
          remaining: lotsData[i][lotRemainingCol]
        });
      }
    }

    // เรียงตามวันหมดอายุ
    lots.sort(function(a, b) {
      return new Date(a.expiry) - new Date(b.expiry);
    });

    // หักจากล็อต
    for (var i = 0; i < lots.length && needed > 0; i++) {
      var lot = lots[i];
      var deduct = Math.min(lot.remaining, needed);
      lotsSheet.getRange(lot.row, lotRemainingCol + 1).setValue(lot.remaining - deduct);
      needed -= deduct;
    }

    // อัพเดทจำนวนรวมในวัตถุดิบ
    for (var i = 1; i < materialsData.length; i++) {
      if (materialsData[i][matIdCol] === materialId) {
        var newQty = materialsData[i][matQtyCol] - materialsNeeded[materialId];
        materialsSheet.getRange(i + 1, matQtyCol + 1).setValue(Math.max(0, newQty));
        break;
      }
    }
  }
}

function getTodaySales(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var data = salesSheet.getDataRange().getValues();
  var headers = data[0];

  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Bangkok', 'yyyyMMdd');

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var orderDate = data[i][headers.indexOf('วันที่')];
    if (orderDate instanceof Date) {
      var orderDateStr = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'yyyyMMdd');
      if (orderDateStr === todayStr) {
        var sale = {};
        for (var j = 0; j < headers.length; j++) {
          var value = data[i][j];
          if (value instanceof Date) {
            value = value.toISOString();
          } else if (headers[j] === 'รายการสินค้า' || headers[j] === 'การชำระเงิน') {
            try {
              value = JSON.parse(value);
            } catch (e) {
              value = [];
            }
          }
          sale[headers[j]] = value;
        }
        result.push(sale);
      }
    }
  }

  return result;
}

function getSalesByDateRange(sheetId, startDate, endDate) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var data = salesSheet.getDataRange().getValues();
  var headers = data[0];

  var start = new Date(startDate);
  start.setHours(0, 0, 0, 0);
  var end = new Date(endDate);
  end.setHours(23, 59, 59, 999);

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var orderDate = data[i][headers.indexOf('วันที่')];
    if (orderDate instanceof Date && orderDate >= start && orderDate <= end) {
      var sale = {};
      for (var j = 0; j < headers.length; j++) {
        var value = data[i][j];
        if (value instanceof Date) {
          value = value.toISOString();
        } else if (headers[j] === 'รายการสินค้า' || headers[j] === 'การชำระเงิน') {
          try {
            value = JSON.parse(value);
          } catch (e) {
            value = [];
          }
        }
        sale[headers[j]] = value;
      }
      result.push(sale);
    }
  }

  return result;
}

function updateOrderStatus(sheetId, orderId, status) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var queueSheet = ss.getSheetByName('คิว');

  // อัพเดทสถานะในการขาย
  var salesData = salesSheet.getDataRange().getValues();
  var salesHeaders = salesData[0];
  var orderIdCol = salesHeaders.indexOf('รหัสออเดอร์');
  var statusCol = salesHeaders.indexOf('สถานะ');

  for (var i = 1; i < salesData.length; i++) {
    if (salesData[i][orderIdCol] === orderId) {
      salesSheet.getRange(i + 1, statusCol + 1).setValue(status);
      break;
    }
  }

  // อัพเดทสถานะในคิว
  var queueData = queueSheet.getDataRange().getValues();
  var queueHeaders = queueData[0];
  var qOrderIdCol = queueHeaders.indexOf('รหัสออเดอร์');
  var qStatusCol = queueHeaders.indexOf('สถานะ');
  var qFinishCol = queueHeaders.indexOf('เวลาเสร็จ');

  for (var i = 1; i < queueData.length; i++) {
    if (queueData[i][qOrderIdCol] === orderId) {
      queueSheet.getRange(i + 1, qStatusCol + 1).setValue(status);
      if (status === 'พร้อมเสิร์ฟ' || status === 'เสร็จสิ้น') {
        queueSheet.getRange(i + 1, qFinishCol + 1).setValue(new Date());
      }
      break;
    }
  }

  return { success: true };
}

function deleteOrder(sheetId, orderId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var queueSheet = ss.getSheetByName('คิว');

  // ลบจากการขาย
  var salesData = salesSheet.getDataRange().getValues();
  var salesHeaders = salesData[0];
  var orderIdCol = salesHeaders.indexOf('รหัสออเดอร์');

  for (var i = salesData.length - 1; i >= 1; i--) {
    if (salesData[i][orderIdCol] === orderId) {
      salesSheet.deleteRow(i + 1);
      break;
    }
  }

  // ลบจากคิว
  var queueData = queueSheet.getDataRange().getValues();
  var queueHeaders = queueData[0];
  var qOrderIdCol = queueHeaders.indexOf('รหัสออเดอร์');

  for (var i = queueData.length - 1; i >= 1; i--) {
    if (queueData[i][qOrderIdCol] === orderId) {
      queueSheet.deleteRow(i + 1);
      break;
    }
  }

  return { success: true };
}

function getOrderDetails(sheetId, orderId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var data = salesSheet.getDataRange().getValues();
  var headers = data[0];

  var orderIdCol = headers.indexOf('รหัสออเดอร์');

  for (var i = 1; i < data.length; i++) {
    if (data[i][orderIdCol] === orderId) {
      var order = {};
      for (var j = 0; j < headers.length; j++) {
        var value = data[i][j];
        if (value instanceof Date) {
          value = value.toISOString();
        } else if (headers[j] === 'รายการสินค้า' || headers[j] === 'การชำระเงิน') {
          try {
            value = JSON.parse(value);
          } catch (e) {
            value = [];
          }
        }
        order[headers[j]] = value;
      }
      return order;
    }
  }

  return null;
}

// ===== QUEUE =====
function getActiveQueue(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var queueSheet = ss.getSheetByName('คิว');
  var data = queueSheet.getDataRange().getValues();
  var headers = data[0];

  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Bangkok', 'yyyyMMdd');

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var orderTime = data[i][headers.indexOf('เวลาสั่ง')];
    if (orderTime instanceof Date) {
      var orderDateStr = Utilities.formatDate(orderTime, 'Asia/Bangkok', 'yyyyMMdd');
      var status = data[i][headers.indexOf('สถานะ')];

      if (orderDateStr === todayStr && (status === 'รอทำ' || status === 'กำลังทำ')) {
        var queue = {};
        for (var j = 0; j < headers.length; j++) {
          var value = data[i][j];
          if (value instanceof Date) {
            value = value.toISOString();
          } else if (headers[j] === 'รายการ') {
            try {
              value = JSON.parse(value);
            } catch (e) {
              value = [];
            }
          }
          queue[headers[j]] = value;
        }
        result.push(queue);
      }
    }
  }

  return result;
}

function getQueueStatus(sheetId, orderId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var queueSheet = ss.getSheetByName('คิว');
  var data = queueSheet.getDataRange().getValues();
  var headers = data[0];

  var orderIdCol = headers.indexOf('รหัสออเดอร์');

  for (var i = 1; i < data.length; i++) {
    if (data[i][orderIdCol] === orderId) {
      var queue = {};
      for (var j = 0; j < headers.length; j++) {
        var value = data[i][j];
        if (value instanceof Date) {
          value = value.toISOString();
        } else if (headers[j] === 'รายการ') {
          try {
            value = JSON.parse(value);
          } catch (e) {
            value = [];
          }
        }
        queue[headers[j]] = value;
      }
      return queue;
    }
  }

  return null;
}

// ===== COSTS =====
function addDailyCost(sheetId, costData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ต้นทุนรายวัน');

  sheet.appendRow([
    new Date(costData.date),
    costData.item,
    costData.amount,
    costData.category,
    costData.note || ''
  ]);

  return { success: true };
}

function addMonthlyCost(sheetId, costData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ต้นทุนรายเดือน');

  sheet.appendRow([
    costData.month,
    costData.item,
    costData.amount,
    costData.category,
    costData.note || ''
  ]);

  return { success: true };
}

function getDailyCosts(sheetId, month) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ต้นทุนรายวัน');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var date = data[i][headers.indexOf('วันที่')];
    if (date instanceof Date) {
      var dateMonth = Utilities.formatDate(date, 'Asia/Bangkok', 'yyyy-MM');
      if (dateMonth === month) {
        var cost = {};
        for (var j = 0; j < headers.length; j++) {
          var value = data[i][j];
          if (value instanceof Date) {
            value = value.toISOString();
          }
          cost[headers[j]] = value;
        }
        result.push(cost);
      }
    }
  }

  return result;
}

function getMonthlyCosts(sheetId, year) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ต้นทุนรายเดือน');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var monthYear = data[i][headers.indexOf('เดือน/ปี')];
    if (monthYear && monthYear.indexOf(year) !== -1) {
      var cost = {};
      for (var j = 0; j < headers.length; j++) {
        cost[headers[j]] = data[i][j];
      }
      result.push(cost);
    }
  }

  return result;
}

// ===== SETTINGS =====
function getSettings(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('การตั้งค่า');
  var data = sheet.getDataRange().getValues();

  var settings = {};
  for (var i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }

  return settings;
}

function updateSettings(sheetId, newSettings) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('การตั้งค่า');
  var data = sheet.getDataRange().getValues();

  for (var key in newSettings) {
    var found = false;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(newSettings[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, newSettings[key]]);
    }
  }

  return { success: true };
}

function getOrderChannels(sheetId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ช่องทางสั่งซื้อ');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var channel = {};
    for (var j = 0; j < headers.length; j++) {
      channel[headers[j]] = data[i][j];
    }
    if (channel['รหัสช่องทาง']) {
      result.push(channel);
    }
  }

  return result;
}

function addOrderChannel(sheetId, channelData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ช่องทางสั่งซื้อ');

  var channelId = 'CH' + new Date().getTime();

  sheet.appendRow([
    channelId,
    channelData.name,
    channelData.orderType,
    'ใช้งาน'
  ]);

  return { success: true, channelId: channelId };
}

function updateOrderChannel(sheetId, channelId, channelData) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ช่องทางสั่งซื้อ');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสช่องทาง');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === channelId) {
      for (var key in channelData) {
        var col = headers.indexOf(key);
        if (col !== -1) {
          sheet.getRange(i + 1, col + 1).setValue(channelData[key]);
        }
      }
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบช่องทาง' };
}

function deleteOrderChannel(sheetId, channelId) {
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('ช่องทางสั่งซื้อ');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var idCol = headers.indexOf('รหัสช่องทาง');

  for (var i = 1; i < data.length; i++) {
    if (data[i][idCol] === channelId) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบช่องทาง' };
}

// ===== DASHBOARD & REPORTS =====
function getDashboardData(sheetId) {
  var cacheKey = 'dashboard_' + sheetId;
  var cached = getCache(cacheKey);
  if (cached) return cached;

  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var salesData = salesSheet.getDataRange().getValues();
  var salesHeaders = salesData[0];

  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Bangkok', 'yyyyMMdd');

  // ยอดขายวันนี้
  var todaySales = 0;
  var todayCost = 0;
  var productsSold = {};

  var dateCol = salesHeaders.indexOf('วันที่');
  var totalCol = salesHeaders.indexOf('ยอดสุทธิ');
  var itemsCol = salesHeaders.indexOf('รายการสินค้า');

  for (var i = 1; i < salesData.length; i++) {
    var orderDate = salesData[i][dateCol];
    if (orderDate instanceof Date) {
      var orderDateStr = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'yyyyMMdd');
      if (orderDateStr === todayStr) {
        todaySales += salesData[i][totalCol] || 0;

        var items = [];
        try {
          items = JSON.parse(salesData[i][itemsCol]);
        } catch (e) {}

        for (var j = 0; j < items.length; j++) {
          var item = items[j];
          if (!productsSold[item.name]) {
            productsSold[item.name] = { quantity: 0, revenue: 0 };
          }
          productsSold[item.name].quantity += item.quantity;
          productsSold[item.name].revenue += item.total;
          todayCost += (item.cost || 0) * item.quantity;
        }
      }
    }
  }

  // 5 สินค้าขายดี
  var topProducts = [];
  for (var name in productsSold) {
    topProducts.push({
      name: name,
      quantity: productsSold[name].quantity,
      revenue: productsSold[name].revenue
    });
  }
  topProducts.sort(function(a, b) {
    return b.quantity - a.quantity;
  });
  topProducts = topProducts.slice(0, 5);

  // ยอดขาย 7 วันล่าสุด
  var last7Days = [];
  for (var d = 6; d >= 0; d--) {
    var date = new Date();
    date.setDate(date.getDate() - d);
    var dateStr = Utilities.formatDate(date, 'Asia/Bangkok', 'yyyyMMdd');
    var dayLabel = Utilities.formatDate(date, 'Asia/Bangkok', 'dd/MM');

    var dayTotal = 0;
    for (var i = 1; i < salesData.length; i++) {
      var orderDate = salesData[i][dateCol];
      if (orderDate instanceof Date) {
        var orderDateStr = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'yyyyMMdd');
        if (orderDateStr === dateStr) {
          dayTotal += salesData[i][totalCol] || 0;
        }
      }
    }

    last7Days.push({ date: dayLabel, total: dayTotal });
  }

  // วัตถุดิบใกล้หมดอายุ
  var expiringMaterials = getExpiringMaterials(sheetId);

  var result = {
    todaySales: todaySales,
    todayCost: todayCost,
    profit: todaySales - todayCost,
    topProducts: topProducts,
    last7Days: last7Days,
    expiringMaterials: expiringMaterials
  };

  setCache(cacheKey, result, 60);

  return result;
}

function getSalesReport(sheetId, startDate, endDate) {
  var sales = getSalesByDateRange(sheetId, startDate, endDate);

  var totalSales = 0;
  var totalOrders = sales.length;
  var productsSold = {};
  var dailySales = {};

  for (var i = 0; i < sales.length; i++) {
    var sale = sales[i];
    totalSales += sale['ยอดสุทธิ'] || 0;

    var date = new Date(sale['วันที่']);
    var dateStr = Utilities.formatDate(date, 'Asia/Bangkok', 'dd/MM/yyyy');

    if (!dailySales[dateStr]) {
      dailySales[dateStr] = 0;
    }
    dailySales[dateStr] += sale['ยอดสุทธิ'] || 0;

    var items = sale['รายการสินค้า'] || [];
    for (var j = 0; j < items.length; j++) {
      var item = items[j];
      if (!productsSold[item.name]) {
        productsSold[item.name] = { quantity: 0, revenue: 0 };
      }
      productsSold[item.name].quantity += item.quantity;
      productsSold[item.name].revenue += item.total;
    }
  }

  var topProducts = [];
  for (var name in productsSold) {
    topProducts.push({
      name: name,
      quantity: productsSold[name].quantity,
      revenue: productsSold[name].revenue
    });
  }
  topProducts.sort(function(a, b) {
    return b.revenue - a.revenue;
  });

  return {
    totalSales: totalSales,
    totalOrders: totalOrders,
    averageOrder: totalOrders > 0 ? totalSales / totalOrders : 0,
    dailySales: dailySales,
    topProducts: topProducts
  };
}

function getMonthlySalesData(sheetId, year) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var data = salesSheet.getDataRange().getValues();
  var headers = data[0];

  var dateCol = headers.indexOf('วันที่');
  var totalCol = headers.indexOf('ยอดสุทธิ');

  var monthlySales = {};
  for (var m = 1; m <= 12; m++) {
    var monthKey = m < 10 ? '0' + m : '' + m;
    monthlySales[monthKey] = 0;
  }

  for (var i = 1; i < data.length; i++) {
    var orderDate = data[i][dateCol];
    if (orderDate instanceof Date) {
      var orderYear = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'yyyy');
      if (orderYear === year) {
        var orderMonth = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'MM');
        monthlySales[orderMonth] += data[i][totalCol] || 0;
      }
    }
  }

  return monthlySales;
}

function getDailySalesInMonth(sheetId, yearMonth) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var data = salesSheet.getDataRange().getValues();
  var headers = data[0];

  var dateCol = headers.indexOf('วันที่');
  var totalCol = headers.indexOf('ยอดสุทธิ');

  var dailySales = {};

  for (var i = 1; i < data.length; i++) {
    var orderDate = data[i][dateCol];
    if (orderDate instanceof Date) {
      var orderYearMonth = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'yyyy-MM');
      if (orderYearMonth === yearMonth) {
        var day = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'dd');
        if (!dailySales[day]) {
          dailySales[day] = 0;
        }
        dailySales[day] += data[i][totalCol] || 0;
      }
    }
  }

  return dailySales;
}

function getProfitLossReport(sheetId, month) {
  var ss = SpreadsheetApp.openById(sheetId);
  var salesSheet = ss.getSheetByName('การขาย');
  var salesData = salesSheet.getDataRange().getValues();
  var salesHeaders = salesData[0];

  var dateCol = salesHeaders.indexOf('วันที่');
  var totalCol = salesHeaders.indexOf('ยอดสุทธิ');
  var itemsCol = salesHeaders.indexOf('รายการสินค้า');

  var totalRevenue = 0;
  var totalProductCost = 0;

  for (var i = 1; i < salesData.length; i++) {
    var orderDate = salesData[i][dateCol];
    if (orderDate instanceof Date) {
      var orderMonth = Utilities.formatDate(orderDate, 'Asia/Bangkok', 'yyyy-MM');
      if (orderMonth === month) {
        totalRevenue += salesData[i][totalCol] || 0;

        var items = [];
        try {
          items = JSON.parse(salesData[i][itemsCol]);
        } catch (e) {}

        for (var j = 0; j < items.length; j++) {
          totalProductCost += (items[j].cost || 0) * items[j].quantity;
        }
      }
    }
  }

  // ต้นทุนรายวัน
  var dailyCosts = getDailyCosts(sheetId, month);
  var totalDailyCost = 0;
  for (var i = 0; i < dailyCosts.length; i++) {
    totalDailyCost += dailyCosts[i]['จำนวนเงิน'] || 0;
  }

  // ต้นทุนรายเดือน
  var year = month.split('-')[0];
  var monthlyCosts = getMonthlyCosts(sheetId, year);
  var totalMonthlyCost = 0;
  for (var i = 0; i < monthlyCosts.length; i++) {
    if (monthlyCosts[i]['เดือน/ปี'] === month) {
      totalMonthlyCost += monthlyCosts[i]['จำนวนเงิน'] || 0;
    }
  }

  var totalCost = totalProductCost + totalDailyCost + totalMonthlyCost;
  var profit = totalRevenue - totalCost;

  return {
    revenue: totalRevenue,
    productCost: totalProductCost,
    dailyCost: totalDailyCost,
    monthlyCost: totalMonthlyCost,
    totalCost: totalCost,
    profit: profit,
    profitMargin: totalRevenue > 0 ? (profit / totalRevenue * 100) : 0
  };
}

// ===== FILE UPLOAD =====
function uploadSlipImage(sheetId, base64Data, fileName) {
  var ss = SpreadsheetApp.openById(sheetId);
  var settings = getSettings(sheetId);

  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var sheetIdCol = headers.indexOf('SheetID');
  var folderIdCol = headers.indexOf('FolderID');

  var folderId = '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][sheetIdCol] === sheetId) {
      folderId = data[i][folderIdCol];
      break;
    }
  }

  if (!folderId) {
    return { success: false, message: 'ไม่พบโฟลเดอร์ร้าน' };
  }

  var folder = DriveApp.getFolderById(folderId);
  var blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data.split(',')[1]),
    'image/png',
    fileName
  );
  var file = folder.createFile(blob);

  return { success: true, url: file.getUrl(), id: file.getId() };
}

// ===== SUPER ADMIN =====
function getAllShops() {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var result = [];
  for (var i = 1; i < data.length; i++) {
    var shop = {};
    for (var j = 0; j < headers.length; j++) {
      var value = data[i][j];
      if (value instanceof Date) {
        value = value.toISOString();
      }
      shop[headers[j]] = value;
    }
    if (shop['รหัสร้าน']) {
      result.push(shop);
    }
  }

  return result;
}

function updateShopLicense(shopId, newExpiryDate) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var shopIdCol = headers.indexOf('รหัสร้าน');
  var expiryCol = headers.indexOf('วันหมดอายุ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][shopIdCol] === shopId) {
      shopsSheet.getRange(i + 1, expiryCol + 1).setValue(new Date(newExpiryDate));
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบร้านค้า' };
}

function updateShopStatus(shopId, status) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var shopIdCol = headers.indexOf('รหัสร้าน');
  var statusCol = headers.indexOf('สถานะ');

  for (var i = 1; i < data.length; i++) {
    if (data[i][shopIdCol] === shopId) {
      shopsSheet.getRange(i + 1, statusCol + 1).setValue(status);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบร้านค้า' };
}

function deleteShop(shopId) {
  var masterSheet = SpreadsheetApp.openById(MASTER_SHEET_ID);
  var shopsSheet = masterSheet.getSheetByName('ร้านค้า');
  var data = shopsSheet.getDataRange().getValues();
  var headers = data[0];

  var shopIdCol = headers.indexOf('รหัสร้าน');

  for (var i = 1; i < data.length; i++) {
    if (data[i][shopIdCol] === shopId) {
      shopsSheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, message: 'ไม่พบร้านค้า' };
}

// ===== SAMPLE DATA =====
function createSampleShop() {
  var result = createNewShop(
    'ร้านกาแฟตัวอย่าง',
    'demo@example.com',
    'demo123',
    'รายเดือน',
    30
  );

  if (result.success) {
    // เพิ่มข้อมูลตัวอย่าง
    var sheetId = result.sheetId;

    // หมวดหมู่
    addCategory(sheetId, 'กาแฟ');
    addCategory(sheetId, 'ชา');
    addCategory(sheetId, 'น้ำผลไม้');

    // วัตถุดิบ
    addMaterial(sheetId, { name: 'เมล็ดกาแฟ', unit: 'กรัม', pricePerUnit: 0.5, quantity: 1000, minQuantity: 200 });
    addMaterial(sheetId, { name: 'นม', unit: 'มล.', pricePerUnit: 0.05, quantity: 5000, minQuantity: 1000 });
    addMaterial(sheetId, { name: 'น้ำเชื่อม', unit: 'กรัม', pricePerUnit: 0.03, quantity: 2000, minQuantity: 500 });
    addMaterial(sheetId, { name: 'น้ำแข็ง', unit: 'กรัม', pricePerUnit: 0.01, quantity: 10000, minQuantity: 2000 });

    Logger.log('Sample shop created with ID: ' + result.shopId);
  }

  return result;
}
