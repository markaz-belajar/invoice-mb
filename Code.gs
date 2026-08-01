/**
 * ╔══════════════════════════════════════════════════════════╗
 * ║       Markaz Belajar — Google Apps Script (Backend)      ║
 * ║   Kompatibel dengan invoice-mb.html (versi terbaru)      ║
 * ╚══════════════════════════════════════════════════════════╝
 *
 * CARA DEPLOY:
 *  1. Buka https://script.google.com → paste kode ini
 *  2. Deploy → Manage Deployments → Edit (✏️) → New version → Deploy
 *  3. URL tidak berubah jika edit deployment yang sama
 */

// ═══════════════════════════════════════════════════════════
//  CONFIG
// ═══════════════════════════════════════════════════════════

// ID folder Google Drive tempat bukti pembayaran disimpan
const FOLDER_ID = '1ONAXSPkfbW7TBONufyGEGTv-MmSvNDo5';

// ID Google Spreadsheet (dari URL spreadsheet)
const SHEET_ID = '1Jt2fv8NLgBxuzdXvw0ptSreCfGBlSuXvODcg5ia9CfA';

// ═══════════════════════════════════════════════════════════
//  HELPER
// ═══════════════════════════════════════════════════════════
function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _route(action, data) {
  switch (action) {
    case 'getInvoices':  return handleGetInvoices();
    case 'saveInvoice':  return handleSaveInvoice(data.invoice);
    case 'updateInvoice': return handleUpdateInvoice(data.backendId, data.updates, data.updatedBy);
    case 'deleteInvoice': return handleDeleteInvoice(data.backendId, data.deletedBy);
    case 'uploadProof':  return handleUploadProof(data.base64, data.mimeType, data.invoiceNumber);
    // Pengeluaran (expense.html)
    case 'getExpenses':   return handleGetExpenses();
    case 'saveExpense':   return handleSaveExpense(data.expense);
    case 'updateExpense': return handleUpdateExpense(data.id, data.updates);
    case 'deleteExpense': return handleDeleteExpense(data.id);
    // Kategori kustom (tersinkron tim)
    case 'getCategories': return handleGetCategories();
    case 'saveCategory':  return handleSaveCategory(data.category);
    default: return _json({ error: 'Action tidak dikenal: ' + action });
  }
}

// ═══════════════════════════════════════════════════════════
//  doGet — handle GET (termasuk POST-via-GET untuk bypass CORS)
// ═══════════════════════════════════════════════════════════
function doGet(e) {
  try {
    // POST-via-GET: data dikirim lewat query param (hindari CORS preflight)
    if (e.parameter.method === 'post' && e.parameter.data) {
      const data = JSON.parse(decodeURIComponent(e.parameter.data));
      return _route(data.action || '', data);
    }

    // GET biasa
    const action = e.parameter.action || '';
    if (action === 'getInvoices')   return handleGetInvoices();
    if (action === 'getExpenses')   return handleGetExpenses();
    if (action === 'getCategories') return handleGetCategories();

    // Health check
    return _json({ status: 'ok', message: 'Markaz Belajar GAS aktif ✓', timestamp: new Date().toISOString() });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  doPost — handle POST biasa
// ═══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    let data = {};

    if (e.parameter && e.parameter.data) {
      data = JSON.parse(decodeURIComponent(e.parameter.data));
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }

    return _route(data.action || e.parameter.action || '', data);

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HANDLER: getInvoices
// ═══════════════════════════════════════════════════════════
function handleGetInvoices() {
  try {
    const sheet = getInvoiceSheet();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return _json({ invoices: [] });

    const headers = rows[0];
    const invoices = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (typeof val === 'string' && (val.startsWith('{') || val.startsWith('['))) {
          try { val = JSON.parse(val); } catch (_) {}
        }
        obj[h] = val;
      });
      return obj;
    });

    return _json({ invoices });
  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HANDLER: saveInvoice
// ═══════════════════════════════════════════════════════════
function handleSaveInvoice(invoice) {
  try {
    if (!invoice) return _json({ error: 'Data invoice kosong.' });

    const sheet    = getInvoiceSheet();
    const headers  = getOrCreateHeaders(sheet);
    const backendId = invoice.__backendId || Utilities.getUuid();
    invoice.__backendId = backendId;

    const row = headers.map(h => {
      const val = invoice[h];
      if (val !== undefined && val !== null && typeof val === 'object') return JSON.stringify(val);
      return val !== undefined ? val : '';
    });

    sheet.appendRow(row);
    return _json({ ok: true, backendId });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HANDLER: updateInvoice
// ═══════════════════════════════════════════════════════════
function handleUpdateInvoice(backendId, updates, updatedBy) {
  try {
    if (!backendId || !updates) return _json({ error: 'Parameter tidak lengkap.' });

    const sheet   = getInvoiceSheet();
    const headers = getOrCreateHeaders(sheet);
    const rows    = sheet.getDataRange().getValues();
    const idCol   = headers.indexOf('__backendId');

    if (idCol < 0) return _json({ error: 'Kolom __backendId tidak ditemukan.' });

    for (let r = 1; r < rows.length; r++) {
      if (rows[r][idCol] === backendId) {
        Object.keys(updates).forEach(key => {
          const col = headers.indexOf(key);
          if (col >= 0) {
            let val = updates[key];
            if (val !== null && typeof val === 'object') val = JSON.stringify(val);
            sheet.getRange(r + 1, col + 1).setValue(val);
          }
        });
        return _json({ ok: true });
      }
    }

    return _json({ error: 'Invoice tidak ditemukan.' });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HANDLER: deleteInvoice
// ═══════════════════════════════════════════════════════════
function handleDeleteInvoice(backendId, deletedBy) {
  try {
    if (!backendId) return _json({ error: 'backendId tidak ada.' });

    const sheet   = getInvoiceSheet();
    const headers = getOrCreateHeaders(sheet);
    const rows    = sheet.getDataRange().getValues();
    const idCol   = headers.indexOf('__backendId');

    if (idCol < 0) return _json({ error: 'Kolom __backendId tidak ditemukan.' });

    for (let r = 1; r < rows.length; r++) {
      if (rows[r][idCol] === backendId) {
        sheet.deleteRow(r + 1);
        return _json({ ok: true });
      }
    }

    return _json({ error: 'Invoice tidak ditemukan.' });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HANDLER: uploadProof — FIX CORS
//  Menerima dari GET (POST-via-GET) maupun POST biasa
// ═══════════════════════════════════════════════════════════
function handleUploadProof(base64, mimeType, invoiceNumber) {
  try {
    if (!base64) return _json({ error: 'Data gambar kosong.' });

    const folder = DriveApp.getFolderById(FOLDER_ID);

    // Bersihkan prefix data URL jika ada
    const cleanBase64 = base64.replace(/^data:[^;]+;base64,/, '');
    const mime        = mimeType || 'image/jpeg';
    const ext         = mime.split('/')[1] || 'jpg';
    const filename    = 'bukti_' + (invoiceNumber || Date.now()) + '.' + ext;

    const blob = Utilities.newBlob(
      Utilities.base64Decode(cleanBase64),
      mime,
      filename
    );

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = file.getId();

    return _json({
      success:      true,
      fileId:       fileId,
      viewUrl:      'https://drive.google.com/file/d/' + fileId + '/view',
      thumbnailUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w600'
    });

  } catch (err) {
    return _json({ success: false, error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HANDLER: getExpenses / saveExpense / updateExpense / deleteExpense
//  (dipakai oleh expense.html — Pengeluaran Operasional)
// ═══════════════════════════════════════════════════════════
function handleGetExpenses() {
  try {
    const sheet = getExpenseSheet();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return _json({ expenses: [] });

    const headers = rows[0];
    const expenses = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let val = row[i];
        if (typeof val === 'string' && (val.startsWith('{') || val.startsWith('['))) {
          try { val = JSON.parse(val); } catch (_) {}
        }
        obj[h] = val;
      });
      return obj;
    });

    return _json({ expenses });
  } catch (err) {
    return _json({ error: err.toString() });
  }
}

function handleSaveExpense(expense) {
  try {
    if (!expense) return _json({ error: 'Data pengeluaran kosong.' });

    const sheet   = getExpenseSheet();
    const headers = getOrCreateExpenseHeaders(sheet);
    const id      = expense.__id || Utilities.getUuid();
    expense.__id  = id;

    // Normalisasi tanggal jadi string biar tidak digeser timezone oleh Sheets
    if (expense.date) expense.date = String(expense.date);

    const row = headers.map(h => {
      const val = expense[h];
      if (val !== undefined && val !== null && typeof val === 'object') return JSON.stringify(val);
      return val !== undefined ? val : '';
    });

    sheet.appendRow(row);
    return _json({ ok: true, id });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

function handleUpdateExpense(id, updates) {
  try {
    if (!id || !updates) return _json({ error: 'Parameter tidak lengkap.' });

    const sheet   = getExpenseSheet();
    const headers = getOrCreateExpenseHeaders(sheet);
    const rows    = sheet.getDataRange().getValues();
    const idCol   = headers.indexOf('__id');

    if (idCol < 0) return _json({ error: 'Kolom __id tidak ditemukan.' });

    for (let r = 1; r < rows.length; r++) {
      if (rows[r][idCol] === id) {
        Object.keys(updates).forEach(key => {
          const col = headers.indexOf(key);
          if (col >= 0) {
            let val = updates[key];
            if (val !== null && typeof val === 'object') val = JSON.stringify(val);
            sheet.getRange(r + 1, col + 1).setValue(val);
          }
        });
        const updatedCol = headers.indexOf('updated_at');
        if (updatedCol >= 0) sheet.getRange(r + 1, updatedCol + 1).setValue(new Date().toISOString());
        return _json({ ok: true });
      }
    }

    return _json({ error: 'Pengeluaran tidak ditemukan.' });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

function handleDeleteExpense(id) {
  try {
    if (!id) return _json({ error: 'id tidak ada.' });

    const sheet   = getExpenseSheet();
    const headers = getOrCreateExpenseHeaders(sheet);
    const rows    = sheet.getDataRange().getValues();
    const idCol   = headers.indexOf('__id');

    if (idCol < 0) return _json({ error: 'Kolom __id tidak ditemukan.' });

    for (let r = 1; r < rows.length; r++) {
      if (rows[r][idCol] === id) {
        sheet.deleteRow(r + 1);
        return _json({ ok: true });
      }
    }

    return _json({ error: 'Pengeluaran tidak ditemukan.' });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HANDLER: getCategories / saveCategory
//  Kategori kustom buatan tim, tersinkron lewat sheet "categories"
// ═══════════════════════════════════════════════════════════
function handleGetCategories() {
  try {
    const sheet = getCategoriesSheet();
    const rows  = sheet.getDataRange().getValues();
    if (rows.length <= 1) return _json({ categories: [] });

    const headers = rows[0];
    const categories = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    }).filter(c => c.key); // buang baris kosong

    return _json({ categories });
  } catch (err) {
    return _json({ error: err.toString() });
  }
}

function handleSaveCategory(category) {
  try {
    if (!category || !category.key) return _json({ error: 'Data kategori kosong.' });

    const sheet   = getCategoriesSheet();
    const headers = getOrCreateCategoryHeaders(sheet);

    // Cegah duplikat (case-insensitive)
    const rows  = sheet.getDataRange().getValues();
    const keyCol = headers.indexOf('key');
    for (let r = 1; r < rows.length; r++) {
      if (String(rows[r][keyCol]).toLowerCase() === String(category.key).toLowerCase()) {
        return _json({ ok: true, alreadyExists: true });
      }
    }

    const row = headers.map(h => {
      if (h === 'created_at') return new Date().toISOString();
      return category[h] !== undefined ? category[h] : '';
    });

    sheet.appendRow(row);
    return _json({ ok: true });

  } catch (err) {
    return _json({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════
//  HELPER: ambil/buat sheet "expenses"
// ═══════════════════════════════════════════════════════════
function getExpenseSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('expenses');
  if (!sheet) sheet = ss.insertSheet('expenses');
  return sheet;
}

function getOrCreateExpenseHeaders(sheet) {
  const lastCol  = sheet.getLastColumn();
  const firstRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  if (!firstRow || firstRow[0] === '') {
    const defaultHeaders = [
      '__id', 'date', 'category', 'description', 'amount',
      'paid_by', 'notes', 'created_at', 'updated_at'
    ];
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    const hRange = sheet.getRange(1, 1, 1, defaultHeaders.length);
    hRange.setBackground('#1B6CA8');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    return defaultHeaders;
  }

  return firstRow.filter(h => h !== '');
}

// ═══════════════════════════════════════════════════════════
//  HELPER: ambil/buat sheet "categories" (kategori kustom tim)
// ═══════════════════════════════════════════════════════════
function getCategoriesSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('categories');
  if (!sheet) sheet = ss.insertSheet('categories');
  return sheet;
}

function getOrCreateCategoryHeaders(sheet) {
  const lastCol  = sheet.getLastColumn();
  const firstRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  if (!firstRow || firstRow[0] === '') {
    const defaultHeaders = ['key', 'color', 'bg', 'icon', 'created_at'];
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    const hRange = sheet.getRange(1, 1, 1, defaultHeaders.length);
    hRange.setBackground('#1B6CA8');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    return defaultHeaders;
  }

  return firstRow.filter(h => h !== '');
}

// ═══════════════════════════════════════════════════════════
//  HELPER: ambil/buat sheet "invoices"
// ═══════════════════════════════════════════════════════════
function getInvoiceSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName('invoices');
  if (!sheet) sheet = ss.insertSheet('invoices');
  return sheet;
}

// ═══════════════════════════════════════════════════════════
//  HELPER: ambil/buat header row
// ═══════════════════════════════════════════════════════════
function getOrCreateHeaders(sheet) {
  const lastCol  = sheet.getLastColumn();
  const firstRow = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];

  if (!firstRow || firstRow[0] === '') {
    const defaultHeaders = [
      '__backendId', 'invoice_number', 'inv_date', 'due_date',
      'student_name', 'parent_name', 'phone', 'student_id', 'level', 'address',
      'programs', 'reg_fee', 'total', 'status',
      'payment_date', 'payment_method', 'payment_proof', 'payment_note',
      'notes', 'created_by', 'created_at', 'updated_at'
    ];
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
    const hRange = sheet.getRange(1, 1, 1, defaultHeaders.length);
    hRange.setBackground('#1B6CA8');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    return defaultHeaders;
  }

  return firstRow.filter(h => h !== '');
}
