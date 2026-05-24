// =============================
// HAPUS SEMUA DUPLIKAT (INCLUDE TERSEMBUNYI)
// =============================
function removeAllDuplicateLinks() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Exam_Links');
    if (!sheet) return 'Sheet Exam_Links tidak ditemukan';

    // Unhide all rows first
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'Tidak ada data';
    if (sheet.getMaxRows() > lastRow) {
      sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);
    }
    // Unhide all rows (if any are hidden)
    if (sheet.getRowGroupDepth(2) > 0) {
      sheet.expandAllRowGroups();
    }
    // Unhide all columns (if any are hidden)
    if (sheet.getColumnGroupDepth(1) > 0) {
      sheet.expandAllColumnGroups();
    }

    // Remove filter views (if any)
    if (sheet.getFilter()) {
      sheet.getFilter().remove();
    }

    // Remove duplicates by URL (kolom D/4)
    const data = sheet.getDataRange().getValues();
    const uniqueUrls = new Map();
    const rowsToDelete = [];
    for (let i = data.length - 1; i >= 1; i--) {
      const url = data[i][3];
      if (!url || url.toString().trim() === '') {
        rowsToDelete.push(i + 1);
        continue;
      }
      if (uniqueUrls.has(url)) {
        rowsToDelete.push(i + 1);
      } else {
        uniqueUrls.set(url, i);
      }
    }
    rowsToDelete.sort((a, b) => b - a);
    for (const rowIndex of rowsToDelete) {
      sheet.deleteRow(rowIndex);
    }
    // Update nomor urut setelah menghapus duplikat
    const updatedData = sheet.getDataRange().getValues();
    for (let i = 1; i < updatedData.length; i++) {
      sheet.getRange(i + 1, 2).setValue(i); // Update kolom Nomor
    }
    clearLinksCache();
    return `Selesai hapus ${rowsToDelete.length} baris duplikat/kosong (termasuk tersembunyi)`;
  } catch (error) {
    return 'Error: ' + error;
  }
}
/**
 * EXAMUNA - GOOGLE APPS SCRIPT BACKEND
 * 
 * Fitur:
 * - Generate QR Code untuk link ujian
 * - Monitoring session siswa
 * - Tracking violations
 * - Session management
 * - Reports & Analytics
 * - Auto-save link ujian ke Spreadsheet (HAPUS DUPLIKAT otomatis)
 * - API untuk dropdown daftar ujian
 */

// ===========================================
// KONFIGURASI
// ===========================================

const CONFIG = {
  SPREADSHEET_ID: '1li0ycYSk1Y2d19bMfx9KUqJrUnUDszoPU3iCi1T9IhY',
  SCHOOL_NAME: 'MGMP PAI Kab Bangkalan',
  APP_NAME: 'EXAMUNA',
  QR_API: 'https://api.qrserver.com/v1/create-qr-code/'
};

// NOTE: adminKey checks removed — API operations do not require an admin key.

// ===========================================
// CACHING LAYER - OPTIMASI PERFORMA
// ===========================================

const CACHE_CONFIG = {
  EXPIRY_MS: 5 * 60 * 1000 // 5 menit cache validity
};

function getCachedLinks() {
  try {
    const cache = PropertiesService.getUserProperties();
    const cached = cache.getProperty('exam_links_cache');
    const cacheTime = cache.getProperty('exam_links_cache_time');
    
    if (cached && cacheTime) {
      const age = Date.now() - parseInt(cacheTime);
      if (age < CACHE_CONFIG.EXPIRY_MS) {
        Logger.log('Cache HIT - Links returned from cache');
        return JSON.parse(cached);
      }
    }
    return null;
  } catch (e) {
    Logger.log('Cache read error: ' + e);
    return null;
  }
}

function setCachedLinks(links) {
  try {
    const cache = PropertiesService.getUserProperties();
    cache.setProperty('exam_links_cache', JSON.stringify(links));
    cache.setProperty('exam_links_cache_time', Date.now().toString());
    Logger.log('Cache SET - Links cached for 5 minutes');
  } catch (e) {
    Logger.log('Cache write error: ' + e);
  }
}

function clearLinksCache() {
  try {
    const cache = PropertiesService.getUserProperties();
    cache.deleteProperty('exam_links_cache');
    cache.deleteProperty('exam_links_cache_time');
    Logger.log('Cache CLEARED - Links cache invalidated');
  } catch (e) {
    Logger.log('Cache clear error: ' + e);
  }
}

// ===========================================
// WEB APP DEPLOYMENT
// ===========================================

function doGet(e) {
  const page = e.parameter.page || 'browser';
  const action = e.parameter.action;
  
  // Handle API actions (return outputs directly; functions now support JSONP)
  if (action === 'saveLink') {
    return saveExamLinkAPI(e);
  } else if (action === 'getLinks') {
    return getExamLinksAPI(e);
  } else if (action === 'updateSubject') {
    return updateExamSubjectAPI(e);
  } else if (action === 'deleteLink') {
    return deleteExamLinkAPI(e);
  } else if (action === 'toggleStatus') {
    return toggleStatusAPI(e);
  }
  
  // Handle page routing
  if (page === 'browser') {
    return getBrowserPage();
  } else if (page === 'admin') {
    return getAdminDashboard();
  } else if (page === 'qr') {
    return generateQRPage(e.parameter.url);
  }
  
  return getBrowserPage();
}

// Add CORS Headers
function addCorsHeaders(response) {
  return ContentService.createTextOutput(response.getContent())
    .setMimeType(ContentService.MimeType.JSON)
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// Save Exam Link API - OPTIMIZED WITH BATCH OPERATIONS
function saveExamLinkAPI(e) {
  try {
    const url = e.parameter.url;
    const npsn = e.parameter.npsn || '';
    const pin = e.parameter.pin || '';

    if (!url) {
      return _createJsonOrJsonpResponse({ success: false, error: 'URL is required' }, e.parameter.callback);
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Exam_Links');

    // Update header jika belum ada kolom NPSN dan PIN
    if (!sheet) {
      sheet = ss.insertSheet('Exam_Links');
      sheet.appendRow([
        'Timestamp',
        'Nomor',
        'Nama Mapel',
        'Link',
        'Status',
        'NPSN',
        'PIN'
      ]);

      const headerRange = sheet.getRange(1, 1, 1, 7);
      headerRange.setBackground('#10b981');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);

      // Set column widths
      sheet.setColumnWidth(1, 150); // Timestamp
      sheet.setColumnWidth(2, 60);  // Nomor
      sheet.setColumnWidth(3, 200); // Nama Mapel
      sheet.setColumnWidth(4, 400); // Link
      sheet.setColumnWidth(5, 100); // Status
      sheet.setColumnWidth(6, 100); // NPSN
      sheet.setColumnWidth(7, 100); // PIN
    } else {
      // Jika sheet sudah ada, pastikan ada minimal 7 kolom dan header NPSN/PIN di kolom F/G
      const lastCol = sheet.getLastColumn();
      if (lastCol < 7) {
        sheet.insertColumnsAfter(lastCol, 7 - lastCol);
      }
      const headersRow = sheet.getRange(1, 1, 1, 7).getValues()[0].map(h => (h || '').toString().trim());
      if ((headersRow[5] || '').toLowerCase() !== 'npsn') {
        sheet.getRange(1, 6).setValue('NPSN');
      }
      if ((headersRow[6] || '').toLowerCase() !== 'pin') {
        sheet.getRange(1, 7).setValue('PIN');
      }
    }

    const subject = e.parameter.subject || 'Tidak diketahui';

    // OPTIMASI 1: Ambil hanya data yang perlu, bukan semuanya
    const lastRow = sheet.getLastRow();
    let data = [];
    if (lastRow > 1) {
      data = sheet.getRange(1, 1, lastRow, 7).getValues();
    }

    let urlExists = false;
    let updateRow = -1;

    // OPTIMASI 2: Find operation yang efisien
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === url) {
        urlExists = true;
        updateRow = i;
        break;
      }
    }

    // OPTIMASI 3: Batch write operations
    if (!urlExists) {
      // Tambah baris baru
      sheet.appendRow([
        new Date(),
        '', // Nomor akan diupdate nanti
        subject,
        url,
        'Aktif',
        npsn,
        pin
      ]);
    } else {
      // Update existing row
      sheet.getRange(updateRow + 1, 3).setValue(subject);
      sheet.getRange(updateRow + 1, 1).setValue(new Date());
      sheet.getRange(updateRow + 1, 6).setValue(npsn);
      sheet.getRange(updateRow + 1, 7).setValue(pin);
    }

    // OPTIMASI 4: Invalidate cache setelah update
    clearLinksCache();

    return _createJsonOrJsonpResponse({ 
      success: true, 
      message: urlExists ? 'Link sudah ada, data diperbarui' : 'Link berhasil disimpan',
      isDuplicate: urlExists,
      cacheCleared: true
    }, e.parameter.callback);
    
  } catch (error) {
    Logger.log('Error saving link: ' + error);
    return _createJsonOrJsonpResponse({ success: false, error: error.toString() }, e.parameter.callback);
  }
}

// Fungsi untuk menghapus duplikat link berdasarkan URL
function removeDuplicateLinks(sheet) {
  try {
    const data = sheet.getDataRange().getValues();
    const uniqueUrls = new Map();
    const rowsToDelete = [];
    
    // Loop dari bawah ke atas (skip header row 0)
    for (let i = data.length - 1; i >= 1; i--) {
      const url = data[i][3];
      
      if (!url || url.toString().trim() === '') {
        // Hapus baris kosong
        rowsToDelete.push(i + 1);
        continue;
      }
      
      if (uniqueUrls.has(url)) {
        // URL duplikat, tandai untuk dihapus
        rowsToDelete.push(i + 1);
      } else {
        // URL unik, simpan
        uniqueUrls.set(url, i);
      }
    }
    
    // Hapus baris duplikat (dari bawah ke atas agar index tidak berubah)
    rowsToDelete.sort((a, b) => b - a);
    for (const rowIndex of rowsToDelete) {
      sheet.deleteRow(rowIndex);
    }
    
    // Update nomor urut setelah menghapus duplikat
    const updatedData = sheet.getDataRange().getValues();
    for (let i = 1; i < updatedData.length; i++) {
      sheet.getRange(i + 1, 2).setValue(i); // Update kolom Nomor
    }
    
    Logger.log(`Removed ${rowsToDelete.length} duplicate/empty rows`);
    
  } catch (error) {
    Logger.log('Error removing duplicates: ' + error);
  }
}

// Get Exam Links API - OPTIMIZED WITH CACHING
function getExamLinksAPI(e) {
  try {
    // OPTIMASI 1: Cek cache dulu - return langsung jika cache valid
    const cached = getCachedLinks();
    if (cached) {
      return _createJsonOrJsonpResponse({ 
        success: true, 
        links: cached,
        fromCache: true,
        cacheExpiry: CACHE_CONFIG.EXPIRY_MS
      }, e.parameter.callback);
    }
    
    // OPTIMASI 2: Jika cache miss, baru akses spreadsheet
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Exam_Links');
    
    if (!sheet) {
      setCachedLinks([]);
      return _createJsonOrJsonpResponse({ success: true, links: [] }, e.parameter.callback);
    }
    
    // OPTIMASI 3: Ambil hanya jumlah baris yang diperlukan, bukan semua
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      setCachedLinks([]);
      return _createJsonOrJsonpResponse({ success: true, links: [] }, e.parameter.callback);
    }
    
    // OPTIMASI 4: Ambil kolom sesuai header sheet - cari kolom NPSN/PIN secara case-insensitive
    const totalCols = Math.max(7, sheet.getLastColumn());
    const data = sheet.getRange(1, 1, lastRow, totalCols).getValues();
    const headerRow = (data[0] || []).map(h => (h || '').toString().trim());

    // helper: cari index header dengan beberapa varian nama (case-insensitive)
    function findHeaderIndex(variants) {
      for (let j = 0; j < headerRow.length; j++) {
        const val = (headerRow[j] || '').toString().toLowerCase();
        if (variants.indexOf(val) !== -1) return j;
      }
      return -1;
    }

    const npsnCandidate = findHeaderIndex(['npsn']);
    const pinCandidate = findHeaderIndex(['pin', 'pin ujian', 'pin mapel']);
    const npsnIdx = npsnCandidate !== -1 ? npsnCandidate : 5; // fallback kolom F (index 5)
    const pinIdx = pinCandidate !== -1 ? pinCandidate : 6; // fallback kolom G (index 6)

    // jika headerRow kosong atau tidak ada NPSN/PIN maka log supaya admin tahu
    if (headerRow.length < 6 || (headerRow[npsnIdx] || '') === '' || (headerRow[pinIdx] || '') === '') {
      Logger.log('getExamLinksAPI: headerRow detected -> ' + JSON.stringify(headerRow));
    }

    const links = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Normalize NPSN and PIN values: trim and ensure they are strings
      const npsnVal = (row[npsnIdx] === undefined || row[npsnIdx] === null) ? '' : row[npsnIdx].toString().trim();
      const pinVal = (row[pinIdx] === undefined || row[pinIdx] === null) ? '' : row[pinIdx].toString().trim();
      links.push({
        timestamp: row[0],
        number: row[1],
        subject: row[2],
        url: row[3],
        status: row[4] || 'Aktif',
        npsn: npsnVal,
        pin: pinVal
      });
    }

    const reversedLinks = links.reverse();
    // OPTIMASI 5: Simpan hasil ke cache untuk request berikutnya
    setCachedLinks(reversedLinks);
    return _createJsonOrJsonpResponse({
      success: true,
      links: reversedLinks,
      fromCache: false
    }, e.parameter.callback);
    
  } catch (error) {
    Logger.log('Error getting links: ' + error);
    return _createJsonOrJsonpResponse({ 
      success: false, 
      error: error.toString(), 
      links: [] 
    }, e.parameter.callback);
  }
}

// Update Exam Subject API
function updateExamSubjectAPI(e) {
  // EDIT FEATURE DISABLED
  return _createJsonOrJsonpResponse({ success: false, error: 'Fitur edit mata pelajaran dinonaktifkan' }, e.parameter.callback);
}

// Delete Exam Link API
function deleteExamLinkAPI(e) {
  try {
    const url = e.parameter.url;
    
    if (!url) {
      return _createJsonOrJsonpResponse({ success: false, error: 'URL is required' }, e.parameter.callback);
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Exam_Links');
    
    if (!sheet) {
      return _createJsonOrJsonpResponse({ success: false, error: 'Sheet not found' }, e.parameter.callback);
    }
    
    const lastRow = sheet.getLastRow();
    const data = lastRow > 1 ? sheet.getRange(1, 1, lastRow, 5).getValues() : [];
    
    // Find and delete the row
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === url) {
        sheet.deleteRow(i + 1);
        
        // Update nomor urut setelah delete
        const updatedLastRow = sheet.getLastRow();
        if (updatedLastRow > 1) {
          const updatedData = sheet.getRange(1, 1, updatedLastRow, 5).getValues();
          for (let j = 1; j < updatedData.length; j++) {
            sheet.getRange(j + 1, 2).setValue(j);
          }
        }
        
        // OPTIMASI: Clear cache setelah delete
        clearLinksCache();
        
        return _createJsonOrJsonpResponse({ 
          success: true, 
          message: 'Link berhasil dihapus',
          cacheCleared: true
        }, e.parameter.callback);
      }
    }
    
    return _createJsonOrJsonpResponse({ 
      success: false, 
      error: 'URL tidak ditemukan'
    }, e.parameter.callback);
    
  } catch (error) {
    Logger.log('Error deleting link: ' + error);
    return _createJsonOrJsonpResponse({ success: false, error: error.toString() }, e.parameter.callback);
  }
}

// Toggle Status API (Aktif/Non-Aktif)
function toggleStatusAPI(e) {
  try {
    const url = e.parameter.url;
    
    if (!url) {
      return _createJsonOrJsonpResponse({ success: false, error: 'URL is required' }, e.parameter.callback);
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Exam_Links');
    
    if (!sheet) {
      return _createJsonOrJsonpResponse({ success: false, error: 'Sheet not found' }, e.parameter.callback);
    }
    
    const lastRow = sheet.getLastRow();
    const data = lastRow > 1 ? sheet.getRange(1, 1, lastRow, 5).getValues() : [];
    
    // Find and toggle status
    for (let i = 1; i < data.length; i++) {
      if (data[i][3] === url) {
        const currentStatus = data[i][4] || 'Aktif';
        const newStatus = currentStatus === 'Aktif' ? 'Non Aktif' : 'Aktif';
        
        sheet.getRange(i + 1, 5).setValue(newStatus); // Kolom E (index 4)
        sheet.getRange(i + 1, 1).setValue(new Date()); // Update timestamp
        
        // OPTIMASI: Clear cache setelah update
        clearLinksCache();
        
        return _createJsonOrJsonpResponse({ 
          success: true, 
          message: 'Status berhasil diubah menjadi ' + newStatus,
          newStatus: newStatus,
          cacheCleared: true
        }, e.parameter.callback);
      }
    }
    
    return _createJsonOrJsonpResponse({ 
      success: false, 
      error: 'URL tidak ditemukan'
    }, e.parameter.callback);
    
  } catch (error) {
    Logger.log('Error toggling status: ' + error);
    return _createJsonOrJsonpResponse({ success: false, error: error.toString() }, e.parameter.callback);
  }
}

// Helper: return JSON or JSONP depending on callback param
function _createJsonOrJsonpResponse(obj, callbackName) {
  const json = JSON.stringify(obj);
  if (callbackName && callbackName.toString().match(/^[A-Za-z0-9_$.]+$/)) {
    // JSONP: return JavaScript that calls the callback
    return ContentService.createTextOutput(callbackName + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const payload = data.data;
    
    let result;
    
    switch(action) {
      case 'startSession':
        result = startExamSessionDirect(payload);
        break;
      case 'endSession':
        result = endExamSessionDirect(payload);
        break;
      case 'logViolation':
        result = logViolationDirect(payload);
        break;
      default:
        result = { success: false, error: 'Unknown action' };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('doPost error: ' + error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function startExamSessionDirect(sessionData) {
  // Sessions storage disabled: return success without writing to spreadsheet.
  try {
    const sessionId = Utilities.getUuid ? Utilities.getUuid() : ('SESSION_' + Date.now());
    return { success: true, sessionId: sessionId, note: 'sessions_disabled' };
  } catch (error) {
    Logger.log('startExamSessionDirect (noop) error: ' + error);
    return { success: false, error: error.toString() };
  }
}

function endExamSessionDirect(endData) {
  // Sessions storage disabled: acknowledge end request without modifying spreadsheet.
  try {
    return { success: true, note: 'sessions_disabled' };
  } catch (error) {
    Logger.log('endExamSessionDirect (noop) error: ' + error);
    return { success: false, error: error.toString() };
  }
}

function logViolationDirect(violationData) {
  // Violation logging disabled: return success but do not persist violation data.
  try {
    return { success: true, note: 'violations_disabled' };
  } catch (error) {
    Logger.log('logViolationDirect (noop) error: ' + error);
    return { success: false, error: error.toString() };
  }
}

function updateViolationCount(sessionId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sessionsSheet = ss.getSheetByName('Sessions');
    const violationsSheet = ss.getSheetByName('Violations');
    
    if (!sessionsSheet || !violationsSheet) return;
    
    const sessionData = sessionsSheet.getDataRange().getValues();
    const violationData = violationsSheet.getDataRange().getValues();
    
    // Count violations for this session
    let count = 0;
    for (let i = 1; i < violationData.length; i++) {
      if (violationData[i][1] === sessionId) {
        count++;
      }
    }
    
    // Update in sessions sheet
    for (let i = 1; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionId) {
        sessionsSheet.getRange(i + 1, 10).setValue(count);
        break;
      }
    }
    
  } catch (error) {
    Logger.log('Error updating violation count: ' + error);
  }
}

// ===========================================
// HTML PAGES
// ===========================================

function getBrowserPage() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('EXAMUNA')
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/forms_2020q4_48dp.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ===========================================
// QR CODE GENERATOR
// ===========================================

function generateQRCode(examUrl, size = 300) {
  // Generate QR Code URL using QR Server API
  const qrUrl = `${CONFIG.QR_API}?size=${size}x${size}&data=${encodeURIComponent(examUrl)}`;
  
  return {
    qrUrl: qrUrl,
    examUrl: examUrl,
    generatedAt: new Date().toISOString()
  };
}

function generateQRPage(url) {
  if (!url) {
    return HtmlService.createHtmlOutput('<h3>Error: URL tidak ditemukan</h3>');
  }
  
  const qr = generateQRCode(url);
  
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <title>QR Code - Exam Access</title>
      <style>
        body {
          font-family: 'Segoe UI', Tahoma, sans-serif;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          min-height: 100vh;
          margin: 0;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
        }
        .container {
          text-align: center;
          background: white;
          padding: 40px;
          border-radius: 20px;
          box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }
        .title {
          color: #333;
          margin-bottom: 10px;
        }
        .subtitle {
          color: #666;
          margin-bottom: 30px;
        }
        img {
          border: 5px solid #667eea;
          border-radius: 15px;
          margin: 20px 0;
        }
        .url {
          background: #f0f0f0;
          padding: 15px;
          border-radius: 10px;
          font-family: monospace;
          color: #333;
          word-break: break-all;
          margin-top: 20px;
        }
        .btn {
          margin-top: 20px;
          padding: 12px 30px;
          background: #667eea;
          color: white;
          border: none;
          border-radius: 8px;
          font-size: 16px;
          cursor: pointer;
          text-decoration: none;
          display: inline-block;
        }
        .btn:hover {
          background: #5568d3;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <h1 class="title">Akses Ujian</h1>
        <p class="subtitle">Scan QR Code dengan ExamBrowser</p>
        <img src="${qr.qrUrl}" alt="QR Code" width="300" height="300">
        <div class="url">${url}</div>
        <a href="${url}" class="btn" target="_blank">Buka Link Langsung</a>
      </div>
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(html);
}

// ===========================================
// SESSION MANAGEMENT
// ===========================================

function startExamSession(sessionData) {
  // Sessions disabled: return a generated sessionId but do not persist to spreadsheet.
  try {
    const sessionId = Utilities.getUuid ? Utilities.getUuid() : ('SESSION_' + Date.now());
    return { success: true, sessionId: sessionId, note: 'sessions_disabled' };
  } catch (error) {
    Logger.log('startExamSession (noop) error: ' + error);
    return { success: false, error: error.toString() };
  }
}

function endExamSession(sessionId, endData) {
  // Sessions disabled: simply acknowledge the end request without persisting.
  try {
    return { success: true, note: 'sessions_disabled' };
  } catch (error) {
    Logger.log('endExamSession (noop) error: ' + error);
    return { success: false, error: error.toString() };
  }
}

// ===========================================
// VIOLATION TRACKING
// ===========================================

function logViolation(violationData) {
  // Violation logging disabled: return success but do not persist any violation data.
  try {
    return { success: true, note: 'violations_disabled' };
  } catch (error) {
    Logger.log('logViolation (noop) error: ' + error);
    return { success: false, error: error.toString() };
  }
}

function updateViolationCount(sessionId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sessionsSheet = ss.getSheetByName('Sessions');
    const violationsSheet = ss.getSheetByName('Violations');
    
    if (!sessionsSheet || !violationsSheet) return;
    
    const sessionData = sessionsSheet.getDataRange().getValues();
    const violationData = violationsSheet.getDataRange().getValues();
    
    // Count violations for this session
    let count = 0;
    for (let i = 1; i < violationData.length; i++) {
      if (violationData[i][1] === sessionId) {
        count++;
      }
    }
    
    // Update in sessions sheet
    for (let i = 1; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionId) {
        sessionsSheet.getRange(i + 1, 10).setValue(count);
        break;
      }
    }
    
  } catch (error) {
    Logger.log('Error updating violation count: ' + error);
  }
}

// ===========================================
// EXAM LINKS MANAGEMENT
// ===========================================

function createExamLink(examData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Exam Links');
    
    if (!sheet) {
      sheet = ss.insertSheet('Exam Links');
      sheet.appendRow([
        'Link ID',
        'Exam Title',
        'Exam URL',
        'Subject',
        'Class',
        'Duration (minutes)',
        'Created Date',
        'Created By',
        'Expiry Date',
        'Status',
        'QR Code URL',
        'Total Access',
        'Active Sessions'
      ]);
      
      // Format header
      const headerRange = sheet.getRange(1, 1, 1, 13);
      headerRange.setBackground('#10b981');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
    }
    
    const linkId = Utilities.getUuid().substring(0, 8).toUpperCase();
    const qrData = generateQRCode(examData.url);
    
    sheet.appendRow([
      linkId,
      examData.title,
      examData.url,
      examData.subject || '-',
      examData.class || '-',
      examData.duration || 90,
      new Date(),
      examData.createdBy || 'Admin',
      examData.expiryDate || '',
      'Active',
      qrData.qrUrl,
      0,
      0
    ]);
    
    return {
      success: true,
      linkId: linkId,
      qrUrl: qrData.qrUrl
    };
    
  } catch (error) {
    Logger.log('Error creating exam link: ' + error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function getExamLinks() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Exam Links');
    
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const links = [];
    
    for (let i = 1; i < data.length; i++) {
      links.push({
        id: data[i][0],
        title: data[i][1],
        url: data[i][2],
        subject: data[i][3],
        class: data[i][4],
        duration: data[i][5],
        createdDate: data[i][6],
        expiryDate: data[i][8],
        status: data[i][9],
        qrUrl: data[i][10],
        totalAccess: data[i][11]
      });
    }
    
    return links;
    
  } catch (error) {
    Logger.log('Error getting exam links: ' + error);
    return [];
  }
}

function getExamLinkById(linkId) {
  try {
    const links = getExamLinks();
    const link = links.find(l => l.id === linkId);
    
    if (link) {
      // Check if expired
      if (link.expiryDate && new Date(link.expiryDate) < new Date()) {
        return { success: false, error: 'Link telah kadaluarsa' };
      }
      
      // Increment access count
      incrementAccessCount(linkId);
      
      return { success: true, link: link };
    }
    
    return { success: false, error: 'Link ID tidak ditemukan' };
    
  } catch (error) {
    Logger.log('Error getting exam link by ID: ' + error);
    return { success: false, error: error.toString() };
  }
}

function incrementAccessCount(linkId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Exam Links');
    
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === linkId) {
        const currentCount = data[i][11] || 0;
        sheet.getRange(i + 1, 12).setValue(currentCount + 1);
        break;
      }
    }
  } catch (error) {
    Logger.log('Error incrementing access count: ' + error);
  }
}

// ===========================================
// ANALYTICS & REPORTS
// ===========================================

function getSessionAnalytics() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sessionsSheet = ss.getSheetByName('Sessions');
    const violationsSheet = ss.getSheetByName('Violations');
    
    if (!sessionsSheet) {
      return {
        totalSessions: 0,
        activeSessions: 0,
        completedSessions: 0,
        totalViolations: 0,
        averageDuration: 0
      };
    }
    
    const sessionData = sessionsSheet.getDataRange().getValues();
    const violationData = violationsSheet ? violationsSheet.getDataRange().getValues() : [];
    
    let totalSessions = sessionData.length - 1;
    let activeSessions = 0;
    let completedSessions = 0;
    let totalDuration = 0;
    let totalViolations = violationData.length - 1;
    
    for (let i = 1; i < sessionData.length; i++) {
      const status = sessionData[i][5];
      const duration = sessionData[i][8];
      
      if (status === 'Active') activeSessions++;
      if (status === 'Completed') completedSessions++;
      if (duration) totalDuration += parseFloat(duration);
    }
    
    return {
      totalSessions: totalSessions,
      activeSessions: activeSessions,
      completedSessions: completedSessions,
      totalViolations: totalViolations,
      averageDuration: completedSessions > 0 ? Math.round(totalDuration / completedSessions) : 0
    };
    
  } catch (error) {
    Logger.log('Error getting analytics: ' + error);
    return null;
  }
}

function generateSessionReport(sessionId) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sessionsSheet = ss.getSheetByName('Sessions');
    const violationsSheet = ss.getSheetByName('Violations');
    
    const sessionData = sessionsSheet.getDataRange().getValues();
    const violationData = violationsSheet ? violationsSheet.getDataRange().getValues() : [];
    
    let session = null;
    for (let i = 1; i < sessionData.length; i++) {
      if (sessionData[i][0] === sessionId) {
        session = {
          sessionId: sessionData[i][0],
          timestamp: sessionData[i][1],
          studentName: sessionData[i][2],
          studentId: sessionData[i][3],
          examUrl: sessionData[i][4],
          status: sessionData[i][5],
          startTime: sessionData[i][6],
          endTime: sessionData[i][7],
          duration: sessionData[i][8],
          violations: sessionData[i][9]
        };
        break;
      }
    }
    
    if (!session) return null;
    
    // Get violations for this session
    const sessionViolations = [];
    for (let i = 1; i < violationData.length; i++) {
      if (violationData[i][1] === sessionId) {
        sessionViolations.push({
          timestamp: violationData[i][0],
          type: violationData[i][4],
          description: violationData[i][5],
          severity: violationData[i][6]
        });
      }
    }
    
    session.violationsList = sessionViolations;
    
    return session;
    
  } catch (error) {
    Logger.log('Error generating report: ' + error);
    return null;
  }
}

// ===========================================
// SPREADSHEET INITIALIZATION
// ===========================================

function initializeSpreadsheet() {
  try {
    const ss = SpreadsheetApp.create(`ExamBrowser Data - ${new Date().toLocaleDateString()}`);
    const spreadsheetId = ss.getId();
    
    Logger.log('Spreadsheet Created!');
    Logger.log('Spreadsheet ID: ' + spreadsheetId);
    Logger.log('URL: ' + ss.getUrl());
    
    // Create all sheets
    createSessionsSheet(ss);
    createViolationsSheet(ss);
    createExamLinksSheet(ss);
    createConfigSheet(ss);
    
    // Delete default sheet
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);
    
    return {
      success: true,
      spreadsheetId: spreadsheetId,
      url: ss.getUrl()
    };
    
  } catch (error) {
    Logger.log('Error initializing spreadsheet: ' + error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

function createSessionsSheet(ss) {
  const sheet = ss.insertSheet('Sessions');
  sheet.appendRow([
    'Session ID',
    'Timestamp',
    'Student Name',
    'Student ID',
    'Exam URL',
    'Status',
    'Start Time',
    'End Time',
    'Duration (minutes)',
    'Violations',
    'IP Address',
    'User Agent'
  ]);
  
  const headerRange = sheet.getRange(1, 1, 1, 12);
  headerRange.setBackground('#667eea');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

function createViolationsSheet(ss) {
  const sheet = ss.insertSheet('Violations');
  sheet.appendRow([
    'Timestamp',
    'Session ID',
    'Student Name',
    'Student ID',
    'Violation Type',
    'Description',
    'Severity',
    'Action Taken'
  ]);
  
  const headerRange = sheet.getRange(1, 1, 1, 8);
  headerRange.setBackground('#ef4444');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

function createExamLinksSheet(ss) {
  const sheet = ss.insertSheet('Exam Links');
  sheet.appendRow([
    'Link ID',
    'Exam Title',
    'Exam URL',
    'Subject',
    'Class',
    'Duration (minutes)',
    'Created Date',
    'Created By',
    'Expiry Date',
    'Status',
    'QR Code URL',
    'Total Access',
    'Active Sessions'
  ]);
  
  const headerRange = sheet.getRange(1, 1, 1, 13);
  headerRange.setBackground('#10b981');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

function createConfigSheet(ss) {
  const sheet = ss.insertSheet('Config');
  sheet.appendRow(['Setting', 'Value', 'Description']);
  
  const settings = [
    ['School Name', CONFIG.SCHOOL_NAME, 'Nama sekolah'],
    ['App Name', CONFIG.APP_NAME, 'Nama aplikasi'],
    ['Max Violations', '5', 'Maksimal pelanggaran sebelum alert'],
    ['Session Timeout', '180', 'Timeout session dalam menit']
  ];
  
  settings.forEach(setting => sheet.appendRow(setting));
  
  const headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange.setBackground('#8b5cf6');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
}

// ===========================================
// TESTING FUNCTIONS
// ===========================================

function testCreateExamLink() {
  const result = createExamLink({
    title: 'Ujian Matematika Kelas XII',
    url: 'https://forms.gle/example123',
    subject: 'Matematika',
    class: 'XII IPA 1',
    duration: 90,
    createdBy: 'teacher@school.com'
  });
  
  Logger.log(result);
}

function testGenerateQR() {
  const qr = generateQRCode('https://script.google.com/your-web-app-url');
  Logger.log('QR URL: ' + qr.qrUrl);
}

/**
 * TEST FUNGSI: Simpan Link Ujian ke Spreadsheet
 * Jalankan fungsi ini dari Apps Script editor untuk test
 * 
 * Cara menjalankan:
 * 1. Pilih fungsi "testSaveExamLink" dari dropdown
 * 2. Klik Run
 * 3. Cek spreadsheet apakah data tersimpan
 */
function testSaveExamLink() {
  Logger.log('=== TESTING SAVE EXAM LINK ===');
  
  // Simulasi parameter dari web app
  const mockEvent = {
    parameter: {
      url: 'https://forms.google.com/test-matematika',
      subject: 'Matematika',
      callback: null
    }
  };
  
  // Panggil fungsi save
  const result = saveExamLinkAPI(mockEvent);
  
  // Log hasil
  Logger.log('Result: ' + result.getContent());
  
  // Cek spreadsheet
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Exam_Links');
  
  if (sheet) {
    const lastRow = sheet.getLastRow();
    Logger.log('Total rows in Exam_Links: ' + lastRow);
    
    if (lastRow > 1) {
      const lastData = sheet.getRange(lastRow, 1, 1, 5).getValues()[0];
      Logger.log('Last saved data:');
      Logger.log('  Timestamp: ' + lastData[0]);
      Logger.log('  Nomor: ' + lastData[1]);
      Logger.log('  Nama Mapel: ' + lastData[2]);
      Logger.log('  Link: ' + lastData[3]);
      Logger.log('  Status: ' + lastData[4]);
    }
  } else {
    Logger.log('ERROR: Sheet Exam_Links not found!');
  }
  
  Logger.log('=== TEST COMPLETE ===');
  Logger.log('Silakan cek spreadsheet untuk memastikan data tersimpan');
}

/**
 * TEST FUNGSI: Get Exam Links dari Spreadsheet
 */
function testGetExamLinks() {
  Logger.log('=== TESTING GET EXAM LINKS ===');
  
  const mockEvent = {
    parameter: {
      callback: null
    }
  };
  
  const result = getExamLinksAPI(mockEvent);
  Logger.log('Result: ' + result.getContent());
  
  Logger.log('=== TEST COMPLETE ===');
}

/**
 * INSTRUKSI SETUP:
 * 
 * 1. Deploy sebagai Web App
 * 2. Copy Deployment URL
 * 3. Update SCRIPT_URL di index.html
 * 4. Test platform EXAMUNA
 * 
 * FITUR YANG TERSEDIA:
 * - QR Code generation untuk link ujian
 * - Session tracking real-time
 * - Violation logging
 * - Analytics dashboard
 * - Multi-platform support
 */

/**
 * Setup / repair headers and basic formatting for required sheets.
 * Run manually from Apps Script editor: pilih fungsi setupSheetHeaders lalu Run.
 */
function setupSheetHeaders() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const definitions = [
    {
      name: 'Exam_Links',
      headers: ['Timestamp', 'Nomor', 'Nama Mapel', 'Link', 'Status'],
      bg: '#10b981',
      widths: [150, 60, 200, 400, 100]
    },
    {
      name: 'Exam Links',
      headers: [
        'Link ID','Exam Title','Exam URL','Subject','Class','Duration (minutes)',
        'Created Date','Created By','Expiry Date','Status','QR Code URL','Total Access','Active Sessions'
      ],
      bg: '#10b981',
      widths: [100,200,300,140,100,120,140,140,120,80,220,100,120]
    },
    {
      name: 'Sessions',
      headers: [
        'Session ID','Timestamp','Student Name','Student ID','Exam URL','Status',
        'Start Time','End Time','Duration (minutes)','Violations','IP Address','User Agent'
      ],
      bg: '#667eea',
      widths: [140,140,160,100,300,90,160,160,140,100,140,300]
    },
    {
      name: 'Violations',
      headers: [
        'Timestamp','Session ID','Student Name','Student ID','Violation Type',
        'Description','Severity','Action Taken'
      ],
      bg: '#ef4444',
      widths: [140,140,160,100,160,300,100,140]
    },
    {
      name: 'Config',
      headers: ['Setting','Value','Description'],
      bg: '#8b5cf6',
      widths: [220,220,400]
    }
  ];

  definitions.forEach(def => {
    let sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
    }

    // Ensure header row length
    const headerRange = sheet.getRange(1, 1, 1, def.headers.length);
    headerRange.setValues([def.headers]);
    headerRange.setBackground(def.bg);
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);

    // Set column widths where provided
    if (Array.isArray(def.widths)) {
      def.widths.forEach((w, idx) => {
        try { sheet.setColumnWidth(idx + 1, w); } catch (e) { /* ignore */ }
      });
    }

    // Optional: clear residual header cells beyond defined headers (keeps tidy)
    const lastCol = sheet.getLastColumn();
    if (lastCol > def.headers.length) {
      sheet.getRange(1, def.headers.length + 1, 1, lastCol - def.headers.length).clearContent();
    }
  });

  // Feedback log
  Logger.log('setupSheetHeaders completed for spreadsheet: ' + CONFIG.SPREADSHEET_ID);
}
