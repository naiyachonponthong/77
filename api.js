// ============================================================
// API Client — Google Apps Script Backend
// ============================================================

var APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxRwxGGW3fxIB0rKRIU2zh9lEo_yUTVEcW6cWAqMF4YYJBvu0YxCXfy6mUbj8ihTyaRXQ/exec';

/**
 * parseApiResponse — อ่านคำตอบจาก Apps Script
 * ปกติจะได้ JSON แต่ถ้า Apps Script พังหรือยังไม่ได้ตั้งสิทธิ์ให้เข้าถึง
 * มันจะตอบกลับเป็นหน้า HTML แทน — กรณีนั้นให้ดึงข้อความ error จริงออกมาบอกผู้ใช้
 * แทนที่จะขึ้นแค่ "ไม่สามารถเชื่อมต่อระบบได้" ซึ่งหาสาเหตุไม่ได้
 */
function parseApiResponse(res, fnName) {
  return res.text().then(function(text) {
    try {
      return JSON.parse(text);
    } catch (e) {
      var detail = extractAppsScriptError(text);
      console.error('[API] Backend ไม่ได้ตอบเป็น JSON [' + fnName + ']:', detail || text.substring(0, 500));
      var err = new Error(detail || 'Google Apps Script ไม่ได้ตอบกลับเป็นข้อมูล JSON');
      err.isBackendError = true;
      throw err;
    }
  });
}

/** extractAppsScriptError — ดึงข้อความ error ออกจากหน้า HTML ที่ Apps Script ส่งกลับมา */
function extractAppsScriptError(html) {
  if (!html) return '';
  var text = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .trim();
  if (!text) return '';

  // โค้ดฝั่งหน้าเว็บถูกนำไปวางใน Apps Script (อาการที่เจอบ่อยที่สุด)
  if (/localStorage|document|window is not defined/i.test(text)) {
    return 'ดูเหมือนโค้ดที่วางไว้ใน Google Apps Script เป็นไฟล์ฝั่งหน้าเว็บ (app.js) '
         + 'กรุณานำเนื้อหาไฟล์ code.gs ไปวางแทน แล้ว Deploy ใหม่ — รายละเอียด: ' + text.substring(0, 200);
  }
  if (/ต้องได้รับสิทธิ|need(s)? permission|authoriz|ServiceLogin|Sign in/i.test(text)) {
    return 'Google Apps Script ยังไม่เปิดสิทธิ์ให้เข้าถึงแบบสาธารณะ '
         + 'กรุณา Deploy ใหม่โดยตั้ง Who has access = Anyone';
  }
  return text.substring(0, 250);
}

function callAPI(fnName) {
  var args = Array.prototype.slice.call(arguments, 1);
  // ใช้ POST เมื่อ payload ใหญ่เกิน URL length limit (base64 รูป / เพิ่มวัสดุหลายรายการ)
  if (fnName === 'uploadFile' || fnName === 'addItemsBulk' || fnName === 'addWithdrawalBulk' || fnName === 'adjustStock') {
    var body = 'fn=' + encodeURIComponent(fnName) + '&args=' + encodeURIComponent(JSON.stringify(args));
    return fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body
    }).then(function(res) {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return parseApiResponse(res, fnName);
    }).catch(function(err) {
      console.error('API Error [' + fnName + ']:', err);
      throw err;
    });
  }
  var url = APPS_SCRIPT_URL + '?fn=' + encodeURIComponent(fnName) + '&args=' + encodeURIComponent(JSON.stringify(args));
  console.log('[API] GET', url);

  return fetch(url, { method: 'GET', mode: 'cors' }).then(function(res) {
    console.log('[API] Response', res.status);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    return parseApiResponse(res, fnName);
  }).then(function(data) {
    console.log('[API] Data', data);
    return data;
  }).catch(function(err) {
    console.warn('[API] Error for', fnName, err);
    if (window._mockAPI && window._mockAPI[fnName]) {
      return Promise.resolve(window._mockAPI[fnName].apply(null, args));
    }
    throw err;
  });
}

// Helper: แปลง file_id เป็น URL สำหรับแสดงรูป
function getFileDataUrl(fileId) {
  if (!fileId) return '';
  if (String(fileId).indexOf('http') === 0) return fileId;
  return 'https://lh5.googleusercontent.com/d/' + fileId;
}
