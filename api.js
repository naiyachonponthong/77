// ============================================================
// API Client — Google Apps Script Backend
// ============================================================

var APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxRwxGGW3fxIB0rKRIU2zh9lEo_yUTVEcW6cWAqMF4YYJBvu0YxCXfy6mUbj8ihTyaRXQ/exec';

function callAPI(fnName) {
  var args = Array.prototype.slice.call(arguments, 1);
  // ใช้ POST เมื่อ payload ใหญ่เกิน URL length limit (base64 รูป / เพิ่มวัสดุหลายรายการ)
  if (fnName === 'uploadFile' || fnName === 'addItemsBulk' || fnName === 'addWithdrawalBulk') {
    var body = 'fn=' + encodeURIComponent(fnName) + '&args=' + encodeURIComponent(JSON.stringify(args));
    return fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body
    }).then(function(res) {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
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
    return res.json();
  }).then(function(data) {
    console.log('[API] Data', data);
    return data;
  }).catch(function(err) {
    console.warn('[API] Fallback to localStorage mock for', fnName, err);
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
