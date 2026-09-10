/* ============================================================
   Data layer — ONE Excel workbook, TWO sheets:
     sheet "Donors"   : வரிசை எண் | கொடுத்தவர் பெயர் | அமௌன்ட் (₹)  + total row
     sheet "Expenses" : வரிசை எண் | தேதி | விவரம் | தொகை (₹)      + total row
   The totals ARE written into the Excel file (bottom of each sheet).

   Storage modes:
   - Local  : link the project's data folder once; the single file
              data/donors.xlsx is read/written there (auto-created).
   - GitHub : GITHUB_REPO set → the same workbook is committed to
              the repo's data/donors.xlsx (token typed at runtime).
   No localStorage / sessionStorage for donor/expense data.
   ============================================================ */
(function () {
  'use strict';

  var FILE_NAME = 'donors.xlsx';
  var DONORS_SHEET = 'Donors';
  var EXPENSES_SHEET = 'Expenses';
  var DCOL = { serial: 'வரிசை எண்', name: 'கொடுத்தவர் பெயர்', amount: 'அமௌன்ட் (₹)' };
  var ECOL = { serial: 'வரிசை எண்', date: 'தேதி', desc: 'விவரம்', amount: 'தொகை (₹)' };
  var TOTAL_DONORS = 'மொத்த நன்கொடை';
  var TOTAL_EXPENSES = 'மொத்த செலவு';
  var IDB_NAME = 'vg-donors';
  var IDB_KEY = 'data-folder-handle';
  var dirHandle = null;
  var handle = null;

  function supported() {
    return typeof window !== 'undefined' && typeof window.showDirectoryPicker === 'function';
  }

  /* ---------- IndexedDB: remember the data-folder handle ---------- */
  function idbOpen() {
    return new Promise(function (res, rej) {
      var rq = indexedDB.open(IDB_NAME, 1);
      rq.onupgradeneeded = function () { rq.result.createObjectStore('handles'); };
      rq.onsuccess = function () { res(rq.result); };
      rq.onerror = function () { rej(rq.error); };
    });
  }
  function idbGet() {
    return idbOpen().then(function (db) {
      return new Promise(function (res) {
        var rq = db.transaction('handles', 'readonly').objectStore('handles').get(IDB_KEY);
        rq.onsuccess = function () { res(rq.result || null); };
        rq.onerror = function () { res(null); };
      });
    }).catch(function () { return null; });
  }
  function idbSet(v) {
    return idbOpen().then(function (db) {
      return new Promise(function (res) {
        var tx = db.transaction('handles', 'readwrite').objectStore('handles').put(v, IDB_KEY);
        tx.onsuccess = function () { res(); };
        tx.onerror = function () { res(); };
      });
    }).catch(function () {});
  }

  /* ---------- parsing ---------- */
  function isTotalRow(name) { return name === TOTAL_DONORS || name === TOTAL_EXPENSES; }

  function parseDonors(ws) {
    if (!ws) return [];
    var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return rows.map(function (r) {
      return {
        name: String(r[DCOL.name] !== undefined ? r[DCOL.name] : '').trim(),
        amount: Number(r[DCOL.amount]) || 0
      };
    }).filter(function (d) { return (d.name || d.amount > 0) && !isTotalRow(d.name); });
  }

  function parseExpenses(ws) {
    if (!ws) return [];
    var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return rows.map(function (r) {
      return {
        date: String(r[ECOL.date] !== undefined ? r[ECOL.date] : '').trim(),
        desc: String(r[ECOL.desc] !== undefined ? r[ECOL.desc] : '').trim(),
        amount: Number(r[ECOL.amount]) || 0
      };
    }).filter(function (d) { return (d.desc || d.amount > 0) && !isTotalRow(d.desc); });
  }

  function parseWorkbook(buf) {
    var wb = XLSX.read(buf, { type: 'array' });
    var dws = wb.Sheets[DONORS_SHEET] || wb.Sheets[wb.SheetNames[0]];
    var ews = wb.Sheets[EXPENSES_SHEET] ||
              (wb.SheetNames.length > 1 ? wb.Sheets[wb.SheetNames[1]] : null);
    return { donors: parseDonors(dws), expenses: parseExpenses(ews) };
  }

  function sum(list) {
    return list.reduce(function (s, d) { return s + (Number(d.amount) || 0); }, 0);
  }

  /* ---------- building sheets (totals included in the file) ---------- */
  function donorsSheet(list) {
    var rows = list.map(function (d, i) {
      var o = {};
      o[DCOL.serial] = i + 1;
      o[DCOL.name] = d.name;
      o[DCOL.amount] = d.amount;
      return o;
    });
    var t = {};
    t[DCOL.serial] = '';
    t[DCOL.name] = TOTAL_DONORS;
    t[DCOL.amount] = sum(list);
    rows.push(t);
    var ws = XLSX.utils.json_to_sheet(rows, { header: [DCOL.serial, DCOL.name, DCOL.amount] });
    ws['!cols'] = [{ wch: 10 }, { wch: 32 }, { wch: 14 }];
    return ws;
  }

  function expensesSheet(list) {
    var rows = list.map(function (d, i) {
      var o = {};
      o[ECOL.serial] = i + 1;
      o[ECOL.date] = d.date;
      o[ECOL.desc] = d.desc;
      o[ECOL.amount] = d.amount;
      return o;
    });
    var t = {};
    t[ECOL.serial] = '';
    t[ECOL.date] = '';
    t[ECOL.desc] = TOTAL_EXPENSES;
    t[ECOL.amount] = sum(list);
    rows.push(t);
    var ws = XLSX.utils.json_to_sheet(rows, { header: [ECOL.serial, ECOL.date, ECOL.desc, ECOL.amount] });
    ws['!cols'] = [{ wch: 10 }, { wch: 12 }, { wch: 32 }, { wch: 14 }];
    return ws;
  }

  function workbookBytes(state) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, donorsSheet(state.donors || []), DONORS_SHEET);
    XLSX.utils.book_append_sheet(wb, expensesSheet(state.expenses || []), EXPENSES_SHEET);
    return new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  }

  function donorsBytes(list) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, donorsSheet(list || []), DONORS_SHEET);
    return new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  }

  function expensesBytes(list) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, expensesSheet(list || []), EXPENSES_SHEET);
    return new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'array' }));
  }

  function downloadBytes(bytes, name) {
    var blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    if (typeof URL !== 'undefined' && URL.createObjectURL) {
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = name;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
    }
  }

  /* ---------- local data-folder mode ---------- */
  function ensureFile() {
    return dirHandle.getFileHandle(FILE_NAME, { create: true }).then(function (fh) {
      handle = fh;
    });
  }

  function readFileState() {
    return handle.getFile()
      .then(function (f) { return f.arrayBuffer(); })
      .then(parseWorkbook);
  }

  function restore(requestPermission, cb) {
    if (!supported()) { cb(false); return; }
    idbGet().then(function (dh) {
      if (!dh) { cb(false); return; }
      dirHandle = dh;
      var q = dirHandle.queryPermission ? dirHandle.queryPermission({ mode: 'readwrite' }) : Promise.resolve('granted');
      q.then(function (state) {
        if (state === 'granted') {
          ensureFile().then(function () { cb(true); }).catch(function () { cb(false); });
          return;
        }
        if (requestPermission && dirHandle.requestPermission) {
          dirHandle.requestPermission({ mode: 'readwrite' }).then(function (s2) {
            if (s2 === 'granted') {
              ensureFile().then(function () { cb(true); }).catch(function () { cb(false); });
            } else { cb(false); }
          }).catch(function () { cb(false); });
        } else {
          cb(false);
        }
      }).catch(function () { cb(false); });
    });
  }

  function linkFolder(cb) {
    if (!supported()) { cb(false, 'unsupported'); return; }
    window.showDirectoryPicker({ mode: 'readwrite' })
      .then(function (dh) {
        dirHandle = dh;
        return idbSet(dh);
      })
      .then(function () { return ensureFile(); })
      .then(function () { cb(true); })
      .catch(function (e) {
        cb(false, (e && e.name === 'AbortError') ? 'cancelled' : 'error');
      });
  }

  function load(cb) {
    if (handle) {
      readFileState()
        .then(function (state) { cb(state, 'excel'); })
        .catch(function () { cb(null, 'error'); });
      return;
    }
    fetch('data/' + FILE_NAME, { cache: 'no-store' })
      .then(function (r) { if (!r.ok) throw new Error('http ' + r.status); return r.arrayBuffer(); })
      .then(function (buf) {
        var state;
        try { state = parseWorkbook(buf); } catch (e) { state = { donors: [], expenses: [] }; }
        cb(state, 'file');
      })
      .catch(function () { cb(null, 'none'); });
  }

  function save(state, cb) {
    if (!handle) { if (cb) cb(false, 'nolink'); return; }
    handle.createWritable()
      .then(function (w) { return w.write(new Blob([workbookBytes(state)], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      })).then(function () { return w.close(); }); })
      .then(function () { if (cb) cb(true); })
      .catch(function () { if (cb) cb(false, 'locked'); });
  }

  function linked() { return !!handle; }
  function folderName() { return dirHandle ? dirHandle.name : ''; }

  /* ---------- GitHub repo mode ---------- */
  function ghRepo() {
    return (typeof window !== 'undefined' && window.VG_CONFIG && window.VG_CONFIG.GITHUB_REPO) || '';
  }
  function ghBranch() {
    return (typeof window !== 'undefined' && window.VG_CONFIG && window.VG_CONFIG.GITHUB_BRANCH) || 'main';
  }
  function ghPath() { return 'data/' + FILE_NAME; }

  function b64encode(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i += 0x8000) {
      s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(s);
  }
  function b64decode(str) {
    var bin = atob(String(str).replace(/\s/g, ''));
    var out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function ghHeaders(token) {
    var h = { 'Accept': 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
    if (token) h['Authorization'] = 'Bearer ' + token;
    return h;
  }
  function ghMeta(token) {
    return fetch('https://api.github.com/repos/' + ghRepo() + '/contents/' + ghPath() + '?ref=' + ghBranch(), {
      headers: ghHeaders(token)
    }).then(function (r) {
      if (r.status === 404) return { exists: false };
      if (r.status === 401) throw new Error('401');
      if (!r.ok) throw new Error('http ' + r.status);
      return r.json().then(function (d) { return { exists: true, sha: d.sha, content: d.content }; });
    });
  }
  function ghLoad(token, cb) {
    if (!ghRepo()) { cb(null, 'nogithub'); return; }
    ghMeta(token).then(function (m) {
      if (!m.exists) { cb(null, 'missing'); return; }
      var state;
      try { state = parseWorkbook(b64decode(m.content).buffer); }
      catch (e) { cb(null, 'error'); return; }
      cb(state, 'github', m.sha);
    }).catch(function (err) { cb(null, err && err.message === '401' ? 'unauthorized' : 'error'); });
  }
  function ghPutBytes(token, bytes, sha, message, cb) {
    var payload = { message: message, content: b64encode(bytes), branch: ghBranch() };
    if (sha) payload.sha = sha;
    fetch('https://api.github.com/repos/' + ghRepo() + '/contents/' + ghPath(), {
      method: 'PUT',
      headers: Object.assign(ghHeaders(token), { 'Content-Type': 'application/json' }),
      body: JSON.stringify(payload)
    }).then(function (r) {
      if (!r.ok) {
        return r.json().catch(function () { return {}; }).then(function (e) {
          cb(false, 'http ' + r.status + (e && e.message ? ' ' + e.message : ''));
        });
      }
      return r.json().then(function (d) { cb(true, d.content && d.content.sha); });
    }).catch(function () { cb(false, 'network'); });
  }
  function ghSave(token, state, sha, cb) {
    ghPutBytes(token, workbookBytes(state), sha, 'donors & expenses update ' + new Date().toISOString(), cb);
  }
  function ghUploadBytes(token, bytes, cb) {
    ghMeta(token).then(function (m) {
      ghPutBytes(token, bytes, m.exists ? m.sha : null, 'upload donors.xlsx ' + new Date().toISOString(), cb);
    }).catch(function (err) { cb(false, err && err.message === '401' ? '401' : 'network'); });
  }
  function ghVerify(token) {
    return fetch('https://api.github.com/repos/' + ghRepo(), { headers: ghHeaders(token) })
      .then(function (r) { return r.ok; });
  }

  window.VGStore = {
    FILE_NAME: FILE_NAME,
    DCOL: DCOL,
    ECOL: ECOL,
    sum: sum,
    supported: supported,
    restore: restore,
    linkFolder: linkFolder,
    load: load,
    save: save,
    linked: linked,
    folderName: folderName,
    workbookBytes: workbookBytes,
    donorsBytes: donorsBytes,
    expensesBytes: expensesBytes,
    downloadBytes: downloadBytes,
    ghRepo: ghRepo,
    ghBranch: ghBranch,
    ghLoad: ghLoad,
    ghSave: ghSave,
    ghUploadBytes: ghUploadBytes,
    ghVerify: ghVerify
  };
})();
