/* ============================================================
   Data layer — the project file  data/donors.xlsx  is the ONLY
   store. No localStorage / sessionStorage for donor data.

   You connect the project's DATA FOLDER once (browser security
   requires that single permission). After that the app always
   reads & writes the one file "donors.xlsx" INSIDE that folder:
   - file exists  → used as-is
   - file missing → created automatically in the same folder
   The folder handle is remembered (IndexedDB), so every later
   visit re-attaches silently with zero clicks.
   ============================================================ */
(function () {
  'use strict';

  var FILE_NAME = 'donors.xlsx';
  var COL = { serial: 'வரிசை எண்', name: 'கொடுத்தவர் பெயர்', amount: 'அமௌன்ட் (₹)' };
  var SEED = [{ name: 'பூபதி இபி', amount: 1000 }];
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

  /* ---------- Excel parse / build (SheetJS) ---------- */
  function fromSheet(buf) {
    var wb = XLSX.read(buf, { type: 'array' });
    var ws = wb.Sheets[wb.SheetNames[0]];
    if (!ws) return [];
    var rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    return rows.map(function (r) {
      return {
        name: String(r[COL.name] !== undefined ? r[COL.name] : '').trim(),
        amount: Number(r[COL.amount]) || 0
      };
    }).filter(function (d) { return d.name || d.amount > 0; });
  }

  function toBytes(list) {
    var rows = list.map(function (d, i) {
      var o = {};
      o[COL.serial] = i + 1;
      o[COL.name] = d.name;
      o[COL.amount] = d.amount;
      return o;
    });
    var ws = XLSX.utils.json_to_sheet(rows, { header: [COL.serial, COL.name, COL.amount] });
    ws['!cols'] = [{ wch: 10 }, { wch: 32 }, { wch: 14 }];
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Donors');
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return new Uint8Array(out);
  }

  function toBlob(list) {
    return new Blob([toBytes(list)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /* ---------- the single file inside the linked data folder ---------- */
  function ensureFile() {
    /* create:true  → uses existing donors.xlsx, or creates it in the
       data folder when the file is not available                  */
    return dirHandle.getFileHandle(FILE_NAME, { create: true }).then(function (fh) {
      handle = fh;
    });
  }

  function readFile() {
    return handle.getFile()
      .then(function (f) { return f.arrayBuffer(); })
      .then(fromSheet);
  }

  /* ---------- silently re-attach remembered folder ---------- */
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

  /* ---------- one-time folder connect (from a user gesture) ---------- */
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

  /* ---------- read rows: linked file first, then http fetch ---------- */
  function load(cb) {
    if (handle) {
      readFile()
        .then(function (list) { cb(list, 'excel'); })
        .catch(function () { cb(null, 'error'); });
      return;
    }
    fetch('data/' + FILE_NAME, { cache: 'no-store' })
      .then(function (r) { if (!r.ok) throw new Error('http ' + r.status); return r.arrayBuffer(); })
      .then(function (buf) {
        var list;
        try { list = fromSheet(buf); } catch (e) { list = SEED.slice(); }
        cb(list, 'file');
      })
      .catch(function () { cb(null, 'none'); });
  }

  /* ---------- write rows into data/donors.xlsx ---------- */
  function save(list, cb) {
    if (!handle) { if (cb) cb(false, 'nolink'); return; }
    handle.createWritable()
      .then(function (w) { return w.write(toBlob(list)).then(function () { return w.close(); }); })
      .then(function () { if (cb) cb(true); })
      .catch(function () { if (cb) cb(false, 'locked'); }); /* e.g. Excel has the file open */
  }

  function linked() { return !!handle; }
  function folderName() { return dirHandle ? dirHandle.name : ''; }

  /* ---------- GitHub repo mode ---------------------------------
     Hosted site writes data/donors.xlsx into its own repository
     via the GitHub Contents API (one commit per change).
     The token is supplied at runtime by the admin page — it is
     never kept in the project code.                              */
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
      var list;
      try { list = fromSheet(b64decode(m.content).buffer); }
      catch (e) { cb(null, 'error'); return; }
      cb(list, 'github', m.sha);
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

  function ghSave(token, list, sha, cb) {
    ghPutBytes(token, toBytes(list), sha, 'donors update ' + new Date().toISOString(), cb);
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

  /* ---------- manual backup download (any browser) ---------- */
  function download(list) {
    var blob = toBlob(list);
    if (typeof URL !== 'undefined' && URL.createObjectURL) {
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = FILE_NAME;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
    }
  }

  window.VGStore = {
    FILE_NAME: FILE_NAME,
    COL: COL,
    supported: supported,
    restore: restore,
    linkFolder: linkFolder,
    load: load,
    save: save,
    linked: linked,
    folderName: folderName,
    download: download,
    toBlob: toBlob,
    toBytes: toBytes,
    ghRepo: ghRepo,
    ghBranch: ghBranch,
    ghLoad: ghLoad,
    ghSave: ghSave,
    ghUploadBytes: ghUploadBytes,
    ghVerify: ghVerify
  };
})();
