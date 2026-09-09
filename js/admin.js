/* Admin context: login + add / update / delete donors.
   Two storage modes (chosen by js/config.js → GITHUB_REPO):
   - GitHub mode : data/donors.xlsx lives in the GitHub project;
     first-time upload from PC, then one commit per change.
   - Local mode  : data folder on this computer (File System Access).
   Donor data is never kept in browser storage; only the login flag
   and (GitHub mode) the runtime token live in sessionStorage.      */
(function () {
  'use strict';

  var CFG = window.VG_CONFIG || { ADMIN_USER: 'admin', ADMIN_PASS: 'admin123' };
  var SESSION_KEY = 'vg_admin_session';
  var TOKEN_KEY = 'vg_gh_token';
  var GH = !!VGStore.ghRepo();

  function $(id) { return document.getElementById(id); }
  function fmt(n) { return Number(n || 0).toLocaleString('en-IN'); }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  var list = [];
  var ghToken = null;
  try { ghToken = sessionStorage.getItem(TOKEN_KEY); } catch (e) {}
  var ghSha = null;

  /* ---------- toast ---------- */
  var toastInstance = null;
  function toast(msg, ok) {
    if (ok === undefined) ok = true;
    if (!toastInstance) toastInstance = new bootstrap.Toast($('appToast'), { delay: 2600 });
    $('toastMsg').textContent = msg;
    $('appToast').classList.toggle('text-bg-success', ok);
    $('appToast').classList.toggle('text-bg-danger', !ok);
    toastInstance.show();
  }

  /* ---------- status lines ---------- */
  function setStatus(kind) {
    var el = $('excelStatus');
    if (kind === 'linked') {
      el.textContent = '✅ இணைக்கப்பட்டது: ' + (VGStore.folderName() || 'data') + '/donors.xlsx — ஒவ்வொரு மாற்றமும் இந்த ஒரே கோப்பில் எழுதப்படும்.';
      el.className = 'excel-status text-success small mt-2 mb-0';
    } else if (kind === 'locked') {
      el.textContent = '⚠️ Excel கோப்பை எழுத முடியவில்லை — donors.xlsx எக்செல்-ல் திறந்திருந்தால் அதை மூடிவிட்டு மீண்டும் முயற்சிக்கவும்.';
      el.className = 'excel-status text-danger small mt-2 mb-0';
    } else if (kind === 'unsupported') {
      el.textContent = '⚠️ இந்த உலாவி நேரடி Excel எழுத்தை ஆதரிக்கவில்லை (Chrome/Edge பயன்படுத்தவும்). மாற்றங்களை ⬇ பதிவிறக்கு மூலம் data/donors.xlsx ஆகச் சேமிக்கவும்.';
      el.className = 'excel-status text-danger small mt-2 mb-0';
    } else {
      el.textContent = '📁 உங்கள் project இன் DATA FOLDER ஐ ஒரு முறை இணைக்கவும் — அதன் பிறகு donors.xlsx அங்கேயே தானாக உருவாக்கப்பட்டு, எல்லா தரவும் அதில் மட்டுமே சேமிக்கப்படும்.';
      el.className = 'excel-status text-secondary small mt-2 mb-0';
    }
  }
  function ghStatus(text, cls) {
    var el = $('ghStatus');
    el.textContent = text;
    el.className = 'small mt-2 mb-0 ' + (cls || 'text-secondary');
  }

  /* ---------- view switching ---------- */
  function showLogin() {
    try { sessionStorage.removeItem(SESSION_KEY); } catch (e) {}
    $('dashView').classList.add('d-none');
    $('loginView').classList.remove('d-none');
    $('logoutBtn').classList.add('d-none');
  }
  function showDash() {
    $('loginView').classList.add('d-none');
    $('dashView').classList.remove('d-none');
    $('logoutBtn').classList.remove('d-none');
  }

  /* ---------- login / logout ---------- */
  $('loginForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var u = $('loginUser').value.trim();
    var p = $('loginPass').value;
    if (u === CFG.ADMIN_USER && p === CFG.ADMIN_PASS) {
      try { sessionStorage.setItem(SESSION_KEY, '1'); } catch (err) {}
      $('loginError').classList.add('d-none');
      $('loginForm').reset();
      showDash();
      toast('உள்நுழைவு வெற்றி 🙏');
      if (GH) { ghInit(); } else { silentAttach(); }
    } else {
      $('loginError').classList.remove('d-none');
    }
  });

  $('logoutBtn').addEventListener('click', showLogin);

  /* ================= GitHub mode ================= */
  function ghInit() {
    $('ghPanel').classList.remove('d-none');
    $('linkExcelBtn').classList.add('d-none');
    if (ghToken) {
      ghLoadCurrent();
    } else {
      ghStatus('🔑 Fine-grained token உள்ளிட்டு இணைக்கவும் (GitHub → Settings → Developer settings → Tokens → repo: ' + VGStore.ghRepo() + ', Contents: Read and write)', 'text-secondary');
      load();
    }
  }

  $('ghConnect').addEventListener('click', function () {
    var t = $('ghToken').value.trim();
    if (!t) { ghStatus('token தேவை', 'text-danger'); return; }
    VGStore.ghVerify(t).then(function (ok) {
      if (!ok) { ghStatus('⚠️ token / repo சரிபார்ப்பு தவறு', 'text-danger'); return; }
      ghToken = t;
      try { sessionStorage.setItem(TOKEN_KEY, t); } catch (e) {}
      $('ghToken').value = '';
      ghStatus('✅ GitHub இணைப்பு: ' + VGStore.ghRepo(), 'text-success');
      ghLoadCurrent();
    });
  });

  function ghLoadCurrent() {
    VGStore.ghLoad(ghToken, function (l, status, sha) {
      if (status === 'github') {
        ghSha = sha; list = l || []; render();
        ghStatus('✅ ' + VGStore.ghRepo() + '/data/donors.xlsx — ஒவ்வொரு மாற்றமும் நேரடியாக repo-வில் commit ஆகும்.', 'text-success');
      } else if (status === 'missing') {
        list = []; render();
        ghStatus('⚠️ Repo-வில் donors.xlsx இல்லை — “⬆ Excel பதிவேற்று” மூலம் முதல் முறை உங்கள் PC-யிலிருந்து பதிவேற்றவும்.', 'text-danger');
      } else if (status === 'unauthorized') {
        ghToken = null;
        try { sessionStorage.removeItem(TOKEN_KEY); } catch (e) {}
        list = []; render();
        ghStatus('🔑 token தேவை — மீண்டும் இணைக்கவும்.', 'text-secondary');
      } else {
        list = []; render();
        ghStatus('⚠️ GitHub-லிருந்து படிக்க முடியவில்லை', 'text-danger');
      }
    });
  }

  /* first-time upload: PC → straight into the GitHub project */
  $('ghUpload').addEventListener('change', function () {
    var f = this.files && this.files[0];
    this.value = '';
    if (!f) return;
    if (!ghToken) { ghStatus('முதலில் 🔑 token இணைக்கவும்', 'text-danger'); return; }
    f.arrayBuffer().then(function (buf) {
      VGStore.ghUploadBytes(ghToken, new Uint8Array(buf), function (ok, shaOrErr) {
        if (ok) {
          ghSha = shaOrErr;
          toast('GitHub repo-வில் donors.xlsx பதிவேற்றப்பட்டது ✅');
          ghLoadCurrent();
        } else {
          ghStatus('⚠️ பதிவேற்ற முடியவில்லை: ' + shaOrErr, 'text-danger');
        }
      });
    });
  });

  /* ================= local (data folder) mode ================= */
  function connectOnce() {
    VGStore.linkFolder(function (ok2, reason) {
      if (ok2) {
        setStatus('linked'); load();
        toast('data folder இணைக்கப்பட்டது — donors.xlsx இனி தானியங்கு ✅');
      } else if (reason === 'cancelled') {
        setStatus('unlinked'); load();
      } else if (reason === 'unsupported') {
        setStatus('unsupported'); load();
      } else {
        setStatus('unlinked'); load();
        toast('உலாவி அனுமதி கிடைக்கவில்லை — மீண்டும் முயற்சிக்கவும்', false);
      }
    });
  }

  function acquire() {
    if (!VGStore.supported()) { setStatus('unsupported'); load(); return; }
    VGStore.restore(true, function (ok) {
      if (ok) { setStatus('linked'); load(); return; } /* silent auto re-attach */
      connectOnce(); /* one-time only */
    });
  }

  function silentAttach() {
    if (!VGStore.supported()) { setStatus('unsupported'); load(); return; }
    VGStore.restore(false, function (ok) {
      if (ok) { setStatus('linked'); } else { setStatus('unlinked'); }
      load();
    });
  }

  $('linkExcelBtn').addEventListener('click', connectOnce);

  /* ================= shared list handling ================= */
  function load() {
    if (GH) {
      if (ghToken) { ghLoadCurrent(); } else { list = []; render(); }
      return;
    }
    VGStore.load(function (l, status) {
      if (status === 'none' || status === 'error') {
        if (!VGStore.linked()) setStatus('unlinked');
        list = [];
      } else {
        list = l || [];
      }
      render();
    });
  }

  function render() {
    var tbody = $('adminRows');
    if (!list.length) {
      tbody.innerHTML = '<tr><td colspan="4" class="loading">பட்டியல் காலியாக உள்ளது</td></tr>';
    } else {
      tbody.innerHTML = list.map(function (d, i) {
        return '<tr>' +
          '<td class="col-serial">' + (i + 1) + '</td>' +
          '<td class="donor-name">' + esc(d.name) + '</td>' +
          '<td class="col-amt">' + fmt(d.amount) + '</td>' +
          '<td class="text-end" style="white-space:nowrap">' +
            '<button class="btn btn-sm btn-outline-primary action-btn" data-act="edit" data-serial="' + (i + 1) + '" data-name="' + esc(d.name) + '" data-amount="' + d.amount + '" title="திருத்து">✏️</button> ' +
            '<button class="btn btn-sm btn-outline-danger action-btn" data-act="del" data-serial="' + (i + 1) + '" data-name="' + esc(d.name) + '" title="நீக்கு">🗑️</button>' +
          '</td>' +
        '</tr>';
      }).join('');
    }
    var total = list.reduce(function (s, d) { return s + (Number(d.amount) || 0); }, 0);
    $('adminTotal').textContent = fmt(total);
  }

  function persist(msg) {
    render();
    if (GH) {
      if (!ghToken) {
        ghStatus('🔑 முதலில் token இணைக்கவும்', 'text-danger');
        toast('GitHub token தேவை', false);
        return;
      }
      VGStore.ghSave(ghToken, list, ghSha, function (ok, shaOrErr) {
        if (ok) {
          ghSha = shaOrErr;
          ghStatus('✅ commit ஆனது: ' + VGStore.ghRepo() + '/data/donors.xlsx', 'text-success');
          toast(msg + ' (GitHub ✅)');
        } else if (String(shaOrErr).indexOf('http 409') === 0) {
          /* someone else committed — refresh sha and retry once */
          VGStore.ghLoad(ghToken, function (l, st, sha) {
            if (st === 'github') {
              ghSha = sha;
              VGStore.ghSave(ghToken, list, ghSha, function (ok2, err2) {
                if (ok2) { ghSha = err2; toast(msg + ' (GitHub ✅)'); }
                else { ghStatus('⚠️ ' + err2, 'text-danger'); toast('சேமிக்க முடியவில்லை', false); }
              });
            }
          });
        } else if (String(shaOrErr).indexOf('http 401') === 0) {
          ghToken = null;
          try { sessionStorage.removeItem(TOKEN_KEY); } catch (e) {}
          ghStatus('🔑 token மீண்டும் தேவை', 'text-danger');
          toast('அனுமதி இல்லை — மீண்டும் இணைக்கவும்', false);
        } else {
          ghStatus('⚠️ சேமிக்க முடியவில்லை: ' + shaOrErr, 'text-danger');
          toast('சேமிக்க முடியவில்லை', false);
        }
      });
      return;
    }
    VGStore.save(list, function (ok, err) {
      if (ok) { setStatus('linked'); toast(msg + ' (Excel ✅)'); }
      else if (err === 'locked') { setStatus('locked'); toast('Excel-ல் சேமிக்க முடியவில்லை — கோப்பை மூடிவிட்டு மீண்டும் முயற்சிக்கவும்', false); }
      else { setStatus('unsupported'); toast('இணைக்கப்படவில்லை — ⬇ பதிவிறக்கு பயன்படுத்தவும்', false); }
    });
  }

  /* ---------- add ---------- */
  $('addForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var name = $('addName').value.trim();
    var amount = Math.round(Number($('addAmount').value));
    if (!name) { toast('பெயர் தேவை', false); return; }
    if (!isFinite(amount) || amount <= 0) { toast('சரியான தொகையை உள்ளிடவும்', false); return; }
    list.push({ name: name, amount: amount });
    $('addForm').reset();
    persist('நன்கொடை சேர்க்கப்பட்டது');
  });

  /* ---------- edit / delete via modals ---------- */
  function editModal()   { return bootstrap.Modal.getOrCreateInstance($('editModal')); }
  function deleteModal() { return bootstrap.Modal.getOrCreateInstance($('deleteModal')); }

  $('adminRows').addEventListener('click', function (e) {
    var btn = e.target.closest('button[data-act]');
    if (!btn) return;
    if (btn.getAttribute('data-act') === 'edit') {
      $('editSerial').value = btn.getAttribute('data-serial');
      $('editName').value   = btn.getAttribute('data-name');
      $('editAmount').value = btn.getAttribute('data-amount');
      editModal().show();
    } else {
      $('deleteSerial').value = btn.getAttribute('data-serial');
      $('deleteName').textContent = btn.getAttribute('data-name');
      deleteModal().show();
    }
  });

  $('editForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var serial = Number($('editSerial').value);
    var name = $('editName').value.trim();
    var amount = Math.round(Number($('editAmount').value));
    if (!name || !isFinite(amount) || amount <= 0) { toast('சரியான தரவை உள்ளிடவும்', false); return; }
    if (serial >= 1 && serial <= list.length) {
      list[serial - 1] = { name: name, amount: amount };
      editModal().hide();
      persist('மாற்றம் சேமிக்கப்பட்டது');
    }
  });

  $('deleteForm').addEventListener('submit', function (e) {
    e.preventDefault();
    var serial = Number($('deleteSerial').value);
    if (serial >= 1 && serial <= list.length) {
      list.splice(serial - 1, 1);
      deleteModal().hide();
      persist('நன்கொடை நீக்கப்பட்டது');
    }
  });

  /* ---------- manual backup download ---------- */
  $('downloadExcelBtn').addEventListener('click', function () {
    if (!list.length) { toast('பட்டியல் காலியாக உள்ளது', false); return; }
    VGStore.download(list);
    toast('donors.xlsx பதிவிறக்கப்பட்டது ⬇');
  });

  /* ---------- boot ---------- */
  var session = null;
  try { session = sessionStorage.getItem(SESSION_KEY); } catch (e) {}
  if (session) {
    showDash();
    if (GH) { ghInit(); } else { silentAttach(); }
  } else {
    showLogin();
  }
})();
