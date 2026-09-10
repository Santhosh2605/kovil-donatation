/* Home page: render donors + expenses read from data/donors.xlsx
   (both sheets of the one workbook). No browser storage. */
(function () {
  'use strict';

  var donorRowsEl   = document.getElementById('donorRows');
  var expenseRowsEl = document.getElementById('expenseRows');
  var totalEl       = document.getElementById('grandTotal');
  var expenseTotalEl = document.getElementById('expenseTotal');
  var balanceEl     = document.getElementById('balance');

  function fmt(n) { return Number(n || 0).toLocaleString('en-IN'); }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  function renderDonors(list) {
    if (!list || !list.length) {
      donorRowsEl.innerHTML = '<tr><td colspan="3" class="loading">இன்னும் நன்கொடைகள் இல்லை</td></tr>';
      totalEl.textContent = '0';
      return 0;
    }
    var total = 0;
    donorRowsEl.innerHTML = list.map(function (d, i) {
      total += Number(d.amount) || 0;
      return '<tr>' +
        '<td class="col-serial">' + (i + 1) + '</td>' +
        '<td class="donor-name">' + esc(d.name) + '</td>' +
        '<td class="col-amt">' + fmt(d.amount) + '</td>' +
      '</tr>';
    }).join('');
    totalEl.textContent = fmt(total);
    return total;
  }

  function renderExpenses(list) {
    if (!list || !list.length) {
      expenseRowsEl.innerHTML = '<tr><td colspan="4" class="loading">இன்னும் செலவுகள் இல்லை</td></tr>';
      expenseTotalEl.textContent = '0';
      return 0;
    }
    var total = 0;
    expenseRowsEl.innerHTML = list.map(function (d, i) {
      total += Number(d.amount) || 0;
      return '<tr>' +
        '<td class="col-serial">' + (i + 1) + '</td>' +
        '<td>' + esc(d.date) + '</td>' +
        '<td class="donor-name">' + esc(d.desc) + '</td>' +
        '<td class="col-amt">' + fmt(d.amount) + '</td>' +
      '</tr>';
    }).join('');
    expenseTotalEl.textContent = fmt(total);
    return total;
  }

  function render(state) {
    var dt = renderDonors(state.donors);
    var et = renderExpenses(state.expenses);
    balanceEl.textContent = fmt(dt - et);
  }

  function refresh() {
    VGStore.load(function (state, status) {
      if (status === 'excel' || status === 'file') { render(state); return; }
      /* no linked handle and no http fetch (e.g. opened via file://) */
      donorRowsEl.innerHTML =
        '<tr><td colspan="3" class="loading">பட்டியலைக் காட்ட ' +
        '<button id="homeLinkBtn" class="btn btn-sm btn-fest">📂 donors.xlsx இணை</button>' +
        '</td></tr>';
      expenseRowsEl.innerHTML = '<tr><td colspan="4" class="loading">—</td></tr>';
      totalEl.textContent = '0';
      expenseTotalEl.textContent = '0';
      balanceEl.textContent = '0';
    });
  }

  document.addEventListener('click', function (e) {
    if (e.target && e.target.id === 'homeLinkBtn') {
      VGStore.linkFolder(function (ok) { if (ok) refresh(); });
    }
  });

  VGStore.restore(false, function () { refresh(); });
  setInterval(refresh, 25000);
})();
