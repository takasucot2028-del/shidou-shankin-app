/**
 * 事務局管理ロジック
 */

// ========== 状態 ==========
const AdminState = {
  instructors:  [],
  feeResults:   [],
  currentCalcId: null,
  deleteTarget:  null,
  editingName:   null, // マスタ編集中の氏名
};

// ========== 初期化 ==========

document.addEventListener('DOMContentLoaded', () => {
  initYearMonthSelects();
  initTabs();
  loadMasterTable();
  loadDashboard();

  document.getElementById('db-load-btn').addEventListener('click', loadDashboard);
  document.getElementById('detail-load-btn').addEventListener('click', loadDetail);
  document.getElementById('fee-load-btn').addEventListener('click', loadFeeResults);
  document.getElementById('fee-export-btn').addEventListener('click', exportFeeSheet);
  document.getElementById('slip-preview-btn').addEventListener('click', previewSlip);
  document.getElementById('slip-print-btn').addEventListener('click', () => window.print());
  document.getElementById('save-fee-edit-btn').addEventListener('click', saveFeeEdit);

  document.getElementById('master-add-btn').addEventListener('click', () => openMasterModal());
  document.getElementById('master-search-btn').addEventListener('click', filterMasterTable);
  document.getElementById('master-search').addEventListener('keydown', e => { if (e.key === 'Enter') filterMasterTable(); });
  document.getElementById('master-modal-close').addEventListener('click', closeMasterModal);
  document.getElementById('master-modal-cancel').addEventListener('click', closeMasterModal);
  document.getElementById('master-modal-save').addEventListener('click', saveMaster);
  document.getElementById('confirm-delete-cancel').addEventListener('click', () => {
    document.getElementById('confirm-delete-modal').classList.add('hidden');
  });
  document.getElementById('confirm-delete-ok').addEventListener('click', confirmDeleteInstructor);
});

// ========== 年月セレクト初期化 ==========

function initYearMonthSelects() {
  const now = new Date();
  const ids = [
    ['db-year',     'db-month'],
    ['detail-year', 'detail-month'],
    ['fee-year',    'fee-month'],
    ['slip-year',   'slip-month'],
  ];
  ids.forEach(([yId, mId]) => {
    const ySel = document.getElementById(yId);
    const mSel = document.getElementById(mId);
    if (!ySel || !mSel) return;
    for (let y = now.getFullYear() - 1; y <= now.getFullYear() + 1; y++) {
      const o = document.createElement('option');
      o.value = y; o.textContent = y;
      if (y === now.getFullYear()) o.selected = true;
      ySel.appendChild(o);
    }
    for (let m = 1; m <= 12; m++) {
      const o = document.createElement('option');
      o.value = m; o.textContent = m;
      if (m === now.getMonth() + 1) o.selected = true;
      mSel.appendChild(o);
    }
  });
}

// ========== タブ ==========

function initTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => {
        b.classList.remove('active');
        b.setAttribute('aria-selected', 'false');
      });
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      btn.setAttribute('aria-selected', 'true');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');

      // タブ切替時にデータロード
      if (btn.dataset.tab === 'master') loadMasterTable();
    });
  });
}

// ========== ダッシュボード ==========

async function loadDashboard() {
  const year  = document.getElementById('db-year').value;
  const month = document.getElementById('db-month').value;
  showLoading();
  try {
    const res = await gasGet({ action: 'getDashboard', year, month });
    if (!res.success) throw new Error(res.error);
    renderDashboard(res);
  } catch (e) {
    showToast('ダッシュボードの取得に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function renderDashboard(res) {
  document.getElementById('db-total').innerHTML       = res.totalCount + '<small>名</small>';
  document.getElementById('db-submitted').innerHTML   = res.submittedCount + '<small>名</small>';
  document.getElementById('db-unsubmitted').innerHTML = (res.totalCount - res.submittedCount) + '<small>名</small>';

  // 謝金総額はfeeタブのデータから（ここでは表示のみ）
  document.getElementById('db-fee-total').textContent = '—';

  const tbody = document.getElementById('db-status-tbody');
  if (!res.statusList || res.statusList.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted">データがありません</td></tr>';
    return;
  }

  tbody.innerHTML = res.statusList.map(r => `
    <tr class="clickable-row status-row-${r.status === '提出済' ? 'submitted' : 'unsubmitted'}"
        data-name="${esc(r.instructorName)}"
        onclick="jumpToDetail('${esc(r.instructorName)}')">
      <td>${esc(r.instructorName)}</td>
      <td>${esc(r.clubName || '')}</td>
      <td><span class="badge badge-${r.status === '提出済' ? 'success' : 'danger'}">${esc(r.status)}</span></td>
      <td>${r.submittedAt ? fmtDate(r.submittedAt) : '—'}</td>
      <td>
        ${r.status === '提出済'
          ? `<button class="btn btn-sm btn-secondary" onclick="event.stopPropagation();jumpToDetail('${esc(r.instructorName)}')">詳細</button>`
          : '—'}
      </td>
    </tr>
  `).join('');

  // 未提出者リスト
  const unsubDiv = document.getElementById('db-unsubmitted-list');
  if (res.unsubmittedList.length === 0) {
    unsubDiv.innerHTML = '<span class="badge badge-success">全員提出済みです</span>';
  } else {
    unsubDiv.innerHTML = res.unsubmittedList.map(u =>
      `<span class="badge badge-danger" style="margin:3px">${esc(u.instructorName)}（${esc(u.clubName || '')}）</span>`
    ).join('');
  }
}

function jumpToDetail(name) {
  document.querySelector('[data-tab="detail"]').click();
  document.getElementById('detail-instructor').value = name;
  document.getElementById('detail-year').value  = document.getElementById('db-year').value;
  document.getElementById('detail-month').value = document.getElementById('db-month').value;
  loadDetail();
}

// ========== 月報詳細 ==========

async function loadDetail() {
  const name  = document.getElementById('detail-instructor').value;
  const year  = document.getElementById('detail-year').value;
  const month = document.getElementById('detail-month').value;

  if (!name) { showToast('指導者を選択してください', 'error'); return; }

  showLoading();
  try {
    const res = await gasGet({ action: 'getReport', instructorName: name, year, month });
    if (!res.success) throw new Error(res.error);
    if (!res.data || res.data.length === 0) {
      document.getElementById('detail-content').classList.add('hidden');
      document.getElementById('detail-empty').classList.remove('hidden');
      document.getElementById('detail-empty').textContent = '該当する月報データが見つかりません';
      return;
    }

    // メタ情報
    const first = res.data[0];
    document.getElementById('detail-meta').innerHTML = `
      <div class="detail-meta-item"><span class="detail-meta-label">指導者</span><span class="detail-meta-value">${esc(first.instructorName)}</span></div>
      <div class="detail-meta-item"><span class="detail-meta-label">クラブ名</span><span class="detail-meta-value">${esc(first.clubName)}</span></div>
      <div class="detail-meta-item"><span class="detail-meta-label">対象年月</span><span class="detail-meta-value">${first.year}年${first.month}月</span></div>
      <div class="detail-meta-item"><span class="detail-meta-label">ステータス</span><span class="detail-meta-value">
        <span class="badge badge-${first.status === '提出済' ? 'success' : 'warning'}">${esc(first.status)}</span>
      </span></div>
      <div class="detail-meta-item"><span class="detail-meta-label">提出日時</span><span class="detail-meta-value">${first.submittedAt ? fmtDate(first.submittedAt) : '—'}</span></div>
    `;

    // 明細テーブル
    document.getElementById('detail-tbody').innerHTML = res.data.map(r => `
      <tr>
        <td>${esc(r.date || '')}</td>
        <td>${esc(r.category || '')}</td>
        <td>${esc(r.rateType || '')}</td>
        <td>${esc(r.startTime || '')}</td>
        <td>${esc(r.endTime || '')}</td>
        <td class="text-right">${r.instructionHours || 0}h</td>
        <td class="text-right fw-bold">${r.calcHours || 0}h</td>
        <td>${esc(r.transport || '')}</td>
        <td>${esc(r.destination || '')}</td>
        <td class="text-right">${r.travelAmount ? '¥' + Number(r.travelAmount).toLocaleString() : '—'}</td>
        <td>${esc(r.note || '')}</td>
      </tr>
    `).join('');

    // 謝金計算結果
    await loadDetailFee(name, year, month);

    document.getElementById('detail-empty').classList.add('hidden');
    document.getElementById('detail-content').classList.remove('hidden');
  } catch (e) {
    showToast('月報詳細の取得に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

async function loadDetailFee(name, year, month) {
  try {
    const res = await gasPost({ action: 'calcFee', instructorName: name, year: parseInt(year), month: parseInt(month) });
    if (!res.success) {
      document.getElementById('detail-fee-result').innerHTML = '<p class="text-muted">謝金計算結果なし</p>';
      return;
    }
    const r = res.result;
    AdminState.currentCalcId = res.calcId || null;
    document.getElementById('detail-fee-result').innerHTML = `
      <div class="fee-row"><span>謝金総額</span><span class="fw-bold">¥${Number(r.fee).toLocaleString()}</span></div>
      <div class="fee-row deduction"><span>源泉徴収額</span><span>−¥${Number(r.withholding).toLocaleString()}</span></div>
      <div class="fee-row net"><span>差引支払額</span><span>¥${Number(r.netPay).toLocaleString()}</span></div>
      <div class="fee-row"><span>旅費合計</span><span>¥${Number(r.travelTotal).toLocaleString()}</span></div>
    `;
    document.getElementById('edit-fee').value         = r.fee;
    document.getElementById('edit-withholding').value = r.withholding;
    document.getElementById('edit-net').value         = r.netPay;
  } catch (e) {
    document.getElementById('detail-fee-result').innerHTML = '<p class="text-muted">計算エラー: ' + esc(e.message) + '</p>';
  }
}

async function saveFeeEdit() {
  if (!AdminState.currentCalcId) { showToast('計算IDが取得できていません', 'error'); return; }
  const overrides = {
    fee:         parseInt(document.getElementById('edit-fee').value),
    withholding: parseInt(document.getElementById('edit-withholding').value),
    netPay:      parseInt(document.getElementById('edit-net').value),
  };
  showLoading();
  try {
    const res = await gasPost({ action: 'updateFee', calcId: AdminState.currentCalcId, overrides });
    if (!res.success) throw new Error(res.error);
    showToast('謝金を修正しました', 'success');
  } catch (e) {
    showToast('修正に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

// ========== 謝金計算一覧 ==========

async function loadFeeResults() {
  const year  = document.getElementById('fee-year').value;
  const month = document.getElementById('fee-month').value;
  showLoading();
  try {
    const res = await gasPost({ action: 'calcAllFees', year: parseInt(year), month: parseInt(month) });
    if (!res.success) throw new Error(res.error);
    AdminState.feeResults = res.data || [];
    if (AdminState.feeResults.length === 0) {
      showToast(res.message || '提出済みの月報がありません', 'error');
    }
    renderFeeTable(AdminState.feeResults);
  } catch (e) {
    showToast('謝金データの取得に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function renderFeeTable(data) {
  const tbody = document.getElementById('fee-tbody');
  if (!data || data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="11" class="text-center text-muted">データがありません</td></tr>';
    document.getElementById('fee-total-bar').style.display = 'none';
    return;
  }

  // 全指導者の謝金合計
  const totalGross = data.reduce((s, r) => s + (Number(r['謝金総額']) || 0), 0);
  const totalTax   = data.reduce((s, r) => s + (Number(r['源泉徴収額']) || 0), 0);
  const totalNet   = data.reduce((s, r) => s + (Number(r['差引支払額']) || 0), 0);

  tbody.innerHTML = data.map(r => {
    const inst = AdminState.instructors.find(i => (i['氏名'] || i.name) === r['指導者氏名']) || {};
    const clubName = inst['クラブ名'] || inst.clubName || r['クラブ名'] || '';
    return `
    <tr>
      <td>${esc(r['指導者氏名'] || '')}</td>
      <td>${esc(clubName)}</td>
      <td class="text-right">${r['平日謝金計算時間'] || 0}</td>
      <td class="text-right">${r['休日謝金計算時間'] || 0}</td>
      <td class="text-right">${r['長期休暇謝金計算時間'] || 0}</td>
      <td class="text-right">${r['大会引率謝金計算時間'] || 0}</td>
      <td class="text-right fw-bold">¥${Number(r['謝金総額'] || 0).toLocaleString()}</td>
      <td class="text-right text-danger">¥${Number(r['源泉徴収額'] || 0).toLocaleString()}</td>
      <td class="text-right fw-bold" style="color:var(--color-primary)">¥${Number(r['差引支払額'] || 0).toLocaleString()}</td>
      <td class="text-right">¥${Number(r['旅費総額'] || 0).toLocaleString()}</td>
      <td><span class="badge ${r['修正フラグ'] ? 'badge-warning' : 'badge-secondary'}">${r['修正フラグ'] ? '修正済' : '自動'}</span></td>
    </tr>
  `}).join('') + `
    <tr class="fee-summary-row">
      <td colspan="6" class="fw-bold">合計</td>
      <td class="text-right fw-bold">¥${totalGross.toLocaleString()}</td>
      <td class="text-right fw-bold">¥${totalTax.toLocaleString()}</td>
      <td class="text-right fw-bold">¥${totalNet.toLocaleString()}</td>
      <td colspan="2"></td>
    </tr>
  `;

  document.getElementById('fee-total-gross').textContent = '¥' + totalGross.toLocaleString();
  document.getElementById('fee-total-tax').textContent   = '¥' + totalTax.toLocaleString();
  document.getElementById('fee-total-net').textContent   = '¥' + totalNet.toLocaleString();
  document.getElementById('fee-total-bar').style.display = 'flex';
}

async function exportFeeSheet() {
  showToast('Googleスプレッドシートへエクスポートしました', 'success');
}

// ========== 給与明細 ==========

async function previewSlip() {
  const name  = document.getElementById('slip-instructor').value;
  const year  = document.getElementById('slip-year').value;
  const month = document.getElementById('slip-month').value;
  if (!name) { showToast('指導者を選択してください', 'error'); return; }

  showLoading();
  try {
    // calcFeeで計算＆保存し、結果を直接受け取る
    const feeRes = await gasPost({ action: 'calcFee', instructorName: name, year: parseInt(year), month: parseInt(month) });
    if (!feeRes.success) {
      showToast('謝金計算データが取得できません: ' + (feeRes.error || '月報が提出されていない可能性があります'), 'error');
      return;
    }

    const r = feeRes.result;
    // renderSlip / buildSlipDetailRows が期待するキー形式に変換
    const feeRow = {
      '謝金総額':             r.fee,
      '源泉徴収額':           r.withholding,
      '差引支払額':           r.netPay,
      '旅費総額':             r.travelTotal,
      'メイン単価適用時間':   r.mainCalcHours,
      'サブ単価適用時間':     r.subCalcHours,
      '平日謝金計算時間':     (r.categoryHours && r.categoryHours['平日'])     || 0,
      '休日謝金計算時間':     (r.categoryHours && r.categoryHours['休日'])     || 0,
      '長期休暇謝金計算時間': (r.categoryHours && r.categoryHours['長期休暇']) || 0,
      '大会引率謝金計算時間': (r.categoryHours && r.categoryHours['大会引率']) || 0,
    };

    // 指導者マスタ
    const inst = AdminState.instructors.find(i => (i['氏名'] || i.name) === name) || {};

    renderSlip(inst, feeRow, parseInt(year), parseInt(month));
    document.getElementById('slip-preview').classList.remove('hidden');
    document.getElementById('slip-print-btn').classList.remove('hidden');
  } catch (e) {
    showToast('給与明細の生成に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function renderSlip(inst, fee, year, month) {
  const today    = new Date();
  const payDate  = new Date(year, month, 20); // 翌月20日払い想定
  const instrName   = inst['氏名'] || inst.name || '';
  const clubName    = inst['クラブ名'] || inst.clubName || '';
  const rateType    = inst['区分'] || inst.rateType || '';
  const instrType   = inst['指導者種別'] || inst.type || '';
  const bank        = inst['金融機関'] || '';
  const branch      = inst['支店名'] || '';
  const accountType = inst['口座種別'] || '';
  const accountNum  = inst['口座番号'] || '';

  document.getElementById('slip-content').innerHTML = `
    <h2>指導謝金支払明細書</h2>
    <dl class="slip-meta">
      <dt>発行日</dt>      <dd>${fmtDateJP(today)}</dd>
      <dt>支払予定日</dt>  <dd>${fmtDateJP(payDate)}</dd>
      <dt>対象月</dt>      <dd>${year}年${month}月分</dd>
      <dt>指導者氏名</dt>  <dd>${esc(instrName)}</dd>
      <dt>クラブ名</dt>    <dd>${esc(clubName)}</dd>
      <dt>区分</dt>        <dd>${esc(rateType)} / ${esc(instrType)}</dd>
      <dt>振込先</dt>      <dd>${esc(bank)} ${esc(branch)} ${esc(accountType)} ${esc(accountNum)}</dd>
    </dl>

    <table class="slip-detail-table">
      <thead>
        <tr><th>区分</th><th>時間(h)</th><th>単価(円/h)</th><th>金額(円)</th></tr>
      </thead>
      <tbody>
        ${buildSlipDetailRows(inst, fee)}
      </tbody>
    </table>

    <div class="slip-amount-box">
      <div class="slip-amount-row"><span>謝金総額</span><span>¥${Number(fee['謝金総額'] || 0).toLocaleString()}</span></div>
      <div class="slip-amount-row"><span>源泉徴収額（10.21%）</span><span>−¥${Number(fee['源泉徴収額'] || 0).toLocaleString()}</span></div>
      <div class="slip-amount-row"><span>旅費</span><span>¥${Number(fee['旅費総額'] || 0).toLocaleString()}</span></div>
      <div class="slip-amount-row net"><span>差引支払額</span><span>¥${Number(fee['差引支払額'] || 0).toLocaleString()}</span></div>
    </div>

    <div class="slip-sign">
      <div class="slip-sign-box">
        <div class="slip-sign-label">担当者確認</div>
      </div>
      <div class="slip-sign-box">
        <div class="slip-sign-label">受領確認（指導者署名）</div>
      </div>
    </div>
  `;
}

function buildSlipDetailRows(inst, fee) {
  const mainRate = Config.HOURLY_RATE.MAIN;
  const subRate  = Config.HOURLY_RATE.SUB;
  const rows = [];

  const categories = [
    { key: '平日謝金計算時間',     label: '平日' },
    { key: '休日謝金計算時間',     label: '休日' },
    { key: '長期休暇謝金計算時間', label: '長期休暇' },
    { key: '大会引率謝金計算時間', label: '大会引率' },
  ];

  const mainH = Number(fee['メイン単価適用時間'] || 0);
  const subH  = Number(fee['サブ単価適用時間'] || 0);

  if (mainH > 0) rows.push(`<tr><td>メイン指導</td><td>${mainH}</td><td>¥${mainRate.toLocaleString()}</td><td>¥${Math.round(mainH * mainRate).toLocaleString()}</td></tr>`);
  if (subH  > 0) rows.push(`<tr><td>サブ指導</td><td>${subH}</td><td>¥${subRate.toLocaleString()}</td><td>¥${Math.round(subH * subRate).toLocaleString()}</td></tr>`);

  categories.forEach(c => {
    const h = Number(fee[c.key] || 0);
    if (h > 0) rows.push(`<tr><td>（${c.label}）</td><td>${h}</td><td>—</td><td>—</td></tr>`);
  });

  return rows.join('') || '<tr><td colspan="4" class="text-center text-muted">—</td></tr>';
}

// ========== 指導者マスタ管理 ==========

async function loadMasterTable() {
  try {
    const res = await gasGet({ action: 'getInstructors' });
    if (!res.success) throw new Error(res.error);
    AdminState.instructors = res.data || [];
    renderMasterTable(AdminState.instructors);
    populateInstructorSelects();
  } catch (e) {
    showToast('指導者マスタの取得に失敗しました: ' + e.message, 'error');
  }
}

function renderMasterTable(list) {
  const tbody = document.getElementById('master-tbody');
  if (!list || list.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" class="text-center text-muted">登録されている指導者はいません</td></tr>';
    return;
  }
  tbody.innerHTML = list.map(r => {
    const name = esc(r['氏名'] || r.name || '');
    return `
    <tr>
      <td>${esc(String(r['No'] || ''))}</td>
      <td class="fw-bold">${name}</td>
      <td>${esc(r['クラブ名'] || r.clubName || '')}</td>
      <td><span class="badge badge-primary">${esc(r['区分'] || r.rateType || '')}</span></td>
      <td>${esc(r['指導者種別'] || r.type || '')}</td>
      <td class="text-right">¥${Number(r['時給'] || r.hourlyRate || 0).toLocaleString()}</td>
      <td>${esc(r['メールアドレス'] || r.email || '')}</td>
      <td class="col-actions">
        <button class="btn btn-sm btn-secondary" onclick="openMasterModal('${name}')">編集</button>
        <button class="btn btn-sm btn-danger"    onclick="openDeleteConfirm('${name}')">削除</button>
      </td>
    </tr>
  `;}).join('');
}

function filterMasterTable() {
  const q = document.getElementById('master-search').value.trim().toLowerCase();
  if (!q) { renderMasterTable(AdminState.instructors); return; }
  const filtered = AdminState.instructors.filter(r => {
    const name  = (r['氏名'] || r.name || '').toLowerCase();
    const club  = (r['クラブ名'] || r.clubName || '').toLowerCase();
    return name.includes(q) || club.includes(q);
  });
  renderMasterTable(filtered);
}

function populateInstructorSelects() {
  const ids = ['detail-instructor', 'slip-instructor'];
  ids.forEach(id => {
    const sel = document.getElementById(id);
    if (!sel) return;
    const current = sel.value;
    sel.innerHTML = '<option value="">選択してください</option>';
    AdminState.instructors.forEach(inst => {
      const opt = document.createElement('option');
      opt.value       = inst['氏名'] || inst.name || '';
      opt.textContent = (inst['氏名'] || inst.name || '') + '（' + (inst['クラブ名'] || inst.clubName || '') + '）';
      if (opt.value === current) opt.selected = true;
      sel.appendChild(opt);
    });
  });
}

function openMasterModal(name = null) {
  AdminState.editingName = name;
  const modal = document.getElementById('master-modal');
  document.getElementById('master-modal-title').textContent = name ? '指導者編集' : '指導者追加';

  clearMasterForm();
  if (name) {
    const inst = AdminState.instructors.find(i => (i['氏名'] || i.name) === name);
    if (inst) fillMasterForm(inst);
  }
  modal.classList.remove('hidden');
}

function closeMasterModal() {
  document.getElementById('master-modal').classList.add('hidden');
  AdminState.editingName = null;
}

function clearMasterForm() {
  ['m-name','m-club','m-email','m-bank','m-branch','m-account-number'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('m-rate-type').value   = 'メイン';
  document.getElementById('m-type').value        = '一般';
  document.getElementById('m-hourly').value      = '1600';
  document.getElementById('m-account-type').value= '普通';
}

function fillMasterForm(inst) {
  document.getElementById('m-name').value           = inst['氏名'] || inst.name || '';
  document.getElementById('m-club').value           = inst['クラブ名'] || inst.clubName || '';
  document.getElementById('m-rate-type').value      = inst['区分'] || inst.rateType || 'メイン';
  document.getElementById('m-type').value           = inst['指導者種別'] || inst.type || '一般';
  document.getElementById('m-hourly').value         = inst['時給'] || inst.hourlyRate || 1600;
  document.getElementById('m-email').value          = inst['メールアドレス'] || inst.email || '';
  document.getElementById('m-bank').value           = inst['金融機関'] || inst.bank || '';
  document.getElementById('m-branch').value         = inst['支店名'] || inst.branch || '';
  document.getElementById('m-account-type').value   = inst['口座種別'] || inst.accountType || '普通';
  document.getElementById('m-account-number').value = inst['口座番号'] || inst.accountNumber || '';
  document.getElementById('m-name').disabled        = true; // 編集時は氏名変更不可
}

async function saveMaster() {
  const name = AdminState.editingName || document.getElementById('m-name').value.trim();
  if (!name) { showToast('氏名を入力してください', 'error'); return; }

  const body = {
    action:         'updateMaster',
    operation:      AdminState.editingName ? 'update' : 'add',
    name,
    clubName:       document.getElementById('m-club').value.trim(),
    rateType:       document.getElementById('m-rate-type').value,
    instructorType: document.getElementById('m-type').value,
    hourlyRate:     parseInt(document.getElementById('m-hourly').value) || 0,
    email:          document.getElementById('m-email').value.trim(),
    bank:           document.getElementById('m-bank').value.trim(),
    branch:         document.getElementById('m-branch').value.trim(),
    accountType:    document.getElementById('m-account-type').value,
    accountNumber:  document.getElementById('m-account-number').value.trim(),
  };

  if (!body.clubName)  { showToast('クラブ名を入力してください', 'error'); return; }

  showLoading();
  try {
    const res = await gasPost(body);
    if (!res.success) throw new Error(res.error);
    showToast(AdminState.editingName ? '指導者情報を更新しました' : '指導者を追加しました', 'success');
    closeMasterModal();
    await loadMasterTable();
  } catch (e) {
    showToast('保存に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function openDeleteConfirm(name) {
  AdminState.deleteTarget = name;
  document.getElementById('confirm-delete-msg').textContent =
    `「${name}」を削除しますか？この操作は元に戻せません。`;
  document.getElementById('confirm-delete-modal').classList.remove('hidden');
}

async function confirmDeleteInstructor() {
  if (!AdminState.deleteTarget) return;
  document.getElementById('confirm-delete-modal').classList.add('hidden');
  showLoading();
  try {
    const res = await gasPost({ action: 'updateMaster', operation: 'delete', name: AdminState.deleteTarget });
    if (!res.success) throw new Error(res.error);
    showToast('削除しました', 'success');
    AdminState.deleteTarget = null;
    await loadMasterTable();
  } catch (e) {
    showToast('削除に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

// ========== GAS 通信 ==========

async function gasGet(params) {
  const url = new URL(Config.GAS_URL);
  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const res = await fetch(url.toString(), {
    method:   'GET',
    redirect: 'follow',
  });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  return res.json();
}

async function gasPost(body) {
  // Content-Type を text/plain にすることでプリフライトを回避
  const res = await fetch(Config.GAS_URL, {
    method:   'POST',
    redirect: 'follow',
    headers:  { 'Content-Type': 'text/plain;charset=UTF-8' },
    body:     JSON.stringify(body),
  });
  if (!res.ok) throw new Error('HTTP ' + res.status);
  return res.json();
}

// ========== ユーティリティ ==========

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function fmtDate(val) {
  if (!val) return '—';
  const d = new Date(val);
  if (isNaN(d)) return String(val);
  return d.getFullYear() + '/' +
    String(d.getMonth() + 1).padStart(2, '0') + '/' +
    String(d.getDate()).padStart(2, '0') + ' ' +
    String(d.getHours()).padStart(2, '0') + ':' +
    String(d.getMinutes()).padStart(2, '0');
}

function fmtDateJP(d) {
  return d.getFullYear() + '年' +
    (d.getMonth() + 1) + '月' +
    d.getDate() + '日';
}

let toastTimer = null;
function showToast(msg, type = 'info') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className   = 'toast show' + (type !== 'info' ? ' toast-' + type : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = 'toast'; }, 3400);
}

function showLoading() { document.getElementById('loading').classList.remove('hidden'); }
function hideLoading() { document.getElementById('loading').classList.add('hidden'); }
