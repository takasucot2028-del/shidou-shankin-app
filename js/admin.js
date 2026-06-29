/**
 * 事務局管理ロジック
 */

// ========== 状態 ==========
const AdminState = {
  instructors:        [],
  feeResults:         [],
  currentCalcId:      null,
  currentTravelTotal: 0,
  deleteTarget:       null,
  editingName:        null,       // マスタ編集中の氏名
  currentDetailRows:  [],         // loadDetail で取得した生データ
  editRowIndex:       0,
  editInstructorType: '一般',   // 月報修正時の指導者種別
};

// ========== 初期化 ==========

document.addEventListener('DOMContentLoaded', () => {
  initYearMonthSelects();
  initTabs();
  loadMasterTable();
  loadDashboard();

  document.getElementById('db-load-btn').addEventListener('click', loadDashboard);
  document.getElementById('db-bulk-report-print-btn').addEventListener('click', printAllReports);
  document.getElementById('detail-load-btn').addEventListener('click', loadDetail);
  document.getElementById('fee-load-btn').addEventListener('click', loadFeeResults);
  document.getElementById('fee-force-calc-btn').addEventListener('click', forceCalcFeeResults);
  document.getElementById('fee-export-btn').addEventListener('click', exportFeeSheet);
  document.getElementById('fee-transfer-btn').addEventListener('click', generateTransferSheet);
  document.getElementById('fee-club-summary-btn').addEventListener('click', generateClubSummarySheet);
  document.getElementById('slip-preview-btn').addEventListener('click', previewSlip);
  document.getElementById('slip-print-btn').addEventListener('click', () => window.print());
  document.getElementById('slip-bulk-print-btn').addEventListener('click', printAllSlips);
  document.getElementById('slip-email-btn').addEventListener('click', bulkSendPaySlipEmails);
  document.getElementById('save-fee-edit-btn').addEventListener('click', saveFeeEdit);
  document.getElementById('edit-fee').addEventListener('input', autoCalcNetPay);
  document.getElementById('edit-withholding').addEventListener('input', autoCalcNetPay);

  document.getElementById('master-add-btn').addEventListener('click', () => openMasterModal());
  document.getElementById('master-search-btn').addEventListener('click', filterMasterTable);
  document.getElementById('master-search').addEventListener('keydown', e => { if (e.key === 'Enter') filterMasterTable(); });
  document.getElementById('master-modal-close').addEventListener('click', closeMasterModal);
  document.getElementById('master-modal-cancel').addEventListener('click', closeMasterModal);
  document.getElementById('master-modal-save').addEventListener('click', saveMaster);

  // PINの表示切替（モーダル内）
  document.querySelectorAll('.pin-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      const inp = document.getElementById(btn.dataset.target);
      inp.type = inp.type === 'text' ? 'password' : 'text';
    });
  });
  document.getElementById('confirm-delete-cancel').addEventListener('click', () => {
    document.getElementById('confirm-delete-modal').classList.add('hidden');
  });
  document.getElementById('confirm-delete-ok').addEventListener('click', confirmDeleteInstructor);

  document.getElementById('report-edit-btn').addEventListener('click', openReportEdit);
  document.getElementById('report-edit-modal-close').addEventListener('click', closeReportEdit);
  document.getElementById('report-edit-cancel-btn').addEventListener('click', closeReportEdit);
  document.getElementById('report-edit-add-row-btn').addEventListener('click', () => addEditRow());
  document.getElementById('report-edit-save-btn').addEventListener('click', saveReportEdit);
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

// 個人表示と一括印刷で共有する行HTML生成ヘルパー
function buildDetailRowHTML(r) {
  return `<tr>
    <td>${esc(r.date || '')}</td>
    <td>${esc(Config.CATEGORY_LABEL[r.category] || r.category || '')}</td>
    <td>${esc(r.rateType || '')}</td>
    <td>${esc(r.startTime || '')}</td>
    <td>${esc(r.endTime || '')}</td>
    <td class="text-right">${r.instructionHours || 0}h</td>
    <td class="text-right fw-bold">${r.calcHours || 0}h</td>
    <td>${esc(r.transport || '')}</td>
    <td>${esc(r.destination || '')}</td>
    <td class="text-right">${r.travelAmount ? '¥' + Number(r.travelAmount).toLocaleString() : '—'}</td>
    <td>${esc(r.note || '')}</td>
  </tr>`;
}

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

    // 生データを保存（月報修正で使用）
    AdminState.currentDetailRows = res.data;

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
    document.getElementById('detail-tbody').innerHTML = res.data.map(buildDetailRowHTML).join('');

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
  // 前回の計算IDを必ずリセット（別指導者の古いIDが残ることを防ぐ）
  AdminState.currentCalcId = null;
  try {
    console.log('[loadDetailFee] calcFee呼び出し:', { name, year, month });
    const res = await gasPost({ action: 'calcFee', instructorName: name, year: parseInt(year), month: parseInt(month) });
    console.log('[loadDetailFee] calcFeeレスポンス:', res);
    if (!res.success) {
      console.warn('[loadDetailFee] calcFee失敗:', res.error);
      document.getElementById('detail-fee-result').innerHTML = '<p class="text-muted">謝金計算結果なし</p>';
      document.getElementById('edit-fee').value         = '';
      document.getElementById('edit-withholding').value = '';
      document.getElementById('edit-net').value         = '';
      return;
    }
    const r = res.result;
    AdminState.currentCalcId      = res.calcId || null;
    AdminState.currentTravelTotal = Number(r.travelTotal) || 0;
    console.log('[loadDetailFee] calcId取得:', AdminState.currentCalcId,
      '| fee:', r.fee, '| withholding:', r.withholding, '| netPay:', r.netPay);
    renderFeeResult(Number(r.fee), Number(r.withholding), Number(r.netPay), AdminState.currentTravelTotal);
    document.getElementById('edit-fee').value         = r.fee;
    document.getElementById('edit-withholding').value = r.withholding;
    document.getElementById('edit-net').value         = r.netPay;
  } catch (e) {
    console.error('[loadDetailFee] 例外:', e);
    AdminState.currentCalcId = null;
    document.getElementById('detail-fee-result').innerHTML = '<p class="text-muted">計算エラー: ' + esc(e.message) + '</p>';
  }
}

function renderFeeResult(fee, withholding, netPay, travelTotal, highlight = false) {
  console.log('[renderFeeResult] 呼び出し | fee:', fee, '| withholding:', withholding, '| netPay:', netPay, '| travelTotal:', travelTotal);
  const el = document.getElementById('detail-fee-result');
  if (!el) { console.error('[renderFeeResult] detail-fee-result 要素が見つかりません'); return; }
  el.innerHTML = `
    <div class="fee-row"><span>謝金総額</span><span class="fw-bold">¥${fee.toLocaleString()}</span></div>
    <div class="fee-row deduction"><span>源泉徴収額</span><span>−¥${withholding.toLocaleString()}</span></div>
    <div class="fee-row net"><span>差引支払額</span><span>¥${netPay.toLocaleString()}</span></div>
    <div class="fee-row"><span>旅費合計</span><span>¥${travelTotal.toLocaleString()}</span></div>
  `;
  if (highlight) {
    el.classList.remove('fee-result-updated');
    void el.offsetWidth;
    el.classList.add('fee-result-updated');
  }
  console.log('[renderFeeResult] DOM更新完了');
}

function autoCalcNetPay() {
  const fee = parseInt(document.getElementById('edit-fee').value, 10);
  const withholding = parseInt(document.getElementById('edit-withholding').value, 10);
  if (!isNaN(fee) && !isNaN(withholding) && fee >= 0 && withholding >= 0) {
    document.getElementById('edit-net').value = fee - withholding;
  }
}

async function saveFeeEdit() {
  console.log('[saveFeeEdit] 開始 | currentCalcId:', AdminState.currentCalcId);
  if (!AdminState.currentCalcId) {
    showToast('計算IDが取得できていません。月報詳細を再度「表示」してください', 'error');
    return;
  }
  const feeRaw         = document.getElementById('edit-fee').value;
  const withholdingRaw = document.getElementById('edit-withholding').value;
  const feeVal         = parseInt(feeRaw, 10);
  const withholdingVal = parseInt(withholdingRaw, 10);
  console.log('[saveFeeEdit] 入力値 | fee:', feeRaw, '→', feeVal, '| withholding:', withholdingRaw, '→', withholdingVal);
  if (isNaN(feeVal) || feeVal < 0) {
    showToast('謝金総額を正しく入力してください', 'error');
    return;
  }
  // 源泉徴収額は0円（免税・少額）も正常値なので >= 0 を許可
  if (isNaN(withholdingVal) || withholdingVal < 0) {
    showToast('源泉徴収額は0以上の整数を入力してください（0円の場合は「0」を入力）', 'error');
    return;
  }
  const netPay = feeVal - withholdingVal;
  document.getElementById('edit-net').value = netPay;
  const overrides = { fee: feeVal, withholding: withholdingVal, netPay };
  console.log('[saveFeeEdit] GAS送信:', { calcId: AdminState.currentCalcId, overrides });
  showLoading();
  try {
    const res = await gasPost({ action: 'updateFee', calcId: AdminState.currentCalcId, overrides });
    console.log('[saveFeeEdit] GASレスポンス:', res);
    if (!res.success) throw new Error(res.error);
    console.log('[saveFeeEdit] renderFeeResult呼び出し | fee:', feeVal, '| withholding:', withholdingVal, '| netPay:', netPay);
    renderFeeResult(feeVal, withholdingVal, netPay, AdminState.currentTravelTotal, true);
    showToast('謝金を修正しました', 'success');
  } catch (e) {
    console.error('[saveFeeEdit] 失敗:', e);
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

async function forceCalcFeeResults() {
  const year  = document.getElementById('fee-year').value;
  const month = document.getElementById('fee-month').value;
  if (!confirm('修正済みの謝金データを含め、全員分を強制再計算します。\n手動修正した金額は失われます。よろしいですか？')) return;
  showLoading();
  try {
    const res = await gasPost({ action: 'calcAllFees', year: parseInt(year), month: parseInt(month), forceRecalc: true });
    if (!res.success) throw new Error(res.error);
    AdminState.feeResults = res.data || [];
    if (AdminState.feeResults.length === 0) {
      showToast(res.message || '提出済みの月報がありません', 'error');
    } else {
      showToast('強制再計算が完了しました', 'success');
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

async function generateTransferSheet() {
  const year  = document.getElementById('fee-year').value;
  const month = document.getElementById('fee-month').value;
  console.log('[generateTransferSheet] fee-year element value:', year, '(type:', typeof year, ')');
  console.log('[generateTransferSheet] fee-month element value:', month, '(type:', typeof month, ')');
  console.log('[generateTransferSheet] GASに送るペイロード:', { action: 'generateTransferSheet', year, month });
  showLoading();
  try {
    const res = await gasPost({ action: 'generateTransferSheet', year, month });
    console.log('[generateTransferSheet] GASレスポンス:', res);
    if (res.success) {
      showToast(`口座振替データを出力しました（${res.count}件）`, 'success');
    } else {
      showToast(res.error || '出力に失敗しました', 'error');
    }
  } catch (e) {
    showToast('通信エラー: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

async function generateClubSummarySheet() {
  const year  = document.getElementById('fee-year').value;
  const month = document.getElementById('fee-month').value;
  showLoading();
  try {
    const res = await gasPost({ action: 'generateClubSummarySheet', year, month });
    if (res.success) {
      showToast(`クラブ別集計シートを出力しました（${res.count}件）`, 'success');
    } else {
      showToast(res.error || '出力に失敗しました', 'error');
    }
  } catch (e) {
    showToast('通信エラー: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
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

async function bulkSendPaySlipEmails() {
  const year  = document.getElementById('slip-year').value;
  const month = document.getElementById('slip-month').value;
  const ymLabel = year + '年' + month + '月';

  if (!confirm('対象年月（' + ymLabel + '）の給与明細を全指導者にメール送信します。よろしいですか？')) return;

  showLoading();
  try {
    const res = await gasPost({ action: 'sendPaySlipEmails', year: parseInt(year), month: parseInt(month) });
    if (!res.success) {
      showToast('メール送信に失敗しました: ' + (res.error || '不明なエラー'), 'error');
      return;
    }
    showToast(ymLabel + '分の給与明細メールを ' + res.sentCount + ' 件送信しました'
      + (res.skippedCount > 0 ? '（メールアドレス未登録 ' + res.skippedCount + ' 件スキップ）' : ''), 'success');
  } catch (e) {
    showToast('メール送信に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

// ========== 支払予定日計算 ==========

function nthMonday(year, month, n) {
  const d = new Date(year, month - 1, 1);
  let count = 0;
  while (true) {
    if (d.getDay() === 1) { count++; if (count === n) return d.getDate(); }
    d.setDate(d.getDate() + 1);
  }
}

function calcShunbun(year) {
  return Math.floor(20.8431 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
}

function calcShubun(year) {
  return Math.floor(23.2488 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
}

function getJPHolidays(year) {
  const set = new Set();
  const add = (m, d) => set.add(new Date(year, m - 1, d).toDateString());

  add(1, 1);
  add(2, 11);
  if (year >= 2020) add(2, 23);
  add(3, calcShunbun(year));
  add(4, 29);
  add(5, 3); add(5, 4); add(5, 5);
  add(8, 11);
  add(9, calcShubun(year));
  add(11, 3);
  add(11, 23);

  add(1,  nthMonday(year, 1,  2));
  add(7,  nthMonday(year, 7,  3));
  add(9,  nthMonday(year, 9,  3));
  add(10, nthMonday(year, 10, 2));

  // 振替休日：祝日が日曜なら翌月曜（既存祝日と重なれば更に翌日）
  const base = new Set(set);
  base.forEach(ds => {
    const d = new Date(ds);
    if (d.getDay() === 0) {
      const sub = new Date(d);
      sub.setDate(sub.getDate() + 1);
      while (set.has(sub.toDateString())) sub.setDate(sub.getDate() + 1);
      set.add(sub.toDateString());
    }
  });

  return set;
}

/**
 * 月報対象月に対する支払予定日を返す（翌月10日、土日祝の場合は前の平日）
 */
function calcPaymentDate(reportYear, reportMonth) {
  let y = reportYear, m = reportMonth + 1;
  if (m > 12) { m = 1; y++; }

  const holidays = getJPHolidays(y);
  const date = new Date(y, m - 1, 10);

  while (date.getDay() === 0 || date.getDay() === 6 || holidays.has(date.toDateString())) {
    date.setDate(date.getDate() - 1);
  }
  return date;
}

// ========== 給与明細レンダリング ==========

function buildSlipHTML(inst, fee, year, month) {
  const today    = new Date();
  const payDate  = calcPaymentDate(year, month);
  const instrName   = inst['氏名'] || inst.name || '';
  const clubName    = inst['クラブ名'] || inst.clubName || '';
  const rateType    = inst['区分'] || inst.rateType || '';
  const instrType   = inst['指導者種別'] || inst.type || '';
  const bank        = inst['金融機関'] || '';
  const branch      = inst['支店名'] || '';
  const accountType = inst['口座種別'] || '';
  const accountNum  = inst['口座番号'] || '';

  return `
    <div class="slip-org-name">一般社団法人たかすスポーツクラブ</div>
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
      <div class="slip-amount-row net"><span>差引支払額</span><span>¥${Number(fee['差引支払額'] || 0).toLocaleString()}</span></div>
      <div class="slip-amount-row travel-section"><span>旅費</span><span>¥${Number(fee['旅費総額'] || 0).toLocaleString()}</span></div>
    </div>

    <div class="slip-sign">
      <div class="slip-sign-box">
        <div class="slip-sign-label">指導者署名・押印</div>
      </div>
      <div class="slip-sign-box">
        <div class="slip-sign-label">確認者署名・押印</div>
      </div>
    </div>
  `;
}

function renderSlip(inst, fee, year, month) {
  document.getElementById('slip-content').innerHTML = buildSlipHTML(inst, fee, year, month);
}

async function printAllSlips() {
  const year  = parseInt(document.getElementById('slip-year').value);
  const month = parseInt(document.getElementById('slip-month').value);

  showLoading();
  try {
    const res = await gasPost({ action: 'calcAllFees', year, month });
    if (!res.success) throw new Error(res.error);

    const allFees = (res.data || []).filter(r => Number(r['謝金総額'] || 0) > 0);
    if (allFees.length === 0) {
      showToast('謝金が発生している指導者がいません', 'error');
      return;
    }

    const slipHTMLs = allFees.map(feeRow => {
      const instrName = feeRow['指導者氏名'];
      const inst = AdminState.instructors.find(i => (i['氏名'] || i.name) === instrName) || {};
      const feeData = {
        '謝金総額':             feeRow['謝金総額'],
        '源泉徴収額':           feeRow['源泉徴収額'],
        '差引支払額':           feeRow['差引支払額'],
        '旅費総額':             feeRow['旅費総額'],
        'メイン単価適用時間':   feeRow['メイン単価適用時間'],
        'サブ単価適用時間':     feeRow['サブ単価適用時間'],
        '平日謝金計算時間':     feeRow['平日謝金計算時間'],
        '休日謝金計算時間':     feeRow['休日謝金計算時間'],
        '長期休暇謝金計算時間': feeRow['長期休暇謝金計算時間'],
        '大会引率謝金計算時間': feeRow['大会引率謝金計算時間'],
      };
      return buildSlipHTML(inst, feeData, year, month);
    });

    const printWin = window.open('', '_blank');
    if (!printWin) {
      showToast('ポップアップがブロックされました。ブラウザの設定でポップアップを許可してください', 'error');
      return;
    }
    printWin.document.open();
    printWin.document.write(buildBulkPrintHTML(slipHTMLs, year, month));
    printWin.document.close();
  } catch (e) {
    showToast('一括印刷に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

// ========== 全指導者月報一括印刷 ==========

async function printAllReports() {
  const year  = parseInt(document.getElementById('db-year').value);
  const month = parseInt(document.getElementById('db-month').value);

  showLoading();
  try {
    const dashRes = await gasGet({ action: 'getDashboard', year, month });
    if (!dashRes.success) throw new Error(dashRes.error);

    const submittedList = (dashRes.statusList || []).filter(r => r.status === '提出済');
    if (submittedList.length === 0) {
      showToast('提出済みの月報がありません', 'error');
      return;
    }

    const [reportResults, feeRes] = await Promise.all([
      Promise.all(submittedList.map(r =>
        gasGet({ action: 'getReport', instructorName: r.instructorName, year, month })
      )),
      gasPost({ action: 'calcAllFees', year, month }),
    ]);

    const feeMap = {};
    if (feeRes.success) {
      (feeRes.data || []).forEach(f => { feeMap[f['指導者氏名']] = f; });
    }

    const reportHTMLs = reportResults
      .filter(res => res.success && res.data && res.data.length > 0)
      .map(res => {
        const rows  = res.data;
        const first = rows[0];
        const inst  = AdminState.instructors.find(i => (i['氏名'] || i.name) === first.instructorName) || {};
        const feeData = feeMap[first.instructorName] || null;
        return buildReportPrintHTML(inst, rows, year, month, feeData);
      });

    if (reportHTMLs.length === 0) {
      showToast('印刷できる月報がありません', 'error');
      return;
    }

    const printWin = window.open('', '_blank');
    if (!printWin) {
      showToast('ポップアップがブロックされました。ブラウザの設定でポップアップを許可してください', 'error');
      return;
    }
    printWin.document.open();
    printWin.document.write(buildBulkReportPrintHTML(reportHTMLs, year, month));
    printWin.document.close();
  } catch (e) {
    showToast('一括印刷に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function buildReportPrintHTML(inst, rows, year, month, feeData) {
  const first     = rows[0] || {};
  const instrName = inst['氏名']      || inst.name     || first.instructorName || '';
  const clubName  = inst['クラブ名']  || inst.clubName || first.clubName       || '';
  const status      = first.status      || '';
  const submittedAt = first.submittedAt ? fmtDate(first.submittedAt) : '—';
  const statusClass = status === '提出済' ? 'badge-success' : 'badge-warning';

  const feeSection = feeData ? `
    <div class="fee-section-title">謝金計算結果</div>
    <div class="fee-result">
      <div class="fee-row"><span>謝金総額</span><span class="fw-bold">¥${Number(feeData['謝金総額'] || 0).toLocaleString()}</span></div>
      <div class="fee-row deduction"><span>源泉徴収額</span><span>−¥${Number(feeData['源泉徴収額'] || 0).toLocaleString()}</span></div>
      <div class="fee-row net"><span>差引支払額</span><span>¥${Number(feeData['差引支払額'] || 0).toLocaleString()}</span></div>
      <div class="fee-row"><span>旅費合計</span><span>¥${Number(feeData['旅費総額'] || 0).toLocaleString()}</span></div>
    </div>
  ` : '';

  return `
    <div class="report-issuer">一般社団法人たかすスポーツクラブ</div>
    <h2 class="report-title">部活動地域展開 指導月報</h2>

    <div class="detail-meta">
      <div class="detail-meta-item">
        <span class="detail-meta-label">指導者</span>
        <span class="detail-meta-value">${esc(instrName)}</span>
      </div>
      <div class="detail-meta-item">
        <span class="detail-meta-label">クラブ名</span>
        <span class="detail-meta-value">${esc(clubName)}</span>
      </div>
      <div class="detail-meta-item">
        <span class="detail-meta-label">対象年月</span>
        <span class="detail-meta-value">${year}年${month}月</span>
      </div>
      <div class="detail-meta-item">
        <span class="detail-meta-label">ステータス</span>
        <span class="detail-meta-value"><span class="badge ${statusClass}">${esc(status)}</span></span>
      </div>
      <div class="detail-meta-item">
        <span class="detail-meta-label">提出日時</span>
        <span class="detail-meta-value">${submittedAt}</span>
      </div>
    </div>

    <div class="table-wrapper">
      <table class="data-table">
        <thead>
          <tr>
            <th>日付</th><th>内容</th><th>時給区分</th>
            <th>開始</th><th>終了</th><th>指導時間</th><th>計算時間</th>
            <th>交通手段</th><th>場所</th><th>旅費</th><th>備考</th>
          </tr>
        </thead>
        <tbody>${rows.map(buildDetailRowHTML).join('')}</tbody>
      </table>
    </div>

    ${feeSection}
  `;
}

function buildBulkReportPrintHTML(reportHTMLs, year, month) {
  const pages = reportHTMLs.map(html => `<div class="report-page">${html}</div>`).join('');
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>${year}年${month}月分 指導月報一括印刷（${reportHTMLs.length}名）</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif;
      font-size: 13px;
      line-height: 1.7;
      color: #2c3e50;
      background: #fff;
    }
    .report-page {
      padding: 12mm 14mm;
      width: 210mm;
      min-height: 297mm;
      page-break-after: always;
      break-after: page;
    }
    .report-page:last-child {
      page-break-after: auto;
      break-after: auto;
    }
    /* ヘッダー */
    .report-issuer {
      text-align: center;
      font-size: 11px;
      color: #6c757d;
      margin-bottom: 4px;
    }
    .report-title {
      font-size: 18px;
      font-weight: 700;
      text-align: center;
      margin-bottom: 12px;
      padding-bottom: 6px;
      border-bottom: 2px solid #2c3e50;
    }
    /* メタ情報（個人表示の detail-meta と同一） */
    .detail-meta {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
      gap: 12px;
      margin-bottom: 16px;
      background: #f4f6f9;
      border-radius: 8px;
      padding: 12px 16px;
    }
    .detail-meta-item { display: flex; flex-direction: column; gap: 2px; }
    .detail-meta-label { font-size: 12px; color: #6c757d; }
    .detail-meta-value { font-weight: 600; font-size: 13px; }
    /* バッジ */
    .badge {
      display: inline-flex;
      align-items: center;
      padding: 2px 10px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      white-space: nowrap;
    }
    .badge-success { background: #d4edda; color: #155724; }
    .badge-warning { background: #fff3cd; color: #856404; }
    /* テーブル（個人表示の data-table と同一） */
    .table-wrapper { overflow-x: auto; margin-bottom: 16px; }
    .data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    .data-table th {
      background: #e8f0fb;
      color: #1a56a0;
      font-weight: 700;
      padding: 10px 12px;
      text-align: left;
      white-space: nowrap;
      border-bottom: 2px solid #1a56a0;
    }
    .data-table td {
      padding: 10px 12px;
      border-bottom: 1px solid #f0f0f0;
      vertical-align: middle;
    }
    /* 謝金計算結果（個人表示の fee-row と同一） */
    .fee-section-title {
      font-size: 15px;
      font-weight: 700;
      color: #1a56a0;
      margin-bottom: 8px;
      padding-bottom: 4px;
      border-bottom: 2px solid #e8f0fb;
    }
    .fee-result { margin-bottom: 12px; }
    .fee-row {
      display: flex;
      justify-content: space-between;
      padding: 6px 0;
      border-bottom: 1px solid #f0f0f0;
      font-size: 13px;
    }
    .fee-row:last-child { border-bottom: none; }
    .fee-row.deduction { color: #dc3545; }
    .fee-row.net {
      font-size: 18px;
      font-weight: 700;
      color: #dc3545;
      border-top: 2px solid #dee2e6;
      margin-top: 4px;
      padding-top: 8px;
    }
    .text-right { text-align: right; }
    .fw-bold { font-weight: 700; }
    @page { size: A4 portrait; margin: 0; }
    @media screen {
      body { background: #e0e0e0; }
      .report-page {
        background: #fff;
        margin: 20px auto;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
      }
    }
  </style>
</head>
<body>
  ${pages}
  <script>window.print();<\/script>
</body>
</html>`;
}

function buildBulkPrintHTML(slipHTMLs, year, month) {
  const pages = slipHTMLs.map(html => `<div class="slip-page">${html}</div>`).join('');
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>${year}年${month}月分 給与明細一括印刷（${slipHTMLs.length}名）</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', 'Hiragino Sans', Meiryo, sans-serif;
      font-size: 13px;
      line-height: 1.7;
      color: #2c3e50;
      background: #fff;
    }
    .slip-page {
      padding: 15mm;
      width: 210mm;
      page-break-after: always;
      break-after: page;
    }
    .slip-page:last-child {
      page-break-after: auto;
      break-after: auto;
    }
    .slip-org-name {
      text-align: center;
      font-size: 14px;
      font-weight: 600;
      margin-bottom: 6px;
      color: #6c757d;
    }
    h2 {
      font-size: 20px;
      font-weight: 700;
      text-align: center;
      margin-bottom: 24px;
      padding-bottom: 8px;
      border-bottom: 2px solid #2c3e50;
    }
    dl.slip-meta {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 4px 24px;
      margin-bottom: 20px;
      font-size: 13px;
    }
    dl.slip-meta dt { color: #6c757d; }
    dl.slip-meta dd { font-weight: 600; }
    .slip-detail-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 20px;
      font-size: 12px;
    }
    .slip-detail-table th {
      background: #f0f0f0;
      border: 1px solid #ccc;
      padding: 6px 10px;
      text-align: center;
    }
    .slip-detail-table td {
      border: 1px solid #ccc;
      padding: 6px 10px;
      text-align: right;
    }
    .slip-detail-table td:first-child { text-align: left; }
    .slip-amount-box {
      border: 2px solid #dc3545;
      border-radius: 4px;
      padding: 16px;
      margin-bottom: 20px;
    }
    .slip-amount-row {
      display: flex;
      justify-content: space-between;
      padding: 4px 0;
      font-size: 13px;
    }
    .slip-amount-row.net {
      font-size: 18px;
      font-weight: 800;
      color: #dc3545;
      border-top: 2px solid #dc3545;
      margin-top: 8px;
      padding-top: 8px;
    }
    .slip-amount-row.travel-section {
      margin-top: 12px;
      border-top: 1px solid #dee2e6;
      padding-top: 8px;
    }
    .slip-sign {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 24px;
      margin-top: 24px;
    }
    .slip-sign-box {
      border: 1px solid #dee2e6;
      border-radius: 4px;
      padding: 12px;
      min-height: 60px;
    }
    .slip-sign-label { font-size: 11px; color: #6c757d; margin-bottom: 4px; }
    .text-center { text-align: center; }
    .text-muted { color: #888; }
    @page { size: A4 portrait; margin: 0; }
    @media screen {
      body { background: #e0e0e0; }
      .slip-page {
        background: #fff;
        margin: 20px auto;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
      }
    }
  </style>
</head>
<body>
  ${pages}
  <script>window.print();<\/script>
</body>
</html>`;
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
  ['m-name','m-club','m-email','m-bank','m-branch','m-account-number','m-pin'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('m-rate-type').value    = 'メイン';
  document.getElementById('m-type').value         = '一般';
  document.getElementById('m-hourly').value       = '1600';
  document.getElementById('m-account-type').value = '普通';
  // PIN入力欄をマスク表示（password）に戻す
  document.getElementById('m-pin').type = 'text';
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
  document.getElementById('m-pin').value            = '';   // セキュリティのため既存PINは表示しない
}

async function saveMaster() {
  const name = AdminState.editingName || document.getElementById('m-name').value.trim();
  if (!name) { showToast('氏名を入力してください', 'error'); return; }

  const pin = document.getElementById('m-pin').value.trim();
  if (pin && !/^\d{4}$/.test(pin)) {
    showToast('PINは4桁の数字で入力してください', 'error');
    return;
  }
  if (!AdminState.editingName && !pin) {
    showToast('新規追加時はPINを設定してください', 'error');
    return;
  }

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
    pin:            pin || undefined,
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

// ========== 月報修正 ==========

function openReportEdit() {
  const rows = AdminState.currentDetailRows;
  if (!rows || rows.length === 0) {
    showToast('月報データがありません。先に「表示」を押してください。', 'error');
    return;
  }

  const first = rows[0];
  const inst = AdminState.instructors.find(i => (i['氏名'] || i.name) === first.instructorName) || {};
  AdminState.editInstructorType = inst['指導者種別'] || inst.type || '一般';

  document.getElementById('report-edit-meta').innerHTML = `
    <div class="detail-meta-item"><span class="detail-meta-label">指導者</span><span class="detail-meta-value">${esc(first.instructorName)}</span></div>
    <div class="detail-meta-item"><span class="detail-meta-label">クラブ名</span><span class="detail-meta-value">${esc(first.clubName)}</span></div>
    <div class="detail-meta-item"><span class="detail-meta-label">対象年月</span><span class="detail-meta-value">${first.year}年${first.month}月</span></div>
    <div class="detail-meta-item"><span class="detail-meta-label">指導者種別</span><span class="detail-meta-value">${esc(AdminState.editInstructorType)}</span></div>
  `;

  document.getElementById('report-edit-tbody').innerHTML = '';
  AdminState.editRowIndex = 0;
  rows.forEach(r => addEditRow(r));

  document.getElementById('report-edit-modal').classList.remove('hidden');
}

function closeReportEdit() {
  document.getElementById('report-edit-modal').classList.add('hidden');
}

function addEditRow(data = null) {
  const tbody = document.getElementById('report-edit-tbody');
  if (tbody.rows.length >= 31) {
    showToast('1ヶ月の最大行数（31行）に達しました', 'warning');
    return;
  }

  const id = 'edit-row-' + (AdminState.editRowIndex++);
  const tr = document.createElement('tr');
  tr.id = id;
  tr.innerHTML = `
    <td><input type="date" class="edit-inp-date" aria-label="日付"></td>
    <td>
      <select class="edit-sel-category" aria-label="内容">
        <option value="平日">平日練習</option>
        <option value="休日">休日練習</option>
        <option value="長期休暇">長期休暇</option>
        <option value="大会引率">大会引率</option>
      </select>
    </td>
    <td>
      <select class="edit-sel-rate" aria-label="時給区分">
        <option value="メイン">メイン</option>
        <option value="サブ">サブ</option>
      </select>
    </td>
    <td><input type="time" class="edit-inp-start" aria-label="開始時刻"></td>
    <td><input type="time" class="edit-inp-end" aria-label="終了時刻"></td>
    <td><input type="text" class="edit-inp-hours" readonly placeholder="自動" aria-label="指導時間"></td>
    <td>
      <select class="edit-sel-transport" aria-label="交通手段">
        <option value="">なし</option>
        <option value="バス代">バス代</option>
        <option value="JR代">JR代</option>
      </select>
    </td>
    <td><input type="text" class="edit-inp-dest" placeholder="場所" aria-label="場所"></td>
    <td><input type="number" class="edit-inp-travel" min="0" step="1" placeholder="0" aria-label="旅費金額"></td>
    <td><input type="text" class="edit-inp-note" placeholder="任意" aria-label="備考"></td>
    <td><button type="button" class="btn-icon edit-delete-row-btn" title="削除" aria-label="削除">✕</button></td>
  `;

  if (data) {
    tr.querySelector('.edit-inp-date').value     = parseAdminDateStr(data.date);
    tr.querySelector('.edit-sel-category').value  = data.category || '平日';
    tr.querySelector('.edit-sel-rate').value      = data.rateType || 'メイン';
    tr.querySelector('.edit-inp-start').value     = parseAdminTimeStr(data.startTime);
    tr.querySelector('.edit-inp-end').value       = parseAdminTimeStr(data.endTime);
    tr.querySelector('.edit-sel-transport').value = data.transport || '';
    tr.querySelector('.edit-inp-dest').value      = data.destination || '';
    tr.querySelector('.edit-inp-travel').value    = data.travelAmount || '';
    tr.querySelector('.edit-inp-note').value      = data.note || '';
    recalcEditRow(tr);
  }

  tr.querySelector('.edit-inp-start').addEventListener('change', () => recalcEditRow(tr));
  tr.querySelector('.edit-inp-end').addEventListener('change',   () => recalcEditRow(tr));
  tr.querySelector('.edit-sel-category').addEventListener('change', () => recalcEditRow(tr));
  tr.querySelector('.edit-delete-row-btn').addEventListener('click', () => {
    if (document.getElementById('report-edit-tbody').rows.length > 1) {
      tr.remove();
    } else {
      showToast('最低1行は必要です', 'warning');
    }
  });

  tbody.appendChild(tr);
}

function recalcEditRow(tr) {
  const start    = tr.querySelector('.edit-inp-start').value;
  const end      = tr.querySelector('.edit-inp-end').value;
  const category = tr.querySelector('.edit-sel-category').value;
  const hours    = calcAdminInstructionHours(start, end, category, AdminState.editInstructorType);
  const inp      = tr.querySelector('.edit-inp-hours');
  inp.value          = hours > 0 ? formatAdminHours(hours) : '';
  inp.dataset.hours  = hours;
}

function calcAdminInstructionHours(startTime, endTime, category, instructorType) {
  const startMin = adminTimeToMinutes(startTime);
  const endMin   = adminTimeToMinutes(endTime);
  if (startMin === null || endMin === null || endMin <= startMin) return 0;

  let effectiveMin;
  const isTeacher     = instructorType === '教員';
  const isWeekdayLike = category === '平日' || category === '長期休暇';

  if (isTeacher && isWeekdayLike) {
    const effectiveStart = Math.max(startMin, Config.TEACHER_WORK_END_MINUTES);
    effectiveMin = Math.max(endMin - effectiveStart, 0);
  } else {
    effectiveMin = endMin - startMin;
  }

  const maxMin     = (Config.MAX_HOURS[category] || 0) * 60;
  const cappedMin  = Math.min(effectiveMin, maxMin);
  const flooredMin = Math.floor(cappedMin / 15) * 15;
  return flooredMin / 60;
}

function adminTimeToMinutes(t) {
  if (!t) return null;
  const [h, m] = t.split(':').map(Number);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function formatAdminHours(h) {
  const hh = Math.floor(h);
  const mm = Math.round((h - hh) * 60);
  return hh + '時間' + String(mm).padStart(2, '0') + '分';
}

function parseAdminDateStr(val) {
  if (!val) return '';
  const s = String(val);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (s.includes('T')) return s.slice(0, 10);
  try {
    const d = new Date(s);
    if (!isNaN(d)) {
      return d.getFullYear() + '-' +
        String(d.getMonth() + 1).padStart(2, '0') + '-' +
        String(d.getDate()).padStart(2, '0');
    }
  } catch (_) {}
  return s;
}

function parseAdminTimeStr(val) {
  if (!val) return '';
  const s = String(val);
  if (/^\d{2}:\d{2}$/.test(s)) return s;
  try {
    const d = new Date(s);
    if (!isNaN(d)) {
      return String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
    }
  } catch (_) {}
  return s;
}

async function saveReportEdit() {
  if (!confirm('月報データを上書きします。よろしいですか？')) return;

  const rows = AdminState.currentDetailRows;
  if (!rows || rows.length === 0) return;
  const first = rows[0];

  const editRows = [];
  document.querySelectorAll('#report-edit-tbody tr').forEach(tr => {
    const date  = tr.querySelector('.edit-inp-date')?.value;
    const start = tr.querySelector('.edit-inp-start')?.value;
    const end   = tr.querySelector('.edit-inp-end')?.value;
    if (!date && !start && !end) return;

    const category = tr.querySelector('.edit-sel-category')?.value || '平日';
    const calcH    = calcAdminInstructionHours(start, end, category, AdminState.editInstructorType);

    editRows.push({
      date,
      category,
      rateType:         tr.querySelector('.edit-sel-rate')?.value || 'メイン',
      startTime:        start,
      endTime:          end,
      instructionHours: parseFloat(tr.querySelector('.edit-inp-hours')?.dataset.hours || 0) || 0,
      calcHours:        calcH,
      transport:        tr.querySelector('.edit-sel-transport')?.value || '',
      destination:      tr.querySelector('.edit-inp-dest')?.value || '',
      travelAmount:     parseFloat(tr.querySelector('.edit-inp-travel')?.value || 0) || 0,
      note:             tr.querySelector('.edit-inp-note')?.value || '',
    });
  });

  if (editRows.length === 0) {
    showToast('1件以上の記録を入力してください', 'error');
    return;
  }

  showLoading();
  try {
    const res = await gasPost({
      action:         'updateReport',
      submitId:       first.submitId,
      instructorName: first.instructorName,
      clubName:       first.clubName,
      year:           parseInt(first.year),
      month:          parseInt(first.month),
      submittedAt:    first.submittedAt,
      rows:           editRows,
    });
    if (!res.success) throw new Error(res.error);
    showToast('月報を修正し、謝金を再計算しました', 'success');
    closeReportEdit();
    loadDetail();
  } catch (e) {
    showToast('修正に失敗しました: ' + e.message, 'error');
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
