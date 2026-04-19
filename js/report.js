/**
 * 月報入力ロジック
 */

// ========== 状態 ==========
const State = {
  instructors: [],       // 指導者マスタ一覧
  currentInstructor: null,
  submitId: null,        // 下書きID（初回保存時に取得）
  rows: [],              // 入力行データのキャッシュ（提出確認用）
};

// ========== 初期化 ==========

document.addEventListener('DOMContentLoaded', () => {
  initYearMonth();
  loadInstructors();
  addRow();              // 最初の1行を表示

  document.getElementById('instructor-name').addEventListener('change', onInstructorChange);
  document.getElementById('add-row-btn').addEventListener('click', addRow);
  document.getElementById('save-btn').addEventListener('click', onSave);
  document.getElementById('submit-btn').addEventListener('click', onShowConfirm);
  document.getElementById('back-btn').addEventListener('click', showFormSection);
  document.getElementById('confirm-submit-btn').addEventListener('click', onConfirmSubmit);
  document.getElementById('new-report-btn').addEventListener('click', resetForm);
});

function initYearMonth() {
  const now = new Date();
  const yearSel  = document.getElementById('target-year');
  const monthSel = document.getElementById('target-month');

  for (let y = now.getFullYear() - 1; y <= now.getFullYear() + 1; y++) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y;
    if (y === now.getFullYear()) opt.selected = true;
    yearSel.appendChild(opt);
  }
  for (let m = 1; m <= 12; m++) {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = m;
    if (m === now.getMonth() + 1) opt.selected = true;
    monthSel.appendChild(opt);
  }
}

async function loadInstructors() {
  try {
    const data = await gasGet({ action: 'getInstructors' });
    if (!data.success) throw new Error(data.error);
    State.instructors = data.data || [];
    populateInstructorSelect();
  } catch (e) {
    showToast('指導者一覧の取得に失敗しました: ' + e.message, 'error');
  }
}

function populateInstructorSelect() {
  const sel = document.getElementById('instructor-name');
  sel.innerHTML = '<option value="">選択してください</option>';
  State.instructors.forEach(inst => {
    const opt = document.createElement('option');
    opt.value = inst['氏名'] || inst.name || '';
    opt.textContent = (inst['氏名'] || inst.name || '') + '（' + (inst['クラブ名'] || inst.clubName || '') + '）';
    sel.appendChild(opt);
  });
}

function onInstructorChange() {
  const name = document.getElementById('instructor-name').value;
  const inst = State.instructors.find(i => (i['氏名'] || i.name) === name);
  State.currentInstructor = inst || null;
  document.getElementById('club-name').value = inst ? (inst['クラブ名'] || inst.clubName || '') : '';

  // 全行の時給区分デフォルトを更新
  if (inst) {
    const defaultRate = inst['区分'] || inst.rateType || 'メイン';
    document.querySelectorAll('.sel-rate').forEach(sel => { sel.value = defaultRate; });
    updateTotals();
  }
}

// ========== 行操作 ==========

let rowIndex = 0;

function addRow() {
  const tbody = document.getElementById('report-tbody');
  if (tbody.rows.length >= 31) {
    showToast('1ヶ月の最大行数（31行）に達しました', 'warning');
    return;
  }

  const id = 'row-' + (rowIndex++);
  const defaultRate = State.currentInstructor
    ? (State.currentInstructor['区分'] || State.currentInstructor.rateType || 'メイン')
    : 'メイン';

  const tr = document.createElement('tr');
  tr.id = id;
  tr.innerHTML = `
    <td data-label="日付">
      <input type="date" class="inp-date" aria-label="日付">
    </td>
    <td data-label="区分">
      <select class="sel-category" aria-label="区分">
        <option value="平日">平日</option>
        <option value="休日">休日</option>
        <option value="長期休暇">長期休暇</option>
        <option value="大会引率">大会引率</option>
      </select>
    </td>
    <td data-label="時給区分">
      <select class="sel-rate" aria-label="時給区分">
        <option value="メイン">メイン</option>
        <option value="サブ">サブ</option>
      </select>
    </td>
    <td data-label="開始時刻">
      <input type="time" class="inp-start" aria-label="開始時刻">
    </td>
    <td data-label="終了時刻">
      <input type="time" class="inp-end" aria-label="終了時刻">
    </td>
    <td data-label="指導時間" class="hours-cell">
      <input type="text" class="inp-hours" readonly placeholder="自動">
    </td>
    <td data-label="交通手段">
      <select class="sel-transport" aria-label="交通手段">
        <option value="">なし</option>
        <option value="バス代">バス代</option>
        <option value="JR代">JR代</option>
      </select>
    </td>
    <td data-label="行先">
      <input type="text" class="inp-dest" placeholder="行先" aria-label="行先">
    </td>
    <td data-label="旅費(円)">
      <input type="number" class="inp-travel" min="0" step="1" placeholder="0" aria-label="旅費金額">
    </td>
    <td data-label="備考">
      <input type="text" class="inp-note" placeholder="任意" aria-label="備考">
    </td>
    <td class="col-action">
      <button type="button" class="btn-icon delete-row-btn" title="この行を削除" aria-label="削除">✕</button>
    </td>
  `;

  tr.querySelector('.sel-rate').value = defaultRate;

  // イベント登録
  tr.querySelector('.inp-start').addEventListener('change', () => recalcRow(tr));
  tr.querySelector('.inp-end').addEventListener('change',   () => recalcRow(tr));
  tr.querySelector('.sel-category').addEventListener('change', () => recalcRow(tr));
  tr.querySelector('.sel-rate').addEventListener('change', updateTotals);
  tr.querySelector('.inp-travel').addEventListener('input', updateTotals);
  tr.querySelector('.delete-row-btn').addEventListener('click', () => {
    if (document.getElementById('report-tbody').rows.length > 1) {
      tr.remove();
      updateTotals();
    } else {
      showToast('最低1行は必要です', 'warning');
    }
  });

  tbody.appendChild(tr);
}

function recalcRow(tr) {
  const start    = tr.querySelector('.inp-start').value;
  const end      = tr.querySelector('.inp-end').value;
  const category = tr.querySelector('.sel-category').value;
  const instrType = State.currentInstructor
    ? (State.currentInstructor['指導者種別'] || State.currentInstructor.type || '一般')
    : '一般';

  const hours = calcInstructionHours(start, end, category, instrType);
  const hoursInput = tr.querySelector('.inp-hours');
  hoursInput.value = hours > 0 ? formatHours(hours) : '';
  hoursInput.dataset.hours = hours;

  updateTotals();
}

// ========== 謝金計算ロジック（フロント側・概算） ==========

function calcInstructionHours(startTime, endTime, category, instructorType) {
  const startMin = timeToMinutes(startTime);
  const endMin   = timeToMinutes(endTime);
  if (startMin === null || endMin === null || endMin <= startMin) return 0;

  let effectiveMin;
  const isTeacher = instructorType === '教員';
  const isWeekdayLike = category === '平日' || category === '長期休暇';

  if (isTeacher && isWeekdayLike) {
    const effectiveStart = Math.max(startMin, Config.TEACHER_WORK_END_MINUTES);
    effectiveMin = Math.max(endMin - effectiveStart, 0);
  } else {
    effectiveMin = endMin - startMin;
  }

  const maxMin     = (Config.MAX_HOURS[category] || 0) * 60;
  const cappedMin  = Math.min(effectiveMin, maxMin);
  const flooredMin = Math.floor(cappedMin / 15) * 15; // 15分単位切捨

  return flooredMin / 60;
}

function calcFeePreview(rows, instructorType) {
  let mainHours = 0;
  let subHours  = 0;
  let travel    = 0;

  rows.forEach(r => {
    const h = parseFloat(r.calcHours) || 0;
    if (r.rateType === 'メイン') mainHours += h;
    else                          subHours  += h;
    travel += parseFloat(r.travelAmount) || 0;
  });

  const fee         = Math.round(mainHours * Config.HOURLY_RATE.MAIN + subHours * Config.HOURLY_RATE.SUB);
  const withholding = Math.floor(fee * Config.WITHHOLDING_TAX_RATE);
  const netPay      = fee - withholding;

  return { mainHours, subHours, fee, withholding, netPay, travel };
}

// ========== 合計更新 ==========

function updateTotals() {
  let totalMin   = 0;
  let totalTravel = 0;
  let mainHours  = 0;
  let subHours   = 0;

  document.querySelectorAll('#report-tbody tr').forEach(tr => {
    const h    = parseFloat(tr.querySelector('.inp-hours')?.dataset.hours || 0) || 0;
    const rate = tr.querySelector('.sel-rate')?.value;
    const tv   = parseFloat(tr.querySelector('.inp-travel')?.value || 0) || 0;

    totalMin += h * 60;
    if (rate === 'メイン') mainHours += h; else subHours += h;
    totalTravel += tv;
  });

  const h = Math.floor(totalMin / 60);
  const m = Math.round(totalMin % 60);
  document.getElementById('total-hours').textContent = h + '時間' + String(m).padStart(2, '0') + '分';
  document.getElementById('total-travel').textContent = '¥' + totalTravel.toLocaleString();

  const fee = Math.round(mainHours * Config.HOURLY_RATE.MAIN + subHours * Config.HOURLY_RATE.SUB);
  document.getElementById('total-fee').textContent = '¥' + fee.toLocaleString();
}

// ========== フォームデータ収集 ==========

function collectRows() {
  const rows = [];
  document.querySelectorAll('#report-tbody tr').forEach(tr => {
    const date  = tr.querySelector('.inp-date')?.value;
    const start = tr.querySelector('.inp-start')?.value;
    const end   = tr.querySelector('.inp-end')?.value;
    if (!date && !start && !end) return; // 空行スキップ

    const instrType = State.currentInstructor
      ? (State.currentInstructor['指導者種別'] || State.currentInstructor.type || '一般')
      : '一般';
    const category = tr.querySelector('.sel-category')?.value || '平日';
    const calcH = calcInstructionHours(start, end, category, instrType);

    rows.push({
      date,
      category,
      rateType:      tr.querySelector('.sel-rate')?.value || 'メイン',
      startTime:     start,
      endTime:       end,
      instructionHours: parseFloat(tr.querySelector('.inp-hours')?.dataset.hours || 0) || 0,
      calcHours:     calcH,
      transport:     tr.querySelector('.sel-transport')?.value || '',
      destination:   tr.querySelector('.inp-dest')?.value || '',
      travelAmount:  parseFloat(tr.querySelector('.inp-travel')?.value || 0) || 0,
      note:          tr.querySelector('.inp-note')?.value || '',
    });
  });
  return rows;
}

// ========== バリデーション ==========

function validate() {
  const name  = document.getElementById('instructor-name').value;
  const year  = document.getElementById('target-year').value;
  const month = document.getElementById('target-month').value;

  if (!name)  { showToast('指導者氏名を選択してください', 'error'); return false; }
  if (!year || !month) { showToast('対象年月を選択してください', 'error'); return false; }

  const rows = collectRows();
  if (rows.length === 0) { showToast('指導記録を1件以上入力してください', 'error'); return false; }

  for (const r of rows) {
    if (!r.date)      { showToast('日付が未入力の行があります', 'error'); return false; }
    if (!r.startTime) { showToast('開始時刻が未入力の行があります', 'error'); return false; }
    if (!r.endTime)   { showToast('終了時刻が未入力の行があります', 'error'); return false; }
    if (r.endTime <= r.startTime) {
      showToast('終了時刻は開始時刻より後にしてください', 'error');
      return false;
    }
  }
  return true;
}

// ========== 一時保存 ==========

async function onSave() {
  if (!validate()) return;
  showLoading();
  try {
    const body = buildPostBody();
    const res  = await gasPost(body);
    if (!res.success) throw new Error(res.error);
    State.submitId = res.submitId;
    showToast('一時保存しました', 'success');
  } catch (e) {
    showToast('保存に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function buildPostBody(action = 'saveReport') {
  const rows = collectRows();
  State.rows = rows;
  return {
    action,
    submitId:       State.submitId || undefined,
    instructorName: document.getElementById('instructor-name').value,
    clubName:       document.getElementById('club-name').value,
    year:           parseInt(document.getElementById('target-year').value),
    month:          parseInt(document.getElementById('target-month').value),
    rows,
  };
}

// ========== 確認画面表示 ==========

function onShowConfirm() {
  if (!validate()) return;

  const name  = document.getElementById('instructor-name').value;
  const club  = document.getElementById('club-name').value;
  const year  = document.getElementById('target-year').value;
  const month = document.getElementById('target-month').value;
  const rows  = collectRows();
  State.rows  = rows;

  // 基本情報サマリー
  const summary = document.getElementById('confirm-summary');
  summary.innerHTML = `
    <div class="confirm-info-item"><span class="confirm-info-label">指導者氏名</span><span class="confirm-info-value">${esc(name)}</span></div>
    <div class="confirm-info-item"><span class="confirm-info-label">クラブ名</span><span class="confirm-info-value">${esc(club)}</span></div>
    <div class="confirm-info-item"><span class="confirm-info-label">対象年月</span><span class="confirm-info-value">${year}年${month}月</span></div>
    <div class="confirm-info-item"><span class="confirm-info-label">入力行数</span><span class="confirm-info-value">${rows.length}行</span></div>
  `;

  // 謝金プレビュー
  const instrType = State.currentInstructor
    ? (State.currentInstructor['指導者種別'] || State.currentInstructor.type || '一般')
    : '一般';
  const preview = calcFeePreview(rows, instrType);
  const feeDiv  = document.getElementById('fee-preview-content');
  feeDiv.innerHTML = `
    <div class="fee-row"><span>メイン指導時間</span><span>${preview.mainHours.toFixed(2)}時間 × ¥${Config.HOURLY_RATE.MAIN.toLocaleString()}</span></div>
    <div class="fee-row"><span>サブ指導時間</span><span>${preview.subHours.toFixed(2)}時間 × ¥${Config.HOURLY_RATE.SUB.toLocaleString()}</span></div>
    <div class="fee-row total"><span>謝金総額</span><span>¥${preview.fee.toLocaleString()}</span></div>
    <div class="fee-row deduction"><span>源泉徴収額（10.21%）</span><span>−¥${preview.withholding.toLocaleString()}</span></div>
    <div class="fee-row net"><span>差引支払額</span><span>¥${preview.netPay.toLocaleString()}</span></div>
    <div class="fee-row"><span>旅費合計</span><span>¥${preview.travel.toLocaleString()}</span></div>
  `;

  document.getElementById('form-section').classList.add('hidden');
  document.getElementById('confirm-section').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function showFormSection() {
  document.getElementById('confirm-section').classList.add('hidden');
  document.getElementById('form-section').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ========== 提出 ==========

async function onConfirmSubmit() {
  showLoading();
  try {
    const body = buildPostBody('submitReport');
    const res  = await gasPost(body);
    if (!res.success) throw new Error(res.error);
    State.submitId = res.submitId;
    showCompleteSection();
  } catch (e) {
    showToast('提出に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function showCompleteSection() {
  const name  = document.getElementById('instructor-name').value;
  const year  = document.getElementById('target-year').value;
  const month = document.getElementById('target-month').value;

  document.getElementById('complete-message').textContent =
    `${name} さんの ${year}年${month}月分 指導月報が提出されました。`;

  document.getElementById('confirm-section').classList.add('hidden');
  document.getElementById('complete-section').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function resetForm() {
  document.getElementById('instructor-name').value = '';
  document.getElementById('club-name').value = '';
  document.getElementById('report-tbody').innerHTML = '';
  State.currentInstructor = null;
  State.submitId = null;
  State.rows = [];
  rowIndex = 0;
  addRow();
  updateTotals();
  document.getElementById('complete-section').classList.add('hidden');
  document.getElementById('form-section').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
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

function timeToMinutes(t) {
  if (!t) return null;
  const [h, m] = t.split(':').map(Number);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

function formatHours(h) {
  const hh = Math.floor(h);
  const mm = Math.round((h - hh) * 60);
  return hh + '時間' + String(mm).padStart(2, '0') + '分';
}

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

let toastTimer = null;
function showToast(msg, type = 'info') {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className   = 'toast show' + (type !== 'info' ? ' toast-' + type : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.className = 'toast'; }, 3200);
}

function showLoading() { document.getElementById('loading').classList.remove('hidden'); }
function hideLoading() { document.getElementById('loading').classList.add('hidden'); }
