/**
 * 月報入力ロジック
 */

// ========== 状態 ==========
const State = {
  instructors: [],       // 指導者マスタ一覧
  currentInstructor: null,
  submitId: null,        // 下書きID（初回保存時に取得）
  rows: [],              // 入力行データのキャッシュ（提出確認用）
  isSubmitted: false,    // 提出済みデータ表示中フラグ
};

// ========== 認証（sessionStorage） ==========

const AUTH_KEY = 'report_auth_v1';

function getAuth() {
  try { return JSON.parse(sessionStorage.getItem(AUTH_KEY)); } catch { return null; }
}
function setAuth(name) { sessionStorage.setItem(AUTH_KEY, JSON.stringify({ name })); }
function clearAuth() { sessionStorage.removeItem(AUTH_KEY); }

// ========== 初期化 ==========

document.addEventListener('DOMContentLoaded', () => {
  initYearMonth();
  initPayslipSelectors();
  loadInstructors();     // ログイン画面のドロップダウンを先に埋める

  // ログインイベント
  document.getElementById('login-btn').addEventListener('click', onLogin);
  document.getElementById('login-pin').addEventListener('keydown', e => { if (e.key === 'Enter') onLogin(); });
  document.getElementById('logout-btn').addEventListener('click', onLogout);

  // PINの表示切替
  document.querySelectorAll('.pin-toggle').forEach(btn => {
    btn.addEventListener('click', () => {
      const inp = document.getElementById(btn.dataset.target);
      inp.type = inp.type === 'text' ? 'password' : 'text';
    });
  });

  document.getElementById('instructor-name').addEventListener('change', onInstructorChange);
  document.getElementById('target-year').addEventListener('change', onYearMonthChange);
  document.getElementById('target-month').addEventListener('change', onYearMonthChange);
  document.getElementById('add-row-btn').addEventListener('click', addRow);
  document.getElementById('save-btn').addEventListener('click', onSave);
  document.getElementById('submit-btn').addEventListener('click', onShowConfirm);
  document.getElementById('back-btn').addEventListener('click', showFormSection);
  document.getElementById('confirm-submit-btn').addEventListener('click', onConfirmSubmit);
  document.getElementById('new-report-btn').addEventListener('click', resetForm);
  document.getElementById('resubmit-btn').addEventListener('click', onResubmit);

  // タブ切替
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  // 給与明細ボタン
  document.getElementById('payslip-show-btn').addEventListener('click', onShowPaySlip);
  document.getElementById('payslip-pdf-btn').addEventListener('click', printPaySlip);

  // 既存セッションがあればそのままフォームへ
  const auth = getAuth();
  if (auth && auth.name) {
    showFormAfterLogin(auth.name);
  } else {
    showLoginSection();
  }
});

// ========== ログイン/ログアウト ==========

function showLoginSection() {
  document.getElementById('login-section').classList.remove('hidden');
  document.getElementById('tab-nav').classList.add('hidden');
  document.getElementById('form-section').classList.add('hidden');
  document.getElementById('payslip-section').classList.add('hidden');
  document.getElementById('confirm-section').classList.add('hidden');
  document.getElementById('complete-section').classList.add('hidden');
  document.getElementById('logout-btn').classList.add('hidden');
}

async function onLogin() {
  const name = document.getElementById('login-instructor').value;
  const pin  = document.getElementById('login-pin').value;
  const errEl = document.getElementById('login-error');

  errEl.classList.add('hidden');
  errEl.textContent = '';

  if (!name) { showLoginError('指導者氏名を選択してください'); return; }
  if (!pin)  { showLoginError('PINを入力してください'); return; }
  if (!/^\d{4}$/.test(pin)) { showLoginError('PINは4桁の数字で入力してください'); return; }

  document.getElementById('login-btn').disabled = true;
  showLoading();
  try {
    const res = await gasPost({ action: 'checkPin', name, pin });
    if (!res.success) {
      showLoginError(res.error || 'PINが正しくありません');
      return;
    }
    setAuth(name);
    showFormAfterLogin(name);
  } catch (e) {
    showLoginError('通信エラーが発生しました: ' + e.message);
  } finally {
    document.getElementById('login-btn').disabled = false;
    hideLoading();
  }
}

function showLoginError(msg) {
  const errEl = document.getElementById('login-error');
  errEl.textContent = msg;
  errEl.classList.remove('hidden');
}

function showFormAfterLogin(name) {
  State._pendingLoginName = name;

  document.getElementById('login-section').classList.add('hidden');
  document.getElementById('tab-nav').classList.remove('hidden');
  document.getElementById('form-section').classList.remove('hidden');
  document.getElementById('payslip-section').classList.add('hidden');
  document.getElementById('confirm-section').classList.add('hidden');
  document.getElementById('complete-section').classList.add('hidden');
  document.getElementById('logout-btn').classList.remove('hidden');
  document.getElementById('login-pin').value = '';

  // 月報タブをアクティブにする
  document.querySelectorAll('.tab-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === 'report');
  });

  // 指導者一覧がすでにあれば即反映
  if (State.instructors.length > 0) {
    applyLoginToForm(name);
  }
}

function applyLoginToForm(name) {
  const sel = document.getElementById('instructor-name');
  // ログイン者のみの選択肢に絞り込む
  sel.innerHTML = '';
  const inst = State.instructors.find(i => (i['氏名'] || i.name) === name);
  if (inst) {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name + '（' + (inst['クラブ名'] || inst.clubName || '') + '）';
    sel.appendChild(opt);
  } else {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  }
  sel.value = name;
  sel.disabled = true;

  // クラブ名・指導者情報を自動セット
  State.currentInstructor = inst || null;
  document.getElementById('club-name').value = inst ? (inst['クラブ名'] || inst.clubName || '') : '';

  if (document.getElementById('report-tbody').rows.length === 0) {
    addRow();
  }
  if (inst) {
    const defaultRate = inst['区分'] || inst.rateType || 'メイン';
    document.querySelectorAll('.sel-rate').forEach(s => { s.value = defaultRate; });
    updateTotals();
    loadDraftReport();
  }
}

function onLogout() {
  if (!confirm('ログアウトしますか？')) return;
  clearAuth();
  State.currentInstructor = null;
  State.submitId = null;
  State.rows = [];
  State.isSubmitted = false;
  State._pendingLoginName = null;
  rowIndex = 0;

  // フォームリセット
  document.getElementById('instructor-name').disabled = false;
  document.getElementById('instructor-name').value = '';
  document.getElementById('club-name').value = '';
  document.getElementById('report-tbody').innerHTML = '';
  updateSubmitButtonsUI();

  // ログイン画面のドロップダウンをリセット
  document.getElementById('login-instructor').value = '';
  document.getElementById('login-pin').value = '';
  document.getElementById('login-error').classList.add('hidden');

  showLoginSection();
}

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
    populateLoginSelect();
    // ログイン済みの場合はフォームにも反映
    if (State._pendingLoginName) {
      applyLoginToForm(State._pendingLoginName);
    }
  } catch (e) {
    showToast('指導者一覧の取得に失敗しました: ' + e.message, 'error');
  }
}

function populateLoginSelect() {
  const sel = document.getElementById('login-instructor');
  sel.innerHTML = '<option value="">選択してください</option>';
  State.instructors.forEach(inst => {
    const opt = document.createElement('option');
    opt.value = inst['氏名'] || inst.name || '';
    opt.textContent = (inst['氏名'] || inst.name || '') + '（' + (inst['クラブ名'] || inst.clubName || '') + '）';
    sel.appendChild(opt);
  });
}

function populateInstructorSelect() {
  // ログイン後はapplyLoginToFormが担うため、ここでは何もしない
}

async function onInstructorChange() {
  const name = document.getElementById('instructor-name').value;
  const inst = State.instructors.find(i => (i['氏名'] || i.name) === name);
  State.currentInstructor = inst || null;
  document.getElementById('club-name').value = inst ? (inst['クラブ名'] || inst.clubName || '') : '';

  if (inst) {
    const defaultRate = inst['区分'] || inst.rateType || 'メイン';
    document.querySelectorAll('.sel-rate').forEach(sel => { sel.value = defaultRate; });
    updateTotals();
    await loadDraftReport();
  }
}

function onYearMonthChange() {
  loadDraftReport();
}

async function loadDraftReport() {
  const name  = document.getElementById('instructor-name').value;
  const year  = document.getElementById('target-year').value;
  const month = document.getElementById('target-month').value;
  if (!name || !year || !month) return;

  State.isSubmitted = false;
  updateSubmitButtonsUI();

  showLoading();
  try {
    const data = await gasGet({ action: 'getReport', instructorName: name, year, month });
    if (!data.success) throw new Error(data.error);

    const allRows   = data.data || [];
    const submitted = allRows.filter(r => (r.status || '').trim() === '提出済');
    const drafts    = allRows.filter(r => (r.status || '').trim() === '下書き');

    if (submitted.length > 0) {
      State.submitId    = submitted[0].submitId;
      State.isSubmitted = true;
      document.getElementById('report-tbody').innerHTML = '';
      rowIndex = 0;
      submitted.forEach(r => addRowWithData(r));
      updateTotals();
      updateSubmitButtonsUI();
      showToast('提出済みデータを読み込みました', 'info');
    } else if (drafts.length > 0) {
      State.submitId    = drafts[0].submitId;
      State.isSubmitted = false;
      document.getElementById('report-tbody').innerHTML = '';
      rowIndex = 0;
      drafts.forEach(r => addRowWithData(r));
      updateTotals();
      updateSubmitButtonsUI();
      showToast('下書きデータを読み込みました', 'success');
    }
  } catch (e) {
    console.error('[loadDraftReport] エラー:', e);
    showToast('データの読み込みに失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function addRowWithData(r) {
  const tr = addRow();
  if (!tr) return;

  tr.querySelector('.inp-date').value        = parseDateStr(r.date);
  tr.querySelector('.sel-category').value    = r.category || '平日';
  tr.querySelector('.sel-rate').value        = r.rateType || 'メイン';
  tr.querySelector('.inp-start').value       = parseTimeStr(r.startTime);
  tr.querySelector('.inp-end').value         = parseTimeStr(r.endTime);
  tr.querySelector('.sel-transport').value   = r.transport || '';
  tr.querySelector('.inp-dest').value        = r.destination || '';
  tr.querySelector('.inp-travel').value      = r.travelAmount || '';
  tr.querySelector('.inp-note').value        = r.note || '';

  recalcRow(tr);
}

function parseDateStr(val) {
  if (!val) return '';
  const s = String(val);
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (s.includes('T')) return s.slice(0, 10);
  try {
    const d = new Date(s);
    if (!isNaN(d)) {
      const y  = d.getFullYear();
      const mo = String(d.getMonth() + 1).padStart(2, '0');
      const dy = String(d.getDate()).padStart(2, '0');
      return `${y}-${mo}-${dy}`;
    }
  } catch (_) {}
  return s;
}

function parseTimeStr(val) {
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

// ========== 行操作 ==========

let rowIndex = 0;

function addRow() {
  const tbody = document.getElementById('report-tbody');
  if (tbody.rows.length >= 31) {
    showToast('1ヶ月の最大行数（31行）に達しました', 'warning');
    return null;
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

  const prevStart = tbody.rows[0] ? tbody.rows[0].querySelector('.inp-start').value : '';
  const prevEnd   = tbody.rows[0] ? tbody.rows[0].querySelector('.inp-end').value   : '';

  tbody.insertBefore(tr, tbody.firstChild);

  if (prevStart) tr.querySelector('.inp-start').value = prevStart;
  if (prevEnd)   tr.querySelector('.inp-end').value   = prevEnd;
  if (prevStart || prevEnd) recalcRow(tr);

  return tr;
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

  document.getElementById('tab-nav').classList.add('hidden');
  document.getElementById('form-section').classList.add('hidden');
  document.getElementById('confirm-section').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function showFormSection() {
  document.getElementById('confirm-section').classList.add('hidden');
  document.getElementById('complete-section').classList.add('hidden');
  document.getElementById('tab-nav').classList.remove('hidden');
  document.getElementById('form-section').classList.remove('hidden');
  document.querySelectorAll('.tab-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === 'report');
  });
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
  document.getElementById('tab-nav').classList.add('hidden');
  document.getElementById('complete-section').classList.remove('hidden');
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function resetForm() {
  // ログイン状態は維持したまま月報入力だけリセット
  document.getElementById('report-tbody').innerHTML = '';
  State.submitId = null;
  State.rows = [];
  State.isSubmitted = false;
  rowIndex = 0;

  const auth = getAuth();
  if (auth && auth.name) {
    // 同一ユーザーで継続
    applyLoginToForm(auth.name);
  } else {
    document.getElementById('instructor-name').value = '';
    document.getElementById('club-name').value = '';
    State.currentInstructor = null;
    addRow();
  }

  updateTotals();
  updateSubmitButtonsUI();
  document.getElementById('complete-section').classList.add('hidden');
  document.getElementById('tab-nav').classList.remove('hidden');
  document.getElementById('form-section').classList.remove('hidden');
  document.querySelectorAll('.tab-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === 'report');
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function updateSubmitButtonsUI() {
  const submitted = State.isSubmitted;
  document.getElementById('save-btn').classList.toggle('hidden', submitted);
  document.getElementById('submit-btn').classList.toggle('hidden', submitted);
  document.getElementById('resubmit-btn').classList.toggle('hidden', !submitted);
  document.getElementById('submitted-notice').classList.toggle('hidden', !submitted);
}

function onResubmit() {
  if (!validate()) return;
  if (!confirm('提出済みのデータを上書きします。よろしいですか？')) return;
  State.isSubmitted = false;
  updateSubmitButtonsUI();
  onShowConfirm();
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

// ========== タブ切替 ==========

function switchTab(tab) {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tab);
  });

  if (tab === 'report') {
    document.getElementById('form-section').classList.remove('hidden');
    document.getElementById('payslip-section').classList.add('hidden');
  } else if (tab === 'payslip') {
    document.getElementById('form-section').classList.add('hidden');
    document.getElementById('payslip-section').classList.remove('hidden');
    // プレビューを初期化（年月選択のみ表示）
    document.getElementById('payslip-preview').classList.add('hidden');
    document.getElementById('payslip-not-calculated').classList.add('hidden');
  }
}

// ========== 給与明細 ==========

function initPayslipSelectors() {
  const now = new Date();
  const yearSel  = document.getElementById('payslip-year');
  const monthSel = document.getElementById('payslip-month');

  for (let y = now.getFullYear() - 1; y <= now.getFullYear() + 1; y++) {
    const opt = document.createElement('option');
    opt.value = y;
    opt.textContent = y;
    // 先月を初期選択
    const prevMonth = now.getMonth() === 0 ? 12 : now.getMonth();
    const prevYear  = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
    if (y === prevYear) opt.selected = true;
    yearSel.appendChild(opt);
  }
  for (let m = 1; m <= 12; m++) {
    const opt = document.createElement('option');
    opt.value = m;
    opt.textContent = m;
    const prevMonth = now.getMonth() === 0 ? 12 : now.getMonth();
    if (m === prevMonth) opt.selected = true;
    monthSel.appendChild(opt);
  }
}

async function onShowPaySlip() {
  const auth = getAuth();
  if (!auth || !auth.name) { showToast('ログインが必要です', 'error'); return; }

  const year  = parseInt(document.getElementById('payslip-year').value);
  const month = parseInt(document.getElementById('payslip-month').value);

  showLoading();
  document.getElementById('payslip-preview').classList.add('hidden');
  document.getElementById('payslip-not-calculated').classList.add('hidden');

  try {
    const data = await gasGet({
      action: 'getPaySlip',
      instructorName: auth.name,
      year,
      month,
    });
    if (!data.success) throw new Error(data.error);

    if (data.notCalculated) {
      document.getElementById('payslip-not-calculated').classList.remove('hidden');
    } else {
      renderPaySlip(data, year, month);
    }
  } catch (e) {
    showToast('給与明細の取得に失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

function renderPaySlip(data, year, month) {
  const { instructor, fee } = data;
  const today    = new Date();
  const payDate  = calcPaymentDate(year, month);

  const todayStr  = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;
  const payDateStr = `${payDate.getFullYear()}年${payDate.getMonth() + 1}月${payDate.getDate()}日`;

  // 指導内容明細行（メイン・サブ）
  const mainFee = Math.round((fee.mainCalcHours || 0) * Config.HOURLY_RATE.MAIN);
  const subFee  = Math.round((fee.subCalcHours  || 0) * Config.HOURLY_RATE.SUB);
  let rateRows = '';
  if (fee.mainCalcHours > 0) {
    rateRows += `<tr>
      <td>メイン指導</td>
      <td class="text-right">${fmtH(fee.mainCalcHours)}</td>
      <td class="text-right">¥${Config.HOURLY_RATE.MAIN.toLocaleString()}</td>
      <td class="text-right">¥${mainFee.toLocaleString()}</td>
    </tr>`;
  }
  if (fee.subCalcHours > 0) {
    rateRows += `<tr>
      <td>サブ指導</td>
      <td class="text-right">${fmtH(fee.subCalcHours)}</td>
      <td class="text-right">¥${Config.HOURLY_RATE.SUB.toLocaleString()}</td>
      <td class="text-right">¥${subFee.toLocaleString()}</td>
    </tr>`;
  }
  if (!rateRows) {
    rateRows = '<tr><td colspan="4" class="text-center text-muted">データなし</td></tr>';
  }

  // 区分別指導時間行
  const cats = [
    { label: '平日',     hours: fee.weekdayHours    || 0 },
    { label: '休日',     hours: fee.holidayHours    || 0 },
    { label: '長期休暇', hours: fee.longVacHours    || 0 },
    { label: '大会引率', hours: fee.tournamentHours || 0 },
  ].filter(c => c.hours > 0);

  const catTotal = (fee.weekdayHours || 0) + (fee.holidayHours || 0) +
                   (fee.longVacHours || 0) + (fee.tournamentHours || 0);
  const catRows = cats.map(c =>
    `<tr><td>${esc(c.label)}</td><td class="text-right">${fmtH(c.hours)}</td></tr>`
  ).join('');

  // 振込先口座
  const bankParts = [instructor.bank, instructor.branch, instructor.accountType, instructor.accountNumber]
    .filter(Boolean);
  const bankInfo = bankParts.length ? esc(bankParts.join('　')) : '未登録';

  const travelRow = (fee.travelTotal || 0) > 0
    ? `<tr><td>旅費</td><td class="text-right">¥${(fee.travelTotal).toLocaleString()}</td></tr>`
    : '';

  const catSection = cats.length > 0 ? `
    <h3 class="payslip-section-heading">区分別指導時間</h3>
    <table class="payslip-table">
      <thead><tr><th>指導区分</th><th class="text-right">謝金計算時間</th></tr></thead>
      <tbody>
        ${catRows}
        <tr class="payslip-subtotal">
          <td>合計</td>
          <td class="text-right">${fmtH(catTotal)}</td>
        </tr>
      </tbody>
    </table>` : '';

  document.getElementById('payslip-doc').innerHTML = `
    <div class="payslip-header-area">
      <div class="payslip-issuer-block">
        <div class="payslip-issuer-name">一般社団法人たかすスポーツクラブ</div>
        <div class="payslip-meta">
          <span>発行日：${todayStr}</span>
          <span>支払予定日：${payDateStr}</span>
        </div>
      </div>
      <div class="payslip-title-block">
        <div class="payslip-main-title">指導謝金支払明細書</div>
        <div class="payslip-period">${year}年${month}月分</div>
      </div>
    </div>

    <div class="payslip-info-grid">
      <table class="payslip-info-table">
        <tr><th>指導者氏名</th><td>${esc(instructor.name || '')}</td></tr>
        <tr><th>クラブ名</th><td>${esc(instructor.clubName || '')}</td></tr>
        <tr><th>区分</th><td>${esc(instructor.rateType || '')}</td></tr>
      </table>
      <table class="payslip-info-table">
        <tr><th>振込先口座</th><td>${bankInfo}</td></tr>
      </table>
    </div>

    <h3 class="payslip-section-heading">指導内容明細</h3>
    <table class="payslip-table">
      <thead>
        <tr>
          <th>区分</th>
          <th class="text-right">謝金計算時間</th>
          <th class="text-right">時給単価</th>
          <th class="text-right">謝金額</th>
        </tr>
      </thead>
      <tbody>${rateRows}</tbody>
    </table>

    ${catSection}

    <h3 class="payslip-section-heading" style="margin-top:20px;">支払金額</h3>
    <table class="payslip-table payslip-summary-table">
      <tbody>
        <tr><td>謝金総額</td><td class="text-right fw-bold">¥${(fee.totalFee || 0).toLocaleString()}</td></tr>
        <tr><td>源泉徴収額（10.21%）</td><td class="text-right text-danger">−¥${(fee.withholding || 0).toLocaleString()}</td></tr>
        <tr class="payslip-netpay">
          <td>差引支払額</td>
          <td class="text-right">¥${(fee.netPay || 0).toLocaleString()}</td>
        </tr>
        ${travelRow}
      </tbody>
    </table>
  `;

  document.getElementById('payslip-preview').classList.remove('hidden');
  document.getElementById('payslip-not-calculated').classList.add('hidden');
}

// 支払予定日：翌月10日、土日の場合は前の金曜日
function calcPaymentDate(year, month) {
  let payYear = year, payMonth = month + 1;
  if (payMonth > 12) { payMonth = 1; payYear++; }

  const d   = new Date(payYear, payMonth - 1, 10);
  const dow = d.getDay();
  let day   = 10;
  if      (dow === 6) day = 9;  // 土曜 → 前日(金)
  else if (dow === 0) day = 8;  // 日曜 → 2日前(金)

  return new Date(payYear, payMonth - 1, day);
}

// 時間表示（例: 1.5 → "1時間30分"）
function fmtH(h) {
  if (!h || h === 0) return '0時間';
  const hh = Math.floor(h);
  const mm = Math.round((h - hh) * 60);
  return hh + '時間' + (mm > 0 ? mm + '分' : '');
}

function printPaySlip() {
  window.print();
}
