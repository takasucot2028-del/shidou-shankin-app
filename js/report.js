/**
 * 月報入力ロジック
 */

// ========== 祝日・曜日定数 ==========

const WEEKDAY_NAMES = ['日', '月', '火', '水', '木', '金', '土'];

/**
 * 祝日は年から計算する。
 *
 * ※以前は2027年までの固定リストだったため、2028年以降は祝日が0件になり、
 *   指導者画面の概算（教員の長期休暇の扱い・曜日表示）がGAS側の確定値と
 *   ズレる状態だった。GAS側 getJPHolidaySet / 事務局画面 getJPHolidays と
 *   同一ロジックに揃えている。3実装のいずれかだけを変更しないこと。
 */
const __jpHolidayCache = {};

function getJPHolidaySetForYear(year) {
  if (__jpHolidayCache[year]) return __jpHolidayCache[year];

  const p = n => String(n).padStart(2, '0');
  const set = new Set();
  const add = (m, d) => set.add(year + '-' + p(m) + '-' + p(d));

  const shunbun = Math.floor(20.8431 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  const shubun  = Math.floor(23.2488 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  const nthMon = (m, n) => {
    const d = new Date(year, m - 1, 1);
    let c = 0;
    for (;;) {
      if (d.getDay() === 1) { c++; if (c === n) return d.getDate(); }
      d.setDate(d.getDate() + 1);
    }
  };

  add(1, 1);
  add(2, 11);
  if (year >= 2020) add(2, 23);
  add(3, shunbun);
  add(4, 29);
  add(5, 3); add(5, 4); add(5, 5);
  add(8, 11);
  add(9, shubun);
  add(11, 3);
  add(11, 23);
  add(1,  nthMon(1,  2));
  add(7,  nthMon(7,  3));
  add(9,  nthMon(9,  3));
  add(10, nthMon(10, 2));

  // 振替休日：祝日が日曜なら翌月曜（既存祝日と重なれば更に翌日）
  const key2date = k => { const q = k.split('-'); return new Date(Number(q[0]), Number(q[1]) - 1, Number(q[2])); };
  const toKey = d => d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  Array.from(set).forEach(k => {
    const d = key2date(k);
    if (d.getDay() === 0) {
      const sub = new Date(d);
      sub.setDate(sub.getDate() + 1);
      while (set.has(toKey(sub))) sub.setDate(sub.getDate() + 1);
      set.add(toKey(sub));
    }
  });

  __jpHolidayCache[year] = set;
  return set;
}

/** "yyyy-MM-dd" 形式の日付が祝日かを返す */
function isJPHolidayKey(key) {
  if (!key) return false;
  const s = String(key).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return false;
  return getJPHolidaySetForYear(Number(m[1])).has(m[1] + '-' + m[2] + '-' + m[3]);
}

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
  try {
    const a = JSON.parse(sessionStorage.getItem(AUTH_KEY));
    if (!a || !a.token) return null;
    // 期限切れのトークンは保持しない（サーバー側でも検証している）
    if (a.expiresAt && a.expiresAt < Date.now()) { clearAuth(); return null; }
    return a;
  } catch { return null; }
}
function setAuth(name, token, expiresAt) {
  sessionStorage.setItem(AUTH_KEY, JSON.stringify({ name, token, expiresAt }));
}
function clearAuth() { sessionStorage.removeItem(AUTH_KEY); }

/** 現在のトークン。未ログインなら空文字 */
function authToken() {
  const a = getAuth();
  return a && a.token ? a.token : '';
}

/**
 * サーバーが認証切れを返したときの共通処理。
 * 該当すれば true を返すので、呼び出し側はそこで中断する。
 */
function handleAuthError(res) {
  if (res && (res.authRequired || res.forbidden)) {
    clearAuth();
    showLoginSection();
    showToast(res.error || '再ログインが必要です', 'error');
    return true;
  }
  return false;
}

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
  document.getElementById('pdf-report-btn').addEventListener('click', onPrintReport);

  // タブ切替
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  // 給与明細ボタン
  document.getElementById('payslip-show-btn').addEventListener('click', onShowPaySlip);
  document.getElementById('payslip-pdf-btn').addEventListener('click', printPaySlip);

  // アプリ内ブラウザ（LINE等）ではPDF保存不可の案内を表示
  if (isInAppBrowser()) {
    const warn = document.getElementById('inapp-warning');
    if (warn) warn.classList.remove('hidden');
  }
  const copyUrlBtn = document.getElementById('copy-url-btn');
  if (copyUrlBtn) {
    copyUrlBtn.addEventListener('click', async () => {
      try {
        await navigator.clipboard.writeText(location.href);
        showToast('URLをコピーしました。Chromeに貼り付けて開いてください', 'success');
      } catch {
        showToast('コピーできませんでした。アドレスバーのURLを手動でコピーしてください', 'error');
      }
    });
  }

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
    setAuth(res.name || name, res.token, res.expiresAt);
    showFormAfterLogin(res.name || name);
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
  // サーバー側のトークンも失効させる（失敗しても画面はログアウトする）
  const t = authToken();
  if (t) { gasPost({ action: 'logout', token: t }).catch(() => {}); }

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
    // 口座情報を含まない公開用の一覧を使う（この画面はログイン前に呼ぶ）
    const data = await gasPost({ action: 'getInstructorsPublic' });
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

  // 対象年月を切り替えたら、前月の行と提出IDを必ず破棄する。
  // ※残したままだと、データのない月に切り替えた際に前月の日付の行が
  //   そのまま別月の月報として提出でき、同じ指導が二重に支給される。
  State.isSubmitted = false;
  State.submitId    = null;
  document.getElementById('report-tbody').innerHTML = '';
  rowIndex = 0;
  updateSubmitButtonsUI();

  showLoading();
  try {
    // 月報はトークン必須のためPOSTで取得する（トークンをURLに残さない）
    const data = await gasPost({
      action: 'getReport', token: authToken(), instructorName: name, year, month,
    });
    if (handleAuthError(data)) return;
    if (!data.success) throw new Error(data.error);

    const allRows   = data.data || [];
    // 内容が完全一致する重複行を除外（過去の同時保存等でシートに残った二重行対策）
    const submitted = dedupeReportRows(allRows.filter(r => (r.status || '').trim() === '提出済'));
    const drafts    = dedupeReportRows(allRows.filter(r => (r.status || '').trim() === '下書き'));

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
    } else {
      // この月のデータは無い。空行1行だけの状態から入力を始める
      addRow();
      updateTotals();
      updateSubmitButtonsUI();
    }
  } catch (e) {
    console.error('[loadDraftReport] エラー:', e);
    showToast('データの読み込みに失敗しました: ' + e.message, 'error');
  } finally {
    hideLoading();
  }
}

// 内容が完全一致する行を重複とみなして除外する（同日複数コマ等の正当な別行は残す）
function dedupeReportRows(rows) {
  const seen = new Set();
  const out = [];
  rows.forEach(r => {
    const key = [
      r.date, r.category, r.rateType, r.startTime, r.endTime,
      r.transport, r.destination, r.travelAmount, r.note,
    ].join('|');
    if (seen.has(key)) return;
    seen.add(key);
    out.push(r);
  });
  return out;
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
  updateWeekdayLabel(tr);
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
      <div class="date-cell-wrap">
        <input type="date" class="inp-date" aria-label="日付">
        <span class="weekday-label"></span>
      </div>
    </td>
    <td data-label="内容">
      <select class="sel-category" aria-label="内容">
        <option value="平日">平日練習</option>
        <option value="休日">休日練習</option>
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
    <td data-label="場所">
      <input type="text" class="inp-dest" placeholder="場所" aria-label="場所">
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
  tr.querySelector('.inp-date').addEventListener('change', () => updateWeekdayLabel(tr));
  tr.querySelector('.delete-row-btn').addEventListener('click', () => {
    if (document.getElementById('report-tbody').rows.length > 1) {
      tr.remove();
      updateTotals();
    } else {
      showToast('最低1行は必要です', 'warning');
    }
  });

  const prevRow   = tbody.rows[0] || null;
  const prevStart = prevRow ? prevRow.querySelector('.inp-start').value : '';
  const prevEnd   = prevRow ? prevRow.querySelector('.inp-end').value   : '';
  const prevDest  = prevRow ? prevRow.querySelector('.inp-dest').value  : '';

  tbody.insertBefore(tr, tbody.firstChild);

  if (prevStart) tr.querySelector('.inp-start').value = prevStart;
  if (prevEnd)   tr.querySelector('.inp-end').value   = prevEnd;
  if (prevDest)  tr.querySelector('.inp-dest').value  = prevDest;
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

  const date  = tr.querySelector('.inp-date').value;
  const hours = calcInstructionHours(start, end, category, instrType, date);
  const hoursInput = tr.querySelector('.inp-hours');
  hoursInput.value = hours > 0 ? formatHours(hours) : '';
  hoursInput.dataset.hours = hours;

  updateTotals();
}

// ========== 謝金計算ロジック（フロント側・概算） ==========

/**
 * 教員の勤務時間控除（16:40以降のみ有効）を適用すべき行かを判定する。
 * 平日は常に対象。長期休暇は日付が平日のときだけ対象とし、土日は
 * 学校勤務がないため通常の土日と同じ扱い（控除なし）にする。
 * ※GAS側 isTeacherCutTarget と同一ロジック。片方だけ変更しないこと。
 */
function isTeacherCutTargetRow(category, dateValue) {
  const cat = String(category).trim();
  if (cat === '平日') return true;
  if (cat !== '長期休暇') return false;

  const d = parseDateValueRow(dateValue);
  if (d === null) return true;        // 日付不明なら従来どおり控除する
  const day = d.getDay();
  if (day === 0 || day === 6) return false;   // 土日は控除しない
  // 祝日も土日と同じ扱い
  const p = n => String(n).padStart(2, '0');
  const key = d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  if (isJPHolidayKey(key)) return false;
  return true;
}

function parseDateValueRow(dateValue) {
  if (!dateValue) return null;
  if (dateValue instanceof Date && !isNaN(dateValue.getTime())) return dateValue;
  const m = String(dateValue).trim().match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (!m) return null;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

function calcInstructionHours(startTime, endTime, category, instructorType, dateValue) {
  const startMin = timeToMinutes(startTime);
  const endMin   = timeToMinutes(endTime);
  if (startMin === null || endMin === null || endMin <= startMin) return 0;

  let effectiveMin;
  // 長期休暇中の土日は学校勤務がないため、教員の勤務時間控除を行わない
  const isTeacher = String(instructorType).trim() === '教員';
  const isWeekdayLike = isTeacherCutTargetRow(category, dateValue);

  if (isTeacher && isWeekdayLike) {
    const effectiveStart = Math.max(startMin, Config.TEACHER_WORK_END_MINUTES);
    effectiveMin = Math.max(endMin - effectiveStart, 0);
  } else {
    effectiveMin = endMin - startMin;
  }

  const maxMin     = (Config.MAX_HOURS[String(category).trim()] || 0) * 60;
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
    const calcH = calcInstructionHours(start, end, category, instrType, date);

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
    // 日付が対象年月の範囲内か検証する（別月の実績が混入すると二重支給になる）
    const ym = String(r.date).slice(0, 7);
    const expected = year + '-' + String(month).padStart(2, '0');
    if (ym !== expected) {
      showToast('対象年月（' + year + '年' + month + '月）以外の日付が含まれています: ' + r.date, 'error');
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
    if (handleAuthError(res)) return;
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
    token:          authToken(),
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
    if (handleAuthError(res)) return;
    if (!res.success) throw new Error(res.error);
    State.submitId   = res.submitId;
    State.isSubmitted = true;
    updateSubmitButtonsUI();
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
  // ※ここで isSubmitted を降ろすと、確認画面から「戻って修正する」で
  //   提出済バッジが消えて一時保存ボタンが復活し、提出済行を残したまま
  //   下書き行が併存する（その後の修正が画面に出なくなる）。
  //   実際に提出が成功した時点で更新する。
  onShowConfirm();
}

// ========== GAS 通信 ==========

/**
 * GET通信。
 * ※認証が必要なアクションはトークンがURLに残るため使用しないこと。
 *   現在このアプリでは使用していない（すべてPOSTに統一）。
 */
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

/**
 * POST通信。GASの仕様上まれに起きる「POSTがGETとして処理される」ケースを再送する。
 *
 * この現象が起きたとき、サーバーは getNotAllowed を返す。これは doGet に
 * 着弾したという意味で、アクションは一切実行されていない。副作用が出ていない
 * ことが確定しているため、再送しても二重登録・二重送信は起きない。
 * 逆に、通信エラーやタイムアウトは「実行されたかどうか不明」なので再送しない。
 */
async function gasPost(body) {
  const MAX_ATTEMPTS = 3;
  const RETRY_DELAYS = [600, 1800];   // ミリ秒

  for (let attempt = 1; ; attempt++) {
    const res = await fetch(Config.GAS_URL, {
      method:   'POST',
      redirect: 'follow',
      // Content-Type を text/plain にすることでプリフライトを回避
      headers:  { 'Content-Type': 'text/plain;charset=UTF-8' },
      body:     JSON.stringify(body),
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);

    const json = await res.json();
    if (!json || !json.getNotAllowed || attempt >= MAX_ATTEMPTS) return json;

    const wait = RETRY_DELAYS[attempt - 1] || 1800;
    console.warn('[gasPost] POSTがGETとして処理されたため再送します (' +
      attempt + '/' + (MAX_ATTEMPTS - 1) + ', ' + wait + 'ms後): ' + body.action);
    await new Promise(ok => setTimeout(ok, wait));
  }
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
    // 給与明細は口座情報を含むためトークン必須・POSTで取得する
    const data = await gasPost({
      action: 'getPaySlip',
      token: authToken(),
      instructorName: auth.name,
      year,
      month,
    });
    if (handleAuthError(data)) return;
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
  // PDFダウンロード用にメタ情報を保持
  State.lastPaySlip = { name: instructor.name || '', year: year, month: month };
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
    { label: '平日練習', hours: fee.weekdayHours    || 0 },
    { label: '休日練習', hours: fee.holidayHours    || 0 },
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

// 外部スクリプトを一度だけ動的に読み込む
function loadScriptOnce(src) {
  return new Promise((resolve, reject) => {
    const existing = Array.from(document.scripts).find(s => s.src === src);
    if (existing) {
      if (existing.dataset.loaded === '1') return resolve();
      existing.addEventListener('load', () => resolve());
      existing.addEventListener('error', () => reject(new Error('読み込み失敗: ' + src)));
      return;
    }
    const s = document.createElement('script');
    s.src = src;
    s.onload  = () => { s.dataset.loaded = '1'; resolve(); };
    s.onerror = () => reject(new Error('読み込み失敗: ' + src));
    document.head.appendChild(s);
  });
}

// PDF生成ライブラリ（html2canvas / jsPDF）を確実に用意する
async function ensurePdfLibs() {
  if (typeof window.html2canvas !== 'function') {
    await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
  }
  if (!window.jspdf) {
    await loadScriptOnce('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
  }
  if (typeof window.html2canvas !== 'function' || !window.jspdf) {
    throw new Error('PDFライブラリを読み込めませんでした（通信環境をご確認ください）');
  }
}

async function printPaySlip() {
  const el = document.getElementById('payslip-doc');
  if (!el || !el.innerHTML.trim()) {
    showToast('先に「表示」で給与明細を表示してください', 'error');
    return;
  }

  const meta  = State.lastPaySlip || {};
  const fname = `給与明細_${meta.name || ''}_${meta.year || ''}年${meta.month || ''}月.pdf`;

  let clone = null;
  showLoading();
  try {
    // 必要なら CDN からライブラリを取得（index.html の読み込み失敗にも対応）
    await ensurePdfLibs();

    // A4幅(約794px)固定のクローンを画面外に作り、端末によらず同じレイアウトで描画する
    clone = el.cloneNode(true);
    clone.removeAttribute('id');
    clone.style.position  = 'fixed';
    clone.style.left      = '-10000px';
    clone.style.top       = '0';
    clone.style.width     = '794px';
    clone.style.maxWidth  = 'none';
    clone.style.margin    = '0';
    clone.style.boxShadow = 'none';
    clone.style.background = '#ffffff';
    document.body.appendChild(clone);

    const canvas = await html2canvas(clone, {
      scale: 2,
      backgroundColor: '#ffffff',
      windowWidth: 794,
    });
    const imgData = canvas.toDataURL('image/jpeg', 0.95);
    const { jsPDF } = window.jspdf;
    const pdf   = new jsPDF('p', 'mm', 'a4');
    const pageW = pdf.internal.pageSize.getWidth();
    const pageH = pdf.internal.pageSize.getHeight();
    const imgH  = canvas.height * pageW / canvas.width;

    // 1ページに収まらない場合はページ送りして全体を出力
    let heightLeft = imgH;
    let position   = 0;
    pdf.addImage(imgData, 'JPEG', 0, position, pageW, imgH);
    heightLeft -= pageH;
    while (heightLeft > 0) {
      position -= pageH;
      pdf.addPage();
      pdf.addImage(imgData, 'JPEG', 0, position, pageW, imgH);
      heightLeft -= pageH;
    }
    await outputPdf(pdf, fname);
  } catch (e) {
    console.error('[printPaySlip] PDF生成エラー:', e);
    showToast('PDF生成に失敗しました: ' + (e && e.message ? e.message : e), 'error');
  } finally {
    if (clone && clone.parentNode) clone.parentNode.removeChild(clone);
    hideLoading();
  }
}

// LINE・各種SNSのアプリ内ブラウザ（WebView）かどうかを判定する
function isInAppBrowser() {
  const ua = navigator.userAgent || '';
  return /Line\/|FBAN|FBAV|Instagram|Twitter|MicroMessenger|TikTok/i.test(ua);
}

// モバイル端末（スマホ・タブレット）かどうかを判定する
function isMobileDevice() {
  if (navigator.userAgentData && typeof navigator.userAgentData.mobile === 'boolean') {
    return navigator.userAgentData.mobile;
  }
  const ua = navigator.userAgent || '';
  // iPadOS は Mac を名乗るためタッチ対応で補完
  if (/iPad|iPhone|iPod|Android|Mobile/i.test(ua)) return true;
  return navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1;
}

// 生成したPDFを端末に応じて出力する
// スマホは Web Share（共有→「ファイルに保存」等）を優先。非対応環境はダウンロード/別タブ表示にフォールバック。
async function outputPdf(pdf, fname) {
  const blob = pdf.output('blob');

  // 1) Web Share API（スマホで最も確実。アプリ内ブラウザのダウンロード非対応を回避）
  //    PCは共有シートではなく通常のダウンロードを使うため、モバイル端末に限定する。
  try {
    const file = new File([blob], fname, { type: 'application/pdf' });
    if (isMobileDevice() && navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], title: fname });
      return;
    }
  } catch (err) {
    if (err && err.name === 'AbortError') return; // ユーザーが共有をキャンセル
    // それ以外は次の手段へフォールバック
  }

  // 2) アプリ内ブラウザ（LINE等）はダウンロード非対応。外部ブラウザで開くよう案内する。
  if (isInAppBrowser()) {
    alert(
      'LINEなどのアプリ内ブラウザではPDFを保存できません。\n\n' +
      '画面右上のメニュー（「︙」や「↗」アイコン）から\n' +
      '「ブラウザで開く」「Chromeで開く」「他のアプリで開く」を選び、\n' +
      'Chrome などのブラウザで開き直してから、もう一度お試しください。'
    );
    return;
  }

  const url = URL.createObjectURL(blob);

  // 3) ダウンロード（PC・通常のモバイルブラウザ）
  try {
    const a = document.createElement('a');
    a.href = url;
    a.download = fname;
    a.rel = 'noopener';
    document.body.appendChild(a);
    a.click();
    a.remove();
  } catch (err) {
    // 4) ダウンロード不可環境では別タブで表示
    window.open(url, '_blank');
  }

  setTimeout(() => URL.revokeObjectURL(url), 60000);
}

// ========== 曜日表示 ==========

function updateWeekdayLabel(tr) {
  const dateVal = tr.querySelector('.inp-date').value;
  const span    = tr.querySelector('.weekday-label');
  if (!span) return;

  if (!dateVal) {
    span.textContent = '';
    span.className   = 'weekday-label';
    return;
  }

  const d = new Date(dateVal + 'T00:00:00');
  if (isNaN(d)) {
    span.textContent = '';
    span.className   = 'weekday-label';
    return;
  }

  const dow       = d.getDay();
  const isHoliday = isJPHolidayKey(dateVal);
  span.textContent = '(' + WEEKDAY_NAMES[dow] + ')';

  let cls = 'weekday-label';
  if (isHoliday || dow === 0) cls += ' weekday-sun';
  else if (dow === 6)         cls += ' weekday-sat';
  span.className = cls;
}

// ========== 月報 PDF出力 ==========

function onPrintReport() {
  const name  = document.getElementById('instructor-name').value;
  const club  = document.getElementById('club-name').value;
  const year  = document.getElementById('target-year').value;
  const month = document.getElementById('target-month').value;
  const rows  = collectRows();

  if (!name) { showToast('指導者氏名を選択してください', 'error'); return; }
  if (rows.length === 0) { showToast('出力する指導記録がありません', 'error'); return; }

  const sorted = rows.slice().sort((a, b) => (a.date || '').localeCompare(b.date || ''));

  const instrType = State.currentInstructor
    ? (State.currentInstructor['指導者種別'] || State.currentInstructor.type || '一般')
    : '一般';

  const today    = new Date();
  const todayStr = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日`;
  const preview  = calcFeePreview(sorted, instrType);

  const tableRows = sorted.map(r => {
    const d        = r.date ? new Date(r.date + 'T00:00:00') : null;
    const dow      = d ? d.getDay() : -1;
    const wday     = d ? '(' + WEEKDAY_NAMES[dow] + ')' : '';
    const isHol    = r.date ? isJPHolidayKey(r.date) : false;
    const dateDisp = r.date ? r.date.slice(5).replace('-', '/') : '';

    let dateColor = '';
    if (isHol || dow === 0) dateColor = 'color:#c62828;';
    else if (dow === 6)     dateColor = 'color:#1565c0;';

    return `<tr>
      <td style="${dateColor}white-space:nowrap;">${esc(dateDisp)} ${wday}</td>
      <td>${esc(Config.CATEGORY_LABEL[r.category] || r.category || '')}</td>
      <td>${esc(r.rateType || '')}</td>
      <td>${r.startTime || ''}</td>
      <td>${r.endTime || ''}</td>
      <td>${r.instructionHours > 0 ? formatHours(r.instructionHours) : ''}</td>
      <td>${esc(r.transport || '')}</td>
      <td>${esc(r.destination || '')}</td>
      <td>${r.travelAmount > 0 ? '¥' + Number(r.travelAmount).toLocaleString() : ''}</td>
      <td>${esc(r.note || '')}</td>
    </tr>`;
  }).join('');

  const catMap = {};
  sorted.forEach(r => {
    const cat = r.category || '平日';
    if (!catMap[cat]) catMap[cat] = { count: 0, hours: 0 };
    catMap[cat].count++;
    catMap[cat].hours += r.instructionHours || 0;
  });
  const catRows = ['平日', '休日', '長期休暇', '大会引率']
    .filter(c => catMap[c])
    .map(c => `<tr><td>${Config.CATEGORY_LABEL[c] || c}</td><td>${catMap[c].count}回</td><td>${formatHours(catMap[c].hours)}</td></tr>`)
    .join('');

  const mainFeeAmt = Math.round(preview.mainHours * Config.HOURLY_RATE.MAIN);
  const subFeeAmt  = Math.round(preview.subHours  * Config.HOURLY_RATE.SUB);

  const html = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>${esc(year)}年${esc(month)}月分 指導月報 - ${esc(name)}</title>
  <style>
    * { box-sizing: border-box; }
    body { font-family: 'Hiragino Sans', 'Meiryo', sans-serif; font-size: 10pt; margin: 0; padding: 14mm 16mm; color: #222; }
    h1 { font-size: 15pt; text-align: center; margin: 0 0 4px; color: #1F4E79; }
    .period { text-align: center; font-size: 12pt; margin: 0 0 14px; color: #444; }
    .info-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-bottom: 14px; border: 1px solid #c8d8ec; padding: 10px 14px; border-radius: 4px; background: #f5f9ff; }
    .info-item { display: flex; flex-direction: column; gap: 1px; }
    .info-label { font-size: 8pt; color: #666; }
    .info-value { font-weight: bold; font-size: 10pt; }
    .section-title { font-size: 10pt; font-weight: bold; color: #1F4E79; border-left: 3px solid #1F4E79; padding-left: 7px; margin: 0 0 6px; }
    table { width: 100%; border-collapse: collapse; font-size: 8.5pt; margin-bottom: 14px; }
    th { background: #1F4E79; color: white; padding: 5px 5px; text-align: center; border: 1px solid #1F4E79; white-space: nowrap; }
    td { padding: 4px 5px; border: 1px solid #ccc; vertical-align: middle; text-align: center; }
    tr:nth-child(even) td { background: #f5f8ff; }
    .summary-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
    .fee-table td:first-child { text-align: left; }
    .fee-net td { font-weight: bold; font-size: 12pt; color: #c62828; background: #fff0f0 !important; border-top: 2px solid #c62828; }
    .footer { margin-top: 20px; border-top: 1px solid #ccc; padding-top: 8px; display: flex; justify-content: space-between; font-size: 8pt; color: #666; }
    @media print { body { padding: 10mm 13mm; } }
  </style>
</head>
<body>
  <h1>指導月報</h1>
  <div class="period">${esc(year)}年${esc(month)}月分</div>

  <div class="info-grid">
    <div class="info-item"><span class="info-label">指導者氏名</span><span class="info-value">${esc(name)}</span></div>
    <div class="info-item"><span class="info-label">クラブ名</span><span class="info-value">${esc(club)}</span></div>
    <div class="info-item"><span class="info-label">指導者種別</span><span class="info-value">${esc(instrType)}</span></div>
  </div>

  <div class="section-title">指導記録</div>
  <table>
    <thead>
      <tr>
        <th>日付</th><th>内容</th><th>時給区分</th><th>開始</th><th>終了</th>
        <th>指導時間</th><th>交通手段</th><th>場所</th><th>旅費</th><th>備考</th>
      </tr>
    </thead>
    <tbody>${tableRows}</tbody>
  </table>

  <div class="summary-grid">
    <div>
      <div class="section-title">内容別集計</div>
      <table>
        <thead><tr><th>内容</th><th>回数</th><th>指導時間</th></tr></thead>
        <tbody>${catRows}</tbody>
      </table>
    </div>
    <div>
      <div class="section-title">謝金概算</div>
      <table class="fee-table">
        <tbody>
          <tr><td>メイン指導（${preview.mainHours.toFixed(2)}h × ¥${Config.HOURLY_RATE.MAIN.toLocaleString()}）</td><td>¥${mainFeeAmt.toLocaleString()}</td></tr>
          <tr><td>サブ指導（${preview.subHours.toFixed(2)}h × ¥${Config.HOURLY_RATE.SUB.toLocaleString()}）</td><td>¥${subFeeAmt.toLocaleString()}</td></tr>
          <tr><td>謝金総額</td><td><strong>¥${preview.fee.toLocaleString()}</strong></td></tr>
          <tr><td>源泉徴収額（10.21%）</td><td style="color:#c62828;">−¥${preview.withholding.toLocaleString()}</td></tr>
          <tr class="fee-net"><td>差引支払額</td><td>¥${preview.netPay.toLocaleString()}</td></tr>
          <tr><td>旅費合計</td><td>¥${preview.travel.toLocaleString()}</td></tr>
        </tbody>
      </table>
    </div>
  </div>

  <div class="footer">
    <span>作成日：${todayStr}</span>
    <span>一般社団法人たかすスポーツクラブ</span>
  </div>
  <script>window.onload = function() { window.print(); };<\/script>
</body>
</html>`;

  const win = window.open('', '_blank');
  if (!win) {
    showToast('ポップアップがブロックされました。ブラウザの設定で許可してください。', 'error');
    return;
  }
  win.document.write(html);
  win.document.close();
}
