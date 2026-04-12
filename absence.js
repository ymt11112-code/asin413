const { ipcRenderer } = require('electron');

let appData = {};
let studentList = [];
let absenceRecords = [];
let absenceEditRowIndex = null;
let parsedPreviewRecords = [];
let selectedLeaveType = '事假';
let selectedMeal = '否';
let speechRecognition = null;
let speechListening = false;

const LEAVE_TYPES = ['事假', '病假', '公假', '喪假', '曠課'];
const SUMMARY_TYPES = ['事假', '病假', '公假', '喪假'];
const TARDY_NOTE_TEMPLATE = '到校時間（）';
const EARLY_LEAVE_NOTE_TEMPLATE = '離校時間（）';

function $(id) {
  return document.getElementById(id);
}

function todayIso() {
  const now = new Date();
  return [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, '0'),
    String(now.getDate()).padStart(2, '0')
  ].join('-');
}

function parseIsoDate(iso) {
  if (!iso) return null;
  const date = new Date(iso + 'T00:00:00');
  return Number.isNaN(date.getTime()) ? null : date;
}

function toIsoFromParts(year, month, day) {
  if (!year || !month || !day) return '';
  return [
    year,
    String(month).padStart(2, '0'),
    String(day).padStart(2, '0')
  ].join('-');
}

function addWeekdays(startIso, extraDays) {
  const start = parseIsoDate(startIso);
  if (!start) return startIso;
  if (extraDays <= 0) return startIso;
  const cursor = new Date(start);
  let remaining = extraDays;
  while (remaining > 0) {
    cursor.setDate(cursor.getDate() + 1);
    const day = cursor.getDay();
    if (day !== 0 && day !== 6) remaining -= 1;
  }
  return toIsoFromParts(cursor.getFullYear(), cursor.getMonth() + 1, cursor.getDate());
}

function countWeekdaysInclusive(startIso, endIso) {
  const start = parseIsoDate(startIso);
  const end = parseIsoDate(endIso);
  if (!start || !end || end < start) return 0;
  const cursor = new Date(start);
  let count = 0;
  while (cursor <= end) {
    const day = cursor.getDay();
    if (day !== 0 && day !== 6) count += 1;
    cursor.setDate(cursor.getDate() + 1);
  }
  return count;
}

function calculateDays(startIso, endIso, periods) {
  if (!startIso) return '';
  if (!endIso || startIso === endIso) {
    const periodCount = parseFloat(periods) || 0;
    if (!periodCount) return '';
    return periodCount < 4 ? '0.5' : '1';
  }
  const weekdays = countWeekdaysInclusive(startIso, endIso);
  return weekdays ? String(weekdays) : '';
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatDateRange(startDate, endDate) {
  if (!startDate) return '';
  if (!endDate || endDate === startDate) return startDate;
  return `${startDate} ~ ${endDate}`;
}

function chineseNumberToInt(text) {
  if (!text) return null;
  if (/^\d+$/.test(text)) return parseInt(text, 10);
  const map = { 一: 1, 二: 2, 兩: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
  if (text === '十') return 10;
  if (text.length === 2 && text[0] === '十') return 10 + (map[text[1]] || 0);
  if (text.length === 2 && text[1] === '十') return (map[text[0]] || 0) * 10;
  if (text.length === 3 && text[1] === '十') return (map[text[0]] || 0) * 10 + (map[text[2]] || 0);
  return map[text] || null;
}

function parseDateTokens(text) {
  const currentYear = new Date().getFullYear();
  const results = [];

  for (const match of text.matchAll(/(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})(?:日|號)?/g)) {
    results.push(toIsoFromParts(parseInt(match[1], 10), parseInt(match[2], 10), parseInt(match[3], 10)));
  }
  for (const match of text.matchAll(/(^|[^\d])(\d{1,2})月\s*(\d{1,2})(?:日|號)?/g)) {
    results.push(toIsoFromParts(currentYear, parseInt(match[2], 10), parseInt(match[3], 10)));
  }
  for (const match of text.matchAll(/(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/g)) {
    results.push(toIsoFromParts(parseInt(match[1], 10), parseInt(match[2], 10), parseInt(match[3], 10)));
  }
  for (const match of text.matchAll(/(^|[^\d])(\d{1,2})[\/.-](\d{1,2})(?!\d)/g)) {
    results.push(toIsoFromParts(currentYear, parseInt(match[2], 10), parseInt(match[3], 10)));
  }

  return Array.from(new Set(results)).filter(Boolean);
}

function stripDatePhrases(text) {
  return text
    .replace(/\d{4}年\s*\d{1,2}月\s*\d{1,2}(?:日|號)?/g, ' ')
    .replace(/\d{1,2}月\s*\d{1,2}(?:日|號)?/g, ' ')
    .replace(/\d{4}[\/.-]\d{1,2}[\/.-]\d{1,2}/g, ' ')
    .replace(/\d{1,2}[\/.-]\d{1,2}/g, ' ');
}

function inferModeFromText(text) {
  if (text.includes('早退')) return 'early';
  if (text.includes('遲到')) return 'late';
  return 'leave';
}

function inferLeaveTypeFromText(text) {
  return LEAVE_TYPES.find((type) => text.includes(type)) || '事假';
}

function inferMealChoiceFromText(text) {
  if (text.includes('不退餐') || text.includes('免退餐')) return '否';
  if (text.includes('要退餐') || text.includes('需退餐') || text.includes('需要退餐') || text.includes('退餐')) return '是';
  return null;
}

function inferPeriodsFromText(text) {
  const normalized = text
    .replace(/第/g, '')
    .replace(/節/g, '')
    .replace(/到/g, '-')
    .replace(/至/g, '-')
    .replace(/~/g, '-')
    .replace(/～/g, '-');

  const rangeMatch = normalized.match(/(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)/);
  if (rangeMatch) {
    const start = parseFloat(rangeMatch[1]);
    const end = parseFloat(rangeMatch[2]);
    if (!Number.isNaN(start) && !Number.isNaN(end) && end >= start) {
      return String(end - start + 1);
    }
  }

  const singleMatch = text.match(/(?:第\s*)?(\d+(?:\.\d+)?)\s*節/);
  if (singleMatch) return String(singleMatch[1]);

  return '';
}

function inferDurationDays(text) {
  const match = text.match(/([一二兩三四五六七八九十\d]+)\s*天/);
  return match ? chineseNumberToInt(match[1]) : null;
}

function findStudentBySeatNumber(seatNumber) {
  const normalizedSeat = String(parseInt(seatNumber, 10));
  return studentList.find((student) => String(parseInt(student.num, 10)) === normalizedSeat) || null;
}

function buildStudentLabel(student) {
  return `${student.num} ${student.name}`;
}

function inferStudentsFromText(text) {
  if (text.includes('全班')) return ['全班'];

  const matches = [];
  const textWithoutDates = stripDatePhrases(text);

  for (const match of textWithoutDates.matchAll(/(^|[^\d])(\d{1,2})號(?:同學)?(?!\d)/g)) {
    const student = findStudentBySeatNumber(match[2]);
    if (student) matches.push(buildStudentLabel(student));
  }

  studentList.forEach((student) => {
    const label = buildStudentLabel(student);
    const byFullLabel = text.includes(label);
    const byName = text.includes(student.name);
    const bySeatContext = new RegExp(`(^|\\D)${parseInt(student.num, 10)}號(?:同學)?(\\D|$)`).test(textWithoutDates);
    if (byFullLabel || byName || bySeatContext) matches.push(label);
  });

  return Array.from(new Set(matches));
}

function buildParsedRecords(text) {
  const mode = inferModeFromText(text);
  const students = inferStudentsFromText(text);
  const dates = parseDateTokens(text);
  const durationDays = inferDurationDays(text);
  const startDate = dates[0] || todayIso();

  let endDate = dates[1] || startDate;
  if (mode === 'leave' && dates.length <= 1 && durationDays && durationDays > 1) {
    endDate = addWeekdays(startDate, durationDays - 1);
  }

  const periods = inferPeriodsFromText(text) || (mode === 'leave' ? '7' : '');
  const leaveType = mode === 'leave' ? inferLeaveTypeFromText(text) : (mode === 'late' ? '遲到' : '早退');
  const meal = mode === 'leave' ? (inferMealChoiceFromText(text) || '否') : '';
  const note = text;
  const days = mode === 'leave'
    ? calculateDays(startDate, endDate, periods)
    : (countWeekdaysInclusive(startDate, endDate) || (startDate ? '1' : ''));

  return students.map((student) => ({
    mode,
    student,
    startDate,
    endDate,
    periods,
    days,
    leaveType,
    meal,
    mealReturned: false,
    note
  }));
}

function getSelectedMode() {
  return document.querySelector('input[name="absence-mode"]:checked')?.value || 'leave';
}

function setSelectedMode(mode) {
  document.querySelectorAll('input[name="absence-mode"]').forEach((input) => {
    input.checked = input.value === mode;
  });
}

function setLeaveType(type) {
  selectedLeaveType = type;
  document.querySelectorAll('#absence-leave-type-buttons .absence-pill').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.type === type);
  });
}

function setMealChoice(choice) {
  selectedMeal = choice;
  document.querySelectorAll('#absence-meal-buttons .absence-pill').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.meal === choice);
  });
  $('absence-meal-returned-row').classList.toggle('hidden', choice !== '是');
  if (choice !== '是') $('absence-meal-returned').checked = false;
}

function renderLeaveTypeButtons() {
  const container = $('absence-leave-type-buttons');
  container.innerHTML = '';
  LEAVE_TYPES.forEach((type) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'absence-pill';
    btn.dataset.type = type;
    btn.textContent = type;
    btn.addEventListener('click', () => setLeaveType(type));
    container.appendChild(btn);
  });
  setLeaveType(selectedLeaveType);
}

function ensureParsePreviewUi() {
  if ($('absence-parse-preview-wrap')) return;
  const parseBox = $('absence-parse-text');
  if (!parseBox || !parseBox.parentElement) return;

  const wrap = document.createElement('div');
  wrap.id = 'absence-parse-preview-wrap';
  wrap.className = 'hidden';
  wrap.style.marginTop = '8px';
  wrap.innerHTML = `
    <div id="absence-parse-preview-empty" class="absence-empty">尚未產生預覽。</div>
    <div id="absence-parse-preview-list" style="display:flex;flex-direction:column;gap:6px;"></div>
  `;
  parseBox.parentElement.appendChild(wrap);
}

function ensureMicButtonUi() {
  if ($('btn-absence-mic')) return;
  const toolbar = $('btn-absence-parse')?.parentElement;
  if (!toolbar) return;
  const btn = document.createElement('button');
  btn.id = 'btn-absence-mic';
  btn.type = 'button';
  btn.className = 'btn-small';
  btn.textContent = '🎤 語音';
  btn.title = '開始語音輸入';
  toolbar.insertBefore(btn, $('absence-parse-status'));
}

function renderParsePreview() {
  ensureParsePreviewUi();
  const wrap = $('absence-parse-preview-wrap');
  const empty = $('absence-parse-preview-empty');
  const list = $('absence-parse-preview-list');

  if (!wrap || !empty || !list) return;

  list.innerHTML = '';
  if (!parsedPreviewRecords.length) {
    wrap.classList.add('hidden');
    return;
  }

  wrap.classList.remove('hidden');
  empty.classList.add('hidden');

  parsedPreviewRecords.forEach((record, index) => {
    const item = document.createElement('label');
    item.className = 'cr-student-checkbox';
    item.style.padding = '8px 10px';
    item.style.border = '1px solid #d8dee9';
    item.style.borderRadius = '10px';
    item.style.display = 'flex';
    item.style.alignItems = 'flex-start';
    item.style.gap = '8px';
    item.innerHTML = `
      <input type="checkbox" class="absence-preview-student" value="${index}" checked>
      <span>${escapeHtml(record.student)}｜${escapeHtml(formatDateRange(record.startDate, record.endDate))}｜${escapeHtml(record.leaveType)}</span>
    `;
    list.appendChild(item);
  });
}

function getSelectedPreviewRecords() {
  const indexes = Array.from(document.querySelectorAll('.absence-preview-student:checked')).map((input) => parseInt(input.value, 10));
  return indexes
    .filter((index) => Number.isInteger(index) && parsedPreviewRecords[index])
    .map((index) => parsedPreviewRecords[index]);
}

function renderAbsenceRecords() {
  const empty = $('absence-records-empty');
  const table = $('absence-records-table');
  const tbody = $('absence-records-tbody');
  tbody.innerHTML = '';

  if (!absenceRecords.length) {
    empty.classList.remove('hidden');
    table.classList.add('hidden');
    return;
  }

  empty.classList.add('hidden');
  table.classList.remove('hidden');

  absenceRecords
    .slice()
    .sort((a, b) => String(b.startDate || '').localeCompare(String(a.startDate || '')) || b.rowIndex - a.rowIndex)
    .forEach((record) => {
      const tr = document.createElement('tr');
      const notePreview = record.note && record.note.length > 24 ? `${record.note.substring(0, 24)}…` : (record.note || '');
      tr.innerHTML = `
        <td>${escapeHtml(formatDateRange(record.startDate, record.endDate))}</td>
        <td>${escapeHtml(record.student)}</td>
        <td>${escapeHtml(record.leaveType)}</td>
        <td>${escapeHtml(record.days || '')}</td>
        <td>${escapeHtml(notePreview)}</td>
        <td></td>
      `;

      const actions = document.createElement('div');
      actions.style.display = 'flex';
      actions.style.gap = '6px';

      const editBtn = document.createElement('button');
      editBtn.className = 'btn-small';
      editBtn.textContent = '編輯';
      editBtn.addEventListener('click', () => openAbsenceEdit(record));

      const deleteBtn = document.createElement('button');
      deleteBtn.className = 'btn-small';
      deleteBtn.textContent = '刪除';
      deleteBtn.addEventListener('click', () => deleteAbsenceRecord(record));

      actions.appendChild(editBtn);
      actions.appendChild(deleteBtn);
      tr.lastElementChild.appendChild(actions);
      tbody.appendChild(tr);
    });
}

function renderAbsenceSummary() {
  const empty = $('absence-summary-empty');
  const table = $('absence-summary-table');
  const tbody = $('absence-summary-tbody');
  tbody.innerHTML = '';

  const labels = studentList.map((student) => buildStudentLabel(student));
  const uniqueLabels = Array.from(new Set([
    ...labels,
    ...absenceRecords.map((record) => record.student).filter(Boolean)
  ]));

  const summary = new Map();
  uniqueLabels.forEach((label) => {
    summary.set(label, { '事假': 0, '病假': 0, '公假': 0, '喪假': 0, '遲到早退': 0 });
  });

  absenceRecords.forEach((record) => {
    const targetLabels = record.student === '全班'
      ? (labels.length ? labels : ['全班'])
      : [record.student];

    targetLabels.forEach((label) => {
      if (!summary.has(label)) {
        summary.set(label, { '事假': 0, '病假': 0, '公假': 0, '喪假': 0, '遲到早退': 0 });
      }
      const bucket = summary.get(label);
      const days = parseFloat(record.days) || 0;
      if (SUMMARY_TYPES.includes(record.leaveType)) {
        bucket[record.leaveType] += days;
      } else if (record.leaveType === '遲到' || record.leaveType === '早退' || record.leaveType === '遲到早退') {
        bucket['遲到早退'] += 1;
      }
    });
  });

  const rows = Array.from(summary.entries()).filter(([, totals]) => Object.values(totals).some((value) => value > 0));
  if (!rows.length) {
    empty.classList.remove('hidden');
    table.classList.add('hidden');
    return;
  }

  empty.classList.add('hidden');
  table.classList.remove('hidden');

  rows.forEach(([label, totals]) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${escapeHtml(label)}</td>
      <td>${totals['事假'] || ''}</td>
      <td>${totals['病假'] || ''}</td>
      <td>${totals['公假'] || ''}</td>
      <td>${totals['喪假'] || ''}</td>
      <td>${totals['遲到早退'] || ''}</td>
    `;
    tbody.appendChild(tr);
  });
}

function updateDaysDisplay() {
  const startDate = $('absence-start-date').value;
  const endDate = $('absence-end-date').value || startDate;
  const periodsInput = $('absence-periods');
  const mode = getSelectedMode();
  const isMultiDay = !!startDate && !!endDate && startDate !== endDate;

  if (mode !== 'leave') {
    $('absence-days').value = countWeekdaysInclusive(startDate, endDate) || (startDate ? '1' : '');
    periodsInput.disabled = true;
    return;
  }

  periodsInput.disabled = isMultiDay;
  if (isMultiDay) periodsInput.value = '7';
  $('absence-days').value = calculateDays(startDate, endDate, periodsInput.value);
}

function updateAbsenceModeUi() {
  const mode = getSelectedMode();
  const isLeave = mode === 'leave';
  const isLate = mode === 'late';

  $('absence-end-date-row').classList.remove('hidden');
  $('absence-periods-row').classList.toggle('hidden', !isLeave);
  $('absence-days-row').classList.remove('hidden');
  $('absence-meal-section').classList.toggle('hidden', !isLeave);
  $('absence-leave-type-buttons').classList.toggle('hidden', !isLeave);
  $('absence-late-type-display').classList.toggle('hidden', isLeave);

  const noteEl = $('absence-note');
  if (!isLeave) {
    if (!$('absence-end-date').value) $('absence-end-date').value = $('absence-start-date').value;
    if (!noteEl.value || noteEl.value === TARDY_NOTE_TEMPLATE || noteEl.value === EARLY_LEAVE_NOTE_TEMPLATE) {
      noteEl.value = isLate ? TARDY_NOTE_TEMPLATE : EARLY_LEAVE_NOTE_TEMPLATE;
    }
    noteEl.placeholder = isLate ? TARDY_NOTE_TEMPLATE : EARLY_LEAVE_NOTE_TEMPLATE;
    $('absence-late-type-display').textContent = isLate ? '遲到' : '早退';
  } else {
    if (noteEl.value === TARDY_NOTE_TEMPLATE || noteEl.value === EARLY_LEAVE_NOTE_TEMPLATE) noteEl.value = '';
    noteEl.placeholder = '備註';
  }

  updateDaysDisplay();
}

function resetAbsenceForm() {
  absenceEditRowIndex = null;
  parsedPreviewRecords = [];
  setSelectedMode('leave');
  $('absence-student').value = '全班';
  $('absence-start-date').value = todayIso();
  $('absence-end-date').value = todayIso();
  $('absence-periods').value = '7';
  $('absence-days').value = '1';
  $('absence-parse-text').value = '';
  $('absence-parse-status').textContent = '';
  $('absence-note').value = '';
  $('absence-meal-returned').checked = false;
  $('absence-meal-note-card').classList.add('hidden');
  $('absence-parse-preview-wrap')?.classList.add('hidden');
  document.querySelector('[data-target="absence-meal-note-card"]')?.setAttribute('aria-expanded', 'false');
  setLeaveType('事假');
  setMealChoice('否');
  renderParsePreview();
  updateAbsenceModeUi();
}

function openAbsenceNew() {
  resetAbsenceForm();
  $('absence-list-view').style.display = 'none';
  $('absence-form').classList.remove('hidden');
  $('btn-absence-back').classList.remove('hidden');
}

function openAbsenceEdit(record) {
  absenceEditRowIndex = record.rowIndex;
  parsedPreviewRecords = [];
  $('absence-student').value = record.student;
  $('absence-start-date').value = record.startDate || todayIso();
  $('absence-end-date').value = record.endDate || record.startDate || todayIso();
  $('absence-periods').value = record.periods || '7';
  $('absence-days').value = record.days || '';
  $('absence-parse-status').textContent = '';
  $('absence-note').value = record.note || '';
  $('absence-meal-returned').checked = !!record.mealReturned;
  $('absence-parse-preview-wrap')?.classList.add('hidden');
  setMealChoice(record.meal === '是' ? '是' : '否');

  if (record.leaveType === '遲到') {
    setSelectedMode('late');
  } else if (record.leaveType === '早退') {
    setSelectedMode('early');
  } else {
    setSelectedMode('leave');
    setLeaveType(record.leaveType || '事假');
  }

  updateAbsenceModeUi();
  $('absence-list-view').style.display = 'none';
  $('absence-form').classList.remove('hidden');
  $('btn-absence-back').classList.remove('hidden');
}

async function deleteAbsenceRecord(record) {
  const gasUrl = appData.scheduleGasUrl || appData.gasApiUrl;
  if (!gasUrl) {
    alert('請先設定 GAS API 網址');
    return;
  }
  if (!confirm('確定要刪除這筆出缺席紀錄嗎？')) return;
  const result = await ipcRenderer.invoke('delete-absence-record', { gasUrl, rowIndex: record.rowIndex });
  if (!result.ok) {
    alert('刪除失敗：' + result.error);
    return;
  }
  await loadAbsenceRecords();
}

function validateAbsenceRecord(record) {
  if (!record.student) return '請選擇座號姓名';
  if (!record.startDate) return '請選擇開始日期';
  if (!record.endDate) return '請選擇結束日期';
  if (record.mode === 'leave' && !record.leaveType) return '請選擇假別';
  if (record.mode !== 'leave' && !record.note) return '請填寫備註';
  return '';
}

function collectAbsenceRecord() {
  const mode = getSelectedMode();
  const startDate = $('absence-start-date').value;
  const endDate = $('absence-end-date').value || startDate;
  const periods = mode === 'leave' ? $('absence-periods').value : '';
  const days = mode === 'leave' ? $('absence-days').value : (countWeekdaysInclusive(startDate, endDate) || (startDate ? '1' : ''));

  return {
    mode,
    student: $('absence-student').value,
    startDate,
    endDate,
    periods,
    days,
    leaveType: mode === 'leave' ? selectedLeaveType : (mode === 'late' ? '遲到' : '早退'),
    meal: mode === 'leave' ? selectedMeal : '',
    mealReturned: mode === 'leave' ? $('absence-meal-returned').checked : false,
    note: $('absence-note').value.trim()
  };
}

function applyParsedAbsenceText() {
  const text = $('absence-parse-text').value.trim();
  const status = $('absence-parse-status');
  if (!text) {
    status.textContent = '請先輸入文字';
    status.style.color = '#e53e3e';
    return;
  }

  const parsedRecords = buildParsedRecords(text);
  if (!parsedRecords.length) {
    parsedPreviewRecords = [];
    renderParsePreview();
    status.textContent = '沒有辨識到學生，請再檢查文字內容。';
    status.style.color = '#e53e3e';
    return;
  }

  parsedPreviewRecords = parsedRecords;
  const first = parsedRecords[0];

  setSelectedMode(first.mode);
  $('absence-student').value = first.student;
  $('absence-start-date').value = first.startDate;
  $('absence-end-date').value = first.endDate;
  if (first.mode === 'leave') {
    $('absence-periods').value = first.periods || '7';
    setLeaveType(first.leaveType);
    setMealChoice(first.meal || '否');
  }

  const noteEl = $('absence-note');
  if (!noteEl.value || noteEl.value === TARDY_NOTE_TEMPLATE || noteEl.value === EARLY_LEAVE_NOTE_TEMPLATE) {
    noteEl.value = first.note;
  }

  renderParsePreview();
  updateAbsenceModeUi();
  updateDaysDisplay();
  status.textContent = `已帶入資料，並產生 ${parsedPreviewRecords.length} 筆預覽；請先確認清單。`;
  status.style.color = '#38a169';
}

function ensureSpeechRecognition() {
  const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!Recognition) return null;
  if (speechRecognition) return speechRecognition;

  speechRecognition = new Recognition();
  speechRecognition.lang = 'zh-TW';
  speechRecognition.continuous = true;
  speechRecognition.interimResults = true;

  speechRecognition.onresult = (event) => {
    let transcript = '';
    for (let i = event.resultIndex; i < event.results.length; i += 1) {
      transcript += event.results[i][0].transcript;
    }
    $('absence-parse-text').value = transcript.trim();
  };

  speechRecognition.onend = () => {
    speechListening = false;
    updateMicButtonState();
  };

  speechRecognition.onerror = () => {
    speechListening = false;
    updateMicButtonState();
  };

  return speechRecognition;
}

function updateMicButtonState() {
  const btn = $('btn-absence-mic');
  if (!btn) return;
  btn.textContent = speechListening ? '⏹ 停止' : '🎤 語音';
  btn.title = speechListening ? '停止語音輸入' : '開始語音輸入';
}

function toggleSpeechInput() {
  const recognition = ensureSpeechRecognition();
  if (!recognition) {
    alert('這台裝置目前不支援語音辨識。');
    return;
  }

  if (speechListening) {
    recognition.stop();
    speechListening = false;
    updateMicButtonState();
    return;
  }

  $('absence-parse-status').textContent = '語音辨識中...';
  $('absence-parse-status').style.color = '#2563eb';
  recognition.start();
  speechListening = true;
  updateMicButtonState();
}

async function saveAbsenceRecord() {
  const gasUrl = appData.scheduleGasUrl || appData.gasApiUrl;
  if (!gasUrl) {
    alert('請先設定 GAS API 網址');
    return;
  }

  const formRecord = collectAbsenceRecord();
  const error = validateAbsenceRecord(formRecord);
  if (error) {
    alert(error);
    return;
  }

  const previewRecords = getSelectedPreviewRecords();
  const targets = absenceEditRowIndex
    ? [formRecord]
    : (previewRecords.length ? previewRecords : [formRecord]);

  const btn = $('btn-save-absence');
  btn.disabled = true;
  btn.textContent = absenceEditRowIndex ? '更新中...' : '儲存中...';

  let result = null;
  if (absenceEditRowIndex) {
    result = await ipcRenderer.invoke('update-absence-record', {
      gasUrl,
      rowIndex: absenceEditRowIndex,
      record: formRecord
    });
  } else {
    for (const record of targets) {
      result = await ipcRenderer.invoke('save-absence-record', { gasUrl, record });
      if (!result.ok) break;
    }
  }

  btn.disabled = false;
  btn.textContent = absenceEditRowIndex ? '更新' : '儲存';

  if (!result || !result.ok) {
    alert('儲存失敗：' + (result?.error || '未知錯誤'));
    return;
  }

  await loadAbsenceRecords();
  $('absence-list-view').style.display = '';
  $('absence-form').classList.add('hidden');
  $('btn-absence-back').classList.add('hidden');
}

async function loadStudentList() {
  const gasUrl = appData.scheduleGasUrl || appData.gasApiUrl;
  if (!gasUrl) return;
  const result = await ipcRenderer.invoke('fetch-student-list', gasUrl);
  if (!result.ok) return;

  studentList = result.students || [];
  const select = $('absence-student');
  select.innerHTML = '';
  [{ num: '', name: '全班' }, ...studentList].forEach((student) => {
    const option = document.createElement('option');
    option.value = student.name === '全班' ? '全班' : buildStudentLabel(student);
    option.textContent = option.value;
    select.appendChild(option);
  });
}

async function init() {
  appData = await ipcRenderer.invoke('load-data');
  ensureParsePreviewUi();
  ensureMicButtonUi();
  renderLeaveTypeButtons();
  await loadStudentList();
  await loadAbsenceRecords();
  updateMicButtonState();
}

$('btn-absence-new').addEventListener('click', openAbsenceNew);
$('btn-absence-back').addEventListener('click', () => {
  $('absence-list-view').style.display = '';
  $('absence-form').classList.add('hidden');
  $('btn-absence-back').classList.add('hidden');
});
$('btn-close-absence').addEventListener('click', () => {
  ipcRenderer.invoke('close-absence-window');
});
$('btn-absence-parse').addEventListener('click', applyParsedAbsenceText);
$('absence-student').addEventListener('change', () => {
  parsedPreviewRecords = [];
  renderParsePreview();
});
$('absence-start-date').addEventListener('change', () => {
  const startDate = $('absence-start-date').value;
  const endDateEl = $('absence-end-date');
  endDateEl.min = startDate;
  if (!endDateEl.value || endDateEl.value < startDate) endDateEl.value = startDate;
  updateAbsenceModeUi();
});
$('absence-end-date').addEventListener('change', updateDaysDisplay);
$('absence-periods').addEventListener('input', updateDaysDisplay);
$('btn-save-absence').addEventListener('click', saveAbsenceRecord);

document.addEventListener('click', (event) => {
  if (event.target && event.target.id === 'btn-absence-mic') {
    toggleSpeechInput();
  }
});

document.querySelectorAll('input[name="absence-mode"]').forEach((input) => {
  input.addEventListener('change', updateAbsenceModeUi);
});
document.querySelectorAll('#absence-meal-buttons .absence-pill').forEach((btn) => {
  btn.addEventListener('click', () => setMealChoice(btn.dataset.meal));
});

init();
