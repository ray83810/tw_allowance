/* ============================================================
   Asurion 排班分析工具 - 核心邏輯
   ============================================================ */

// ==================== 全域狀態 ====================
const state = {
  files: { application: null, schedule: null },
  workbooks: { application: null, schedule: null },
  parsed: { leave: [], overtime: [], origLeaves: [], origOvertimes: [], employees: [], ptoData: [], holidays: [], scheduleMonth: null, scheduleYear: null, scheduleSheet: null },
  results: null
};

// ==================== 假別對照表 ====================
const LEAVE_TYPE_MAP = {
  'Annual Leave特別休假': { code: 'PTO', label: 'Annual Leave\n特別休假', shortLabel: '特休', category: 'pto' },
  'Asurion Leave亞勝假期': { code: 'PTO-AL', label: 'Asurion Leave\n亞勝假期', shortLabel: '亞勝假期', category: 'pto' },
  'Sick Leave病假': { code: 'SL', label: 'Sick Leave\n病假', shortLabel: '病假', category: 'other' },
  'Menstrual Leave病假（生理假)': { code: 'SL-M', label: 'Menstrual Leave\n病假（生理假)', shortLabel: '生理假', category: 'other' },
  'Personal Leave事假': { code: 'PL', label: 'Personal Leave事假', shortLabel: '事假', category: 'other' },
  'Official Leave公假': { code: 'LOA', label: 'Official Leave公假', shortLabel: '公假', category: 'other' },
  'Official Leave公假（金融市場常識與職業道德考試）': { code: 'LOA', label: 'Official Leave公假', shortLabel: '公假', category: 'other' },
  'Official Leave公假（財產保險業務員資格證照考試)': { code: 'LOA', label: 'Official Leave公假', shortLabel: '公假', category: 'other' },
  'Official Leave公假（健檢)': { code: 'LOA', label: 'Official Leave公假', shortLabel: '公假', category: 'other' },
  'Marriage Leave婚假': { code: 'ML', label: 'Marriage Leave婚假', shortLabel: '婚假', category: 'other' },
  'Bereavement Leave喪假': { code: 'BL', label: 'Bereavement Leave喪假', shortLabel: '喪假', category: 'other' },
  'Family Care Leave家庭照顧假': { code: 'FL', label: 'Family Care Leave家庭照顧假', shortLabel: '家庭假', category: 'other' }
};

const LEAVE_CODE_ORDER = ['PTO', 'PTO-AL', 'SL', 'SL-M', 'PL', 'LOA', 'ML', 'BL', 'FL'];
const OTHER_LEAVE_CODES = ['SL', 'SL-M', 'PL', 'LOA', 'ML', 'BL', 'FL'];
const MONTH_NAMES_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MONTH_NAMES_ZH = ['一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'];

// ==================== 工具函數 ====================
function parseExcelDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    if (val.__isNormalized) return val;
    const localDate = new Date(val.getUTCFullYear(), val.getUTCMonth(), val.getUTCDate(), val.getUTCHours(), val.getUTCMinutes(), val.getUTCSeconds());
    localDate.__isNormalized = true;
    return localDate;
  }
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    if (d) {
      const localDate = new Date(d.y, d.m - 1, d.d, d.H, d.M, d.S);
      localDate.__isNormalized = true;
      return localDate;
    }
  }
  if (typeof val === 'string') {
    const cleanStr = val.trim().replace(/[\/\\]/g, '-');
    const match = cleanStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (match) {
      const y = parseInt(match[1], 10);
      const m = parseInt(match[2], 10) - 1;
      const d = parseInt(match[3], 10);
      const timeMatch = cleanStr.match(/\s+(\d{1,2}):(\d{1,2}):(\d{1,2})/);
      let localDate;
      if (timeMatch) {
        localDate = new Date(y, m, d, parseInt(timeMatch[1], 10), parseInt(timeMatch[2], 10), parseInt(timeMatch[3], 10));
      } else {
        localDate = new Date(y, m, d);
      }
      localDate.__isNormalized = true;
      return localDate;
    }
    const d = new Date(val);
    if (!isNaN(d.getTime())) {
      const localDate = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds());
      localDate.__isNormalized = true;
      return localDate;
    }
  }
  return null;
}

function toNum(val) {
  if (val === null || val === undefined || val === '') return 0;
  const n = parseFloat(val);
  return isNaN(n) ? 0 : n;
}

function normalizeName(name) {
  if (!name) return '';
  return String(name).trim().replace(/\s+/g, ' ');
}

function getCanonicalName(name, activeEmployeesList) {
  if (!name) return '';
  const norm = name.trim().toLowerCase().replace(/[\s\-_]+/g, '');
  
  const employeesList = activeEmployeesList || (state.parsed.employees && state.parsed.employees.map(e => e.name)) || [];
  
  // Clean active list names for matching
  for (const emp of employeesList) {
    const empNorm = emp.trim().toLowerCase().replace(/[\s\-_]+/g, '');
    if (norm === empNorm) return emp;
  }
  
  // Try matching last name and first name overlap
  for (const emp of employeesList) {
    const empParts = emp.trim().split(/\s+/);
    const nameParts = name.trim().split(/\s+/);
    
    const empLast = empParts[empParts.length - 1].toLowerCase();
    const nameLast = nameParts[nameParts.length - 1].toLowerCase();
    
    if (empLast === nameLast) {
      const empFirst = empParts.slice(0, -1).join('').toLowerCase();
      const nameFirst = nameParts.slice(0, -1).join('').toLowerCase();
      if (empFirst.includes(nameFirst) || nameFirst.includes(empFirst) || norm.includes(empFirst)) {
        return emp;
      }
    }
  }

  // Fallback to name parts overlap
  for (const emp of employeesList) {
    const empParts = emp.toLowerCase().split(/\s+/);
    const nameParts = name.toLowerCase().trim().split(/\s+/);
    if (nameParts.some(p => empParts.includes(p))) {
      return emp;
    }
  }
  
  // Custom check for the Ding Kai / Jian Kai Ding / Jlian Kai Ding typo
  for (const emp of employeesList) {
    const empNorm = emp.toLowerCase().replace(/[\s\-_]+/g, '');
    if (empNorm.includes('ding') && norm.includes('ding')) {
      if (norm.includes('kai') || norm.includes('jian') || norm.includes('jlian')) {
        return emp;
      }
    }
  }

  return name.trim();
}

function getMonthKey(date) {
  if (!date) return null;
  const d = parseExcelDate(date);
  if (!d) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

function formatLocalDate(date) {
  if (!date) return '';
  const d = parseExcelDate(date);
  if (!d) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dateVal = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dateVal}`;
}

function getColLetter(colIndex) {
  let temp, letter = '';
  while (colIndex > 0) {
    temp = (colIndex - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    colIndex = (colIndex - temp - 1) / 26;
  }
  return letter;
}

function makeYearlySumFormula(colLetter, rowIdx) {
  const parts = MONTH_NAMES_EN.map((m, idx) => {
    const sheetName = m + (idx === 0 ? ' ' : '');
    return `'${sheetName}'!${colLetter}${rowIdx}`;
  }).join(',');
  return `SUM(${parts})`;
}

function showToast(message, type = 'success') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  container.appendChild(toast);
  setTimeout(() => { toast.classList.add('exit'); setTimeout(() => toast.remove(), 400); }, 3000);
}

// ==================== 假別字串規整 ====================
function normalizeLeaveType(type) {
  if (!type) return '';
  type = String(type).trim();
  if (type.includes('特別休假') || type.includes('Annual Leave') || (type.includes('PTO') && !type.includes('PTO-AL'))) {
    return 'Annual Leave特別休假';
  }
  if (type.includes('亞勝') || type.includes('Asurion Leave') || type.includes('PTO-AL')) {
    return 'Asurion Leave亞勝假期';
  }
  if (type.includes('生理假') || type.includes('Menstrual')) {
    return 'Menstrual Leave病假（生理假)';
  }
  if (type.includes('病假') || type.includes('Sick')) {
    return 'Sick Leave病假';
  }
  if (type.includes('事假') || type.includes('Personal')) {
    return 'Personal Leave事假';
  }
  if (type.includes('金融市場')) {
    return 'Official Leave公假（金融市場常識與職業道德考試）';
  }
  if (type.includes('財產保險')) {
    return 'Official Leave公假（財產保險業務員資格證照考試)';
  }
  if (type.includes('健檢')) {
    return 'Official Leave公假（健檢)';
  }
  if (type.includes('公假') || type.includes('Official')) {
    return 'Official Leave公假';
  }
  if (type.includes('婚假') || type.includes('Marriage')) {
    return 'Marriage Leave婚假';
  }
  if (type.includes('喪假') || type.includes('Bereavement')) {
    return 'Bereavement Leave喪假';
  }
  if (type.includes('家庭') || type.includes('Family')) {
    return 'Family Care Leave家庭照顧假';
  }
  return type;
}

function isNightShiftVal(val) {
  if (!val) return false;
  const s = String(val).trim();
  return /^(15|16|17|18|19)\b/.test(s);
}

function formatShiftDate(date) {
  if (!date) return '';
  const d = parseExcelDate(date);
  if (!d) return '';
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

// ==================== 檔案偵測 ====================
function detectFileType(wb) {
  const sheets = wb.SheetNames;
  if (sheets.some(s => /\d{6}/.test(s)) || sheets.includes('特休日數') || sheets.includes('特休') || sheets.includes('國定假日') || sheets.includes('排班')) {
    return 'schedule';
  }
  if (sheets.includes('Sheet1') || sheets.includes('班表相關申請表單') || sheets.includes('請假原始檔') || sheets.includes('加班原始數據')) {
    return 'application';
  }
  return null;
}

function parseLeaveSheet(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (data.length < 2) return { daily: [], original: [] };
  const daily = [];
  const original = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[5]) continue;
    const applicant = normalizeName(row[5]);
    const leaveType = normalizeLeaveType(row[6]);
    const startDate = parseExcelDate(row[8]);
    const days = toNum(row[10]);
    const timeRange = String(row[11] || '');
    if (!applicant || !leaveType || !startDate) continue;

    const endDate = parseExcelDate(row[9]) || startDate;
    original.push({ applicant, leaveType, startDate, endDate, days, timeRange });

    let remainingDays = days;
    let currentDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    currentDate.__isNormalized = true;

    while (remainingDays > 0) {
      const dayVal = Math.min(remainingDays, 1.0);
      const mDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
      mDate.__isNormalized = true;

      daily.push({
        applicant,
        leaveType,
        monthDate: mDate,
        startDate: new Date(currentDate.getTime()),
        endDate: new Date(currentDate.getTime()),
        days: dayVal,
        timeRange,
        monthKey: getMonthKey(mDate)
      });

      remainingDays -= dayVal;
      currentDate.setDate(currentDate.getDate() + 1);
    }
  }
  return { daily, original };
}

function parseOvertimeSheet(ws) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  if (data.length < 2) return [];
  const records = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[5]) continue;
    const applicant = normalizeName(row[5]);
    const otDate = parseExcelDate(row[7]);
    const hours = toNum(row[11]);
    const approved = String(row[15] || '').trim();
    if (!applicant || !otDate) continue;
    if (approved && approved !== 'Approved') continue;

    const startTime = String(row[8] || '').trim();
    const endDate = parseExcelDate(row[9]) || otDate;
    const endTime = String(row[10] || '').trim();

    records.push({ applicant, otDate, startTime, endDate, endTime, hours, dateKey: getMonthKey(otDate) });
  }
  return records;
}

function parseCombinedSheet(data) {
  const dailyLeaves = [];
  const origLeaves = [];
  const overtime = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[5]) continue;
    
    const applicant = normalizeName(row[5]);
    const formType = String(row[6] || '').trim();
    if (!applicant || !formType) continue;

    // 加班申請
    if (formType.startsWith('加班') || formType === '加班申請表') {
      const otDate = parseExcelDate(row[14]); // O (index 14)
      const hours = toNum(row[18]); // S (index 18)
      if (otDate && hours > 0) {
        const startTime = String(row[15] || '').trim();
        const endDate = parseExcelDate(row[16]) || otDate;
        const endTime = String(row[17] || '').trim();

        overtime.push({ applicant, otDate, startTime, endDate, endTime, hours, dateKey: getMonthKey(otDate) });
      }
    }
    // 請假申請 (包含長假與公假，排除事前申請)
    else if ((formType.startsWith('請假') || formType.startsWith('長假') || formType.startsWith('公假')) && !formType.includes('(事前申請)')) {
      let leaveType = '';
      let startDate = null;
      let endDate = null;
      let days = 0;
      let timeRange = '';

      if (formType.startsWith('請假')) {
        leaveType = normalizeLeaveType(row[7]); // H (index 7)
        startDate = parseExcelDate(row[8]); // I (index 8)
        endDate = parseExcelDate(row[9]); // J (index 9)
        days = toNum(row[10]); // K (index 10)
        timeRange = String(row[11] || ''); // L (index 11)
      } else if (formType.startsWith('長假')) {
        leaveType = normalizeLeaveType(row[28]); // AC (index 28)
        startDate = parseExcelDate(row[25]); // Z (index 25)
        endDate = parseExcelDate(row[26]); // AA (index 26)
        days = toNum(row[27]); // AB (index 27)
      } else if (formType.startsWith('公假')) {
        leaveType = 'Official Leave公假';
        startDate = parseExcelDate(row[29]); // AD (index 29)
        endDate = startDate;
        days = 1;
      }

      if (leaveType && startDate) {
        endDate = endDate || startDate;
        origLeaves.push({ applicant, leaveType, startDate, endDate, days, timeRange });

        let remainingDays = days;
        let currentDate = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
        currentDate.__isNormalized = true;

        while (remainingDays > 0) {
          const dayVal = Math.min(remainingDays, 1.0);
          const monthDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
          monthDate.__isNormalized = true;

          dailyLeaves.push({
            applicant,
            leaveType,
            monthDate,
            startDate: new Date(currentDate.getTime()),
            endDate: new Date(currentDate.getTime()),
            days: dayVal,
            timeRange,
            monthKey: getMonthKey(monthDate)
          });

          remainingDays -= dayVal;
          currentDate.setDate(currentDate.getDate() + 1);
        }
      }
    }
  }
  return { dailyLeaves, origLeaves, overtime };
}

function parseApplicationData(wb) {
  const result = { leave: [], overtime: [], origLeaves: [], origOvertimes: [] };
  
  // 1. 支援舊版請假原始檔
  if (wb.Sheets['請假原始檔']) {
    const parsed = parseLeaveSheet(wb.Sheets['請假原始檔']);
    result.leave = parsed.daily;
    result.origLeaves = parsed.original;
  }
  // 2. 支援舊版加班原始數據
  if (wb.Sheets['加班原始數據']) {
    result.overtime = parseOvertimeSheet(wb.Sheets['加班原始數據']);
    result.origOvertimes = result.overtime;
  }
  
  // 3. 支援新版合併表單 (Sheet1)
  const combinedSheetName = wb.SheetNames.find(s => s === 'Sheet1' || (!['請假原始檔', '請假整理', '加班原始數據', '加班整理'].includes(s) && !/\d{6}/.test(s)));
  if (combinedSheetName) {
    const ws = wb.Sheets[combinedSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    if (data.length >= 2) {
      const parsed = parseCombinedSheet(data);
      if (parsed.dailyLeaves.length > 0) {
        result.leave = parsed.dailyLeaves;
        result.origLeaves = parsed.origLeaves;
      }
      if (parsed.overtime.length > 0) {
        result.overtime = parsed.overtime;
        result.origOvertimes = parsed.overtime;
      }
    }
  }
  
  return result;
}

// ==================== 解析排班表 ====================
function parseScheduleData(wb) {
  const result = { employees: [], ptoData: [], holidays: [], month: null, year: null, scheduleSheetName: null };

  // 偵測月份分頁 (如 "202605" 或 "202605v2")
  const scheduleSheet = wb.SheetNames.find(s => /\d{6}/.test(s));
  if (scheduleSheet) {
    const match = scheduleSheet.match(/(\d{4})(\d{2})/);
    if (match) {
      result.year = parseInt(match[1]);
      result.month = parseInt(match[2]);
      result.scheduleSheetName = scheduleSheet;
    }
  }

  // 解析月份分頁 - 取得員工列表和班表
  if (scheduleSheet) {
    const ws = wb.Sheets[scheduleSheet];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

    // 取得日期列 (動態尋找第二列中第一個為 2000 年之後的日期欄位)
    const dates = [];
    let dateStartCol = -1;
    if (data.length > 1) {
      for (let c = 0; c < data[1].length; c++) {
        const d = parseExcelDate(data[1][c]);
        if (d && d.getFullYear() >= 2000) {
          dateStartCol = c;
          break;
        }
      }
      
      if (dateStartCol !== -1) {
        for (let c = dateStartCol; c < data[1].length; c++) {
          const d = parseExcelDate(data[1][c]);
          if (d && d.getFullYear() >= 2000) dates.push({ col: c, date: d });
        }
      }
    }

    // 取得員工資料 (rows where col A has a name and col B has shift time)
    if (dateStartCol !== -1) {
      for (let r = 2; r < data.length; r++) {
        const row = data[r];
        const name = normalizeName(row[0]);
        const shift = String(row[1] || '').trim();

        if (!name || !shift.match(/\d{2}:\d{2}/)) continue;

        const offDays = String(row[2] || '');
        const workDays = toNum(row[dateStartCol - 5]);
        const offCount = toNum(row[dateStartCol - 4]);
        const ptoCount = toNum(row[dateStartCol - 3]);
        const ptoAlCount = toNum(row[dateStartCol - 2]);
        const loaCount = toNum(row[dateStartCol - 1]);

        // 取得每日班表
        const schedule = {};
        for (const { col, date } of dates) {
          const val = String(row[col] || '').trim();
          schedule[formatLocalDate(date)] = val;
        }

        result.employees.push({ name, shift, offDays, workDays, offCount, ptoCount, ptoAlCount, loaCount, schedule });
      }
    }
  }

  // 解析 特休日數 分頁 (可能為 "特休日數" 或 "特休")
  const ptoSheetName = wb.SheetNames.find(s => s === '特休日數' || s === '特休');
  if (ptoSheetName) {
    const ws = wb.Sheets[ptoSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const activeNames = result.employees.map(e => e.name);
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const name = normalizeName(row[0]);
      if (!name) continue;

      const canonicalName = getCanonicalName(name, activeNames);

      const pto = toNum(row[1]);
      const ptoAl = toNum(row[2]);
      const ttl = toNum(row[3]);
      const monthly = {};
      for (let m = 0; m < 12; m++) {
        monthly[m + 1] = toNum(row[4 + m]);
      }

      result.ptoData.push({ name: canonicalName, pto, ptoAl, ttl, monthly });
    }
  }

  // 解析 國定假日 分頁 (可能為 "國定假日" 或 "排班")
  const holidaySheetName = wb.SheetNames.find(s => s === '國定假日' || s === '排班');
  if (holidaySheetName) {
    const ws = wb.Sheets[holidaySheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    for (const row of data) {
      if (row && row[0]) {
        const d = parseExcelDate(row[0]);
        if (d && d.getFullYear() >= 2000) result.holidays.push(d);
      }
    }
  }

  return result;
}

// ==================== 計算邏輯 ====================
function calculateAll() {
  const { leave: leaveRecords, overtime: otRecords, employees, ptoData } = state.parsed;
  const { scheduleMonth, scheduleYear } = state.parsed;
  const targetMonthKey = `${scheduleYear}-${String(scheduleMonth).padStart(2, '0')}`;

  // 取得所有員工名 (以排班表 active 名單為主，已去重排序)
  const allEmployees = [...new Set(employees.map(e => e.name))].sort();

  // ===== 1. 當月請假統計 =====
  const monthlyLeave = {};
  const monthlyLeaveTypes = new Set();

  for (const rec of leaveRecords) {
    if (rec.monthKey !== targetMonthKey) continue;

    const empName = getCanonicalName(rec.applicant, allEmployees);
    if (!allEmployees.includes(empName)) continue; // 排除非本月班表名單的人

    monthlyLeaveTypes.add(rec.leaveType);
    if (!monthlyLeave[empName]) monthlyLeave[empName] = {};
    monthlyLeave[empName][rec.leaveType] = (monthlyLeave[empName][rec.leaveType] || 0) + rec.days;
  }

  // ===== 2. 年度請假統計 =====
  const yearlyLeave = {};
  const yearlyLeaveTypes = new Set();

  for (const rec of leaveRecords) {
    const empName = getCanonicalName(rec.applicant, allEmployees);
    if (!allEmployees.includes(empName)) continue;

    const year = parseInt(rec.monthKey.split('-')[0]);
    if (year !== scheduleYear) continue;

    const typeInfo = LEAVE_TYPE_MAP[rec.leaveType];
    if (!typeInfo) continue;

    yearlyLeaveTypes.add(rec.leaveType);
    const mk = rec.monthKey;
    if (!yearlyLeave[mk]) yearlyLeave[mk] = {};
    yearlyLeave[mk][rec.leaveType] = (yearlyLeave[mk][rec.leaveType] || 0) + rec.days;
  }

  // ===== 3. 加班統計 =====
  const otStats = {};
  const otDates = new Set();

  for (const rec of otRecords) {
    const empName = getCanonicalName(rec.applicant, allEmployees);
    if (!allEmployees.includes(empName)) continue;

    if (rec.dateKey !== targetMonthKey) continue;

    const dateStr = formatLocalDate(rec.otDate);
    otDates.add(dateStr);
    if (!otStats[empName]) otStats[empName] = {};
    otStats[empName][dateStr] = (otStats[empName][dateStr] || 0) + rec.hours;
  }

  // ===== 4. 總請假統計 =====
  const totalStats = {};

  for (const emp of allEmployees) {
    totalStats[emp] = {
      monthlyPTO: {},
      otherLeave: {}
    };
    for (let m = 1; m <= 12; m++) totalStats[emp].monthlyPTO[m] = 0;
    for (const code of OTHER_LEAVE_CODES) totalStats[emp].otherLeave[code] = 0;
  }

  for (const rec of leaveRecords) {
    const empName = getCanonicalName(rec.applicant, allEmployees);
    if (!allEmployees.includes(empName)) continue;

    const year = parseInt(rec.monthKey.split('-')[0]);
    if (year !== scheduleYear) continue;

    const typeInfo = LEAVE_TYPE_MAP[rec.leaveType];
    if (!typeInfo || !totalStats[empName]) continue;

    const monthNum = parseInt(rec.monthKey.split('-')[1]);

    if (typeInfo.code === 'PTO' || typeInfo.code === 'PTO-AL') {
      totalStats[empName].monthlyPTO[monthNum] = (totalStats[empName].monthlyPTO[monthNum] || 0) + rec.days;
    } else {
      totalStats[empName].otherLeave[typeInfo.code] = (totalStats[empName].otherLeave[typeInfo.code] || 0) + rec.days;
    }
  }

  // ===== 5. 月度明細統計 (for 總請假統計 各月分頁) =====
  const monthlyDetail = {};
  for (let m = 1; m <= 12; m++) {
    monthlyDetail[m] = {};
    for (const emp of allEmployees) {
      monthlyDetail[m][emp] = {};
      for (const code of LEAVE_CODE_ORDER) monthlyDetail[m][emp][code] = 0;
    }
  }

  for (const rec of leaveRecords) {
    const empName = getCanonicalName(rec.applicant, allEmployees);
    if (!allEmployees.includes(empName)) continue;

    const year = parseInt(rec.monthKey.split('-')[0]);
    if (year !== scheduleYear) continue;

    const typeInfo = LEAVE_TYPE_MAP[rec.leaveType];
    if (!typeInfo) continue;

    const monthNum = parseInt(rec.monthKey.split('-')[1]);
    if (monthlyDetail[monthNum] && monthlyDetail[monthNum][empName]) {
      monthlyDetail[monthNum][empName][typeInfo.code] = (monthlyDetail[monthNum][empName][typeInfo.code] || 0) + rec.days;
    }
  }

  // ===== 6. 特休日數計算 =====
  const ptoSummary = {};
  for (const emp of allEmployees) {
    const p = ptoData.find(pd => pd.name === emp) || { pto: 0, ptoAl: 0 };
    const monthlyUsage = {};
    let totalUsed = 0;
    for (let m = 1; m <= 12; m++) {
      const usage = totalStats[emp] ? totalStats[emp].monthlyPTO[m] : 0;
      monthlyUsage[m] = usage;
      totalUsed += usage;
    }

    ptoSummary[emp] = {
      pto: p.pto,
      ptoAl: p.ptoAl,
      ttl: p.pto + p.ptoAl,
      monthlyUsage,
      totalUsed,
      ptoRemaining: p.pto - totalUsed,
      ptoAlRemaining: p.ptoAl,
      totalRemaining: p.pto + p.ptoAl - totalUsed
    };
  }

  // ===== 7. 津貼統計 (小夜班) =====
  const allowanceStats = {
    summary: [], // { name, days }
    details: [], // { name, startDate, endDate, total }
  };

  const nightShiftSummary = [];
  const nightShiftDetails = [];

  for (const emp of allEmployees) {
    const employee = employees.find(e => e.name === emp);
    if (!employee || !employee.schedule) continue;

    const nightShiftDates = [];
    for (const dateStr of Object.keys(employee.schedule)) {
      if (isNightShiftVal(employee.schedule[dateStr])) {
        const parts = dateStr.split('-');
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const d = parseInt(parts[2], 10);
        const dateObj = new Date(y, m, d);
        dateObj.__isNormalized = true;
        nightShiftDates.push(dateObj);
      }
    }

    if (nightShiftDates.length === 0) continue;

    // Sort chronologically
    nightShiftDates.sort((a, b) => a.getTime() - b.getTime());

    // Group consecutive dates
    const groups = [];
    let currentGroup = [nightShiftDates[0]];
    for (let i = 1; i < nightShiftDates.length; i++) {
      const prev = nightShiftDates[i - 1];
      const curr = nightShiftDates[i];
      const diffTime = curr.getTime() - prev.getTime();
      const diffDays = Math.round(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays === 1) {
        currentGroup.push(curr);
      } else {
        groups.push(currentGroup);
        currentGroup = [curr];
      }
    }
    if (currentGroup.length > 0) {
      groups.push(currentGroup);
    }

    // Add to summary
    nightShiftSummary.push({
      name: emp,
      days: nightShiftDates.length
    });

    // Add to details
    for (const g of groups) {
      nightShiftDetails.push({
        name: emp,
        startDate: g[0],
        endDate: g[g.length - 1],
        total: g.length
      });
    }
  }

  // Sort summary alphabetically by name
  nightShiftSummary.sort((a, b) => a.name.localeCompare(b.name));

  // Sort details alphabetically by name, then by startDate
  nightShiftDetails.sort((a, b) => {
    const nameComp = a.name.localeCompare(b.name);
    if (nameComp !== 0) return nameComp;
    return a.startDate.getTime() - b.startDate.getTime();
  });

  allowanceStats.summary = nightShiftSummary;
  allowanceStats.details = nightShiftDetails;

  // ===== 8. 請假/加班 明細統計 =====
  const leaveDetails = [];
  for (const rec of state.parsed.origLeaves) {
    const empName = getCanonicalName(rec.applicant, allEmployees);
    if (!allEmployees.includes(empName)) continue;
    const y = rec.startDate.getFullYear();
    const m = rec.startDate.getMonth() + 1;
    if (y === scheduleYear && m === scheduleMonth) {
      leaveDetails.push({
        name: empName,
        leaveType: rec.leaveType,
        startDate: rec.startDate,
        endDate: rec.endDate,
        total: rec.days,
        timeRange: rec.timeRange
      });
    }
  }

  const otDetails = [];
  for (const rec of state.parsed.origOvertimes) {
    const empName = getCanonicalName(rec.applicant, allEmployees);
    if (!allEmployees.includes(empName)) continue;
    const y = rec.otDate.getFullYear();
    const m = rec.otDate.getMonth() + 1;
    if (y === scheduleYear && m === scheduleMonth) {
      otDetails.push({
        name: empName,
        startDate: rec.otDate,
        startTime: rec.startTime,
        endDate: rec.endDate,
        endTime: rec.endTime,
        total: rec.hours
      });
    }
  }

  return {
    monthlyLeave, monthlyLeaveTypes: [...monthlyLeaveTypes],
    yearlyLeave, yearlyLeaveTypes: [...yearlyLeaveTypes],
    otStats, otDates: [...otDates].sort(),
    totalStats, monthlyDetail,
    ptoSummary,
    allowanceStats,
    leaveDetails,
    otDetails,
    allEmployees, targetMonthKey,
    scheduleMonth, scheduleYear,
  };
}

// ==================== Excel 生成 ====================
function generateLeaveExcel(results) {
  const wb = XLSX.utils.book_new();

  // --- Sheet 1: 請假整理 ---
  const { monthlyLeave, monthlyLeaveTypes, yearlyLeave, yearlyLeaveTypes, allEmployees, scheduleYear, scheduleMonth, leaveDetails } = results;

  const sortedMonthlyTypes = monthlyLeaveTypes.sort((a, b) => {
    const ia = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[a]?.code);
    const ib = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[b]?.code);
    return ia - ib;
  });

  const sortedYearlyTypes = yearlyLeaveTypes.sort((a, b) => {
    const ia = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[a]?.code);
    const ib = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[b]?.code);
    return ia - ib;
  });

  // Build left-hand side rows
  const leftRows = [];
  leftRows.push(['月份', `${scheduleYear}/${scheduleMonth}/1`]);
  leftRows.push([]);
  leftRows.push(['當月請假統計']);
  
  const monthlyHdr = ['', ...sortedMonthlyTypes.map(t => LEAVE_TYPE_MAP[t].label), '總計'];
  leftRows.push(monthlyHdr);

  const empWithLeave = allEmployees.filter(e => monthlyLeave[e]);
  for (let i = 0; i < empWithLeave.length; i++) {
    const emp = empWithLeave[i];
    const rowIdx = i + 6; // excel row index (1-based, starts at row 6)
    const row = [emp];
    for (const lt of sortedMonthlyTypes) {
      row.push(monthlyLeave[emp]?.[lt] || '');
    }
    const lastColLetter = getColLetter(sortedMonthlyTypes.length + 1);
    row.push({ f: `SUM(B${rowIdx}:${lastColLetter}${rowIdx})` });
    leftRows.push(row);
  }

  // Totals row for monthly summary
  const totalRowIdx = empWithLeave.length + 6;
  const totalRow = ['總計'];
  for (let c = 2; c <= sortedMonthlyTypes.length + 1; c++) {
    const colLetter = getColLetter(c);
    totalRow.push({ f: `SUM(${colLetter}6:${colLetter}${totalRowIdx - 1})` });
  }
  const lastColLetter = getColLetter(sortedMonthlyTypes.length + 1);
  totalRow.push({ f: `SUM(${lastColLetter}6:${lastColLetter}${totalRowIdx - 1})` });
  leftRows.push(totalRow);

  leftRows.push([]);
  leftRows.push(['申請人', '(全部)']);
  leftRows.push([]);

  // 年度請假統計
  leftRows.push([`${scheduleYear}年度請假統計`]);
  const yHdr = ['', ...sortedYearlyTypes.map(t => LEAVE_TYPE_MAP[t].label), '總計'];
  leftRows.push(yHdr);

  const sortedMonths = Object.keys(yearlyLeave).sort();
  const yearlyStartRowIdx = totalRowIdx + 5; // e.g. 14 + 5 = 19
  for (let i = 0; i < sortedMonths.length; i++) {
    const mk = sortedMonths[i];
    const pts = mk.split('-');
    const dateStr = `${pts[0]}/${parseInt(pts[1])}/1`;
    const rowIdx = yearlyStartRowIdx + i + 1;
    const row = [dateStr];
    for (const lt of sortedYearlyTypes) {
      row.push(yearlyLeave[mk]?.[lt] || '');
    }
    const lastYColLetter = getColLetter(sortedYearlyTypes.length + 1);
    row.push({ f: `SUM(B${rowIdx}:${lastYColLetter}${rowIdx})` });
    leftRows.push(row);
  }

  // Year totals
  const yTotalRowIdx = yearlyStartRowIdx + sortedMonths.length + 1;
  const yTotal = ['總計'];
  for (let c = 2; c <= sortedYearlyTypes.length + 1; c++) {
    const colLetter = getColLetter(c);
    yTotal.push({ f: `SUM(${colLetter}${yearlyStartRowIdx + 1}:${colLetter}${yTotalRowIdx - 1})` });
  }
  const lastYColLetter = getColLetter(sortedYearlyTypes.length + 1);
  yTotal.push({ f: `SUM(${lastYColLetter}${yearlyStartRowIdx + 1}:${lastYColLetter}${yTotalRowIdx - 1})` });
  leftRows.push(yTotal);

  // Combine Left and Right Tables
  const rows1 = [];
  rows1.push([]); // Row 1 is empty

  const maxLen = Math.max(leftRows.length, leaveDetails.length + 4);
  const rightStartColIdx = sortedMonthlyTypes.length + 3; // e.g. 2 types -> index 5 (Column F)

  for (let i = 0; i < maxLen; i++) {
    const row = [];

    // Fill left table
    if (i < leftRows.length) {
      const leftRow = leftRows[i];
      for (let c = 0; c < leftRow.length; c++) {
        row[c] = leftRow[c];
      }
    }

    // Fill right table (starts at Row 4 index 2 of leftRows)
    if (i >= 2) {
      const rightRowIdx = i - 2; // 0 for Row 4 (Excel Row 4)
      if (rightRowIdx === 0) {
        row[rightStartColIdx] = 'Name';
        row[rightStartColIdx + 1] = 'Type of Leave';
        row[rightStartColIdx + 2] = 'Start Date';
        row[rightStartColIdx + 3] = 'End Date';
        row[rightStartColIdx + 4] = 'Total';
        row[rightStartColIdx + 5] = 'Time Period';
      } else if (rightRowIdx - 1 < leaveDetails.length) {
        const rec = leaveDetails[rightRowIdx - 1];
        row[rightStartColIdx] = rec.name;
        row[rightStartColIdx + 1] = rec.leaveType;
        row[rightStartColIdx + 2] = formatShiftDate(rec.startDate);
        row[rightStartColIdx + 3] = formatShiftDate(rec.endDate);
        row[rightStartColIdx + 4] = rec.total;
        row[rightStartColIdx + 5] = rec.timeRange;
      }
    }

    // Clean undefined cells to empty strings
    for (let c = 0; c < row.length; c++) {
      if (row[c] === undefined) row[c] = '';
    }

    rows1.push(row);
  }

  const ws1 = XLSX.utils.aoa_to_sheet(rows1);

  // Column width config
  const colsConfig = [];
  colsConfig[0] = { wch: 16 };
  for (let c = 1; c <= sortedMonthlyTypes.length; c++) {
    colsConfig[c] = { wch: 12 };
  }
  colsConfig[sortedMonthlyTypes.length + 1] = { wch: 10 };
  colsConfig[sortedMonthlyTypes.length + 2] = { wch: 5 }; // empty divider
  colsConfig[rightStartColIdx] = { wch: 16 };
  colsConfig[rightStartColIdx + 1] = { wch: 25 };
  colsConfig[rightStartColIdx + 2] = { wch: 12 };
  colsConfig[rightStartColIdx + 3] = { wch: 12 };
  colsConfig[rightStartColIdx + 4] = { wch: 8 };
  colsConfig[rightStartColIdx + 5] = { wch: 15 };
  ws1['!cols'] = colsConfig;

  XLSX.utils.book_append_sheet(wb, ws1, '請假整理');

  return wb;
}

function generateOvertimeExcel(results) {
  const wb = XLSX.utils.book_new();

  // --- Sheet 1: 加班整理 ---
  const { otStats, otDates, allEmployees, otDetails } = results;
  const empsWithOT = allEmployees.filter(e => otStats[e]);
  const sortedDates = otDates.sort();

  // Build left-hand side rows
  const leftRows = [];
  leftRows.push([]);
  leftRows.push([]);
  leftRows.push(['加總 - 共加班幾個小時']);

  // Header row with dates
  const formattedDates = sortedDates.map(d => {
    const pts = d.split('-');
    return `${parseInt(pts[0])}/${parseInt(pts[1])}/${parseInt(pts[2])}`;
  });
  const hdr = ['', ...formattedDates, '總計'];
  leftRows.push(hdr);

  // Employee rows (Formula-based row sums)
  const lastDateColLetter = getColLetter(sortedDates.length + 1);
  for (let i = 0; i < empsWithOT.length; i++) {
    const emp = empsWithOT[i];
    const rowIdx = i + 5;
    const row = [emp];
    for (const d of sortedDates) {
      row.push(otStats[emp]?.[d] || '');
    }
    row.push({ f: `SUM(B${rowIdx}:${lastDateColLetter}${rowIdx})` });
    leftRows.push(row);
  }

  // Totals Row (Formula-based column sums)
  const totalRowIdx = empsWithOT.length + 5;
  const totalRow = ['總計'];
  for (let c = 2; c <= sortedDates.length + 1; c++) {
    const colLetter = getColLetter(c);
    totalRow.push({ f: `SUM(${colLetter}5:${colLetter}${totalRowIdx - 1})` });
  }
  totalRow.push({ f: `SUM(${lastDateColLetter}5:${lastDateColLetter}${totalRowIdx - 1})` });
  leftRows.push(totalRow);

  // Combine Left and Right Tables
  const rows = [];
  const maxLen = Math.max(leftRows.length, otDetails.length + 4);
  const rightStartColIdx = sortedDates.length + 3; // e.g. Col A to ... plus 1 empty divider

  for (let i = 0; i < maxLen; i++) {
    const row = [];

    // Fill left table
    if (i < leftRows.length) {
      const leftRow = leftRows[i];
      for (let c = 0; c < leftRow.length; c++) {
        row[c] = leftRow[c];
      }
    }

    // Fill right table (starts at Row 4 index 3 of leftRows)
    if (i >= 3) {
      const rightRowIdx = i - 3; // 0 for Row 4 (Excel Row 4)
      if (rightRowIdx === 0) {
        row[rightStartColIdx] = 'Name';
        row[rightStartColIdx + 1] = 'Start Date';
        row[rightStartColIdx + 2] = 'Start Time';
        row[rightStartColIdx + 3] = 'End Date';
        row[rightStartColIdx + 4] = 'End Time';
        row[rightStartColIdx + 5] = 'Total';
      } else if (rightRowIdx - 1 < otDetails.length) {
        const rec = otDetails[rightRowIdx - 1];
        row[rightStartColIdx] = rec.name;
        row[rightStartColIdx + 1] = formatShiftDate(rec.startDate);
        row[rightStartColIdx + 2] = rec.startTime;
        row[rightStartColIdx + 3] = formatShiftDate(rec.endDate);
        row[rightStartColIdx + 4] = rec.endTime;
        row[rightStartColIdx + 5] = rec.total;
      }
    }

    // Clean undefined cells to empty strings
    for (let c = 0; c < row.length; c++) {
      if (row[c] === undefined) row[c] = '';
    }

    rows.push(row);
  }

  const ws1 = XLSX.utils.aoa_to_sheet(rows);

  // Column width config
  const colsConfig = [];
  colsConfig[0] = { wch: 16 };
  for (let c = 1; c <= sortedDates.length; c++) {
    colsConfig[c] = { wch: 8 };
  }
  colsConfig[sortedDates.length + 1] = { wch: 10 };
  colsConfig[sortedDates.length + 2] = { wch: 5 }; // empty divider
  colsConfig[rightStartColIdx] = { wch: 16 };
  colsConfig[rightStartColIdx + 1] = { wch: 12 };
  colsConfig[rightStartColIdx + 2] = { wch: 10 };
  colsConfig[rightStartColIdx + 3] = { wch: 12 };
  colsConfig[rightStartColIdx + 4] = { wch: 10 };
  colsConfig[rightStartColIdx + 5] = { wch: 8 };
  ws1['!cols'] = colsConfig;

  XLSX.utils.book_append_sheet(wb, ws1, '加班整理');

  return wb;
}

function generateTotalLeaveExcel(results) {
  const wb = XLSX.utils.book_new();
  const { totalStats, ptoSummary, allEmployees, monthlyDetail, scheduleYear, scheduleMonth } = results;

  // --- Sheet 1: 統計表 ---
  const rows = [];

  // Row 1: headers
  const r1 = [scheduleYear, '', '', '特休'];
  for (let i = 0; i < 13; i++) r1.push('');
  r1.push('病假/事假/曠職/遲到');
  for (let i = 0; i < 10; i++) r1.push('');
  rows.push(r1);

  // Row 2: sub headers
  const r2 = ['', '特別\n休假', '亞勝\n假期'];
  for (let m = 1; m <= 12; m++) {
    r2.push(`${scheduleYear}/${m}/1`); // formatted date string
  }
  r2.push('已休\n總天數', '剩餘\n天數', '', '', '病假', '生理假', '事假', '公假', '婚假', '喪假', '家庭假', '遲到', '曠職', '');
  rows.push(r2);

  // Row 3: code headers
  const r3 = ['', '', '', ...Array(12).fill('已用天數'), '', '', '病假', '生理假', 'SL', 'SL-M', 'PL', 'LOA', 'ML', 'BL', 'FL', '', '', ''];
  rows.push(r3);

  // Employee rows with full formulas matching target
  for (let i = 0; i < allEmployees.length; i++) {
    const emp = allEmployees[i];
    const rowIdx = i + 4; // Excel Row index is 4-based
    const basePto = state.parsed.ptoData.find(p => getCanonicalName(p.name, allEmployees) === emp) || { pto: 0, ptoAl: 0 };

    const row = [emp, basePto.pto || 0, basePto.ptoAl || 0];

    // Col D to O (Jan to Dec PTO formulas)
    for (let m = 1; m <= 12; m++) {
      if (m <= scheduleMonth) {
        const monthSheet = MONTH_NAMES_EN[m - 1] + (m === 1 ? ' ' : '');
        row.push({ f: `SUM('${monthSheet}'!D${rowIdx}:E${rowIdx})` });
      } else {
        row.push('');
      }
    }

    // Col P (已休總天數) & Col Q (剩餘天數)
    row.push({ f: `SUM(D${rowIdx}:O${rowIdx})` });
    row.push({ f: `B${rowIdx}+C${rowIdx}-P${rowIdx}` });

    // Col R to AB (Other leave cumulative formulas)
    row.push({ f: makeYearlySumFormula('F', rowIdx) }); // Col R (病假)
    row.push({ f: makeYearlySumFormula('G', rowIdx) }); // Col S (生理假)
    row.push({ f: `R${rowIdx}+MAX(S${rowIdx}-3,0)` }); // Col T (病假付給)
    row.push({ f: `MIN(S${rowIdx},3)` }); // Col U (生理假付給)
    row.push({ f: makeYearlySumFormula('H', rowIdx) }); // Col V (事假)
    row.push({ f: makeYearlySumFormula('I', rowIdx) }); // Col W (公假)
    row.push({ f: makeYearlySumFormula('J', rowIdx) }); // Col X (婚假)
    row.push({ f: makeYearlySumFormula('K', rowIdx) }); // Col Y (喪假)
    row.push({ f: makeYearlySumFormula('L', rowIdx) }); // Col Z (家庭假)
    row.push({ f: makeYearlySumFormula('M', rowIdx) }); // Col AA (遲到)
    row.push({ f: makeYearlySumFormula('N', rowIdx) }); // Col AB (曠職)
    rows.push(row);
  }

  const ws1 = XLSX.utils.aoa_to_sheet(rows);
  ws1['!cols'] = [{ wch: 16 }, { wch: 8 }, { wch: 8 }, ...Array(12).fill({ wch: 10 }), { wch: 10 }, { wch: 8 }, { wch: 6 }, { wch: 6 }, ...Array(9).fill({ wch: 6 })];
  
  const lastDay = new Date(scheduleYear, scheduleMonth, 0).getDate();
  const summarySheetName = `統計表(統計至${scheduleMonth}.${lastDay}`;
  XLSX.utils.book_append_sheet(wb, ws1, summarySheetName);

  // --- Monthly Sheets (Jan ~ Dec) ---
  const leaveTypeLabels = [
    ['PTO', 'Annual Leave特別休假'],
    ['PTO-AL', 'Asurion Leave亞勝假期'],
    ['Half-PTO/PTO-Half', '上特/下特'],
    ['Half-AL/AL-Half', '上亞勝假期/下亞勝假期\n'],
    ['SL', 'Sick Leave病假'],
    ['SL-M', 'Menstrual Leave病假（生理假)'],
    ['PL', 'Personal Leave事假'],
    ['LOA', 'Official Leave公假'],
    ['LOA', 'Official Leave公假（金融市場常識與職業道德考試）'],
    ['LOA', 'Official Leave公假（財產保險業務員資格證照考試)'],
    ['LOA', 'Official Leave公假（健檢)'],
    ['ML', 'Marriage Leave婚假'],
    ['', 'Maternity Leave產假'],
    ['', 'Paternity Leave陪產假'],
    ['', 'Pregnancy Checkup Accompaniment Leave陪產檢假'],
    ['BL', 'Bereavement Leave喪假'],
    ['FL', 'Family Care Leave家庭照顧假'],
  ];

  for (let m = 1; m <= 12; m++) {
    const mRows = [];
    const monthNameZh = MONTH_NAMES_ZH[m - 1];
    const sheetName = MONTH_NAMES_EN[m - 1] + (m === 1 ? ' ' : '');

    // Row 1: header
    mRows.push(['', '', `${scheduleYear}${monthNameZh}`]);
    // Row 2: column headers
    mRows.push(['', '', '', '特休', '亞勝假期', '病假', '生理假', '事假', '公假', '婚假', '喪假', '家庭假', '遲到', '曠職']);
    // Row 3: code headers
    mRows.push(['', '', '', 'PTO', 'PTO-AL', 'SL', 'SL-M', 'PL', 'LOA', 'ML', 'BL', 'FL', '', '']);

    // Employee data rows (aligned with leave type labels)
    const empIndex = allEmployees.slice(); // copy
    for (let lt = 0; lt < leaveTypeLabels.length; lt++) {
      const [code, label] = leaveTypeLabels[lt];
      const row = [code, label];

      if (lt < empIndex.length) {
        const emp = empIndex[lt];
        row.push(emp);

        const detail = monthlyDetail[m]?.[emp] || {};
        const getVal = (code) => {
          const val = detail[code];
          return (val === undefined || val === null || val === 0) ? null : val;
        };
        row.push(getVal('PTO'));
        row.push(getVal('PTO-AL'));
        row.push(getVal('SL'));
        row.push(getVal('SL-M'));
        row.push(getVal('PL'));
        row.push(getVal('LOA'));
        row.push(getVal('ML'));
        row.push(getVal('BL'));
        row.push(getVal('FL'));
        row.push(null);
        row.push(null);
      } else if (lt === empIndex.length) {
        // Totals row (Row 13 in Excel) with dynamic formula
        row.push('');
        const endRow = empIndex.length + 3; // rows 4 to endRow
        row.push(
          { f: `SUM(D4:D${endRow})` },
          { f: `SUM(E4:E${endRow})` },
          { f: `SUM(F4:F${endRow})` },
          { f: `SUM(G4:G${endRow})` },
          { f: `SUM(H4:H${endRow})` },
          { f: `SUM(I4:I${endRow})` },
          { f: `SUM(J4:J${endRow})` },
          { f: `SUM(K4:K${endRow})` },
          { f: `SUM(L4:L${endRow})` },
          { f: `SUM(M4:M${endRow})` },
          { f: `SUM(N4:N${endRow})` }
        );
      } else {
        row.push('');
      }

      mRows.push(row);
    }

    const ws = XLSX.utils.aoa_to_sheet(mRows);
    ws['!cols'] = [{ wch: 20 }, { wch: 52 }, { wch: 16 }, ...Array(11).fill({ wch: 8 })];
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  return wb;
}

function generateUpdatedPtoSheet(results) {
  const { allEmployees, scheduleMonth, totalStats } = results;
  const rows = [];
  
  // Headers
  rows.push(['English Name ', 'PTO', 'PTO-AL', 'TTL', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', '已休', 'PTO', 'PTO-AL', '未休']);
  
  for (let i = 0; i < allEmployees.length; i++) {
    const emp = allEmployees[i];
    const rowIdx = i + 2; // Row index is 2-based (Row 1 is headers)
    
    // Find matching employee base numbers in parsed.ptoData
    const basePto = state.parsed.ptoData.find(p => getCanonicalName(p.name, allEmployees) === emp) || { pto: 0, ptoAl: 0 };
    
    const row = [];
    row.push(emp); // Col A
    row.push(basePto.pto || 0); // Col B
    row.push(basePto.ptoAl || 0); // Col C
    row.push({ f: `B${rowIdx}+C${rowIdx}` }); // Col D (TTL)
    
    // Col E to P (Jan to Dec monthly PTO usage)
    for (let m = 1; m <= 12; m++) {
      if (m <= scheduleMonth) {
        const usage = totalStats[emp]?.monthlyPTO[m] || 0;
        row.push(usage || '');
      } else {
        row.push('');
      }
    }
    
    row.push({ f: `SUM(E${rowIdx}:P${rowIdx})` }); // Col Q (已休)
    row.push({ f: `IF(T${rowIdx}<=C${rowIdx}, 0, MIN(B${rowIdx}, T${rowIdx}-C${rowIdx}))` }); // Col R (PTO remaining)
    row.push({ f: `IF(T${rowIdx}<=C${rowIdx}, T${rowIdx}, C${rowIdx})` }); // Col S (PTO-AL remaining)
    row.push({ f: `D${rowIdx}-Q${rowIdx}` }); // Col T (未休)
    
    rows.push(row);
  }
  
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 16 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, ...Array(12).fill({ wch: 8 }), { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }];
  return ws;
}

function generateAllowanceSheet(results) {
  const { allowanceStats } = results;
  const { summary, details } = allowanceStats;

  const rows = [];
  // Row 1: Headers
  rows.push(['津貼統計(以天數計)', '小夜天數', '', 'Name', 'Start Date', 'End Date', 'Total']);

  const maxRows = Math.max(summary.length + 1, details.length); // excluding header

  for (let i = 0; i < maxRows; i++) {
    const row = ['', '', '', '', '', '', ''];

    // Left table (Summary + Total)
    if (i < summary.length) {
      row[0] = summary[i].name;
      row[1] = summary[i].days;
    } else if (i === summary.length) {
      row[0] = 'Total';
      row[1] = { f: `SUM(B2:B${summary.length + 1})` };
    }

    // Right table (Details)
    if (i < details.length) {
      row[3] = details[i].name;
      row[4] = formatShiftDate(details[i].startDate);
      row[5] = formatShiftDate(details[i].endDate);
      row[6] = details[i].total;
    }

    rows.push(row);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{ wch: 22 }, { wch: 12 }, { wch: 5 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 8 }];
  return ws;
}

// ==================== 預覽渲染 ====================
function renderPreview(results) {
  renderLeavePreview(results);
  renderOvertimePreview(results);
  renderTotalPreview(results);
  renderPTOPreview(results);
  renderAllowancePreview(results);
  renderPersonalSelect(results);
}

function createTable(headers, rows, footerRow) {
  let html = '<div class="table-wrapper"><table><thead><tr>';
  for (const h of headers) html += `<th>${h}</th>`;
  html += '</tr></thead><tbody>';
  for (const row of rows) {
    html += '<tr>';
    for (const cell of row) html += `<td>${cell === '' || cell === 0 ? '' : cell}</td>`;
    html += '</tr>';
  }
  if (footerRow) {
    html += '</tbody><tfoot><tr>';
    for (const cell of footerRow) html += `<td>${cell === '' || cell === 0 ? '' : cell}</td>`;
    html += '</tr></tfoot>';
  } else {
    html += '</tbody>';
  }
  html += '</table></div>';
  return html;
}

function renderLeavePreview(results) {
  const { monthlyLeave, monthlyLeaveTypes, yearlyLeave, yearlyLeaveTypes, allEmployees, scheduleMonth, scheduleYear, leaveDetails } = results;

  const sortedTypes = monthlyLeaveTypes.sort((a, b) => {
    const ia = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[a]?.code);
    const ib = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[b]?.code);
    return ia - ib;
  });

  let html = `<h3>📋 ${scheduleYear} 年 ${scheduleMonth} 月 請假統計</h3>`;
  html += '<div class="preview-layout">';

  // 1. Left Table: Monthly Summary
  html += '<div class="preview-col-summary">';
  html += '<h4>當月請假統計 (天數)</h4>';
  const headers = ['員工', ...sortedTypes.map(t => LEAVE_TYPE_MAP[t]?.shortLabel || t), '總計'];
  const rows = [];
  const emps = allEmployees.filter(e => monthlyLeave[e]);
  const totals = ['總計'];

  for (const emp of emps) {
    const row = [emp];
    let total = 0;
    for (const lt of sortedTypes) {
      const val = monthlyLeave[emp]?.[lt] || 0;
      row.push(val);
      total += val;
    }
    row.push(total);
    rows.push(row);
  }

  for (let c = 0; c < sortedTypes.length; c++) {
    let sum = 0;
    for (const emp of emps) sum += (monthlyLeave[emp]?.[sortedTypes[c]] || 0);
    totals.push(sum);
  }
  totals.push(rows.reduce((s, r) => s + r[r.length - 1], 0));

  html += createTable(headers, rows, totals);
  html += '</div>';

  // 2. Right Table: Details
  html += '<div class="preview-col-details">';
  html += '<h4>請假明細</h4>';
  const detailHeaders = ['Name', 'Type of Leave', 'Start Date', 'End Date', 'Total', 'Time Period'];
  const detailRows = leaveDetails.map(rec => [
    rec.name,
    rec.leaveType,
    formatShiftDate(rec.startDate),
    formatShiftDate(rec.endDate),
    rec.total,
    rec.timeRange
  ]);
  html += createTable(detailHeaders, detailRows, null);
  html += '</div>';

  html += '</div>'; // end preview-layout

  // Yearly table
  const sortedYearTypes = yearlyLeaveTypes.sort((a, b) => {
    const ia = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[a]?.code);
    const ib = LEAVE_CODE_ORDER.indexOf(LEAVE_TYPE_MAP[b]?.code);
    return ia - ib;
  });

  html += `<h3>📊 ${scheduleYear} 年度請假統計</h3>`;
  const yHeaders = ['月份', ...sortedYearTypes.map(t => LEAVE_TYPE_MAP[t]?.shortLabel || t), '總計'];
  const yRows = [];
  const sortedMonths = Object.keys(yearlyLeave).sort();

  for (const mk of sortedMonths) {
    const [y, m] = mk.split('-');
    const row = [`${y}-${m}`];
    let mTotal = 0;
    for (const lt of sortedYearTypes) {
      const val = yearlyLeave[mk]?.[lt] || 0;
      row.push(val);
      mTotal += val;
    }
    row.push(mTotal);
    yRows.push(row);
  }

  const yTotals = ['總計'];
  for (const lt of sortedYearTypes) {
    let sum = 0;
    for (const mk of sortedMonths) sum += (yearlyLeave[mk]?.[lt] || 0);
    yTotals.push(sum);
  }
  yTotals.push(yRows.reduce((s, r) => s + r[r.length - 1], 0));

  html += createTable(yHeaders, yRows, yTotals);

  document.getElementById('tab-leave').innerHTML = html;
}

function renderOvertimePreview(results) {
  const { otStats, otDates, allEmployees, otDetails } = results;
  const emps = allEmployees.filter(e => otStats[e]);
  const sorted = otDates.sort();

  let html = '<h3>⏰ 加班統計</h3>';
  html += '<div class="preview-layout">';

  // 1. Left Table: Monthly Summary
  html += '<div class="preview-col-summary">';
  html += '<h4>當月加班統計 (小時)</h4>';
  const headers = ['員工', ...sorted.map(d => d.slice(5)), '總計'];
  const rows = [];
  for (const emp of emps) {
    const row = [emp];
    let total = 0;
    for (const d of sorted) {
      const val = otStats[emp]?.[d] || 0;
      row.push(val);
      total += val;
    }
    row.push(total);
    rows.push(row);
  }

  const totals = ['總計'];
  for (const d of sorted) {
    let sum = 0;
    for (const emp of emps) sum += (otStats[emp]?.[d] || 0);
    totals.push(sum);
  }
  totals.push(rows.reduce((s, r) => s + r[r.length - 1], 0));

  html += createTable(headers, rows, totals);
  html += '</div>';

  // 2. Right Table: Details
  html += '<div class="preview-col-details">';
  html += '<h4>加班明細</h4>';
  const detailHeaders = ['Name', 'Start Date', 'Start Time', 'End Date', 'End Time', 'Total'];
  const detailRows = otDetails.map(rec => [
    rec.name,
    formatShiftDate(rec.startDate),
    rec.startTime,
    formatShiftDate(rec.endDate),
    rec.endTime,
    rec.total
  ]);
  html += createTable(detailHeaders, detailRows, null);
  html += '</div>';

  html += '</div>'; // end preview-layout

  document.getElementById('tab-overtime').innerHTML = html;
}

function renderTotalPreview(results) {
  const { totalStats, ptoSummary, allEmployees, scheduleYear, scheduleMonth } = results;

  let html = `<h3>📊 ${scheduleYear} 總請假統計 (統計至 ${scheduleMonth} 月)</h3>`;

  const headers = ['員工', '特休', '亞勝', ...MONTH_NAMES_EN.slice(0, scheduleMonth).map(m => m.slice(0, 3)), '已休', '剩餘', 'SL', 'SL-M', 'PL', 'LOA'];
  const rows = [];

  for (const emp of allEmployees) {
    const ptoInfo = ptoSummary[emp] || {};
    const empStats = totalStats[emp] || { monthlyPTO: {}, otherLeave: {} };

    const row = [emp, ptoInfo.pto || 0, ptoInfo.ptoAl || 0];

    let totalUsed = 0;
    for (let m = 1; m <= scheduleMonth; m++) {
      const usage = empStats.monthlyPTO[m] || 0;
      row.push(usage || '');
      totalUsed += usage;
    }

    row.push(totalUsed || '');
    row.push(((ptoInfo.pto || 0) + (ptoInfo.ptoAl || 0) - totalUsed) || '');
    row.push(empStats.otherLeave['SL'] || '');
    row.push(empStats.otherLeave['SL-M'] || '');
    row.push(empStats.otherLeave['PL'] || '');
    row.push(empStats.otherLeave['LOA'] || '');
    rows.push(row);
  }

  html += createTable(headers, rows);
  document.getElementById('tab-total').innerHTML = html;
}

function renderPTOPreview(results) {
  const { ptoSummary, allEmployees, scheduleMonth } = results;

  let html = '<h3>🏖 特休日數</h3>';
  const headers = ['員工', 'PTO', 'PTO-AL', '合計', ...MONTH_NAMES_EN.slice(0, scheduleMonth), '已休', '剩餘PTO', '剩餘AL', '剩餘合計'];
  const rows = [];

  for (const emp of allEmployees) {
    const info = ptoSummary[emp];
    if (!info) continue;

    const row = [emp, info.pto, info.ptoAl, info.ttl];
    for (let m = 1; m <= scheduleMonth; m++) {
      row.push(info.monthlyUsage[m] || '');
    }
    row.push(info.totalUsed || '');
    row.push(info.ptoRemaining);
    row.push(info.ptoAlRemaining);
    row.push(info.totalRemaining);
    rows.push(row);
  }

  html += createTable(headers, rows);
  document.getElementById('tab-pto').innerHTML = html;
}

function renderAllowancePreview(results) {
  const { allowanceStats } = results;
  const { summary, details } = allowanceStats;

  let html = '<div class="allowance-layout">';

  // 1. Left Table: Summary
  html += '<div class="allowance-col-summary">';
  html += '<h3>🌙 津貼統計 (以天數計)</h3>';
  
  const summaryHeaders = ['客服人員', '小夜天數'];
  const summaryRows = summary.map(item => [item.name, item.days]);
  const totalDays = summary.reduce((sum, item) => sum + item.days, 0);
  const summaryFooter = ['Total', totalDays];
  
  html += createTable(summaryHeaders, summaryRows, summaryFooter);
  html += '</div>';

  // 2. Right Table: Details
  html += '<div class="allowance-col-details">';
  html += '<h3>📋 小夜班明細</h3>';
  
  const detailHeaders = ['Name', 'Start Date', 'End Date', 'Total'];
  const detailRows = details.map(item => [
    item.name,
    formatShiftDate(item.startDate),
    formatShiftDate(item.endDate),
    item.total
  ]);
  
  html += createTable(detailHeaders, detailRows, null);
  html += '</div>';

  html += '</div>';

  document.getElementById('tab-allowance').innerHTML = html;
}

// ==================== 檔案下載 ====================
function downloadAll() {
  if (!state.results) return;

  const { scheduleYear, scheduleMonth } = state.results;

  const mergedWb = XLSX.utils.book_new();

  // 1. 請假統計 (請假整理 sheet renamed to 請假統計)
  const leaveWb = generateLeaveExcel(state.results);
  XLSX.utils.book_append_sheet(mergedWb, leaveWb.Sheets['請假整理'], '請假統計');

  // 2. 加班統計 (加班整理 sheet renamed to 加班統計)
  const otWb = generateOvertimeExcel(state.results);
  XLSX.utils.book_append_sheet(mergedWb, otWb.Sheets['加班整理'], '加班統計');

  // 3. 總請假統計 - 統計表 renamed to 總請假統計 & Jan-Dec monthly sheets (will be hidden)
  const totalWb = generateTotalLeaveExcel(state.results);
  const totalSheetNames = totalWb.SheetNames;
  const summarySheetName = totalSheetNames.find(name => name.includes('統計表'));
  if (summarySheetName) {
    XLSX.utils.book_append_sheet(mergedWb, totalWb.Sheets[summarySheetName], '總請假統計');
  }
  for (const sheetName of totalSheetNames) {
    if (sheetName !== summarySheetName) {
      XLSX.utils.book_append_sheet(mergedWb, totalWb.Sheets[sheetName], sheetName);
    }
  }

  // 4. 特休日數 sheet (Updated with formulas)
  const ptoWs = generateUpdatedPtoSheet(state.results);
  XLSX.utils.book_append_sheet(mergedWb, ptoWs, '特休日數');

  // 5. 津貼統計 sheet
  const allowanceWs = generateAllowanceSheet(state.results);
  XLSX.utils.book_append_sheet(mergedWb, allowanceWs, '津貼統計');

  // Set sheet visibility. Hide the monthly detailed sheets (Jan ~ Dec).
  // This keeps formulas working while displaying only the 4 main items.
  mergedWb.Workbook = {
    Sheets: mergedWb.SheetNames.map(name => {
      const isMonthly = MONTH_NAMES_EN.some(m => name.trim() === m);
      return {
        Hidden: isMonthly ? 1 : 0
      };
    })
  };

  XLSX.writeFile(mergedWb, `${scheduleYear}.${scheduleMonth}合併分析報表_Soluto_&_Care.xlsx`);

  showToast('合併分析報表已下載！', 'success');
}

// ==================== UI 事件處理 ====================
function initUI() {
  // Upload cards
  document.querySelectorAll('.upload-card').forEach(card => {
    const dropZone = card.querySelector('.drop-zone');
    const fileInput = card.querySelector('.file-input');
    const dataType = card.dataset.type;

    // Click to upload
    dropZone.addEventListener('click', () => fileInput.click());

    // File input change
    fileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) handleFile(e.target.files[0], dataType, card);
    });

    // Drag & drop
    dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0], dataType, card);
    });
  });

  // Analyze button
  document.getElementById('btn-analyze').addEventListener('click', runAnalysis);

  // Download button
  document.getElementById('btn-download').addEventListener('click', downloadAll);

  // Tabs
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');
    });
  });

  // Personal report controls
  const agentSelect = document.getElementById('personal-agent-select');
  if (agentSelect) {
    agentSelect.addEventListener('change', (e) => {
      renderPersonalCard(e.target.value);
    });
  }

  const btnExportCard = document.getElementById('btn-export-card');
  if (btnExportCard) {
    btnExportCard.addEventListener('click', () => {
      const empName = agentSelect.value;
      if (!empName) {
        showToast('請先選擇客服人員', 'warning');
        return;
      }
      exportPersonalCardImage(empName);
    });
  }

  const btnExportAllZip = document.getElementById('btn-export-all-zip');
  if (btnExportAllZip) {
    btnExportAllZip.addEventListener('click', () => {
      exportAllCardsToZip();
    });
  }
}

// ==================== 個人統計與圖片匯出邏輯 ====================
function renderPersonalSelect(results) {
  const select = document.getElementById('personal-agent-select');
  if (!select) return;

  const activeEmployees = results.allEmployees;
  select.innerHTML = '<option value="">-- 請選擇人員 --</option>';

  // Sort alphabetically using Chinese collation
  const sortedEmps = [...activeEmployees].sort((a, b) => a.localeCompare(b, 'zh-Hant'));

  for (const emp of sortedEmps) {
    const opt = document.createElement('option');
    opt.value = emp;
    opt.textContent = emp;
    select.appendChild(opt);
  }

  // Clear card content to placeholder
  document.getElementById('personal-report-card').innerHTML = `
    <div class="empty-card-placeholder">
        <div class="placeholder-icon">👤</div>
        <p>請選擇客服人員以檢視個人出勤與津貼統計</p>
    </div>
  `;
}

function renderPersonalCard(empName) {
  const card = document.getElementById('personal-report-card');
  if (!card) return;

  if (!empName) {
    card.innerHTML = `
      <div class="empty-card-placeholder">
          <div class="placeholder-icon">👤</div>
          <p>請選擇客服人員以檢視個人出勤與津貼統計</p>
      </div>
    `;
    return;
  }

  const { scheduleYear, scheduleMonth, leaveDetails, otDetails, allowanceStats } = state.results;

  // 1. Filter Leaves
  const leaves = leaveDetails.filter(rec => rec.name === empName);
  const totalLeaveDays = leaves.reduce((sum, rec) => sum + toNum(rec.total), 0);

  // 2. Filter Overtimes
  const ots = otDetails.filter(rec => rec.name === empName);
  const totalOtHours = ots.reduce((sum, rec) => sum + toNum(rec.total), 0);

  // 3. Night shifts
  const nightShiftSummary = allowanceStats.summary.find(s => s.name === empName);
  const totalNightDays = nightShiftSummary ? nightShiftSummary.days : 0;
  const nightShiftRanges = allowanceStats.details.filter(rec => rec.name === empName);

  // Build the Card HTML
  let html = `
    <!-- Card Header -->
    <div class="card-header">
        <div class="card-header-logo">
            <div class="logo-brand">
                <span class="brand-blue">ASU</span><span class="brand-orange">RION</span>
            </div>
            <div class="card-header-title">月度出勤與津貼明細</div>
        </div>
        <div class="card-header-meta">
            <div class="card-meta-name">${empName}</div>
            <div class="card-meta-month">${scheduleYear} 年 ${scheduleMonth} 月</div>
        </div>
    </div>

    <!-- Summary Metrics -->
    <div class="card-summary-grid">
        <div class="summary-stat-card leave">
            <div class="stat-value">${totalLeaveDays}<span>天</span></div>
            <div class="stat-label">請假總計</div>
        </div>
        <div class="summary-stat-card overtime">
            <div class="stat-value">${totalOtHours}<span>小時</span></div>
            <div class="stat-label">加班總計</div>
        </div>
        <div class="summary-stat-card allowance">
            <div class="stat-value">${totalNightDays}<span>天</span></div>
            <div class="stat-label">小夜津貼天數</div>
        </div>
    </div>

    <!-- Details 1: Please-leave -->
    <div class="card-detail-section">
        <div class="section-title">📋 請假明細</div>
  `;

  if (leaves.length === 0) {
    html += `<div class="empty-details">無請假記錄</div>`;
  } else {
    html += `
      <table>
          <thead>
              <tr>
                  <th>假別</th>
                  <th>開始時間</th>
                  <th>結束時間</th>
                  <th>總計(天)</th>
                  <th>時間段</th>
              </tr>
          </thead>
          <tbody>
    `;
    for (const rec of leaves) {
      html += `
              <tr>
                  <td>${rec.leaveType}</td>
                  <td>${formatShiftDate(rec.startDate)}</td>
                  <td>${formatShiftDate(rec.endDate)}</td>
                  <td>${rec.total}</td>
                  <td>${rec.timeRange || ''}</td>
              </tr>
      `;
    }
    html += `
          </tbody>
      </table>
    `;
  }
  html += `</div>`;

  // Details 2: Overtime
  html += `
    <div class="card-detail-section">
        <div class="section-title">⏰ 加班明細</div>
  `;
  if (ots.length === 0) {
    html += `<div class="empty-details">無加班記錄</div>`;
  } else {
    html += `
      <table>
          <thead>
              <tr>
                  <th>開始日期</th>
                  <th>開始時間</th>
                  <th>結束日期</th>
                  <th>結束時間</th>
                  <th>時數</th>
              </tr>
          </thead>
          <tbody>
    `;
    for (const rec of ots) {
      html += `
              <tr>
                  <td>${formatShiftDate(rec.startDate)}</td>
                  <td>${rec.startTime || ''}</td>
                  <td>${formatShiftDate(rec.endDate)}</td>
                  <td>${rec.endTime || ''}</td>
                  <td>${rec.total}</td>
              </tr>
      `;
    }
    html += `
          </tbody>
      </table>
    `;
  }
  html += `</div>`;

  // Details 3: Night shift allowance
  html += `
    <div class="card-detail-section">
        <div class="section-title">🌙 小夜班明細 (津貼)</div>
  `;
  if (nightShiftRanges.length === 0) {
    html += `<div class="empty-details">無小夜班記錄</div>`;
  } else {
    html += `
      <table>
          <thead>
              <tr>
                  <th>開始日期</th>
                  <th>結束日期</th>
                  <th>天數</th>
              </tr>
          </thead>
          <tbody>
    `;
    for (const rec of nightShiftRanges) {
      html += `
              <tr>
                  <td>${formatShiftDate(rec.startDate)}</td>
                  <td>${formatShiftDate(rec.endDate)}</td>
                  <td>${rec.total}</td>
              </tr>
      `;
    }
    html += `
          </tbody>
      </table>
    `;
  }
  html += `</div>`;

  card.innerHTML = html;
}

function exportPersonalCardImage(empName) {
  if (!state.results) return;
  const cardElement = document.getElementById('personal-report-card');
  const { scheduleYear, scheduleMonth } = state.results;

  showToast('正在生成個人圖片，請稍候...', 'warning');

  html2canvas(cardElement, {
    scale: 2,
    backgroundColor: '#ffffff',
    useCORS: true,
    logging: false
  }).then(canvas => {
    canvas.toBlob(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${empName}_${scheduleYear}.${scheduleMonth}_出勤明細.png`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast('個人出勤圖片下載成功！', 'success');
    }, 'image/png');
  }).catch(err => {
    console.error(err);
    showToast(`生成圖片失敗: ${err.message}`, 'error');
  });
}

async function exportAllCardsToZip() {
  if (!state.results) return;

  const select = document.getElementById('personal-agent-select');
  if (!select) return;

  const originalSelected = select.value;
  const { scheduleYear, scheduleMonth, allEmployees } = state.results;

  const overlay = document.getElementById('loading-overlay');
  const loadingText = overlay.querySelector('.loading-text');
  
  const originalOverlayText = loadingText.textContent;
  loadingText.textContent = '正在準備批次匯出...';
  overlay.style.display = 'flex';

  const sortedEmps = [...allEmployees].sort((a, b) => a.localeCompare(b, 'zh-Hant'));
  const zip = new JSZip();

  try {
    for (let i = 0; i < sortedEmps.length; i++) {
      const emp = sortedEmps[i];
      loadingText.textContent = `正在生成圖片 (${i + 1}/${sortedEmps.length}): ${emp}`;

      // 1. Render card in visible DOM
      select.value = emp;
      renderPersonalCard(emp);

      // Wait 2 frames for rendering engine to paint the card layout
      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));

      // 2. Snapshot the card
      const canvas = await html2canvas(document.getElementById('personal-report-card'), {
        scale: 2,
        backgroundColor: '#ffffff',
        useCORS: true,
        logging: false
      });

      // 3. Convert to blob and add to ZIP file
      const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
      zip.file(`${emp}_${scheduleYear}.${scheduleMonth}_出勤明細.png`, blob);
    }

    loadingText.textContent = '正在壓縮圖片並打包 ZIP...';
    const content = await zip.generateAsync({ type: 'blob' });

    // Download ZIP
    const url = URL.createObjectURL(content);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${scheduleYear}.${scheduleMonth}人員出勤明細圖片.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    showToast('所有人員圖片打包 ZIP 下載成功！', 'success');
  } catch (err) {
    console.error(err);
    showToast(`打包批次圖片失敗: ${err.message}`, 'error');
  } finally {
    // Restore select value and card display
    select.value = originalSelected;
    renderPersonalCard(originalSelected);

    // Hide overlay and reset text
    overlay.style.display = 'none';
    loadingText.textContent = originalOverlayText;
  }
}

async function handleFile(file, expectedType, card) {
  try {
    const data = await file.arrayBuffer();
    const wb = XLSX.read(data, { cellDates: false });

    // Auto-detect file type
    const detectedType = detectFileType(wb);

    if (detectedType && detectedType !== expectedType) {
      // The file was dropped in the wrong zone - auto-redirect
      const correctCard = document.querySelector(`.upload-card[data-type="${detectedType}"]`);
      if (correctCard) {
        card = correctCard;
        expectedType = detectedType;
      }
    }

    const finalType = detectedType || expectedType;

    state.files[finalType] = file;
    state.workbooks[finalType] = wb;

    // Parse immediately
    if (finalType === 'application') {
      const parsedData = parseApplicationData(wb);
      state.parsed.leave = parsedData.leave;
      state.parsed.overtime = parsedData.overtime;
      state.parsed.origLeaves = parsedData.origLeaves;
      state.parsed.origOvertimes = parsedData.origOvertimes;
    } else if (finalType === 'schedule') {
      const schedData = parseScheduleData(wb);
      state.parsed.employees = schedData.employees;
      state.parsed.ptoData = schedData.ptoData;
      state.parsed.holidays = schedData.holidays;
      state.parsed.scheduleMonth = schedData.month;
      state.parsed.scheduleYear = schedData.year;
      state.parsed.scheduleSheetName = schedData.scheduleSheetName;
    }

    // Update UI
    const targetCard = document.querySelector(`.upload-card[data-type="${finalType}"]`);
    if (targetCard) {
      targetCard.classList.add('loaded');
      const status = targetCard.querySelector('.file-status');
      status.classList.add('show');
      status.innerHTML = `<span class="check">✓</span> ${file.name}`;
      targetCard.querySelector('.drop-zone').style.display = 'none';
    }

    // Show detected info
    if (state.parsed.scheduleMonth && state.parsed.scheduleYear) {
      const info = document.getElementById('detected-info');
      info.style.display = 'flex';
      
      const currentYear = state.parsed.scheduleYear;
      const currentMonth = state.parsed.scheduleMonth;
      const targetMonthKey = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
      
      const yearLeaves = state.parsed.leave.filter(rec => {
        const year = parseInt(rec.monthKey.split('-')[0]);
        return year === currentYear;
      }).length;
      
      const monthOvertimes = state.parsed.overtime.filter(rec => {
        return rec.dateKey === targetMonthKey;
      }).length;
      
      info.innerHTML = `<span class="info-icon">📅</span> 偵測到月份：<strong>${currentYear} 年 ${currentMonth} 月</strong>
        <span class="info-detail">（員工：${state.parsed.employees.length} 人 | ${currentYear} 年請假記錄：${yearLeaves} 筆 | 當月加班記錄：${monthOvertimes} 筆）</span>`;
    }

    showToast(`${file.name} 載入成功`, 'success');
    checkReady();
  } catch (err) {
    console.error(err);
    showToast(`讀取檔案失敗: ${err.message}`, 'error');
  }
}

function checkReady() {
  const allLoaded = state.workbooks.application && state.workbooks.schedule;
  document.getElementById('btn-analyze').disabled = !allLoaded;
}

function runAnalysis() {
  const overlay = document.getElementById('loading-overlay');
  overlay.style.display = 'flex';

  setTimeout(() => {
    try {
      state.results = calculateAll();
      renderPreview(state.results);

      document.getElementById('results-section').style.display = 'block';
      document.getElementById('btn-download').style.display = 'inline-flex';

      // Activate first tab
      document.querySelector('.tab-btn[data-tab="leave"]').click();

      showToast('分析完成！', 'success');

      // Scroll to results
      document.getElementById('results-section').scrollIntoView({ behavior: 'smooth' });
    } catch (err) {
      console.error(err);
      showToast(`分析失敗: ${err.message}`, 'error');
    } finally {
      overlay.style.display = 'none';
    }
  }, 300);
}

// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', initUI);
