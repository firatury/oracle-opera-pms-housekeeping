/* global XLSX, pdfjsLib */
const startMenu = document.getElementById('startMenu');
const appShell = document.getElementById('appShell');
const startDeparturesBtn = document.getElementById('startDeparturesBtn');
const startArrivalsBtn = document.getElementById('startArrivalsBtn');
const startDndBtn = document.getElementById('startDndBtn');
const backMenuBtn = document.getElementById('backMenuBtn');
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const preview = document.getElementById('preview');
const summary = document.getElementById('summary');
const statusEl = document.getElementById('status');
const printBtn = document.getElementById('printBtn');
const excelBtn = document.getElementById('excelBtn');
const clearBtn = document.getElementById('clearBtn');
const officeBtn = document.getElementById('officeBtn');
const officeExcelBtn = document.getElementById('officeExcelBtn');
const assignmentPanel = document.getElementById('assignmentPanel');
const chiefControls = document.getElementById('chiefControls');
const assignmentStatus = document.getElementById('assignmentStatus');
const appTitle = document.getElementById('appTitle');
const departuresModeBtn = document.getElementById('departuresModeBtn');
const arrivalsModeBtn = document.getElementById('arrivalsModeBtn');
const dndModeBtn = document.getElementById('dndModeBtn');
const vacantModeBtn = document.getElementById('vacantModeBtn');
const greenPanel = document.getElementById('greenPanel');
const greenRoomsInput = document.getElementById('greenRoomsInput');
const currentRoomsPanel = document.getElementById('currentRoomsPanel');
const currentRoomsInput = document.getElementById('currentRoomsInput');
const currentRoomsStatus = document.getElementById('currentRoomsStatus');
const uploadTitle = document.getElementById('uploadTitle');

const MODE_DEPARTURES = 'departures';
const MODE_ARRIVALS = 'arrivals';
const MODE_DND = 'dnd';
const MODE_VACANT = 'vacant';
const ETD_HIGHLIGHT = '17:00';
const ROWS_PER_PAGE = 33;
const PRINT_PAGE_LIMITS = {
  departures: { hardMax: 29, unitBudget: 31 },
  arrivals: { hardMax: 24, unitBudget: 26 },
  vacant: { hardMax: 28, unitBudget: 30 },
};
const CHIEF_GROUPS = ['1000ler', '2000ler', '3000ler', '4000ler', '5000ler'];
const LEAVE_ELIGIBLE_GROUPS = CHIEF_GROUPS.filter(group => group !== '5000ler');

let currentMode = MODE_DEPARTURES;
let originalGroups = new Map();
let printableGroups = new Map();
let leaveGroups = new Set();
let sectionAssignments = new Map(); // key: "1000ler::1100ler", value: target chief group
let lastFileName = '';
let lastWorkbook = null;
let lastWorkbooks = [];
let greenRooms = new Set();
let currentRoomFilter = new Map(); // room -> { room, arrivalDate, source }
let currentRoomFileNames = [];
let dndResults = [];
let dndDateWindowText = '';
let dndFilterStats = { active: false, currentRooms: 0, skippedOldRooms: 0, stoppedBeforeArrival: 0 };

const FIELD_DEFS = [
  { key: 'room', out: 'Room', aliases: ['room', 'oda', 'oda no', 'room no', 'room number'] },
  { key: 'eta', out: 'ETA', aliases: ['eta', 'arrival time', 'geliş saati', 'giris saati'] },
  { key: 'arrival', out: 'Arrival', aliases: ['arrival', 'arrıval', 'geliş', 'gelis', 'arrival date'] },
  { key: 'adults', out: 'Adults', aliases: ['adults', 'adult', 'adult(s)', 'yetişkin', 'yetiskin', 'adults count'] },
  { key: 'children', out: 'Children', aliases: ['children', 'childr', 'child', 'çocuk', 'cocuk'] },
  { key: 'childAges', out: 'Child Ages', aliases: ['child ages', 'child age', 'ages', 'çocuk yaş', 'cocuk yas'] },
  { key: 'departure', out: 'Departure', aliases: ['departure', 'departures', 'ayrılış', 'ayrilis', 'departure date'] },
  { key: 'etd', out: 'ETD', aliases: ['etd', 'departure time', 'çıkış saati', 'cikis saati'] },
  { key: 'name', out: 'Name', aliases: ['name', 'guest name', 'guest', 'misafir', 'ad soyad'] },
  { key: 'travelAgent', out: 'Travel Agent', aliases: ['travel agent', 'agent', 'agency', 'acenta', 'travelagent'] },
];

const REQUIRED_KEYS = {
  [MODE_DEPARTURES]: ['room', 'eta', 'arrival', 'adults', 'children', 'childAges', 'departure', 'etd', 'travelAgent'],
  [MODE_ARRIVALS]: ['room', 'eta', 'arrival', 'adults', 'children', 'childAges', 'departure', 'etd', 'name', 'travelAgent'],
};

function modeLabel(mode = currentMode) {
  if (mode === MODE_ARRIVALS) return 'Arrivals';
  if (mode === MODE_DND) return 'DND / TİST';
  if (mode === MODE_VACANT) return 'Vacant Rooms';
  return 'Departures';
}

function requiredFields(mode = currentMode) {
  return (REQUIRED_KEYS[mode] || []).map(key => FIELD_DEFS.find(field => field.key === key));
}

function normalizeHeader(value) {
  return String(value ?? '')
    .trim()
    .toLocaleLowerCase('tr-TR')
    .replace(/\s+/g, ' ')
    .replace(/[._-]/g, ' ');
}

function canonical(value) {
  return normalizeHeader(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/ı/g, 'i')
    .replace(/[^a-z0-9 ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function clean(value) {
  if (value === null || value === undefined) return '';
  return String(value).replace(/\s+/g, ' ').trim();
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function setStatus(message, type = '') {
  statusEl.textContent = message;
  statusEl.className = `status ${type}`.trim();
}

function excelSerialToDate(serial, date1904 = false) {
  const offset = date1904 ? 1462 : 0;
  const utcDays = Math.floor(Number(serial) - 25569 + offset);
  const utcValue = utcDays * 86400;
  return new Date(utcValue * 1000);
}

function excelSerialToLocalDate(serial, date1904 = false) {
  const utcDate = excelSerialToDate(serial, date1904);
  if (!(utcDate instanceof Date) || isNaN(utcDate)) return null;
  // Excel tarihleri gün bazlıdır. UTC 00:00 olarak gelen değer bazı tarayıcı/saat dilimlerinde
  // bir önceki gün gibi görünebilir. Bu yüzden UTC gün/ay/yıl parçalarıyla yerel tarih oluşturuyoruz.
  return new Date(utcDate.getUTCFullYear(), utcDate.getUTCMonth(), utcDate.getUTCDate(), 12, 0, 0, 0);
}

function formatDateParts(day, month, year) {
  const y = String(year).padStart(4, '0');
  return `${String(day).padStart(2, '0')}.${String(month).padStart(2, '0')}.${y}`;
}

function formatDateValue(value) {
  if (value instanceof Date && !isNaN(value)) {
    return formatDateParts(value.getDate(), value.getMonth() + 1, value.getFullYear());
  }
  if (typeof value === 'number' && value > 20000 && value < 80000) {
    const d = excelSerialToLocalDate(value);
    return d ? formatDateParts(d.getDate(), d.getMonth() + 1, d.getFullYear()) : '';
  }
  const text = clean(value);
  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) return `${iso[3].padStart(2, '0')}.${iso[2].padStart(2, '0')}.${iso[1]}`;
  const slash = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (slash) {
    const year = slash[3].length === 2 ? `20${slash[3]}` : slash[3];
    return `${slash[1].padStart(2, '0')}.${slash[2].padStart(2, '0')}.${year}`;
  }
  return text;
}

function formatTimeValue(value) {
  if (value instanceof Date && !isNaN(value)) {
    return value.toLocaleTimeString('tr-TR', { hour: '2-digit', minute: '2-digit', hour12: false });
  }
  if (typeof value === 'number') {
    const fraction = value >= 1 ? value - Math.floor(value) : value;
    if (fraction >= 0 && fraction < 1) {
      const minutes = Math.round(fraction * 24 * 60) % (24 * 60);
      const h = String(Math.floor(minutes / 60)).padStart(2, '0');
      const m = String(minutes % 60).padStart(2, '0');
      return `${h}:${m}`;
    }
  }
  const text = clean(value);
  const hasNextDay = /next\s*day/i.test(text);
  const match = text.match(/(\d{1,2})[:.](\d{2})/);
  if (match) {
    const time = `${match[1].padStart(2, '0')}:${match[2]}`;
    return hasNextDay ? `${time}Next Day` : time;
  }
  return text;
}

function stripNextDayFromTime(text) {
  return clean(text).replace(/\s*next\s*day\s*/ig, '').trim();
}

function hasNextDayMarker(value, displayValue = '') {
  return /next\s*day/i.test(clean(value)) || /next\s*day/i.test(clean(displayValue));
}

function formatEtaValue(value, displayValue = '', mode = currentMode) {
  const formatted = formatTimeValue(value);
  if (mode === MODE_ARRIVALS) {
    const cleaned = stripNextDayFromTime(formatted);
    if (cleaned) return cleaned;
    return stripNextDayFromTime(displayValue);
  }
  return formatted;
}

function looksLikeDateString(text) {
  return /\bGMT\b|^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\b/i.test(text)
    || /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(text);
}

function stripAgePart(value) {
  const text = String(value ?? '').trim();
  if (!text) return '';
  const stripped = text.replace(/^0+(?=\d)/, '');
  return stripped === '' ? '0' : stripped;
}

function parseChildrenCount(value) {
  const text = clean(value).replace(',', '.');
  if (!text) return null;
  const num = Number(text);
  return Number.isFinite(num) ? Math.max(0, Math.round(num)) : null;
}

function agePartsToText(parts) {
  return parts
    .map(stripAgePart)
    .filter(part => part !== '')
    .join(',');
}

function normalizePlainAgeList(value, childrenCount = null) {
  const count = parseChildrenCount(childrenCount);
  let text = clean(value);
  if (!text) return '';

  // 0.4 / 4.6 / 10.12 gibi Excel'in ondalık sayıya çevirdiği yaş listelerini geri al.
  if (/^\d{1,2}\.\d{1,2}$/.test(text) && count && count >= 2) {
    text = text.replace('.', ',');
  }

  const parts = text
    .replace(/\s*;\s*/g, ',')
    .split(',')
    .map(part => part.trim())
    .filter(Boolean);

  return agePartsToText(parts);
}

function childAgesFromDateParts(day, month, yearText, childrenCount = null) {
  const d = Number(day);
  const m = Number(month);
  const yText = String(yearText ?? '').trim();
  const y = Number(yText);
  const count = parseChildrenCount(childrenCount);
  const fullYear = yText.length <= 2 ? (y >= 10 ? 2000 + y : 1900 + y) : y;

  if (!Number.isFinite(d) || !Number.isFinite(m) || !Number.isFinite(fullYear)) return '';
  if (count === 0) return '';

  const dayIsAge = d >= 0 && d <= 18;
  const monthIsAge = m >= 0 && m <= 18;
  const yearAge = fullYear - 2010;
  const yearCanBeAge = yearAge >= 0 && yearAge <= 18;

  // Excel bazı iki yaş değerlerini 04.01.1900 gibi eski tarih yapabiliyor.
  // 0,4 gibi değerlerde doğru sonuç 0,4 olmalı; 4,1,0 olmamalı.
  if ((fullYear === 1899 || fullYear === 1900 || fullYear === 2000) && m === 1 && dayIsAge) {
    if (count === 1) return agePartsToText([d]);
    return agePartsToText([0, d]);
  }

  // Excel 2,6,9 gibi üçlü yaş listesini Sat Jun 09 2012 / 09.06.2012 yapabiliyor.
  // Yıl 2012 -> ilk yaş 2, ay -> ikinci yaş, gün -> üçüncü yaş.
  if (yearCanBeAge && dayIsAge && monthIsAge) {
    if (count >= 3 || !count) return agePartsToText([yearAge, m, d]);
    if (count === 2) return agePartsToText([m, d]);
    if (count === 1) return agePartsToText([d]);
  }

  // Güncel yıl ile oluşan tarihler genelde Excel'in 7,7 değerini 07.07.2026
  // tarihine çevirmesinden kaynaklanıyor. Children sayısı 2 ise gün,ay olarak al.
  if (dayIsAge && monthIsAge) {
    if (count === 2) return agePartsToText([d, m]);
    if (count === 1) return agePartsToText([d]);
  }

  // Child Ages kolonunda gerçek tarih göstermeyelim; emin değilsek boş bırak.
  return '';
}

function decodeDateLikeChildAgesText(text, childrenCount = null) {
  const value = clean(text);
  if (!value) return '';

  const jsText = childAgesFromJsDateText(value, childrenCount);
  if (jsText) return jsText;

  // 07.07.2026, 7/7/2026, 07-07-26 ve sonuna saat eklenmiş haller.
  const numericDate = value.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/);
  if (numericDate) return childAgesFromDateParts(numericDate[1], numericDate[2], numericDate[3], childrenCount);

  const parsed = new Date(value);
  return !isNaN(parsed) ? inferChildAgesFromDate(parsed, childrenCount) : '';
}

function inferChildAgesFromDate(date, childrenCount = null) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  return childAgesFromDateParts(date.getDate(), date.getMonth() + 1, date.getFullYear(), childrenCount);
}

function childAgesFromJsDateText(text, childrenCount = null) {
  const value = clean(text);
  const match = value.match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+(\d{4})/i);
  if (!match) return '';
  const months = {
    jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
    jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
  };
  return childAgesFromDateParts(match[2], months[match[1].toLowerCase()], match[3], childrenCount);
}

function normalizeAgeListText(text, childrenCount = null) {
  const value = clean(text);
  if (!value) return '';

  if (looksLikeDateString(value)) {
    return decodeDateLikeChildAgesText(value, childrenCount);
  }

  return normalizePlainAgeList(value, childrenCount);
}

function formatChildAgesValue(value, displayValue = '', childrenCount = null) {
  const count = parseChildrenCount(childrenCount);
  if (count === 0) return '';

  const displayRaw = clean(displayValue);
  const raw = clean(value);

  const candidates = [displayRaw, value, raw];
  for (const candidate of candidates) {
    if (candidate instanceof Date && !isNaN(candidate)) {
      const decoded = inferChildAgesFromDate(candidate, count);
      if (decoded) return decoded;
      continue;
    }

    const text = clean(candidate);
    if (!text) continue;

    if (looksLikeDateString(text)) {
      const decoded = decodeDateLikeChildAgesText(text, count);
      if (decoded) return decoded;
      continue;
    }

    if (typeof candidate === 'number') {
      // Child Ages tek tam sayı ise tek yaş olarak kalır. Ondalık değerler yaş listesi olabilir.
      if (Number.isInteger(candidate)) return String(candidate);
      return normalizePlainAgeList(String(candidate), count);
    }

    const normalized = normalizePlainAgeList(text, count);
    if (normalized) return normalized;
  }

  return '';
}


function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function sameDayKey(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const d = startOfDay(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

function formatShortDate(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  return formatDateParts(date.getDate(), date.getMonth() + 1, date.getFullYear());
}

function dateFromParts(day, month, year) {
  const y = Number(String(year).length === 2 ? `20${year}` : year);
  const m = Number(month);
  const d = Number(day);
  if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
  if (m < 1 || m > 12 || d < 1 || d > 31 || y < 1900 || y > 2100) return null;
  const date = new Date(y, m - 1, d, 12, 0, 0, 0);
  return isNaN(date) ? null : startOfDay(date);
}

function parseDateText(text) {
  const value = clean(text);
  if (!value) return null;

  const iso = value.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) return dateFromParts(iso[3], iso[2], iso[1]);

  const match = value.match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (match) return dateFromParts(match[1], match[2], match[3]);

  // Oda numarası gibi düz sayıları tarih sanmasın.
  // JavaScript new Date('1001') değerini tarih kabul edebiliyor; bu DND tablo algısını bozuyordu.
  if (/^\d+$/.test(value)) return null;

  const parsed = new Date(value);
  if (isNaN(parsed)) return null;
  const year = parsed.getFullYear();
  if (year < 1900 || year > 2100) return null;
  return startOfDay(parsed);
}

function parseMatrixDate(value, displayValue = '') {
  // DND / TİST formlarında tarih hücreleri bazen Excel'de mm-dd-yy olarak görünür.
  // Görünen metni önce okursak 07-03-26 değerini 07.03.2026 sanabilir.
  // Bu yüzden gerçek Excel tarih değerini (Date / seri numarası) önce kullanıyoruz.
  if (value instanceof Date && !isNaN(value)) {
    return startOfDay(value);
  }

  if (typeof value === 'number' && value > 20000 && value < 80000) {
    const d = excelSerialToLocalDate(value);
    return d ? startOfDay(d) : null;
  }

  const valueTextDate = parseDateText(value);
  if (valueTextDate) return valueTextDate;

  const displayDate = parseDateText(displayValue);
  if (displayDate) return displayDate;

  return null;
}

function detectDndStatus(value) {
  const text = canonical(value).toUpperCase();
  if (!text) return '';
  if (text === 'D' || text.includes('DND')) return 'DND';
  if (text === 'T' || text.includes('TIST') || text.includes('TİST')) return 'TİST';
  return '';
}

function isWednesday(date) {
  return date instanceof Date && !isNaN(date) && date.getDay() === 3;
}

function realDndStatusForCell(rowIndex, colIndex, rows, displayRows) {
  const raw = rows[rowIndex][colIndex];
  const display = (displayRows[rowIndex] || [])[colIndex];
  return detectDndStatus(raw || display);
}

function dndStatusForCell(rowIndex, colIndex, date, rows, displayRows) {
  const status = realDndStatusForCell(rowIndex, colIndex, rows, displayRows);
  if (status) return status;
  return isWednesday(date) ? 'ÇARŞAMBA' : '';
}

function dateColumnHasDndMark(colIndex, rows, displayRows, startRowIndex = 0) {
  for (let r = startRowIndex; r < rows.length; r += 1) {
    if (realDndStatusForCell(r, colIndex, rows, displayRows)) return true;
  }
  return false;
}

function findRowDndStartIndex(rowIndex, activeDateCols, rows, displayRows, arrivalLimit = null) {
  if (!activeDateCols.length) return -1;
  const firstDate = activeDateCols[0].date;
  if (arrivalLimit && firstDate < arrivalLimit) return -1;

  // Eski ana kural korunur: DND / TİST dünden geriye doğru kontrol edilir.
  // Burada sadece başlangıç/bitiş tarihinin yanlış görünmesine sebep olan tarih okuma kısmı düzeltilmiştir.
  return 0;
}

function canWednesdayBridgeDnd(rowIndex, activeDateCols, index, rows, displayRows, streak) {
  const current = activeDateCols[index];
  return Boolean(current && isWednesday(current.date));
}

function processDndWorkbook(workbook) {
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: true, defval: '' });
  const displayRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false, defval: '' });
  if (!rows.length) throw new Error('Excel sayfası boş görünüyor.');

  let dateRowIndex = 0;
  let dateCols = [];

  rows.slice(0, 15).forEach((row, rowIndex) => {
    const cols = row
      .map((cell, colIndex) => ({ colIndex, date: parseMatrixDate(cell, (displayRows[rowIndex] || [])[colIndex]) }))
      .filter(item => item.date);
    if (cols.length > dateCols.length) {
      dateCols = cols;
      dateRowIndex = rowIndex;
    }
  });

  if (dateCols.length < 2) {
    throw new Error('Tarih satırı bulunamadı. İlk satırda 30.06.2026 gibi tarih başlıkları olmalı.');
  }

  const dateColIndexes = new Set(dateCols.map(item => item.colIndex));
  let roomColIndex = -1;
  let bestRoomCount = 0;
  const maxCols = Math.max(...rows.map(row => row.length));

  for (let col = 0; col < maxCols; col += 1) {
    if (dateColIndexes.has(col)) continue;
    let count = 0;
    for (let r = dateRowIndex + 1; r < rows.length; r += 1) {
      const text = clean((displayRows[r] || [])[col] || rows[r][col]);
      if (/^\d{3,5}$/.test(text.replace(/\.0$/, ''))) count += 1;
    }
    if (count > bestRoomCount) {
      bestRoomCount = count;
      roomColIndex = col;
    }
  }

  if (roomColIndex === -1 || bestRoomCount === 0) {
    throw new Error('Oda numarası sütunu bulunamadı. İlk sütunda 1001, 1002 gibi oda numaraları olmalı.');
  }

  const yesterday = startOfDay(new Date());
  yesterday.setDate(yesterday.getDate() - 1);

  const sortedDateCols = dateCols.slice().sort((a, b) => b.date - a.date);
  const pastDateCols = sortedDateCols.filter(item => item.date <= yesterday);

  // Ana kural: dünden geriye doğru bakılır.
  // Ancak formda ilerideki boş tarih kolonları varsa veya bilgisayar tarihi formun tarih aralığına denk gelmiyorsa
  // en son gerçekten D/T işaretli günü başlangıç kabul ediyoruz. Böylece boş gelecek kolon yüzünden sonuçlar kaybolmaz.
  const markedDateCols = sortedDateCols.filter(item => dateColumnHasDndMark(item.colIndex, rows, displayRows, dateRowIndex + 1));
  const latestMarkedDate = markedDateCols[0]?.date || null;
  const latestPastCol = pastDateCols[0] || null;
  const latestPastHasMark = latestPastCol ? dateColumnHasDndMark(latestPastCol.colIndex, rows, displayRows, dateRowIndex + 1) : false;

  let controlStartDate = null;
  if (latestPastCol && latestPastHasMark) {
    controlStartDate = latestPastCol.date;
  } else if (latestMarkedDate && (!latestPastCol || latestMarkedDate <= latestPastCol.date)) {
    controlStartDate = latestMarkedDate;
  } else if (latestPastCol) {
    controlStartDate = latestPastCol.date;
  } else {
    controlStartDate = sortedDateCols[0]?.date || null;
  }

  const controlDateCols = controlStartDate
    ? sortedDateCols.filter(item => item.date <= controlStartDate)
    : [];
  const latestControlDate = controlDateCols[0]?.date;
  const oldestControlDate = controlDateCols[controlDateCols.length - 1]?.date;
  dndDateWindowText = controlDateCols.length
    ? `${formatShortDate(latestControlDate)} tarihinden geriye doğru / Çarşamba tüm odalar işaretli sayılır`
    : 'Tarih bulunamadı / Çarşamba tüm odalar işaretli sayılır';

  const results = [];
  const hasCurrentRoomFilter = currentRoomFilter.size > 0;
  dndFilterStats = {
    active: hasCurrentRoomFilter,
    currentRooms: currentRoomFilter.size,
    skippedOldRooms: 0,
    stoppedBeforeArrival: 0,
  };

  for (let r = dateRowIndex + 1; r < rows.length; r += 1) {
    const rawRoom = normalizeRoomId((displayRows[r] || [])[roomColIndex] || rows[r][roomColIndex]);
    if (!rawRoom) continue;

    const currentInfo = currentRoomFilter.get(rawRoom);
    if (hasCurrentRoomFilter && !currentInfo) {
      dndFilterStats.skippedOldRooms += 1;
      continue;
    }

    const arrivalLimit = currentInfo?.arrivalDate ? startOfDay(currentInfo.arrivalDate) : null;
    const streak = [];
    const startIndex = findRowDndStartIndex(r, controlDateCols, rows, displayRows, arrivalLimit);

    if (startIndex === -1) continue;

    for (let c = startIndex; c < controlDateCols.length; c += 1) {
      const { colIndex, date } = controlDateCols[c];
      if (arrivalLimit && date < arrivalLimit) {
        dndFilterStats.stoppedBeforeArrival += 1;
        break;
      }

      const realStatus = realDndStatusForCell(r, colIndex, rows, displayRows);
      const status = realStatus || (canWednesdayBridgeDnd(r, controlDateCols, c, rows, displayRows, streak) ? 'ÇARŞAMBA' : '');
      if (!status) break;

      if (streak.length) {
        const prev = streak[streak.length - 1].date;
        const diffDays = Math.round((prev - date) / 86400000);
        if (diffDays !== 1) break;
      }

      streak.push({ date, status, real: Boolean(realStatus) });
    }

    if (streak.length >= 2) {
      const sortedAsc = streak.slice().sort((a, b) => a.date - b.date);
      const realSortedAsc = sortedAsc.filter(item => item.real || item.status !== 'ÇARŞAMBA');
      const startItem = realSortedAsc[0] || sortedAsc[0];
      const endItem = realSortedAsc[realSortedAsc.length - 1] || sortedAsc[sortedAsc.length - 1];
      const daysWithoutWednesday = streak.filter(item => item.status !== 'ÇARŞAMBA').length;
      results.push({
        room: rawRoom,
        start: startItem.date,
        end: endItem.date,
        daysWithoutWednesday,
        days: streak.length,
        details: sortedAsc.map(item => `${formatShortDate(item.date)} ${item.status}`).join(' / '),
      });
    }
  }

  if (hasCurrentRoomFilter) {
    dndDateWindowText += ` / Güncel oda filtresi: ${currentRoomFilter.size} oda`;
  }

  results.sort((a, b) => {
    const aNoWed = a.daysWithoutWednesday ?? a.days;
    const bNoWed = b.daysWithoutWednesday ?? b.days;
    return bNoWed - aNoWed
      || b.days - a.days
      || roomSortValue(a.room) - roomSortValue(b.room)
      || a.start - b.start;
  });
  return results;
}

function renderDndPreview(results) {
  preview.classList.remove('empty');
  preview.classList.add('preview-dnd');
  preview.classList.remove('preview-arrivals', 'preview-departures');
  preview.innerHTML = '';

  const page = document.createElement('article');
  page.className = 'sheet-page dnd-page';

  const rowsHtml = results.map((item, index) => `
    <tr>
      <td class="idx">${index + 1}</td>
      <td class="room">${escapeHtml(item.room)}</td>
      <td>${escapeHtml(item.daysWithoutWednesday ?? item.days)}</td>
      <td>${escapeHtml(item.days)}</td>
      <td class="dnd-detail">${escapeHtml(item.details)}</td>
    </tr>`).join('');

  page.innerHTML = `
    <div class="dnd-title">DND / TİST ARKA ARKAYA ODALAR</div>
    <div class="dnd-range">Kontrol: ${escapeHtml(dndDateWindowText || 'Tüm tarihler')}</div>
    <div class="table-wrap">
      <table class="departure-table dnd-table">
        <colgroup>
          <col class="idx"><col class="room"><col class="small"><col class="small"><col class="notes">
        </colgroup>
        <thead>
          <tr><th></th><th>Room</th><th>Çarşamba Hariç Gün</th><th>Gün</th><th>Detay</th></tr>
        </thead>
        <tbody>${rowsHtml || '<tr><td colspan="5">Arka arkaya DND / TİST oda bulunamadı.</td></tr>'}</tbody>
      </table>
    </div>`;

  preview.appendChild(page);
}

function renderDndSummary(results) {
  summary.hidden = false;
  const roomCount = new Set(results.map(item => item.room)).size;
  const maxDays = results.reduce((max, item) => Math.max(max, item.days), 0);
  const maxDaysNoWed = results.reduce((max, item) => Math.max(max, item.daysWithoutWednesday ?? item.days), 0);
  const filterCard = dndFilterStats.active
    ? `<div class="summary-card"><strong>${dndFilterStats.currentRooms}</strong><span>Güncel oda filtresi</span><small>${dndFilterStats.skippedOldRooms} eski oda atlandı</small></div>`
    : '';

  summary.innerHTML = `
    <div class="summary-card summary-total"><strong>${results.length}</strong><span>Arka arkaya kayıt</span></div>
    <div class="summary-card"><strong>${roomCount}</strong><span>Oda</span></div>
    <div class="summary-card"><strong>${maxDaysNoWed}</strong><span>Çarşamba hariç en uzun</span></div>
    <div class="summary-card"><strong>${maxDays}</strong><span>En uzun gün</span></div>
    ${filterCard}`;
}

function renderDndOutput(message = '') {
  assignmentPanel.hidden = true;
  greenPanel.hidden = true;
  updateCurrentRoomsPanel();
  chiefControls.innerHTML = '';
  assignmentStatus.innerHTML = '';
  renderDndPreview(dndResults);
  renderDndSummary(dndResults);
  setButtons({ printable: true, clearable: Boolean(lastWorkbook) });
  setStatus(message || `${lastFileName} DND / TİST olarak işlendi.`, 'ok');
}

function findHeaderRow(rows, fields = requiredFields()) {
  let best = { index: 0, score: -1 };
  rows.slice(0, 25).forEach((row, index) => {
    const joined = row.map(canonical);
    let score = 0;
    fields.forEach(field => {
      if (joined.some(cell => field.aliases.map(canonical).includes(cell))) score++;
    });
    if (score > best.score) best = { index, score };
  });
  return best.score >= 3 ? best.index : 0;
}

function mapHeaders(headerRow, fields = requiredFields()) {
  const map = {};
  const normalizedCells = headerRow.map(canonical);

  fields.forEach(field => {
    const normalizedAliases = field.aliases.map(canonical);
    let index = normalizedCells.findIndex(cell => normalizedAliases.includes(cell));

    if (index === -1) {
      index = normalizedCells.findIndex(cell => normalizedAliases.some(alias => cell.includes(alias) || alias.includes(cell)));
    }

    if (index !== -1) map[field.key] = index;
  });

  return map;
}

function fileNameModeHint(fileName = '', workbook = null) {
  const names = [fileName, ...(workbook?.SheetNames || [])].join(' ');
  const text = canonical(names);

  if (/\b(vacant|hkvacroom|vac|bos|boş)\b/.test(text)) return { mode: MODE_VACANT, confidence: 0.98, reason: 'dosya adında Vacant var' };
  if (/\b(dnd|tist|temizlik)\b/.test(text)) return { mode: MODE_DND, confidence: 0.98, reason: 'dosya adında DND / TİST var' };
  if (/\b(arrival|arrivals|arrivas|giris|gelis)\b/.test(text)) return { mode: MODE_ARRIVALS, confidence: 0.98, reason: 'dosya adında Arrivals var' };
  if (/\b(departure|departures|departus|depertur|cikis|ayrilis)\b/.test(text)) return { mode: MODE_DEPARTURES, confidence: 0.98, reason: 'dosya adında Departures var' };

  return null;
}

function inspectWorkbookShape(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
  const displayRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '', dateNF: 'dd.mm.yyyy' });
  if (!rows.length) return { rows, displayRows, dndMatrix: false, dateCols: [] };

  let bestDateCols = [];
  rows.slice(0, 15).forEach((row, rowIndex) => {
    const cols = row
      .map((cell, colIndex) => ({ colIndex, date: parseMatrixDate(cell, (displayRows[rowIndex] || [])[colIndex]) }))
      .filter(item => item.date);
    if (cols.length > bestDateCols.length) bestDateCols = cols;
  });

  const dateColIndexes = new Set(bestDateCols.map(item => item.colIndex));
  let bestRoomCount = 0;
  const maxCols = Math.max(0, ...rows.map(row => row.length));
  for (let col = 0; col < maxCols; col += 1) {
    if (dateColIndexes.has(col)) continue;
    let count = 0;
    for (let r = 0; r < rows.length; r += 1) {
      const room = normalizeRoomId((displayRows[r] || [])[col] || rows[r][col]);
      if (room) count += 1;
    }
    if (count > bestRoomCount) bestRoomCount = count;
  }

  return {
    rows,
    displayRows,
    dndMatrix: bestDateCols.length >= 5 && bestRoomCount >= 5,
    dateCols: bestDateCols,
  };
}

function mostCommonRatio(values) {
  const list = values.filter(Boolean);
  if (!list.length) return 0;
  const counts = new Map();
  list.forEach(value => counts.set(value, (counts.get(value) || 0) + 1));
  return Math.max(...counts.values()) / list.length;
}

function detectWorkbookMode(workbook, fileName = '') {
  const nameHint = fileNameModeHint(fileName, workbook);
  if (nameHint) return nameHint;

  const shape = inspectWorkbookShape(workbook);
  if (shape.dndMatrix) return { mode: MODE_DND, confidence: 0.9, reason: 'ilk satırda çoklu tarih ve ilk sütunda oda listesi var' };

  const allFields = FIELD_DEFS;
  const headerIndex = findHeaderRow(shape.displayRows, allFields);
  const headerMap = mapHeaders(shape.displayRows[headerIndex] || [], allFields);
  if (headerMap.room === undefined || headerMap.arrival === undefined || headerMap.departure === undefined) return null;

  const records = shape.rows.slice(headerIndex + 1, headerIndex + 101);
  const displayRecords = shape.displayRows.slice(headerIndex + 1, headerIndex + 101);
  const arrivalDates = [];
  const departureDates = [];
  let etdLateCount = 0;
  let etdCount = 0;

  records.forEach((row, index) => {
    const displayRow = displayRecords[index] || [];
    const room = normalizeRoomId(displayRow[headerMap.room] || row[headerMap.room]);
    if (!room) return;

    const arrivalDate = parseMatrixDate(row[headerMap.arrival], displayRow[headerMap.arrival]);
    const departureDate = parseMatrixDate(row[headerMap.departure], displayRow[headerMap.departure]);
    if (arrivalDate) arrivalDates.push(sameDayKey(arrivalDate));
    if (departureDate) departureDates.push(sameDayKey(departureDate));

    if (headerMap.etd !== undefined) {
      const etd = formatTimeValue(row[headerMap.etd] || displayRow[headerMap.etd]);
      if (etd) {
        etdCount += 1;
        if (etd === ETD_HIGHLIGHT) etdLateCount += 1;
      }
    }
  });

  const arrivalCluster = mostCommonRatio(arrivalDates);
  const departureCluster = mostCommonRatio(departureDates);
  const lateRatio = etdCount ? etdLateCount / etdCount : 0;

  // Departures dosyalarında ETD 17:00 çoğunlukla sabit/yüksek olur. Arrivals dosyalarında saatler daha dağınıktır.
  if (lateRatio >= 0.45) return { mode: MODE_DEPARTURES, confidence: 0.78, reason: 'ETD 17:00 oranı yüksek' };
  if (arrivalCluster >= 0.7 && arrivalCluster > departureCluster + 0.2) return { mode: MODE_ARRIVALS, confidence: 0.72, reason: 'Arrival tarihi tek güne yoğunlaşıyor' };
  if (departureCluster >= 0.7 && departureCluster > arrivalCluster + 0.2) return { mode: MODE_DEPARTURES, confidence: 0.72, reason: 'Departure tarihi tek güne yoğunlaşıyor' };

  return null;
}

function wrongFileWarning(workbook, fileName = '', selectedMode = currentMode) {
  const detected = detectWorkbookMode(workbook, fileName);
  if (!detected || detected.mode === selectedMode || detected.confidence < 0.7) return '';

  return `${modeLabel(selectedMode)} seçili ama yüklenen dosya ${modeLabel(detected.mode)} gibi görünüyor. Lütfen doğru bölümü seçip dosyayı tekrar yükle.`;
}

function roomNumber(roomValue) {
  const match = clean(roomValue).match(/\d+/);
  return match ? Number(match[0]) : NaN;
}

function roomGroup(roomValue) {
  const num = roomNumber(roomValue);
  if (!Number.isFinite(num)) return 'Diğer';
  const group = Math.floor(num / 1000) * 1000;
  if (group >= 1000 && group <= 5000) return `${group}ler`;
  return 'Diğer';
}

function hundredSection(roomValue) {
  const num = roomNumber(roomValue);
  if (!Number.isFinite(num)) return 'Diğer';
  const section = Math.floor(num / 100) * 100;
  return `${section}ler`;
}

function roomSortValue(roomValue) {
  const num = roomNumber(roomValue);
  return Number.isFinite(num) ? num : Number.MAX_SAFE_INTEGER;
}

function normalizeRoomId(roomValue) {
  const text = clean(roomValue).replace(/\.0$/, '');
  const exact = text.match(/^\d{3,5}$/);
  if (exact) return exact[0];
  const match = text.match(/\b\d{3,5}\b/);
  return match ? match[0] : '';
}

function chooseCurrentArrivalDate(oldDate, newDate) {
  if (!(newDate instanceof Date) || isNaN(newDate)) return oldDate || null;
  if (!(oldDate instanceof Date) || isNaN(oldDate)) return startOfDay(newDate);
  // Aynı oda birden fazla güncel listede varsa en yeni arrival tarihini al.
  // Böylece eski konaklamanın DND/TİST kayıtları yanlış sayılmaz.
  return newDate > oldDate ? startOfDay(newDate) : oldDate;
}

function mergeCurrentRoomRecord(targetMap, room, arrivalDate, source) {
  const normalizedRoom = normalizeRoomId(room);
  if (!normalizedRoom) return false;

  const existing = targetMap.get(normalizedRoom) || { room: normalizedRoom, arrivalDate: null, sources: new Set() };
  existing.arrivalDate = chooseCurrentArrivalDate(existing.arrivalDate, arrivalDate);
  if (!existing.sources) existing.sources = new Set();
  if (source) existing.sources.add(source);
  targetMap.set(normalizedRoom, existing);
  return true;
}

function findCurrentRoomHeader(rows) {
  const roomField = FIELD_DEFS.find(field => field.key === 'room');
  const arrivalField = FIELD_DEFS.find(field => field.key === 'arrival');
  let best = { index: -1, roomCol: -1, arrivalCol: -1, score: -1 };

  rows.slice(0, 25).forEach((row, index) => {
    const normalizedCells = row.map(canonical);
    const roomAliases = roomField.aliases.map(canonical);
    const arrivalAliases = arrivalField.aliases.map(canonical);

    let roomCol = normalizedCells.findIndex(cell => roomAliases.includes(cell) || roomAliases.some(alias => cell.includes(alias) || alias.includes(cell)));
    let arrivalCol = normalizedCells.findIndex(cell => arrivalAliases.includes(cell) || arrivalAliases.some(alias => cell.includes(alias) || alias.includes(cell)));
    let score = 0;
    if (roomCol !== -1) score += 4;
    if (arrivalCol !== -1) score += 2;

    if (score > best.score) best = { index, roomCol, arrivalCol, score };
  });

  return best.roomCol !== -1 ? best : null;
}

function findBestRoomColumn(rows, displayRows, startRow = 0) {
  const maxCols = Math.max(0, ...rows.map(row => row.length));
  let best = { col: -1, count: 0 };

  for (let col = 0; col < maxCols; col += 1) {
    let count = 0;
    for (let r = startRow; r < rows.length; r += 1) {
      const room = normalizeRoomId((displayRows[r] || [])[col] || rows[r][col]);
      if (room) count += 1;
    }
    if (count > best.count) best = { col, count };
  }

  return best.count ? best.col : -1;
}

function extractCurrentRoomMapFromWorkbook(workbook, sourceName = '') {
  const map = new Map();

  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: '' });
    const displayRows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '', dateNF: 'dd.mm.yyyy' });
    if (!rows.length) return;

    const header = findCurrentRoomHeader(displayRows);
    let roomCol = header ? header.roomCol : -1;
    let arrivalCol = header ? header.arrivalCol : -1;
    let startRow = header ? header.index + 1 : 0;

    if (roomCol === -1) {
      roomCol = findBestRoomColumn(rows, displayRows, 0);
      startRow = 0;
    }

    if (roomCol === -1) return;

    for (let r = startRow; r < rows.length; r += 1) {
      const displayRow = displayRows[r] || [];
      const row = rows[r] || [];
      const room = normalizeRoomId(displayRow[roomCol] || row[roomCol]);
      if (!room) continue;

      let arrivalDate = null;
      if (arrivalCol !== -1) {
        arrivalDate = parseMatrixDate(row[arrivalCol], displayRow[arrivalCol]);
      }

      mergeCurrentRoomRecord(map, room, arrivalDate, sourceName);
    }
  });

  return map;
}

function updateCurrentRoomsStatus() {
  if (!currentRoomsStatus) return;
  if (!currentRoomFilter.size) {
    currentRoomsStatus.textContent = 'Güncel oda listesi yüklenmedi. Yüklenirse DND / TİST sadece o odalara göre hesaplanır.';
    currentRoomsStatus.className = 'current-rooms-status';
    return;
  }

  const withArrivalDate = [...currentRoomFilter.values()].filter(item => item.arrivalDate).length;
  const sourceText = currentRoomFileNames.length ? ` / ${currentRoomFileNames.join(', ')}` : '';
  currentRoomsStatus.textContent = `${currentRoomFilter.size} güncel oda yüklendi. ${withArrivalDate} odada Arrival tarihi bulundu${sourceText}`;
  currentRoomsStatus.className = 'current-rooms-status ok';
}

function updateCurrentRoomsPanel() {
  if (!currentRoomsPanel) return;
  currentRoomsPanel.hidden = currentMode !== MODE_DND;
  if (currentRoomsInput) currentRoomsInput.disabled = currentMode !== MODE_DND;
  updateCurrentRoomsStatus();
}

function groupSortValue(groupName) {
  if (groupName === 'Diğer') return 999999;
  if (groupName === 'Ofis') return 0;
  const parsed = parseInt(groupName, 10);
  return Number.isFinite(parsed) ? parsed : 999998;
}

function normalizeAgent(agent) {
  const value = clean(agent);
  return value.replace(/\s*Limited$/i, ' Limited');
}


function ensurePdfLibrary() {
  if (typeof pdfjsLib === 'undefined') {
    throw new Error('PDF okuyucu yüklenemedi. İnternet bağlantısını kontrol edip sayfayı yenile.');
  }
  if (pdfjsLib.GlobalWorkerOptions && !pdfjsLib.GlobalWorkerOptions.workerSrc) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  }
}

function vacantLineClusters(items, yTolerance = 3) {
  const lines = [];
  items
    .filter(item => clean(item.text))
    .sort((a, b) => (b.y - a.y) || (a.x - b.x))
    .forEach(item => {
      let line = lines.find(candidate => Math.abs(candidate.y - item.y) <= yTolerance);
      if (!line) {
        line = { y: item.y, items: [] };
        lines.push(line);
      }
      line.items.push(item);
      line.y = (line.y * (line.items.length - 1) + item.y) / line.items.length;
    });

  return lines
    .sort((a, b) => b.y - a.y)
    .map(line => line.items
      .sort((a, b) => a.x - b.x)
      .map(item => clean(item.text))
      .filter(Boolean)
      .join(' ')
      .replace(/\s+([,.;:])/g, '$1')
      .replace(/\s+/g, ' ')
      .trim())
    .filter(Boolean);
}

function vacantColumnText(rowItems, minX, maxX, { separator = ' / ' } = {}) {
  const lines = vacantLineClusters(rowItems.filter(item => item.x >= minX && item.x < maxX));
  return lines.join(separator).replace(/\s+/g, ' ').trim();
}

function normalizeVacantStatus(value) {
  return clean(value).replace(/\bDue\s+Out\b/gi, 'Due Out').replace(/\bDue\s+In\b/gi, 'Due In').replace(/\bChecked\s+In\b/gi, 'Checked In');
}

function isVacantPdfArtifactText(value) {
  const text = clean(value);
  if (!text) return true;
  return /^page\s+\d+\s+of\s+\d+$/i.test(text)
    || /^hkvacroom$/i.test(text)
    || /^roomclassall$/i.test(text);
}

function stripVacantPdfArtifacts(value, { roomClass = false } = {}) {
  let text = clean(value);
  if (!text) return '';

  // HK Vacant PDF'lerinde bazı sayfalarda footer/kaynak yazıları tablo satırıyla
  // aynı hizaya düşebiliyor. Bunlar tarih/Next Blocked/Class hücrelerine karışmasın.
  text = text
    .replace(/\bPage\s+\d+\s+of\s+\d+\b/gi, '')
    .replace(/\bhkvacroom\b/gi, '');

  if (roomClass) {
    text = text.replace(/[_\s-]*RoomClassAll\b/gi, '');
  }

  return text
    .replace(/\s*\/\s*(?=\/|$)/g, '')
    .replace(/(^|\s)\/\s*$/g, '')
    .replace(/^\s*\/\s*/g, '')
    .replace(/\s*\/\s*/g, ' / ')
    .replace(/\s{2,}/g, ' ')
    .trim();
}

function parseVacantPageItems(textContent) {
  const pageItems = (textContent.items || [])
    .map(item => {
      const transform = item.transform || [];
      return {
        text: clean(item.str),
        x: Number(transform[4] || 0),
        y: Number(transform[5] || 0),
        width: Number(item.width || transform[0] || 0),
      };
    })
    .filter(item => item.text)
    .filter(item => !isVacantPdfArtifactText(item.text));

  const roomAnchors = pageItems
    .filter(item => item.x >= 25 && item.x <= 60 && /^\d{3,5}$/.test(item.text.replace(/\.0$/, '')))
    .map(item => ({ ...item, room: normalizeRoomId(item.text), rowItems: [] }))
    .filter(item => item.room);

  if (!roomAnchors.length) return [];

  const tableItems = pageItems.filter(item => item.x >= 25 && item.x <= 560);
  tableItems.forEach(item => {
    let best = null;
    let bestDistance = Infinity;
    roomAnchors.forEach(anchor => {
      const distance = Math.abs(anchor.y - item.y);
      if (distance < bestDistance) {
        best = anchor;
        bestDistance = distance;
      }
    });
    // pdf.js koordinatlarında Y aşağı indikçe küçülür.
    // Başlık satırını ilk odaya karıştırmamak için sadece oda satırıyla aynı hizada
    // veya onun altında kalan devam satırlarını alıyoruz.
    const verticalDelta = best ? best.y - item.y : Infinity;
    if (best && verticalDelta >= -3 && verticalDelta <= 22) best.rowItems.push(item);
  });

  return roomAnchors.map(anchor => {
    const rowItems = anchor.rowItems;
    const roomClass = stripVacantPdfArtifacts(vacantColumnText(rowItems, 58, 90, { separator: ' ' }), { roomClass: true });
    const occupancy = stripVacantPdfArtifacts(vacantColumnText(rowItems, 90, 116, { separator: ' / ' }));
    const foStatus = stripVacantPdfArtifacts(vacantColumnText(rowItems, 116, 140, { separator: ' / ' }));
    const nightsVacant = stripVacantPdfArtifacts(vacantColumnText(rowItems, 140, 177, { separator: ' / ' }));
    const name = stripVacantPdfArtifacts(vacantColumnText(rowItems, 177, 278, { separator: ' / ' }));
    const arrival = stripVacantPdfArtifacts(vacantColumnText(rowItems, 278, 318, { separator: ' / ' }));
    const departure = stripVacantPdfArtifacts(vacantColumnText(rowItems, 318, 358, { separator: ' / ' }));
    const reservationStatus = normalizeVacantStatus(stripVacantPdfArtifacts(vacantColumnText(rowItems, 358, 420, { separator: ' / ' })));
    const adults = stripVacantPdfArtifacts(vacantColumnText(rowItems, 420, 448, { separator: ' / ' }));
    const children = stripVacantPdfArtifacts(vacantColumnText(rowItems, 448, 475, { separator: ' / ' }));
    const discrepantStatus = stripVacantPdfArtifacts(vacantColumnText(rowItems, 475, 515, { separator: ' / ' }));
    const nextBlocked = stripVacantPdfArtifacts(vacantColumnText(rowItems, 515, 560, { separator: ' / ' }));

    return {
      room: anchor.room,
      roomClass: clean(roomClass).replace(/\s+/g, ''),
      roomType: occupancy,
      foStatus,
      nightsVacant,
      name,
      arrival,
      departure,
      reservationStatus,
      adults,
      children,
      discrepantStatus,
      nextBlocked,
      notes: '',
    };
  }).filter(record => record.room);
}

async function readVacantPdfFile(file) {
  if (!/\.pdf$/i.test(file.name || '')) {
    throw new Error('Vacant Rooms için PDF dosyası yükle.');
  }
  ensurePdfLibrary();
  const buffer = await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data: buffer });
  const pdf = await loadingTask.promise;
  const records = [];

  for (let pageNo = 1; pageNo <= pdf.numPages; pageNo += 1) {
    const page = await pdf.getPage(pageNo);
    const textContent = await page.getTextContent();
    records.push(...parseVacantPageItems(textContent));
  }

  if (!records.length) throw new Error('Vacant Rooms PDF içinde oda satırı bulunamadı.');
  records.sort((a, b) => roomSortValue(a.room) - roomSortValue(b.room));
  return { name: file.name, records, type: MODE_VACANT };
}

function mergeVacantRecordItems(items) {
  const allRecords = [];
  items.forEach(item => {
    (item.records || []).forEach(record => allRecords.push(record));
  });
  if (!allRecords.length) throw new Error('Vacant Rooms PDF içinde oda satırı bulunamadı.');
  allRecords.sort((a, b) => roomSortValue(a.room) - roomSortValue(b.room));
  const groups = new Map();
  allRecords.forEach(record => {
    const groupName = roomGroup(record.room);
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(record);
  });
  return new Map([...groups.entries()].sort((a, b) => groupSortValue(a[0]) - groupSortValue(b[0])));
}

function rowToRecord(row, map, displayRow = [], mode = currentMode) {
  const children = clean(row[map.children]);
  const etaNextDay = mode === MODE_ARRIVALS && hasNextDayMarker(row[map.eta], displayRow[map.eta]);
  return {
    room: clean(row[map.room]),
    eta: formatEtaValue(row[map.eta], displayRow[map.eta], mode),
    etaNextDay,
    arrival: formatDateValue(row[map.arrival]),
    adults: clean(row[map.adults]),
    children,
    childAges: formatChildAgesValue(row[map.childAges], displayRow[map.childAges], children),
    departure: formatDateValue(row[map.departure]),
    etd: formatTimeValue(row[map.etd]),
    name: clean(row[map.name]),
    travelAgent: normalizeAgent(row[map.travelAgent]),
    notes: '',
  };
}


function formatLoadedFileNames(filesOrItems = []) {
  const names = [...filesOrItems].map(item => item?.name || item?.fileName || '').filter(Boolean);
  if (!names.length) return '';
  if (names.length <= 2) return names.join(', ');
  return `${names.length} Excel dosyası`;
}

function mergeWorkbookGroups(workbookItems, mode = currentMode) {
  const items = [...(workbookItems || [])].filter(Boolean);
  if (!items.length) throw new Error('Excel dosyası bulunamadı.');

  const allRecords = [];
  items.forEach(item => {
    const workbook = item.workbook || item;
    const groups = processWorkbook(workbook, mode);
    groups.forEach(records => allRecords.push(...records));
  });

  if (!allRecords.length) throw new Error('Oda numarası olan satır bulunamadı.');

  allRecords.sort((a, b) => roomSortValue(a.room) - roomSortValue(b.room));
  const groups = new Map();
  allRecords.forEach(record => {
    const groupName = roomGroup(record.room);
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(record);
  });

  return new Map([...groups.entries()].sort((a, b) => groupSortValue(a[0]) - groupSortValue(b[0])));
}

function processWorkbook(workbook, mode = currentMode) {
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: true, defval: '' });
  const displayRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false, defval: '', dateNF: 'dd.mm.yyyy' });
  if (!rows.length) throw new Error('Excel sayfası boş görünüyor.');

  const fields = requiredFields(mode);
  const headerIndex = findHeaderRow(rows, fields);
  const headerMap = mapHeaders(rows[headerIndex], fields);
  const missing = fields.filter(field => headerMap[field.key] === undefined).map(field => field.out);
  if (missing.length) {
    throw new Error(`Eksik kolon bulundu: ${missing.join(', ')}. ${modeLabel(mode)} için başlık satırında bu kolonlar olmalı.`);
  }

  const records = rows.slice(headerIndex + 1)
    .map((row, index) => rowToRecord(row, headerMap, displayRows[headerIndex + 1 + index] || [], mode))
    .filter(record => record.room && /^\d+/.test(record.room));

  if (!records.length) throw new Error('Oda numarası olan satır bulunamadı.');

  records.sort((a, b) => roomSortValue(a.room) - roomSortValue(b.room));

  const groups = new Map();
  records.forEach(record => {
    const groupName = roomGroup(record.room);
    if (!groups.has(groupName)) groups.set(groupName, []);
    groups.get(groupName).push(record);
  });

  return new Map([...groups.entries()].sort((a, b) => groupSortValue(a[0]) - groupSortValue(b[0])));
}

function chunkArray(items, size) {
  const chunks = [];
  for (let i = 0; i < items.length; i += size) chunks.push(items.slice(i, i + size));
  return chunks;
}

function printLineCount(value, charsPerLine) {
  const text = clean(value);
  if (!text) return 1;
  return Math.max(1, Math.ceil(text.length / charsPerLine));
}

function estimatePrintUnits(record) {
  if (currentMode === MODE_ARRIVALS) {
    const maxLines = Math.max(
      printLineCount(record.name, 12),
      printLineCount(record.travelAgent, 9),
      printLineCount(record.notes, 34),
      printLineCount(record.childAges, 8)
    );
    return Math.min(4, Math.max(1, maxLines));
  }

  if (currentMode === MODE_DEPARTURES) {
    const maxLines = Math.max(
      printLineCount(record.travelAgent, 12),
      printLineCount(record.notes, 38),
      printLineCount(record.childAges, 9)
    );
    return Math.min(4, Math.max(1, maxLines));
  }

  if (currentMode === MODE_VACANT) {
    const maxLines = Math.max(
      printLineCount(record.name, 18),
      printLineCount(record.reservationStatus, 12),
      printLineCount(record.nextBlocked, 10)
    );
    return Math.min(4, Math.max(1, maxLines));
  }

  return 1;
}

function splitNumberedRecordsForPrint(numberedRecords) {
  const limits = currentMode === MODE_ARRIVALS
    ? PRINT_PAGE_LIMITS.arrivals
    : currentMode === MODE_VACANT
      ? PRINT_PAGE_LIMITS.vacant
      : PRINT_PAGE_LIMITS.departures;

  const pages = [];
  let page = [];
  let units = 0;

  numberedRecords.forEach(item => {
    const itemUnits = estimatePrintUnits(item.record);
    const mustBreak = page.length > 0 && (
      page.length >= limits.hardMax ||
      units + itemUnits > limits.unitBudget
    );

    if (mustBreak) {
      pages.push(page);
      page = [];
      units = 0;
    }

    page.push(item);
    units += itemUnits;
  });

  if (page.length) pages.push(page);
  return pages;
}

function getRecords(groupName) {
  return originalGroups.get(groupName) || [];
}

function allOriginalRecords() {
  return [...originalGroups.values()]
    .flat()
    .sort((a, b) => roomSortValue(a.room) - roomSortValue(b.room));
}

function buildOfficeGroups() {
  const records = allOriginalRecords();
  return records.length ? new Map([['Ofis', records]]) : new Map();
}

function countLate(records) {
  return (records || []).filter(record => record.etd === ETD_HIGHLIGHT).length;
}

function lateText(records) {
  if (currentMode !== MODE_DEPARTURES) return '';
  return `Total Late: ${countLate(records)}`;
}

function updateGreenRooms() {
  if (currentMode !== MODE_ARRIVALS) {
    greenRooms = new Set();
    return;
  }

  const tokens = String(greenRoomsInput?.value || '')
    .split(/[\s,;]+/)
    .map(token => token.replace(/\D/g, ''))
    .filter(Boolean);
  greenRooms = new Set(tokens);
}

function isGreenRoom(record) {
  if (currentMode !== MODE_ARRIVALS) return false;

  const raw = clean(record.room).replace(/\D/g, '');
  const numeric = roomNumber(record.room);
  return (raw && greenRooms.has(raw)) || (Number.isFinite(numeric) && greenRooms.has(String(numeric)));
}

function getLeaveSections() {
  const sections = [];

  [...leaveGroups].sort((a, b) => groupSortValue(a) - groupSortValue(b)).forEach(sourceGroup => {
    const buckets = new Map();
    getRecords(sourceGroup).forEach(record => {
      const sectionName = hundredSection(record.room);
      if (!buckets.has(sectionName)) buckets.set(sectionName, []);
      buckets.get(sectionName).push(record);
    });

    [...buckets.entries()]
      .sort((a, b) => groupSortValue(a[0]) - groupSortValue(b[0]))
      .forEach(([sectionName, records]) => {
        sections.push({
          key: `${sourceGroup}::${sectionName}`,
          sourceGroup,
          sectionName,
          records: [...records].sort((a, b) => roomSortValue(a.room) - roomSortValue(b.room)),
        });
      });
  });

  return sections;
}

function cleanAssignments() {
  leaveGroups = new Set([...leaveGroups].filter(group => LEAVE_ELIGIBLE_GROUPS.includes(group)));
  const validSectionKeys = new Set(getLeaveSections().map(section => section.key));
  [...sectionAssignments.keys()].forEach(key => {
    const target = sectionAssignments.get(key);
    if (!validSectionKeys.has(key) || leaveGroups.has(target) || !CHIEF_GROUPS.includes(target)) {
      sectionAssignments.delete(key);
    }
  });
}

function buildPrintableGroups() {
  cleanAssignments();

  const result = new Map();

  originalGroups.forEach((records, groupName) => {
    if (CHIEF_GROUPS.includes(groupName) && leaveGroups.has(groupName)) return;
    if (!result.has(groupName)) result.set(groupName, []);
    result.get(groupName).push(...records);
  });

  const unassigned = [];
  getLeaveSections().forEach(section => {
    const target = sectionAssignments.get(section.key);
    if (target && !leaveGroups.has(target)) {
      if (!result.has(target)) result.set(target, []);
      result.get(target).push(...section.records);
    } else if (section.records.length) {
      unassigned.push(section);
    }
  });

  result.forEach((records, groupName) => {
    records.sort((a, b) => roomSortValue(a.room) - roomSortValue(b.room));
    if (!records.length) result.delete(groupName);
  });

  const sorted = new Map([...result.entries()].sort((a, b) => groupSortValue(a[0]) - groupSortValue(b[0])));
  return { groups: sorted, unassigned };
}

function renderPage(records, groupName) {
  const isArrivals = currentMode === MODE_ARRIVALS;
  const isVacant = currentMode === MODE_VACANT;
  const page = document.createElement('article');
  page.className = 'sheet-page';
  page.dataset.group = groupName;

  const rowsHtml = records.map(({ record, rowNumber }) => {
    const etdClass = currentMode === MODE_DEPARTURES && record.etd === ETD_HIGHLIGHT ? 'etd-highlight' : '';
    const roomClass = `room${isGreenRoom(record) ? ' room-green' : ''}`;

    if (isVacant) {
      return `<tr>
        <td class="idx">${rowNumber}</td>
        <td class="${roomClass}">${escapeHtml(record.room)}</td>
        <td>${escapeHtml(record.roomClass)}</td>
        <td>${escapeHtml(record.roomType)}</td>
        <td>${escapeHtml(record.foStatus)}</td>
        <td>${escapeHtml(record.nightsVacant)}</td>
        <td class="name">${escapeHtml(record.name)}</td>
        <td>${escapeHtml(record.arrival)}</td>
        <td>${escapeHtml(record.departure)}</td>
        <td class="status-cell">${escapeHtml(record.reservationStatus)}</td>
        <td>${escapeHtml(record.adults)}</td>
        <td>${escapeHtml(record.children)}</td>
        <td>${escapeHtml(record.discrepantStatus)}</td>
        <td>${escapeHtml(record.nextBlocked)}</td>
      </tr>`;
    }

    if (isArrivals) {
      return `<tr>
        <td class="idx">${rowNumber}</td>
        <td class="${roomClass}">${escapeHtml(record.room)}</td>
        <td class="${record.etaNextDay ? 'eta-next-day' : ''}">${escapeHtml(record.eta)}</td>
        <td>${escapeHtml(record.arrival)}</td>
        <td>${escapeHtml(record.adults)}</td>
        <td>${escapeHtml(record.children)}</td>
        <td>${escapeHtml(record.childAges)}</td>
        <td>${escapeHtml(record.departure)}</td>
        <td>${escapeHtml(record.etd)}</td>
        <td class="name">${escapeHtml(record.name)}</td>
        <td class="agent">${escapeHtml(record.travelAgent)}</td>
        <td class="notes">${escapeHtml(record.notes)}</td>
      </tr>`;
    }

    return `<tr>
      <td class="idx">${rowNumber}</td>
      <td class="${roomClass}">${escapeHtml(record.room)}</td>
      <td>${escapeHtml(record.eta)}</td>
      <td>${escapeHtml(record.arrival)}</td>
      <td>${escapeHtml(record.adults)}</td>
      <td>${escapeHtml(record.children)}</td>
      <td>${escapeHtml(record.childAges)}</td>
      <td>${escapeHtml(record.departure)}</td>
      <td class="${etdClass}">${escapeHtml(record.etd)}</td>
      <td class="agent">${escapeHtml(record.travelAgent)}</td>
      <td class="notes">${escapeHtml(record.notes)}</td>
    </tr>`;
  }).join('');

  const colgroup = isVacant
    ? `<colgroup>
        <col class="idx"><col class="room"><col class="vac-class"><col class="vac-type"><col class="vac-fo"><col class="vac-nights"><col class="name"><col class="date"><col class="date"><col class="vac-status"><col class="small"><col class="small"><col class="vac-disc"><col class="date">
      </colgroup>`
    : isArrivals
      ? `<colgroup>
        <col class="idx"><col class="room"><col class="time"><col class="date"><col class="small"><col class="small"><col class="age"><col class="date"><col class="time"><col class="name"><col class="agent"><col class="notes">
      </colgroup>`
      : `<colgroup>
        <col class="idx"><col class="room"><col class="time"><col class="date"><col class="small"><col class="small"><col class="age"><col class="date"><col class="time"><col class="agent"><col class="notes">
      </colgroup>`;

  const header = isVacant
    ? `<tr>
        <th></th>
        <th>Room</th>
        <th>Class</th>
        <th>Type</th>
        <th>FO</th>
        <th>Nights<br>Vac.</th>
        <th>Name</th>
        <th>Arr.</th>
        <th>Dep.</th>
        <th>Res.<br>Status</th>
        <th>Ad.</th>
        <th>Ch.</th>
        <th>Disc.</th>
        <th>Next<br>Blocked</th>
      </tr>`
    : isArrivals
      ? `<tr>
        <th></th>
        <th>Room</th>
        <th>ETA</th>
        <th>Arrival</th>
        <th>Adult<br>s</th>
        <th>Childr<br>en</th>
        <th>Child<br>Ages</th>
        <th>Departure</th>
        <th>ETD</th>
        <th>Name</th>
        <th>Travel Agent</th>
        <th>NOTLAR</th>
      </tr>`
      : `<tr>
        <th></th>
        <th>Room</th>
        <th>ETA</th>
        <th>Arrival</th>
        <th>Adul<br>ts</th>
        <th>Childr<br>en</th>
        <th>Child<br>Ages</th>
        <th>Departure</th>
        <th>ETD</th>
        <th>Travel Agent</th>
        <th>NOTLAR</th>
      </tr>`;

  const columnCount = isVacant ? 14 : (isArrivals ? 12 : 11);
  const officeHeader = groupName === 'Ofis'
    ? `<tr class="office-head-row"><th colspan="${columnCount}">OFİS</th></tr>`
    : '';

  page.innerHTML = `
    <div class="table-wrap">
      <table class="departure-table ${isVacant ? 'vacant-table' : (isArrivals ? 'arrival-table' : 'departure-mode-table')}">
        ${colgroup}
        <thead>${officeHeader}${header}</thead>
        <tbody>${rowsHtml}</tbody>
      </table>
    </div>`;

  preview.appendChild(page);
}

function renderPrintablePreview(groups) {
  printableGroups = groups;
  preview.classList.remove('empty');
  preview.classList.toggle('preview-arrivals', currentMode === MODE_ARRIVALS);
  preview.classList.toggle('preview-departures', currentMode === MODE_DEPARTURES);
  preview.classList.toggle('preview-vacant', currentMode === MODE_VACANT);
  preview.innerHTML = '';

  if (!groups.size) {
    preview.className = 'preview empty';
    preview.innerHTML = `<div class="empty-state no-print"><h2>Yazdırılacak grup yok</h2><p>En az bir kat şefi aktif olmalı veya izinli bölümler aktif bir kat şefine atanmalı.</p></div>`;
    return;
  }

  groups.forEach((records, groupName) => {
    const numbered = records.map((record, index) => ({ record, rowNumber: index + 1 }));

    // Ofis çıktısı tek tablo olarak bırakılır. Böylece uzun Arrivals listelerinde
    // tarayıcı sayfaları doğal böler; 28-33 gibi az satırlı ara sayfalar ve büyük
    // boşluklar oluşmaz. Normal kat şefi çıktıları ise ayrı sayfalara bölünmeye devam eder.
    if (groupName === 'Ofis') {
      renderPage(numbered, groupName);
      return;
    }

    // Uzun isim/acenteler satırı büyütebildiği için sabit 33 satır bazen
    // yazdırmada sayfa sınırında satırın kaybolmasına neden oluyordu.
    // Burada satır yüksekliği tahminiyle daha güvenli sayfalara bölüyoruz.
    const pages = splitNumberedRecordsForPrint(numbered);
    pages.forEach(pageRecords => renderPage(pageRecords, groupName));
  });
}

function renderSummary(groups) {
  summary.hidden = false;
  const groupEntries = [...groups.entries()];
  const totalRows = groupEntries.reduce((sum, [, rows]) => sum + rows.length, 0);
  const totalLate = groupEntries.reduce((sum, [, rows]) => sum + countLate(rows), 0);
  const lateTotalHtml = currentMode === MODE_DEPARTURES ? `<small>Total Late: ${totalLate}</small>` : '';
  const totalCard = `<div class="summary-card summary-total"><strong>${totalRows}</strong><span>Toplam</span>${lateTotalHtml}</div>`;
  const groupCards = groupEntries
    .map(([name, rows]) => {
      const lateHtml = currentMode === MODE_DEPARTURES ? `<small>Total Late: ${countLate(rows)}</small>` : '';
      return `<div class="summary-card"><strong>${rows.length}</strong><span>${escapeHtml(name)}</span>${lateHtml}</div>`;
    })
    .join('');
  summary.innerHTML = `${totalCard}${groupCards}`;
}

function setButtons({ printable = false, clearable = false } = {}) {
  printBtn.disabled = !printable;
  excelBtn.disabled = !printable;
  officeBtn.disabled = !clearable || currentMode === MODE_DND;
  officeExcelBtn.disabled = !clearable || currentMode === MODE_DND;
  clearBtn.disabled = !clearable;
  greenRoomsInput.disabled = !clearable || currentMode !== MODE_ARRIVALS;
  if (currentRoomsInput) currentRoomsInput.disabled = currentMode !== MODE_DND;
}

function activeChiefGroups() {
  return CHIEF_GROUPS.filter(group => !leaveGroups.has(group));
}

function formatLateInline(records) {
  const text = lateText(records);
  return text ? `<span class="late-count">${text}</span>` : '';
}

function formatLateOption(records) {
  const text = lateText(records);
  return text ? ` / ${text}` : '';
}

function renderAssignmentControls(currentGroups = new Map()) {
  if (!originalGroups.size) {
    assignmentPanel.hidden = true;
    greenPanel.hidden = true;
    updateCurrentRoomsPanel();
    chiefControls.innerHTML = '';
    assignmentStatus.innerHTML = '';
    return;
  }

  greenPanel.hidden = currentMode !== MODE_ARRIVALS;
  assignmentPanel.hidden = false;
  const sections = getLeaveSections();
  const activeGroups = activeChiefGroups();

  chiefControls.innerHTML = CHIEF_GROUPS.map(groupName => {
    const originalRecords = getRecords(groupName);
    const isLeave = leaveGroups.has(groupName);
    const canBeLeave = LEAVE_ELIGIBLE_GROUPS.includes(groupName);
    const displayRecords = isLeave ? originalRecords : (currentGroups.get(groupName) || []);
    const displayCount = displayRecords.length;
    const originalCount = originalRecords.length;
    const disabled = originalCount === 0;
    const selectableSections = sections.filter(section => section.sourceGroup !== groupName);

    const leaveToggleHtml = canBeLeave ? `
          <label class="leave-toggle">
            <input type="checkbox" data-leave-group="${groupName}" ${isLeave ? 'checked' : ''} ${disabled ? 'disabled' : ''}>
            <span>İzinli</span>
          </label>` : '';

    const selectHtml = !isLeave && selectableSections.length ? `
      <label class="assign-label" for="assign-${groupName}">Eklenecek bölümler</label>
      <select id="assign-${groupName}" class="assignment-select" data-target="${groupName}" multiple size="${Math.min(Math.max(selectableSections.length, 3), 7)}">
        ${selectableSections.map(section => {
          const assignedTarget = sectionAssignments.get(section.key);
          const selected = assignedTarget === groupName ? 'selected' : '';
          const optionDisabled = assignedTarget && assignedTarget !== groupName ? 'disabled' : '';
          const label = `${section.sourceGroup} → ${section.sectionName} (${section.records.length} oda${formatLateOption(section.records)})`;
          return `<option value="${escapeHtml(section.key)}" ${selected} ${optionDisabled}>${escapeHtml(label)}</option>`;
        }).join('')}
      </select>` : '';

    const leaveNoteLate = lateText(originalRecords);
    const leaveNote = leaveNoteLate ? `İzinli: ${originalCount} oda / ${leaveNoteLate}` : `İzinli: ${originalCount} oda`;

    return `
      <div class="chief-card ${isLeave ? 'is-leave' : ''} ${displayCount === 0 ? 'is-empty' : ''}">
        <div class="chief-top">
          <div>
            <strong>${escapeHtml(groupName)}</strong>
            <span>${displayCount} oda</span>
            ${formatLateInline(displayRecords)}
          </div>
          ${leaveToggleHtml}
        </div>
        ${isLeave ? `<p class="chief-note">${leaveNote}</p>` : ''}
        ${selectHtml}
      </div>`;
  }).join('');

  if (sections.length && activeGroups.length) {
    const assignedCount = sections.filter(section => sectionAssignments.has(section.key)).length;
    assignmentStatus.innerHTML = `<b>${sections.length}</b> izinli oda bölümü var, <b>${assignedCount}</b> tanesi eklendi.`;
  } else if (sections.length && !activeGroups.length) {
    assignmentStatus.innerHTML = 'Yazdırmak için en az bir kat şefi aktif olmalı.';
  } else {
    assignmentStatus.innerHTML = 'İzinli kat şefi seçilmedi.';
  }
}

function updateOutput(message) {
  updateGreenRooms();
  const { groups, unassigned } = buildPrintableGroups();
  renderAssignmentControls(groups);
  renderPrintablePreview(groups);
  renderSummary(groups);

  const hasAnyActiveChief = activeChiefGroups().some(group => getRecords(group).length || [...sectionAssignments.values()].includes(group));
  const canPrint = groups.size > 0 && unassigned.length === 0 && (hasAnyActiveChief || !leaveGroups.size);
  setButtons({ printable: canPrint, clearable: originalGroups.size > 0 });

  if (message) {
    setStatus(message, canPrint ? 'ok' : 'error');
    return;
  }

  if (unassigned.length) {
    const labels = unassigned.map(section => `${section.sourceGroup} → ${section.sectionName}`).join(', ');
    setStatus(`İzinli bölüm atanmamış: ${labels}. PDF almak için bunları aktif kat şeflerine ekle.`, 'error');
  } else if (!groups.size) {
    setStatus('Yazdırılacak grup kalmadı. En az bir kat şefi aktif olmalı.', 'error');
  } else if (lastFileName) {
    setStatus(`${lastFileName} ${modeLabel()} olarak işlendi. PDF / Excel çıktısı hazır: ${[...groups.keys()].join(', ')}.`, 'ok');
  }
}

function resetAssignmentsForNewData() {
  leaveGroups = new Set();
  sectionAssignments = new Map();
}

function processCurrentWorkbook(message = '') {
  if (!lastWorkbook && !lastWorkbooks.length) return;
  try {
    if (currentMode === MODE_VACANT) {
      dndResults = [];
      const itemsToProcess = lastWorkbooks.length ? lastWorkbooks : [];
      originalGroups = mergeVacantRecordItems(itemsToProcess);
      printableGroups = new Map(originalGroups);
      resetAssignmentsForNewData();
      updateOutput(message || `${lastFileName} ${modeLabel()} olarak işlendi.`);
      return;
    }

    if (currentMode === MODE_DND) {
      dndResults = processDndWorkbook(lastWorkbook);
      originalGroups = new Map([['DND / TİST', dndResults]]);
      printableGroups = new Map();
      resetAssignmentsForNewData();
      renderDndOutput(message || `${lastFileName} DND / TİST olarak işlendi.`);
      return;
    }

    dndResults = [];
    const workbooksToProcess = lastWorkbooks.length ? lastWorkbooks : [{ workbook: lastWorkbook, name: lastFileName }];
    originalGroups = mergeWorkbookGroups(workbooksToProcess, currentMode);
    printableGroups = new Map(originalGroups);
    resetAssignmentsForNewData();
    updateOutput(message || `${lastFileName} ${modeLabel()} olarak işlendi.`);
  } catch (error) {
    console.error(error);
    originalGroups = new Map();
    printableGroups = new Map();
    dndResults = [];
    resetAssignmentsForNewData();
    setStatus(error.message || 'Dosya işlenirken hata oluştu.', 'error');
    setButtons({ printable: false, clearable: Boolean(lastWorkbook || lastWorkbooks.length) });
    assignmentPanel.hidden = true;
    greenPanel.hidden = true;
    updateCurrentRoomsPanel();
    summary.hidden = true;
    preview.className = 'preview empty';
    preview.innerHTML = emptyPreviewHtml();
  }
}

async function readWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  // DND / TİST formlarında tarih başlıklarını Date objesine çevirmek bazı
  // tarayıcı/saat dilimlerinde 1 gün geri kaydırabiliyor. Bu yüzden DND'de
  // Excel seri numarasını ham bırakıp parseMatrixDate içinde güvenli çeviriyoruz.
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: currentMode !== MODE_DND });
  return { name: file.name, workbook };
}

async function handleFiles(files) {
  const fileList = [...files].filter(Boolean);
  if (!fileList.length) return;

  if (currentMode === MODE_DND && fileList.length > 1) {
    const warning = 'DND / TİST için tek Excel dosyası yükle. Departures, Arrivals ve Vacant için birden fazla dosya seçebilirsin.';
    fileInput.value = '';
    setStatus(warning, 'error');
    window.alert(warning);
    return;
  }

  try {
    const loadingLabel = currentMode === MODE_VACANT ? 'PDF' : 'Excel';
    setStatus(fileList.length === 1 ? `${fileList[0].name} okunuyor...` : `${fileList.length} ${loadingLabel} dosyası okunuyor...`);
    const loaded = [];

    for (const file of fileList) {
      const item = currentMode === MODE_VACANT ? await readVacantPdfFile(file) : await readWorkbookFile(file);
      if (currentMode !== MODE_VACANT) {
        const warning = wrongFileWarning(item.workbook, item.name, currentMode);
        if (warning) {
          fileInput.value = '';
          setStatus(warning, 'error');
          window.alert(warning);
          return;
        }
      }
      loaded.push(item);
    }

    lastWorkbooks = loaded;
    lastWorkbook = loaded[0]?.workbook || null;
    lastFileName = formatLoadedFileNames(loaded);

    const fileText = loaded.length === 1 ? loaded[0].name : `${loaded.length} ${currentMode === MODE_VACANT ? 'PDF' : 'Excel'} dosyası`;
    processCurrentWorkbook(`${fileText} ${modeLabel()} olarak işlendi. PDF veya Excel alabilirsin.`);
  } catch (error) {
    console.error(error);
    setStatus(error.message || 'Dosya işlenirken hata oluştu.', 'error');
    setButtons({ printable: Boolean(printableGroups.size || dndResults.length), clearable: Boolean(lastWorkbook || lastWorkbooks.length) });
    assignmentPanel.hidden = !printableGroups.size;
  }
}

async function handleFile(file) {
  return handleFiles([file]);
}


async function handleCurrentRoomFiles(files) {
  const fileList = [...files];
  if (!fileList.length) return;

  const merged = new Map();
  const names = [];

  try {
    setStatus('Güncel oda listeleri okunuyor...');

    for (const file of fileList) {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const fileMap = extractCurrentRoomMapFromWorkbook(workbook, file.name);
      fileMap.forEach(item => {
        mergeCurrentRoomRecord(merged, item.room, item.arrivalDate, file.name);
      });
      names.push(file.name);
    }

    currentRoomFilter = merged;
    currentRoomFileNames = names;
    updateCurrentRoomsStatus();

    if (!currentRoomFilter.size) {
      setStatus('Güncel oda listelerinde oda numarası bulunamadı. Dosyada 1001 gibi oda numaraları olmalı.', 'error');
      return;
    }

    if (currentMode === MODE_DND && lastWorkbook) {
      processCurrentWorkbook(`${currentRoomFilter.size} güncel oda ile DND / TİST yeniden hesaplandı.`);
    } else {
      setStatus(`${currentRoomFilter.size} güncel oda yüklendi. Şimdi DND / TİST Excel dosyasını yükle.`, 'ok');
    }
  } catch (error) {
    console.error(error);
    currentRoomFilter = new Map();
    currentRoomFileNames = [];
    updateCurrentRoomsStatus();
    setStatus(error.message || 'Güncel oda listesi okunurken hata oluştu.', 'error');
  }
}

function emptyPreviewHtml() {
  const helper = currentMode === MODE_DND
    ? 'İlk satırı tarih, ilk sütunu oda numarası olan DND / TİST Excel dosyasını yükle.'
    : currentMode === MODE_VACANT
      ? 'Vacant Rooms PDF dosyasını yükle. Oda listesi 1000ler / 2000ler gibi parçalanır.'
      : `${modeLabel()} Excel dosyasını yükle. Kolon isimleri örnekteki gibi olmalı.`;
  return `<div class="empty-state no-print"><h2>PDF önizlemesi burada görünecek</h2><p>${helper}</p></div>`;
}

function clearAll() {
  fileInput.value = '';
  originalGroups = new Map();
  printableGroups = new Map();
  leaveGroups = new Set();
  sectionAssignments = new Map();
  lastFileName = '';
  lastWorkbook = null;
  lastWorkbooks = [];
  greenRooms = new Set();
  currentRoomFilter = new Map();
  currentRoomFileNames = [];
  dndResults = [];
  dndDateWindowText = '';
  dndFilterStats = { active: false, currentRooms: 0, skippedOldRooms: 0, stoppedBeforeArrival: 0 };
  greenRoomsInput.value = '';
  if (currentRoomsInput) currentRoomsInput.value = '';
  preview.className = 'preview empty';
  preview.innerHTML = emptyPreviewHtml();
  summary.hidden = true;
  summary.innerHTML = '';
  assignmentPanel.hidden = true;
  greenPanel.hidden = true;
  updateCurrentRoomsPanel();
  chiefControls.innerHTML = '';
  assignmentStatus.innerHTML = '';
  setButtons({ printable: false, clearable: false });
  setStatus('Henüz dosya yüklenmedi.');
}

function updateModeUi() {
  appTitle.textContent = modeLabel();
  document.title = modeLabel();
  departuresModeBtn.classList.toggle('active', currentMode === MODE_DEPARTURES);
  arrivalsModeBtn.classList.toggle('active', currentMode === MODE_ARRIVALS);
  dndModeBtn.classList.toggle('active', currentMode === MODE_DND);
  vacantModeBtn?.classList.toggle('active', currentMode === MODE_VACANT);
  departuresModeBtn.setAttribute('aria-pressed', String(currentMode === MODE_DEPARTURES));
  arrivalsModeBtn.setAttribute('aria-pressed', String(currentMode === MODE_ARRIVALS));
  dndModeBtn.setAttribute('aria-pressed', String(currentMode === MODE_DND));
  vacantModeBtn?.setAttribute('aria-pressed', String(currentMode === MODE_VACANT));
  document.body.classList.toggle('mode-arrivals', currentMode === MODE_ARRIVALS);
  document.body.classList.toggle('mode-departures', currentMode === MODE_DEPARTURES);
  document.body.classList.toggle('mode-dnd', currentMode === MODE_DND);
  document.body.classList.toggle('mode-vacant', currentMode === MODE_VACANT);
  greenPanel.hidden = currentMode !== MODE_ARRIVALS || !originalGroups.size;
  greenRoomsInput.disabled = currentMode !== MODE_ARRIVALS || !originalGroups.size;
  if (fileInput) {
    fileInput.multiple = currentMode !== MODE_DND;
    fileInput.accept = currentMode === MODE_VACANT ? '.pdf' : '.xlsx,.xls,.csv';
  }
  if (uploadTitle) uploadTitle.textContent = currentMode === MODE_VACANT ? 'Vacant PDF Yükle' : (currentMode === MODE_DND ? 'Excel Yükle' : 'Excel Yükle / Birden Fazla Seç');
  updateCurrentRoomsPanel();
  if (currentMode === MODE_DND) {
    assignmentPanel.hidden = true;
    greenPanel.hidden = true;
  }
}


function showStartMenu() {
  if (startMenu) startMenu.hidden = true;
  appShell.hidden = false;
}

function showApp() {
  if (startMenu) startMenu.hidden = true;
  appShell.hidden = false;
}

function hasLoadedMainFile() {
  return Boolean(lastWorkbook || lastWorkbooks.length || originalGroups.size || printableGroups.size || dndResults.length);
}

function switchModeAndMaybeClear(mode) {
  if (![MODE_DEPARTURES, MODE_ARRIVALS, MODE_DND, MODE_VACANT].includes(mode)) return;
  if (mode === currentMode) return;

  const previousMode = currentMode;
  const hadData = hasLoadedMainFile();
  currentMode = mode;
  updateModeUi();

  if (hadData) {
    clearAll();
    updateModeUi();
    setStatus(`${modeLabel()} seçildi. Önceki ${modeLabel(previousMode)} dosyası otomatik temizlendi. Yeni ${mode === MODE_VACANT ? 'PDF' : 'Excel'} yükle.`, 'ok');
  } else {
    preview.innerHTML = emptyPreviewHtml();
    setStatus(`${modeLabel()} seçildi. ${mode === MODE_VACANT ? 'PDF' : 'Excel'} dosyası yükle.`);
  }
}

function selectModeFromMenu(mode) {
  showApp();
  switchModeAndMaybeClear(mode);
}

function setMode(mode) {
  if (appShell.hidden) {
    selectModeFromMenu(mode);
    return;
  }
  switchModeAndMaybeClear(mode);
}

fileInput.addEventListener('change', event => {
  const files = event.target.files;
  if (files?.length) handleFiles(files);
});

currentRoomsInput?.addEventListener('change', event => {
  if (event.target.files?.length) handleCurrentRoomFiles(event.target.files);
});

['dragenter', 'dragover'].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach(eventName => {
  dropZone.addEventListener(eventName, event => {
    event.preventDefault();
    dropZone.classList.remove('dragover');
  });
});

dropZone.addEventListener('drop', event => {
  const files = event.dataTransfer.files;
  if (files?.length) handleFiles(files);
});

chiefControls.addEventListener('change', event => {
  const leaveInput = event.target.closest('input[data-leave-group]');
  if (leaveInput) {
    const groupName = leaveInput.dataset.leaveGroup;
    if (!LEAVE_ELIGIBLE_GROUPS.includes(groupName)) return;
    if (leaveInput.checked) {
      leaveGroups.add(groupName);
    } else {
      leaveGroups.delete(groupName);
    }
    cleanAssignments();
    updateOutput();
    return;
  }

  const select = event.target.closest('select.assignment-select');
  if (select) {
    const targetGroup = select.dataset.target;
    const currentTargetSections = new Set([...sectionAssignments.entries()]
      .filter(([, target]) => target === targetGroup)
      .map(([key]) => key));
    const selectedKeys = new Set([...select.selectedOptions].map(option => option.value));

    currentTargetSections.forEach(key => {
      if (!selectedKeys.has(key)) sectionAssignments.delete(key);
    });

    selectedKeys.forEach(key => {
      sectionAssignments.set(key, targetGroup);
    });

    cleanAssignments();
    updateOutput();
  }
});


function safeSheetName(name, usedNames = new Set()) {
  const base = clean(name || 'Sayfa')
    .replace(/[\\/?*\[\]:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 28) || 'Sayfa';
  let candidate = base;
  let index = 2;
  while (usedNames.has(candidate)) {
    const suffix = ` ${index}`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }
  usedNames.add(candidate);
  return candidate;
}

function cellRef(rowIndex, colIndex) {
  return XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
}

function excelBorder(color = '555555') {
  return {
    top: { style: 'thin', color: { rgb: color } },
    right: { style: 'thin', color: { rgb: color } },
    bottom: { style: 'thin', color: { rgb: color } },
    left: { style: 'thin', color: { rgb: color } },
  };
}

function excelCellStyle({ fill = 'FFFFFF', bold = true, size = 10, align = 'center', valign = 'center', border = true, fontColor = '111111' } = {}) {
  return {
    fill: { patternType: 'solid', fgColor: { rgb: fill } },
    font: { name: 'Arial', bold, sz: size, color: { rgb: fontColor } },
    alignment: { horizontal: align, vertical: valign, wrapText: true },
    border: border ? excelBorder() : undefined,
  };
}

function excelHeaders(isArrivals, isVacant = false) {
  if (isVacant) {
    return ['', 'Room', 'Class', 'Type', 'FO', 'Nights\nVac.', 'Name', 'Arr.', 'Dep.', 'Res.\nStatus', 'Ad.', 'Ch.', 'Disc.', 'Next\nBlocked'];
  }
  if (isArrivals) {
    return ['', 'Room', 'ETA', 'Arrival', 'Adult\ns', 'Childr\nen', 'Child\nAges', 'Departure', 'ETD', 'Name', 'Travel\nAgent', 'NOTLAR'];
  }
  return ['', 'Room', 'ETA', 'Arrival', 'Adul\nts', 'Childr\nen', 'Child\nAges', 'Departure', 'ETD', 'Travel Agent', 'NOTLAR'];
}

function recordToExcelRow(record, rowNumber, isArrivals, isVacant = false) {
  if (isVacant) {
    return [
      rowNumber,
      record.room,
      record.roomClass,
      record.roomType,
      record.foStatus,
      record.nightsVacant,
      record.name,
      record.arrival,
      record.departure,
      record.reservationStatus,
      record.adults,
      record.children,
      record.discrepantStatus,
      record.nextBlocked,
    ];
  }

  if (isArrivals) {
    return [
      rowNumber,
      record.room,
      record.eta,
      record.arrival,
      record.adults,
      record.children,
      record.childAges,
      record.departure,
      record.etd,
      record.name,
      record.travelAgent,
      record.notes,
    ];
  }

  return [
    rowNumber,
    record.room,
    record.eta,
    record.arrival,
    record.adults,
    record.children,
    record.childAges,
    record.departure,
    record.etd,
    record.travelAgent,
    record.notes,
  ];
}

function excelColumnWidths(isArrivals, isVacant = false) {
  if (isVacant) {
    return [4.5, 7, 8, 7, 6, 7.5, 22, 9, 9, 14, 5, 5, 7, 13].map(wch => ({ wch }));
  }
  if (isArrivals) {
    return [4.5, 8, 9, 11, 6, 6.5, 8.5, 11, 8.5, 13, 12, 34].map(wch => ({ wch }));
  }
  return [4.5, 8, 8.5, 11, 6, 6.5, 8.5, 11, 8.5, 16, 34].map(wch => ({ wch }));
}

function styleExcelSheet(ws, records, groupName, isArrivals, startRow, isVacant = false) {
  const headers = excelHeaders(isArrivals, isVacant);
  const columnCount = headers.length;
  const headerFill = isVacant ? '4B5563' : (isArrivals ? '79A9D4' : 'C8755C');
  const headerStyle = excelCellStyle({ fill: headerFill, bold: true, size: 10 });
  const bodyStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 10 });
  const notesStyle = excelCellStyle({ fill: 'FFFFFF', bold: false, size: 10, align: 'left' });
  const smallTextStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 9 });
  const compactArrivalTextStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 8 });
  const largerArrivalStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 12 });
  const etaNextDayStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 12, fontColor: '5B5FC7' });
  const etdHighlightStyle = excelCellStyle({ fill: 'FFF176', bold: true, size: 10 });
  const roomGreenStyle = excelCellStyle({ fill: 'B8D8BD', bold: true, size: isArrivals ? 12 : 10 });
  const vacantSmallStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 8 });

  ws['!cols'] = excelColumnWidths(isArrivals, isVacant);
  ws['!rows'] = [];

  if (groupName === 'Ofis') {
    ws['!rows'][0] = { hpt: 24 };
    ws['!merges'] = ws['!merges'] || [];
    ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: columnCount - 1 } });
    const titleCell = ws[cellRef(0, 0)] || { t: 's', v: 'OFİS' };
    titleCell.v = 'OFİS';
    titleCell.t = 's';
    titleCell.s = {
      fill: { patternType: 'solid', fgColor: { rgb: 'FFFFFF' } },
      font: { name: 'Arial', bold: true, sz: 16, color: { rgb: '111111' } },
      alignment: { horizontal: 'right', vertical: 'center' },
    };
    ws[cellRef(0, 0)] = titleCell;
  }

  ws['!rows'][startRow] = { hpt: 24 };
  for (let c = 0; c < columnCount; c += 1) {
    const cell = ws[cellRef(startRow, c)] || { t: 's', v: headers[c] };
    cell.s = headerStyle;
    ws[cellRef(startRow, c)] = cell;
  }

  records.forEach((record, index) => {
    const rowIndex = startRow + 1 + index;
    ws['!rows'][rowIndex] = { hpt: isArrivals ? 32 : (isVacant ? 30 : 26) };

    for (let c = 0; c < columnCount; c += 1) {
      const ref = cellRef(rowIndex, c);
      const cell = ws[ref] || { t: 's', v: '' };
      cell.t = typeof cell.v === 'number' ? 'n' : 's';
      cell.s = bodyStyle;

      // Child Ages kolonu Excel'de tarih formatına dönmesin; her zaman metin kalsın.
      if (!isVacant && c === 6) {
        cell.t = 's';
        cell.v = String(cell.v ?? '');
        cell.z = '@';
      }

      if (isVacant && [2, 3, 4, 5, 9, 12, 13].includes(c)) {
        cell.s = vacantSmallStyle;
      }
      if (isArrivals && [1, 2, 4, 5, 6].includes(c)) {
        cell.s = largerArrivalStyle;
      }
      if (isArrivals && c === 2 && record.etaNextDay) {
        cell.s = etaNextDayStyle;
      }
      if (isArrivals && [3, 7, 8, 9, 10].includes(c)) {
        cell.s = compactArrivalTextStyle;
      }
      if (!isVacant && c === columnCount - 1) {
        cell.s = notesStyle;
      }
      if (c === 1 && isGreenRoom(record)) {
        cell.s = roomGreenStyle;
      }
      if (!isArrivals && !isVacant && c === 8 && record.etd === ETD_HIGHLIGHT) {
        cell.s = etdHighlightStyle;
      }

      ws[ref] = cell;
    }
  });

  ws['!margins'] = { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25, header: 0, footer: 0 };
  ws['!pageSetup'] = { paperSize: 9, orientation: 'portrait', fitToWidth: 1, fitToHeight: 0 };
}

function buildExcelSheet(records, groupName, isArrivals, isVacant = false) {
  const headers = excelHeaders(isArrivals, isVacant);
  const rows = [];
  let startRow = 0;

  if (groupName === 'Ofis') {
    rows.push(['OFİS', ...Array(headers.length - 1).fill('')]);
    startRow = 1;
  }

  rows.push(headers);
  records.forEach((record, index) => rows.push(recordToExcelRow(record, index + 1, isArrivals, isVacant)));

  const ws = XLSX.utils.aoa_to_sheet(rows);
  styleExcelSheet(ws, records, groupName, isArrivals, startRow, isVacant);
  return ws;
}

function downloadExcelForGroups(groups, { filePrefix = modeLabel(), office = false } = {}) {
  if (!groups || !groups.size) {
    setStatus('Excel için satır bulunamadı.', 'error');
    return;
  }

  updateGreenRooms();
  const isArrivals = currentMode === MODE_ARRIVALS;
  const isVacant = currentMode === MODE_VACANT;
  const wb = XLSX.utils.book_new();
  const usedNames = new Set();

  groups.forEach((records, groupName) => {
    const sheetName = safeSheetName(groupName === 'Ofis' ? 'OFİS' : groupName, usedNames);
    const ws = buildExcelSheet(records, groupName, isArrivals, isVacant);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  const now = new Date();
  const datePart = now.toLocaleDateString('tr-TR').replace(/\./g, '-');
  const cleanPrefix = clean(filePrefix).replace(/[^a-z0-9ığüşöçİĞÜŞÖÇ_-]+/gi, '_') || modeLabel();
  const fileName = `${cleanPrefix}${office ? '_Ofis' : ''}_${datePart}.xlsx`;
  XLSX.writeFile(wb, fileName, { bookType: 'xlsx', cellStyles: true });
  setStatus(`${fileName} indirildi.`, 'ok');
}

function exportCurrentExcel() {
  if (!originalGroups.size) {
    setStatus(currentMode === MODE_VACANT ? 'Önce Vacant PDF dosyası yükle.' : 'Önce Excel dosyası yükle.', 'error');
    return;
  }

  if (currentMode === MODE_DND) {
    downloadDndExcel();
    return;
  }

  const { groups, unassigned } = buildPrintableGroups();
  if (unassigned.length) {
    const labels = unassigned.map(section => `${section.sourceGroup} → ${section.sectionName}`).join(', ');
    setStatus(`Excel alınamadı. Önce atanmayan izinli bölümleri seç: ${labels}.`, 'error');
    return;
  }

  downloadExcelForGroups(groups, { filePrefix: modeLabel() });
}

function exportOfficeExcelDirect() {
  if (!originalGroups.size) {
    setStatus(currentMode === MODE_VACANT ? 'Önce Vacant PDF dosyası yükle.' : 'Önce Excel dosyası yükle.', 'error');
    return;
  }

  const officeGroups = buildOfficeGroups();
  downloadExcelForGroups(officeGroups, { filePrefix: modeLabel(), office: true });
}


function downloadDndExcel() {
  if (!dndResults.length) {
    setStatus('DND / TİST Excel için arka arkaya kayıt bulunamadı.', 'error');
    return;
  }

  const headers = ['', 'Room', 'Çarşamba Hariç Gün', 'Gün', 'Detay'];
  const rows = [
    ['DND / TİST ARKA ARKAYA ODALAR', '', '', '', ''],
    [`Kontrol: ${dndDateWindowText || 'Tüm tarihler'}`, '', '', '', ''],
    headers,
    ...dndResults.map((item, index) => [
      index + 1,
      item.room,
      item.daysWithoutWednesday ?? item.days,
      item.days,
      item.details,
    ]),
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [4.5, 9, 16, 7, 60].map(wch => ({ wch }));
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 4 } },
  ];
  ws['!rows'] = [{ hpt: 24 }, { hpt: 18 }, { hpt: 24 }];

  const titleStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 16, border: false });
  titleStyle.alignment.horizontal = 'left';
  const subtitleStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 10, border: false });
  subtitleStyle.alignment.horizontal = 'left';
  const headerStyle = excelCellStyle({ fill: 'B71C1C', bold: true, size: 10 });
  headerStyle.font.color = { rgb: 'FFFFFF' };
  const bodyStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 10 });
  const detailStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 10, align: 'left' });

  ws[cellRef(0, 0)].s = titleStyle;
  ws[cellRef(1, 0)].s = subtitleStyle;
  for (let c = 0; c < headers.length; c += 1) {
    ws[cellRef(2, c)].s = headerStyle;
  }
  dndResults.forEach((_, index) => {
    const rowIndex = 3 + index;
    ws['!rows'][rowIndex] = { hpt: 24 };
    for (let c = 0; c < headers.length; c += 1) {
      ws[cellRef(rowIndex, c)].s = c === headers.length - 1 ? detailStyle : bodyStyle;
    }
  });
  ws['!margins'] = { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25, header: 0, footer: 0 };
  ws['!pageSetup'] = { paperSize: 9, orientation: 'portrait', fitToWidth: 1, fitToHeight: 0 };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'DND TIST');
  const now = new Date();
  const datePart = now.toLocaleDateString('tr-TR').replace(/\./g, '-');
  const fileName = `DND_TIST_${datePart}.xlsx`;
  XLSX.writeFile(wb, fileName, { bookType: 'xlsx', cellStyles: true });
  setStatus(`${fileName} indirildi.`, 'ok');
}

function printCleanPdf(options = {}) {
  const { skipAssignmentCheck = false, printTitle = modeLabel(), afterRestore = null } = options;

  if (!skipAssignmentCheck) {
    const { unassigned } = buildPrintableGroups();
    if (unassigned.length) {
      const labels = unassigned.map(section => `${section.sourceGroup} → ${section.sectionName}`).join(', ');
      setStatus(`PDF alınamadı. Önce atanmayan izinli bölümleri seç: ${labels}.`, 'error');
      return;
    }
  }

  const oldTitle = document.title;
  document.title = printTitle;
  document.body.classList.add('printing-clean');

  let restored = false;
  const restore = () => {
    if (restored) return;
    restored = true;
    document.title = oldTitle;
    document.body.classList.remove('printing-clean');
    window.removeEventListener('afterprint', restore);
    if (typeof afterRestore === 'function') setTimeout(afterRestore, 0);
  };

  window.addEventListener('afterprint', restore);
  setTimeout(() => window.print(), 50);
  setTimeout(restore, 3000);
}

function printOfficeDirect() {
  if (!originalGroups.size) {
    setStatus(currentMode === MODE_VACANT ? 'Önce Vacant PDF dosyası yükle.' : 'Önce Excel dosyası yükle.', 'error');
    return;
  }

  const officeGroups = buildOfficeGroups();
  if (!officeGroups.size) {
    setStatus('Ofis çıktısı için satır bulunamadı.', 'error');
    return;
  }

  updateGreenRooms();
  renderPrintablePreview(officeGroups);
  renderSummary(officeGroups);
  setStatus(`Ofis çıktısı ${modeLabel()} için ayrıştırmadan hazırlandı. Yazdırma ekranı açılıyor...`, 'ok');
  printCleanPdf({
    skipAssignmentCheck: true,
    printTitle: modeLabel(),
    afterRestore: () => updateOutput(),
  });
}

greenRoomsInput.addEventListener('input', () => {
  if (originalGroups.size) updateOutput();
});

startDeparturesBtn?.addEventListener('click', () => selectModeFromMenu(MODE_DEPARTURES));
startArrivalsBtn?.addEventListener('click', () => selectModeFromMenu(MODE_ARRIVALS));
startDndBtn?.addEventListener('click', () => selectModeFromMenu(MODE_DND));
backMenuBtn?.addEventListener('click', showStartMenu);
departuresModeBtn.addEventListener('click', () => setMode(MODE_DEPARTURES));
arrivalsModeBtn.addEventListener('click', () => setMode(MODE_ARRIVALS));
dndModeBtn.addEventListener('click', () => setMode(MODE_DND));
vacantModeBtn?.addEventListener('click', () => setMode(MODE_VACANT));
officeBtn.addEventListener('click', printOfficeDirect);
officeExcelBtn.addEventListener('click', exportOfficeExcelDirect);
excelBtn.addEventListener('click', exportCurrentExcel);
printBtn.addEventListener('click', () => printCleanPdf({ printTitle: modeLabel(), skipAssignmentCheck: currentMode === MODE_DND }));
clearBtn.addEventListener('click', clearAll);

updateModeUi();
setButtons({ printable: false, clearable: false });
showApp();
