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
const lateCoutModeBtn = document.getElementById('lateCoutModeBtn');
const greenPanel = document.getElementById('greenPanel');
const greenRoomsInput = document.getElementById('greenRoomsInput');
const vacantRoomsPdfInput = document.getElementById('vacantRoomsPdfInput');
const vacantRoomsStatus = document.getElementById('vacantRoomsStatus');
const currentRoomsPanel = document.getElementById('currentRoomsPanel');
const currentRoomsInput = document.getElementById('currentRoomsInput');
const currentRoomsStatus = document.getElementById('currentRoomsStatus');
const uploadTitle = document.getElementById('uploadTitle');
const settingsBtn = document.getElementById('settingsBtn');
const settingsOverlay = document.getElementById('settingsOverlay');
const settingsSaveBtn = document.getElementById('settingsSaveBtn');
const settingsExportBtn = document.getElementById('settingsExportBtn');
const settingsImportBtn = document.getElementById('settingsImportBtn');
const settingsImportInput = document.getElementById('settingsImportInput');
const settingsResetBtn = document.getElementById('settingsResetBtn');
const settingsCloseBtn = document.getElementById('settingsCloseBtn');
const downloadsAutoPanel = document.getElementById('downloadsAutoPanel');
const downloadsAutoStatus = document.getElementById('downloadsAutoStatus');
const downloadsConnectBtn = document.getElementById('downloadsConnectBtn');
const downloadsRefreshBtn = document.getElementById('downloadsRefreshBtn');

const MODE_DEPARTURES = 'departures';
const MODE_ARRIVALS = 'arrivals';
const MODE_DND = 'dnd';
const MODE_VACANT = 'vacant';
const MODE_LATECOUT = 'latecout';
let ETD_HIGHLIGHT = '17:00';
const PRINT_ROWS_PER_PAGE = {
  departures: 38,
  arrivals: 38,
  vacant: 38,
  latecout: 41,
};
const CHIEF_GROUPS = ['1000ler', '2000ler', '3000ler', '4000ler', '5000ler'];
const LEAVE_ELIGIBLE_GROUPS = CHIEF_GROUPS.filter(group => group !== '5000ler');

let currentMode = MODE_ARRIVALS;
let originalGroups = new Map();
let printableGroups = new Map();
let leaveGroups = new Set();
let sectionAssignments = new Map(); // key: "1000ler::1100ler", value: target chief group
let lastFileName = '';
let lastWorkbook = null;
let lastWorkbooks = [];
let greenRooms = new Set();
let vacantRoomsFileName = '';
let currentRoomFilter = new Map(); // room -> { room, arrivalDate, source }
let currentRoomFileNames = [];
let dndResults = [];
let lateCoutResults = [];
let lateCoutDateText = '';
let dndDateWindowText = '';
let dndFilterStats = { active: false, currentRooms: 0, skippedOldRooms: 0, stoppedBeforeArrival: 0 };

/* ---------- İndirilenler klasörü otomatik yükleme ---------- */
const DOWNLOADS_DB_NAME = 'operaDownloadsFolderV1';
const DOWNLOADS_DB_STORE = 'handles';
const DOWNLOADS_HANDLE_KEY = 'downloadsDirectory';
const DOWNLOADS_SCAN_INTERVAL_MS = 15000;
const AUTO_FILES_PER_TYPE = 3;
let downloadsDirectoryHandle = null;
let downloadsLastSignature = '';
let downloadsScanBusy = false;
let downloadsAutoLoadBusy = false;
let downloadsScanTimer = null;
let autoDownloadFiles = { arrivals: [], departures: [], vacant: [] };


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

/* ---------- Ayarlar ---------- */
const SETTINGS_KEY = 'listeAyarlarV1';
const DEFAULT_SETTINGS = {
  headerColor: '#c8755c',
  arrivalsHeaderColor: '#79a9d4',
  accentColor: '#8e4d3b',
  etdColor: '#fff176',
  greenColor: '#b8d8bd',
  screenFontSize: 13,
  printFontSize: 11,
  screenRowHeight: 31,
  printRowHeight: 24,
  rowsPerPage: 38,
  etdLateTime: '17:00',
  titleDepartures: 'Check Out List',
  titleArrivals: 'Check In List',
  titleVacant: 'Vacant List',
  fontFamily: 'Arial, Helvetica, sans-serif',
  headerTextColor: '#2a120c',
  bodyTextColor: '#171717',
  borderColor: '#635d56',
  borderWidth: 1.5,
  boldBody: true,
  etdHighlightEnabled: true,
  showReportTitle: true,
  fillBlankRows: true,
  showSummary: true,
  printPagePadding: 8,
  dndMinDays: 2,
  excelRowHeight: 0,
  lateEtdList: '17:00, 18:00',
  titleLateCout: 'PALACE  17:00  LATE CHECK OUT GÖREV DAĞILIMI',
  lateCoutSupervisor: 'RECEP KESKİN',
  lateCoutServiceTime: '18:30',
  wIdx: 0, wRoom: 0, wTime: 0, wDate: 0, wSmall: 0,
  wAge: 0, wName: 0, wAgent: 0, wNotes: 0,
};
let appSettings = { ...DEFAULT_SETTINGS };

const SETTING_COLOR_KEYS = ['headerColor', 'arrivalsHeaderColor', 'accentColor', 'etdColor', 'greenColor', 'headerTextColor', 'bodyTextColor', 'borderColor'];
const SETTING_BOOL_KEYS = ['boldBody', 'etdHighlightEnabled', 'showReportTitle', 'fillBlankRows', 'showSummary'];
const SETTING_WIDTH_KEYS = ['wIdx', 'wRoom', 'wTime', 'wDate', 'wSmall', 'wAge', 'wName', 'wAgent', 'wNotes'];
const SETTING_ZERO_OK_KEYS = [...SETTING_WIDTH_KEYS, 'printPagePadding', 'excelRowHeight'];
const SETTING_COL_CLASS = {
  wIdx: 'idx', wRoom: 'room', wTime: 'time', wDate: 'date', wSmall: 'small',
  wAge: 'age', wName: 'name', wAgent: 'agent', wNotes: 'notes',
};

function clampNumber(value, min, max, fallback) {
  const num = Number(value);
  if (!Number.isFinite(num)) return fallback;
  return Math.min(max, Math.max(min, num));
}

function sanitizeSettings(input = {}) {
  const merged = { ...DEFAULT_SETTINGS, ...input };
  merged.screenFontSize = clampNumber(merged.screenFontSize, 7, 24, DEFAULT_SETTINGS.screenFontSize);
  merged.printFontSize = clampNumber(merged.printFontSize, 6, 20, DEFAULT_SETTINGS.printFontSize);
  merged.screenRowHeight = clampNumber(merged.screenRowHeight, 16, 80, DEFAULT_SETTINGS.screenRowHeight);
  merged.printRowHeight = clampNumber(merged.printRowHeight, 14, 60, DEFAULT_SETTINGS.printRowHeight);
  merged.rowsPerPage = Math.round(clampNumber(merged.rowsPerPage, 5, 80, DEFAULT_SETTINGS.rowsPerPage));
  merged.borderWidth = clampNumber(merged.borderWidth, 0.5, 5, DEFAULT_SETTINGS.borderWidth);
  merged.printPagePadding = clampNumber(merged.printPagePadding, 0, 25, DEFAULT_SETTINGS.printPagePadding);
  merged.dndMinDays = Math.round(clampNumber(merged.dndMinDays, 2, 10, DEFAULT_SETTINGS.dndMinDays));
  merged.excelRowHeight = clampNumber(merged.excelRowHeight, 0, 60, DEFAULT_SETTINGS.excelRowHeight);
  SETTING_WIDTH_KEYS.forEach(key => {
    merged[key] = clampNumber(merged[key], 0, 400, 0);
  });
  SETTING_BOOL_KEYS.forEach(key => {
    merged[key] = Boolean(merged[key]);
  });
  if (!/^\d{2}:\d{2}$/.test(String(merged.etdLateTime || ''))) merged.etdLateTime = DEFAULT_SETTINGS.etdLateTime;
  SETTING_COLOR_KEYS.forEach(key => {
    if (!/^#[0-9a-fA-F]{6}$/.test(String(merged[key] || ''))) merged[key] = DEFAULT_SETTINGS[key];
  });
  merged.fontFamily = String(merged.fontFamily || '').replace(/[;}<>]/g, '').trim() || DEFAULT_SETTINGS.fontFamily;
  merged.lateEtdList = String(merged.lateEtdList || '').trim() || DEFAULT_SETTINGS.lateEtdList;
  merged.lateCoutSupervisor = String(merged.lateCoutSupervisor ?? '').trim();
  merged.lateCoutServiceTime = String(merged.lateCoutServiceTime ?? '').trim();
  return merged;
}

function lateEtdList() {
  const list = String(appSettings.lateEtdList || '')
    .split(/[,;\s]+/)
    .map(token => token.trim())
    .filter(token => /^\d{1,2}:\d{2}$/.test(token))
    .map(token => token.padStart(5, '0'));
  return list.length ? [...new Set(list)] : ['17:00', '18:00'];
}

function loadSettings() {
  try {
    const rawText = localStorage.getItem(SETTINGS_KEY);
    if (rawText) appSettings = sanitizeSettings(JSON.parse(rawText));
  } catch (error) {
    console.warn('Ayarlar okunamadı, varsayılanlar kullanılıyor.', error);
    appSettings = { ...DEFAULT_SETTINGS };
  }
}

function saveSettingsToStorage() {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(appSettings));
  } catch (error) {
    console.warn('Ayarlar kaydedilemedi.', error);
  }
}

function buildSettingsCss(s) {
  let widthCss = '';
  Object.entries(SETTING_COL_CLASS).forEach(([key, cls]) => {
    if (s[key] > 0) widthCss += `.departure-table col.${cls} { width: ${s[key]}px !important; }\n`;
  });

  return `
:root { --header: ${s.headerColor}; --accent: ${s.accentColor}; --etd: ${s.etdColor}; }
.departure-table.arrival-table thead th { background: ${s.arrivalsHeaderColor}; }
.departure-table thead th { color: ${s.headerTextColor}; }
.departure-table .room-green { background: ${s.greenColor} !important; }
table.departure-table { font-size: ${s.screenFontSize}px; font-family: ${s.fontFamily}; color: ${s.bodyTextColor}; }
.departure-table th, .departure-table td {
  height: ${s.screenRowHeight}px;
  border: ${s.borderWidth}px solid ${s.borderColor} !important;
}
.departure-table tbody td { font-weight: ${s.boldBody ? 900 : 400}; }
.departure-table .blank-fill-row td { border-color: transparent !important; }
${s.etdHighlightEnabled ? '' : '.departure-table .etd-highlight { background: transparent !important; }'}
${widthCss}
@media print {
  table.departure-table, table.departure-table.arrival-table { font-size: ${s.printFontSize}px !important; }
  .departure-table th, .departure-table td { height: ${s.printRowHeight}px !important; }
  .sheet-page { padding: ${s.printPagePadding}mm ${s.printPagePadding}mm !important; }
}`;
}

function applySettings() {
  ETD_HIGHLIGHT = appSettings.etdLateTime;
  PRINT_ROWS_PER_PAGE.departures = appSettings.rowsPerPage;
  PRINT_ROWS_PER_PAGE.arrivals = appSettings.rowsPerPage;
  PRINT_ROWS_PER_PAGE.vacant = appSettings.rowsPerPage;

  let styleTag = document.getElementById('settingsStyle');
  if (!styleTag) {
    styleTag = document.createElement('style');
    styleTag.id = 'settingsStyle';
    document.head.appendChild(styleTag);
  }
  styleTag.textContent = buildSettingsCss(appSettings);
}

function fillSettingsForm() {
  document.querySelectorAll('#settingsOverlay [data-setting]').forEach(input => {
    const key = input.dataset.setting;
    if (!(key in appSettings)) return;
    if (input.type === 'checkbox') {
      input.checked = Boolean(appSettings[key]);
    } else if (SETTING_ZERO_OK_KEYS.includes(key)) {
      input.value = Number(appSettings[key]) > 0 ? appSettings[key] : '';
    } else {
      input.value = appSettings[key];
    }
  });
}

function collectSettingsForm() {
  const next = { ...appSettings };
  document.querySelectorAll('#settingsOverlay [data-setting]').forEach(input => {
    const key = input.dataset.setting;
    if (!(key in DEFAULT_SETTINGS)) return;
    if (input.type === 'checkbox') {
      next[key] = input.checked;
    } else if (input.type === 'number') {
      const num = Number(input.value);
      if (SETTING_ZERO_OK_KEYS.includes(key)) {
        next[key] = Number.isFinite(num) && num >= 0 ? num : 0;
      } else if (Number.isFinite(num) && num > 0) {
        next[key] = num;
      }
    } else if (clean(input.value)) {
      next[key] = input.value.trim();
    }
  });
  return sanitizeSettings(next);
}

function rerenderAfterSettings() {
  if (currentMode === MODE_DND) {
    if (dndResults.length) renderDndOutput();
    return;
  }
  if (currentMode === MODE_LATECOUT) {
    if (lastWorkbooks.length || lastWorkbook) {
      try {
        lateCoutResults = processLateCoutFiles(lastWorkbooks.length ? lastWorkbooks : [{ workbook: lastWorkbook, name: lastFileName }]);
      } catch (error) {
        console.error(error);
      }
    }
    renderLateCoutOutput();
    return;
  }
  if (originalGroups.size) updateOutput();
}

function openSettings() {
  fillSettingsForm();
  settingsOverlay.hidden = false;
}

function closeSettings() {
  settingsOverlay.hidden = true;
}

function exportSettingsFile() {
  const blob = new Blob([JSON.stringify(appSettings, null, 2)], { type: 'application/json' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = 'liste-ayarlar.json';
  link.click();
  URL.revokeObjectURL(link.href);
}

function importSettingsFile(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const parsed = JSON.parse(String(reader.result || '{}'));
      appSettings = sanitizeSettings(parsed);
      saveSettingsToStorage();
      applySettings();
      fillSettingsForm();
      rerenderAfterSettings();
      setStatus('Ayarlar dosyadan yüklendi.', 'ok');
    } catch (error) {
      setStatus('Ayar dosyası okunamadı. Geçerli bir JSON dosyası seç.', 'error');
    }
  };
  reader.readAsText(file);
}

function hexToRgb(value, fallback = 'FFFFFF') {
  const text = String(value || '').replace('#', '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(text) ? text : fallback;
}
/* ---------- Ayarlar sonu ---------- */

function modeLabel(mode = currentMode) {
  if (mode === MODE_ARRIVALS) return 'Arrivals';
  if (mode === MODE_DND) return 'DND / TİST';
  if (mode === MODE_VACANT) return 'Vacant Rooms';
  if (mode === MODE_LATECOUT) return 'Late Check Out';
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

// Arrivals için Child Ages: Excel hücresinde görünen metni hiçbir dönüşüm yapmadan aynen göster.
// displayValue, SheetJS'in hücredeki biçimlendirilmiş görünümüdür; varsa o, yoksa ham değer kullanılır.
function childAgesDisplayValue(value, displayValue = '') {
  const display = clean(displayValue);
  if (display) return display;
  return clean(value);
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

    if (streak.length >= (appSettings.dndMinDays || 2)) {
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
  if (!appSettings.showSummary) {
    summary.hidden = true;
    summary.innerHTML = '';
    return;
  }
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

/* ---------- Late Check Out ---------- */
function flattenGroupsToRecords(groups) {
  return [...groups.values()].flat();
}

function parsePaxCount(value) {
  const num = parseInt(clean(value), 10);
  return Number.isFinite(num) ? num : 0;
}

function processLateCoutFiles(workbookItems) {
  const items = [...(workbookItems || [])].filter(Boolean);
  if (!items.length) throw new Error('Excel dosyası bulunamadı.');

  let depItem = null;
  let arrItem = null;
  items.forEach(item => {
    const detected = detectWorkbookMode(item.workbook, item.name);
    if (detected?.mode === MODE_DEPARTURES && !depItem) depItem = item;
    else if (detected?.mode === MODE_ARRIVALS && !arrItem) arrItem = item;
  });

  // Algılama başarısızsa dosya adı ipucuna göre tekrar dene.
  if (!depItem || !arrItem) {
    items.forEach(item => {
      const hint = fileNameModeHint(item.name, item.workbook);
      if (hint?.mode === MODE_DEPARTURES && !depItem) depItem = item;
      else if (hint?.mode === MODE_ARRIVALS && !arrItem) arrItem = item;
    });
  }

  if (!depItem) {
    throw new Error('Departures dosyası bulunamadı. Late C/Out için Departures ve Arrivals Excel dosyalarını birlikte yükle.');
  }

  const depRecords = flattenGroupsToRecords(processWorkbook(depItem.workbook, MODE_DEPARTURES));
  const arrRecords = arrItem ? flattenGroupsToRecords(processWorkbook(arrItem.workbook, MODE_ARRIVALS)) : [];

  // Gelen misafir bilgisi oda numarasına göre eşleştirilir.
  const arrivalsByRoom = new Map();
  arrRecords.forEach(record => {
    const roomKey = normalizeRoomId(record.room);
    if (!roomKey || arrivalsByRoom.has(roomKey)) return;
    const adults = parsePaxCount(record.adults);
    const children = parsePaxCount(record.children);
    const ages = clean(record.childAges)
      .split(/[,;.\s]+/)
      .map(part => parseInt(part, 10))
      .filter(age => Number.isFinite(age));
    const hasBaby = ages.some(age => age >= 0 && age <= 3);

    let pax = children > 0 ? `${adults}+${children}` : (adults > 0 ? String(adults) : '');
    if (pax && hasBaby) pax += '+BEBEK';

    arrivalsByRoom.set(roomKey, { pax, eta: clean(record.eta) });
  });

  const lateEtds = lateEtdList();
  const results = depRecords
    .filter(record => lateEtds.includes(clean(record.etd)))
    .map(record => {
      const arrivalInfo = arrivalsByRoom.get(normalizeRoomId(record.room)) || { pax: '', eta: '' };
      return {
        room: clean(record.room),
        checkout: clean(record.etd),
        pax: arrivalInfo.pax,
        checkin: arrivalInfo.eta,
      };
    });

  results.sort((a, b) => a.checkout.localeCompare(b.checkout) || roomSortValue(a.room) - roomSortValue(b.room));

  // Başlıktaki tarih: Departures kayıtlarındaki en yaygın çıkış günü.
  const dateCounts = new Map();
  depRecords.forEach(record => {
    const dateText = clean(record.departure);
    if (dateText) dateCounts.set(dateText, (dateCounts.get(dateText) || 0) + 1);
  });
  lateCoutDateText = [...dateCounts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || '';

  if (!arrItem) {
    setStatus('Arrivals dosyası bulunamadı; C/IN Kişi Sayısı ve Check In Saati boş bırakıldı.', 'error');
  }

  return results;
}

// Şablonun (latecout-sablon.xlsx) birebir ölçüleri: Excel sütun genişlikleri px'e,
// satır yükseklikleri pt->px (x4/3) çevrildi. Tasarım bu doğal boyutta çizilir,
// ekrana ve A4'e transform: scale() ile sığdırılır (zoom Firefox'ta çalışmaz).
const LATECOUT_GRID = {
  colA: 43,        // şablondaki boş A sütunu (5.43 birim)
  spaceRow: 240,   // şablondaki boş 1. satır (180pt)
  titleRow: 130,   // 2. satır: 97.5pt
  headerRow: 124,  // 3. satır: 93pt
  dataRow: 80,     // veri satırları: 60pt
  templateRows: 20,// şablonda görünür veri satırı: 4..23
  cols: [83, 181, 199, 294, 166, 337, 445, 357, 137, 455] // B..K sütunları px
};
const LATECOUT_TABLE_W = LATECOUT_GRID.cols.reduce((sum, w) => sum + w, 0); // 2654
const LATECOUT_TITLE_W = LATECOUT_GRID.cols.slice(0, 7).reduce((sum, w) => sum + w, 0);  // B2:H2 = 1705
const LATECOUT_DATE_W = LATECOUT_TABLE_W - LATECOUT_TITLE_W;                              // I2:K2 = 949

function formatLateCoutDateTitle() {
  // Başlıkta her zaman bugünün tarihi görünür (GG.AA.YYYY).
  const today = new Date();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  return `${dd}.${mm}.${today.getFullYear()}`;
}

function renderLateCoutPreview(results) {
  preview.classList.remove('empty');
  preview.classList.remove('preview-arrivals', 'preview-departures', 'preview-dnd', 'preview-vacant');
  preview.innerHTML = '';

  const page = document.createElement('article');
  page.className = 'sheet-page latecout-page';

  const supervisor = clean(appSettings.lateCoutSupervisor);
  const serviceTime = clean(appSettings.lateCoutServiceTime);
  const serviceTimeText = serviceTime ? `SERVİS SAATİ   ${serviceTime}` : '';

  // Şablonun ızgarası sabit 20 satır (4-23); oda sayısı azsa kalan satırlar boş kalır.
  const totalRows = Math.max(results.length, LATECOUT_GRID.templateRows);

  let rowsHtml = '';
  for (let i = 0; i < totalRows; i += 1) {
    const item = results[i];
    const sideCells = i === 0
      ? `<td class="lc-sup-name" rowspan="${totalRows}">${escapeHtml(supervisor)}</td><td class="lc-sup-time" rowspan="${totalRows}">${escapeHtml(serviceTimeText)}</td>`
      : '';
    rowsHtml += `<tr>
      <td>${i + 1}</td>
      <td>${item ? escapeHtml(item.room) : ''}</td>
      <td>${item ? escapeHtml(item.checkout) : ''}</td>
      <td class="lc-pax">${item ? escapeHtml(item.pax) : ''}</td>
      <td>${item ? escapeHtml(item.checkin) : ''}</td>
      <td class="lc-maid" colspan="2"></td>
      <td class="lc-sup"></td>
      ${sideCells}
    </tr>`;
  }

  const titleText = appSettings.showReportTitle ? appSettings.titleLateCout : '';
  const colsHtml = LATECOUT_GRID.cols.map(w => `<col style="width:${w}px">`).join('');
  const logoHtml = window.LATECOUT_LOGO_B64
    ? `<img class="lc-logo" alt="" src="data:image/png;base64,${window.LATECOUT_LOGO_B64}">`
    : '';

  page.innerHTML = `
    <div class="lc-scaler">
      <div class="latecout-natural">
        <div class="lc-space-row">${logoHtml}</div>
        <div class="lc-grid">
          <div class="lc-titlebar">
            <div class="lc-title">${escapeHtml(titleText)}</div>
            <div class="lc-date">${escapeHtml(formatLateCoutDateTitle())}</div>
          </div>
          <table class="lc-table">
            <colgroup>${colsHtml}</colgroup>
            <thead>
              <tr>
                <th>SIRA</th>
                <th>ODA NO</th>
                <th>CHECK OUT SAATİ</th>
                <th>C/IN KİŞİ SAYISI</th>
                <th>CHECK IN SAATI</th>
                <th colspan="2">GÖREVLİ MAID</th>
                <th colspan="3">GÖREVLİ KAT ŞEFİ -SUPERVİSOR</th>
              </tr>
            </thead>
            <tbody>${rowsHtml}</tbody>
          </table>
          <div class="lc-gap"></div>
          <div class="lc-blackbar"></div>
          ${buildLateCoutStaffHtml(colsHtml)}
        </div>
      </div>
    </div>`;

  preview.appendChild(page);
  applyLateCoutScale(page);
}

// Şablonun alt bölümü (satır 66-98): siyah ayraç, 17:30 servisi personel ızgarası,
// geç çıkacak odalar ve izinli personeller. İçeriği şablonda olduğu gibi sabittir.
function buildLateCoutStaffHtml(colsHtml) {
  const STAFF_ROWS = 31; // şablon satırları 68-98
  let bodyRows = '';
  for (let i = 1; i <= STAFF_ROWS; i += 1) {
    if (i === 1) {
      bodyRows += `<tr>
        <td class="lc-st-num">1</td>
        <td colspan="3" class="lc-st-name"></td>
        <td class="lc-st-f"></td>
        <td class="lc-st-red2">ODA NO</td>
        <td class="lc-st-red2">ÇIKIŞ SAATI</td>
        <td colspan="2" class="lc-st-ij"></td>
        <td class="lc-st-k"></td>
      </tr>`;
    } else if (i <= 28) {
      bodyRows += `<tr>
        <td class="lc-st-num">${i}</td>
        <td colspan="3" class="lc-st-name"></td>
        <td class="lc-st-f"></td>
        <td class="lc-st-g"></td>
        <td class="lc-st-h"></td>
        <td colspan="2" class="lc-st-ij"></td>
        <td class="lc-st-k"></td>
      </tr>`;
    } else if (i === 29) {
      bodyRows += `<tr>
        <td class="lc-st-num">29</td>
        <td colspan="3" class="lc-st-name"></td>
        <td class="lc-st-f"></td>
        <td class="lc-st-blank"></td>
        <td class="lc-st-blank"></td>
        <td colspan="3" class="lc-st-exec lc-st-exec-top">EXEC HOUSEKEEPER</td>
      </tr>`;
    } else if (i === 30) {
      bodyRows += `<tr>
        <td class="lc-st-num">30</td>
        <td colspan="3" class="lc-st-name"></td>
        <td class="lc-st-f"></td>
        <td class="lc-st-blank"></td>
        <td class="lc-st-blank"></td>
        <td colspan="3" class="lc-st-exec">BAYRAM GÜÇLÜ</td>
      </tr>`;
    } else {
      bodyRows += `<tr>
        <td class="lc-st-num">31</td>
        <td colspan="3" class="lc-st-name"></td>
        <td class="lc-st-f"></td>
      </tr>`;
    }
  }
  return `<table class="lc-table lc-staff">
    <colgroup>${colsHtml}</colgroup>
    <tbody>
      <tr class="lc-st-head">
        <td colspan="4" class="lc-st-yellow">SAAT 17:30 SERVİSİ İLE GİDECEK PERSONELLER</td>
        <td class="lc-st-f"></td>
        <td colspan="2" class="lc-st-red">GEÇ ÇIKACAK ODALAT</td>
        <td colspan="2" class="lc-st-amber"></td>
        <td class="lc-st-yellow2">İZİNLİ PERSONELLER</td>
      </tr>
      ${bodyRows}
    </tbody>
  </table>`;
}

// Doğal boyuttaki tasarımı ekran ve yazdırma alanına transform: scale() ile sığdırır.
function applyLateCoutScale(page) {
  const scaler = page.querySelector('.lc-scaler');
  const natural = page.querySelector('.latecout-natural');
  if (!scaler || !natural) return;

  const nw = natural.offsetWidth;
  const nh = natural.offsetHeight;
  if (!nw || !nh) return;

  // Ekran: sayfanın iç genişliğine sığdır
  const availScreen = Math.max(200, page.clientWidth - 32); // .sheet-page padding 16+16
  const screenScale = Math.min(1, availScreen / nw);
  natural.style.transform = `scale(${screenScale})`;
  scaler.style.width = `${Math.round(nw * screenScale)}px`;
  scaler.style.height = `${Math.round(nh * screenScale)}px`;

  // Yazdırma: A4 kullanılabilir alana (sayfa dolgusu düşülmüş) sığdır
  const padMm = clampNumber(parseFloat(appSettings.printPagePadding), 0, 25, 8);
  const mmToPx = 96 / 25.4;
  const availW = (210 - padMm * 2) * mmToPx;
  const availH = (297 - padMm * 2) * mmToPx;
  const printScale = Math.min(1, availW / nw, availH / nh);

  let printStyle = document.getElementById('latecoutPrintScale');
  if (!printStyle) {
    printStyle = document.createElement('style');
    printStyle.id = 'latecoutPrintScale';
    document.head.appendChild(printStyle);
  }
  printStyle.textContent = `@media print {
  .latecout-page .latecout-natural { transform: scale(${printScale}) !important; }
  .latecout-page .lc-scaler { width: ${Math.ceil(nw * printScale)}px !important; height: ${Math.ceil(nh * printScale)}px !important; }
}`;
}

function renderLateCoutSummary(results) {
  if (!appSettings.showSummary) {
    summary.hidden = true;
    summary.innerHTML = '';
    return;
  }
  summary.hidden = false;
  const withArrival = results.filter(item => item.pax || item.checkin).length;
  const etdCards = lateEtdList().map(etd => {
    const count = results.filter(item => item.checkout === etd).length;
    return `<div class="summary-card"><strong>${count}</strong><span>ETD ${escapeHtml(etd)}</span></div>`;
  }).join('');

  summary.innerHTML = `
    <div class="summary-card summary-total"><strong>${results.length}</strong><span>Late C/Out oda</span></div>
    ${etdCards}
    <div class="summary-card"><strong>${withArrival}</strong><span>C/IN eşleşen oda</span></div>`;
}

function renderLateCoutOutput(message = '') {
  assignmentPanel.hidden = true;
  greenPanel.hidden = true;
  updateCurrentRoomsPanel();
  chiefControls.innerHTML = '';
  assignmentStatus.innerHTML = '';
  renderLateCoutPreview(lateCoutResults);
  renderLateCoutSummary(lateCoutResults);
  setButtons({ printable: lateCoutResults.length > 0, clearable: Boolean(lastWorkbooks.length || lastWorkbook) });
  if (message) setStatus(message, lateCoutResults.length ? 'ok' : 'error');
}

function lateCoutExcelFileName() {
  const now = new Date();
  const datePart = now.toLocaleDateString('tr-TR').replace(/\./g, '-');
  return `Late_Check_Out_${datePart}.xlsx`;
}

/* Şablon xlsx'in içindeki sheet XML'ini doğrudan düzenler.
   Böylece yazı tipleri, kenarlıklar, dolgular, birleştirmeler ve diğer sayfalar
   %100 korunur; sadece hedef hücrelerin içeriği değişir. */
const XLSX_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function xmlCreateCellChild(sheetDoc, tag, text) {
  const el = sheetDoc.createElementNS(XLSX_MAIN_NS, tag);
  if (text !== undefined) el.textContent = text;
  return el;
}

function xmlSetCellText(sheetDoc, ref, text) {
  const cell = xmlFindCell(sheetDoc, ref);
  if (!cell) return;
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  if (text !== '') {
    cell.setAttribute('t', 'inlineStr');
    const is = xmlCreateCellChild(sheetDoc, 'is');
    is.appendChild(xmlCreateCellChild(sheetDoc, 't', text));
    cell.appendChild(is);
  } else {
    cell.removeAttribute('t');
  }
}

function xmlSetCellDate(sheetDoc, ref, date) {
  const cell = xmlFindCell(sheetDoc, ref);
  if (!cell) return;
  while (cell.firstChild) cell.removeChild(cell.firstChild);
  cell.removeAttribute('t');
  const serial = Math.round((date - new Date(1899, 11, 30)) / 86400000);
  cell.appendChild(xmlCreateCellChild(sheetDoc, 'v', String(serial)));
}

function xmlFindCell(sheetDoc, ref) {
  const cells = sheetDoc.getElementsByTagName('c');
  for (let i = 0; i < cells.length; i += 1) {
    if (cells[i].getAttribute('r') === ref) return cells[i];
  }
  // Hücre yoksa satırın içine oluştur (şablon satır aralığı dışına taşan ekstra odalar için).
  const rowMatch = ref.match(/^([A-Z]+)(\d+)$/);
  if (!rowMatch) return null;
  const rows = sheetDoc.getElementsByTagName('row');
  for (let i = 0; i < rows.length; i += 1) {
    if (rows[i].getAttribute('r') === rowMatch[2]) {
      const cell = xmlCreateCellChild(sheetDoc, 'c');
      cell.setAttribute('r', ref);
      rows[i].appendChild(cell);
      return cell;
    }
  }
  return null;
}

// Hazır tasarım şablonunu (latecout-sablon.xlsx) açıp verileri içine yazar.
function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

async function downloadLateCoutExcelFromTemplate() {
  if (typeof JSZip === 'undefined') throw new Error('JSZip yüklenemedi.');
  // Şablon önce gömülü kopyadan okunur: dosya file:// ile açıldığında ya da
  // şablon dosyası sunucuda bulunmadığında fetch düşer ve sade tasarıma geçilirdi.
  let buffer;
  if (window.LATECOUT_TEMPLATE_XLSX_B64) {
    buffer = base64ToArrayBuffer(window.LATECOUT_TEMPLATE_XLSX_B64);
  } else {
    const response = await fetch('latecout-sablon.xlsx');
    if (!response.ok) throw new Error('Şablon dosyası okunamadı.');
    buffer = await response.arrayBuffer();
  }
  const zip = await JSZip.loadAsync(buffer);
  const parser = new DOMParser();

  // İlk sayfanın XML dosyasını bul.
  const workbookDoc = parser.parseFromString(await zip.file('xl/workbook.xml').async('text'), 'application/xml');
  const firstSheetEl = workbookDoc.getElementsByTagName('sheet')[0];
  const rid = firstSheetEl.getAttribute('r:id') || firstSheetEl.getAttribute('id');
  const relsDoc = parser.parseFromString(await zip.file('xl/_rels/workbook.xml.rels').async('text'), 'application/xml');
  let sheetPath = 'xl/worksheets/sheet1.xml';
  const rels = relsDoc.getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i += 1) {
    if (rels[i].getAttribute('Id') === rid) {
      sheetPath = `xl/${rels[i].getAttribute('Target').replace(/^\//, '')}`;
      break;
    }
  }

  const sheetDoc = parser.parseFromString(await zip.file(sheetPath).async('text'), 'application/xml');

  // C, D, E, F kolonları: ODA NO, CHECK OUT SAATİ, C/IN KİŞİ SAYISI, CHECK IN SAATI.
  const DATA_COLS = ['C', 'D', 'E', 'F'];
  const TEMPLATE_DATA_ROWS = 62; // 4-65 arası şablon veri bölgesi; önce temizle sonra yaz.
  const rowCount = Math.max(lateCoutResults.length, TEMPLATE_DATA_ROWS);
  for (let i = 0; i < rowCount; i += 1) {
    const item = lateCoutResults[i];
    const values = item
      ? [item.room, item.checkout, item.pax, item.checkin]
      : ['', '', '', ''];
    DATA_COLS.forEach((col, idx) => {
      xmlSetCellText(sheetDoc, `${col}${4 + i}`, String(values[idx] ?? ''));
    });
  }

  // I2 hücresine bugünün tarihi yazılır.
  xmlSetCellDate(sheetDoc, 'I2', startOfDay(new Date()));

  // J4 (kat şefi) ve K4 (servis saati): ayarlardaki değerleri yaz ki Excel çıktısı
  // ekrandaki/PDF'teki önizlemeyle birebir aynı olsun.
  const supervisorText = clean(appSettings.lateCoutSupervisor);
  const serviceTimeText = clean(appSettings.lateCoutServiceTime);
  xmlSetCellText(sheetDoc, 'J4', supervisorText ? `${supervisorText} ` : '');
  xmlSetCellText(sheetDoc, 'K4', serviceTimeText ? `SERVİS SAATİ   ${serviceTimeText}` : '');

  zip.file(sheetPath, new XMLSerializer().serializeToString(sheetDoc));
  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  const fileName = lateCoutExcelFileName();
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(link.href);
  setStatus(`${fileName} şablon tasarımıyla indirildi.`, 'ok');
}

// Şablon okunamazsa (örn. dosya file:// ile açılmışsa) basit tasarımla üretir.
function downloadLateCoutExcelSimple() {
  const headers = ['SIRA', 'ODA NO', 'CHECK OUT SAATİ', 'C/IN KİŞİ SAYISI', 'CHECK IN SAATI', 'GÖREVLİ MAID', 'GÖREVLİ KAT ŞEFİ -SUPERVİSOR'];
  const rows = [
    [appSettings.titleLateCout, '', '', '', '', '', formatLateCoutDateTitle()],
    headers,
    ...lateCoutResults.map((item, index) => [index + 1, item.room, item.checkout, item.pax, item.checkin, '', '']),
  ];

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [6, 9, 12, 12, 12, 18, 22].map(wch => ({ wch }));
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];
  ws['!rows'] = [{ hpt: 24 }, { hpt: 22 }];

  const titleStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 14, border: false });
  titleStyle.alignment.horizontal = 'left';
  const headerStyle = excelCellStyle({ fill: hexToRgb(appSettings.headerColor, 'C8755C'), bold: true, size: 10 });
  const bodyStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 11 });
  const dateStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 11, border: false });

  ws[cellRef(0, 0)].s = titleStyle;
  ws[cellRef(0, 6)].s = dateStyle;
  for (let c = 0; c < headers.length; c += 1) {
    ws[cellRef(1, c)].s = headerStyle;
  }
  lateCoutResults.forEach((item, index) => {
    const rowIndex = 2 + index;
    ws['!rows'][rowIndex] = { hpt: appSettings.excelRowHeight > 0 ? appSettings.excelRowHeight : 22 };
    for (let c = 0; c < headers.length; c += 1) {
      ws[cellRef(rowIndex, c)].s = bodyStyle;
    }
  });
  ws['!margins'] = { left: 0.25, right: 0.25, top: 0.25, bottom: 0.25, header: 0, footer: 0 };
  ws['!pageSetup'] = { paperSize: 9, orientation: 'portrait', fitToWidth: 1, fitToHeight: 0 };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'LATE CHECK OUT');
  const fileName = lateCoutExcelFileName();
  XLSX.writeFile(wb, fileName, { bookType: 'xlsx', cellStyles: true });
  setStatus(`${fileName} indirildi.`, 'ok');
}

async function downloadLateCoutExcel() {
  if (!lateCoutResults.length) {
    setStatus('Late C/Out Excel için kayıt bulunamadı.', 'error');
    return;
  }

  try {
    await downloadLateCoutExcelFromTemplate();
  } catch (error) {
    console.warn('Şablon ile üretim başarısız, basit tasarıma dönülüyor.', error);
    downloadLateCoutExcelSimple();
  }
}
/* ---------- Late Check Out sonu ---------- */

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
    .filter(item => item.text);

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
    const roomClass = vacantColumnText(rowItems, 58, 90, { separator: ' ' });
    const occupancy = vacantColumnText(rowItems, 90, 116, { separator: ' / ' });
    const foStatus = vacantColumnText(rowItems, 116, 140, { separator: ' / ' });
    const nightsVacant = vacantColumnText(rowItems, 140, 177, { separator: ' / ' });
    const name = vacantColumnText(rowItems, 177, 278, { separator: ' / ' });
    const arrival = vacantColumnText(rowItems, 278, 318, { separator: ' / ' });
    const departure = vacantColumnText(rowItems, 318, 358, { separator: ' / ' });
    const reservationStatus = normalizeVacantStatus(vacantColumnText(rowItems, 358, 420, { separator: ' / ' }));
    const adults = vacantColumnText(rowItems, 420, 448, { separator: ' / ' });
    const children = vacantColumnText(rowItems, 448, 475, { separator: ' / ' });
    const discrepantStatus = vacantColumnText(rowItems, 475, 515, { separator: ' / ' });
    const nextBlocked = vacantColumnText(rowItems, 515, 560, { separator: ' / ' });

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
  return { name: file.name, records, type: MODE_VACANT, lastModified: Number(file.lastModified || 0), size: Number(file.size || 0) };
}


function isVacantCleanRecord(record) {
  const roomType = canonical(record?.roomType || '');
  return roomType.split(/\s+/).includes('vac');
}

function sortedGreenRoomList() {
  return [...greenRooms].sort((a, b) => roomSortValue(a) - roomSortValue(b));
}

function arrivalsGreenMatchCount() {
  if (currentMode !== MODE_ARRIVALS || !originalGroups.size || !greenRooms.size) return 0;
  return allOriginalRecords().filter(record => isGreenRoom(record)).length;
}

function updateVacantRoomsStatus({ error = '' } = {}) {
  if (!vacantRoomsStatus) return;

  if (error) {
    vacantRoomsStatus.textContent = error;
    vacantRoomsStatus.className = 'vacant-rooms-status error';
    return;
  }

  if (!greenRooms.size) {
    vacantRoomsStatus.textContent = 'Vacant Rooms PDF yükle; Room Type = VAC olan odalar otomatik bulunur.';
    vacantRoomsStatus.className = 'vacant-rooms-status';
    return;
  }

  const matchCount = arrivalsGreenMatchCount();
  const source = vacantRoomsFileName ? `${vacantRoomsFileName}: ` : '';
  vacantRoomsStatus.textContent = `${source}${greenRooms.size} VAC oda bulundu. Arrivals listesinde ${matchCount} oda temiz olarak eşleşti.`;
  vacantRoomsStatus.className = 'vacant-rooms-status ok';
}

async function handleArrivalsVacantPdfs(files) {
  const fileList = [...(files || [])].filter(Boolean);
  if (!fileList.length) return;
  if (currentMode !== MODE_ARRIVALS) {
    updateVacantRoomsStatus({ error: 'Temiz oda PDF yalnızca Arrivals bölümünde kullanılır.' });
    return;
  }

  try {
    if (vacantRoomsPdfInput) vacantRoomsPdfInput.disabled = true;
    if (vacantRoomsStatus) {
      vacantRoomsStatus.textContent = fileList.length === 1 ? `${fileList[0].name} okunuyor...` : `${fileList.length} Vacant PDF okunuyor...`;
      vacantRoomsStatus.className = 'vacant-rooms-status loading';
    }

    const uniqueRooms = new Set();
    const usedNames = [];
    for (const file of fileList) {
      try {
        const item = await readVacantPdfFile(file);
        item.records
          .filter(isVacantCleanRecord)
          .map(record => normalizeRoomId(record.room))
          .filter(Boolean)
          .forEach(room => uniqueRooms.add(room));
        usedNames.push(item.name);
      } catch (error) {
        console.warn('Vacant PDF atlandı:', file.name, error);
      }
    }

    if (!uniqueRooms.size) throw new Error('Vacant Rooms PDF dosyalarında Room Type = VAC olan oda bulunamadı.');

    greenRooms = uniqueRooms;
    vacantRoomsFileName = usedNames.length <= 1 ? (usedNames[0] || '') : `${usedNames.length} Vacant PDF`;
    if (greenRoomsInput) greenRoomsInput.value = sortedGreenRoomList().join(', ');

    if (originalGroups.size) {
      updateOutput(`${vacantRoomsFileName}: ${greenRooms.size} benzersiz VAC oda alındı; Arrivals temiz odaları yeşil işaretlendi.`);
    } else {
      updateVacantRoomsStatus();
      setStatus(`${vacantRoomsFileName}: ${greenRooms.size} benzersiz VAC oda hazır. Şimdi Arrivals Excel dosyasını yükle.`, 'ok');
    }
  } catch (error) {
    console.error(error);
    greenRooms = new Set();
    vacantRoomsFileName = '';
    if (greenRoomsInput) greenRoomsInput.value = '';
    updateVacantRoomsStatus({ error: error.message || 'Vacant Rooms PDF okunamadı.' });
    setStatus(error.message || 'Vacant Rooms PDF okunamadı.', 'error');
  } finally {
    if (vacantRoomsPdfInput) {
      vacantRoomsPdfInput.disabled = currentMode !== MODE_ARRIVALS || !originalGroups.size;
      vacantRoomsPdfInput.value = '';
    }
  }
}

async function handleArrivalsVacantPdf(file) {
  return handleArrivalsVacantPdfs([file]);
}

function mergeVacantRecordItems(items) {
  const sortedItems = [...(items || [])].filter(Boolean)
    .sort((a, b) => Number(b?.lastModified || 0) - Number(a?.lastModified || 0));
  const byRoom = new Map();
  sortedItems.forEach(item => {
    (item.records || []).forEach(record => {
      const key = normalizeRoomId(record.room);
      if (key && !byRoom.has(key)) byRoom.set(key, record);
    });
  });
  const allRecords = [...byRoom.values()];
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
    childAges: mode === MODE_ARRIVALS
      ? childAgesDisplayValue(row[map.childAges], displayRow[map.childAges])
      : formatChildAgesValue(row[map.childAges], displayRow[map.childAges], children),
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
  const items = [...(workbookItems || [])].filter(Boolean)
    .sort((a, b) => Number(b?.lastModified || 0) - Number(a?.lastModified || 0));
  if (!items.length) throw new Error('Excel dosyası bulunamadı.');

  // Aynı oda birden fazla parça/raporda varsa en yeni dosyadaki kayıt korunur.
  const byRoom = new Map();
  items.forEach(item => {
    const workbook = item.workbook || item;
    const groups = processWorkbook(workbook, mode);
    groups.forEach(records => records.forEach(record => {
      const key = normalizeRoomId(record.room);
      if (key && !byRoom.has(key)) byRoom.set(key, record);
    }));
  });

  const allRecords = [...byRoom.values()];
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
      printLineCount(record.nextBlocked, 10),
      printLineCount(record.notes, 32)
    );
    return Math.min(4, Math.max(1, maxLines));
  }

  return 1;
}

function splitNumberedRecordsForPrint(numberedRecords) {
  const rowsPerPage = currentMode === MODE_ARRIVALS
    ? PRINT_ROWS_PER_PAGE.arrivals
    : currentMode === MODE_VACANT
      ? PRINT_ROWS_PER_PAGE.vacant
      : PRINT_ROWS_PER_PAGE.departures;

  const pages = [];
  for (let i = 0; i < numberedRecords.length; i += rowsPerPage) {
    pages.push(numberedRecords.slice(i, i + rowsPerPage));
  }
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

function printedRowsPerPage() {
  if (currentMode === MODE_ARRIVALS) return PRINT_ROWS_PER_PAGE.arrivals;
  if (currentMode === MODE_VACANT) return PRINT_ROWS_PER_PAGE.vacant;
  return PRINT_ROWS_PER_PAGE.departures;
}

function reportTitleForGroup(groupName) {
  if (groupName === 'Ofis') return '';
  const numeric = String(groupName || '').match(/\d+/)?.[0] || String(groupName || '').replace(/ler$/i, '');
  if (currentMode === MODE_DEPARTURES) return appSettings.titleDepartures;
  if (currentMode === MODE_ARRIVALS) return appSettings.titleArrivals;
  if (currentMode === MODE_VACANT) return appSettings.titleVacant;
  return String(groupName || '');
}

function blankTableRows(count, columnCount) {
  if (!count || count < 1) return '';
  const cells = Array.from({ length: columnCount }, (_, index) => `<td${index === columnCount - 1 ? ' class="notes"' : ''}>&nbsp;</td>`).join('');
  return Array.from({ length: count }, () => `<tr class="blank-fill-row">${cells}</tr>`).join('');
}

function renderPage(records, groupName, options = {}) {
  const { padToRows = 0 } = options;
  const isArrivals = currentMode === MODE_ARRIVALS;
  const isVacant = currentMode === MODE_VACANT;
  const columnCount = isVacant ? 15 : (isArrivals ? 12 : 11);
  const page = document.createElement('article');
  page.className = `sheet-page ${groupName === 'Ofis' ? 'office-page' : 'auto-fill-page'}`;
  page.dataset.group = groupName;

  const rowsHtml = records.map(({ record, rowNumber }) => {
    const etdClass = currentMode === MODE_DEPARTURES && appSettings.etdHighlightEnabled && record.etd === ETD_HIGHLIGHT ? 'etd-highlight' : '';
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
        <td class="notes">${escapeHtml(record.notes)}</td>
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

  const emptyRowsHtml = blankTableRows(
    Math.max(0, padToRows - records.length),
    columnCount,
  );

  const colgroup = isVacant
    ? `<colgroup>
        <col class="idx"><col class="room"><col class="vac-class"><col class="vac-type"><col class="vac-fo"><col class="vac-nights"><col class="name"><col class="date"><col class="date"><col class="vac-status"><col class="small"><col class="small"><col class="vac-disc"><col class="date"><col class="notes">
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
        <th>NOTLAR</th>
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

  const officeHeader = groupName === 'Ofis'
    ? `<tr class="office-head-row"><th colspan="${columnCount}">OFİS</th></tr>`
    : '';
  const reportTitle = appSettings.showReportTitle ? reportTitleForGroup(groupName) : '';
  const titleHtml = reportTitle ? `<div class="report-title">${escapeHtml(reportTitle)}</div>` : '';

  page.innerHTML = `
    ${titleHtml}
    <div class="table-wrap">
      <table class="departure-table ${isVacant ? 'vacant-table' : (isArrivals ? 'arrival-table' : 'departure-mode-table')}">
        ${colgroup}
        <thead>${officeHeader}${header}</thead>
        <tbody>${rowsHtml}${emptyRowsHtml}</tbody>
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

    // Arrivals ve Departures Ofis çıktısı da sayfa başına satır ayarına uyar.
    // Varsayılan 38 satırdır; son sayfa gerekirse boş satırlarla tamamlanır.
    if (groupName === 'Ofis') {
      const rowsPerPage = printedRowsPerPage();
      const pages = splitNumberedRecordsForPrint(numbered);
      pages.forEach(pageRecords => renderPage(pageRecords, groupName, {
        padToRows: appSettings.fillBlankRows ? rowsPerPage : 0,
      }));
      return;
    }

    // PDF/yazdırmada her 1000/2000/3000/4000/5000 grubu ayrı A4 sayfa olarak çıkar.
    // Az kayıt varsa kalan alan boş tablo satırlarıyla doldurulur; sonraki grup aynı sayfaya girmez.
    const rowsPerPage = printedRowsPerPage();
    const pages = splitNumberedRecordsForPrint(numbered);
    pages.forEach(pageRecords => renderPage(pageRecords, groupName, { padToRows: appSettings.fillBlankRows ? rowsPerPage : 0 }));
  });
}

function renderSummary(groups) {
  if (!appSettings.showSummary) {
    summary.hidden = true;
    summary.innerHTML = '';
    return;
  }
  summary.hidden = false;
  const groupEntries = [...groups.entries()];
  const totalRows = groupEntries.reduce((sum, [, rows]) => sum + rows.length, 0);
  const totalLate = groupEntries.reduce((sum, [, rows]) => sum + countLate(rows), 0);
  const lateTotalHtml = currentMode === MODE_DEPARTURES ? `<small>Total Late: ${totalLate}</small>` : '';
  const cleanTotal = currentMode === MODE_ARRIVALS ? groupEntries.reduce((sum, [, rows]) => sum + rows.filter(isGreenRoom).length, 0) : 0;
  const cleanCard = currentMode === MODE_ARRIVALS && greenRooms.size
    ? `<div class="summary-card summary-clean"><strong>${cleanTotal}</strong><span>Temiz / VAC</span><small>Vacant PDF ile eşleşen</small></div>`
    : '';
  const totalCard = `<div class="summary-card summary-total"><strong>${totalRows}</strong><span>Toplam</span>${lateTotalHtml}</div>`;
  const groupCards = groupEntries
    .map(([name, rows]) => {
      const lateHtml = currentMode === MODE_DEPARTURES ? `<small>Total Late: ${countLate(rows)}</small>` : '';
      const cleanCount = currentMode === MODE_ARRIVALS ? rows.filter(isGreenRoom).length : 0;
      const cleanHtml = currentMode === MODE_ARRIVALS
        ? `<small class="group-clean-count">Total Temiz: ${cleanCount}</small>`
        : '';
      return `<div class="summary-card"><strong>${rows.length}</strong><span>${escapeHtml(name)}</span>${cleanHtml}${lateHtml}</div>`;
    })
    .join('');
  summary.innerHTML = `${totalCard}${cleanCard}${groupCards}`;
}

function setButtons({ printable = false, clearable = false } = {}) {
  printBtn.disabled = !printable;
  excelBtn.disabled = !printable;
  officeBtn.disabled = !clearable || currentMode === MODE_DND || currentMode === MODE_LATECOUT;
  officeExcelBtn.disabled = !clearable || currentMode === MODE_DND || currentMode === MODE_LATECOUT;
  clearBtn.disabled = !clearable;
  greenRoomsInput.disabled = !clearable || currentMode !== MODE_ARRIVALS;
  if (vacantRoomsPdfInput) vacantRoomsPdfInput.disabled = !clearable || currentMode !== MODE_ARRIVALS;
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
    const cleanCount = currentMode === MODE_ARRIVALS ? displayRecords.filter(isGreenRoom).length : 0;
    const originalCount = originalRecords.length;
    const disabled = originalCount === 0;
    const selectableSections = sections.filter(section => section.sourceGroup !== groupName);

    const leaveToggleHtml = canBeLeave ? `
          <label class="leave-toggle">
            <input type="checkbox" data-leave-group="${groupName}" ${isLeave ? 'checked' : ''} ${disabled ? 'disabled' : ''}>
            <span>İzinli</span>
          </label>` : '';

    const selectHtml = !isLeave && selectableSections.length ? `
      <div class="assign-head-row">
        <label class="assign-label">Eklenecek bölümler</label>
        <div class="assign-actions">
          <button type="button" class="mini-assign-btn" data-assign-all="${groupName}">Boştakileri al</button>
          <button type="button" class="mini-assign-btn ghost" data-assign-clear="${groupName}">Temizle</button>
        </div>
      </div>
      <div class="assignment-checklist" data-target="${groupName}">
        ${selectableSections.map(section => {
          const assignedTarget = sectionAssignments.get(section.key);
          const checked = assignedTarget === groupName ? 'checked' : '';
          const disabled = assignedTarget && assignedTarget !== groupName ? 'disabled' : '';
          const ownerText = assignedTarget && assignedTarget !== groupName ? ` • ${assignedTarget} aldı` : '';
          const label = `${section.sectionName} (${section.records.length} oda${formatLateOption(section.records)})${ownerText}`;
          return `<label class="assignment-check ${disabled ? 'is-disabled' : ''}">
            <input type="checkbox" value="${escapeHtml(section.key)}" data-target="${groupName}" ${checked} ${disabled}>
            <span>${escapeHtml(label)}</span>
          </label>`;
        }).join('')}
      </div>` : '';

    const leaveNoteLate = lateText(originalRecords);
    const leaveNote = leaveNoteLate ? `İzinli: ${originalCount} oda / ${leaveNoteLate}` : `İzinli: ${originalCount} oda`;

    return `
      <div class="chief-card ${isLeave ? 'is-leave' : ''} ${displayCount === 0 ? 'is-empty' : ''}">
        <div class="chief-top">
          <div>
            <strong>${escapeHtml(groupName)}</strong>
            <span>${displayCount} oda</span>
            ${currentMode === MODE_ARRIVALS ? `<span class="clean-count-inline">Total Temiz: ${cleanCount}</span>` : ''}
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
  updateVacantRoomsStatus();
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
    const labels = unassigned.map(section => `${section.sectionName}`).join(', ');
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

    if (currentMode === MODE_LATECOUT) {
      dndResults = [];
      const itemsToProcess = lastWorkbooks.length ? lastWorkbooks : [{ workbook: lastWorkbook, name: lastFileName }];
      lateCoutResults = processLateCoutFiles(itemsToProcess);
      originalGroups = new Map([['Late Check Out', lateCoutResults]]);
      printableGroups = new Map();
      resetAssignmentsForNewData();
      renderLateCoutOutput(message || `${lastFileName} Late Check Out olarak işlendi.`);
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
  // raw: true -> HTML tabanlı .xls dosyalarında SheetJS'in "6, 9, 12" gibi
  // Child Ages metinlerini otomatik tarihe çevirmesini engeller.
  // Gerçek .xlsx dosyalarında bu seçenek etkisizdir, hücre tipleri dosyadan gelir.
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: currentMode !== MODE_DND, raw: true });
  return { name: file.name, workbook, lastModified: Number(file.lastModified || 0), size: Number(file.size || 0) };
}

function setDownloadsAutoStatus(message, type = '') {
  if (!downloadsAutoStatus) return;
  downloadsAutoStatus.textContent = message;
  downloadsAutoStatus.className = type;
}

function supportsDownloadsFolderAccess() {
  return typeof window.showDirectoryPicker === 'function' && 'indexedDB' in window;
}

function openDownloadsDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DOWNLOADS_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(DOWNLOADS_DB_STORE)) db.createObjectStore(DOWNLOADS_DB_STORE);
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveDownloadsDirectoryHandle(handle) {
  const db = await openDownloadsDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(DOWNLOADS_DB_STORE, 'readwrite');
    tx.objectStore(DOWNLOADS_DB_STORE).put(handle, DOWNLOADS_HANDLE_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

async function restoreDownloadsDirectoryHandle() {
  const db = await openDownloadsDb();
  const handle = await new Promise((resolve, reject) => {
    const tx = db.transaction(DOWNLOADS_DB_STORE, 'readonly');
    const request = tx.objectStore(DOWNLOADS_DB_STORE).get(DOWNLOADS_HANDLE_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return handle;
}

async function directoryHasReadPermission(handle, requestIfNeeded = false) {
  if (!handle) return false;
  const opts = { mode: 'read' };
  if (typeof handle.queryPermission === 'function') {
    const state = await handle.queryPermission(opts);
    if (state === 'granted') return true;
  }
  if (requestIfNeeded && typeof handle.requestPermission === 'function') {
    return (await handle.requestPermission(opts)) === 'granted';
  }
  return false;
}

function autoFileLabel(item) { return item?.file?.name || '—'; }
function autoFilesLabel(items = []) {
  const list = [...items].filter(Boolean);
  if (!list.length) return '—';
  if (list.length === 1) return autoFileLabel(list[0]);
  return `${list.length} dosya: ${list.map(autoFileLabel).join(', ')}`;
}
function updateDownloadsButtons() {
  if (downloadsRefreshBtn) downloadsRefreshBtn.disabled = !downloadsDirectoryHandle || downloadsScanBusy;
  if (downloadsConnectBtn) downloadsConnectBtn.disabled = downloadsScanBusy;
}
function downloadsSummaryText() {
  const parts = [];
  if (autoDownloadFiles.arrivals.length) parts.push(`Arrivals: ${autoFilesLabel(autoDownloadFiles.arrivals)}`);
  if (autoDownloadFiles.departures.length) parts.push(`Departures: ${autoFilesLabel(autoDownloadFiles.departures)}`);
  if (autoDownloadFiles.vacant.length) parts.push(`Temiz odalar: ${autoFilesLabel(autoDownloadFiles.vacant)}`);
  return parts.length ? parts.join(' • ') : 'Uygun Arrivals, Departures veya Vacant dosyası bulunamadı.';
}
function addAutoFile(key, file) {
  if (!file || !autoDownloadFiles[key]) return;
  const item = { file, lastModified: Number(file.lastModified || 0), size: Number(file.size || 0) };
  const fingerprint = `${file.name}::${item.lastModified}::${item.size}`;
  if (autoDownloadFiles[key].some(existing => `${existing.file.name}::${existing.lastModified}::${existing.size}` === fingerprint)) return;
  autoDownloadFiles[key].push(item);
  autoDownloadFiles[key].sort((a, b) => b.lastModified - a.lastModified);
  autoDownloadFiles[key] = autoDownloadFiles[key].slice(0, AUTO_FILES_PER_TYPE);
}

async function readWorkbookForAutoDetection(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellDates: true, raw: true });
}

async function buildDownloadsSignature(handle) {
  const entries = [];
  for await (const entry of handle.values()) {
    if (entry.kind !== 'file') continue;
    const name = entry.name || '';
    if (!/\.(xlsx|xls|csv|pdf)$/i.test(name)) continue;
    try {
      const file = await entry.getFile();
      entries.push(`${file.name}:${file.lastModified}:${file.size}`);
    } catch (error) { console.warn('Dosya bilgisi okunamadı:', name, error); }
  }
  return entries.sort().join('|');
}

async function collectDownloadsCandidates(handle) {
  const files = [];
  for await (const entry of handle.values()) {
    if (entry.kind !== 'file') continue;
    if (!/\.(xlsx|xls|csv|pdf)$/i.test(entry.name || '')) continue;
    try { files.push(await entry.getFile()); }
    catch (error) { console.warn('İndirilenler dosyası açılamadı:', entry.name, error); }
  }
  files.sort((a, b) => Number(b.lastModified || 0) - Number(a.lastModified || 0));
  return files;
}

async function classifyDownloadsFiles(files) {
  autoDownloadFiles = { arrivals: [], departures: [], vacant: [] };
  const ambiguousExcels = [];
  const ambiguousPdfs = [];

  for (const file of files) {
    if (/\.pdf$/i.test(file.name)) {
      const hint = fileNameModeHint(file.name);
      if (hint?.mode === MODE_VACANT) addAutoFile('vacant', file);
      else ambiguousPdfs.push(file);
      continue;
    }
    if (!/\.(xlsx|xls|csv)$/i.test(file.name)) continue;
    const hint = fileNameModeHint(file.name);
    if (hint?.mode === MODE_ARRIVALS) addAutoFile('arrivals', file);
    else if (hint?.mode === MODE_DEPARTURES) addAutoFile('departures', file);
    else ambiguousExcels.push(file);
  }

  for (const file of ambiguousExcels.slice(0, 15)) {
    if (autoDownloadFiles.arrivals.length >= AUTO_FILES_PER_TYPE && autoDownloadFiles.departures.length >= AUTO_FILES_PER_TYPE) break;
    try {
      const workbook = await readWorkbookForAutoDetection(file);
      const detected = detectWorkbookMode(workbook, file.name);
      if (detected?.mode === MODE_ARRIVALS && detected.confidence >= 0.7) addAutoFile('arrivals', file);
      if (detected?.mode === MODE_DEPARTURES && detected.confidence >= 0.7) addAutoFile('departures', file);
    } catch (error) { console.warn('Excel otomatik sınıflandırılamadı:', file.name, error); }
  }

  if (autoDownloadFiles.vacant.length < AUTO_FILES_PER_TYPE) {
    for (const file of ambiguousPdfs.slice(0, 8)) {
      if (autoDownloadFiles.vacant.length >= AUTO_FILES_PER_TYPE) break;
      try {
        const item = await readVacantPdfFile(file);
        if (item.records.some(isVacantCleanRecord)) addAutoFile('vacant', file);
      } catch (error) { /* Vacant değil; geç. */ }
    }
  }
}

async function applyAutoDownloadsForCurrentMode({ silentMissing = false } = {}) {
  if (downloadsAutoLoadBusy) return;
  if (![MODE_ARRIVALS, MODE_DEPARTURES].includes(currentMode)) return;
  const mainItems = currentMode === MODE_ARRIVALS ? autoDownloadFiles.arrivals : autoDownloadFiles.departures;
  if (!mainItems.length) {
    if (!silentMissing) setDownloadsAutoStatus(`${modeLabel()} için uygun Excel bulunamadı. ${downloadsSummaryText()}`, 'error');
    return;
  }
  downloadsAutoLoadBusy = true;
  try {
    await handleFiles(mainItems.map(item => item.file));
    if (currentMode === MODE_ARRIVALS && autoDownloadFiles.vacant.length) {
      await handleArrivalsVacantPdfs(autoDownloadFiles.vacant.map(item => item.file));
    }
    setDownloadsAutoStatus(`Otomatik bağlı • ${downloadsSummaryText()}`, 'ok');
  } catch (error) {
    console.error(error);
    setDownloadsAutoStatus(error.message || 'İndirilenler dosyaları otomatik yüklenemedi.', 'error');
  } finally { downloadsAutoLoadBusy = false; }
}

async function scanDownloadsFolder({ force = false, autoApply = true } = {}) {
  if (!downloadsDirectoryHandle || downloadsScanBusy) return;
  const allowed = await directoryHasReadPermission(downloadsDirectoryHandle, false);
  if (!allowed) {
    setDownloadsAutoStatus('İndirilenler klasörü kayıtlı, fakat okuma izni gerekiyor. Bağla düğmesine bas.', 'error');
    updateDownloadsButtons();
    return;
  }
  downloadsScanBusy = true;
  updateDownloadsButtons();
  try {
    const signature = await buildDownloadsSignature(downloadsDirectoryHandle);
    if (!force && signature === downloadsLastSignature) return;
    downloadsLastSignature = signature;
    setDownloadsAutoStatus('İndirilenler taranıyor…', 'loading');
    const files = await collectDownloadsCandidates(downloadsDirectoryHandle);
    await classifyDownloadsFiles(files);
    setDownloadsAutoStatus(`Otomatik bağlı • ${downloadsSummaryText()}`, 'ok');
    if (autoApply) await applyAutoDownloadsForCurrentMode({ silentMissing: true });
  } catch (error) {
    console.error(error);
    setDownloadsAutoStatus(error.message || 'İndirilenler klasörü taranamadı.', 'error');
  } finally { downloadsScanBusy = false; updateDownloadsButtons(); }
}

function startDownloadsAutoScanTimer() {
  if (downloadsScanTimer) clearInterval(downloadsScanTimer);
  downloadsScanTimer = setInterval(() => {
    scanDownloadsFolder({ force: false, autoApply: true }).catch(console.error);
  }, DOWNLOADS_SCAN_INTERVAL_MS);
}

async function connectDownloadsFolder() {
  if (!supportsDownloadsFolderAccess()) {
    setDownloadsAutoStatus('Bu özellik masaüstü Chrome/Edge üzerinde HTTPS veya localhost ile çalışır.', 'error');
    return;
  }
  try {
    let handle = downloadsDirectoryHandle;
    if (handle) {
      const granted = await directoryHasReadPermission(handle, true);
      if (!granted) handle = null;
    }
    if (!handle) {
      handle = await window.showDirectoryPicker({ id: 'opera-downloads', mode: 'read', startIn: 'downloads' });
      const granted = await directoryHasReadPermission(handle, true);
      if (!granted) throw new Error('Klasör okuma izni verilmedi.');
    }
    downloadsDirectoryHandle = handle;
    await saveDownloadsDirectoryHandle(handle);
    downloadsLastSignature = '';
    updateDownloadsButtons();
    await scanDownloadsFolder({ force: true, autoApply: true });
    startDownloadsAutoScanTimer();
  } catch (error) {
    if (error?.name === 'AbortError') { setDownloadsAutoStatus('Klasör seçimi iptal edildi.', ''); return; }
    console.error(error);
    setDownloadsAutoStatus(error.message || 'İndirilenler klasörü bağlanamadı.', 'error');
  }
}

async function initDownloadsAutomation() {
  if (!supportsDownloadsFolderAccess()) {
    if (downloadsAutoPanel) downloadsAutoPanel.hidden = false;
    setDownloadsAutoStatus('Otomatik klasör okuma için masaüstü Chrome/Edge ve HTTPS/localhost gerekir.', 'error');
    return;
  }
  try {
    downloadsDirectoryHandle = await restoreDownloadsDirectoryHandle();
    updateDownloadsButtons();
    if (!downloadsDirectoryHandle) return;
    const allowed = await directoryHasReadPermission(downloadsDirectoryHandle, false);
    if (!allowed) {
      setDownloadsAutoStatus('İndirilenler klasörü hatırlandı. Erişimi etkinleştirmek için Bağla düğmesine bir kez bas.', '');
      return;
    }
    await scanDownloadsFolder({ force: true, autoApply: true });
    startDownloadsAutoScanTimer();
  } catch (error) {
    console.warn('İndirilenler otomatik başlatılamadı:', error);
    setDownloadsAutoStatus('Klasör bağlantısı yeniden kurulmalı. Bağla düğmesine bas.', 'error');
  }
}

/* ---------- İndirilenler klasörü otomatik yükleme sonu ---------- */

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
      if (currentMode !== MODE_VACANT && currentMode !== MODE_LATECOUT) {
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
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, raw: true });
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
      : currentMode === MODE_LATECOUT
        ? 'Departures ve Arrivals Excel dosyalarını birlikte seç (Ctrl ile iki dosya). ETD\'si late listesinde olan odalar tabloya yazılır.'
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
  vacantRoomsFileName = '';
  currentRoomFilter = new Map();
  currentRoomFileNames = [];
  dndResults = [];
  lateCoutResults = [];
  lateCoutDateText = '';
  dndDateWindowText = '';
  dndFilterStats = { active: false, currentRooms: 0, skippedOldRooms: 0, stoppedBeforeArrival: 0 };
  greenRoomsInput.value = '';
  if (vacantRoomsPdfInput) vacantRoomsPdfInput.value = '';
  updateVacantRoomsStatus();
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
  dndModeBtn?.classList.toggle('active', currentMode === MODE_DND);
  vacantModeBtn?.classList.toggle('active', currentMode === MODE_VACANT);
  lateCoutModeBtn?.classList.toggle('active', currentMode === MODE_LATECOUT);
  departuresModeBtn.setAttribute('aria-pressed', String(currentMode === MODE_DEPARTURES));
  arrivalsModeBtn.setAttribute('aria-pressed', String(currentMode === MODE_ARRIVALS));
  dndModeBtn?.setAttribute('aria-pressed', String(currentMode === MODE_DND));
  vacantModeBtn?.setAttribute('aria-pressed', String(currentMode === MODE_VACANT));
  lateCoutModeBtn?.setAttribute('aria-pressed', String(currentMode === MODE_LATECOUT));
  document.body.classList.toggle('mode-arrivals', currentMode === MODE_ARRIVALS);
  document.body.classList.toggle('mode-departures', currentMode === MODE_DEPARTURES);
  document.body.classList.toggle('mode-dnd', currentMode === MODE_DND);
  document.body.classList.toggle('mode-vacant', currentMode === MODE_VACANT);
  document.body.classList.toggle('mode-latecout', currentMode === MODE_LATECOUT);
  greenPanel.hidden = currentMode !== MODE_ARRIVALS || !originalGroups.size;
  greenRoomsInput.disabled = currentMode !== MODE_ARRIVALS || !originalGroups.size;
  if (vacantRoomsPdfInput) vacantRoomsPdfInput.disabled = currentMode !== MODE_ARRIVALS || !originalGroups.size;
  if (fileInput) {
    fileInput.multiple = currentMode !== MODE_DND;
    fileInput.accept = currentMode === MODE_VACANT ? '.pdf' : '.xlsx,.xls,.csv';
  }
  if (uploadTitle) {
    uploadTitle.textContent = currentMode === MODE_VACANT
      ? 'Vacant PDF Yükle'
      : currentMode === MODE_DND
        ? 'Excel Yükle'
        : currentMode === MODE_LATECOUT
          ? 'Departures + Arrivals Birlikte Yükle'
          : 'Excel Yükle / Birden Fazla Seç';
  }
  updateCurrentRoomsPanel();
  if (currentMode === MODE_DND || currentMode === MODE_LATECOUT) {
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
  if (![MODE_DEPARTURES, MODE_ARRIVALS].includes(mode)) return;
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

  if (downloadsDirectoryHandle) {
    setTimeout(() => applyAutoDownloadsForCurrentMode({ silentMissing: true }), 0);
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

vacantRoomsPdfInput?.addEventListener('change', event => {
  const file = event.target.files?.[0];
  if (file) handleArrivalsVacantPdf(file);
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

  const assignCheck = event.target.closest('.assignment-check input[type="checkbox"]');
  if (assignCheck) {
    const targetGroup = assignCheck.dataset.target;
    const key = assignCheck.value;
    if (assignCheck.checked) {
      sectionAssignments.set(key, targetGroup);
    } else if (sectionAssignments.get(key) === targetGroup) {
      sectionAssignments.delete(key);
    }
    cleanAssignments();
    updateOutput();
  }
});

chiefControls.addEventListener('click', event => {
  const assignAllBtn = event.target.closest('[data-assign-all]');
  if (assignAllBtn) {
    const targetGroup = assignAllBtn.dataset.assignAll;
    getLeaveSections().forEach(section => {
      if (section.sourceGroup !== targetGroup && !sectionAssignments.has(section.key)) {
        sectionAssignments.set(section.key, targetGroup);
      }
    });
    cleanAssignments();
    updateOutput();
    return;
  }

  const assignClearBtn = event.target.closest('[data-assign-clear]');
  if (assignClearBtn) {
    const targetGroup = assignClearBtn.dataset.assignClear;
    [...sectionAssignments.entries()].forEach(([key, target]) => {
      if (target === targetGroup) sectionAssignments.delete(key);
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
    border: border ? excelBorder(hexToRgb(appSettings.borderColor, '555555')) : undefined,
  };
}

function excelHeaders(isArrivals) {
  if (isArrivals) {
    return ['', 'Room', 'ETA', 'Arrival', 'Adult\ns', 'Childr\nen', 'Child\nAges', 'Departure', 'ETD', 'Name', 'Travel\nAgent', 'NOTLAR'];
  }
  return ['', 'Room', 'ETA', 'Arrival', 'Adul\nts', 'Childr\nen', 'Child\nAges', 'Departure', 'ETD', 'Travel Agent', 'NOTLAR'];
}

function recordToExcelRow(record, rowNumber, isArrivals) {
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
    return [4.5, 7, 8, 7, 6, 7.5, 18, 9, 9, 12, 5, 5, 7, 11, 28].map(wch => ({ wch }));
  }
  if (isArrivals) {
    return [4.5, 8, 9, 11, 6, 6.5, 8.5, 11, 8.5, 13, 12, 34].map(wch => ({ wch }));
  }
  return [4.5, 8, 8.5, 11, 6, 6.5, 8.5, 11, 8.5, 16, 34].map(wch => ({ wch }));
}

function styleExcelSheet(ws, records, groupName, isArrivals, startRow, isVacant = false) {
  const headers = excelHeaders(isArrivals, isVacant);
  const columnCount = headers.length;
  const headerFill = isVacant ? '4B5563' : (isArrivals ? hexToRgb(appSettings.arrivalsHeaderColor, '79A9D4') : hexToRgb(appSettings.headerColor, 'C8755C'));
  const headerStyle = excelCellStyle({ fill: headerFill, bold: true, size: 10 });
  const bodyStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 10 });
  const notesStyle = excelCellStyle({ fill: 'FFFFFF', bold: false, size: 10, align: 'left' });
  const smallTextStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 9 });
  const compactArrivalTextStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 8 });
  const largerArrivalStyle = excelCellStyle({ fill: 'FFFFFF', bold: true, size: 12 });
  const etdHighlightStyle = excelCellStyle({ fill: hexToRgb(appSettings.etdColor, 'FFF176'), bold: true, size: 10 });
  const roomGreenStyle = excelCellStyle({ fill: hexToRgb(appSettings.greenColor, 'B8D8BD'), bold: true, size: isArrivals ? 12 : 10 });
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
    ws['!rows'][rowIndex] = { hpt: appSettings.excelRowHeight > 0 ? appSettings.excelRowHeight : (isArrivals ? 32 : (isVacant ? 30 : 26)) };

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
      if (isArrivals && [3, 7, 8, 9, 10].includes(c)) {
        cell.s = compactArrivalTextStyle;
      }
      if (c === columnCount - 1) {
        cell.s = notesStyle;
      }
      if (c === 1 && isGreenRoom(record)) {
        cell.s = roomGreenStyle;
      }
      if (!isArrivals && !isVacant && c === 8 && appSettings.etdHighlightEnabled && record.etd === ETD_HIGHLIGHT) {
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

  if (currentMode === MODE_LATECOUT) {
    downloadLateCoutExcel();
    return;
  }

  const { groups, unassigned } = buildPrintableGroups();
  if (unassigned.length) {
    const labels = unassigned.map(section => `${section.sectionName}`).join(', ');
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
      const labels = unassigned.map(section => `${section.sectionName}`).join(', ');
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

downloadsConnectBtn?.addEventListener('click', connectDownloadsFolder);
downloadsRefreshBtn?.addEventListener('click', () => scanDownloadsFolder({ force: true, autoApply: true }));

greenRoomsInput.addEventListener('input', () => {
  if (originalGroups.size) updateOutput();
});

startDeparturesBtn?.addEventListener('click', () => selectModeFromMenu(MODE_DEPARTURES));
startArrivalsBtn?.addEventListener('click', () => selectModeFromMenu(MODE_ARRIVALS));
startDndBtn?.addEventListener('click', () => selectModeFromMenu(MODE_DND));
backMenuBtn?.addEventListener('click', showStartMenu);
departuresModeBtn.addEventListener('click', () => setMode(MODE_DEPARTURES));
arrivalsModeBtn.addEventListener('click', () => setMode(MODE_ARRIVALS));
dndModeBtn?.addEventListener('click', () => setMode(MODE_DND));
vacantModeBtn?.addEventListener('click', () => setMode(MODE_VACANT));
lateCoutModeBtn?.addEventListener('click', () => setMode(MODE_LATECOUT));
officeBtn.addEventListener('click', printOfficeDirect);
officeExcelBtn.addEventListener('click', exportOfficeExcelDirect);
excelBtn.addEventListener('click', exportCurrentExcel);
printBtn.addEventListener('click', () => printCleanPdf({ printTitle: modeLabel(), skipAssignmentCheck: currentMode === MODE_DND || currentMode === MODE_LATECOUT }));
clearBtn.addEventListener('click', clearAll);

settingsBtn?.addEventListener('click', openSettings);
settingsCloseBtn?.addEventListener('click', closeSettings);
settingsOverlay?.addEventListener('click', event => {
  if (event.target === settingsOverlay) closeSettings();
});
settingsSaveBtn?.addEventListener('click', () => {
  appSettings = collectSettingsForm();
  saveSettingsToStorage();
  applySettings();
  closeSettings();
  rerenderAfterSettings();
  setStatus('Ayarlar kaydedildi ve uygulandı.', 'ok');
});
settingsResetBtn?.addEventListener('click', () => {
  appSettings = { ...DEFAULT_SETTINGS };
  saveSettingsToStorage();
  applySettings();
  fillSettingsForm();
  rerenderAfterSettings();
  setStatus('Ayarlar varsayılanlara döndürüldü.', 'ok');
});
settingsExportBtn?.addEventListener('click', exportSettingsFile);
settingsImportBtn?.addEventListener('click', () => settingsImportInput?.click());
settingsImportInput?.addEventListener('change', event => {
  importSettingsFile(event.target.files?.[0]);
  event.target.value = '';
});

loadSettings();
applySettings();
updateModeUi();
setButtons({ printable: false, clearable: false });
showApp();
initDownloadsAutomation();
