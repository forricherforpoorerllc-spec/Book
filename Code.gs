/* =====================================================================
 *  My Reading Journey — Google Apps Script Backend (Code.gs)
 *  Connects the SPA (index.html) to Google Sheets as a database.
 *
 *  Sheet Tabs:
 *    Library       – Every book in the user's collection
 *    Challenges    – Reading goals / challenges
 *    Shelves       – Custom shelves + book→shelf map
 *    Profile       – Single-row: identity, prefs, settings
 *    Audiobooks    – Saved audiobooks & playback positions
 *
 *  Deployment: Deploy → Web app → Execute as *me*, access "Anyone"
 * ===================================================================== */

// ── Serve the UI ────────────────────────────────────────────────────────
function doGet() {
  var title = _buildJourneyTitle();
  var output = HtmlService.createHtmlOutputFromFile('index')
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  try {
    // XFrameOptionsMode may be unavailable in some script contexts
    var xfMode = HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.SAMEORIGIN;
    if (xfMode != null) output.setXFrameOptionsMode(xfMode);
  } catch(e) {}
  return output;
}

/** Build "{Name}'s Reading Journey" (or "My Reading Journey" if no name) */
function _buildJourneyTitle() {
  var sheet = _ss().getSheetByName(SHEET_PROFILE);
  if (sheet && sheet.getLastRow() >= 2) {
    var name = String(sheet.getRange(2, 1).getValue() || '').trim();
    if (name) {
      var possessive = name.endsWith('s') ? name + "'" : name + "'s";
      return possessive + ' Reading Journey';
    }
  }
  return 'My Reading Journey';
}

// ── Constants ───────────────────────────────────────────────────────────
var SHEET_DASHBOARD   = 'Dashboard';
var SHEET_LIBRARY     = 'Library';
var SHEET_STATS       = 'Stats';
var SHEET_CHALLENGES  = 'Challenges';
var SHEET_SHELVES     = 'Shelves';
var SHEET_PROFILE     = 'Profile';
var SHEET_AUDIOBOOKS  = 'Audiobooks';
var SHEET_COVER       = 'Cover';

// Data model — 'Cover' is a display-only IMAGE column placed RIGHT AFTER BookId
// so it becomes the first visible column in the Library tab.
// Stored value for Cover is '' — the IMAGE() formula is written by _initLibrarySheet.
var LIBRARY_HEADERS = [
  'BookId','Cover','Title','Author','Status','Rating','Pages','Genre',
  'DateAdded','DateStarted','DateFinished','CurrentPage',
  'Series','SeriesNumber','TbrPriority','Format','Source','SpiceLevel',
  'Tags','Shelves','Notes','Review','Quotes',
  'Favorite','CoverEmoji','CoverUrl','Gradient1','Gradient2',
  'ISBN','OLID','AuthorKey'
];

/** Letter for a LIBRARY_HEADERS column name. Safe against reordering. */
function _L(name) {
  var idx = LIBRARY_HEADERS.indexOf(name);
  if (idx < 0) return 'A';
  return _colLetter(idx + 1);
}

/** Full qualified Library reference for one column (e.g. `'Library'!D:D`). */
function _LR(name) {
  var letter = _L(name);
  return "'" + SHEET_LIBRARY + "'!" + letter + ":" + letter;
}

/** Library reference starting at row 2 (e.g. `'Library'!D2:D`). */
function _LR2(name) {
  var letter = _L(name);
  return "'" + SHEET_LIBRARY + "'!" + letter + "2:" + letter;
}

var CHALLENGE_HEADERS = ['ChallengeId','Name','Icon','Current','Target'];

var SHELF_HEADERS = ['ShelfId','Name','Icon'];

var PROFILE_HEADERS = [
  'Name','Motto','PhotoData','Theme',
  'YearlyGoal','Onboarded','DemoCleared','ShowSpoilers',
  'ReadingOrder','RecentIds','SortBy','LibViewMode',
  'SelectedFilter','ActiveShelf','ChallengeBarCollapsed','LibToolsOpen',
  'LibraryName',
  // --- Full-sync fields (mirror the last remaining localStorage-only keys) ---
  'CustomQuotes','CoversEnabled','TutorialCompleted',
  'LastAudioId','TotalListeningMins'
];

var AUDIOBOOK_HEADERS = [
  'AudiobookId','Title','Author','Duration','CoverEmoji','CoverUrl',
  'ChapterCount','LibrivoxProjectId','CurrentChapterIndex',
  'CurrentTime','PlaybackSpeed','TotalListeningMins'
];

// ── Theme palettes (match UI themes exactly) ────────────────────────────
var THEME_PALETTES = {
  romantic: {
    bg: '#FFF1F8', header: '#BE185D', headerText: '#FFFFFF', accent: '#F472B6',
    border: '#F9A8D4', altRow: '#FEF0F5', text: '#4B5563', tabColor: '#D94684',
    titleCol: '#9D174D', subtleText: '#6B7280', favHighlight: '#FFF1F2'
  },
  spicy: {
    bg: '#2A0008', header: '#8F001F', headerText: '#FFE4E8', accent: '#FF4D4D',
    border: '#5A001A', altRow: '#3D0E14', text: '#FFE4E8', tabColor: '#FF1F3D',
    titleCol: '#FF6A6A', subtleText: '#E8C5D8', favHighlight: '#4A0010'
  },
  dreamy: {
    bg: '#F5F3FF', header: '#8B5CF6', headerText: '#FFFFFF', accent: '#A78BFA',
    border: '#DDD6FE', altRow: '#EDE9FE', text: '#4B5563', tabColor: '#A78BFA',
    titleCol: '#7C3AED', subtleText: '#6B7280', favHighlight: '#FDF2F8'
  },
  fresh: {
    bg: '#ECFDF5', header: '#10B981', headerText: '#FFFFFF', accent: '#34D399',
    border: '#A7F3D0', altRow: '#D1FAE5', text: '#374151', tabColor: '#10B981',
    titleCol: '#047857', subtleText: '#6B7280', favHighlight: '#FEF3C7'
  },
  midnight: {
    bg: '#0F172A', header: '#312E81', headerText: '#E0E7FF', accent: '#818CF8',
    border: '#334155', altRow: '#1E293B', text: '#CBD5E1', tabColor: '#6366F1',
    titleCol: '#A5B4FC', subtleText: '#94A3B8', favHighlight: '#1E1B4B'
  },
  sunset: {
    bg: '#130824', header: '#1A0525', headerText: '#FFF0F5', accent: '#FF6D00',
    border: '#6A1F5C', altRow: '#2A0E30', text: '#E8C5D8', tabColor: '#FF6D00',
    titleCol: '#FFB300', subtleText: '#B090A8', favHighlight: '#3D0E32'
  },
  // ── P1 additional ──
  velvet: {
    bg: '#1A1040', header: '#4C1D95', headerText: '#F5F3FF', accent: '#D946EF',
    border: '#3B1F6E', altRow: '#240A5E', text: '#DDD6FE', tabColor: '#9333EA',
    titleCol: '#E879F9', subtleText: '#A78BFA', favHighlight: '#2D1060'
  },
  // ── P2 themes ──
  horizon: {
    bg: '#f0f9ff', header: '#0284c7', headerText: '#FFFFFF', accent: '#22d3ee',
    border: '#bae6fd', altRow: '#e0f2fe', text: '#1e3a5f', tabColor: '#0284c7',
    titleCol: '#0369a1', subtleText: '#5A7FA8', favHighlight: '#fffbf0'
  },
  arctic: {
    bg: '#030d1a', header: '#0c2a4a', headerText: '#e0f2fe', accent: '#38bdf8',
    border: '#1e3a5f', altRow: '#0a1e30', text: '#bae6fd', tabColor: '#0369a1',
    titleCol: '#7dd3fc', subtleText: '#60a5fa', favHighlight: '#051525'
  },
  sahara: {
    bg: '#fffbf0', header: '#d97706', headerText: '#FFFFFF', accent: '#f59e0b',
    border: '#fde68a', altRow: '#fef3c7', text: '#4a3720', tabColor: '#d97706',
    titleCol: '#92400e', subtleText: '#b45309', favHighlight: '#fff7e6'
  },
  ember: {
    bg: '#060e06', header: '#122010', headerText: '#e8dfc8', accent: '#d4a030',
    border: '#1e3420', altRow: '#0a160a', text: '#c4b898', tabColor: '#5a9e48',
    titleCol: '#d4a030', subtleText: '#9a9070', favHighlight: '#040c04'
  },
  volcano: {
    bg: '#FFF8F5', header: '#B91C1C', headerText: '#FFFFFF', accent: '#F59E0B',
    border: '#FCA5A5', altRow: '#FFF0E8', text: '#4A1010', tabColor: '#DC2626',
    titleCol: '#1C0808', subtleText: '#7A3030', favHighlight: '#FFF5F0'
  },
  dusk: {
    bg: '#130824', header: '#1A0525', headerText: '#FFF0F5', accent: '#FFCA0A',
    border: '#6A1F5C', altRow: '#3D0E32', text: '#E8C5D8', tabColor: '#FF6D00',
    titleCol: '#FF6D00', subtleText: '#B090A8', favHighlight: '#0D0418'
  },
  // ── legacy/unused themes ──
  petal: {
    bg: '#FFF8FC', header: '#C06C84', headerText: '#FFFFFF', accent: '#F67280',
    border: '#F9B2D7', altRow: '#FFF0F8', text: '#3a2030', tabColor: '#C06C84',
    titleCol: '#8B3A55', subtleText: '#7a5060', favHighlight: '#F6FFDC'
  },
  coral: {
    bg: '#1a2e3f', header: '#355C7D', headerText: '#fdf2f8', accent: '#F8B195',
    border: '#3d5a7a', altRow: '#243548', text: '#F8B195', tabColor: '#F67280',
    titleCol: '#F8B195', subtleText: '#C06C84', favHighlight: '#152535'
  },
  lagoon: {
    bg: '#F5F5F5', header: '#229799', headerText: '#FFFFFF', accent: '#48CFCB',
    border: '#A2D5C6', altRow: '#E8FAF8', text: '#1f2937', tabColor: '#229799',
    titleCol: '#166d6f', subtleText: '#3A8B95', favHighlight: '#F0FFF4'
  },
  'mint mist': {
    bg: '#F5F5F5', header: '#229799', headerText: '#FFFFFF', accent: '#48CFCB',
    border: '#A2D5C6', altRow: '#E8FAF8', text: '#1f2937', tabColor: '#229799',
    titleCol: '#166d6f', subtleText: '#3A8B95', favHighlight: '#F0FFF4'
  },
  jade: {
    bg: '#0a1510', header: '#237227', headerText: '#f0fdf4', accent: '#CFFFE2',
    border: '#1a3d22', altRow: '#0f1f15', text: '#CFFFE2', tabColor: '#519A66',
    titleCol: '#A2D5C6', subtleText: '#66D0BC', favHighlight: '#071210'
  },
  'sage forest': {
    bg: '#0a1510', header: '#237227', headerText: '#f0fdf4', accent: '#CFFFE2',
    border: '#1a3d22', altRow: '#0f1f15', text: '#CFFFE2', tabColor: '#519A66',
    titleCol: '#A2D5C6', subtleText: '#66D0BC', favHighlight: '#071210'
  },
  // ── P4 themes ──
  champagne: {
    bg: '#faf4e8', header: '#6B3800', headerText: '#FEFBF2', accent: '#C9A84C',
    border: '#DEC87A', altRow: '#f2e8cc', text: '#1A0E00', tabColor: '#C9A84C',
    titleCol: '#8B5200', subtleText: '#7A5A30', favHighlight: '#fffbf2'
  },
  obsidian: {
    bg: '#080807', header: '#1c1c10', headerText: '#F5F5F0', accent: '#D4AF37',
    border: '#3a3520', altRow: '#111108', text: '#F5F5F0', tabColor: '#D4AF37',
    titleCol: '#F0D060', subtleText: '#A89050', favHighlight: '#050504'
  },
  pearl: {
    bg: '#F8F9FC', header: '#2A3A52', headerText: '#FFFFFF', accent: '#5A7FA8',
    border: '#C8D4E4', altRow: '#EEF2F8', text: '#12181F', tabColor: '#5A7FA8',
    titleCol: '#344868', subtleText: '#5A6A80', favHighlight: '#F2F4F8'
  },
  opal: {
    bg: '#F8F9FC', header: '#2A3A52', headerText: '#FFFFFF', accent: '#5A7FA8',
    border: '#C8D4E4', altRow: '#EEF2F8', text: '#12181F', tabColor: '#5A7FA8',
    titleCol: '#344868', subtleText: '#5A6A80', favHighlight: '#F2F4F8'
  },
  onyx: {
    bg: '#0a0a0a', header: '#0a1428', headerText: '#F0F0F0', accent: '#1A60CC',
    border: '#1a1a2e', altRow: '#111111', text: '#F0F0F0', tabColor: '#2870E0',
    titleCol: '#4A8EF0', subtleText: '#7090D8', favHighlight: '#060606'
  },
  bunny: {
    bg: '#0A0A0A', header: '#1A000F', headerText: '#FFFFFF', accent: '#FF007F',
    border: '#FF99CC', altRow: '#100008', text: '#FFFFFF', tabColor: '#FF007F',
    titleCol: '#FF80BF', subtleText: '#FF99C8', favHighlight: '#050504'
  },
  // ── P3 themes (index3) ──
  blossom: {
    bg: '#FFF7FA', header: '#C85888', headerText: '#FFFFFF', accent: '#F8D4A0',
    border: '#F4C8DC', altRow: '#F8FEFF', text: '#503048', tabColor: '#C0407C',
    titleCol: '#1E1018', subtleText: '#8a5878', favHighlight: '#FFFBF4'
  },
  lavenderhaze: {
    bg: '#F8F6FF', header: '#7858B0', headerText: '#FFFFFF', accent: '#F0D898',
    border: '#D8CCF0', altRow: '#F4F8FF', text: '#402860', tabColor: '#6848A8',
    titleCol: '#18101E', subtleText: '#7858A0', favHighlight: '#FFFCF4'
  },
  sorbet: {
    bg: '#FFFAF6', header: '#C87050', headerText: '#FFFFFF', accent: '#F8E890',
    border: '#F8D0B8', altRow: '#F6FFFD', text: '#503828', tabColor: '#C85860',
    titleCol: '#1C1008', subtleText: '#886858', favHighlight: '#FEFFF4'
  },
  cloud: {
    bg: '#F4FBFF', header: '#3E90C4', headerText: '#FFFFFF', accent: '#F4B8D8',
    border: '#B8DCF4', altRow: '#F8F6FF', text: '#2C4860', tabColor: '#2888C8',
    titleCol: '#101828', subtleText: '#507898', favHighlight: '#FFF8FC'
  },
  meadow: {
    bg: '#F6FAF4', header: '#3A4E30', headerText: '#FFFFFF', accent: '#FFD8EC',
    border: '#C8D8C0', altRow: '#ECF4EA', text: '#3A4E38', tabColor: '#94A684',
    titleCol: '#1A2418', subtleText: '#5A7058', favHighlight: '#FFEEF4'
  },
  sherbet: {
    bg: '#F7F9F2', header: '#0D5044', headerText: '#FFFFFF', accent: '#7048A0',
    border: '#B8E4D8', altRow: '#EDF7F5', text: '#1A3828', tabColor: '#1A7A68',
    titleCol: '#0A1814', subtleText: '#2A5848', favHighlight: '#F5EDF7'
  }
};

// ── Human-readable display headers ──────────────────────────────────────
var DISPLAY_MAP = {
  // Library
  'BookId':'ID', 'Cover':'Cover', 'DateAdded':'Date Added', 'DateStarted':'Date Started',
  'DateFinished':'Date Finished', 'CurrentPage':'Current Page',
  'SeriesNumber':'Series #', 'TbrPriority':'TBR Priority',
  'SpiceLevel':'Spice Level', 'CoverEmoji':'Cover', 'CoverUrl':'Cover URL',
  'Gradient1':'Grad 1', 'Gradient2':'Grad 2', 'AuthorKey':'Author Key',
  // Challenges
  'ChallengeId':'ID', 'Current':'Progress', 'Target':'Goal',
  // Shelves
  'ShelfId':'ID',
  // Profile
  'PhotoData':'Photo', 'YearlyGoal':'Yearly Goal', 'DemoCleared':'Demo Cleared',
  'ShowSpoilers':'Show Spoilers', 'ReadingOrder':'Reading Order',
  'RecentIds':'Recent IDs', 'SortBy':'Sort By', 'LibViewMode':'View Mode',
  'SelectedFilter':'Active Filter', 'ActiveShelf':'Active Shelf',
  'ChallengeBarCollapsed':'Goals Collapsed', 'LibToolsOpen':'Tools Open',
  'LibraryName':'Library Name',
  'CustomQuotes':'Custom Quotes', 'CoversEnabled':'Covers Enabled',
  'TutorialCompleted':'Tutorial Completed', 'LastAudioId':'Last Audiobook',
  'TotalListeningMins':'Total Listening (min)',
  // Audiobooks
  'AudiobookId':'ID', 'ChapterCount':'Chapters',
  'LibrivoxProjectId':'Project ID', 'CurrentChapterIndex':'Current Chapter',
  'CurrentTime':'Position', 'PlaybackSpeed':'Speed',
  'TotalListeningMins':'Listened (min)'
};

function _displayHeaders(internalHeaders) {
  return internalHeaders.map(function(h) { return DISPLAY_MAP[h] || h; });
}

// ── Utility ─────────────────────────────────────────────────────────────
function _uuid() {
  return Utilities.getUuid();
}

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ── Enterprise utilities ────────────────────────────────────────────────

/** Structured logger — always use this instead of raw Logger.log() */
function _log(level, fn, msg) {
  Logger.log('[' + level.toUpperCase() + '] ' + fn + ': ' + String(msg));
}

/**
 * Validate that an ID value is a non-empty string of reasonable length.
 * Rejects null, undefined, empty string, and values > 200 chars.
 */
function _validateId(val) {
  var s = String(val || '').trim();
  return s.length >= 4 && s.length <= 200;
}

/**
 * Validate a theme name against the known palette keys.
 * Returns the sanitized theme name, or 'blossom' if unknown.
 */
var _VALID_THEME_KEYS = (function() {
  var keys = {};
  Object.keys(THEME_PALETTES).forEach(function(k) { keys[k] = true; });
  return keys;
}());

function _validateTheme(name) {
  var t = String(name || '').toLowerCase().trim();
  return _VALID_THEME_KEYS[t] ? t : 'blossom';
}

function _getOrCreateSheet(name, headers) {
  var ss = _ss();
  var sheet = ss.getSheetByName(name);
  var displayRow = _displayHeaders(headers);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    _ensureColumns(sheet, displayRow.length);
    sheet.getRange(1, 1, 1, displayRow.length).setValues([displayRow]);
    sheet.setFrozenRows(1);
  } else {
    _ensureColumns(sheet, displayRow.length);
    // ── Schema migration for existing Library sheets ──
    // Older versions shipped without the 'Cover' column at index 2.
    // If we detect that shape (col A = BookId/ID, col B = Title), inject a new col B.
    if (name === SHEET_LIBRARY && headers === LIBRARY_HEADERS) {
      var h1 = String(sheet.getRange(1, 1).getValue() || '').toLowerCase();
      var h2 = String(sheet.getRange(1, 2).getValue() || '').toLowerCase();
      if ((h1 === 'id' || h1 === 'bookid') && (h2 === 'title')) {
        try {
          sheet.insertColumnAfter(1);
          _log('info', '_getOrCreateSheet', 'Migrated Library: inserted Cover column at position 2');
        } catch (migErr) {
          _log('error', '_getOrCreateSheet', 'Library migration failed: ' + migErr);
        }
      }
    }
    var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    var normalizedCurrent = currentHeaders.filter(function(value) { return value !== ''; });
    var needsHeaderSync = normalizedCurrent.length !== displayRow.length || displayRow.some(function(header, index) {
      return normalizedCurrent[index] !== header;
    });
    if (needsHeaderSync) {
      sheet.getRange(1, 1, 1, displayRow.length).setValues([displayRow]);
    }
  }
  return sheet;
}

/** Ensure a sheet has at least `needed` columns. Prevents "out of bounds" errors. */
function _ensureColumns(sheet, needed) {
  var current = sheet.getMaxColumns();
  if (current < needed) {
    sheet.insertColumnsAfter(current, needed - current);
  }
}

/** Ensure a sheet has at least `needed` rows. */
function _ensureRows(sheet, needed) {
  var current = sheet.getMaxRows();
  if (current < needed) {
    sheet.insertRowsAfter(current, needed - current);
  }
}

function _sheetToObjects(sheet, internalHeaders) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = internalHeaders || data[0];
  return data.slice(1).filter(function(row) {
    // Ignore display-only rows that don't carry data for the mapped headers.
    for (var i = 0; i < headers.length; i++) {
      if (row[i] !== '' && row[i] !== null) return true;
    }
    return false;
  }).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function _findRowByCol(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][colIndex]) === String(value)) return r + 1; // 1-indexed
  }
  return -1;
}

/** Convert 1-based column number to letter (1→A, 27→AA) */
function _colLetter(col) {
  var s = '';
  while (col > 0) {
    col--;
    s = String.fromCharCode(65 + (col % 26)) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

function _getThemePalette(themeName) {
  return THEME_PALETTES[String(themeName || '').toLowerCase()] || THEME_PALETTES.blossom;
}

/**
 * Write a live IMAGE() formula to the Cover column (col 2) for a specific row.
 * Called after clientAddBook / clientUpdateBook to restore the formula that
 * appendRow / setValues would otherwise overwrite with an empty string.
 */
function _writeCoverFormula(sheet, rowNum) {
  try {
    var coverCol   = LIBRARY_HEADERS.indexOf('Cover') + 1;
    var urlLetter  = _colLetter(LIBRARY_HEADERS.indexOf('CoverUrl') + 1);
    var isbnLetter = _colLetter(LIBRARY_HEADERS.indexOf('ISBN') + 1);
    sheet.getRange(rowNum, coverCol).setFormula(
      '=IFERROR(IMAGE(IF(' + urlLetter + rowNum + '<>"",' + urlLetter + rowNum +
      ',IF(' + isbnLetter + rowNum + '<>"",' +
      '"https://covers.openlibrary.org/b/isbn/"&' + isbnLetter + rowNum + '&"-L.jpg","")),4,80,60),"")'
    );
  } catch(e) {
    _log('warn', '_writeCoverFormula', 'row ' + rowNum + ': ' + e);
  }
}

/** Blend two hex colors by averaging RGB channels (for hero gradient mid-tone) */
function _blendColors(hex1, hex2) {
  function _parse(h) {
    h = h.replace('#', '');
    return [parseInt(h.substring(0,2),16), parseInt(h.substring(2,4),16), parseInt(h.substring(4,6),16)];
  }
  var c1 = _parse(hex1), c2 = _parse(hex2);
  var r = Math.round((c1[0]+c2[0])/2).toString(16);
  var g = Math.round((c1[1]+c2[1])/2).toString(16);
  var b = Math.round((c1[2]+c2[2])/2).toString(16);
  return '#' + (r.length<2?'0'+r:r) + (g.length<2?'0'+g:g) + (b.length<2?'0'+b:b);
}

function _getCurrentTheme() {
  var sheet = _ss().getSheetByName(SHEET_PROFILE);
  if (!sheet || sheet.getLastRow() < 2) return 'blossom';
  var themeCol = PROFILE_HEADERS.indexOf('Theme') + 1;
  return String(sheet.getRange(2, themeCol).getValue() || 'blossom').toLowerCase();
}

// ── Sheet Initialization & Styling ──────────────────────────────────────
function initializeSheets() {
  // New lightweight sheet experience lives in code1.gs.
  // Keep this delegation so existing menu actions and web app init still work.
  if (typeof _dbLiteInitializeSheets === 'function') {
    _dbLiteInitializeSheets();
    return;
  }

  var theme = _getCurrentTheme();
  // Each init is isolated — one failure must NEVER block other tabs from being built.
  function _safe(name, fn) {
    try { fn(); }
    catch (err) { _log('error', 'initializeSheets', name + ' failed: ' + (err && err.message ? err.message : err)); }
  }
  // Data/utility sheets (hidden)
  _safe('Library',     function(){ _initLibrarySheet(theme); });
  _safe('Challenges',  function(){ _initChallengesSheet(theme); });
  _safe('Shelves',     function(){ _initShelvesSheet(theme); });
  _safe('Profile',     function(){ _initProfileSheet(theme); });
  _safe('Audiobooks',  function(){ _initAudiobooksSheet(theme); });
  // Presentation sheets (visible, screenshot-worthy)
  _safe('Dashboard',   function(){ _initDashboardSheet(theme); });
  _safe('Stats',       function(){ _initStatsSheet(theme); });

  // Seed demo data on first run (only if library is empty)
  _seedDemoData();

  var ss = _ss();

  // Legacy cleanup — remove the old 'Start Here' tab from earlier versions.
  var legacyCover = ss.getSheetByName('Start Here');
  if (legacyCover) {
    try { ss.deleteSheet(legacyCover); } catch(e) {}
  }

  // Tab order: hero presentation first, then library, then analytics; utilities last.
  var order = [SHEET_DASHBOARD, SHEET_LIBRARY, SHEET_STATS,
               SHEET_CHALLENGES, SHEET_SHELVES, SHEET_PROFILE, SHEET_AUDIOBOOKS];
  order.forEach(function(name, i) {
    var s = ss.getSheetByName(name);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(i + 1);
    }
  });

  // Hide utility/data-only tabs. Presentation tabs stay visible.
  [SHEET_CHALLENGES, SHEET_SHELVES, SHEET_PROFILE, SHEET_AUDIOBOOKS].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) {
      try { s.hideSheet(); } catch(e) {}
    }
  });
  // Ensure presentation tabs are visible (in case they were hidden from a prior version).
  [SHEET_DASHBOARD, SHEET_LIBRARY, SHEET_STATS].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) {
      try { s.showSheet(); } catch(e) {}
    }
  });

  // Land users on the Dashboard — the signature presentation tab.
  var dash = ss.getSheetByName(SHEET_DASHBOARD);
  if (dash) ss.setActiveSheet(dash);
}

function _styleHeaderRow(sheet, numCols, palette) {
  var headerRange = sheet.getRange(1, 1, 1, numCols);
  headerRange
    .setBackground(palette.header)
    .setFontColor(palette.headerText)
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontFamily('Google Sans')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(false, false, true, false, false, false, palette.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(1, 48);
}

function _applySheetBase(sheet, numCols, palette) {
  // Tab color
  sheet.setTabColor(palette.tabColor || palette.header);
  var maxRows = Math.max(sheet.getMaxRows(), 200);

  // Hide default gridlines — custom borders carry the structure
  sheet.setHiddenGridlines(true);

  // Data area: consistent font, color, vertical alignment
  var dataRange = sheet.getRange(2, 1, maxRows - 1, numCols);
  dataRange
    .setBackground(palette.bg)
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontColor(palette.text || '#374151')
    .setVerticalAlignment('middle');

  // Modern row heights — denser than cards, roomier than raw database defaults.
  sheet.setRowHeightsForced(2, maxRows - 1, 36);

  // Alternating row banding
  var banding = sheet.getBandings();
  banding.forEach(function(b) { b.remove(); });
  if (maxRows > 1) {
    sheet.getRange(1, 1, maxRows, numCols).applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY
    );
    var newBanding = sheet.getBandings()[0];
    if (newBanding) {
      newBanding
        .setHeaderRowColor(palette.header)
        .setFirstRowColor(palette.bg)
        .setSecondRowColor(palette.altRow);
    }
  }

  // Gridline borders with theme color
  sheet.getRange(1, 1, maxRows, numCols)
    .setBorder(null, null, null, null, true, true, palette.border, SpreadsheetApp.BorderStyle.SOLID);

  // Outer frame for stronger visual hierarchy.
  sheet.getRange(1, 1, maxRows, numCols)
    .setBorder(true, true, true, true, false, false, palette.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function _initLibrarySheet(theme) {
  var sheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
  var palette = _getThemePalette(theme);
  var numCols = LIBRARY_HEADERS.length;   // 31 data columns
  var PROGRESS_COL = numCols + 1;         // col 32 — helper column, never touched by webapp
  _ensureColumns(sheet, PROGRESS_COL + 2);
  _ensureRows(sheet, 700);
  var lastRow = Math.max(sheet.getMaxRows(), 700);

  // ─── Column widths ─────────────────────────────────────────────────────────
  // Width 0 = hidden. Only columns a reader actively uses get real widths.
  var widths = {
    'BookId':0,        'Cover':68,       'Title':240,      'Author':158,
    'Status':118,      'Rating':88,      'Pages':70,       'Genre':120,
    'DateAdded':0,     'DateStarted':98, 'DateFinished':98,'CurrentPage':0,
    'Series':140,      'SeriesNumber':50,'TbrPriority':0,  'Format':96,
    'Source':0,        'SpiceLevel':62,  'Tags':0,         'Shelves':0,
    'Notes':0,         'Review':0,       'Quotes':0,       'Favorite':52,
    'CoverEmoji':0,    'CoverUrl':0,     'Gradient1':0,    'Gradient2':0,
    'ISBN':0,          'OLID':0,         'AuthorKey':0
  };
  LIBRARY_HEADERS.forEach(function(h, i) {
    sheet.setColumnWidth(i + 1, widths[h] || 30);
  });
  sheet.setColumnWidth(PROGRESS_COL, 162);  // Progress helper column

  // ─── Show all, then hide columns where width = 0 ───────────────────────────
  try { sheet.showColumns(1, numCols); } catch(e) {}
  LIBRARY_HEADERS.forEach(function(h, i) {
    if ((widths[h] || 0) === 0) {
      try { sheet.hideColumns(i + 1); } catch(e) {}
    }
  });

  // ─── Row heights — tall enough for cover thumbnails ───────────────────────
  sheet.setRowHeightsForced(2, lastRow - 1, 88);

  _styleHeaderRow(sheet, numCols, palette);
  sheet.setRowHeight(1, 52);  // slightly taller header for breathing room
  _applySheetBase(sheet, numCols, palette);

  // Progress helper column header
  sheet.getRange(1, PROGRESS_COL)
    .setValue('Progress')
    .setBackground(palette.header)
    .setFontColor(palette.headerText)
    .setFontWeight('bold').setFontSize(12)
    .setFontFamily('Google Sans')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, palette.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setNote('Auto-computed: ████░░░░ 42% for active reads · ✓ Done · → Queued · • DNF');

  // ─── Freeze: header row + Cover + Title stay locked while scrolling ────────
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);  // BookId(hidden) col1 + Cover col2 + Title col3

  // ─── Column-specific alignment ───
  var leftCols = ['Title','Author','Series','Tags','Shelves','Notes','Review','Quotes','Source'];
  var centerCols = ['Status','Rating','Pages','Genre','DateAdded','DateStarted','DateFinished',
                    'CurrentPage','SeriesNumber','TbrPriority','Format','SpiceLevel','Favorite',
                    'CoverEmoji','ISBN'];
  leftCols.forEach(function(h) {
    var col = LIBRARY_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment('left');
  });
  centerCols.forEach(function(h) {
    var col = LIBRARY_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment('center');
  });

  // ─── Title column: bold, theme-tinted, wrappable ──────────────────────────
  var titleCol = LIBRARY_HEADERS.indexOf('Title') + 1;
  sheet.getRange(2, titleCol, lastRow - 1, 1)
    .setFontWeight('bold')
    .setFontColor(palette.titleCol || palette.header)
    .setWrap(true);

  // ─── Text wrapping for long text columns ─────────────────────────────────
  ['Notes','Review','Quotes','Tags','Shelves'].forEach(function(h) {
    var col = LIBRARY_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(2, col, lastRow - 1, 1).setWrap(true);
  });

  // ─── Data validation: Status dropdown ───
  var statusCol = LIBRARY_HEADERS.indexOf('Status') + 1;
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Reading', 'Finished', 'Want to Read', 'DNF'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusCol, lastRow - 1, 1).setDataValidation(statusRule);

  // ─── Data validation: Rating 0-5 ───
  var ratingCol = LIBRARY_HEADERS.indexOf('Rating') + 1;
  var ratingRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 5)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, ratingCol, lastRow - 1, 1).setDataValidation(ratingRule);

  // ─── Data validation: SpiceLevel 0-5 ───
  var spiceCol = LIBRARY_HEADERS.indexOf('SpiceLevel') + 1;
  var spiceRule = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 5)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, spiceCol, lastRow - 1, 1).setDataValidation(spiceRule);

  // ─── Data validation: Format dropdown ───
  var formatCol = LIBRARY_HEADERS.indexOf('Format') + 1;
  var formatRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Paperback', 'Hardcover', 'Ebook', 'Audio'], true)
    .setAllowInvalid(true) // allow custom entries
    .build();
  sheet.getRange(2, formatCol, lastRow - 1, 1).setDataValidation(formatRule);

  // Genre curation keeps analytics and slicers clean.
  var genreCol = LIBRARY_HEADERS.indexOf('Genre') + 1;
  var genreRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Romance','Fantasy','Mystery','Thriller','SciFi','Historical','Memoir','Self-Help','Fiction','Other'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, genreCol, lastRow - 1, 1).setDataValidation(genreRule);

  // ─── Favorite checkboxes ───
  var favCol = LIBRARY_HEADERS.indexOf('Favorite') + 1;
  sheet.getRange(2, favCol, lastRow - 1, 1).insertCheckboxes();

  // ─── Date formats ─────────────────────────────────────────────────────────
  // DateAdded stays ISO (hidden internal sync field)
  sheet.getRange(2, LIBRARY_HEADERS.indexOf('DateAdded') + 1, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  // DateStarted and DateFinished are visible — show as "Apr 5, 2026"
  ['DateStarted','DateFinished'].forEach(function(h) {
    var col = LIBRARY_HEADERS.indexOf(h) + 1;
    sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat('mmm d, yyyy');
  });

  // ─── Conditional formatting ────────────────────────────────────────────────
  var rules = sheet.getConditionalFormatRules();
  // Clear managed columns: Status, Rating, Favorite, Spice, Genre, Progress
  var managedCols = [statusCol, ratingCol, favCol, spiceCol, genreCol, PROGRESS_COL];
  rules = rules.filter(function(r) {
    var ranges = r.getRanges();
    return !ranges.some(function(rng) {
      return managedCols.indexOf(rng.getColumn()) !== -1;
    });
  });

  // Status: solid vivid pills — white text on a strong background color
  var statusRange = sheet.getRange(2, statusCol, lastRow - 1, 1);
  [
    { text: 'Reading',      bg: '#1E40AF', font: '#FFFFFF' },
    { text: 'Finished',     bg: '#166534', font: '#FFFFFF' },
    { text: 'Want to Read', bg: '#B45309', font: '#FFFFFF' },
    { text: 'DNF',          bg: '#7F1D1D', font: '#FFFFFF' }
  ].forEach(function(sc) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(sc.text)
      .setBackground(sc.bg).setFontColor(sc.font).setBold(true)
      .setRanges([statusRange]).build());
  });

  // Rating: warm gold gradient
  var ratingRange = sheet.getRange(2, ratingCol, lastRow - 1, 1);
  var ratingColors = [
    { n: 5, font: '#92400E', bg: '#FEF3C7' },
    { n: 4, font: '#D97706', bg: '#FFFBEB' },
    { n: 3, font: '#F59E0B', bg: null },
    { n: 2, font: '#9CA3AF', bg: null },
    { n: 1, font: '#9CA3AF', bg: null }
  ];
  ratingColors.forEach(function(rc) {
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(rc.n)
      .setFontColor(rc.font)
      .setBold(true)
      .setRanges([ratingRange]);
    if (rc.bg) rule.setBackground(rc.bg);
    rules.push(rule.build());
  });

  // Favorite: highlight full row in theme's accent color
  var fullRowRange = sheet.getRange(2, 1, lastRow - 1, numCols);
  var favFormula = '=$' + _colLetter(favCol) + '2=TRUE';
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(favFormula)
    .setBackground(palette.favHighlight || '#FFF1F2')
    .setRanges([fullRowRange])
    .build());

  // SpiceLevel: red tint for heat level 3+
  var spiceRange = sheet.getRange(2, spiceCol, lastRow - 1, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(3)
    .setBackground('#FEE2E2')
    .setFontColor('#DC2626')
    .setBold(true)
    .setRanges([spiceRange])
    .build());

  // Genre: semantic color coding
  var genreRange = sheet.getRange(2, genreCol, lastRow - 1, 1);
  var genreColors = [
    { text: 'Romance',    bg: '#FCE7F3', font: '#BE185D' },
    { text: 'Fantasy',    bg: '#EDE9FE', font: '#6D28D9' },
    { text: 'Mystery',    bg: '#E0E7FF', font: '#3730A3' },
    { text: 'Thriller',   bg: '#FEE2E2', font: '#991B1B' },
    { text: 'SciFi',      bg: '#CFFAFE', font: '#0E7490' },
    { text: 'Historical', bg: '#FEF3C7', font: '#92400E' },
    { text: 'Memoir',     bg: '#DBEAFE', font: '#1E40AF' },
    { text: 'Self-Help',  bg: '#D1FAE5', font: '#065F46' },
    { text: 'Fiction',    bg: '#F3F4F6', font: '#4B5563' }
  ];
  genreColors.forEach(function(gc) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(gc.text)
      .setBackground(gc.bg)
      .setFontColor(gc.font)
      .setBold(true)
      .setRanges([genreRange])
      .build());
  });

  // Progress column: state-based coloring
  var progressRange = sheet.getRange(2, PROGRESS_COL, lastRow - 1, 1);
  [
    { contains: 'Done',   bg: '#D1FAE5', font: '#065F46' },
    { contains: 'Queued', bg: null,      font: '#B45309' },
    { contains: 'DNF',    bg: '#FEE2E2', font: '#7F1D1D' },
    { contains: '%',      bg: null,      font: '#1E40AF' }
  ].forEach(function(pc) {
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(pc.contains).setFontColor(pc.font).setBold(true)
      .setRanges([progressRange]);
    if (pc.bg) rule.setBackground(pc.bg);
    rules.push(rule.build());
  });

  sheet.setConditionalFormatRules(rules);

  // ─── Rating rendered as stars while storing numeric values ───────────────
  sheet.getRange(2, ratingCol, lastRow - 1, 1)
    .setNumberFormat('[=0]"☆☆☆☆☆";[=1]"★☆☆☆☆";[=2]"★★☆☆☆";[=3]"★★★☆☆";[=4]"★★★★☆";"★★★★★"');

  // Pages with thousands separator
  sheet.getRange(2, LIBRARY_HEADERS.indexOf('Pages') + 1, lastRow - 1, 1).setNumberFormat('#,##0');

  // ─── Header notes (hover descriptions) ────────────────────────────────────
  var headerNotes = {
    'Cover':       'Book cover art — auto-loaded from CoverUrl or ISBN via OpenLibrary',
    'Title':       'Book title',
    'Author':      'Primary author name',
    'Status':      'Reading / Finished / Want to Read / DNF',
    'Rating':      'Your rating 0–5, displayed as stars',
    'Pages':       'Total page count',
    'Genre':       'Primary genre — drives the Stats tab charts',
    'DateStarted': 'Date you started reading (Apr 5, 2026)',
    'DateFinished':'Date you finished reading (Apr 5, 2026)',
    'Series':      'Series name (if part of a series)',
    'SeriesNumber':'Book # in series',
    'Format':      'Paperback, Hardcover, Ebook, or Audio',
    'SpiceLevel':  'Romance heat level 0–5',
    'Favorite':    'Check = all-time favorite'
  };
  Object.keys(headerNotes).forEach(function(h) {
    var col = LIBRARY_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(1, col).setNote(headerNotes[h]);
  });

  // ─── Protect header row ────────────────────────────────────────────────────
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  existing.forEach(function(p) { if (p.getDescription() === 'Header row — do not edit') p.remove(); });
  sheet.getRange(1, 1, 1, PROGRESS_COL).protect().setDescription('Header row — do not edit').setWarningOnly(true);

  // ─── Cover IMAGE formulas — batch-written to col 2 (Cover in LIBRARY_HEADERS) ──
  // IMAGE() renders live book art at 80×60 px portrait portrait.
  // Falls back: CoverUrl → ISBN via OpenLibrary → blank.
  // clientAddBook / clientUpdateBook call _writeCoverFormula() after each
  // write to restore this formula that appendRow/setValues would overwrite.
  var coverCol    = LIBRARY_HEADERS.indexOf('Cover') + 1;
  var urlLet      = _colLetter(LIBRARY_HEADERS.indexOf('CoverUrl') + 1);
  var isbnLet     = _colLetter(LIBRARY_HEADERS.indexOf('ISBN') + 1);
  var imgFormulas = [];
  for (var imgR = 2; imgR <= lastRow; imgR++) {
    imgFormulas.push([
      '=IFERROR(IMAGE(IF(' + urlLet + imgR + '<>"",' + urlLet + imgR +
      ',IF(' + isbnLet + imgR + '<>"",' +
      '"https://covers.openlibrary.org/b/isbn/"&' + isbnLet + imgR + '&"-L.jpg","")),4,80,60),"")'
    ]);
  }
  sheet.getRange(2, coverCol, imgFormulas.length, 1)
    .setFormulas(imgFormulas)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // Silence old col 31 content from prior version (AuthorKey was incorrectly
  // repurposed as a second image column — clear and hide it cleanly).
  try {
    sheet.setColumnWidth(numCols, 0);
    sheet.hideColumns(numCols);
  } catch(e) {}

  // ─── Progress helper column (col 32) — Unicode block bars ─────────────────
  // Reading  → "████░░░░  42%"   |   Finished → "✓  Done"
  // Want to Read → "→  Queued"   |   DNF → "•  DNF"
  var sLet  = _colLetter(statusCol);
  var cpLet = _colLetter(LIBRARY_HEADERS.indexOf('CurrentPage') + 1);
  var pgLet = _colLetter(LIBRARY_HEADERS.indexOf('Pages') + 1);
  var progressFormulas = [];
  for (var pr = 2; pr <= lastRow; pr++) {
    progressFormulas.push([
      '=IFERROR(IF(' + sLet + pr + '="Reading",' +
        'REPT(CHAR(9608),ROUND(' + cpLet + pr + '/MAX(1,' + pgLet + pr + ')*8,0))&' +
        'REPT(CHAR(9617),8-ROUND(' + cpLet + pr + '/MAX(1,' + pgLet + pr + ')*8,0))&' +
        '"  "&ROUND(100*' + cpLet + pr + '/MAX(1,' + pgLet + pr + '),0)&"%",' +
      'IF(' + sLet + pr + '="Finished","✓  Done",' +
      'IF(' + sLet + pr + '="Want to Read","→  Queued",' +
      'IF(' + sLet + pr + '="DNF","•  DNF","")))),"")'
    ]);
  }
  sheet.getRange(2, PROGRESS_COL, progressFormulas.length, 1)
    .setFormulas(progressFormulas)
    .setFontFamily('Google Sans').setFontSize(10)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // Alternate banding on progress column
  var progBg = [];
  for (var br = 2; br <= lastRow; br++) {
    progBg.push([(br % 2 === 0) ? palette.altRow : palette.bg]);
  }
  sheet.getRange(2, PROGRESS_COL, progBg.length, 1).setBackgrounds(progBg);
}


function _initChallengesSheet(theme) {
  var sheet = _getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
  var palette = _getThemePalette(theme);
  var numCols = CHALLENGE_HEADERS.length;
  _ensureColumns(sheet, 8);
  _ensureRows(sheet, 50);
  var lastRow = Math.max(sheet.getMaxRows(), 50);

  // Core columns
  sheet.setColumnWidth(1, 30);   // ChallengeId — hidden UUID
  sheet.setColumnWidth(2, 280);  // Name — wide readable label
  sheet.setColumnWidth(3, 70);   // Icon token
  sheet.setColumnWidth(4, 90);   // Current progress value
  sheet.setColumnWidth(5, 90);   // Target value
  // Helper display-only columns (not part of data model)
  sheet.setColumnWidth(6, 200);  // Progress bar sparkline
  sheet.setColumnWidth(7, 80);   // % Complete
  sheet.setColumnWidth(8, 90);   // Remaining
  try { sheet.showColumns(1, 8); } catch(e){}
  sheet.hideColumns(1);

  _styleHeaderRow(sheet, numCols, palette);
  _applySheetBase(sheet, numCols, palette);

  sheet.setFrozenRows(1);

  // ─── Header labels for helper columns ───
  var helperHeaders = ['Progress', '% Done', 'Left'];
  helperHeaders.forEach(function(label, i) {
    var col = 6 + i;
    var cell = sheet.getRange(1, col);
    cell.setValue(label)
      .setBackground(palette.header)
      .setFontColor(palette.headerText)
      .setFontWeight('bold')
      .setFontSize(12)
      .setFontFamily('Google Sans')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  });
  sheet.setRowHeight(1, 48);
  // Match footer border on helper header cells
  sheet.getRange(1, 6, 1, 3)
    .setBorder(false, false, true, false, false, false, palette.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // ─── Alignment for data columns ───
  sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('left').setFontWeight('bold');
  sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment('center').setFontSize(10).setFontWeight('bold');
  sheet.getRange(2, 4, lastRow - 1, 2).setHorizontalAlignment('center');

  // ─── Number validation for Current and Target ───
  [4, 5].forEach(function(col) {
    sheet.getRange(2, col, lastRow - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0)
        .setAllowInvalid(false)
        .build()
    );
  });

  // ─── Helper column formulas (rows 2 onward) ───
  for (var r = 2; r <= lastRow; r++) {
    // Progress bar sparkline (col 6)
    sheet.getRange(r, 6)
      .setFormula('=IFERROR(SPARKLINE(D' + r + '/MAX(1,E' + r + '),{"charttype","bar";"max",1;"color1","' + palette.accent + '";"color2","' + palette.border + '"}),"")');
    // % Done (col 7)
    sheet.getRange(r, 7)
      .setFormula('=IFERROR(IF(E' + r + '>0,ROUND(100*D' + r + '/E' + r + ',0)&"%",""),"")');
    // Remaining (col 8)
    sheet.getRange(r, 8)
      .setFormula('=IFERROR(IF(E' + r + '>0,MAX(0,E' + r + '-D' + r + ')&" left",""),"")');
  }
  sheet.getRange(2, 6, lastRow - 1, 1).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange(2, 7, lastRow - 1, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontFamily('Google Sans').setFontSize(10);
  sheet.getRange(2, 8, lastRow - 1, 1).setHorizontalAlignment('center').setFontColor(palette.subtleText || '#6B7280').setFontFamily('Google Sans').setFontSize(10);

  // Style helper cols background — batched write
  var chalHelperBg = [];
  for (var br = 2; br <= lastRow; br++) {
    var rowBand = (br % 2 === 0) ? palette.altRow : palette.bg;
    chalHelperBg.push([rowBand, rowBand, rowBand]);
  }
  sheet.getRange(2, 6, chalHelperBg.length, 3).setBackgrounds(chalHelperBg);

  // ─── Conditional formatting: highlight completed ───
  var rules = sheet.getConditionalFormatRules().filter(function(r) {
    return !r.getRanges().some(function(rng) { return rng.getColumn() <= 8; });
  });
  var fullRange = sheet.getRange(2, 1, lastRow - 1, 8);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($D2>0,$D2>=$E2)')
    .setBackground('#DCFCE7')
    .setFontColor('#166534')
    .setBold(true)
    .setRanges([fullRange])
    .build());
  sheet.setConditionalFormatRules(rules);

  // Header notes
  sheet.getRange(1, 2).setNote('Challenge or goal name');
  sheet.getRange(1, 3).setNote('Short display token (e.g. Books, Daily, Authors)');
  sheet.getRange(1, 4).setNote('Current progress value — update here or via the web app');
  sheet.getRange(1, 5).setNote('Target value to complete the challenge');
  sheet.getRange(1, 6).setNote('Visual progress bar (auto-calculated)');
  sheet.getRange(1, 7).setNote('Percentage complete (auto-calculated)');
  sheet.getRange(1, 8).setNote('How many remain until goal (auto-calculated)');

  // Protect header
  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  existing.forEach(function(p) { if (p.getDescription() === 'Header row') p.remove(); });
  sheet.getRange(1, 1, 1, 8).protect().setDescription('Header row').setWarningOnly(true);
}

function _initShelvesSheet(theme) {
  var sheet = _getOrCreateSheet(SHEET_SHELVES, SHELF_HEADERS);
  var palette = _getThemePalette(theme);
  var numCols = SHELF_HEADERS.length;
  _ensureColumns(sheet, 4);
  _ensureRows(sheet, 50);
  var lastRow = Math.max(sheet.getMaxRows(), 50);

  sheet.setColumnWidth(1, 30);   // ShelfId — hidden UUID
  sheet.setColumnWidth(2, 300);  // Name — wide
  sheet.setColumnWidth(3, 80);   // Icon token
  // Helper display-only column
  sheet.setColumnWidth(4, 120);  // Books count
  try { sheet.showColumns(1, 4); } catch(e){}
  sheet.hideColumns(1);

  _styleHeaderRow(sheet, numCols, palette);
  _applySheetBase(sheet, numCols, palette);

  sheet.setFrozenRows(1);

  // Header label for helper col
  sheet.getRange(1, 4)
    .setValue('Books on Shelf')
    .setBackground(palette.header)
    .setFontColor(palette.headerText)
    .setFontWeight('bold')
    .setFontSize(12)
    .setFontFamily('Google Sans')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNote('Auto-counted from Library (books whose Shelves field contains this shelf name)');
  sheet.setRowHeight(1, 48);
  sheet.getRange(1, 4)
    .setBorder(false, false, true, false, false, false, palette.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange(2, 2, lastRow - 1, 1).setHorizontalAlignment('left').setFontWeight('bold');
  sheet.getRange(2, 3, lastRow - 1, 1).setHorizontalAlignment('center').setFontSize(10).setFontWeight('bold');

  // ─── Book count formulas ───
  for (var r = 2; r <= lastRow; r++) {
    sheet.getRange(r, 4)
      .setFormula('=IFERROR(IF(B' + r + '<>"",COUNTIF(\'' + SHEET_LIBRARY + '\'!S:S,"*"&B' + r + '&"*"),0),"")');
  }
  sheet.getRange(2, 4, lastRow - 1, 1)
    .setHorizontalAlignment('center')
    .setFontWeight('bold')
    .setFontFamily('Google Sans')
    .setFontSize(10)
    .setFontColor(palette.titleCol || palette.header);
  // Style background to match banding — batched write
  var shelfHelperBg = [];
  for (var br = 2; br <= lastRow; br++) {
    shelfHelperBg.push([(br % 2 === 0) ? palette.altRow : palette.bg]);
  }
  sheet.getRange(2, 4, shelfHelperBg.length, 1).setBackgrounds(shelfHelperBg);

  sheet.getRange(1, 2).setNote('Shelf name (e.g., "Beach Reads", "Book Club")');
  sheet.getRange(1, 3).setNote('Short display token for this shelf');

  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  existing.forEach(function(p) { if (p.getDescription() === 'Header row') p.remove(); });
  sheet.getRange(1, 1, 1, 4).protect().setDescription('Header row').setWarningOnly(true);
}

function _initProfileSheet(theme) {
  var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
  var palette = _getThemePalette(theme);
  var numCols = PROFILE_HEADERS.length;
  _ensureColumns(sheet, numCols);

  // ─── Column widths ───
  var profileWidths = {
    'Name':200, 'Motto':260, 'PhotoData':30, 'Theme':130,
    'YearlyGoal':100, 'Onboarded':30, 'DemoCleared':30, 'ShowSpoilers':110,
    'ReadingOrder':30, 'RecentIds':30, 'SortBy':30, 'LibViewMode':30,
    'SelectedFilter':30, 'ActiveShelf':30, 'ChallengeBarCollapsed':30,
    'LibToolsOpen':30, 'LibraryName':200
  };
  PROFILE_HEADERS.forEach(function(h, i) {
    sheet.setColumnWidth(i + 1, profileWidths[h] || 30);
  });

  // ─── Show all first, then hide internal-only columns ───
  // Visible: Name, Motto, Theme, YearlyGoal, ShowSpoilers, LibraryName
  // Hidden: everything else (internal app state)
  try { sheet.showColumns(1, numCols); } catch(e) {}
  var hideProfileCols = [
    'PhotoData', 'Onboarded', 'DemoCleared',
    'ReadingOrder', 'RecentIds', 'SortBy', 'LibViewMode',
    'SelectedFilter', 'ActiveShelf', 'ChallengeBarCollapsed', 'LibToolsOpen'
  ];
  hideProfileCols.forEach(function(h) {
    var col = PROFILE_HEADERS.indexOf(h) + 1;
    if (col > 0) { try { sheet.hideColumns(col); } catch(e) {} }
  });

  _styleHeaderRow(sheet, numCols, palette);
  _applySheetBase(sheet, numCols, palette);

  sheet.setFrozenRows(1);

  // ─── Alignment ───
  sheet.getRange(2, 1, 1, numCols).setHorizontalAlignment('center');
  ['Name', 'Motto', 'LibraryName'].forEach(function(h) {
    var col = PROFILE_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(2, col).setHorizontalAlignment('left');
  });

  // ─── ShowSpoilers checkbox ───
  var spoilersCol = PROFILE_HEADERS.indexOf('ShowSpoilers') + 1;
  if (spoilersCol > 0) sheet.getRange(2, spoilersCol).insertCheckboxes();

  // ─── Theme dropdown (only user-facing themes) ───
  var themeCol = PROFILE_HEADERS.indexOf('Theme') + 1;
  if (themeCol > 0) {
    sheet.getRange(2, themeCol).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList([
          'romantic','spicy','dreamy','fresh','midnight','sunset','velvet',
          'horizon','arctic','sahara','ember',
          'lagoon','jade','petal','coral',
          'mint mist','sage forest',
          'champagne','obsidian','pearl','opal','onyx','bunny',
          'blossom','lavenderhaze','sorbet','cloud','meadow','sherbet','volcano','dusk'
        ], true)
        .setAllowInvalid(false)
        .build()
    );
  }

  // ─── Header notes (only for visible columns) ───
  var profileNotes = {
    'Name':        'Your display name in the web app',
    'Motto':       'Your personal reading motto shown on your profile',
    'Theme':       'App color theme — change here to re-style all sheets',
    'YearlyGoal':  'Number of books you aim to read this year',
    'ShowSpoilers':'Show spoiler content in reviews',
    'LibraryName': 'Custom name shown as your library title'
  };
  Object.keys(profileNotes).forEach(function(h) {
    var col = PROFILE_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(1, col).setNote(profileNotes[h]);
  });

  // ─── Seed default row if empty ───
  if (sheet.getLastRow() < 2) {
    var defaults = PROFILE_HEADERS.map(function(h) {
      switch(h) {
        case 'Theme': return theme || 'blossom';
        case 'YearlyGoal': return 50;
        case 'Onboarded': return false;
        case 'DemoCleared': return false;
        case 'ShowSpoilers': return false;
        case 'Motto': return 'A focused place to track every book';
        case 'SortBy': return 'default';
        case 'LibViewMode': return 'grid';
        case 'ReadingOrder': return '[]';
        case 'RecentIds': return '[]';
        case 'SelectedFilter': return 'all';
        case 'ActiveShelf': return '';
        case 'ChallengeBarCollapsed': return false;
        case 'LibToolsOpen': return false;
        default: return '';
      }
    });
    sheet.appendRow(defaults);
  }

  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  existing.forEach(function(p) { if (p.getDescription() === 'Header row') p.remove(); });
  sheet.getRange(1, 1, 1, numCols).protect().setDescription('Header row').setWarningOnly(true);
}

function _initAudiobooksSheet(theme) {
  var sheet = _getOrCreateSheet(SHEET_AUDIOBOOKS, AUDIOBOOK_HEADERS);
  var palette = _getThemePalette(theme);
  var numCols = AUDIOBOOK_HEADERS.length;
  _ensureColumns(sheet, numCols + 2);
  _ensureRows(sheet, 50);
  var lastRow = Math.max(sheet.getMaxRows(), 50);

  var audioWidths = {
    'AudiobookId':30,  // hidden UUID
    'Title':230,       'Author':160,
    'Duration':100,    'CoverEmoji':30,  // CoverEmoji hidden — fallback token
    'CoverUrl':30,     'ChapterCount':86,
    'LibrivoxProjectId':30,              // hidden — internal ID
    'CurrentChapterIndex':110,  'CurrentTime':30,  // CurrentTime hidden (seconds, not useful in sheet)
    'PlaybackSpeed':90, 'TotalListeningMins':130
  };
  AUDIOBOOK_HEADERS.forEach(function(h, i) {
    sheet.setColumnWidth(i + 1, audioWidths[h] || 30);
  });

  // ─── Hide technical columns ───
  try { sheet.showColumns(1, numCols); } catch(e) {}
  ['AudiobookId', 'CoverEmoji', 'CoverUrl', 'LibrivoxProjectId', 'CurrentTime'].forEach(function(h) {
    var col = AUDIOBOOK_HEADERS.indexOf(h) + 1;
    if (col > 0) { try { sheet.hideColumns(col); } catch(e) {} }
  });

  // ─── Listening progress helper column ───
  sheet.setColumnWidth(numCols + 1, 180);  // Progress bar
  sheet.setColumnWidth(numCols + 2, 90);   // Chapter %

  _styleHeaderRow(sheet, numCols, palette);
  _applySheetBase(sheet, numCols, palette);

  // Helper header labels
  var helperLabels = ['Listening Progress', 'Chapter %'];
  helperLabels.forEach(function(label, i) {
    var col = numCols + 1 + i;
    sheet.getRange(1, col)
      .setValue(label)
      .setBackground(palette.header)
      .setFontColor(palette.headerText)
      .setFontWeight('bold').setFontSize(12)
      .setFontFamily('Google Sans')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setBorder(false, false, true, false, false, false, palette.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });
  sheet.setRowHeight(1, 48);

  sheet.setFrozenRows(1);

  // ─── Alignment ───
  var titleCol = AUDIOBOOK_HEADERS.indexOf('Title') + 1;
  var authorCol = AUDIOBOOK_HEADERS.indexOf('Author') + 1;
  sheet.getRange(2, titleCol, lastRow - 1, 1).setHorizontalAlignment('left').setFontWeight('bold').setFontColor(palette.titleCol || palette.header);
  sheet.getRange(2, authorCol, lastRow - 1, 1).setHorizontalAlignment('left');

  var centerCols = ['Duration','ChapterCount','CurrentChapterIndex','PlaybackSpeed','TotalListeningMins'];
  centerCols.forEach(function(h) {
    var col = AUDIOBOOK_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment('center');
  });

  // ─── Helper column formulas ───
  var chapCountCol = _colLetter(AUDIOBOOK_HEADERS.indexOf('ChapterCount') + 1);
  var chapIdxCol   = _colLetter(AUDIOBOOK_HEADERS.indexOf('CurrentChapterIndex') + 1);
  var totalMinCol  = _colLetter(AUDIOBOOK_HEADERS.indexOf('TotalListeningMins') + 1);
  for (var r = 2; r <= lastRow; r++) {
    // Listening progress bar (chapters listened / total chapters)
    sheet.getRange(r, numCols + 1)
      .setFormula('=IFERROR(SPARKLINE(' + chapIdxCol + r + '/MAX(1,' + chapCountCol + r + '),{"charttype","bar";"max",1;"color1","' + palette.accent + '";"color2","' + palette.border + '"}),"")');
    // Chapter % text
    sheet.getRange(r, numCols + 2)
      .setFormula('=IFERROR(IF(' + chapCountCol + r + '>0,ROUND(100*' + chapIdxCol + r + '/' + chapCountCol + r + ',0)&"%",""),"")');
  }
  sheet.getRange(2, numCols + 1, lastRow - 1, 1).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.getRange(2, numCols + 2, lastRow - 1, 1).setHorizontalAlignment('center').setFontWeight('bold').setFontFamily('Google Sans').setFontSize(10).setFontColor(palette.titleCol || palette.header);
  // Style helper cols background — batched write
  var audioHelperBg = [];
  for (var br = 2; br <= lastRow; br++) {
    var aBand = (br % 2 === 0) ? palette.altRow : palette.bg;
    audioHelperBg.push([aBand, aBand]);
  }
  sheet.getRange(2, numCols + 1, audioHelperBg.length, 2).setBackgrounds(audioHelperBg);

  // Header notes
  var audioNotes = {
    'Title': 'Audiobook title',
    'Author': 'Author or narrator name',
    'Duration': 'Total runtime (HH:MM:SS)',
    'ChapterCount': 'Total number of chapters',
    'CurrentChapterIndex': 'Chapter index you are currently on (0-based)',
    'PlaybackSpeed': 'Speed multiplier  (1.0 = normal, 1.5 = fast)',
    'TotalListeningMins': 'Total minutes you have listened to this book'
  };
  Object.keys(audioNotes).forEach(function(h) {
    var col = AUDIOBOOK_HEADERS.indexOf(h) + 1;
    if (col > 0) sheet.getRange(1, col).setNote(audioNotes[h]);
  });

  var existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  existing.forEach(function(p) { if (p.getDescription() === 'Header row') p.remove(); });
  sheet.getRange(1, 1, 1, numCols + 2).protect().setDescription('Header row').setWarningOnly(true);
}

// ── Stats Sheet (live analytics presentation tab) ───────────────────────
function _initStatsSheet(theme) {
  var ss = _ss();
  var sheet = ss.getSheetByName(SHEET_STATS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_STATS);
  } else {
    sheet.clear();
    sheet.getBandings().forEach(function(b) { b.remove(); });
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(pp) { pp.remove(); });
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(pp) { pp.remove(); });
    try { sheet.getCharts().forEach(function(c) { sheet.removeChart(c); }); } catch(e){}
  }

  var p   = _getThemePalette(theme);
  var lib = SHEET_LIBRARY;
  var prf = SHEET_PROFILE;
  var year = new Date().getFullYear();
  var NC = 14;
  _ensureColumns(sheet, NC);
  _ensureRows(sheet, 80);

  // ── Canvas setup ──
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(p.tabColor || p.header);
  sheet.setColumnWidth(1, 20);   // left gutter
  sheet.setColumnWidth(14, 20);  // right gutter
  for (var c = 2; c <= 13; c++) sheet.setColumnWidth(c, 118);
  if (sheet.getMaxColumns() > NC) sheet.deleteColumns(NC + 1, sheet.getMaxColumns() - NC);
  var totalRows = 50;
  if (sheet.getMaxRows() < totalRows) sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  if (sheet.getMaxRows() > totalRows) sheet.deleteRows(totalRows + 1, sheet.getMaxRows() - totalRows);

  sheet.getRange(1, 1, totalRows, NC).setBackground(p.bg);

  var softPanel = p.altRow;
  var headerBg  = p.header;
  var headerTxt = p.headerText;
  var textCol   = p.text || '#1F2937';
  var subtleCol = p.subtleText || '#6B7280';
  var accent    = p.accent;
  var border    = p.border;
  var titleCol  = p.titleCol || p.header;

  // ╔═══ HERO BANNER (rows 1-3) ═══
  sheet.setRowHeight(1, 90);
  sheet.getRange(1, 2, 1, 12).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(28).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setFormula('=IFERROR(IF(\'' + prf + '\'!A2<>"",\'' + prf + '\'!A2&"\'s Reading Stats","Reading Stats"),"Reading Stats")');

  sheet.setRowHeight(2, 26);
  sheet.getRange(2, 2, 1, 12).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('normal')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('Live analytics · refreshes automatically when books are added or updated');

  sheet.setRowHeight(3, 6);
  sheet.getRange(3, 2, 1, 12).merge().setBackground(accent);

  sheet.setRowHeight(4, 12);

  // ╔═══ KPI ROW 1 — Library at a glance (rows 5-7) ═══
  sheet.setRowHeight(5, 32);
  sheet.getRange(5, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  LIBRARY AT A GLANCE');

  var kpis1 = [
    { col:2,  val: '=IFERROR(COUNTA(\'' + lib + '\'!B2:B),0)',
      label: 'TOTAL BOOKS' },
    { col:5,  val: '=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Finished"),0)',
      label: 'FINISHED' },
    { col:8,  val: '=IFERROR(TEXT(AVERAGEIF(\'' + lib + '\'!E:E,">0",\'' + lib + '\'!E:E),"0.0")&" ★","—")',
      label: 'AVG RATING' },
    { col:11, val: '=IFERROR(COUNTIFS(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!J:J,">="&DATE(' + year + ',1,1)),0)',
      label: year + ' BOOKS' }
  ];
  sheet.setRowHeight(6, 64);
  sheet.setRowHeight(7, 24);
  kpis1.forEach(function(k) {
    sheet.getRange(6, k.col, 1, 3).merge()
      .setBackground(softPanel).setFontColor(titleCol)
      .setFontFamily('Google Sans').setFontSize(28).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula(k.val);
    sheet.getRange(7, k.col, 1, 3).merge()
      .setBackground(softPanel).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setValue(k.label);
    sheet.getRange(6, k.col, 2, 3)
      .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, k.col, 1, 3)
      .setBorder(true, null, null, null, null, null, accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  sheet.setRowHeight(8, 12);

  // ╔═══ KPI ROW 2 — Reading depth (rows 9-11) ═══
  sheet.setRowHeight(9, 32);
  sheet.getRange(9, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  READING DEPTH');

  var kpis2 = [
    { col:2,  val: '=IFERROR(TEXT(SUMIF(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!F:F),"#,##0"),"0")',
      label: 'PAGES READ' },
    { col:5,  val: '=IFERROR(TEXT(ROUND(AVERAGEIF(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!F:F),0),"#,##0"),"0")',
      label: 'AVG PAGES/BOOK' },
    { col:8,  val: '=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Reading"),0)',
      label: 'ACTIVE READS' },
    { col:11, val: '=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Want to Read"),0)',
      label: 'ON THE TBR' }
  ];
  sheet.setRowHeight(10, 64);
  sheet.setRowHeight(11, 24);
  kpis2.forEach(function(k) {
    sheet.getRange(10, k.col, 1, 3).merge()
      .setBackground(softPanel).setFontColor(titleCol)
      .setFontFamily('Google Sans').setFontSize(28).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula(k.val);
    sheet.getRange(11, k.col, 1, 3).merge()
      .setBackground(softPanel).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setValue(k.label);
    sheet.getRange(10, k.col, 2, 3)
      .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(10, k.col, 1, 3)
      .setBorder(true, null, null, null, null, null, accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  sheet.setRowHeight(12, 14);

  // ╔═══ BOOKS BY GENRE (rows 13-24) ═══
  sheet.setRowHeight(13, 32);
  sheet.getRange(13, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  BOOKS BY GENRE');

  // Sub-header
  sheet.setRowHeight(14, 32);
  sheet.getRange(14, 2, 1, 4).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setValue('  GENRE');
  sheet.getRange(14, 6, 1, 5).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('DISTRIBUTION');
  sheet.getRange(14, 11, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setValue('COUNT  ');

  // 10 genre rows
  var genreQuery = 'QUERY(\'' + lib + '\'!G2:G,"select G, count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 10 label G \'\', count(G) \'\'",0)';
  var genreMaxQ  = 'QUERY(\'' + lib + '\'!G2:G,"select count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 1 label count(G) \'\'",0)';
  for (var gi = 0; gi < 10; gi++) {
    var gRow = 15 + gi;
    sheet.setRowHeight(gRow, 26);
    var rowBg = (gi % 2 === 0) ? softPanel : p.bg;
    sheet.getRange(gRow, 2, 1, 4).merge()
      .setBackground(rowBg).setFontColor(textCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setFormula('=IFERROR(INDEX(' + genreQuery + ',' + (gi + 1) + ',1),"")');
    sheet.getRange(gRow, 6, 1, 5).merge()
      .setBackground(rowBg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(SPARKLINE(INDEX(' + genreQuery + ',' + (gi + 1) + ',2)/MAX(1,INDEX(' + genreMaxQ + ',1,1)),{"charttype","bar";"max",1;"color1","' + accent + '";"color2","' + border + '"}),"")');
    sheet.getRange(gRow, 11, 1, 2).merge()
      .setBackground(rowBg).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle')
      .setFormula('=IFERROR(INDEX(' + genreQuery + ',' + (gi + 1) + ',2),"")');
  }
  sheet.getRange(14, 2, 11, 11)
    .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(25, 14);

  // ╔═══ STATUS BREAKDOWN (rows 26-31) ═══
  sheet.setRowHeight(26, 32);
  sheet.getRange(26, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  STATUS BREAKDOWN');

  sheet.setRowHeight(27, 32);
  sheet.getRange(27, 2, 1, 4).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setValue('  STATUS');
  sheet.getRange(27, 6, 1, 5).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('SHARE OF LIBRARY');
  sheet.getRange(27, 11, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setValue('COUNT  ');

  var statusRows = [
    { label:'Finished',     count:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Finished"),0)',      color:'#166534' },
    { label:'Reading',      count:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Reading"),0)',       color:'#1E40AF' },
    { label:'Want to Read', count:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Want to Read"),0)',  color:'#92400E' },
    { label:'DNF',          count:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"DNF"),0)',           color:'#991B1B' }
  ];
  var totalBooksF = 'MAX(1,COUNTA(\'' + lib + '\'!B2:B))';
  statusRows.forEach(function(sr, si) {
    var sRow = 28 + si;
    sheet.setRowHeight(sRow, 28);
    var rowBg = (si % 2 === 0) ? softPanel : p.bg;
    sheet.getRange(sRow, 2, 1, 4).merge()
      .setBackground(rowBg).setFontColor(sr.color)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setValue('  ' + sr.label);
    // Bar — share of total library
    var countExpr = sr.count.replace(/^=/, '');
    sheet.getRange(sRow, 6, 1, 5).merge()
      .setBackground(rowBg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(SPARKLINE((' + countExpr + ')/(' + totalBooksF + '),{"charttype","bar";"max",1;"color1","' + accent + '";"color2","' + border + '"}),"")');
    sheet.getRange(sRow, 11, 1, 2).merge()
      .setBackground(rowBg).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle')
      .setFormula(sr.count);
  });
  sheet.getRange(27, 2, 5, 11)
    .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(32, 14);

  // ╔═══ YEAR OVER YEAR (rows 33-41) ═══
  sheet.setRowHeight(33, 32);
  sheet.getRange(33, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  YEAR OVER YEAR');

  sheet.setRowHeight(34, 32);
  sheet.getRange(34, 2, 1, 3).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('YEAR');
  sheet.getRange(34, 5, 1, 5).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('BOOKS FINISHED (vs. yearly goal)');
  sheet.getRange(34, 10, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('COUNT');
  sheet.getRange(34, 12, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('PAGES');

  for (var yi = 0; yi < 7; yi++) {
    var yRow = 35 + yi;
    var targetYear = year - 6 + yi;  // oldest to newest (current = last)
    sheet.setRowHeight(yRow, 28);
    var rowBg = (yi % 2 === 0) ? softPanel : p.bg;
    var isCurrent = (targetYear === year);
    var yearColor = isCurrent ? titleCol : textCol;
    var countF = 'COUNTIFS(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!J:J,">="&DATE(' + targetYear + ',1,1),\'' + lib + '\'!J:J,"<="&DATE(' + targetYear + ',12,31))';

    sheet.getRange(yRow, 2, 1, 3).merge()
      .setBackground(rowBg).setFontColor(yearColor)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight(isCurrent ? 'bold' : 'normal')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setValue(targetYear);
    // Bar proportional to yearly goal (from Profile!E2, fallback 50)
    sheet.getRange(yRow, 5, 1, 5).merge()
      .setBackground(rowBg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(SPARKLINE(' + countF + '/MAX(1,IFERROR(\'' + prf + '\'!E2,50)),{"charttype","bar";"max",1;"color1","' + accent + '";"color2","' + border + '"}),"")');
    sheet.getRange(yRow, 10, 1, 2).merge()
      .setBackground(rowBg).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight(isCurrent ? 'bold' : 'normal')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(' + countF + ',0)');
    sheet.getRange(yRow, 12, 1, 2).merge()
      .setBackground(rowBg).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(TEXT(SUMPRODUCT((\'' + lib + '\'!D2:D="Finished")*(YEAR(\'' + lib + '\'!J2:J)=' + targetYear + ')*(\'' + lib + '\'!F2:F)),"#,##0"),0)');
  }
  sheet.getRange(34, 2, 8, 13)
    .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(42, 14);

  // Hide rows beyond content so tab feels tight
  if (sheet.getMaxRows() > 43) {
    try { sheet.hideRows(43, sheet.getMaxRows() - 43); } catch(e){}
  }

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(0);

  sheet.protect().setDescription('Stats — live analytics, do not edit').setWarningOnly(true);
}

// ── Cover / Welcome Page ────────────────────────────────────────────────
function _initCoverSheet(theme) {
  var ss = _ss();
  var sheet = ss.getSheetByName(SHEET_COVER);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COVER);
  } else {
    sheet.clear();
    sheet.getBandings().forEach(function(b) { b.remove(); });
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(prot) { prot.remove(); });
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(prot) { prot.remove(); });
  }

  var p = _getThemePalette(theme);
  var lib = SHEET_LIBRARY;
  var prf = SHEET_PROFILE;
  var year = new Date().getFullYear();
  var NC = 14;
  var webAppUrl = '';
  try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e2) {}
  _ensureColumns(sheet, NC);
  _ensureRows(sheet, 90);

  sheet.setHiddenGridlines(true);
  sheet.setTabColor(p.tabColor || p.header);
  sheet.setColumnWidth(1, 16);   // left gutter
  sheet.setColumnWidth(14, 16);  // right gutter
  for (var c = 2; c <= 13; c++) sheet.setColumnWidth(c, 112);
  if (sheet.getMaxColumns() > NC) sheet.deleteColumns(NC + 1, sheet.getMaxColumns() - NC);

  sheet.getRange(1, 1, 90, NC).setBackground(p.bg);

  sheet.setRowHeight(1, 64);
  sheet.getRange(1, 2, 1, 12).merge()
    .setBackground(p.header).setFontColor(p.headerText)
    .setFontFamily('Google Sans').setFontSize(27).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setFormula('=IFERROR(IF(\'' + prf + '\'!A2<>"",\'' + prf + '\'!A2&"\'s Reading Database","Reading Database"),"Reading Database")');

  sheet.setRowHeight(2, 30);
  sheet.getRange(2, 2, 1, 12).merge()
    .setBackground(p.header).setFontColor(p.headerText)
    .setFontFamily('Google Sans').setFontSize(11)
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setFormula('=IFERROR(COUNTA(\'' + lib + '\'!A2:A)&" books tracked • "&COUNTIFS(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!J:J,">="&DATE(' + year + ',1,1))&" finished this year • Avg "&ROUND(AVERAGEIF(\'' + lib + '\'!E:E,">0",\'' + lib + '\'!E:E),1)&" stars","Reading analytics ready")');

  sheet.setRowHeight(3, 8);
  sheet.getRange(3, 2, 1, 12).merge().setBackground(p.accent);

  sheet.setRowHeight(5, 26);
  sheet.getRange(5, 2, 1, 12).merge()
    .setBackground(p.bg).setFontColor(p.accent)
    .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('Executive Snapshot');

  var tiles = [
    { col:2,  label:'Total Books', formula:'=IFERROR(COUNTA(\'' + lib + '\'!A2:A),0)' },
    { col:5,  label:'Finished', formula:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Finished"),0)' },
    { col:8,  label:'Reading', formula:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Reading"),0)' },
    { col:11, label:'Avg Rating', formula:'=IFERROR(ROUND(AVERAGEIF(\'' + lib + '\'!E:E,">0",\'' + lib + '\'!E:E),1),0)' }
  ];

  sheet.setRowHeight(6, 52);
  sheet.setRowHeight(7, 22);
  tiles.forEach(function(tile) {
    sheet.getRange(6, tile.col, 1, 3).merge()
      .setBackground(p.altRow).setFontColor(p.text)
      .setFontFamily('Google Sans').setFontSize(22).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula(tile.formula)
      .setBorder(true, true, true, true, false, false, p.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.getRange(7, tile.col, 1, 3).merge()
      .setBackground(p.altRow).setFontColor(p.subtleText || '#6B7280')
      .setFontFamily('Google Sans').setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setValue(tile.label)
      .setBorder(true, true, true, true, false, false, p.border, SpreadsheetApp.BorderStyle.SOLID);
  });

  sheet.setRowHeight(8, 12);
  sheet.setRowHeight(9, 26);
  sheet.getRange(9, 2, 1, 12).merge()
    .setBackground(p.bg).setFontColor(p.accent)
    .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('Top Genres');

  sheet.setRowHeight(10, 34);
  sheet.getRange(10, 2, 1, 12).merge()
    .setBackground(p.altRow).setFontColor(p.text)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setFormula('=IFERROR(TEXTJOIN("  •  ",TRUE,QUERY({' +
      'FILTER(\'' + lib + '\'!G2:G,\'' + lib + '\'!G2:G<>""),' +
      'FILTER(\'' + lib + '\'!G2:G,\'' + lib + '\'!G2:G<>"")},' +
      '"select Col1, count(Col2) group by Col1 order by count(Col2) desc limit 3 label count(Col2) \"\"",0)),' +
      '"Add books to build genre analytics")')
    .setBorder(true, true, true, true, false, false, p.border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(12, 24);
  sheet.getRange(12, 2, 1, 12).merge()
    .setBackground(p.bg).setFontColor(p.accent)
    .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('Launch');

  sheet.setRowHeight(13, 54);
  var launchCell = sheet.getRange(13, 2, 1, 12).merge()
    .setBackground(p.altRow).setFontColor(p.text)
    .setFontFamily('Google Sans').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, p.border, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  if (webAppUrl) {
    launchCell.setFormula('=HYPERLINK("' + webAppUrl + '","Open Reading Web App")');
  } else {
    launchCell.setValue('Use the custom menu to open the web app.');
  }

  sheet.setRowHeight(15, 22);
  sheet.getRange(15, 2, 1, 12).merge()
    .setBackground(p.bg).setFontColor(p.subtleText || '#6B7280')
    .setFontFamily('Google Sans').setFontSize(8)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setValue('Data is stored directly in this spreadsheet and surfaced in the web app UI. Library tab is optimized as the operational database.');

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(0);

  var sheetExisting = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  sheetExisting.forEach(function(prot) { prot.remove(); });
  sheet.protect().setDescription('Cover - do not edit directly').setWarningOnly(true);
}

// ── Dashboard (hero presentation tab) ───────────────────────────────────
function _initDashboardSheet(theme) {
  var ss = _ss();
  var sheet = ss.getSheetByName(SHEET_DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DASHBOARD);
  } else {
    sheet.clear();
    sheet.getBandings().forEach(function(b) { b.remove(); });
    sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(pp) { pp.remove(); });
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(pp) { pp.remove(); });
    try { sheet.getCharts().forEach(function(c) { sheet.removeChart(c); }); } catch(e){}
  }

  var p    = _getThemePalette(theme);
  var lib  = SHEET_LIBRARY;
  var prf  = SHEET_PROFILE;
  var chal = SHEET_CHALLENGES;
  var year = new Date().getFullYear();
  var NC   = 14; // A gutter + 12 data cols + N gutter
  var webAppUrl = '';
  try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) {}
  _ensureColumns(sheet, NC);
  _ensureRows(sheet, 120);

  // ── Canvas setup ──
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(p.tabColor || p.header);
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(14, 20);
  for (var c = 2; c <= 13; c++) sheet.setColumnWidth(c, 118);
  if (sheet.getMaxColumns() > NC) sheet.deleteColumns(NC + 1, sheet.getMaxColumns() - NC);
  if (sheet.getMaxRows() < 40) sheet.insertRowsAfter(sheet.getMaxRows(), 40 - sheet.getMaxRows());
  if (sheet.getMaxRows() > 40) sheet.deleteRows(41, sheet.getMaxRows() - 40);

  // Paint the entire canvas with the theme bg
  sheet.getRange(1, 1, 40, NC).setBackground(p.bg);

  var softPanel  = p.altRow;
  var headerBg   = p.header;
  var headerTxt  = p.headerText;
  var textCol    = p.text || '#1F2937';
  var subtleCol  = p.subtleText || '#6B7280';
  var accent     = p.accent;
  var border     = p.border;
  var titleCol   = p.titleCol || p.header;

  // ╔═══ HERO BANNER (rows 1-3) ═══
  sheet.setRowHeight(1, 90);
  sheet.getRange(1, 2, 1, 12).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(28).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setFormula('=IFERROR(IF(\'' + prf + '\'!A2<>"",\'' + prf + '\'!A2&"\'s Reading Journey","My Reading Journey"),"My Reading Journey")');

  sheet.setRowHeight(2, 26);
  sheet.getRange(2, 2, 1, 12).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('normal')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setFormula('=IFERROR(IF(\'' + prf + '\'!B2<>"",\'' + prf + '\'!B2,"Where stories come alive"),"Where stories come alive")');

  sheet.setRowHeight(3, 6);
  sheet.getRange(3, 2, 1, 12).merge().setBackground(accent);

  sheet.setRowHeight(4, 12);

  // ╔═══ SNAPSHOT — 4 KPI cards (rows 5-7) ═══
  sheet.setRowHeight(5, 32);
  sheet.getRange(5, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  MY LIBRARY');

  var kpis = [
    { col:2,  label:'BOOKS',     formula:'=IFERROR(COUNTA(\'' + lib + '\'!B2:B),0)' },
    { col:5,  label:'FINISHED',  formula:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Finished"),0)' },
    { col:8,  label:'READING',   formula:'=IFERROR(COUNTIF(\'' + lib + '\'!D:D,"Reading"),0)' },
    { col:11, label:'AVG RATING',formula:'=IFERROR(TEXT(AVERAGEIF(\'' + lib + '\'!E:E,">0",\'' + lib + '\'!E:E),"0.0")&" ★","—")' }
  ];
  sheet.setRowHeight(6, 64);
  sheet.setRowHeight(7, 24);
  kpis.forEach(function(k) {
    sheet.getRange(6, k.col, 1, 3).merge()
      .setBackground(softPanel).setFontColor(titleCol)
      .setFontFamily('Google Sans').setFontSize(28).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula(k.formula);
    sheet.getRange(7, k.col, 1, 3).merge()
      .setBackground(softPanel).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setValue(k.label);
    sheet.getRange(6, k.col, 2, 3)
      .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, k.col, 1, 3)
      .setBorder(true, null, null, null, null, null, accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });

  sheet.setRowHeight(8, 14);

  // ╔═══ YEAR GOAL (left) | TOP GENRES (right) (rows 9-14) ═══
  sheet.setRowHeight(9, 32);
  sheet.getRange(9, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  ' + year + ' READING GOAL  ·  TOP GENRES');

  // Left panel: goal progress card (rows 10-14, cols B-G)
  sheet.setRowHeight(10, 54);
  sheet.getRange(10, 2, 1, 6).merge()
    .setBackground(softPanel).setFontColor(titleCol)
    .setFontFamily('Google Sans').setFontSize(24).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFormula('=IFERROR(COUNTIFS(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!J:J,">="&DATE(' + year + ',1,1))&" / "&\'' + prf + '\'!E2,"—")')
    .setBorder(true, true, false, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(11, 20);
  sheet.getRange(11, 2, 1, 6).merge()
    .setBackground(softPanel).setFontColor(subtleCol)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setValue('BOOKS COMPLETED')
    .setBorder(false, true, false, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(12, 22);
  sheet.getRange(12, 2, 1, 6).merge()
    .setBackground(softPanel)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFormula('=IFERROR(SPARKLINE(COUNTIFS(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!J:J,">="&DATE(' + year + ',1,1))/MAX(1,\'' + prf + '\'!E2),{"charttype","bar";"max",1;"color1","' + accent + '";"color2","' + border + '"}),"")')
    .setBorder(false, true, false, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(13, 22);
  sheet.getRange(13, 2, 1, 6).merge()
    .setBackground(softPanel).setFontColor(titleCol)
    .setFontFamily('Google Sans').setFontSize(12).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFormula('=IFERROR(ROUND(100*COUNTIFS(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!J:J,">="&DATE(' + year + ',1,1))/MAX(1,\'' + prf + '\'!E2),0)&"% complete","0% complete")')
    .setBorder(false, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(14, 4);

  // Right panel: top 5 genres (rows 10-14, cols H-M)
  var genreFormula = '=IFERROR(ARRAYFORMULA(IF(LEN(QUERY(\'' + lib + '\'!G2:G,"select G, count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 5 label G \'\', count(G) \'\'",0))>0,' +
                     'QUERY(\'' + lib + '\'!G2:G,"select G, count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 5 label G \'\', count(G) \'\'",0),"")),"Add books to see genres")';
  sheet.setRowHeight(10, 54); // already set; keep
  // Render genre list: 5 rows × 6 cols (H-M), each row = genre name + count + bar
  for (var gi = 0; gi < 5; gi++) {
    var gRow = 10 + gi;
    sheet.setRowHeight(gRow, 22);
    // Genre name (H-I)
    sheet.getRange(gRow, 8, 1, 2).merge()
      .setBackground(softPanel).setFontColor(textCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle')
      .setFormula('=IFERROR(INDEX(QUERY(\'' + lib + '\'!G2:G,"select G, count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 5 label G \'\', count(G) \'\'",0),' + (gi + 1) + ',1),"")');
    // Bar (J-L) — sparkline bar of count
    sheet.getRange(gRow, 10, 1, 3).merge()
      .setBackground(softPanel)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(SPARKLINE(INDEX(QUERY(\'' + lib + '\'!G2:G,"select G, count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 5 label G \'\', count(G) \'\'",0),' + (gi + 1) + ',2)/MAX(1,MAX(QUERY(\'' + lib + '\'!G2:G,"select count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 1 label count(G) \'\'",0))),{"charttype","bar";"max",1;"color1","' + accent + '";"color2","' + border + '"}),"")');
    // Count (M)
    sheet.getRange(gRow, 13)
      .setBackground(softPanel).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('right').setVerticalAlignment('middle')
      .setFormula('=IFERROR(INDEX(QUERY(\'' + lib + '\'!G2:G,"select G, count(G) where G is not null and G <> \'\' group by G order by count(G) desc limit 5 label G \'\', count(G) \'\'",0),' + (gi + 1) + ',2),"")');
  }
  // Border the genre panel
  sheet.getRange(10, 8, 5, 6)
    .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(15, 14);

  // ╔═══ CURRENTLY READING (left) | RECENTLY FINISHED (right) (rows 16-22) ═══
  sheet.setRowHeight(16, 32);
  sheet.getRange(16, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  NOW READING  ·  JUST FINISHED');

  // Column headers — redesigned to include cover thumbnail slot
  sheet.setRowHeight(17, 24);
  sheet.getRange(17, 2)
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(8).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('CVR');
  sheet.getRange(17, 3, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setValue('  TITLE');
  sheet.getRange(17, 5, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('PROGRESS');
  sheet.getRange(17, 7)
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('%');
  sheet.getRange(17, 8)
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(8).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('CVR');
  sheet.getRange(17, 9, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setValue('  TITLE');
  sheet.getRange(17, 11, 1, 2).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle').setValue('  AUTHOR');
  sheet.getRange(17, 13)
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setValue('★');

  // 5 rows of data each side (rows 18-22) — 58px rows to show cover thumbnails
  for (var ri = 0; ri < 5; ri++) {
    var dRow = 18 + ri;
    sheet.setRowHeight(dRow, 72);
    var rowBg = (ri % 2 === 0) ? softPanel : p.bg;

    // LEFT: Currently Reading
    // Cover col (B) — IMAGE from CoverUrl (col Y) of the matching reading book
    sheet.getRange(dRow, 2)
      .setBackground(rowBg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(IMAGE(INDEX(QUERY(\'' + lib + '\'!B2:Y,"select Y where D = \'Reading\' order by I desc limit 5",0),' + (ri + 1) + ',1),4,52,39),"")');
    // Title (C-D)
    sheet.getRange(dRow, 3, 1, 2).merge()
      .setBackground(rowBg).setFontColor(textCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
      .setFormula('=IFERROR(" "&INDEX(QUERY(\'' + lib + '\'!B2:K,"select B where D = \'Reading\' order by I desc limit 5",0),' + (ri + 1) + ',1),"")');
    // Progress sparkline (E-F)
    sheet.getRange(dRow, 5, 1, 2).merge()
      .setBackground(rowBg)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(SPARKLINE(INDEX(QUERY(\'' + lib + '\'!B2:K,"select K/F where D = \'Reading\' order by I desc limit 5 label K/F \'\'",0),' + (ri + 1) + ',1),{"charttype","bar";"max",1;"color1","' + accent + '";"color2","' + border + '"}),"")');
    // % complete (G)
    sheet.getRange(dRow, 7)
      .setBackground(rowBg).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(ROUND(100*INDEX(QUERY(\'' + lib + '\'!B2:K,"select K/F where D = \'Reading\' order by I desc limit 5 label K/F \'\'",0),' + (ri + 1) + ',1),0)&"%","")');

    // RIGHT: Recently Finished
    // Cover col (H) — IMAGE from CoverUrl (col Y) of the matching finished book
    sheet.getRange(dRow, 8)
      .setBackground(rowBg).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(IMAGE(INDEX(QUERY(\'' + lib + '\'!B2:Y,"select Y where D = \'Finished\' and J is not null order by J desc limit 5",0),' + (ri + 1) + ',1),4,52,39),"")');
    // Title (I-J)
    sheet.getRange(dRow, 9, 1, 2).merge()
      .setBackground(rowBg).setFontColor(textCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
      .setFormula('=IFERROR(" "&INDEX(QUERY(\'' + lib + '\'!B2:J,"select B where D = \'Finished\' and J is not null order by J desc limit 5",0),' + (ri + 1) + ',1),"")');
    // Author (K-L)
    sheet.getRange(dRow, 11, 1, 2).merge()
      .setBackground(rowBg).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('normal')
      .setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
      .setFormula('=IFERROR(" "&INDEX(QUERY(\'' + lib + '\'!B2:J,"select C where D = \'Finished\' and J is not null order by J desc limit 5",0),' + (ri + 1) + ',1),"")');
    // Rating (M)
    sheet.getRange(dRow, 13)
      .setBackground(rowBg).setFontColor('#D97706')
      .setFontFamily('Google Sans').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula('=IFERROR(REPT("★",INDEX(QUERY(\'' + lib + '\'!B2:J,"select E where D = \'Finished\' and J is not null order by J desc limit 5 label E \'\'",0),' + (ri + 1) + ',1)),"")');
  }
  sheet.getRange(17, 2, 6, 6)
    .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(17, 8, 6, 6)
    .setBorder(true, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(23, 14);

  // ╔═══ 12-MONTH READING VELOCITY (rows 24-27) ═══
  sheet.setRowHeight(24, 32);
  sheet.getRange(24, 1, 1, NC).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(11).setFontWeight('bold')
    .setHorizontalAlignment('left').setVerticalAlignment('middle')
    .setValue('  ◼  12-MONTH READING VELOCITY');

  // Hidden monthly helper vector on row 25 — transparent to the eye via matching bg/font
  sheet.setRowHeight(25, 2);
  for (var mi = 0; mi < 12; mi++) {
    var helperCol = 2 + mi;
    var mo = mi - 11;
    var vf = '=COUNTIFS(\'' + lib + '\'!D:D,"Finished",\'' + lib + '\'!J:J,">="&EOMONTH(TODAY(),' + (mo - 1) + ')+1,\'' + lib + '\'!J:J,"<="&EOMONTH(TODAY(),' + mo + '))';
    sheet.getRange(25, helperCol).setFormula(vf).setFontColor(p.bg).setBackground(p.bg).setFontSize(7);
  }

  sheet.setRowHeight(26, 58);
  sheet.getRange(26, 2, 1, 12).merge()
    .setBackground(softPanel)
    .setFormula('=SPARKLINE(B25:M25,{"charttype","column";"color","' + accent + '";"highcolor","' + titleCol + '";"empty","zero"})')
    .setBorder(true, true, false, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);

  // Month labels under the chart
  sheet.setRowHeight(27, 20);
  for (var lb = 0; lb < 12; lb++) {
    var lbCol = 2 + lb;
    var labelFormula = '=UPPER(LEFT(TEXT(EOMONTH(TODAY(),' + (lb - 11) + '),"mmm"),3))';
    sheet.getRange(27, lbCol)
      .setBackground(softPanel).setFontColor(subtleCol)
      .setFontFamily('Google Sans').setFontSize(8).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setFormula(labelFormula);
  }
  sheet.getRange(27, 2, 1, 12)
    .setBorder(false, true, true, true, false, false, border, SpreadsheetApp.BorderStyle.SOLID);

  sheet.setRowHeight(28, 14);

  // ╔═══ LAUNCH CTA (rows 29-30) ═══
  sheet.setRowHeight(29, 56);
  var cta = sheet.getRange(29, 2, 1, 12).merge()
    .setBackground(headerBg).setFontColor(headerTxt)
    .setFontFamily('Google Sans').setFontSize(14).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  if (webAppUrl) {
    cta.setFormula('=HYPERLINK("' + webAppUrl + '","➜  Open Reading Web App")');
  } else {
    cta.setValue('Open Reading Web App from the custom menu');
  }

  sheet.setRowHeight(30, 20);
  sheet.getRange(30, 2, 1, 12).merge()
    .setBackground(p.bg).setFontColor(subtleCol)
    .setFontFamily('Google Sans').setFontSize(8).setFontStyle('italic')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setValue('All metrics refresh automatically when you add, rate, or finish books in the web app.');

  // Hide extra rows beyond viewport so the tab feels tight
  if (sheet.getMaxRows() > 31) {
    try { sheet.hideRows(31, sheet.getMaxRows() - 31); } catch(e){}
  }

  sheet.setFrozenRows(3);
  sheet.setFrozenColumns(0);

  sheet.protect().setDescription('Dashboard protection').setWarningOnly(true);
}

// ── Re-style all sheets when theme changes ──────────────────────────────
function _reStyleAllSheets(themeName) {
  if (typeof _dbLiteInitializeSheets === 'function') {
    _dbLiteInitializeSheets();
    _incrementSyncVersion();
    return;
  }

  // Full re-init each sheet with the new theme — this updates tab colors,
  // header/data row backgrounds, title font tint, banding, and all styling.
  // Isolate each init so one failure cannot block the rest of the theme switch.
  function _safe(name, fn) {
    try { fn(); }
    catch (err) { _log('error', '_reStyleAllSheets', name + ' failed: ' + (err && err.message ? err.message : err)); }
  }
  _safe('Cover',      function(){ _initCoverSheet(themeName); });
  _safe('Dashboard',  function(){ _initDashboardSheet(themeName); });
  _safe('Stats',      function(){ _initStatsSheet(themeName); });
  _safe('Library',    function(){ _initLibrarySheet(themeName); });
  _safe('Challenges', function(){ _initChallengesSheet(themeName); });
  _safe('Shelves',    function(){ _initShelvesSheet(themeName); });
  _safe('Profile',    function(){ _initProfileSheet(themeName); });
  _safe('Audiobooks', function(){ _initAudiobooksSheet(themeName); });
  // Bump sync version so webapp detects the theme change
  _incrementSyncVersion();
}

// ── Status Mapping ──────────────────────────────────────────────────────
function _uiStatusToSheet(status) {
  var map = { 'finished':'Finished', 'reading':'Reading', 'want-to-read':'Want to Read', 'dnf':'DNF' };
  return map[String(status).toLowerCase()] || 'Want to Read';
}
function _sheetStatusToUi(status) {
  var val = String(status || '').toLowerCase();
  if (val === 'finished' || val === 'read') return 'finished';
  if (val === 'reading' || val === 'listening') return 'reading';
  if (val === 'dnf') return 'dnf';
  return 'want-to-read';
}

// =====================================================================
//  CLIENT API — called from UI via google.script.run
// =====================================================================

/** Bootstrap: returns all data for initial render */
function clientGetInitialData() {
  // Only run full sheet initialization when sheets don't yet exist (first run).
  // Skipping on every subsequent load avoids ~250 unnecessary Sheets API calls.
  var _profileCheck = _ss().getSheetByName(SHEET_PROFILE);
  if (!_profileCheck || _profileCheck.getLastRow() < 2) {
    initializeSheets();
  }

  var libSheet = _ss().getSheetByName(SHEET_LIBRARY);
  var chalSheet = _ss().getSheetByName(SHEET_CHALLENGES);
  var shelfSheet = _ss().getSheetByName(SHEET_SHELVES);
  var profileSheet = _ss().getSheetByName(SHEET_PROFILE);
  var audioSheet = _ss().getSheetByName(SHEET_AUDIOBOOKS);

  var library = _sheetToObjects(libSheet, LIBRARY_HEADERS).map(function(row) {
    var rawIsbn = String(row.ISBN || '').trim();
    var isbnNorm = _normalizeIsbn(rawIsbn);
    var coverPrimary = row.CoverUrl || '';
    var isbnCoverUrl = isbnNorm ? ('https://covers.openlibrary.org/b/isbn/' + isbnNorm + '-L.jpg') : '';
    var coverFallback = (isbnCoverUrl && coverPrimary !== isbnCoverUrl) ? isbnCoverUrl : '';
    return {
      BookId:           row.BookId,
      Title:            row.Title,
      Author:           row.Author,
      Status:           row.Status,
      Rating:           Number(row.Rating) || 0,
      PageCount:        Number(row.Pages) || 0,
      Genres:           row.Genre || '',
      DateAdded:        row.DateAdded ? Utilities.formatDate(new Date(row.DateAdded), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      DateStarted:      row.DateStarted ? Utilities.formatDate(new Date(row.DateStarted), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      DateFinished:     row.DateFinished ? Utilities.formatDate(new Date(row.DateFinished), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      CurrentPage:      Number(row.CurrentPage) || 0,
      Series:           row.Series || '',
      SeriesOrder:      row.SeriesNumber || '',
      TbrPriority:      Number(row.TbrPriority) || 0,
      Format:           row.Format || '',
      Source:           row.Source || '',
      SpiceLevel:       Number(row.SpiceLevel) || 0,
      Moods:            row.Tags || '',
      Shelves:          row.Shelves || '',
      Notes:            row.Notes || '',
      Review:           row.Review || '',
      Quotes:           row.Quotes || '',
      Favorite:         row.Favorite === true || row.Favorite === 'TRUE',
      CoverEmoji:       row.CoverEmoji || 'BK',
      CoverUrlPrimary:  coverPrimary,
      CoverUrlFallback: coverFallback,
      Gradient1:        row.Gradient1 || '',
      Gradient2:        row.Gradient2 || '',
      ISBN:             isbnNorm,
      OLID:             row.OLID || '',
      AuthorKey:        row.AuthorKey || ''
    };
  });

  var nytBundle = clientGetNYTBadgesForLibrary();
  var nytFeedBundle = clientGetNYTFeed();

  var goals = _sheetToObjects(chalSheet, CHALLENGE_HEADERS).map(function(row) {
    return {
      GoalId:       row.ChallengeId,
      GoalType:     row.Name,
      Icon:         row.Icon || 'GOAL',
      CurrentValue: Number(row.Current) || 0,
      TargetValue:  Number(row.Target) || 1
    };
  });

  var shelves = _sheetToObjects(shelfSheet, SHELF_HEADERS).map(function(row) {
    return {
      ShelfId:   row.ShelfId,
      ShelfName: row.Name,
      Icon:      row.Icon || ''
    };
  });

  // Profile (single row)
  var profileData = {};
  if (profileSheet && profileSheet.getLastRow() >= 2) {
    var pRow = profileSheet.getRange(2, 1, 1, PROFILE_HEADERS.length).getValues()[0];
    PROFILE_HEADERS.forEach(function(h, i) { profileData[h] = pRow[i]; });
  }

  var settings = {
    Theme:                String(profileData.Theme || 'blossom'),
    ShowSpoilers:         String(profileData.ShowSpoilers || 'false')
  };

  var audiobooks = _sheetToObjects(audioSheet, AUDIOBOOK_HEADERS).map(function(row) {
    return {
      AudiobookId:       row.AudiobookId,
      Title:             row.Title,
      Author:            row.Author,
      Duration:          row.Duration || '',
      CoverEmoji:        row.CoverEmoji || 'AUDIO',
      CoverUrl:          row.CoverUrl || '',
      ChapterCount:      Number(row.ChapterCount) || 0,
      LibrivoxProjectId: row.LibrivoxProjectId || '',
      CurrentChapterIndex: Number(row.CurrentChapterIndex) || 0,
      CurrentTime:       Number(row.CurrentTime) || 0,
      PlaybackSpeed:     Number(row.PlaybackSpeed) || 1,
      TotalListeningMins: Number(row.TotalListeningMins) || 0
    };
  });

  return {
    library:    library,
    goals:      goals,
    shelves:    shelves,
    settings:   settings,
    profile: {
      name:      String(profileData.Name || ''),
      motto:     String(profileData.Motto || 'A focused place to track every book'),
      photoData: String(profileData.PhotoData || '')
    },
    yearlyGoal:   Number(profileData.YearlyGoal) || 50,
    readingOrder: _safeJsonParse(profileData.ReadingOrder, []),
    recentIds:    _safeJsonParse(profileData.RecentIds, []),
    sortBy:       String(profileData.SortBy || 'default'),
    libViewMode:  String(profileData.LibViewMode || 'grid'),
    onboarded:    String(profileData.Onboarded) === 'true' || profileData.Onboarded === true,
    demoCleared:  String(profileData.DemoCleared) === 'true' || profileData.DemoCleared === true,
    selectedFilter: String(profileData.SelectedFilter || 'all'),
    activeShelf:    String(profileData.ActiveShelf || ''),
    challengeBarCollapsed: String(profileData.ChallengeBarCollapsed) === 'true' || profileData.ChallengeBarCollapsed === true,
    libToolsOpen: String(profileData.LibToolsOpen) === 'true' || profileData.LibToolsOpen === true,
    libraryName:  String(profileData.LibraryName || ''),
    customQuotes: _safeJsonParse(profileData.CustomQuotes, []),
    coversEnabled: !(String(profileData.CoversEnabled) === 'false' || profileData.CoversEnabled === false),
    tutorialCompleted: String(profileData.TutorialCompleted) === 'true' || profileData.TutorialCompleted === true,
    lastAudioId: String(profileData.LastAudioId || ''),
    totalListeningMins: Number(profileData.TotalListeningMins) || 0,
    audiobooks:   audiobooks,
    nytBadges:    nytBundle.byBookId || {},
    nytBadgesByIsbn: nytBundle.byIsbn || {},
    nytFeed:      nytFeedBundle.lists || [],
    nytCacheDate: nytFeedBundle.updatedAt || PropertiesService.getScriptProperties().getProperty('NYT_CACHE_DATE') || ''
  };
}

function _safeJsonParse(str, fallback) {
  try { var parsed = JSON.parse(str); return Array.isArray(parsed) ? parsed : fallback; }
  catch(e) { return fallback; }
}

function _normalizeIsbn(isbn) {
  return String(isbn || '').toUpperCase().replace(/[^0-9X]/g, '');
}

function _humanizeListName(name) {
  return String(name || '')
    .replace(/-/g, ' ')
    .replace(/\b\w/g, function(c) { return c.toUpperCase(); })
    .trim();
}

/** Add a book */
function clientAddBook(payload) {
  try {
    var sheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
    var bookId = _uuid();
    var row = _bookPayloadToRow(bookId, payload);
    sheet.appendRow(row);
    _writeCoverFormula(sheet, sheet.getLastRow());
    return { BookId: bookId };
  } catch(e) {
    return { error: e.message };
  }
}

/** Update a book (partial payload) */
function clientUpdateBook(bookId, updates) {
  try {
    if (!_validateId(bookId)) return { error: 'Invalid book ID.' };
    if (!updates || typeof updates !== 'object') return { error: 'Updates object required.' };
    var sheet = _ss().getSheetByName(SHEET_LIBRARY);
    if (!sheet) return { error: 'Library sheet not found' };
    var rowIdx = _findRowByCol(sheet, 0, bookId);
    if (rowIdx < 0) return { error: 'Book not found' };

    var dataRow = sheet.getRange(rowIdx, 1, 1, LIBRARY_HEADERS.length).getValues()[0];

  // Map update keys to column indices
  var keyMap = {
    'Title':       'Title',
    'Author':      'Author',
    'Status':      'Status',
    'Rating':      'Rating',
    'Pages':       'Pages',
    'PageCount':   'Pages',
    'Genre':       'Genre',
    'Genres':      'Genre',
    'DateAdded':   'DateAdded',
    'DateStarted': 'DateStarted',
    'DateFinished':'DateFinished',
    'CurrentPage': 'CurrentPage',
    'Series':      'Series',
    'SeriesOrder': 'SeriesNumber',
    'SeriesNumber':'SeriesNumber',
    'TbrPriority': 'TbrPriority',
    'Format':      'Format',
    'Source':      'Source',
    'SpiceLevel':  'SpiceLevel',
    'Tags':        'Tags',
    'Moods':       'Tags',
    'Shelves':     'Shelves',
    'Notes':       'Notes',
    'Review':      'Review',
    'Quotes':      'Quotes',
    'Favorite':    'Favorite',
    'CoverEmoji':  'CoverEmoji',
    'CoverUrl':    'CoverUrl',
    'CoverUrlPrimary': 'CoverUrl',
    'Gradient1':   'Gradient1',
    'Gradient2':   'Gradient2',
    'ISBN':        'ISBN',
    'OLID':        'OLID',
    'AuthorKey':   'AuthorKey'
  };

  Object.keys(updates).forEach(function(k) {
    var colName = keyMap[k];
    if (!colName) return;
    var colIdx = LIBRARY_HEADERS.indexOf(colName);
    if (colIdx < 0) return;
    var val = updates[k];
    // Normalize status to sheet format
    if (colName === 'Status') val = _uiStatusToSheet(val);
    dataRow[colIdx] = val;
  });

  sheet.getRange(rowIdx, 1, 1, LIBRARY_HEADERS.length).setValues([dataRow]);
  _writeCoverFormula(sheet, rowIdx);
  return { success: true };
  } catch(e) { return { error: e.message }; }
}

/** Delete a book */
function clientDeleteBook(bookId) {
  try {
    if (!_validateId(bookId)) return;
    var sheet = _ss().getSheetByName(SHEET_LIBRARY);
    if (!sheet) return;
    var rowIdx = _findRowByCol(sheet, 0, bookId);
    if (rowIdx >= 2) sheet.deleteRow(rowIdx);
  } catch(e) { return { error: e.message }; }
}

/** Bulk import CSV rows — accepts either raw Goodreads CSV or structured payload rows */
function clientImportGoodreadsCSV(rows) {
  try {
  if (!rows || rows.length < 2) return { imported: 0 };
  var sheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
  var headers = rows[0];

  function findCol(name) {
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).toLowerCase().trim() === name.toLowerCase()) return i;
    }
    return -1;
  }

  var newRows = [];
  for (var r = 1; r < rows.length; r++) {
    var d = rows[r];
    var title  = d[findCol('Title')]  || d[0] || '';
    var author = d[findCol('Author')] || d[1] || '';
    if (!title) continue;

    // Build payload from whichever columns are available
    var payload = {
      Title: title,
      Author: author,
      Status: d[findCol('Status')] || (findCol('Exclusive Shelf') > -1 ? d[findCol('Exclusive Shelf')] : '') || 'Want to Read',
      Rating: Number(d[findCol('Rating')] || d[findCol('My Rating')] || 0),
      PageCount: Number(d[findCol('PageCount')] || d[findCol('Pages')] || d[findCol('Number of Pages')] || 0),
      Genre: d[findCol('Genre')] || d[findCol('Genres')] || '',
      DateAdded: d[findCol('DateAdded')] || d[findCol('Date Read')] || new Date().toISOString().slice(0,10),
      DateStarted: d[findCol('DateStarted')] || '',
      DateFinished: d[findCol('DateFinished')] || '',
      CurrentPage: Number(d[findCol('CurrentPage')] || 0),
      Series: d[findCol('Series')] || '',
      SeriesNumber: d[findCol('SeriesNumber')] || '',
      TbrPriority: Number(d[findCol('TbrPriority')] || 0),
      Format: d[findCol('Format')] || '',
      Source: d[findCol('Source')] || '',
      SpiceLevel: Number(d[findCol('SpiceLevel')] || 0),
      Tags: d[findCol('Tags')] || d[findCol('Moods')] || '',
      Shelves: d[findCol('Shelves')] || '',
      Notes: d[findCol('Notes')] || '',
      Review: d[findCol('Review')] || '',
      Quotes: d[findCol('Quotes')] || '',
      Favorite: d[findCol('Favorite')] === true || d[findCol('Favorite')] === 'true',
      CoverEmoji: d[findCol('CoverEmoji')] || '',
      CoverUrl: d[findCol('CoverUrl')] || '',
      Gradient1: d[findCol('Gradient1')] || '',
      Gradient2: d[findCol('Gradient2')] || '',
      ISBN: d[findCol('ISBN')] || '',
      OLID: d[findCol('OLID')] || '',
      AuthorKey: d[findCol('AuthorKey')] || ''
    };

    newRows.push(_bookPayloadToRow(_uuid(), payload));
  }

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, LIBRARY_HEADERS.length).setValues(newRows);
  }
  return { imported: newRows.length };
  } catch(e) { return { error: e.message, imported: 0 }; }
}

function _bookPayloadToRow(bookId, p) {
  function _cap(v, n) { return String(v || '').slice(0, n); }
  return [
    bookId,
    '', // Cover — display-only IMAGE formula is written by _initLibrarySheet
    _cap(p.Title, 500),
    _cap(p.Author, 300),
    p.Status ? _uiStatusToSheet(p.Status) : 'Want to Read',
    Number(p.Rating) || 0,
    Number(p.PageCount || p.Pages) || 0,
    _cap(p.Genres || p.Genre, 200),
    p.DateAdded || new Date().toISOString().slice(0,10),
    p.DateStarted || '',
    p.DateFinished || '',
    Number(p.CurrentPage) || 0,
    _cap(p.Series, 200),
    p.SeriesOrder || p.SeriesNumber || '',
    Number(p.TbrPriority) || 0,
    _cap(p.Format, 100),
    _cap(p.Source, 100),
    Number(p.SpiceLevel) || 0,
    _cap(p.Moods || p.Tags, 500),
    _cap(p.Shelves, 500),
    _cap(p.Notes, 5000),
    _cap(p.Review, 5000),
    _cap(p.Quotes, 5000),
    p.Favorite === true || p.Favorite === 'true',
    p.CoverEmoji || 'BK',
    _cap(p.CoverUrlPrimary || p.CoverUrl, 2000),
    _cap(p.Gradient1, 50),
    _cap(p.Gradient2, 50),
    _cap(p.ISBN, 20),
    _cap(p.OLID, 50),
    _cap(p.AuthorKey, 100)
  ];
}

// ── Shelves ─────────────────────────────────────────────────────────────
function clientAddShelf(name, icon) {
  try {
    var n = String(name || '').trim().slice(0, 200);
    if (!n) return { error: 'Shelf name is required.' };
    var sheet = _getOrCreateSheet(SHEET_SHELVES, SHELF_HEADERS);
    var shelfId = _uuid();
    sheet.appendRow([shelfId, n, String(icon || '').slice(0, 50)]);
    return { ShelfId: shelfId };
  } catch(e) { _log('ERROR', 'clientAddShelf', e.message); return { error: e.message }; }
}

function clientDeleteShelf(shelfId) {
  try {
    if (!_validateId(shelfId)) return;
    var sheet = _ss().getSheetByName(SHEET_SHELVES);
    if (!sheet) return;
    var rowIdx = _findRowByCol(sheet, 0, shelfId);
    if (rowIdx >= 2) sheet.deleteRow(rowIdx);
  } catch(e) { _log('ERROR', 'clientDeleteShelf', e.message); }
}

function clientRenameShelf(shelfId, newName) {
  try {
    if (!_validateId(shelfId)) return;
    var sheet = _ss().getSheetByName(SHEET_SHELVES);
    if (!sheet) return;
    var rowIdx = _findRowByCol(sheet, 0, shelfId);
    if (rowIdx >= 2) sheet.getRange(rowIdx, 2).setValue(String(newName || '').slice(0, 200));
  } catch(e) { _log('ERROR', 'clientRenameShelf', e.message); }
}

function clientUpdateShelf(shelfId, updates) {
  try {
    if (!updates || !_validateId(shelfId)) return;
    var sheet = _ss().getSheetByName(SHEET_SHELVES);
    if (!sheet) return;
    var rowIdx = _findRowByCol(sheet, 0, shelfId);
    if (rowIdx < 2) return;
    var newName = updates.Name !== undefined ? updates.Name : updates.name;
    if (newName !== undefined) sheet.getRange(rowIdx, 2).setValue(String(newName).slice(0, 200));
    var newIcon = updates.Icon !== undefined ? updates.Icon : updates.icon;
    if (newIcon !== undefined) sheet.getRange(rowIdx, 3).setValue(String(newIcon).slice(0, 50));
  } catch(e) { _log('ERROR', 'clientUpdateShelf', e.message); }
}

// ── Challenges / Goals ──────────────────────────────────────────────────
function clientAddChallenge(payload) {
  try {
    if (!payload) return { error: 'Payload required.' };
    var sheet = _getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
    var id = _uuid();
    sheet.appendRow([id, String(payload.name || 'New Challenge').slice(0, 200), String(payload.icon || 'GOAL').slice(0, 50), Number(payload.current) || 0, Number(payload.target) || 10]);
    return { ChallengeId: id };
  } catch(e) { _log('ERROR', 'clientAddChallenge', e.message); return { error: e.message }; }
}

function clientUpdateChallenge(challengeId, updates) {
  try {
    if (!_validateId(challengeId)) return;
    var sheet = _ss().getSheetByName(SHEET_CHALLENGES);
    if (!sheet) return;
    var rowIdx = _findRowByCol(sheet, 0, challengeId);
    if (rowIdx < 2) return;
    var row = sheet.getRange(rowIdx, 1, 1, CHALLENGE_HEADERS.length).getValues()[0];
    if (updates.name !== undefined)    row[1] = String(updates.name).slice(0, 200);
    if (updates.icon !== undefined)    row[2] = String(updates.icon).slice(0, 50);
    if (updates.current !== undefined) row[3] = Number(updates.current);
    if (updates.target !== undefined)  row[4] = Number(updates.target);
    sheet.getRange(rowIdx, 1, 1, CHALLENGE_HEADERS.length).setValues([row]);
  } catch(e) { _log('ERROR', 'clientUpdateChallenge', e.message); }
}

function clientDeleteChallenge(challengeId) {
  try {
    if (!_validateId(challengeId)) return;
    var sheet = _ss().getSheetByName(SHEET_CHALLENGES);
    if (!sheet) return;
    var rowIdx = _findRowByCol(sheet, 0, challengeId);
    if (rowIdx >= 2) sheet.deleteRow(rowIdx);
  } catch(e) { _log('ERROR', 'clientDeleteChallenge', e.message); }
}

function clientSyncChallenges(challengesArray) {
  try {
    // Guard: never wipe all challenges when an empty array is passed (e.g. race condition).
    if (!challengesArray || challengesArray.length === 0) return;
    var sheet = _getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
    // Build the full row set before touching the sheet to minimise the write window.
    var rows = challengesArray.map(function(c) {
      return [
        c._serverChallengeId || _uuid(),
        String(c.name || '').slice(0, 200),
        String(c.icon || 'GOAL').slice(0, 50),
        Number(c.current) || 0,
        Number(c.target) || 1
      ];
    });
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, CHALLENGE_HEADERS.length).clearContent();
    }
    sheet.getRange(2, 1, rows.length, CHALLENGE_HEADERS.length).setValues(rows);
  } catch(e) { _log('ERROR', 'clientSyncChallenges', e.message); }
}

// ── Settings / Preferences ──────────────────────────────────────────────
function clientSetSetting(key, value) {
  try {
    var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
    if (sheet.getLastRow() < 2) {
      _initProfileSheet(_getCurrentTheme());
    }
    var colIdx = PROFILE_HEADERS.indexOf(key);
    if (colIdx < 0) return;
    // Validate theme names before writing — prevents invalid values breaking sheet styling.
    var safeValue = (key === 'Theme') ? _validateTheme(value) : value;
    sheet.getRange(2, colIdx + 1).setValue(safeValue);
    if (key === 'Theme') {
      _reStyleAllSheets(safeValue);
    }
  } catch(e) { _log('ERROR', 'clientSetSetting', e.message); }
}

function clientSetSettings(settingsObj) {
  try {
    if (!settingsObj || typeof settingsObj !== 'object') return;
    Object.keys(settingsObj).forEach(function(k) {
      clientSetSetting(k, settingsObj[k]);
    });
  } catch(e) { _log('ERROR', 'clientSetSettings', e.message); }
}

// ── Profile ─────────────────────────────────────────────────────────────
function clientSaveProfile(profileData) {
  try {
    var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
    if (sheet.getLastRow() < 2) _initProfileSheet(_getCurrentTheme());

    var mapping = {
      'name':     'Name',
      'motto':    'Motto',
      'photoData':'PhotoData'
    };
    Object.keys(mapping).forEach(function(k) {
      if (profileData[k] !== undefined) {
        var colIdx = PROFILE_HEADERS.indexOf(mapping[k]);
        if (colIdx >= 0) {
          // Google Sheets cells have a 50,000 character limit.
          // Truncate base64 photo data and other fields to avoid an unhandled cell-size exception.
          var val;
          if (k === 'photoData') {
            val = String(profileData[k] || '').slice(0, 49000);
          } else {
            val = String(profileData[k] || '').slice(0, 500);
          }
          sheet.getRange(2, colIdx + 1).setValue(val);
        }
      }
    });

    // Strip control characters before using the name as a spreadsheet title.
    var name = String(profileData.name || '').replace(/[\x00-\x1F\x7F]/g, '').trim().slice(0, 100);
    if (name) {
      var possessive = name.endsWith('s') ? name + "'" : name + "'s";
      _ss().rename(possessive + ' Reading Journey');
    } else {
      _ss().rename('My Reading Journey');
    }
  } catch(e) { _log('ERROR', 'clientSaveProfile', e.message); }
}

function clientSaveYearlyGoal(goal) {
  try {
    var n = Number(goal);
    if (!isFinite(n) || n < 1 || n > 10000) return;
    clientSetSetting('YearlyGoal', n);
  } catch(e) { _log('ERROR', 'clientSaveYearlyGoal', e.message); }
}

function clientSaveReadingOrder(orderArray) {
  try {
    if (!Array.isArray(orderArray)) return;
    var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
    if (sheet.getLastRow() < 2) _initProfileSheet(_getCurrentTheme());
    var colIdx = PROFILE_HEADERS.indexOf('ReadingOrder');
    if (colIdx >= 0) sheet.getRange(2, colIdx + 1).setValue(JSON.stringify(orderArray));
  } catch(e) { _log('ERROR', 'clientSaveReadingOrder', e.message); }
}

function clientSaveRecentIds(idsArray) {
  try {
    if (!Array.isArray(idsArray)) return;
    var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
    if (sheet.getLastRow() < 2) _initProfileSheet(_getCurrentTheme());
    var colIdx = PROFILE_HEADERS.indexOf('RecentIds');
    if (colIdx >= 0) sheet.getRange(2, colIdx + 1).setValue(JSON.stringify(idsArray));
  } catch(e) { _log('ERROR', 'clientSaveRecentIds', e.message); }
}

function clientSaveUiPrefs(prefs) {
  try {
    if (!prefs || typeof prefs !== 'object') return;
    if (prefs.sortBy !== undefined)               clientSetSetting('SortBy', prefs.sortBy);
    if (prefs.libViewMode !== undefined)           clientSetSetting('LibViewMode', prefs.libViewMode);
    if (prefs.onboarded !== undefined)             clientSetSetting('Onboarded', prefs.onboarded);
    if (prefs.demoCleared !== undefined)           clientSetSetting('DemoCleared', prefs.demoCleared);
    if (prefs.selectedFilter !== undefined)        clientSetSetting('SelectedFilter', prefs.selectedFilter);
    if (prefs.activeShelf !== undefined)           clientSetSetting('ActiveShelf', prefs.activeShelf);
    if (prefs.challengeBarCollapsed !== undefined) clientSetSetting('ChallengeBarCollapsed', prefs.challengeBarCollapsed);
    if (prefs.libToolsOpen !== undefined)          clientSetSetting('LibToolsOpen', prefs.libToolsOpen);
    if (prefs.libraryName !== undefined)           clientSetSetting('LibraryName', prefs.libraryName);
    // --- Full-sync fields ---
    if (prefs.customQuotes !== undefined) {
      // Arrays are stored as JSON strings to keep the sheet cell scalar & predictable.
      var cq = prefs.customQuotes;
      clientSetSetting('CustomQuotes', typeof cq === 'string' ? cq : JSON.stringify(cq || []));
    }
    if (prefs.coversEnabled !== undefined)         clientSetSetting('CoversEnabled', !!prefs.coversEnabled);
    if (prefs.tutorialCompleted !== undefined)     clientSetSetting('TutorialCompleted', !!prefs.tutorialCompleted);
    if (prefs.lastAudioId !== undefined)           clientSetSetting('LastAudioId', String(prefs.lastAudioId || ''));
    if (prefs.totalListeningMins !== undefined)    clientSetSetting('TotalListeningMins', Number(prefs.totalListeningMins) || 0);
  } catch(e) { _log('ERROR', 'clientSaveUiPrefs', e.message); }
}

// ── Audiobooks ──────────────────────────────────────────────────────────
function clientSaveAudiobook(audioData) {
  try {
  var sheet = _getOrCreateSheet(SHEET_AUDIOBOOKS, AUDIOBOOK_HEADERS);
  var existingRow = _findRowByCol(sheet, 0, audioData.id);
  var row = [
    audioData.id || _uuid(),
    audioData.title || '',
    audioData.author || '',
    audioData.duration || '',
    audioData.cover || 'AUDIO',
    audioData.coverUrl || '',
    Number(audioData.chapterCount) || 0,
    audioData.audiobookId || '',
    Number(audioData.chapterIndex) || 0,
    Number(audioData.currentTime) || 0,
    Number(audioData.speed) || 1,
    Number(audioData.totalListeningMins) || 0
  ];
  if (existingRow > 1) {
    sheet.getRange(existingRow, 1, 1, AUDIOBOOK_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  } catch(e) { return { error: e.message }; }
}

function clientSaveAudioPosition(audioId, chapterIndex, currentTime, speed, totalListeningMins) {
  try {
    if (!_validateId(audioId)) return;
    var sheet = _ss().getSheetByName(SHEET_AUDIOBOOKS);
    if (!sheet) return;
    var rowIdx = _findRowByCol(sheet, 0, String(audioId));
    if (rowIdx < 2) return;
    // Single read-modify-write instead of 4 separate setValues round-trips.
    var row = sheet.getRange(rowIdx, 1, 1, AUDIOBOOK_HEADERS.length).getValues()[0];
    row[AUDIOBOOK_HEADERS.indexOf('CurrentChapterIndex')] = Number(chapterIndex) || 0;
    row[AUDIOBOOK_HEADERS.indexOf('CurrentTime')]         = Number(currentTime) || 0;
    row[AUDIOBOOK_HEADERS.indexOf('PlaybackSpeed')]       = Number(speed) || 1;
    if (totalListeningMins !== undefined && totalListeningMins !== null) {
      row[AUDIOBOOK_HEADERS.indexOf('TotalListeningMins')] = Number(totalListeningMins) || 0;
    }
    sheet.getRange(rowIdx, 1, 1, AUDIOBOOK_HEADERS.length).setValues([row]);
  } catch(e) { _log('ERROR', 'clientSaveAudioPosition', e.message); }
}

// ── Demo Data Management ────────────────────────────────────────────────
function _clearSheetDataRows(sheet, headers) {
  if (!sheet || sheet.getLastRow() < 2) return;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
}

/** Seed demo library, challenges, and shelves for first-run WOW factor */
function _seedDemoData() {
  var libSheet = _ss().getSheetByName(SHEET_LIBRARY);
  // Only seed if library is empty (first run)
  if (libSheet && libSheet.getLastRow() >= 2) return;
  if (!libSheet) libSheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);

  var now = new Date();
  function _monthDate(monthsAgo, day) {
    var d = new Date(now.getFullYear(), now.getMonth() - monthsAgo, day);
    return d.toISOString().slice(0, 10);
  }
  function _weeksAgo(w, dayOff) {
    var d = new Date(now);
    d.setDate(now.getDate() - w * 7 + (dayOff || 0));
    return d.toISOString().slice(0, 10);
  }

  var demoBooks = [
    { t:'Beach Read', a:'Emily Henry', g:'Romance', isbn:'9781984806734', pg:352, r:5, e:'BK', g1:'#F472B6', g2:'#F43F5E', stat:'Finished', da:_weeksAgo(11,2), df:_monthDate(0,2) },
    { t:'Circe', a:'Madeline Miller', g:'Fantasy', isbn:'9780316556347', pg:393, r:5, e:'BK', g1:'#A78BFA', g2:'#8B5CF6', stat:'Finished', da:_weeksAgo(10,0), df:_monthDate(2,14) },
    { t:'The Silent Patient', a:'Alex Michaelides', g:'Thriller', isbn:'9781250301697', pg:325, r:4, e:'BK', g1:'#FB923C', g2:'#F59E0B', stat:'Finished', da:_weeksAgo(9,4), df:_monthDate(1,3) },
    { t:'Atomic Habits', a:'James Clear', g:'Self-Help', isbn:'9780735211292', pg:306, r:5, e:'BK', g1:'#34D399', g2:'#10B981', stat:'Finished', da:_weeksAgo(8,1), df:_monthDate(5,22) },
    { t:'The Song of Achilles', a:'Madeline Miller', g:'Fantasy', isbn:'9780062060624', pg:352, r:5, e:'BK', g1:'#60A5FA', g2:'#6366F1', stat:'Finished', da:_weeksAgo(0,0), df:_monthDate(2,27) },
    { t:'Where the Crawdads Sing', a:'Delia Owens', g:'Mystery', isbn:'9780735224292', pg:368, r:5, e:'BK', g1:'#2DD4BF', g2:'#06B6D4', stat:'Finished', da:_weeksAgo(6,0), df:_monthDate(1,27) },
    { t:'Project Hail Mary', a:'Andy Weir', g:'SciFi', isbn:'9780593135204', pg:476, r:5, e:'BK', g1:'#FB7185', g2:'#EC4899', stat:'Finished', da:_weeksAgo(4,5), df:_monthDate(1,20) },
    { t:'The Guest List', a:'Lucy Foley', g:'Mystery', isbn:'9780062868930', pg:312, r:4, e:'BK', g1:'#E879F9', g2:'#8B5CF6', stat:'Finished', da:_weeksAgo(8,5), df:_monthDate(4,24) },
    { t:'Educated', a:'Tara Westover', g:'Memoir', isbn:'9780399590504', pg:334, r:5, e:'BK', g1:'#FBBF24', g2:'#F97316', stat:'Finished', da:_weeksAgo(3,3), df:_monthDate(1,8) },
    { t:'The Invisible Life of Addie LaRue', a:'V.E. Schwab', g:'Fantasy', isbn:'9780765387561', pg:448, r:5, e:'BK', g1:'#4ADE80', g2:'#22C55E', stat:'Finished', da:_weeksAgo(1,4), df:_monthDate(2,25) },
    { t:'The Vanishing Half', a:'Brit Bennett', g:'Fiction', isbn:'9780525536291', pg:343, r:4, e:'BK', g1:'#818CF8', g2:'#3B82F6', stat:'Finished', da:_weeksAgo(4,1), df:_monthDate(1,14) },
    { t:'Verity', a:'Colleen Hoover', g:'Thriller', isbn:'9781538724736', pg:374, r:5, e:'BK', g1:'#22D3EE', g2:'#14B8A6', stat:'Finished', da:_weeksAgo(7,3), df:_monthDate(4,3) },
    { t:'Book Lovers', a:'Emily Henry', g:'Romance', isbn:'9780593334836', pg:368, r:5, e:'BK', g1:'#EC4899', g2:'#E11D48', stat:'Finished', da:_weeksAgo(3,0), df:_monthDate(2,4) },
    { t:'The Spanish Love Deception', a:'Elena Armas', g:'Romance', isbn:'9781982177010', pg:358, r:4, e:'BK', g1:'#A78BFA', g2:'#9333EA', stat:'Reading', da:_weeksAgo(5,2), ds:'2026-02-10', cp:125 },
    { t:'A Court of Thorns and Roses', a:'Sarah J. Maas', g:'Fantasy', isbn:'9781635575569', pg:419, r:5, e:'BK', g1:'#F97316', g2:'#D97706', stat:'Finished', da:_weeksAgo(6,4), df:_monthDate(3,19) },
    { t:'The Thursday Murder Club', a:'Richard Osman', g:'Mystery', isbn:'9781984880963', pg:369, r:4, e:'BK', g1:'#22C55E', g2:'#10B981', stat:'DNF', da:_weeksAgo(2,5) },
    { t:'The Four Winds', a:'Kristin Hannah', g:'Historical', isbn:'9781250178602', pg:454, r:5, e:'BK', g1:'#3B82F6', g2:'#4F46E5', stat:'Finished', da:_weeksAgo(2,2), df:_monthDate(5,8) },
    { t:'Normal People', a:'Sally Rooney', g:'Romance', isbn:'9781984822185', pg:266, r:4, e:'BK', g1:'#14B8A6', g2:'#06B6D4', stat:'Finished', da:_weeksAgo(9,4), df:_monthDate(5,15) },
    { t:'The House in the Cerulean Sea', a:'TJ Klune', g:'Fantasy', isbn:'9781250217288', pg:396, r:5, e:'BK', g1:'#F43F5E', g2:'#EC4899', stat:'Finished', da:_weeksAgo(6,0), df:_monthDate(4,29) },
    { t:'Malibu Rising', a:'Taylor Jenkins Reid', g:'Fiction', isbn:'9780593158203', pg:369, r:5, e:'BK', g1:'#8B5CF6', g2:'#D946EF', stat:'Finished', da:_weeksAgo(0,4), df:_monthDate(3,5) },
    { t:'The Love Hypothesis', a:'Ali Hazelwood', g:'Romance', isbn:'9780593336823', pg:357, r:4, e:'BK', g1:'#F59E0B', g2:'#EA580C', stat:'Reading', da:_weeksAgo(2,2), ds:'2026-02-01', cp:89 },
    { t:'Daisy Jones & The Six', a:'Taylor Jenkins Reid', g:'Fiction', isbn:'9781524798628', pg:368, r:5, e:'BK', g1:'#10B981', g2:'#16A34A', stat:'Finished', da:_weeksAgo(1,1), df:_monthDate(3,12) },
    { t:'The Atlas Six', a:'Olivie Blake', g:'Fantasy', isbn:'9781250854513', pg:374, r:4, e:'BK', g1:'#4F46E5', g2:'#2563EB', stat:'Finished', da:_weeksAgo(3,3), df:_monthDate(3,26) },
    { t:'Red, White & Royal Blue', a:'Casey McQuiston', g:'Romance', isbn:'9781250316776', pg:352, r:5, e:'BK', g1:'#06B6D4', g2:'#0D9488', stat:'Finished', da:_weeksAgo(0,2), df:_monthDate(2,20) },
    { t:'It Ends With Us', a:'Colleen Hoover', g:'Romance', isbn:'9781501110375', pg:376, r:5, e:'BK', g1:'#A78BFA', g2:'#8B5CF6', stat:'Reading', da:_weeksAgo(0,0), ds:'2026-03-22', cp:169 },
    { t:'The Midnight Library', a:'Matt Haig', g:'Fiction', isbn:'9780525559474', pg:304, r:4, e:'BK', g1:'#60A5FA', g2:'#22D3EE', stat:'Reading', da:_weeksAgo(0,2), ds:'2026-03-01', cp:249 }
  ];

  var bookIds = [];
  var readingIds = [];
  var rows = demoBooks.map(function(b) {
    var bookId = _uuid();
    bookIds.push(bookId);
    if (b.stat === 'Reading') readingIds.push(bookId);
    // LIBRARY_HEADERS order: BookId, Cover, Title, Author, Status, Rating, Pages, Genre,
    //   DateAdded, DateStarted, DateFinished, CurrentPage, Series, SeriesNumber,
    //   TbrPriority, Format, Source, SpiceLevel, Tags, Shelves, Notes, Review, Quotes,
    //   Favorite, CoverEmoji, CoverUrl, Gradient1, Gradient2, ISBN, OLID, AuthorKey
    // Cover (index 1) is left blank — IMAGE formula is stamped after setValues.
    return [
      bookId, '',  // BookId, Cover (formula added below)
      b.t, b.a, b.stat, b.r, b.pg, b.g,
      b.da || '', b.ds || '', b.df || '', b.cp || 0,
      '', '', 0, '', '', 0,
      '', '', '', '', '',
      b.r === 5, 'BK',
      'https://covers.openlibrary.org/b/isbn/' + b.isbn + '-L.jpg',
      b.g1, b.g2, b.isbn, '', ''
    ];
  });

  if (rows.length > 0) {
    libSheet.getRange(2, 1, rows.length, LIBRARY_HEADERS.length).setValues(rows);
    // Stamp cover IMAGE() formulas for every seeded row
    for (var si = 0; si < rows.length; si++) {
      _writeCoverFormula(libSheet, 2 + si);
    }
  }

  // Seed challenges
  var chalSheet = _getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
  if (chalSheet.getLastRow() < 2) {
    chalSheet.getRange(2, 1, 3, CHALLENGE_HEADERS.length).setValues([
      [_uuid(), '50 Books Challenge', 'Books', 42, 50],
      [_uuid(), 'Read 30 Min Daily', 'Daily', 27, 30],
      [_uuid(), 'Try 5 New Authors', 'Authors', 4, 5]
    ]);
  }

  // Seed shelves
  var shelfSheet = _getOrCreateSheet(SHEET_SHELVES, SHELF_HEADERS);
  if (shelfSheet.getLastRow() < 2) {
    shelfSheet.getRange(2, 1, 3, SHELF_HEADERS.length).setValues([
      [_uuid(), 'Book Club', 'Club'],
      [_uuid(), 'Comfort Reads', 'Calm'],
      [_uuid(), 'Summer TBR', 'Seasonal']
    ]);
  }

  // Set reading order in profile
  var profileSheet = _ss().getSheetByName(SHEET_PROFILE);
  if (profileSheet && profileSheet.getLastRow() >= 2 && readingIds.length > 0) {
    var roCol = PROFILE_HEADERS.indexOf('ReadingOrder') + 1;
    if (roCol > 0) profileSheet.getRange(2, roCol).setValue(JSON.stringify(readingIds));
  }
}

function clientClearDemoData() {
  try {
    var ss = _ss();
    _clearSheetDataRows(ss.getSheetByName(SHEET_LIBRARY), LIBRARY_HEADERS);
    _clearSheetDataRows(ss.getSheetByName(SHEET_CHALLENGES), CHALLENGE_HEADERS);
    _clearSheetDataRows(ss.getSheetByName(SHEET_SHELVES), SHELF_HEADERS);
    _clearSheetDataRows(ss.getSheetByName(SHEET_AUDIOBOOKS), AUDIOBOOK_HEADERS);

    var profileSheet = ss.getSheetByName(SHEET_PROFILE);
    if (profileSheet && profileSheet.getLastRow() >= 2) {
      var profileResets = {
        ReadingOrder: '[]',
        RecentIds: '[]',
        SelectedFilter: 'all',
        ActiveShelf: '',
        SortBy: 'default',
        LibViewMode: 'grid',
        ChallengeBarCollapsed: false,
        LibToolsOpen: false
      };
      // Single read-modify-write — avoids 8 separate Sheets API calls.
      var profileRow = profileSheet.getRange(2, 1, 1, PROFILE_HEADERS.length).getValues()[0];
      Object.keys(profileResets).forEach(function(key) {
        var colIdx = PROFILE_HEADERS.indexOf(key);
        if (colIdx >= 0) profileRow[colIdx] = profileResets[key];
      });
      profileSheet.getRange(2, 1, 1, PROFILE_HEADERS.length).setValues([profileRow]);
    }

    return { cleared: true };
  } catch(e) { _log('ERROR', 'clientClearDemoData', e.message); return { error: e.message }; }
}

// ── Search proxies (LibriVox & Open Library are called client-side) ─────
// These are no-ops if you want search done entirely client-side via fetch().
// Include them only if you prefer to proxy searches through Apps Script
// to avoid CORS issues in the iframe.

function clientSearchBooks(query) {
  if (!query) return [];
  var url = 'https://openlibrary.org/search.json?q=' + encodeURIComponent(query) +
    '&limit=15&fields=key,title,author_name,author_key,first_publish_year,isbn,cover_i,number_of_pages_median,subject';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    return (json.docs || []).map(function(doc) {
      var isbnCandidates = [];
      if (doc.isbn && doc.isbn.length) {
        for (var i = 0; i < doc.isbn.length; i++) {
          var candidate = _normalizeIsbn(doc.isbn[i]);
          if (!candidate) continue;
          if (isbnCandidates.indexOf(candidate) === -1) isbnCandidates.push(candidate);
        }
      }
      isbnCandidates = isbnCandidates.slice(0, 10);
      var isbn = '';
      for (var j = 0; j < isbnCandidates.length; j++) {
        if (isbnCandidates[j].length === 13) { isbn = isbnCandidates[j]; break; }
      }
      if (!isbn && isbnCandidates.length) isbn = isbnCandidates[0];
      var coverId = doc.cover_i || '';
      return {
        title: doc.title || '',
        author: (doc.author_name || [])[0] || '',
        authorKey: (doc.author_key || [])[0] || '',
        year: doc.first_publish_year || '',
        isbn: isbn,
        isbnCandidates: isbnCandidates,
        isbns: isbnCandidates,
        olid: doc.key ? doc.key.replace('/works/', '') : '',
        coverId: coverId,
        coverUrlPrimary: isbn ? ('https://covers.openlibrary.org/b/isbn/' + isbn + '-L.jpg') : (coverId ? ('https://covers.openlibrary.org/b/id/' + coverId + '-L.jpg') : ''),
        coverUrlFallback: coverId ? ('https://covers.openlibrary.org/b/id/' + coverId + '-L.jpg') : '',
        pageCount: doc.number_of_pages_median || '',
        subjects: (doc.subject || []).slice(0, 5)
      };
    });
  } catch(e) {
    Logger.log('Search error: ' + e);
    return [];
  }
}

function clientSearchAudiobook(query) {
  if (!query) return [];
  var url = 'https://librivox.org/api/feed/audiobooks?title=' + encodeURIComponent(query) + '&format=json&limit=10&extended=1';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    var json = JSON.parse(resp.getContentText());
    return (json.books || []).map(function(b) {
      var a = (b.authors || [{}])[0] || {};
      return {
        audiobookId: b.id,
        title: b.title || '',
        author: ((a.first_name || '') + ' ' + (a.last_name || '')).trim() || 'Unknown Author',
        totalTime: b.totaltime || '',
        numSections: Number(b.num_sections) || 0,
        coverUrl: b.url_image || ''
      };
    });
  } catch(e) {
    Logger.log('Audio search error: ' + e);
    return [];
  }
}

function clientSearchPodcastDiscussions(query) {
  if (!query) return [];
  // Store credentials in Apps Script → Project Settings → Script Properties:
  //   PODCAST_INDEX_API_KEY   = your key from podcastindex.org
  //   PODCAST_INDEX_API_SECRET = your secret from podcastindex.org
  var props = PropertiesService.getScriptProperties();
  var apiKey    = props.getProperty('PODCAST_INDEX_API_KEY');
  var apiSecret = props.getProperty('PODCAST_INDEX_API_SECRET');
  if (!apiKey || !apiSecret) return [];
  var ts = String(Math.floor(Date.now() / 1000));
  var authBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    apiKey + apiSecret + ts,
    Utilities.Charset.UTF_8
  );
  var auth = authBytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
  var url = 'https://api.podcastindex.org/api/1.0/search/byterm?q=' +
    encodeURIComponent(query + ' book') + '&max=10';
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'X-Auth-Key': apiKey,
        'X-Auth-Date': ts,
        'Authorization': auth,
        'User-Agent': 'PageVault/1.0'
      }
    });
    if (resp.getResponseCode() !== 200) return [];
    var json = JSON.parse(resp.getContentText());
    return (json.feeds || []).slice(0, 10).map(function(feed) {
      return {
        id: feed.id || '',
        title: feed.title || '',
        author: feed.author || '',
        description: feed.description ? String(feed.description).slice(0, 140) : '',
        coverUrl: feed.image || feed.artwork || '',
        podcastLink: feed.link || feed.url || ''
      };
    });
  } catch(e) {
    Logger.log('Podcast discussion search error: ' + e);
    return [];
  }
}

function clientGetAudiobookChapters(projectId) {
  if (!projectId) return [];
  var url = 'https://librivox.org/api/feed/audiotracks?project_id=' + projectId + '&format=json';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var json = JSON.parse(resp.getContentText());
    return (json.sections || []).map(function(s, i) {
      return {
        chapterIndex: i,
        title: s.title || ('Chapter ' + (i + 1)),
        duration: s.playtime || '',
        url: s.listen_url || '',
        reader: (s.readers || [])[0] ? s.readers[0].display_name : ''
      };
    });
  } catch(e) {
    Logger.log('Chapter fetch error: ' + e);
    return [];
  }
}

// ── Menu trigger for manual re-init ─────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(_buildJourneyTitle())
    .addItem('Open Web App', '_openWebApp')
    .addItem('Refresh Dashboard Styling', '_reStyleCurrentTheme')
    .addSeparator()
    .addItem('Refresh NYT Bestseller Cache', 'clientRefreshNYTCache')
    .addItem('Install Weekly NYT Trigger', 'installNYTWeeklyTrigger')
    .addItem('Install Sync Trigger', 'installSyncTrigger')
    .addSeparator()
    .addItem('Re-initialize Sheets', 'initializeSheets')
    .addToUi();
}

/** Open the deployed web app URL */
function _openWebApp() {
  var url = ScriptApp.getService().getUrl();
  // Use JSON.stringify to safely encode the URL into the JS string literal,
  // avoiding any concatenation-based injection if the URL ever contains quotes.
  var safeSrc = '<script>window.open(' + JSON.stringify(url) + ');google.script.host.close();\x3c/script>';
  var html = HtmlService.createHtmlOutput(safeSrc).setWidth(1).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

function _reStyleCurrentTheme() {
  _reStyleAllSheets(_getCurrentTheme());
  SpreadsheetApp.getUi().alert('Sheet styles updated to match your current theme.');
}

// =====================================================================
//  OPEN LIBRARY — BOOK DETAIL, AUTHOR, FREE EBOOK CHECK
// =====================================================================

/**
 * Fetch full book detail from Open Library Works API.
 * @param {string} olid  e.g. "OL27448W"
 */
function clientGetBookDetails(olid) {
  if (!olid) return null;
  var workKey = olid.startsWith('/works/') ? olid : '/works/' + olid;
  var url = 'https://openlibrary.org' + workKey + '.json';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var data = JSON.parse(resp.getContentText());

    // Description may be a string or {type, value} object
    var desc = '';
    if (data.description) {
      desc = (typeof data.description === 'object') ? (data.description.value || '') : String(data.description);
    }

    // Subjects array — cap at 10
    var subjects = (data.subjects || []).slice(0, 10);

    // Cover IDs
    var coverId = (data.covers || [])[0] || null;
    var coverUrl = coverId ? 'https://covers.openlibrary.org/b/id/' + coverId + '-L.jpg' : null;

    // Author key from first author entry
    var authorKey = null;
    if (data.authors && data.authors[0] && data.authors[0].author) {
      authorKey = data.authors[0].author.key || null; // e.g. "/authors/OL23919A"
    }

    return {
      olid:        olid,
      description: desc,
      subjects:    subjects,
      coverUrl:    coverUrl,
      coverId:     coverId,
      authorKey:   authorKey,
      firstPublish: data.first_publish_date || ''
    };
  } catch(e) {
    Logger.log('clientGetBookDetails error: ' + e);
    return null;
  }
}

/**
 * Fetch author bio and photo from Open Library Authors API.
 * @param {string} authorKey  e.g. "/authors/OL23919A" or "OL23919A"
 */
function clientGetAuthorDetails(authorKey) {
  if (!authorKey) return null;
  var key = authorKey.startsWith('/authors/') ? authorKey : '/authors/' + authorKey;
  var olid = key.replace('/authors/', '');
  var url = 'https://openlibrary.org' + key + '.json';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var data = JSON.parse(resp.getContentText());

    var bio = '';
    if (data.bio) {
      bio = (typeof data.bio === 'object') ? (data.bio.value || '') : String(data.bio);
    }

    var photoId = (data.photos || [])[0] || null;
    var photoUrl = photoId ? 'https://covers.openlibrary.org/a/olid/' + olid + '-M.jpg' : null;

    return {
      authorKey:  key,
      name:       data.name || '',
      bio:        bio,
      birthDate:  data.birth_date || '',
      deathDate:  data.death_date || '',
      photoUrl:   photoUrl
    };
  } catch(e) {
    Logger.log('clientGetAuthorDetails error: ' + e);
    return null;
  }
}

/**
 * Check if a book is freely readable via Internet Archive / Open Library.
 * Returns { available: bool, readUrl: string, previewLevel: string }
 * @param {string} isbn
 */
function clientCheckFreeEbook(isbn) {
  if (!isbn) return { available: false };
  var url = 'https://openlibrary.org/api/books?bibkeys=ISBN:' + encodeURIComponent(isbn) +
            '&jscmd=viewapi&format=json';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return { available: false };
    var data = JSON.parse(resp.getContentText());
    var entry = data['ISBN:' + isbn];
    if (!entry) return { available: false };
    var preview = entry.preview || 'noview';
    return {
      available:    preview === 'full' || preview === 'limited',
      previewLevel: preview,
      readUrl:      entry.read_url || entry.info_url || '',
      thumbnail:    entry.thumbnail_url || ''
    };
  } catch(e) {
    Logger.log('clientCheckFreeEbook error: ' + e);
    return { available: false };
  }
}

// =====================================================================
//  INTERNET ARCHIVE — AUDIO SEARCH
// =====================================================================

/**
 * Search Internet Archive for free audio recordings of a book.
 * Returns up to 8 results with identifiers and stream info.
 * @param {string} query  title (+ optionally author)
 */
function clientSearchArchiveAudio(query) {
  if (!query) return [];
  var url = 'https://archive.org/advancedsearch.php' +
    '?q=' + encodeURIComponent(query + ' mediatype:audio') +
    '&fl[]=identifier,title,creator,description,runtime' +
    '&rows=8&output=json';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    var data = JSON.parse(resp.getContentText());
    return ((data.response || {}).docs || []).map(function(doc) {
      return {
        identifier:  doc.identifier || '',
        title:       doc.title || '',
        author:      (Array.isArray(doc.creator) ? doc.creator[0] : doc.creator) || '',
        description: (Array.isArray(doc.description) ? doc.description[0] : doc.description) || '',
        runtime:     doc.runtime || '',
        streamBase:  'https://archive.org/download/' + (doc.identifier || '')
      };
    });
  } catch(e) {
    Logger.log('clientSearchArchiveAudio error: ' + e);
    return [];
  }
}

/**
 * Get the file listing for an Internet Archive audio item so we can build
 * per-chapter stream URLs.
 * @param {string} identifier  Archive.org item identifier
 */
function clientGetArchiveAudioFiles(identifier) {
  if (!identifier) return [];
  var url = 'https://archive.org/metadata/' + encodeURIComponent(identifier);
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    var data = JSON.parse(resp.getContentText());
    var files = (data.files || [])
      .filter(function(f) { return /\.(mp3|ogg|flac|opus)$/i.test(f.name); })
      .sort(function(a, b) { return String(a.name).localeCompare(b.name); });
    return files.map(function(f, i) {
      return {
        chapterIndex: i,
        title:    f.title || f.name,
        duration: f.length || '',
        url:      'https://archive.org/download/' + identifier + '/' + f.name
      };
    });
  } catch(e) {
    Logger.log('clientGetArchiveAudioFiles error: ' + e);
    return [];
  }
}

// =====================================================================
//  NYT BOOKS API — SERVER-SIDE CACHE
// =====================================================================
//  Store your free NYT API key in Script Properties:
//    Project Settings → Script Properties → NYT_API_KEY = <your key>
//  Free key: developer.nytimes.com — email only, no credit card.
//
//  Cache lives in PropertiesService.getScriptProperties() under key
//  "NYT_CACHE" as a JSON-stringified object:
//    { isbn: { rank, weeksOn, list, title, author }, ... }
//
//  A time-based trigger fires clientRefreshNYTCache() once a week.
// =====================================================================

var NYT_LISTS = [
  'hardcover-fiction',
  'hardcover-nonfiction',
  'paperback-nonfiction',
  'young-adult-hardcover',
  'childrens-middle-grade-hardcover',
  'graphic-books-and-manga',
  'science',
  'business-books'
];

function _getNytApiKey() {
  // Store your NYT Books API key in Apps Script → Project Settings → Script Properties
  // as key: NYT_API_KEY.  Get a free key at https://developer.nytimes.com/
  return PropertiesService.getScriptProperties().getProperty('NYT_API_KEY') || null;
}

/**
 * Called by the weekly time-based trigger.
 * Fetches all major NYT lists and caches ISBN→rank data.
 */
function clientRefreshNYTCache() {
  var props = PropertiesService.getScriptProperties();
  var apiKey = _getNytApiKey();
  if (!apiKey) {
    Logger.log('NYT_API_KEY not set in Script Properties. Add it under Project Settings → Script Properties.');
    return;
  }

  var cache = {};
  var currentFeedLists = [];
  NYT_LISTS.forEach(function(listName) {
    try {
      var url = 'https://api.nytimes.com/svc/books/v3/lists/current/' +
                encodeURIComponent(listName) + '.json?api-key=' + apiKey;
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) return;
      var data = JSON.parse(resp.getContentText());
      var results = data.results || {};
      var books = results.books || [];
      currentFeedLists.push({
        list: listName,
        listDisplay: results.display_name || _humanizeListName(listName),
        updatedAt: results.published_date || '',
        books: books.slice(0, 8).map(function(b) {
          return {
            rank: Number(b.rank) || 0,
            weeksOn: Number(b.weeks_on_list) || 0,
            title: b.title || '',
            author: b.author || '',
            description: b.description || '',
            isbn13: b.primary_isbn13 || '',
            isbn10: b.primary_isbn10 || '',
            bookImage: b.book_image || ''
          };
        })
      });
      books.forEach(function(b) {
        var isbns = [b.primary_isbn13, b.primary_isbn10].filter(Boolean);
        isbns.forEach(function(isbn) {
          var norm = _normalizeIsbn(isbn);
          if (!norm) return;
          cache[norm] = {
            rank:     b.rank,
            weeksOn:  b.weeks_on_list,
            list:     listName,
            listDisplay: _humanizeListName(listName),
            title:    b.title,
            author:   b.author
          };
        });
      });
      Utilities.sleep(6500); // NYT free tier: max 10 req/min → must wait ≥6 s between requests
    } catch(e) {
      Logger.log('NYT fetch error for ' + listName + ': ' + e);
    }
  });

  // Enrich cache with historical best-seller metadata for library ISBNs not in current lists.
  // This surfaces older books that are no longer in "current" NYT lists.
  var libraryIsbns = _getLibraryIsbnsForNyt().filter(function(isbn) { return !cache[isbn]; });
  var maxHistoryLookups = 20;
  for (var i = 0; i < libraryIsbns.length && i < maxHistoryLookups; i++) {
    var isbn = libraryIsbns[i];
    try {
      var historyUrl = 'https://api.nytimes.com/svc/books/v3/lists/best-sellers/history.json?isbn=' + encodeURIComponent(isbn) + '&api-key=' + apiKey;
      var historyResp = UrlFetchApp.fetch(historyUrl, { muteHttpExceptions: true });
      if (historyResp.getResponseCode() !== 200) { Utilities.sleep(1200); continue; }
      var historyData = JSON.parse(historyResp.getContentText());
      var entries = (historyData.results || []);
      if (!entries.length) { Utilities.sleep(1200); continue; }

      var entry = entries[0];
      var ranksHistory = Array.isArray(entry.ranks_history) ? entry.ranks_history : [];
      var bestRank = 0;
      var bestWeeks = Number(entry.weeks_on_list) || 0;
      var bestList = '';
      var bestListDisplay = '';

      ranksHistory.forEach(function(rh) {
        var rankVal = Number(rh.rank) || 0;
        var weeksVal = Number(rh.weeks_on_list) || 0;
        if (!bestRank || (rankVal > 0 && rankVal < bestRank)) bestRank = rankVal;
        if (weeksVal > bestWeeks) bestWeeks = weeksVal;
        if (!bestList && rh.list_name_encoded) bestList = rh.list_name_encoded;
        if (!bestListDisplay && rh.display_name) bestListDisplay = rh.display_name;
      });

      if (!bestList) bestList = entry.list_name_encoded || 'best-sellers-history';
      if (!bestListDisplay) bestListDisplay = entry.list_name || _humanizeListName(bestList);

      cache[isbn] = {
        rank: bestRank,
        weeksOn: bestWeeks,
        list: bestList,
        listDisplay: bestListDisplay,
        title: entry.title || '',
        author: entry.author || ''
      };
    } catch (e) {
      Logger.log('NYT history fetch error for ' + isbn + ': ' + e);
    }
    Utilities.sleep(1200);
  }

  props.setProperty('NYT_CACHE', JSON.stringify(cache));
  props.setProperty('NYT_CACHE_DATE', new Date().toISOString().slice(0, 10));
  props.setProperty('NYT_FEED_CURRENT', JSON.stringify({
    updatedAt: new Date().toISOString().slice(0, 10),
    lists: currentFeedLists
  }));
  Logger.log('NYT cache refreshed: ' + Object.keys(cache).length + ' ISBNs cached.');
}

function _getLibraryIsbnsForNyt() {
  var sheet = _ss().getSheetByName(SHEET_LIBRARY);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var isbnCol = LIBRARY_HEADERS.indexOf('ISBN');
  if (isbnCol < 0) return [];
  var values = sheet.getRange(2, isbnCol + 1, sheet.getLastRow() - 1, 1).getValues();
  var seen = {};
  var result = [];
  for (var i = 0; i < values.length; i++) {
    var isbn = _normalizeIsbn(values[i][0]);
    if (!isbn || seen[isbn]) continue;
    seen[isbn] = true;
    result.push(isbn);
  }
  return result;
}

/**
 * Cross-reference the user's library ISBNs against the NYT cache.
 * Returns a map of { isbn: { rank, weeksOn, list, listDisplay } }
 * for all matched books. Called from getDashboardData flow.
 */
function clientGetNYTBadgesForLibrary() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('NYT_CACHE');
  if (!raw) return { byBookId: {}, byIsbn: {} };
  var cache;
  try { cache = JSON.parse(raw); } catch(e) { return { byBookId: {}, byIsbn: {} }; }

  // Read all ISBNs from the Library sheet
  var sheet = _ss().getSheetByName(SHEET_LIBRARY);
  if (!sheet || sheet.getLastRow() < 2) return { byBookId: {}, byIsbn: {} };

  var isbnCol = LIBRARY_HEADERS.indexOf('ISBN');
  var bookIdCol = LIBRARY_HEADERS.indexOf('BookId');
  var data = sheet.getDataRange().getValues();
  var byBookId = {};
  var byIsbn = {};

  for (var r = 1; r < data.length; r++) {
    var isbn = _normalizeIsbn(data[r][isbnCol] || '');
    var bookId = String(data[r][bookIdCol] || '').trim();
    if (isbn && cache[isbn]) {
      var badge = {
        rank:        cache[isbn].rank,
        weeksOn:     cache[isbn].weeksOn,
        list:        cache[isbn].list,
        listDisplay: cache[isbn].listDisplay
      };
      byBookId[bookId] = badge;
      byIsbn[isbn] = badge;
    }
  }
  return { byBookId: byBookId, byIsbn: byIsbn };
}

function clientGetNYTFeed() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('NYT_FEED_CURRENT');
  if (!raw) return { updatedAt: '', lists: [] };
  try {
    var parsed = JSON.parse(raw);
    return {
      updatedAt: parsed.updatedAt || '',
      lists: Array.isArray(parsed.lists) ? parsed.lists : []
    };
  } catch (e) {
    return { updatedAt: '', lists: [] };
  }
}

/**
 * Install the weekly time-based trigger for NYT cache refresh.
 * Run this manually once from the Apps Script editor after deployment.
 * It is safe to run multiple times — it checks for duplicates first.
 */
function installNYTWeeklyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'clientRefreshNYTCache') {
      Logger.log('Trigger already installed.');
      return;
    }
  }
  ScriptApp.newTrigger('clientRefreshNYTCache')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(3)
    .create();
  Logger.log('Weekly NYT refresh trigger installed (Mondays at 3 AM).');
}

// =====================================================================
//  SYNC VERSION — Sheet→Webapp change detection
// =====================================================================

/** Return the current sync version counter. Webapp polls this to detect sheet edits. */
function clientGetSyncVersion() {
  var props = PropertiesService.getScriptProperties();
  return Number(props.getProperty('SYNC_VERSION') || '0');
}

/** Increment sync version. Called by onEdit trigger when user edits in the sheet. */
function _incrementSyncVersion() {
  var props = PropertiesService.getScriptProperties();
  var current = Number(props.getProperty('SYNC_VERSION') || '0');
  props.setProperty('SYNC_VERSION', String(current + 1));
}

/**
 * Installable onEdit trigger handler — detects when data is edited in the
 * Library, Challenges, Shelves, or Profile tabs and bumps the sync version
 * so the webapp's poll loop can detect the change.
 */
function onEditSyncHandler(e) {
  if (!e || !e.range) return;
  var sheetName = e.range.getSheet().getName();
  var syncedSheets = [SHEET_LIBRARY, SHEET_CHALLENGES, SHEET_SHELVES, SHEET_PROFILE, SHEET_AUDIOBOOKS];
  if (syncedSheets.indexOf(sheetName) >= 0) {
    _incrementSyncVersion();
  }
}

/**
 * Install the installable onEdit trigger for sync version tracking.
 * Safe to run multiple times — checks for duplicates.
 */
function installSyncTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEditSyncHandler') {
      Logger.log('Sync trigger already installed.');
      SpreadsheetApp.getUi().alert('Sync trigger is already installed.');
      return;
    }
  }
  ScriptApp.newTrigger('onEditSyncHandler')
    .forSpreadsheet(_ss())
    .onEdit()
    .create();
  Logger.log('Sync trigger installed.');
  SpreadsheetApp.getUi().alert('Sync trigger installed. Changes in the sheet will now sync to the web app.');
}

// =====================================================================
//  STORYGRAPH CSV IMPORT
// =====================================================================

/**
 * Import a StoryGraph CSV export.
 * Accepts either raw StoryGraph CSV columns or structured payload rows
 * (from toServerBookPayload). Auto-detects by checking header names.
 * @param {Array<Array>} rows  2D array, first row = headers
 */
function clientImportStoryGraphCSV(rows) {
  try {
  if (!rows || rows.length < 2) return { imported: 0 };
  var sheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
  var headers = rows[0].map(function(h) { return String(h).toLowerCase().trim(); });

  function col(name) {
    var idx = headers.indexOf(name.toLowerCase());
    return idx >= 0 ? idx : -1;
  }

  function cell(row, name) {
    var i = col(name);
    return i >= 0 ? String(row[i] || '').trim() : '';
  }

  // Detect if this is a structured payload (from toServerBookPayload)
  var isStructured = headers.includes('title') && headers.includes('pagecount');

  var _sgStatusMap = {
    'read':          'Finished',
    'currently reading': 'Reading',
    'to read':       'Want to Read',
    'did not finish':'DNF'
  };

  var newRows = [];
  for (var r = 1; r < rows.length; r++) {
    var d = rows[r];

    if (isStructured) {
      // Structured payload from UI — just pass through to _bookPayloadToRow
      var title = cell(d, 'title');
      if (!title) continue;
      var payload = {
        Title: title,
        Author: cell(d, 'author'),
        Status: cell(d, 'status') || 'Want to Read',
        Rating: Number(cell(d, 'rating') || 0),
        PageCount: Number(cell(d, 'pagecount') || 0),
        Genre: cell(d, 'genre') || cell(d, 'genres') || '',
        DateAdded: cell(d, 'dateadded') || new Date().toISOString().slice(0,10),
        DateStarted: cell(d, 'datestarted') || '',
        DateFinished: cell(d, 'datefinished') || '',
        CurrentPage: Number(cell(d, 'currentpage') || 0),
        Series: cell(d, 'series') || '',
        SeriesNumber: cell(d, 'seriesnumber') || '',
        TbrPriority: Number(cell(d, 'tbrpriority') || 0),
        Format: cell(d, 'format') || '',
        Source: cell(d, 'source') || '',
        SpiceLevel: Number(cell(d, 'spicelevel') || 0),
        Tags: cell(d, 'tags') || '',
        Shelves: cell(d, 'shelves') || '',
        Notes: cell(d, 'notes') || '',
        Review: cell(d, 'review') || '',
        Quotes: cell(d, 'quotes') || '',
        Favorite: cell(d, 'favorite') === 'true',
        CoverEmoji: cell(d, 'coveremoji') || '',
        CoverUrl: cell(d, 'coverurl') || '',
        Gradient1: cell(d, 'gradient1') || '',
        Gradient2: cell(d, 'gradient2') || '',
        ISBN: cell(d, 'isbn') || '',
        OLID: cell(d, 'olid') || '',
        AuthorKey: cell(d, 'authorkey') || ''
      };
      newRows.push(_bookPayloadToRow(_uuid(), payload));
      continue;
    }

    // Original StoryGraph CSV format
    var sgTitle  = cell(d, 'title');
    var author = cell(d, 'authors') || cell(d, 'author');
    if (!sgTitle) continue;

    var sgStatus = cell(d, 'read status').toLowerCase();
    var sheetStatus = _sgStatusMap[sgStatus] || 'Want to Read';

    var rating = 0;
    var ratingRaw = cell(d, 'star rating');
    if (ratingRaw) rating = Math.min(5, Math.max(0, Math.round(parseFloat(ratingRaw) || 0)));

    var dateRead = cell(d, 'last date read') || cell(d, 'dates read') || '';
    // StoryGraph date format may be "YYYY/MM/DD" — normalize to YYYY-MM-DD
    if (dateRead && dateRead.indexOf('/') >= 0) {
      dateRead = dateRead.split(' ')[0].replace(/\//g, '-');
    }

    var genres  = cell(d, 'genres') || cell(d, 'genre') || '';
    var moods   = cell(d, 'moods') || cell(d, 'tags') || '';
    var review  = cell(d, 'review') || '';
    var pages   = Number(cell(d, 'number of pages') || cell(d, 'pages') || 0);
    var pace    = cell(d, 'pace') || '';
    var notes   = (pace ? 'Pace: ' + pace + '\n' : '') + (moods ? 'Moods: ' + moods : '');

    newRows.push(_bookPayloadToRow(_uuid(), {
      Title:      sgTitle,
      Author:     author,
      Status:     sheetStatus,
      Rating:     rating,
      PageCount:  pages,
      Genres:     genres,
      Moods:      moods,
      Review:     review,
      Notes:      notes.trim(),
      DateAdded:  dateRead || new Date().toISOString().slice(0, 10),
      DateFinished: sheetStatus === 'Finished' ? dateRead : ''
    }));
  }

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, LIBRARY_HEADERS.length).setValues(newRows);
  }
  return { imported: newRows.length };
  } catch(e) { return { error: e.message, imported: 0 }; }
}

// ── Add menu items for new functions ────────────────────────────────────
// (appended to existing onOpen menu)

