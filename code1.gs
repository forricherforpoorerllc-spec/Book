/* =====================================================================
 *  code1.gs — My Reading Journey (Standalone)
 *
 *  Single-file Google Apps Script backend + sheet layout for the
 *  My Reading Journey template (Etsy release).
 *
 *  WHAT THIS FILE DOES:
 *  - Serves the web app (doGet → index.html)
 *  - Owns all sheet constants & schema
 *  - Builds the visible Library tab to match the product screenshot:
 *      Row 1 = merged "📚 LIBRARY ▾" title chip
 *      Row 2 = column headers with icon glyphs
 *      Row 3+ = data rows (no cover images, normal ~32px rows)
 *  - Keeps every utility sheet hidden (Challenges, Shelves, Profile, Audiobooks)
 *  - Seeds 26 demo books automatically on first open
 *  - Exposes the full client* API surface used by index.html / index2.html / index3.html
 *
 *  DEPLOY:
 *  - Extensions → Apps Script → paste this file → Save
 *  - Delete any older Code.gs if present (all runtime now lives here)
 *  - Run initializeSheets() once from the menu to build the sheet + seed demo
 *  - Deploy → New deployment → Web app → Execute as *me*, access Anyone
 *  - (Optional) Menu → Install Sync Trigger  for live sheet↔web-app sync
 * ===================================================================== */

function _dbLiteInitializeSheets() {
	var theme = _getCurrentTheme();
	var ss = _ss();

	// Ensure core data sheets exist with correct headers.
	var library = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
	_getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
	_getOrCreateSheet(SHEET_SHELVES, SHELF_HEADERS);
	var profile = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
	_getOrCreateSheet(SHEET_AUDIOBOOKS, AUDIOBOOK_HEADERS);

	_dbLiteEnsureProfileDefaults(profile);
	_dbLiteInitLibrarySheet(library, theme);
	_dbLiteInitMyYearSheet(ss, theme);

	// Seed demo on first run only (also runs unconditionally on first open via onOpen).
	_seedDemoData();

	_dbLiteArrangeTabs(ss);
}

function _dbLiteTheme(themeName) {
	var p = _getThemePalette(themeName);
	var accent = p && p.header ? p.header : '#C85888';
	var accent2 = p && p.accent ? p.accent : '#F8D4A0';
	var isDark = _dbLiteIsDark(accent);
	return {
		accent: accent,
		accent2: accent2,
		headerBg: accent,
		headerText: isDark ? '#FFFFFF' : '#111111',
		border: accent,
		text: '#111111',
		white: '#FFFFFF',
		lightGray: '#F7F7F7',
		grid: '#E6E6E6'
	};
}

function _dbLiteIsDark(hex) {
	var h = String(hex || '').replace('#', '');
	if (h.length !== 6) return true;
	var r = parseInt(h.substring(0, 2), 16);
	var g = parseInt(h.substring(2, 4), 16);
	var b = parseInt(h.substring(4, 6), 16);
	var yiq = (r * 299 + g * 587 + b * 114) / 1000;
	return yiq < 145;
}

function _dbLiteEnsureProfileDefaults(sheet) {
	if (!sheet) return;
	if (sheet.getLastRow() >= 2) return;

	var defaults = PROFILE_HEADERS.map(function(h) {
		switch (h) {
			case 'Name': return '';
			case 'Motto': return 'A focused place to track every book';
			case 'PhotoData': return '';
			case 'Theme': return 'blossom';
			case 'YearlyGoal': return 50;
			case 'Onboarded': return false;
			case 'DemoCleared': return false;
			case 'ShowSpoilers': return true;
			case 'ReadingOrder': return '[]';
			case 'RecentIds': return '[]';
			case 'SortBy': return 'default';
			case 'LibViewMode': return 'grid';
			case 'SelectedFilter': return 'all';
			case 'ActiveShelf': return '';
			case 'ChallengeBarCollapsed': return false;
			case 'LibToolsOpen': return false;
			case 'LibraryName': return 'My Library';
			case 'CustomQuotes': return '[]';
			case 'CoversEnabled': return true;
			case 'TutorialCompleted': return false;
			case 'LastAudioId': return '';
			case 'TotalListeningMins': return 0;
			default: return '';
		}
	});
	sheet.getRange(2, 1, 1, PROFILE_HEADERS.length).setValues([defaults]);
}

function _dbLiteArrangeTabs(ss) {
	var order = [
		SHEET_LIBRARY,
		SHEET_MYYEAR,
		SHEET_CHALLENGES,
		SHEET_SHELVES,
		SHEET_PROFILE,
		SHEET_AUDIOBOOKS
	];

	order.forEach(function(name, i) {
		var s = ss.getSheetByName(name);
		if (!s) return;
		ss.setActiveSheet(s);
		ss.moveActiveSheet(i + 1);
	});

	// Library and My Year are visible; all data sheets are hidden.
	var allSheets = ss.getSheets();
	allSheets.forEach(function(s) {
		var n = s.getName();
		if (n === SHEET_LIBRARY || n === SHEET_MYYEAR) {
			try { s.showSheet(); } catch (e) {}
		} else {
			try { s.hideSheet(); } catch (e) {}
		}
	});

	var lib = ss.getSheetByName(SHEET_LIBRARY);
	if (lib) ss.setActiveSheet(lib);
}

function _dbLiteInitLibrarySheet(sheet, themeName) {
	var t = _dbLiteTheme(themeName);
	// Col A = row-number formula. LIBRARY_HEADERS data starts at LIBRARY_DATA_COL (B).
	var totalCols = LIBRARY_HEADERS.length + LIBRARY_DATA_COL - 1;
	_ensureColumns(sheet, totalCols);
	_ensureRows(sheet, 5008); // rows 1-8 template + 5000 data rows

	// Strip prior conditional formatting and filter only (rows 1-8 visuals untouched)
	try { sheet.clearConditionalFormatRules(); } catch(e) {}
	try { if (sheet.getFilter()) sheet.getFilter().remove(); } catch(e) {}
	try { sheet.getBandings().forEach(function(b) { b.remove(); }); } catch(e) {}

	sheet.setTabColor(t.accent);
	sheet.setHiddenGridlines(true);

	// ── Rows 1–6: banner background = theme primary color ──────────────────
	// User manually places a floating image on top; code only sets the color.
	// Row 7 stays white as a spacer before the row-8 header band.
	try {
		sheet.getRange(1, 1, 6, totalCols).setBackground(t.headerBg);
	} catch (e) {}

	// ── Column A: auto row-number formula (=1,2,3... fills as books are added) ──
	sheet.setColumnWidth(1, 40);
	var rowNumFormulas = [];
	for (var rn = 0; rn < 5000; rn++) {
		var dr = LIBRARY_DATA_ROW + rn;
		rowNumFormulas.push(['=IF(B' + dr + '<>"",ROW()-' + (LIBRARY_DATA_ROW - 1) + ',"")']);
	}
	sheet.getRange(LIBRARY_DATA_ROW, 1, 5000, 1)
		.setFormulas(rowNumFormulas)
		.setFontFamily('Montserrat').setFontSize(9)
		.setFontColor('#9CA3AF').setHorizontalAlignment('center')
		.setVerticalAlignment('middle').setNumberFormat('#');

	// ── Visible column widths (B–M: Title through Favorite) ─────────────────
	var visibleWidths = [270, 190, 130, 140, 110, 120, 80, 115, 120, 160, 70, 90];
	for (var vi = 0; vi < LIBRARY_VISIBLE_COUNT; vi++) {
		sheet.setColumnWidth(LIBRARY_DATA_COL + vi, visibleWidths[vi] || 120);
	}

	// ── Hide all hidden columns (N onward — webapp internals) ────────────────
	var hiddenStart = LIBRARY_DATA_COL + LIBRARY_VISIBLE_COUNT;
	var hiddenCount = LIBRARY_HEADERS.length - LIBRARY_VISIBLE_COUNT;
	try { sheet.showColumns(hiddenStart, hiddenCount); } catch(e) {}
	try { sheet.hideColumns(hiddenStart, hiddenCount); } catch(e) {}

	// ── Data area: font, background, wrap (rows 9–5008, visible cols B–M) ───
	sheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, 5000, LIBRARY_VISIBLE_COUNT)
		.setFontFamily('Montserrat').setFontSize(11)
		.setFontColor('#1F2937').setVerticalAlignment('middle')
		.setBackground('#FFFFFF').setWrap(false);

	// ── Row heights (all 5000 data rows at once) ──────────────────────────────
	sheet.setRowHeights(LIBRARY_DATA_ROW, 5000, 44);

	// ── Per-column alignment + number formats ─────────────────────────────────
	var titleCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Title');
	var authorCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Author');
	var pagesCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Pages');
	var dsCol     = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('DateStarted');
	var dfCol     = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('DateFinished');
	var snCol     = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('SeriesNumber');
	var favCol    = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Favorite');

	sheet.getRange(LIBRARY_DATA_ROW, titleCol,  5000, 1).setFontWeight('bold').setHorizontalAlignment('left');
	sheet.getRange(LIBRARY_DATA_ROW, authorCol, 5000, 1).setHorizontalAlignment('left');
	sheet.getRange(LIBRARY_DATA_ROW, pagesCol,  5000, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
	sheet.getRange(LIBRARY_DATA_ROW, dsCol,     5000, 1).setNumberFormat('mmm d, yyyy').setHorizontalAlignment('center');
	sheet.getRange(LIBRARY_DATA_ROW, dfCol,     5000, 1).setNumberFormat('mmm d, yyyy').setHorizontalAlignment('center');
	sheet.getRange(LIBRARY_DATA_ROW, snCol,     5000, 1).setHorizontalAlignment('center');

	// Center-align chip columns
	['Status','Genre','Rating','Format'].forEach(function(h) {
		var col = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf(h);
		sheet.getRange(LIBRARY_DATA_ROW, col, 5000, 1).setHorizontalAlignment('center');
	});

	// ── Hairline bottom border on every data row ──────────────────────────────
	sheet.getRange(LIBRARY_DATA_ROW, 1, 5000, totalCols)
		.setBorder(null, null, true, null, null, null, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);

	// ── Freeze template header row + filter on visible columns ───────────────
	sheet.setFrozenRows(LIBRARY_HEADER_ROW);
	sheet.setFrozenColumns(0);
	try {
		sheet.getRange(LIBRARY_HEADER_ROW, LIBRARY_DATA_COL, 5001, LIBRARY_VISIBLE_COUNT).createFilter();
	} catch(e) {}

	// ── Favorite checkboxes ───────────────────────────────────────────────────
	sheet.getRange(LIBRARY_DATA_ROW, favCol, 5000, 1).insertCheckboxes();

	// ── Chip dropdowns + colored pill conditional formatting ─────────────────
	_dbLiteApplyValidations(sheet);
	_dbLiteApplyPillFormatting(sheet);
}

function _dbLiteInitMyYearSheet(ss, themeName) {
	var sheet = ss.getSheetByName(SHEET_MYYEAR);
	if (!sheet) sheet = ss.insertSheet(SHEET_MYYEAR);
	var t = _dbLiteTheme(themeName);

	sheet.clearContents();
	sheet.clearFormats();
	try { if (sheet.getFilter()) sheet.getFilter().remove(); } catch(e) {}
	try { sheet.clearConditionalFormatRules(); } catch(e) {}
	try { sheet.getBandings().forEach(function(b) { b.remove(); }); } catch(e) {}

	var COVERS_PER_ROW = 6;
	var COVER_W = 90;
	var COVER_H = 135;
	var NUM_COLS = 7; // col A: labels, cols B-G: covers
	_ensureColumns(sheet, NUM_COLS);
	_ensureRows(sheet, 2000);
	sheet.setHiddenGridlines(true);
	sheet.setTabColor(t.accent);
	sheet.setColumnWidth(1, 120);
	for (var c = 2; c <= NUM_COLS; c++) sheet.setColumnWidth(c, COVER_W);

	var row = 1;

	// ── Banner ────────────────────────────────────────────────────────────
	sheet.getRange(row, 1, 1, NUM_COLS).merge()
		.setValue('📚  MY YEAR')
		.setBackground(t.headerBg).setFontColor(t.headerText)
		.setFontFamily('Montserrat').setFontSize(24).setFontWeight('bold')
		.setHorizontalAlignment('center').setVerticalAlignment('middle');
	sheet.setRowHeight(row, 72);
	row++;

	// ── Challenges ────────────────────────────────────────────────────────
	var chalSheet = ss.getSheetByName(SHEET_CHALLENGES);
	var challenges = chalSheet ? _sheetToObjects(chalSheet, CHALLENGE_HEADERS) : [];

	if (challenges.length) {
		sheet.getRange(row, 1, 1, NUM_COLS).merge()
			.setValue('🎯  READING CHALLENGES')
			.setBackground(t.accent2 || '#F3F4F6').setFontColor(t.headerBg)
			.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold')
			.setHorizontalAlignment('left').setVerticalAlignment('middle')
			.setBorder(null, null, true, null, null, null, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);
		sheet.setRowHeight(row, 32);
		row++;

		challenges.forEach(function(ch) {
			var current = Number(ch.Current) || 0;
			var target  = Math.max(1, Number(ch.Target) || 1);
			var pct     = Math.min(1, current / target);
			var filled  = Math.round(pct * 20);
			var bar     = new Array(filled + 1).join('█') + new Array(20 - filled + 1).join('░');
			sheet.getRange(row, 1)
				.setValue(String(ch.Icon || '📖').slice(0, 4) + '  ' + String(ch.Name || ''))
				.setFontFamily('Montserrat').setFontSize(10).setFontWeight('bold')
				.setFontColor('#1F2937').setVerticalAlignment('middle');
			sheet.getRange(row, 2, 1, 3).merge()
				.setValue(bar).setFontColor(t.headerBg).setFontSize(9).setVerticalAlignment('middle');
			sheet.getRange(row, 5)
				.setValue(current + ' / ' + target)
				.setFontFamily('Montserrat').setFontSize(10).setFontColor('#4B5563')
				.setHorizontalAlignment('center').setVerticalAlignment('middle');
			sheet.getRange(row, 6)
				.setValue(Math.round(pct * 100) + '%')
				.setFontFamily('Montserrat').setFontSize(10).setFontWeight('bold')
				.setFontColor(t.headerBg).setHorizontalAlignment('center').setVerticalAlignment('middle');
			sheet.getRange(row, 1, 1, NUM_COLS)
				.setBorder(null, null, true, null, null, null, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);
			sheet.setRowHeight(row, 36);
			row++;
		});
	}

	// Gap
	sheet.setRowHeight(row, 24);
	row++;

	// ── Book Covers header ─────────────────────────────────────────────────
	sheet.getRange(row, 1, 1, NUM_COLS).merge()
		.setValue('📖  BOOK COVERS')
		.setBackground(t.accent2 || '#F3F4F6').setFontColor(t.headerBg)
		.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold')
		.setHorizontalAlignment('left').setVerticalAlignment('middle')
		.setBorder(null, null, true, null, null, null, '#E5E7EB', SpreadsheetApp.BorderStyle.SOLID);
	sheet.setRowHeight(row, 32);
	row++;

	// ── Read finished books with cover URLs from Library ──────────────────
	var libSheet = ss.getSheetByName(SHEET_LIBRARY);
	var byMonth  = {};
	var monthOrder = [];
	if (libSheet && libSheet.getLastRow() >= LIBRARY_DATA_ROW) {
		var numDR = libSheet.getLastRow() - LIBRARY_DATA_ROW + 1;
		var libData = libSheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, numDR, LIBRARY_HEADERS.length).getValues();
		var tIdx   = LIBRARY_HEADERS.indexOf('Title');
		var cuIdx  = LIBRARY_HEADERS.indexOf('CoverUrl');
		var dfIdx  = LIBRARY_HEADERS.indexOf('DateFinished');
		var MN     = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
		libData.forEach(function(dr) {
			if (!String(dr[tIdx] || '').trim()) return;
			var url = String(dr[cuIdx] || '').trim();
			if (!url || url.indexOf('http') !== 0) return;
			var df = dr[dfIdx];
			if (!df) return;
			var d = new Date(df);
			if (isNaN(d.getTime())) return;
			var key   = d.getFullYear() + String(d.getMonth() + 1).padStart(2, '0');
			var label = MN[d.getMonth()] + ' ' + d.getFullYear();
			if (!byMonth[key]) { byMonth[key] = { label: label, covers: [] }; monthOrder.push(key); }
			byMonth[key].covers.push(url.replace(/["']/g, ''));
		});
		monthOrder.sort();
	}

	if (!monthOrder.length) {
		sheet.getRange(row, 1, 1, NUM_COLS).merge()
			.setValue('No finished books with covers yet — add a book in the web app, mark it finished, then sync.')
			.setFontFamily('Montserrat').setFontSize(10).setFontColor('#9CA3AF')
			.setHorizontalAlignment('center').setVerticalAlignment('middle');
		sheet.setRowHeight(row, 44);
		return;
	}

	monthOrder.forEach(function(key) {
		var grp = byMonth[key];
		// Month label row
		sheet.getRange(row, 1, 1, NUM_COLS).merge()
			.setValue(grp.label)
			.setBackground('#F9FAFB').setFontFamily('Montserrat').setFontSize(10)
			.setFontWeight('bold').setFontColor(t.headerBg)
			.setHorizontalAlignment('left').setVerticalAlignment('middle');
		sheet.setRowHeight(row, 28);
		row++;
		// Cover grid
		var col = 2; // start at column B
		grp.covers.forEach(function(url) {
			sheet.getRange(row, col)
				.setFormula('=IMAGE("' + url + '",4,' + COVER_H + ',' + COVER_W + ')')
				.setVerticalAlignment('middle').setHorizontalAlignment('center')
				.setBackground('#FFFFFF');
			sheet.setRowHeight(row, COVER_H);
			col++;
			if (col > COVERS_PER_ROW + 1) { col = 2; row++; sheet.setRowHeight(row, COVER_H); }
		});
		if (col > 2) row++; // finish partial row
		sheet.setRowHeight(row, 14);
		row++;
	});
}

function _dbLiteApplyValidations(sheet) {
	var startRow = LIBRARY_DATA_ROW;
	var dataRows = 5000;
	var statusCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Status');
	var genreCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Genre');
	var ratingCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Rating');
	var formatCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Format');

	sheet.getRange(startRow, statusCol, dataRows, 1).setDataValidation(
		SpreadsheetApp.newDataValidation()
			.requireValueInList(['Reading','Finished','Want to Read','DNF'], true)
			.setAllowInvalid(false).build()
	);
	sheet.getRange(startRow, genreCol, dataRows, 1).setDataValidation(
		SpreadsheetApp.newDataValidation()
			.requireValueInList([
				'Romance','Fantasy','Mystery','Thriller','SciFi','Historical',
				'Memoir','Biography','Self-Help','Nonfiction','Fiction',
				'Horror','YA','Poetry','Classics','Literary','Graphic','Other'
			], true)
			.setAllowInvalid(false).build()
	);
	// Rating stored as star strings (★ through ★★★★★) — chip dropdown
	sheet.getRange(startRow, ratingCol, dataRows, 1).setDataValidation(
		SpreadsheetApp.newDataValidation()
			.requireValueInList(['★','★★','★★★','★★★★','★★★★★'], true)
			.setAllowInvalid(true).build()
	);
	sheet.getRange(startRow, formatCol, dataRows, 1).setDataValidation(
		SpreadsheetApp.newDataValidation()
			.requireValueInList(['Paperback','Hardcover','Ebook','Audiobook'], true)
			.setAllowInvalid(true).build()
	);
}

function _dbLiteApplyPillFormatting(sheet) {
	var startRow  = LIBRARY_DATA_ROW;
	var dataRows  = 5000;
	var statusCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Status');
	var genreCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Genre');
	var ratingCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Rating');
	var formatCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Format');
	var rules = [];

	function pill(col, value, bg, fg) {
		return SpreadsheetApp.newConditionalFormatRule()
			.whenTextEqualTo(value).setBackground(bg).setFontColor(fg).setBold(true)
			.setRanges([sheet.getRange(startRow, col, dataRows, 1)]).build();
	}

	// Status — pastel
	rules.push(pill(statusCol, 'Reading',      '#BFDBFE', '#1E3A8A'));
	rules.push(pill(statusCol, 'Finished',     '#BBF7D0', '#14532D'));
	rules.push(pill(statusCol, 'Want to Read', '#FED7AA', '#7C2D12'));
	rules.push(pill(statusCol, 'DNF',          '#E5E7EB', '#374151'));

	// Genre — vivid deep colors, white text (matches template screenshot exactly)
	rules.push(pill(genreCol, 'Romance',    '#991B1B', '#FFFFFF'));
	rules.push(pill(genreCol, 'Fantasy',    '#5B21B6', '#FFFFFF'));
	rules.push(pill(genreCol, 'Mystery',    '#312E81', '#FFFFFF'));
	rules.push(pill(genreCol, 'Thriller',   '#7F1D1D', '#FFFFFF'));
	rules.push(pill(genreCol, 'SciFi',      '#155E75', '#FFFFFF'));
	rules.push(pill(genreCol, 'Historical', '#78350F', '#FFFFFF'));
	rules.push(pill(genreCol, 'Memoir',     '#115E59', '#FFFFFF'));
	rules.push(pill(genreCol, 'Biography',  '#1E3A8A', '#FFFFFF'));
	rules.push(pill(genreCol, 'Self-Help',  '#14532D', '#FFFFFF'));
	rules.push(pill(genreCol, 'Nonfiction', '#334155', '#FFFFFF'));
	rules.push(pill(genreCol, 'Fiction',    '#1F2937', '#FFFFFF'));
	rules.push(pill(genreCol, 'Horror',     '#0F172A', '#FFFFFF'));
	rules.push(pill(genreCol, 'YA',         '#9D174D', '#FFFFFF'));
	rules.push(pill(genreCol, 'Poetry',     '#6B21A8', '#FFFFFF'));
	rules.push(pill(genreCol, 'Classics',   '#92400E', '#FFFFFF'));
	rules.push(pill(genreCol, 'Literary',   '#1E293B', '#FFFFFF'));
	rules.push(pill(genreCol, 'Graphic',    '#7C3AED', '#FFFFFF'));
	rules.push(pill(genreCol, 'Other',      '#4B5563', '#FFFFFF'));

	// Rating — graduated amber (★ = pale, ★★★★★ = full gold)
	rules.push(pill(ratingCol, '★',     '#FFF7ED', '#C2410C'));
	rules.push(pill(ratingCol, '★★',    '#FEF3C7', '#92400E'));
	rules.push(pill(ratingCol, '★★★',   '#FDE68A', '#78350F'));
	rules.push(pill(ratingCol, '★★★★',  '#FCD34D', '#451A03'));
	rules.push(pill(ratingCol, '★★★★★', '#F59E0B', '#451A03'));

	// Format — soft pastels
	rules.push(pill(formatCol, 'Paperback',  '#EDE9FE', '#5B21B6'));
	rules.push(pill(formatCol, 'Hardcover',  '#DBEAFE', '#1E40AF'));
	rules.push(pill(formatCol, 'Ebook',      '#CFFAFE', '#155E75'));
	rules.push(pill(formatCol, 'Audiobook',  '#FDF4FF', '#6B21A8'));

	sheet.setConditionalFormatRules(rules);
}

/* =====================================================================
 *  ── STANDALONE RUNTIME ─────────────────────────────────────────────
 *  Everything below this line is the web-app engine that used to live
 *  in Code.gs. Paste this file on its own into Apps Script and you're
 *  done — no second .gs file required.
 * ===================================================================== */

// ── Serve the UI ────────────────────────────────────────────────────────
function doGet() {
	var title = _buildJourneyTitle();
	var output = HtmlService.createHtmlOutputFromFile('index')
		.setTitle(title)
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
	try {
		var xfMode = HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.SAMEORIGIN;
		if (xfMode != null) output.setXFrameOptionsMode(xfMode);
	} catch (e) {}
	return output;
}

function _buildJourneyTitle() {
	var sheet = _ss().getSheetByName(SHEET_PROFILE);
	if (sheet && sheet.getLastRow() >= 2) {
		var name = String(sheet.getRange(2, 1).getValue() || '').trim();
		if (name) {
			var possessive = name.slice(-1) === 's' ? name + "'" : name + "'s";
			return possessive + ' Reading Journey';
		}
	}
	return 'My Reading Journey';
}

// ── Sheet names / schema ────────────────────────────────────────────────
var SHEET_LIBRARY     = 'Library';
var SHEET_CHALLENGES  = 'Challenges';
var SHEET_SHELVES     = 'Shelves';
var SHEET_PROFILE     = 'Profile';
var SHEET_AUDIOBOOKS  = 'Audiobooks';
var SHEET_MYYEAR      = 'My Year';

// ── Library layout constants ─────────────────────────────────────────────
// Col A = auto row-number formula. Data starts at column B (LIBRARY_DATA_COL).
// Rows 1-7 = banner/image (never touched by code). Row 8 = headers. Row 9+ = data.
var LIBRARY_DATA_COL      = 2;   // Column B
var LIBRARY_HEADER_ROW    = 8;
var LIBRARY_DATA_ROW      = 9;
var LIBRARY_VISIBLE_COUNT = 12;  // Title → Favorite (cols B–M)

var LIBRARY_HEADERS = [
	// ── Visible columns (B–M, 12 total) ─────────────────────────────────
	'Title','Author','Status','Genre','Rating','Format','Pages',
	'DateStarted','DateFinished','Series','SeriesNumber','Favorite',
	// ── Hidden columns (N onward — webapp internals, never visible) ──────
	'BookId','CoverUrl','CoverEmoji','Gradient1','Gradient2',
	'DateAdded','CurrentPage','TbrPriority','Source','SpiceLevel',
	'Tags','Shelves','Notes','Review','Quotes',
	'ISBN','OLID','AuthorKey'
];

var CHALLENGE_HEADERS = ['ChallengeId','Name','Icon','Current','Target'];
var SHELF_HEADERS = ['ShelfId','Name','Icon'];

var PROFILE_HEADERS = [
	'Name','Motto','PhotoData','Theme',
	'YearlyGoal','Onboarded','DemoCleared','ShowSpoilers',
	'ReadingOrder','RecentIds','SortBy','LibViewMode',
	'SelectedFilter','ActiveShelf','ChallengeBarCollapsed','LibToolsOpen',
	'LibraryName',
	'CustomQuotes','CoversEnabled','TutorialCompleted',
	'LastAudioId','TotalListeningMins'
];

var AUDIOBOOK_HEADERS = [
	'AudiobookId','Title','Author','Duration','CoverEmoji','CoverUrl',
	'ChapterCount','LibrivoxProjectId','CurrentChapterIndex',
	'CurrentTime','PlaybackSpeed','TotalListeningMins'
];

// ── Theme palettes (drive sheet tab + title chip colors) ────────────────
var THEME_PALETTES = {
	romantic:  { header: '#B91C1C', headerText: '#FFFFFF', accent: '#DC2626', border: '#FCA5A5' },
	spicy:     { header: '#8F001F', headerText: '#FFE4E8', accent: '#FF4D4D', border: '#5A001A' },
	dreamy:    { header: '#8B5CF6', headerText: '#FFFFFF', accent: '#A78BFA', border: '#DDD6FE' },
	fresh:     { header: '#10B981', headerText: '#FFFFFF', accent: '#34D399', border: '#A7F3D0' },
	midnight:  { header: '#312E81', headerText: '#E0E7FF', accent: '#818CF8', border: '#334155' },
	sunset:    { header: '#1A0525', headerText: '#FFF0F5', accent: '#FF6D00', border: '#6A1F5C' },
	velvet:    { header: '#4C1D95', headerText: '#F5F3FF', accent: '#D946EF', border: '#3B1F6E' },
	horizon:   { header: '#0284C7', headerText: '#FFFFFF', accent: '#22D3EE', border: '#BAE6FD' },
	arctic:    { header: '#0C2A4A', headerText: '#E0F2FE', accent: '#38BDF8', border: '#1E3A5F' },
	sahara:    { header: '#D97706', headerText: '#FFFFFF', accent: '#F59E0B', border: '#FDE68A' },
	ember:     { header: '#122010', headerText: '#E8DFC8', accent: '#D4A030', border: '#1E3420' },
	volcano:   { header: '#B91C1C', headerText: '#FFFFFF', accent: '#F59E0B', border: '#FCA5A5' },
	dusk:      { header: '#1A0525', headerText: '#FFF0F5', accent: '#FFCA0A', border: '#6A1F5C' },
	petal:     { header: '#C06C84', headerText: '#FFFFFF', accent: '#F67280', border: '#F9B2D7' },
	coral:     { header: '#355C7D', headerText: '#FDF2F8', accent: '#F8B195', border: '#3D5A7A' },
	lagoon:    { header: '#229799', headerText: '#FFFFFF', accent: '#48CFCB', border: '#A2D5C6' },
	'mint mist':{ header: '#229799', headerText: '#FFFFFF', accent: '#48CFCB', border: '#A2D5C6' },
	jade:      { header: '#237227', headerText: '#F0FDF4', accent: '#CFFFE2', border: '#1A3D22' },
	'sage forest': { header: '#237227', headerText: '#F0FDF4', accent: '#CFFFE2', border: '#1A3D22' },
	champagne: { header: '#A16207', headerText: '#FFFFFF', accent: '#FACC15', border: '#FDE68A' },
	obsidian:  { header: '#111827', headerText: '#F9FAFB', accent: '#6366F1', border: '#1F2937' },
	pearl:     { header: '#9CA3AF', headerText: '#FFFFFF', accent: '#D1D5DB', border: '#E5E7EB' },
	opal:      { header: '#0EA5E9', headerText: '#FFFFFF', accent: '#67E8F9', border: '#BAE6FD' },
	onyx:      { header: '#1E293B', headerText: '#F8FAFC', accent: '#64748B', border: '#334155' },
	bunny:     { header: '#EC4899', headerText: '#FFFFFF', accent: '#FBCFE8', border: '#FCE7F3' },
	blossom:   { header: '#C85888', headerText: '#FFFFFF', accent: '#F8D4A0', border: '#F7C9D5' },
	lavenderhaze:{ header: '#7C3AED', headerText: '#FFFFFF', accent: '#C4B5FD', border: '#DDD6FE' },
	sorbet:    { header: '#F59E0B', headerText: '#FFFFFF', accent: '#FB923C', border: '#FED7AA' },
	cloud:     { header: '#6B7280', headerText: '#FFFFFF', accent: '#9CA3AF', border: '#E5E7EB' },
	meadow:    { header: '#65A30D', headerText: '#FFFFFF', accent: '#A3E635', border: '#D9F99D' },
	sherbet:   { header: '#0D5044', headerText: '#FFFFFF', accent: '#7048A0', border: '#B8E4D8' }
};

var DISPLAY_MAP = {
	'BookId':'ID', 'DateAdded':'Date Added', 'DateStarted':'Date Started',
	'DateFinished':'Date Finished', 'CurrentPage':'Current Page',
	'SeriesNumber':'Series #', 'TbrPriority':'TBR Priority',
	'SpiceLevel':'Spice Level', 'CoverEmoji':'Cover Emoji', 'CoverUrl':'Cover URL',
	'Gradient1':'Grad 1', 'Gradient2':'Grad 2', 'AuthorKey':'Author Key',
	'ChallengeId':'ID', 'Current':'Progress', 'Target':'Goal',
	'ShelfId':'ID',
	'PhotoData':'Photo', 'YearlyGoal':'Yearly Goal', 'DemoCleared':'Demo Cleared',
	'ShowSpoilers':'Show Spoilers', 'ReadingOrder':'Reading Order',
	'RecentIds':'Recent IDs', 'SortBy':'Sort By', 'LibViewMode':'View Mode',
	'SelectedFilter':'Active Filter', 'ActiveShelf':'Active Shelf',
	'ChallengeBarCollapsed':'Goals Collapsed', 'LibToolsOpen':'Tools Open',
	'LibraryName':'Library Name',
	'CustomQuotes':'Custom Quotes', 'CoversEnabled':'Covers Enabled',
	'TutorialCompleted':'Tutorial Completed', 'LastAudioId':'Last Audiobook',
	'TotalListeningMins':'Total Listening (min)',
	'AudiobookId':'ID', 'ChapterCount':'Chapters',
	'LibrivoxProjectId':'Project ID', 'CurrentChapterIndex':'Current Chapter',
	'CurrentTime':'Position', 'PlaybackSpeed':'Speed'
};

function _displayHeaders(internalHeaders) {
	return internalHeaders.map(function(h) { return DISPLAY_MAP[h] || h; });
}

// ── Utility helpers ─────────────────────────────────────────────────────
function _uuid() { return Utilities.getUuid(); }
function _ss()   { return SpreadsheetApp.getActiveSpreadsheet(); }
function _log(level, fn, msg) { Logger.log('[' + String(level).toUpperCase() + '] ' + fn + ': ' + String(msg)); }

function _validateId(val) {
	var s = String(val || '').trim();
	return s.length >= 4 && s.length <= 200;
}

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
		// Write headers on row 1 of hidden utility sheets so the data model is stable.
		// The VISIBLE Library sheet re-arranges row 1/2 in _dbLiteInitLibrarySheet.
		if (name !== SHEET_LIBRARY) {
			sheet.getRange(1, 1, 1, displayRow.length).setValues([displayRow]);
		}
	}
	return sheet;
}

function _ensureColumns(sheet, needed) {
	var current = sheet.getMaxColumns();
	if (current < needed) sheet.insertColumnsAfter(current, needed - current);
}
function _ensureRows(sheet, needed) {
	var current = sheet.getMaxRows();
	if (current < needed) sheet.insertRowsAfter(current, needed - current);
}

function _sheetToObjects(sheet, internalHeaders) {
	if (!sheet) return [];
	var data = sheet.getDataRange().getValues();
	if (data.length < 2) return [];
	var headers = internalHeaders || data[0];
	return data.slice(1).filter(function(row) {
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
		if (String(data[r][colIndex]) === String(value)) return r + 1;
	}
	return -1;
}

// ── Library-specific read helpers ───────────────────────────────────────
// The Library sheet has a formula counter in col A. Data starts at
// LIBRARY_DATA_COL (B), header at LIBRARY_HEADER_ROW (8), data at LIBRARY_DATA_ROW (9).
function _libSheetToObjects(sheet) {
	if (!sheet) return [];
	var lastRow = sheet.getLastRow();
	if (lastRow < LIBRARY_DATA_ROW) return [];
	var numRows = lastRow - LIBRARY_DATA_ROW + 1;
	var data = sheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, numRows, LIBRARY_HEADERS.length).getValues();
	return data.filter(function(row) {
		return String(row[0] || '').trim() !== ''; // row[0] = Title
	}).map(function(row) {
		var obj = {};
		LIBRARY_HEADERS.forEach(function(h, i) { obj[h] = row[i]; });
		return obj;
	});
}

function _findLibRowByBookId(sheet, bookId) {
	if (!sheet || !bookId) return -1;
	var lastRow = sheet.getLastRow();
	if (lastRow < LIBRARY_DATA_ROW) return -1;
	var numRows = lastRow - LIBRARY_DATA_ROW + 1;
	var bookIdCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('BookId');
	var col = sheet.getRange(LIBRARY_DATA_ROW, bookIdCol, numRows, 1).getValues();
	for (var r = 0; r < col.length; r++) {
		if (String(col[r][0]) === String(bookId)) return LIBRARY_DATA_ROW + r;
	}
	return -1;
}

// Returns the next empty data row in the Library sheet.
// getLastRow() is unreliable because col A has pre-filled formulas (rows 9-5008),
// so we scan backwards through the Title column (col B) to find the last real entry.
function _nextLibDataRow(sheet) {
	var titleCol = LIBRARY_DATA_COL; // Title is index 0 → LIBRARY_DATA_COL + 0
	var maxRow = sheet.getMaxRows();
	if (maxRow < LIBRARY_DATA_ROW) return LIBRARY_DATA_ROW;
	var count = maxRow - LIBRARY_DATA_ROW + 1;
	var vals = sheet.getRange(LIBRARY_DATA_ROW, titleCol, count, 1).getValues();
	for (var i = vals.length - 1; i >= 0; i--) {
		if (String(vals[i][0] || '').trim() !== '') return LIBRARY_DATA_ROW + i + 1;
	}
	return LIBRARY_DATA_ROW;
}

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

function _getCurrentTheme() {
	var sheet = _ss().getSheetByName(SHEET_PROFILE);
	if (!sheet || sheet.getLastRow() < 2) return 'blossom';
	var themeCol = PROFILE_HEADERS.indexOf('Theme') + 1;
	return String(sheet.getRange(2, themeCol).getValue() || 'blossom').toLowerCase();
}

// Covers are NEVER rendered in the visible sheet — this is intentional for
// the template design (no forced row heights, no broken external images).
// The function is kept as a no-op so legacy callers still work.
function _writeCoverFormula(/*sheet, rowNum*/) { /* no-op */ }

// ── Public layout entry points ──────────────────────────────────────────
function initializeSheets() {
	_dbLiteInitializeSheets();
}

function _reStyleAllSheets(themeName) {
	_dbLiteInitializeSheets();
	_incrementSyncVersion();
}

// ── Status mapping ──────────────────────────────────────────────────────
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

function _safeJsonParse(str, fallback) {
	try { var parsed = JSON.parse(str); return Array.isArray(parsed) ? parsed : fallback; }
	catch (e) { return fallback; }
}
function _normalizeIsbn(isbn) { return String(isbn || '').toUpperCase().replace(/[^0-9X]/g, ''); }
function _humanizeListName(name) {
	return String(name || '').replace(/-/g, ' ').replace(/\b\w/g, function(c) { return c.toUpperCase(); }).trim();
}

// =====================================================================
//  CLIENT API — called from UI via google.script.run
// =====================================================================

// ── Rating helpers: sheet stores ★ strings, webapp stores 0-5 numbers ──
function _starsToNum(v) {
	var s = String(v || '').trim();
	if (s === '★★★★★') return 5;
	if (s === '★★★★')  return 4;
	if (s === '★★★')   return 3;
	if (s === '★★')    return 2;
	if (s === '★')     return 1;
	var n = Number(s);
	return (n >= 0 && n <= 5) ? Math.round(n) : 0;
}
function _numToStars(n) {
	var num = Math.max(0, Math.min(5, Math.round(Number(n) || 0)));
	if (num === 0) return '';
	return new Array(num + 1).join('★');
}

function clientGetInitialData() {
	var _profileCheck = _ss().getSheetByName(SHEET_PROFILE);
	if (!_profileCheck || _profileCheck.getLastRow() < 2) {
		initializeSheets();
	}

	var libSheet   = _ss().getSheetByName(SHEET_LIBRARY);
	var chalSheet  = _ss().getSheetByName(SHEET_CHALLENGES);
	var shelfSheet = _ss().getSheetByName(SHEET_SHELVES);
	var profileSheet = _ss().getSheetByName(SHEET_PROFILE);
	var audioSheet = _ss().getSheetByName(SHEET_AUDIOBOOKS);

	var library = _libSheetToObjects(libSheet).map(function(row) {
		var rawIsbn      = String(row.ISBN || '').trim();
		var isbnNorm     = _normalizeIsbn(rawIsbn);
		var coverPrimary = String(row.CoverUrl || '').trim();
		var isbnCoverUrl = isbnNorm ? ('https://covers.openlibrary.org/b/isbn/' + isbnNorm + '-L.jpg') : '';
		var coverFallback = (isbnCoverUrl && coverPrimary !== isbnCoverUrl) ? isbnCoverUrl : '';
		function _fmtDate(v) {
			if (!v) return '';
			try { return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch(e) { return ''; }
		}
		return {
			BookId: row.BookId, Title: row.Title, Author: row.Author, Status: row.Status,
			Rating: _starsToNum(row.Rating),
			PageCount: Number(row.Pages) || 0,
			Genres: row.Genre || '',
			DateAdded:    _fmtDate(row.DateAdded),
			DateStarted:  _fmtDate(row.DateStarted),
			DateFinished: _fmtDate(row.DateFinished),
			CurrentPage: Number(row.CurrentPage) || 0,
			Series: row.Series || '', SeriesOrder: row.SeriesNumber || '',
			TbrPriority: row.TbrPriority || '',
			Format: row.Format || '', Source: row.Source || '',
			SpiceLevel: Number(row.SpiceLevel) || 0,
			Moods: row.Tags || '', Shelves: row.Shelves || '',
			Notes: row.Notes || '', Review: row.Review || '', Quotes: row.Quotes || '',
			Favorite: row.Favorite === true || String(row.Favorite).toUpperCase() === 'TRUE',
			CoverEmoji: row.CoverEmoji || '',
			CoverUrlPrimary: coverPrimary, CoverUrlFallback: coverFallback,
			Gradient1: row.Gradient1 || '', Gradient2: row.Gradient2 || '',
			ISBN: isbnNorm, OLID: row.OLID || '', AuthorKey: row.AuthorKey || ''
		};
	});

	var nytBundle = clientGetNYTBadgesForLibrary();
	var nytFeedBundle = clientGetNYTFeed();

	var goals = _sheetToObjects(chalSheet, CHALLENGE_HEADERS).map(function(row) {
		return {
			GoalId: row.ChallengeId, GoalType: row.Name, Icon: row.Icon || 'GOAL',
			CurrentValue: Number(row.Current) || 0, TargetValue: Number(row.Target) || 1
		};
	});

	var shelves = _sheetToObjects(shelfSheet, SHELF_HEADERS).map(function(row) {
		return { ShelfId: row.ShelfId, ShelfName: row.Name, Icon: row.Icon || '' };
	});

	var profileData = {};
	if (profileSheet && profileSheet.getLastRow() >= 2) {
		var pRow = profileSheet.getRange(2, 1, 1, PROFILE_HEADERS.length).getValues()[0];
		PROFILE_HEADERS.forEach(function(h, i) { profileData[h] = pRow[i]; });
	}

	var settings = {
		Theme: String(profileData.Theme || 'blossom'),
		ShowSpoilers: String(profileData.ShowSpoilers || 'false')
	};

	var audiobooks = _sheetToObjects(audioSheet, AUDIOBOOK_HEADERS).map(function(row) {
		return {
			AudiobookId: row.AudiobookId, Title: row.Title, Author: row.Author,
			Duration: row.Duration || '', CoverEmoji: row.CoverEmoji || 'AUDIO',
			CoverUrl: row.CoverUrl || '', ChapterCount: Number(row.ChapterCount) || 0,
			LibrivoxProjectId: row.LibrivoxProjectId || '',
			CurrentChapterIndex: Number(row.CurrentChapterIndex) || 0,
			CurrentTime: Number(row.CurrentTime) || 0,
			PlaybackSpeed: Number(row.PlaybackSpeed) || 1,
			TotalListeningMins: Number(row.TotalListeningMins) || 0
		};
	});

	return {
		library: library, goals: goals, shelves: shelves, settings: settings,
		profile: {
			name: String(profileData.Name || ''),
			motto: String(profileData.Motto || 'A focused place to track every book'),
			photoData: String(profileData.PhotoData || '')
		},
		yearlyGoal: Number(profileData.YearlyGoal) || 50,
		readingOrder: _safeJsonParse(profileData.ReadingOrder, []),
		recentIds: _safeJsonParse(profileData.RecentIds, []),
		sortBy: String(profileData.SortBy || 'default'),
		libViewMode: String(profileData.LibViewMode || 'grid'),
		onboarded: String(profileData.Onboarded) === 'true' || profileData.Onboarded === true,
		demoCleared: String(profileData.DemoCleared) === 'true' || profileData.DemoCleared === true,
		selectedFilter: String(profileData.SelectedFilter || 'all'),
		activeShelf: String(profileData.ActiveShelf || ''),
		challengeBarCollapsed: String(profileData.ChallengeBarCollapsed) === 'true' || profileData.ChallengeBarCollapsed === true,
		libToolsOpen: String(profileData.LibToolsOpen) === 'true' || profileData.LibToolsOpen === true,
		libraryName: String(profileData.LibraryName || ''),
		customQuotes: _safeJsonParse(profileData.CustomQuotes, []),
		coversEnabled: !(String(profileData.CoversEnabled) === 'false' || profileData.CoversEnabled === false),
		tutorialCompleted: String(profileData.TutorialCompleted) === 'true' || profileData.TutorialCompleted === true,
		lastAudioId: String(profileData.LastAudioId || ''),
		totalListeningMins: Number(profileData.TotalListeningMins) || 0,
		audiobooks: audiobooks,
		nytBadges: nytBundle.byBookId || {},
		nytBadgesByIsbn: nytBundle.byIsbn || {},
		nytFeed: nytFeedBundle.lists || [],
		nytCacheDate: nytFeedBundle.updatedAt || PropertiesService.getScriptProperties().getProperty('NYT_CACHE_DATE') || ''
	};
}

// ── Books ────────────────────────────────────────────────────────────────
function clientAddBook(payload) {
	try {
		var sheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
		var bookId = _uuid();
		var row = _bookPayloadToRow(bookId, payload);
		// Append to first empty data row — _nextLibDataRow scans Title column since
		// getLastRow() is inflated by the 5000 pre-filled col-A formulas.
		var nextRow = _nextLibDataRow(sheet);
		sheet.getRange(nextRow, LIBRARY_DATA_COL, 1, LIBRARY_HEADERS.length).setValues([row]);
		return { BookId: bookId };
	} catch (e) { return { error: e.message }; }
}

function clientUpdateBook(bookId, updates) {
	try {
		if (!_validateId(bookId)) return { error: 'Invalid book ID.' };
		if (!updates || typeof updates !== 'object') return { error: 'Updates object required.' };
		var sheet = _ss().getSheetByName(SHEET_LIBRARY);
		if (!sheet) return { error: 'Library sheet not found' };
		var rowIdx = _findLibRowByBookId(sheet, bookId);
		if (rowIdx < 0) return { error: 'Book not found' };
		var dataRow = sheet.getRange(rowIdx, LIBRARY_DATA_COL, 1, LIBRARY_HEADERS.length).getValues()[0];
		var keyMap = {
			'Title':'Title','Author':'Author','Status':'Status','Rating':'Rating',
			'Pages':'Pages','PageCount':'Pages','Genre':'Genre','Genres':'Genre',
			'DateAdded':'DateAdded','DateStarted':'DateStarted','DateFinished':'DateFinished',
			'CurrentPage':'CurrentPage','Series':'Series','SeriesOrder':'SeriesNumber',
			'SeriesNumber':'SeriesNumber','TbrPriority':'TbrPriority','Format':'Format',
			'Source':'Source','SpiceLevel':'SpiceLevel','Tags':'Tags','Moods':'Tags',
			'Shelves':'Shelves','Notes':'Notes','Review':'Review','Quotes':'Quotes',
			'Favorite':'Favorite','CoverEmoji':'CoverEmoji','CoverUrl':'CoverUrl',
			'CoverUrlPrimary':'CoverUrl','Gradient1':'Gradient1','Gradient2':'Gradient2',
			'ISBN':'ISBN','OLID':'OLID','AuthorKey':'AuthorKey'
		};
		Object.keys(updates).forEach(function(k) {
			var colName = keyMap[k]; if (!colName) return;
			var colIdx = LIBRARY_HEADERS.indexOf(colName); if (colIdx < 0) return;
			var val = updates[k];
			if (colName === 'Status') val = _uiStatusToSheet(val);
			if (colName === 'Rating') val = _numToStars(val);
			dataRow[colIdx] = val;
		});
		sheet.getRange(rowIdx, LIBRARY_DATA_COL, 1, LIBRARY_HEADERS.length).setValues([dataRow]);
		return { success: true };
	} catch (e) { return { error: e.message }; }
}

function clientDeleteBook(bookId) {
	try {
		if (!_validateId(bookId)) return;
		var sheet = _ss().getSheetByName(SHEET_LIBRARY);
		if (!sheet) return;
		var rowIdx = _findLibRowByBookId(sheet, bookId);
		if (rowIdx >= LIBRARY_DATA_ROW) sheet.deleteRow(rowIdx);
	} catch (e) { return { error: e.message }; }
}

function _bookPayloadToRow(bookId, p) {
	function _cap(v, n) { return String(v || '').slice(0, n); }
	// Order must match LIBRARY_HEADERS exactly
	return [
		_cap(p.Title, 500),
		_cap(p.Author, 300),
		p.Status ? _uiStatusToSheet(p.Status) : 'Want to Read',
		_cap(p.Genres || p.Genre, 200),
		_numToStars(p.Rating),
		_cap(p.Format, 100),
		Number(p.PageCount || p.Pages) || 0,
		p.DateStarted || '',
		p.DateFinished || '',
		_cap(p.Series, 200),
		p.SeriesOrder || p.SeriesNumber || '',
		p.Favorite === true || p.Favorite === 'true',
		// Hidden columns
		bookId,
		_cap(p.CoverUrlPrimary || p.CoverUrl, 2000),
		p.CoverEmoji || '',
		_cap(p.Gradient1, 50),
		_cap(p.Gradient2, 50),
		p.DateAdded || new Date().toISOString().slice(0, 10),
		Number(p.CurrentPage) || 0,
		p.TbrPriority || '',
		_cap(p.Source, 100),
		Number(p.SpiceLevel) || 0,
		_cap(p.Moods || p.Tags, 500),
		_cap(p.Shelves, 500),
		_cap(p.Notes, 5000),
		_cap(p.Review, 5000),
		_cap(p.Quotes, 5000),
		_cap(p.ISBN, 20),
		_cap(p.OLID, 50),
		_cap(p.AuthorKey, 100)
	];
}

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
			var title = d[findCol('Title')] || d[0] || '';
			var author = d[findCol('Author')] || d[1] || '';
			if (!title) continue;
			newRows.push(_bookPayloadToRow(_uuid(), {
				Title: title, Author: author,
				Status: d[findCol('Status')] || (findCol('Exclusive Shelf') > -1 ? d[findCol('Exclusive Shelf')] : '') || 'Want to Read',
				Rating: Number(d[findCol('Rating')] || d[findCol('My Rating')] || 0),
				PageCount: Number(d[findCol('PageCount')] || d[findCol('Pages')] || d[findCol('Number of Pages')] || 0),
				Genre: d[findCol('Genre')] || d[findCol('Genres')] || '',
				DateAdded: d[findCol('DateAdded')] || d[findCol('Date Read')] || new Date().toISOString().slice(0, 10),
				DateFinished: d[findCol('DateFinished')] || '',
				Review: d[findCol('Review')] || '', Notes: d[findCol('Notes')] || '',
				ISBN: d[findCol('ISBN')] || ''
			}));
		}
		if (newRows.length > 0) {
			var nextRow = _nextLibDataRow(sheet);
			sheet.getRange(nextRow, LIBRARY_DATA_COL, newRows.length, LIBRARY_HEADERS.length).setValues(newRows);
		}
		return { imported: newRows.length };
	} catch (e) { return { error: e.message, imported: 0 }; }
}

function clientImportStoryGraphCSV(rows) {
	// StoryGraph CSVs use the same basic shape — delegate to Goodreads importer
	// which already accepts header name variants.
	return clientImportGoodreadsCSV(rows);
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
	} catch (e) { return { error: e.message }; }
}
function clientDeleteShelf(shelfId) {
	try {
		if (!_validateId(shelfId)) return;
		var sheet = _ss().getSheetByName(SHEET_SHELVES); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, shelfId);
		if (rowIdx >= 2) sheet.deleteRow(rowIdx);
	} catch (e) {}
}
function clientRenameShelf(shelfId, newName) {
	try {
		if (!_validateId(shelfId)) return;
		var sheet = _ss().getSheetByName(SHEET_SHELVES); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, shelfId);
		if (rowIdx >= 2) sheet.getRange(rowIdx, 2).setValue(String(newName || '').slice(0, 200));
	} catch (e) {}
}
function clientUpdateShelf(shelfId, updates) {
	try {
		if (!updates || !_validateId(shelfId)) return;
		var sheet = _ss().getSheetByName(SHEET_SHELVES); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, shelfId);
		if (rowIdx < 2) return;
		var newName = updates.Name !== undefined ? updates.Name : updates.name;
		if (newName !== undefined) sheet.getRange(rowIdx, 2).setValue(String(newName).slice(0, 200));
		var newIcon = updates.Icon !== undefined ? updates.Icon : updates.icon;
		if (newIcon !== undefined) sheet.getRange(rowIdx, 3).setValue(String(newIcon).slice(0, 50));
	} catch (e) {}
}

// ── Challenges / Goals ──────────────────────────────────────────────────
function clientAddChallenge(payload) {
	try {
		if (!payload) return { error: 'Payload required.' };
		var sheet = _getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
		var id = _uuid();
		sheet.appendRow([id, String(payload.name || 'New Challenge').slice(0, 200),
			String(payload.icon || 'GOAL').slice(0, 50),
			Number(payload.current) || 0, Number(payload.target) || 10]);
		return { ChallengeId: id };
	} catch (e) { return { error: e.message }; }
}
function clientUpdateChallenge(challengeId, updates) {
	try {
		if (!_validateId(challengeId)) return;
		var sheet = _ss().getSheetByName(SHEET_CHALLENGES); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, challengeId); if (rowIdx < 2) return;
		var row = sheet.getRange(rowIdx, 1, 1, CHALLENGE_HEADERS.length).getValues()[0];
		if (updates.name !== undefined)    row[1] = String(updates.name).slice(0, 200);
		if (updates.icon !== undefined)    row[2] = String(updates.icon).slice(0, 50);
		if (updates.current !== undefined) row[3] = Number(updates.current);
		if (updates.target !== undefined)  row[4] = Number(updates.target);
		sheet.getRange(rowIdx, 1, 1, CHALLENGE_HEADERS.length).setValues([row]);
	} catch (e) {}
}
function clientDeleteChallenge(challengeId) {
	try {
		if (!_validateId(challengeId)) return;
		var sheet = _ss().getSheetByName(SHEET_CHALLENGES); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, challengeId);
		if (rowIdx >= 2) sheet.deleteRow(rowIdx);
	} catch (e) {}
}
function clientSyncChallenges(challengesArray) {
	try {
		if (!challengesArray || challengesArray.length === 0) return;
		var sheet = _getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
		var rows = challengesArray.map(function(c) {
			return [
				c._serverChallengeId || _uuid(),
				String(c.name || '').slice(0, 200),
				String(c.icon || 'GOAL').slice(0, 50),
				Number(c.current) || 0, Number(c.target) || 1
			];
		});
		if (sheet.getLastRow() > 1) {
			sheet.getRange(2, 1, sheet.getLastRow() - 1, CHALLENGE_HEADERS.length).clearContent();
		}
		sheet.getRange(2, 1, rows.length, CHALLENGE_HEADERS.length).setValues(rows);
	} catch (e) {}
}

// ── Settings / Profile ──────────────────────────────────────────────────
function clientSetSetting(key, value) {
	try {
		var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
		if (sheet.getLastRow() < 2) _dbLiteEnsureProfileDefaults(sheet);
		var colIdx = PROFILE_HEADERS.indexOf(key);
		if (colIdx < 0) return;
		var safeValue = (key === 'Theme') ? _validateTheme(value) : value;
		sheet.getRange(2, colIdx + 1).setValue(safeValue);
		if (key === 'Theme') _reStyleAllSheets(safeValue);
	} catch (e) {}
}
function clientSetSettings(settingsObj) {
	try {
		if (!settingsObj || typeof settingsObj !== 'object') return;
		Object.keys(settingsObj).forEach(function(k) { clientSetSetting(k, settingsObj[k]); });
	} catch (e) {}
}

function clientSaveProfile(profileData) {
	try {
		var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
		if (sheet.getLastRow() < 2) _dbLiteEnsureProfileDefaults(sheet);
		var mapping = { 'name':'Name', 'motto':'Motto', 'photoData':'PhotoData' };
		Object.keys(mapping).forEach(function(k) {
			if (profileData[k] !== undefined) {
				var colIdx = PROFILE_HEADERS.indexOf(mapping[k]);
				if (colIdx >= 0) {
					var val = (k === 'photoData')
						? String(profileData[k] || '').slice(0, 49000)
						: String(profileData[k] || '').slice(0, 500);
					sheet.getRange(2, colIdx + 1).setValue(val);
				}
			}
		});
		var name = String(profileData.name || '').replace(/[\x00-\x1F\x7F]/g, '').trim().slice(0, 100);
		if (name) {
			var possessive = name.slice(-1) === 's' ? name + "'" : name + "'s";
			_ss().rename(possessive + ' Reading Journey');
		} else {
			_ss().rename('My Reading Journey');
		}
	} catch (e) {}
}

function clientSaveYearlyGoal(goal) {
	try {
		var n = Number(goal);
		if (!isFinite(n) || n < 1 || n > 10000) return;
		clientSetSetting('YearlyGoal', n);
	} catch (e) {}
}
function clientSaveReadingOrder(orderArray) {
	try {
		if (!Array.isArray(orderArray)) return;
		var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
		if (sheet.getLastRow() < 2) _dbLiteEnsureProfileDefaults(sheet);
		var colIdx = PROFILE_HEADERS.indexOf('ReadingOrder');
		if (colIdx >= 0) sheet.getRange(2, colIdx + 1).setValue(JSON.stringify(orderArray));
	} catch (e) {}
}
function clientSaveRecentIds(idsArray) {
	try {
		if (!Array.isArray(idsArray)) return;
		var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
		if (sheet.getLastRow() < 2) _dbLiteEnsureProfileDefaults(sheet);
		var colIdx = PROFILE_HEADERS.indexOf('RecentIds');
		if (colIdx >= 0) sheet.getRange(2, colIdx + 1).setValue(JSON.stringify(idsArray));
	} catch (e) {}
}
function clientSaveUiPrefs(prefs) {
	try {
		if (!prefs || typeof prefs !== 'object') return;
		if (prefs.sortBy !== undefined)                clientSetSetting('SortBy', prefs.sortBy);
		if (prefs.libViewMode !== undefined)           clientSetSetting('LibViewMode', prefs.libViewMode);
		if (prefs.onboarded !== undefined)             clientSetSetting('Onboarded', prefs.onboarded);
		if (prefs.demoCleared !== undefined)           clientSetSetting('DemoCleared', prefs.demoCleared);
		if (prefs.selectedFilter !== undefined)        clientSetSetting('SelectedFilter', prefs.selectedFilter);
		if (prefs.activeShelf !== undefined)           clientSetSetting('ActiveShelf', prefs.activeShelf);
		if (prefs.challengeBarCollapsed !== undefined) clientSetSetting('ChallengeBarCollapsed', prefs.challengeBarCollapsed);
		if (prefs.libToolsOpen !== undefined)          clientSetSetting('LibToolsOpen', prefs.libToolsOpen);
		if (prefs.libraryName !== undefined)           clientSetSetting('LibraryName', prefs.libraryName);
		if (prefs.customQuotes !== undefined) {
			var cq = prefs.customQuotes;
			clientSetSetting('CustomQuotes', typeof cq === 'string' ? cq : JSON.stringify(cq || []));
		}
		if (prefs.coversEnabled !== undefined)      clientSetSetting('CoversEnabled', !!prefs.coversEnabled);
		if (prefs.tutorialCompleted !== undefined)  clientSetSetting('TutorialCompleted', !!prefs.tutorialCompleted);
		if (prefs.lastAudioId !== undefined)        clientSetSetting('LastAudioId', String(prefs.lastAudioId || ''));
		if (prefs.totalListeningMins !== undefined) clientSetSetting('TotalListeningMins', Number(prefs.totalListeningMins) || 0);
	} catch (e) {}
}

// ── Audiobooks ──────────────────────────────────────────────────────────
function clientSaveAudiobook(audioData) {
	try {
		var sheet = _getOrCreateSheet(SHEET_AUDIOBOOKS, AUDIOBOOK_HEADERS);
		var existingRow = _findRowByCol(sheet, 0, audioData.id);
		var row = [
			audioData.id || _uuid(), audioData.title || '', audioData.author || '',
			audioData.duration || '', audioData.cover || 'AUDIO', audioData.coverUrl || '',
			Number(audioData.chapterCount) || 0, audioData.audiobookId || '',
			Number(audioData.chapterIndex) || 0, Number(audioData.currentTime) || 0,
			Number(audioData.speed) || 1, Number(audioData.totalListeningMins) || 0
		];
		if (existingRow > 1) {
			sheet.getRange(existingRow, 1, 1, AUDIOBOOK_HEADERS.length).setValues([row]);
		} else {
			sheet.appendRow(row);
		}
	} catch (e) { return { error: e.message }; }
}
function clientSaveAudioPosition(audioId, chapterIndex, currentTime, speed, totalListeningMins) {
	try {
		if (!_validateId(audioId)) return;
		var sheet = _ss().getSheetByName(SHEET_AUDIOBOOKS); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, String(audioId)); if (rowIdx < 2) return;
		var row = sheet.getRange(rowIdx, 1, 1, AUDIOBOOK_HEADERS.length).getValues()[0];
		row[AUDIOBOOK_HEADERS.indexOf('CurrentChapterIndex')] = Number(chapterIndex) || 0;
		row[AUDIOBOOK_HEADERS.indexOf('CurrentTime')]         = Number(currentTime) || 0;
		row[AUDIOBOOK_HEADERS.indexOf('PlaybackSpeed')]       = Number(speed) || 1;
		if (totalListeningMins !== undefined && totalListeningMins !== null) {
			row[AUDIOBOOK_HEADERS.indexOf('TotalListeningMins')] = Number(totalListeningMins) || 0;
		}
		sheet.getRange(rowIdx, 1, 1, AUDIOBOOK_HEADERS.length).setValues([row]);
	} catch (e) {}
}

// ── Demo Data ───────────────────────────────────────────────────────────
function _clearSheetDataRows(sheet, headers) {
	if (!sheet || sheet.getLastRow() < 2) return;
	sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
}

function _seedDemoData() {
	// Respect the "user has cleared demo data" flag so we never resurrect
	// sample books after a user has explicitly emptied their library.
	if (PropertiesService.getScriptProperties().getProperty('DEMO_CLEARED') === '1') return;

	var libSheet = _ss().getSheetByName(SHEET_LIBRARY);
	if (!libSheet) libSheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);

	// Exit if any real book row already exists (Title column = col B = LIBRARY_DATA_COL).
	var lastRow = libSheet.getLastRow();
	if (lastRow >= LIBRARY_DATA_ROW) {
		var titleVals = libSheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, lastRow - LIBRARY_DATA_ROW + 1, 1).getValues();
		for (var i = 0; i < titleVals.length; i++) {
			if (String(titleVals[i][0] || '').trim() !== '') return;
		}
	}

	var now = new Date();
	function _monthDate(monthsAgo, day) {
		var d = new Date(now.getFullYear(), now.getMonth() - monthsAgo, day);
		return d.toISOString().slice(0, 10);
	}
	function _weeksAgo(w, dayOff) {
		var d = new Date(now); d.setDate(now.getDate() - w * 7 + (dayOff || 0));
		return d.toISOString().slice(0, 10);
	}

	var demoBooks = [
		{ t:'Beach Read', a:'Emily Henry', g:'Romance', isbn:'9781984806734', pg:352, r:5, stat:'Finished', da:_weeksAgo(11,2), df:_monthDate(0,2), fmt:'Paperback' },
		{ t:'Circe', a:'Madeline Miller', g:'Fantasy', isbn:'9780316556347', pg:393, r:5, stat:'Finished', da:_weeksAgo(10,0), df:_monthDate(2,14), fmt:'Hardcover' },
		{ t:'The Silent Patient', a:'Alex Michaelides', g:'Thriller', isbn:'9781250301697', pg:325, r:4, stat:'Finished', da:_weeksAgo(9,4), df:_monthDate(1,3), fmt:'Paperback' },
		{ t:'Atomic Habits', a:'James Clear', g:'Self-Help', isbn:'9780735211292', pg:306, r:5, stat:'Finished', da:_weeksAgo(8,1), df:_monthDate(5,22), fmt:'Audiobook' },
		{ t:'The Song of Achilles', a:'Madeline Miller', g:'Fantasy', isbn:'9780062060624', pg:352, r:5, stat:'Finished', da:_weeksAgo(0,0), df:_monthDate(2,27), fmt:'Paperback' },
		{ t:'Where the Crawdads Sing', a:'Delia Owens', g:'Mystery', isbn:'9780735224292', pg:368, r:5, stat:'Finished', da:_weeksAgo(6,0), df:_monthDate(1,27), fmt:'Hardcover' },
		{ t:'Project Hail Mary', a:'Andy Weir', g:'SciFi', isbn:'9780593135204', pg:476, r:5, stat:'Finished', da:_weeksAgo(4,5), df:_monthDate(1,20), fmt:'Ebook' },
		{ t:'The Guest List', a:'Lucy Foley', g:'Mystery', isbn:'9780062868930', pg:312, r:4, stat:'Finished', da:_weeksAgo(8,5), df:_monthDate(4,24), fmt:'Paperback' },
		{ t:'Educated', a:'Tara Westover', g:'Memoir', isbn:'9780399590504', pg:334, r:5, stat:'Finished', da:_weeksAgo(3,3), df:_monthDate(1,8), fmt:'Hardcover' },
		{ t:'The Invisible Life of Addie LaRue', a:'V.E. Schwab', g:'Fantasy', isbn:'9780765387561', pg:448, r:5, stat:'Finished', da:_weeksAgo(1,4), df:_monthDate(2,25), fmt:'Paperback' },
		{ t:'The Vanishing Half', a:'Brit Bennett', g:'Fiction', isbn:'9780525536291', pg:343, r:4, stat:'Finished', da:_weeksAgo(4,1), df:_monthDate(1,14), fmt:'Hardcover' },
		{ t:'Verity', a:'Colleen Hoover', g:'Thriller', isbn:'9781538724736', pg:374, r:5, stat:'Finished', da:_weeksAgo(7,3), df:_monthDate(4,3), fmt:'Paperback' },
		{ t:'Book Lovers', a:'Emily Henry', g:'Romance', isbn:'9780593334836', pg:368, r:5, stat:'Finished', da:_weeksAgo(3,0), df:_monthDate(2,4), fmt:'Paperback' },
		{ t:'The Spanish Love Deception', a:'Elena Armas', g:'Romance', isbn:'9781982177010', pg:358, r:4, stat:'Finished', da:_weeksAgo(5,2), df:_monthDate(1,10), fmt:'Ebook' },
		{ t:'A Court of Thorns and Roses', a:'Sarah J. Maas', g:'Fantasy', isbn:'9781635575569', pg:419, r:5, stat:'Finished', da:_weeksAgo(6,4), df:_monthDate(3,19), fmt:'Paperback' },
		{ t:'The Thursday Murder Club', a:'Richard Osman', g:'Mystery', isbn:'9781984880963', pg:369, r:4, stat:'Finished', da:_weeksAgo(2,5), df:_monthDate(2,20), fmt:'Hardcover' },
		{ t:'The Four Winds', a:'Kristin Hannah', g:'Historical', isbn:'9781250178602', pg:454, r:5, stat:'Finished', da:_weeksAgo(2,2), df:_monthDate(5,8), fmt:'Hardcover' },
		{ t:'Normal People', a:'Sally Rooney', g:'Romance', isbn:'9781984822185', pg:266, r:4, stat:'Finished', da:_weeksAgo(9,4), df:_monthDate(5,15), fmt:'Paperback' },
		{ t:'The House in the Cerulean Sea', a:'TJ Klune', g:'Fantasy', isbn:'9781250217288', pg:396, r:5, stat:'Finished', da:_weeksAgo(6,0), df:_monthDate(4,29), fmt:'Paperback' },
		{ t:'Malibu Rising', a:'Taylor Jenkins Reid', g:'Fiction', isbn:'9780593158203', pg:369, r:5, stat:'Finished', da:_weeksAgo(0,4), df:_monthDate(3,5), fmt:'Hardcover' },
		{ t:'The Love Hypothesis', a:'Ali Hazelwood', g:'Romance', isbn:'9780593336823', pg:357, r:4, stat:'Finished', da:_weeksAgo(2,2), df:_monthDate(1,25), fmt:'Paperback' },
		{ t:'Daisy Jones & The Six', a:'Taylor Jenkins Reid', g:'Fiction', isbn:'9781524798628', pg:368, r:5, stat:'Finished', da:_weeksAgo(1,1), df:_monthDate(3,12), fmt:'Audiobook' },
		{ t:'The Atlas Six', a:'Olivie Blake', g:'Fantasy', isbn:'9781250854513', pg:374, r:4, stat:'Finished', da:_weeksAgo(3,3), df:_monthDate(3,26), fmt:'Paperback' },
		{ t:'Red, White & Royal Blue', a:'Casey McQuiston', g:'Romance', isbn:'9781250316776', pg:352, r:5, stat:'Finished', da:_weeksAgo(0,2), df:_monthDate(2,20), fmt:'Paperback' },
		{ t:'It Ends With Us', a:'Colleen Hoover', g:'Romance', isbn:'9781501110375', pg:376, r:5, stat:'Reading', da:_weeksAgo(0,0), ds:'2026-03-22', cp:169, fmt:'Paperback' },
		{ t:'The Midnight Library', a:'Matt Haig', g:'Fiction', isbn:'9780525559474', pg:304, r:4, stat:'Reading', da:_weeksAgo(0,2), ds:'2026-03-01', cp:249, fmt:'Hardcover' }
	];

	var readingIds = [];
	var rows = demoBooks.map(function(b) {
		var bookId = _uuid();
		if (b.stat === 'Reading') readingIds.push(bookId);
		// Array order must match LIBRARY_HEADERS exactly:
		// Title, Author, Status, Genre, Rating, Format, Pages, DateStarted, DateFinished,
		// Series, SeriesNumber, Favorite, BookId, CoverUrl, CoverEmoji, Gradient1, Gradient2,
		// DateAdded, CurrentPage, TbrPriority, Source, SpiceLevel, Tags, Shelves, Notes,
		// Review, Quotes, ISBN, OLID, AuthorKey
		return [
			b.t,                          // Title
			b.a,                          // Author
			b.stat,                       // Status
			b.g,                          // Genre
			_numToStars(b.r),             // Rating (★ chip string)
			b.fmt || 'Paperback',         // Format
			b.pg,                         // Pages
			b.ds || '',                   // DateStarted
			b.df || '',                   // DateFinished
			'',                           // Series
			'',                           // SeriesNumber
			b.r === 5,                    // Favorite (checkbox — 5-star books start favorited)
			bookId,                       // BookId (hidden)
			'https://covers.openlibrary.org/b/isbn/' + b.isbn + '-L.jpg', // CoverUrl (hidden)
			'BK',                         // CoverEmoji (hidden)
			'',                           // Gradient1 (hidden)
			'',                           // Gradient2 (hidden)
			b.da || new Date().toISOString().slice(0, 10), // DateAdded (hidden)
			b.cp || 0,                    // CurrentPage (hidden)
			'',                           // TbrPriority (hidden)
			'',                           // Source (hidden)
			0,                            // SpiceLevel (hidden)
			'',                           // Tags (hidden)
			'',                           // Shelves (hidden)
			'',                           // Notes (hidden)
			'',                           // Review (hidden)
			'',                           // Quotes (hidden)
			b.isbn,                       // ISBN (hidden)
			'',                           // OLID (hidden)
			''                            // AuthorKey (hidden)
		];
	});

	if (rows.length > 0) {
		// Write to LIBRARY_DATA_ROW (row 9), starting at LIBRARY_DATA_COL (col B = 2).
		// Col A has pre-filled =IF() formulas and is not touched here.
		libSheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, rows.length, LIBRARY_HEADERS.length).setValues(rows);
	}

	// Seed goals
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

	// Reading order → profile
	var profileSheet = _ss().getSheetByName(SHEET_PROFILE);
	if (profileSheet && profileSheet.getLastRow() >= 2 && readingIds.length > 0) {
		var roCol = PROFILE_HEADERS.indexOf('ReadingOrder') + 1;
		if (roCol > 0) profileSheet.getRange(2, roCol).setValue(JSON.stringify(readingIds));
	}
}

function clientClearDemoData() {
	try {
		// If called from a menu click (has UI), require confirmation. The
		// webapp calls this same function but its own confirm dialog already
		// ran, so a second confirm won't appear there (getUi() throws in
		// webapp context and we fall through).
		try {
			var ui = SpreadsheetApp.getUi();
			var resp = ui.alert(
				'Clear all books & data?',
				'This removes every book, challenge, shelf, and reading progress entry from the sheet. This cannot be undone.',
				ui.ButtonSet.OK_CANCEL
			);
			if (resp !== ui.Button.OK) return { cleared: false };
		} catch (uiErr) { /* no UI (webapp / trigger context) — proceed */ }

		// Flag so onOpen auto-init never re-seeds sample books after a clear.
		PropertiesService.getScriptProperties().setProperty('DEMO_CLEARED', '1');

		var ss = _ss();
		// Library: clear data rows B-onward (col A has pre-filled formulas — leave them intact)
		var libSheet = ss.getSheetByName(SHEET_LIBRARY);
		if (libSheet && libSheet.getLastRow() >= LIBRARY_DATA_ROW) {
			libSheet.getRange(
				LIBRARY_DATA_ROW, LIBRARY_DATA_COL,
				libSheet.getLastRow() - LIBRARY_DATA_ROW + 1,
				LIBRARY_HEADERS.length
			).clearContent();
		}
		_clearSheetDataRows(ss.getSheetByName(SHEET_CHALLENGES), CHALLENGE_HEADERS);
		_clearSheetDataRows(ss.getSheetByName(SHEET_SHELVES), SHELF_HEADERS);
		_clearSheetDataRows(ss.getSheetByName(SHEET_AUDIOBOOKS), AUDIOBOOK_HEADERS);
		var profileSheet = ss.getSheetByName(SHEET_PROFILE);
		if (profileSheet && profileSheet.getLastRow() >= 2) {
			var resets = {
				ReadingOrder: '[]', RecentIds: '[]', SelectedFilter: 'all',
				ActiveShelf: '', SortBy: 'default', LibViewMode: 'grid',
				ChallengeBarCollapsed: false, LibToolsOpen: false
			};
			var pRow = profileSheet.getRange(2, 1, 1, PROFILE_HEADERS.length).getValues()[0];
			Object.keys(resets).forEach(function(k) {
				var c = PROFILE_HEADERS.indexOf(k); if (c >= 0) pRow[c] = resets[k];
			});
			profileSheet.getRange(2, 1, 1, PROFILE_HEADERS.length).setValues([pRow]);
		}
		return { cleared: true };
	} catch (e) { return { error: e.message }; }
}

// =====================================================================
//  EXTERNAL API PROXIES — Open Library, LibriVox, Archive.org, Podcast
//  Proxied through Apps Script to avoid iframe CORS restrictions.
// =====================================================================

function clientSearchBooks(query) {
	if (!query) return [];
	var url = 'https://openlibrary.org/search.json?q=' + encodeURIComponent(query) +
		'&limit=15&fields=key,title,author_name,author_key,first_publish_year,isbn,cover_i,number_of_pages_median,subject';
	try {
		var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
		var json = JSON.parse(resp.getContentText());
		return (json.docs || []).map(function(doc) {
			var candidates = [];
			if (doc.isbn) {
				for (var i = 0; i < doc.isbn.length; i++) {
					var c = _normalizeIsbn(doc.isbn[i]);
					if (c && candidates.indexOf(c) === -1) candidates.push(c);
				}
			}
			candidates = candidates.slice(0, 10);
			var isbn = '';
			for (var j = 0; j < candidates.length; j++) {
				if (candidates[j].length === 13) { isbn = candidates[j]; break; }
			}
			if (!isbn && candidates.length) isbn = candidates[0];
			var coverId = doc.cover_i || '';
			return {
				title: doc.title || '',
				author: (doc.author_name || [])[0] || '',
				authorKey: (doc.author_key || [])[0] || '',
				year: doc.first_publish_year || '',
				isbn: isbn, isbnCandidates: candidates, isbns: candidates,
				olid: doc.key ? doc.key.replace('/works/', '') : '',
				coverId: coverId,
				coverUrlPrimary: isbn ? ('https://covers.openlibrary.org/b/isbn/' + isbn + '-L.jpg')
					: (coverId ? ('https://covers.openlibrary.org/b/id/' + coverId + '-L.jpg') : ''),
				coverUrlFallback: coverId ? ('https://covers.openlibrary.org/b/id/' + coverId + '-L.jpg') : '',
				pageCount: doc.number_of_pages_median || '',
				subjects: (doc.subject || []).slice(0, 5)
			};
		});
	} catch (e) { return []; }
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
				audiobookId: b.id, title: b.title || '',
				author: ((a.first_name || '') + ' ' + (a.last_name || '')).trim() || 'Unknown Author',
				totalTime: b.totaltime || '', numSections: Number(b.num_sections) || 0,
				coverUrl: b.url_image || ''
			};
		});
	} catch (e) { return []; }
}

function clientSearchPodcastDiscussions(query) {
	if (!query) return [];
	var props = PropertiesService.getScriptProperties();
	var apiKey = props.getProperty('PODCAST_INDEX_API_KEY');
	var apiSecret = props.getProperty('PODCAST_INDEX_API_SECRET');
	if (!apiKey || !apiSecret) return [];
	var ts = String(Math.floor(Date.now() / 1000));
	var authBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, apiKey + apiSecret + ts, Utilities.Charset.UTF_8);
	var auth = authBytes.map(function(b) {
		var v = (b < 0 ? b + 256 : b).toString(16);
		return v.length === 1 ? '0' + v : v;
	}).join('');
	var url = 'https://api.podcastindex.org/api/1.0/search/byterm?q=' + encodeURIComponent(query + ' book') + '&max=10';
	try {
		var resp = UrlFetchApp.fetch(url, {
			muteHttpExceptions: true,
			headers: {
				'X-Auth-Key': apiKey, 'X-Auth-Date': ts,
				'Authorization': auth, 'User-Agent': 'PageVault/1.0'
			}
		});
		if (resp.getResponseCode() !== 200) return [];
		var json = JSON.parse(resp.getContentText());
		return (json.feeds || []).slice(0, 10).map(function(feed) {
			return {
				id: feed.id || '', title: feed.title || '', author: feed.author || '',
				description: feed.description ? String(feed.description).slice(0, 140) : '',
				coverUrl: feed.image || feed.artwork || '',
				podcastLink: feed.link || feed.url || ''
			};
		});
	} catch (e) { return []; }
}

function clientGetAudiobookChapters(projectId) {
	if (!projectId) return [];
	var url = 'https://librivox.org/api/feed/audiotracks?project_id=' + projectId + '&format=json';
	try {
		var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
		var json = JSON.parse(resp.getContentText());
		return (json.sections || []).map(function(s, i) {
			return {
				chapterIndex: i, title: s.title || ('Chapter ' + (i + 1)),
				duration: s.playtime || '', url: s.listen_url || '',
				reader: (s.readers || [])[0] ? s.readers[0].display_name : ''
			};
		});
	} catch (e) { return []; }
}

function clientGetBookDetails(olid) {
	if (!olid) return null;
	var workKey = olid.indexOf('/works/') === 0 ? olid : '/works/' + olid;
	var url = 'https://openlibrary.org' + workKey + '.json';
	try {
		var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
		if (resp.getResponseCode() !== 200) return null;
		var data = JSON.parse(resp.getContentText());
		var desc = '';
		if (data.description) desc = (typeof data.description === 'object') ? (data.description.value || '') : String(data.description);
		var subjects = (data.subjects || []).slice(0, 10);
		var coverId = (data.covers || [])[0] || null;
		var coverUrl = coverId ? 'https://covers.openlibrary.org/b/id/' + coverId + '-L.jpg' : null;
		var authorKey = null;
		if (data.authors && data.authors[0] && data.authors[0].author) authorKey = data.authors[0].author.key || null;
		return {
			olid: olid, description: desc, subjects: subjects,
			coverUrl: coverUrl, coverId: coverId,
			authorKey: authorKey, firstPublish: data.first_publish_date || ''
		};
	} catch (e) { return null; }
}

function clientGetAuthorDetails(authorKey) {
	if (!authorKey) return null;
	var key = authorKey.indexOf('/authors/') === 0 ? authorKey : '/authors/' + authorKey;
	var olid = key.replace('/authors/', '');
	var url = 'https://openlibrary.org' + key + '.json';
	try {
		var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
		if (resp.getResponseCode() !== 200) return null;
		var data = JSON.parse(resp.getContentText());
		var bio = '';
		if (data.bio) bio = (typeof data.bio === 'object') ? (data.bio.value || '') : String(data.bio);
		var photoId = (data.photos || [])[0] || null;
		var photoUrl = photoId ? 'https://covers.openlibrary.org/a/olid/' + olid + '-M.jpg' : null;
		return {
			authorKey: key, name: data.name || '', bio: bio,
			birthDate: data.birth_date || '', deathDate: data.death_date || '',
			photoUrl: photoUrl
		};
	} catch (e) { return null; }
}

function clientCheckFreeEbook(isbn) {
	if (!isbn) return { available: false };
	var url = 'https://openlibrary.org/api/books?bibkeys=ISBN:' + encodeURIComponent(isbn) + '&jscmd=viewapi&format=json';
	try {
		var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
		if (resp.getResponseCode() !== 200) return { available: false };
		var data = JSON.parse(resp.getContentText());
		var entry = data['ISBN:' + isbn];
		if (!entry) return { available: false };
		var preview = entry.preview || 'noview';
		return {
			available: preview === 'full' || preview === 'limited',
			previewLevel: preview,
			readUrl: entry.read_url || entry.info_url || '',
			thumbnail: entry.thumbnail_url || ''
		};
	} catch (e) { return { available: false }; }
}

function clientSearchArchiveAudio(query) {
	if (!query) return [];
	var url = 'https://archive.org/advancedsearch.php?q=' + encodeURIComponent(query + ' mediatype:audio') +
		'&fl[]=identifier,title,creator,description,runtime&rows=8&output=json';
	try {
		var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
		if (resp.getResponseCode() !== 200) return [];
		var data = JSON.parse(resp.getContentText());
		return ((data.response || {}).docs || []).map(function(doc) {
			return {
				identifier: doc.identifier || '', title: doc.title || '',
				author: (Array.isArray(doc.creator) ? doc.creator[0] : doc.creator) || '',
				description: (Array.isArray(doc.description) ? doc.description[0] : doc.description) || '',
				runtime: doc.runtime || '',
				streamBase: 'https://archive.org/download/' + (doc.identifier || '')
			};
		});
	} catch (e) { return []; }
}

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
				chapterIndex: i, title: f.title || f.name,
				duration: f.length || '',
				url: 'https://archive.org/download/' + identifier + '/' + f.name
			};
		});
	} catch (e) { return []; }
}

// =====================================================================
//  NYT BESTSELLER CACHE
//  Add NYT_API_KEY in Project Settings → Script Properties to enable.
// =====================================================================

var NYT_LISTS = [
	'hardcover-fiction', 'hardcover-nonfiction', 'paperback-nonfiction',
	'young-adult-hardcover', 'childrens-middle-grade-hardcover',
	'graphic-books-and-manga', 'science', 'business-books'
];

function _getNytApiKey() {
	return PropertiesService.getScriptProperties().getProperty('NYT_API_KEY') || null;
}

function clientRefreshNYTCache() {
	var props = PropertiesService.getScriptProperties();
	var apiKey = _getNytApiKey();
	if (!apiKey) return;
	var cache = {}, currentFeed = [];
	NYT_LISTS.forEach(function(listName) {
		try {
			var url = 'https://api.nytimes.com/svc/books/v3/lists/current/' + encodeURIComponent(listName) + '.json?api-key=' + apiKey;
			var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
			if (resp.getResponseCode() !== 200) return;
			var data = JSON.parse(resp.getContentText());
			var results = data.results || {};
			var books = results.books || [];
			currentFeed.push({
				list: listName,
				listDisplay: results.display_name || _humanizeListName(listName),
				updatedAt: results.published_date || '',
				books: books.slice(0, 8).map(function(b) {
					return {
						rank: Number(b.rank) || 0, weeksOn: Number(b.weeks_on_list) || 0,
						title: b.title || '', author: b.author || '',
						description: b.description || '',
						isbn13: b.primary_isbn13 || '', isbn10: b.primary_isbn10 || '',
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
						rank: b.rank, weeksOn: b.weeks_on_list,
						list: listName, listDisplay: _humanizeListName(listName),
						title: b.title, author: b.author
					};
				});
			});
			Utilities.sleep(6500);
		} catch (e) {}
	});
	props.setProperty('NYT_CACHE', JSON.stringify(cache));
	props.setProperty('NYT_CACHE_DATE', new Date().toISOString().slice(0, 10));
	props.setProperty('NYT_FEED_CURRENT', JSON.stringify({
		updatedAt: new Date().toISOString().slice(0, 10), lists: currentFeed
	}));
}

function _getLibraryIsbnsForNyt() {
	var sheet = _ss().getSheetByName(SHEET_LIBRARY);
	if (!sheet || sheet.getLastRow() < LIBRARY_DATA_ROW) return [];
	var isbnCol = LIBRARY_HEADERS.indexOf('ISBN');
	if (isbnCol < 0) return [];
	var values = sheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL + isbnCol, sheet.getLastRow() - LIBRARY_DATA_ROW + 1, 1).getValues();
	var seen = {}, result = [];
	for (var i = 0; i < values.length; i++) {
		var isbn = _normalizeIsbn(values[i][0]);
		if (!isbn || seen[isbn]) continue;
		seen[isbn] = true; result.push(isbn);
	}
	return result;
}

function clientGetNYTBadgesForLibrary() {
	var raw = PropertiesService.getScriptProperties().getProperty('NYT_CACHE');
	if (!raw) return { byBookId: {}, byIsbn: {} };
	var cache; try { cache = JSON.parse(raw); } catch (e) { return { byBookId: {}, byIsbn: {} }; }
	var sheet = _ss().getSheetByName(SHEET_LIBRARY);
	if (!sheet || sheet.getLastRow() < LIBRARY_DATA_ROW) return { byBookId: {}, byIsbn: {} };
	var isbnCol = LIBRARY_HEADERS.indexOf('ISBN');
	var bookIdCol = LIBRARY_HEADERS.indexOf('BookId');
	var data = sheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, sheet.getLastRow() - LIBRARY_DATA_ROW + 1, LIBRARY_HEADERS.length).getValues();
	var byBookId = {}, byIsbn = {};
	for (var r = 0; r < data.length; r++) {
		var isbn = _normalizeIsbn(data[r][isbnCol] || '');
		var bookId = String(data[r][bookIdCol] || '').trim();
		if (isbn && cache[isbn]) {
			var badge = { rank: cache[isbn].rank, weeksOn: cache[isbn].weeksOn, list: cache[isbn].list, listDisplay: cache[isbn].listDisplay };
			byBookId[bookId] = badge; byIsbn[isbn] = badge;
		}
	}
	return { byBookId: byBookId, byIsbn: byIsbn };
}

function clientGetNYTFeed() {
	var raw = PropertiesService.getScriptProperties().getProperty('NYT_FEED_CURRENT');
	if (!raw) return { updatedAt: '', lists: [] };
	try {
		var parsed = JSON.parse(raw);
		return {
			updatedAt: parsed.updatedAt || '',
			lists: Array.isArray(parsed.lists) ? parsed.lists : []
		};
	} catch (e) { return { updatedAt: '', lists: [] }; }
}

function installNYTWeeklyTrigger() {
	var triggers = ScriptApp.getProjectTriggers();
	for (var i = 0; i < triggers.length; i++) {
		if (triggers[i].getHandlerFunction() === 'clientRefreshNYTCache') return;
	}
	ScriptApp.newTrigger('clientRefreshNYTCache').timeBased().everyWeeks(1)
		.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(3).create();
}

// =====================================================================
//  SYNC VERSION — Sheet→Webapp change detection
// =====================================================================

function clientGetSyncVersion() {
	return Number(PropertiesService.getScriptProperties().getProperty('SYNC_VERSION') || '0');
}
function _incrementSyncVersion() {
	var props = PropertiesService.getScriptProperties();
	var current = Number(props.getProperty('SYNC_VERSION') || '0');
	props.setProperty('SYNC_VERSION', String(current + 1));
}
function onEditSyncHandler(e) {
	if (!e || !e.range) return;
	var sheetName = e.range.getSheet().getName();
	var synced = [SHEET_LIBRARY, SHEET_CHALLENGES, SHEET_SHELVES, SHEET_PROFILE, SHEET_AUDIOBOOKS];
	if (synced.indexOf(sheetName) >= 0) _incrementSyncVersion();
}
function installSyncTrigger() {
	var triggers = ScriptApp.getProjectTriggers();
	for (var i = 0; i < triggers.length; i++) {
		if (triggers[i].getHandlerFunction() === 'onEditSyncHandler') {
			try { SpreadsheetApp.getUi().alert('Sync trigger is already installed.'); } catch (e) {}
			return;
		}
	}
	ScriptApp.newTrigger('onEditSyncHandler').forSpreadsheet(_ss()).onEdit().create();
	try { SpreadsheetApp.getUi().alert('Sync trigger installed.'); } catch (e) {}
}

// =====================================================================
//  MENU + FIRST-OPEN AUTO-SEED
// =====================================================================

function onOpen() {
	// Auto-init on every open — _seedDemoData() no-ops if the Library
	// already has data, so this is safe + instant after the first run.
	try { _dbLiteInitializeSheets(); } catch (e) { _log('error', 'onOpen', e); }

	try {
		var ui = SpreadsheetApp.getUi();
		var advanced = ui.createMenu('Advanced')
			.addItem('Refresh NYT Bestseller Cache', 'clientRefreshNYTCache')
			.addItem('Install Weekly NYT Refresh', 'installNYTWeeklyTrigger')
			.addItem('Install Live Sync Trigger',  'installSyncTrigger')
			.addSeparator()
			.addItem('Rebuild Sheet Structure', 'initializeSheets');

		ui.createMenu(_buildJourneyTitle())
			.addItem('📖 Open Web App',           '_openWebApp')
			.addSeparator()
			.addItem('🎨 Refresh Styling & Colors', '_reStyleCurrentTheme')
			.addItem('🗑  Clear All Books & Data',  'clientClearDemoData')
			.addSeparator()
			.addSubMenu(advanced)
			.addToUi();
	} catch (e) {}
}

function _openWebApp() {
	var url = ScriptApp.getService().getUrl();
	var safeSrc = '<script>window.open(' + JSON.stringify(url) + ');google.script.host.close();\x3c/script>';
	var html = HtmlService.createHtmlOutput(safeSrc).setWidth(1).setHeight(1);
	SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

function _reStyleCurrentTheme() {
	_dbLiteInitializeSheets();
	try { SpreadsheetApp.getUi().alert('Sheet styling refreshed.'); } catch (e) {}
}
