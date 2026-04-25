/* =====================================================================
 *  code1.gs — My Reading Journey (Standalone)
 *
 *  Single-file Google Apps Script backend + sheet layout for the
 *  My Reading Journey template (Etsy release).
 *
 *  WHAT THIS FILE DOES:
 *  - Serves the web app (doGet → index.html / index2.html / index3.html)
 *  - Owns all sheet constants & schema
 *  - Builds the visible Library tab to match the product screenshot:
 *      Rows 1–7 = banner (image + title). Row 8 = column headers. Row 9+ = data.
 *  - Keeps every utility sheet hidden (Challenges, Shelves, Profile, Audiobooks)
 *  - Seeds 72 demo books automatically on first open
 *  - Exposes the full client* API surface used by index.html / index2.html / index3.html
 *
 *  PRODUCTS:
 *  - Product 1 (index.html)  → Romantic theme (pink/red)   — set PRODUCT_VARIANT = 'index'
 *  - Product 2 (index2.html) → Horizon theme  (blue)       — set PRODUCT_VARIANT = 'index2'
 *  - Product 3 (index3.html) → Blossom theme  (mauve)      — set PRODUCT_VARIANT = 'index3'
 *
 *  DEPLOY:
 *  1. Extensions → Apps Script → paste this file → Save
 *  2. Set PRODUCT_VARIANT below to match the product you're distributing
 *  3. Delete any older Code.gs if present (all runtime now lives here)
 *  4. Add your NYT and PodcastIndex keys in the constants below (optional but recommended)
 *  5. Buyer opens their copied sheet: first-open setup runs automatically
 *     (sheet init + sync trigger + NYT warmup trigger when key is present)
 *  6. Buyer uses the one-time setup dialog to deploy Web App and gets their app URL
 * ===================================================================== */

/* =====================================================================
 *  API KEYS — fill these in on your MASTER sheet before publishing.
 *  Buyers will inherit them automatically when they "Make a copy".
 *  All three services have free tiers; rotate here if a key gets abused.
 *    NYT:           https://developer.nytimes.com  (free, 500 req/day)
 *    PodcastIndex:  https://podcastindex.org/      (free)
 * ===================================================================== */
var NYT_API_KEY              = '';
var PODCAST_INDEX_API_KEY    = '';
var PODCAST_INDEX_API_SECRET = '';

function _dbLiteInitializeSheets() {
	// Migrate old default theme: only for Product 1 — 'blossom' (pink) → 'romantic' (red).
	// Scoped to PRODUCT_VARIANT 'index' so Product 3 (blossom) is never affected.
	var props = PropertiesService.getScriptProperties();
	if (PRODUCT_VARIANT === 'index' && props.getProperty('DEFAULT_THEME_MIGRATED_V2') !== '1') {
		try {
			var pSheet = _ss().getSheetByName(SHEET_PROFILE);
			var pRow = _getProfileDataRow(pSheet);
			if (pSheet && pRow >= 2) {
				var tCol = PROFILE_HEADERS.indexOf('Theme') + 1;
				var cur = String(pSheet.getRange(pRow, tCol).getValue() || '').toLowerCase();
				if (!cur || cur === 'blossom') pSheet.getRange(pRow, tCol).setValue('romantic');
			}
		} catch(e) {}
		props.setProperty('DEFAULT_THEME_MIGRATED_V2', '1');
	}

	var ss = _ss();

	// Ensure core data sheets exist with correct headers.
	var library = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
	_getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
	_getOrCreateSheet(SHEET_SHELVES, SHELF_HEADERS);
	var profile = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
	_getOrCreateSheet(SHEET_AUDIOBOOKS, AUDIOBOOK_HEADERS);

	// Seed profile defaults FIRST so that _getCurrentTheme() reads the correct
	// product-specific theme (horizon/blossom) rather than falling back to 'romantic'.
	_dbLiteEnsureProfileDefaults(profile);

	var theme = _getCurrentTheme();
	_dbLiteInitLibrarySheet(library, theme);

	// Style hidden data tabs with the theme palette header so they look intentional if unhidden.
	[
		{ name: SHEET_CHALLENGES, headers: CHALLENGE_HEADERS, label: 'Challenges' },
		{ name: SHEET_SHELVES,    headers: SHELF_HEADERS,     label: 'Shelves'    },
		{ name: SHEET_PROFILE,    headers: PROFILE_HEADERS,   label: 'Profile'    },
		{ name: SHEET_AUDIOBOOKS, headers: AUDIOBOOK_HEADERS, label: 'Audiobooks' }
	].forEach(function(d) {
		var s = ss.getSheetByName(d.name);
		if (s) _dbLiteStyleHiddenHeader(s, d.headers, d.label, theme);
	});

	// Seed demo BEFORE My Year so the cover grid has books to display.
	_seedDemoData();

	_dbLiteInitMyYearSheet(ss, theme);
	_dbLiteArrangeTabs(ss);
}

function _dbLiteStyleHiddenHeader(sheet, headers, label, themeName) {
	if (!sheet) return;
	var t = _dbLiteTheme(themeName);
	var display = _displayHeaders ? _displayHeaders(headers) : headers;
	var cols = display.length;
	try { _ensureColumns(sheet, cols); } catch(e) {}
	// Row 1 = themed palette banner with tab label right-aligned.
	try { sheet.getRange(1, 1, 1, cols).breakApart(); } catch(e) {}
	sheet.getRange(1, 1, 1, cols).merge()
		.setValue(String(label || sheet.getName()).toUpperCase() + '  _')
		.setBackground(t.headerBg).setFontColor(t.headerText || '#FFFFFF')
		.setFontFamily('Montserrat').setFontSize(22).setFontWeight('bold')
		.setHorizontalAlignment('right').setVerticalAlignment('middle');
	sheet.setRowHeight(1, 48);
	// Row 2 = field headers on white, bold with themed bottom border.
	sheet.getRange(2, 1, 1, cols).setValues([display])
		.setBackground('#FFFFFF').setFontColor('#0F172A')
		.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold')
		.setHorizontalAlignment('left').setVerticalAlignment('middle');
	sheet.setRowHeight(2, 32);
	sheet.getRange(2, 1, 1, cols)
		.setBorder(null, null, true, null, null, null, t.accent, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
	try { sheet.setFrozenRows(2); } catch(e) {}
	sheet.setTabColor(t.accent);
	sheet.setHiddenGridlines(true);
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
	if (_getProfileDataRow(sheet) >= HIDDEN_DATA_ROW) return;

	var defaults = PROFILE_HEADERS.map(function(h) {
		switch (h) {
			case 'Name': return '';
			case 'Motto': return 'A focused place to track every book';
			case 'PhotoData': return '';
			case 'Theme': return _VIEW_THEME_MAP[PRODUCT_VARIANT] || PropertiesService.getScriptProperties().getProperty('PRODUCT_DEFAULT_THEME') || 'romantic';
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
	sheet.getRange(HIDDEN_DATA_ROW, 1, 1, PROFILE_HEADERS.length).setValues([defaults]);
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

function _dbLiteUpsertLibraryBannerImage(sheet) {
	// Banner is fully optional. The Library banner already renders beautifully
	// without an image (themed background + large "Reading _ / LIBRARY _" text).
	// Buyers can OPTIONALLY add their own banner by:
	//   1. Uploading an image to their own Drive
	//   2. Project Settings → Script Properties → add BANNER_IMAGE_FILE_ID = <the file id>
	// If unset (the default), we skip image insertion entirely — no failure mode,
	// no permissions issue, no silent image deletion.
	if (!sheet) return;
	var fileId = '';
	try { fileId = PropertiesService.getScriptProperties().getProperty('BANNER_IMAGE_FILE_ID') || ''; } catch (e) {}
	if (!fileId) return;
	var blob = null;
	try { blob = DriveApp.getFileById(fileId).getBlob(); } catch (e) { return; }
	if (!blob) return;
	try {
		// Only remove OUR previously-inserted banner (anchored in rows 1-6, cols 1-6).
		var images = sheet.getImages();
		images.forEach(function(img) {
			try {
				var anchor = img.getAnchorCell();
				if (anchor && anchor.getRow() <= 6 && anchor.getColumn() <= 6) img.remove();
			} catch (e) {}
		});
		var image = sheet.insertImage(blob, 1, 1, 14, 8);
		image.setWidth(470);
		image.setHeight(180);
	} catch (e) {}
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

	// ── Rows 1–8: banner layout ───────────────────────────────────────────
	// 52pt Montserrat cap+descender needs ~80px. Rows 2+3 merged = 84px.
	// Row 1/6 = 14px padding. Row 7 = 8px separator. Row 8 = 48px header.
	sheet.setRowHeight(1, 14);
	sheet.setRowHeights(2, 2, 42);
	sheet.setRowHeights(4, 2, 42);
	sheet.setRowHeight(6, 14);
	sheet.setRowHeight(7, 8);
	sheet.setRowHeight(8, 48);

	// Rows 1–6: themed background (floating cat image sits on top in cols B–F)
	sheet.getRange(1, 1, 6, totalCols).setBackground(t.headerBg);

	// Rows 2–3, cols G–L (7–12): "Reading _" — 52px normal white, right-aligned
	try { sheet.getRange(2, 7, 2, 6).merge(); } catch(e) {}
	sheet.getRange(2, 7, 2, 6)
		.setValue('Reading _')
		.setFontFamily('Montserrat').setFontSize(52).setFontWeight('normal')
		.setFontColor('#FFFFFF').setBackground(t.headerBg)
		.setHorizontalAlignment('right').setVerticalAlignment('bottom');

	// Rows 4–5, cols G–L (7–12): "LIBRARY _" — 52px bold white, right-aligned
	try { sheet.getRange(4, 7, 2, 6).merge(); } catch(e) {}
	sheet.getRange(4, 7, 2, 6)
		.setValue('LIBRARY _')
		.setFontFamily('Montserrat').setFontSize(52).setFontWeight('bold')
		.setFontColor('#FFFFFF').setBackground(t.headerBg)
		.setHorizontalAlignment('right').setVerticalAlignment('top');
	_dbLiteUpsertLibraryBannerImage(sheet);

	// Rows 7–8: header strip on white (matches data area).
	var HDR_BG = '#FFFFFF';
	sheet.getRange(7, 1, 2, totalCols).setBackground(HDR_BG);

	// Row 8: column header labels — 12pt bold, centered. Dates stack via \n + wrap.
	var hdrLabels = [''];  // col A = blank (row-number column)
	LIBRARY_HEADERS.slice(0, LIBRARY_VISIBLE_COUNT).forEach(function(h) {
		var label = (DISPLAY_MAP[h] || h).toUpperCase();
		// Only stack the two date columns (they share limited horizontal space).
		if (h === 'DateStarted') label = 'DATE\nSTARTED';
		else if (h === 'DateFinished') label = 'DATE\nFINISHED';
		hdrLabels.push(label);
	});
	sheet.getRange(LIBRARY_HEADER_ROW, 1, 1, hdrLabels.length)
		.setValues([hdrLabels])
		.setBackground(HDR_BG).setFontColor('#1E293B').setFontWeight('bold')
		.setFontFamily('Montserrat').setFontSize(12)
		.setVerticalAlignment('middle').setHorizontalAlignment('center')
		.setWrap(false);
	// Only date columns get wrap enabled (so the \n renders as two lines).
	var dsColH = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('DateStarted');
	var dfColH = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('DateFinished');
	sheet.getRange(LIBRARY_HEADER_ROW, dsColH, 1, 1).setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
	sheet.getRange(LIBRARY_HEADER_ROW, dfColH, 1, 1).setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
	sheet.getRange(LIBRARY_HEADER_ROW, LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Title'),  1, 1).setHorizontalAlignment('left');
	sheet.getRange(LIBRARY_HEADER_ROW, LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Author'), 1, 1).setHorizontalAlignment('left');

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

	// ── Visible column widths (B–K: Title through Favorite) ─────────────────
	// Order: Title, Author, Status(wider), Genre, Rating, Format, Pages, DateStarted, DateFinished, Favorite
	var visibleWidths = [260, 180, 130, 150, 140, 130, 120, 90, 115, 115, 80];
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

	// Rating column — larger gold stars (base style; conditional formats reinforce)
	var ratingColBase = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Rating');
	sheet.getRange(LIBRARY_DATA_ROW, ratingColBase, 5000, 1)
		.setFontFamily('Montserrat').setFontSize(16).setFontWeight('bold')
		.setFontColor('#F59E0B').setHorizontalAlignment('center')
		.setVerticalAlignment('middle');

	// Favorite column — larger red heart
	var favColBase = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Favorite');
	sheet.getRange(LIBRARY_DATA_ROW, favColBase, 5000, 1)
		.setFontSize(16).setFontWeight('bold').setFontColor('#DC2626')
		.setHorizontalAlignment('center');

	// ── Row heights (all 5000 data rows at once) ──────────────────────────────
	sheet.setRowHeights(LIBRARY_DATA_ROW, 5000, 44);

	// ── Per-column alignment + number formats ─────────────────────────────────
	var titleCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Title');
	var authorCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Author');
	var pagesCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Pages');
	var dsCol     = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('DateStarted');
	var dfCol     = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('DateFinished');

	sheet.getRange(LIBRARY_DATA_ROW, titleCol,  5000, 1).setFontWeight('bold').setHorizontalAlignment('left');
	sheet.getRange(LIBRARY_DATA_ROW, authorCol, 5000, 1).setHorizontalAlignment('left');
	sheet.getRange(LIBRARY_DATA_ROW, pagesCol,  5000, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
	sheet.getRange(LIBRARY_DATA_ROW, dsCol,     5000, 1).setNumberFormat('mmm d, yyyy').setHorizontalAlignment('center');
	sheet.getRange(LIBRARY_DATA_ROW, dfCol,     5000, 1).setNumberFormat('mmm d, yyyy').setHorizontalAlignment('center');

	// Center-align chip columns
	['Status','Genre','Rating','Format'].forEach(function(h) {
		var col = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf(h);
		sheet.getRange(LIBRARY_DATA_ROW, col, 5000, 1).setHorizontalAlignment('center');
	});

	// ── Notebook-style ruled paper (full width, no vertical lines) ─────────
	var NB  = '#A4C2F4';
	var NBS = SpreadsheetApp.BorderStyle.SOLID;
	// One call across all columns: vertical=false removes grid lines, horizontal=true adds ruled lines
	sheet.getRange(LIBRARY_DATA_ROW, 1, 5000, totalCols)
		.setBorder(null, null, true, null, false, true, NB, NBS);
	// Red margin line on the left edge of col B
	sheet.getRange(LIBRARY_DATA_ROW, 2, 5000, 1)
		.setBorder(null, true, null, null, null, null, '#FF4C4C', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

	// ── Freeze template header row (no filter — it adds an unwanted dark outline) ─
	sheet.setFrozenRows(LIBRARY_HEADER_ROW);
	sheet.setFrozenColumns(0);

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

	// Keep My Year width consistent with Library: exactly 12 visible columns (A-L).
	// Library total: 40 + 260+180+130+150+140+130+120+90+115+115+80 = ~1550 px.
	// Match here with 12 columns × 126 = 1512 px so the tab is visually the same width.
	var NUM_COLS = 12;
	var COV_START = 1;
	var COV_PER_ROW = 12;
	var COVER_W = 126;
	var COVER_H = 218;
	var HELPER_ROW = 1800;
	var HELPER_ROWS = 400;

	_ensureColumns(sheet, NUM_COLS);
	_ensureRows(sheet, 2200);
	sheet.setHiddenGridlines(true);
	sheet.setTabColor(t.accent);

	for (var c = 1; c <= NUM_COLS; c++) sheet.setColumnWidth(c, COVER_W);

	// Remove old charts so we can rebuild cleanly every run.
	try {
		sheet.getCharts().forEach(function(ch) { sheet.removeChart(ch); });
	} catch(e) {}

	var MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
	var DOW = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
	var thisYear = new Date().getFullYear();
	var thisMonth = new Date().getMonth();

	var libSheet = ss.getSheetByName(SHEET_LIBRARY);
	var allBooks = [];
	if (libSheet && libSheet.getLastRow() >= LIBRARY_DATA_ROW) {
		var numDR = libSheet.getLastRow() - LIBRARY_DATA_ROW + 1;
		var vals = libSheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, numDR, LIBRARY_HEADERS.length).getValues();
		var tIdx = LIBRARY_HEADERS.indexOf('Title');
		var stIdx = LIBRARY_HEADERS.indexOf('Status');
		var gIdx = LIBRARY_HEADERS.indexOf('Genre');
		var pgIdx = LIBRARY_HEADERS.indexOf('Pages');
		var cpIdx = LIBRARY_HEADERS.indexOf('CurrentPage');
		var rtIdx = LIBRARY_HEADERS.indexOf('Rating');
		var cuIdx = LIBRARY_HEADERS.indexOf('CoverUrl');
		var dsIdx = LIBRARY_HEADERS.indexOf('DateStarted');
		var dfIdx = LIBRARY_HEADERS.indexOf('DateFinished');
		var isIdx = LIBRARY_HEADERS.indexOf('ISBN');
		var auIdx = LIBRARY_HEADERS.indexOf('Author');
		var serIdx = LIBRARY_HEADERS.indexOf('Series');

		vals.forEach(function(r) {
			if (!String(r[tIdx] || '').trim()) return;
			var status = String(r[stIdx] || '').toLowerCase();
			var ratingStr = String(r[rtIdx] || '');
			allBooks.push({
				title: String(r[tIdx] || ''),
				author: String(r[auIdx] || ''),
				series: String(r[serIdx] || ''),
				status: status,
				genre: String(r[gIdx] || ''),
				pages: Number(r[pgIdx]) || 0,
				currentPage: Number(r[cpIdx]) || 0,
				rating: (ratingStr.match(/★/g) || []).length,
				coverUrl: String(r[cuIdx] || '').trim(),
				isbn: String(r[isIdx] || '').replace(/["'\s]/g, ''),
				dateStarted: r[dsIdx],
				dateFinished: r[dfIdx]
			});
		});
	}

	var finished = allBooks.filter(function(b) { return b.status === 'finished'; });
	var finishedThisYear = finished.filter(function(b) {
		if (!b.dateFinished) return false;
		var d = new Date(b.dateFinished);
		return !isNaN(d.getTime()) && d.getFullYear() === thisYear;
	});
	var thisMonthCount = finished.filter(function(b) {
		if (!b.dateFinished) return false;
		var d = new Date(b.dateFinished);
		return !isNaN(d.getTime()) && d.getFullYear() === thisYear && d.getMonth() === thisMonth;
	}).length;

	var totalPages = allBooks.reduce(function(s, b) {
		if (b.status === 'finished') return s + (b.pages || 0);
		if (b.status === 'reading' || b.status === 'dnf') return s + (b.currentPage || 0);
		return s;
	}, 0);

	var ratedBooks = allBooks.filter(function(b) { return b.rating > 0; });
	var avgRating = ratedBooks.length
		? (ratedBooks.reduce(function(s, b) { return s + b.rating; }, 0) / ratedBooks.length).toFixed(1)
		: '0';

	// Reading streak (consecutive days with finished books)
	var finishedDateSet = {};
	finished.forEach(function(b) {
		if (!b.dateFinished) return;
		var key = String(b.dateFinished).slice(0, 10);
		finishedDateSet[key] = true;
	});
	var streak = 0;
	var cur = new Date();
	if (!finishedDateSet[cur.toISOString().slice(0, 10)]) {
		cur.setDate(cur.getDate() - 1);
	}
	while (finishedDateSet[cur.toISOString().slice(0, 10)]) {
		streak++;
		cur.setDate(cur.getDate() - 1);
	}

	// Yearly goal
	var profileSheet = ss.getSheetByName(SHEET_PROFILE);
	var yearlyGoal = 50;
	var profileRow = _getProfileDataRow(profileSheet);
	if (profileSheet && profileRow >= 2) {
		var goalColIdx = PROFILE_HEADERS.indexOf('YearlyGoal') + 1;
		yearlyGoal = Number(profileSheet.getRange(profileRow, goalColIdx).getValue()) || 50;
	}
	var goalDone = finishedThisYear.length;
	var goalPct = Math.min(100, Math.round(goalDone / Math.max(1, yearlyGoal) * 100));
	var goalBar = new Array(Math.round(goalPct / 7) + 1).join('█') + new Array(15 - Math.round(goalPct / 7) + 1).join('░');

	// Aggregate counts for charts
	var genreCounts = {};
	allBooks.forEach(function(b) {
		var g = String(b.genre || 'Other').trim() || 'Other';
		genreCounts[g] = (genreCounts[g] || 0) + 1;
	});

	var statusCounts = {
		Reading: allBooks.filter(function(b) { return b.status === 'reading'; }).length,
		Finished: allBooks.filter(function(b) { return b.status === 'finished'; }).length,
		'Want to Read': allBooks.filter(function(b) { return b.status === 'want to read' || b.status === 'want-to-read'; }).length,
		DNF: allBooks.filter(function(b) { return b.status === 'dnf'; }).length
	};

	var monthlyBooks = [];
	for (var m = 0; m < 12; m++) {
		var ct = 0;
		finishedThisYear.forEach(function(b) {
			var d = new Date(b.dateFinished);
			if (!isNaN(d.getTime()) && d.getMonth() === m) ct++;
		});
		monthlyBooks.push(ct);
	}
	var monthlyGoal = Math.max(1, Math.round(yearlyGoal / 12));

	var weeklyBooks = [0,0,0,0,0,0,0];
	var today = new Date();
	var weekStart = new Date(today);
	weekStart.setDate(today.getDate() - today.getDay());
	finished.forEach(function(b) {
		if (!b.dateFinished) return;
		var d = new Date(b.dateFinished);
		if (isNaN(d.getTime())) return;
		var diff = Math.floor((new Date(d.getFullYear(), d.getMonth(), d.getDate()) - new Date(weekStart.getFullYear(), weekStart.getMonth(), weekStart.getDate())) / 86400000);
		if (diff >= 0 && diff < 7) weeklyBooks[diff]++;
	});

	var topGenre = '—';
	var topGenreCount = 0;
	Object.keys(genreCounts).forEach(function(g) {
		if (genreCounts[g] > topGenreCount) { topGenre = g; topGenreCount = genreCounts[g]; }
	});

	var daysArr = finished.reduce(function(arr, b) {
		if (b.dateStarted && b.dateFinished) {
			var d1 = new Date(b.dateStarted), d2 = new Date(b.dateFinished);
			if (!isNaN(d1.getTime()) && !isNaN(d2.getTime()) && d2 > d1) arr.push((d2 - d1) / 86400000);
		}
		return arr;
	}, []);
	var avgDays = daysArr.length ? (daysArr.reduce(function(s, d) { return s + d; }, 0) / daysArr.length).toFixed(1) : '—';

	// ── Rolling 12-month window (always has data, regardless of calendar year) ──
	var nowD = new Date();
	var rollingMonths = []; // [{label, ym, books, pages, goal}]
	for (var rm = 11; rm >= 0; rm--) {
		var d0 = new Date(nowD.getFullYear(), nowD.getMonth() - rm, 1);
		rollingMonths.push({
			label: MN[d0.getMonth()] + (rm > 0 || nowD.getMonth() !== d0.getMonth() ? " '" + String(d0.getFullYear()).slice(-2) : ''),
			ym: d0.getFullYear() * 100 + d0.getMonth(),
			books: 0,
			pages: 0,
			goal: Math.max(1, Math.round((yearlyGoal || 50) / 12))
		});
	}
	finished.forEach(function(b) {
		if (!b.dateFinished) return;
		var d = new Date(b.dateFinished);
		if (isNaN(d.getTime())) return;
		var k = d.getFullYear() * 100 + d.getMonth();
		for (var i = 0; i < rollingMonths.length; i++) {
			if (rollingMonths[i].ym === k) {
				rollingMonths[i].books++;
				rollingMonths[i].pages += (b.pages || 0);
				break;
			}
		}
	});

	// YTD cumulative goal progress
	var ytdRows = [];
	var cumulative = 0;
	var monthlyGoalNum = Math.max(1, Math.round((yearlyGoal || 50) / 12));
	for (var ym2 = 0; ym2 <= thisMonth; ym2++) {
		var monthBooks = finishedThisYear.filter(function(b) {
			var d = new Date(b.dateFinished);
			return !isNaN(d.getTime()) && d.getMonth() === ym2;
		}).length;
		cumulative += monthBooks;
		ytdRows.push({ label: MN[ym2], books: monthBooks, target: monthlyGoalNum * (ym2 + 1) });
	}

	var coverBooks = [];
	allBooks.forEach(function(b) {
		var url = (b.coverUrl && b.coverUrl.indexOf('http') === 0)
			? b.coverUrl
			: (b.isbn ? 'https://covers.openlibrary.org/b/isbn/' + b.isbn + '-L.jpg' : '');
		if (!url) return;
		coverBooks.push({
			url: url.replace(/["']/g, ''),
			title: b.title,
			author: b.author,
			series: b.series,
			genre: b.genre,
			status: b.status,
			rating: b.rating
		});
	});

	// ── Layout foundation ─────────────────────────────────────────────────
	function _setCoverCell(cell, bk) {
		cell.setFormula('=IMAGE("' + bk.url + '",4,' + COVER_H + ',' + COVER_W + ')')
			.setBackground('#FFFFFF')
			.setHorizontalAlignment('center').setVerticalAlignment('middle');
		var metaParts = [];
		if (bk.series) metaParts.push('Series: ' + bk.series);
		if (bk.genre) metaParts.push(bk.genre);
		if (bk.status) metaParts.push(bk.status.charAt(0).toUpperCase() + bk.status.slice(1));
		if (bk.rating) metaParts.push(new Array(bk.rating + 1).join('★'));
		var noteParts = [bk.title];
		if (bk.author) noteParts.push('by ' + bk.author);
		if (metaParts.length) noteParts.push(metaParts.join('  ·  '));
		cell.setNote(noteParts.join('\n'));
	}

	var heroCount = Math.min(coverBooks.length, 24);
	var heroRows = Math.max(1, Math.ceil(Math.max(heroCount, 1) / COV_PER_ROW));
	// Layout anchors — banner 1-7 (Library-style), KPI cards 8-10, spacer 11,
	// hero header 12, hero covers from 13.
	var KPI_TITLE = 8, KPI_VALUE = 9, KPI_SUB = 10;
	var HERO_HEADER = 12;
	var HERO_TOP = 13;
	var ANALYTICS_HEADER = HERO_TOP + heroRows + 1;
	var CHART_TOP = ANALYTICS_HEADER + 1;
	var CHART_BOTTOM = CHART_TOP + 15;
	var UTIL_HEADER = CHART_BOTTOM + 16;
	var UTIL_TOP = UTIL_HEADER + 1;
	var FULL_HEADER = UTIL_TOP + 9;
	var FULL_TOP = FULL_HEADER + 1;

	// ── Banner (rows 1–7) — mirrors Library banner styling ──────────────
	// Rows 1–6: themed background. Rows 2–3 "My Year _" light, rows 4–5 "2026 _" bold, both right-aligned.
	try { sheet.getRange(1, 1, 7, NUM_COLS).breakApart(); } catch(e) {}
	sheet.getRange(1, 1, 7, NUM_COLS).clearContent();
	sheet.setRowHeight(1, 10);
	sheet.setRowHeights(2, 2, 28);
	sheet.setRowHeights(4, 2, 28);
	sheet.setRowHeight(6, 10);
	sheet.setRowHeight(7, 8);
	sheet.getRange(1, 1, 6, NUM_COLS).setBackground(t.headerBg);
	try { sheet.getRange(2, 7, 2, 6).merge(); } catch(e) {}
	sheet.getRange(2, 7, 2, 6)
		.setValue('My Year _')
		.setFontFamily('Montserrat').setFontSize(34).setFontWeight('normal')
		.setFontColor('#FFFFFF').setBackground(t.headerBg)
		.setHorizontalAlignment('right').setVerticalAlignment('bottom');
	try { sheet.getRange(4, 7, 2, 6).merge(); } catch(e) {}
	sheet.getRange(4, 7, 2, 6)
		.setValue(thisYear + ' _')
		.setFontFamily('Montserrat').setFontSize(34).setFontWeight('bold')
		.setFontColor('#FFFFFF').setBackground(t.headerBg)
		.setHorizontalAlignment('right').setVerticalAlignment('top');
	sheet.getRange(7, 1, 1, NUM_COLS).setBackground('#FFFFFF');

	// KPI cards — colored accent stripe + bold number; no emojis.
	var cardPalette = [
		{ accent:'#6366F1', tint:'#EEF2FF' },
		{ accent:'#EC4899', tint:'#FDF2F8' },
		{ accent:'#10B981', tint:'#ECFDF5' },
		{ accent:'#F59E0B', tint:'#FFFBEB' }
	];
	var cards = [
		{ c1:1, c2:3, title:'BOOKS THIS YEAR', value:String(goalDone), sub:goalDone + ' of ' + yearlyGoal + ' goal  ·  ' + goalPct + '%' },
		{ c1:4, c2:6, title:'THIS MONTH', value:String(thisMonthCount), sub:MN[thisMonth] + ' ' + thisYear + '  ·  ' + streak + ' day streak' },
		{ c1:7, c2:9, title:'PAGES TRACKED', value:(totalPages >= 1000 ? (totalPages / 1000).toFixed(1) + 'K' : String(totalPages)), sub:allBooks.length + ' books  ·  avg ' + (avgDays === '—' ? '—' : avgDays + 'd') },
		{ c1:10, c2:12, title:'AVG RATING', value:avgRating + ' / 5', sub:'Top: ' + topGenre + '  ·  ' + ratedBooks.length + ' rated' }
	];
	cards.forEach(function(cd, idx) {
		var pal = cardPalette[idx % cardPalette.length];
		var w = cd.c2 - cd.c1 + 1;
		sheet.getRange(KPI_TITLE, cd.c1, 1, w).merge()
			.setValue('  ' + cd.title)
			.setBackground(pal.tint)
			.setFontColor(pal.accent).setFontFamily('Montserrat').setFontSize(10).setFontWeight('bold')
			.setHorizontalAlignment('left').setVerticalAlignment('middle');
		sheet.getRange(KPI_VALUE, cd.c1, 1, w).merge()
			.setValue('  ' + cd.value)
			.setBackground('#FFFFFF')
			.setFontColor('#0F172A').setFontFamily('Montserrat').setFontSize(32).setFontWeight('bold')
			.setHorizontalAlignment('left').setVerticalAlignment('middle');
		sheet.getRange(KPI_SUB, cd.c1, 1, w).merge()
			.setValue('  ' + cd.sub)
			.setBackground('#FFFFFF')
			.setFontColor('#64748B').setFontFamily('Montserrat').setFontSize(10)
			.setHorizontalAlignment('left').setVerticalAlignment('middle');
		sheet.getRange(KPI_TITLE, cd.c1, 3, w)
			.setBorder(true, true, true, true, false, false, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
		sheet.getRange(KPI_TITLE, cd.c1, 3, 1)
			.setBorder(null, true, null, null, null, null, pal.accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
	});
	sheet.setRowHeight(KPI_TITLE, 30);
	sheet.setRowHeight(KPI_VALUE, 54);
	sheet.setRowHeight(KPI_SUB, 28);
	sheet.setRowHeight(11, 16);

	// Top cover wall hero strip — match Library tab header styling (dark bar, white text)
	sheet.getRange(HERO_HEADER, 1, 1, NUM_COLS).merge()
		.setValue('   COVER WALL')
		.setBackground(t.headerBg).setFontColor(t.headerText || '#FFFFFF')
		.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold')
		.setHorizontalAlignment('left').setVerticalAlignment('middle');
	sheet.setRowHeight(HERO_HEADER, 32);

	if (!coverBooks.length) {
		sheet.setRowHeight(HERO_TOP, 58);
		sheet.getRange(HERO_TOP, 1, 1, NUM_COLS).merge()
			.setValue('No books with cover images yet. Add books with cover URLs, then run Advanced → Rebuild Sheet Structure.')
			.setFontFamily('Montserrat').setFontSize(10).setFontColor('#9CA3AF')
			.setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#FFFFFF');
		return;
	}

	for (var hr = 0; hr < heroRows; hr++) {
		sheet.setRowHeight(HERO_TOP + hr, COVER_H);
	}
	for (var h = 0; h < heroCount; h++) {
		var heroRow = HERO_TOP + Math.floor(h / COV_PER_ROW);
		var heroCol = 1 + (h % COV_PER_ROW);
		_setCoverCell(sheet.getRange(heroRow, heroCol), coverBooks[h]);
	}
	if (heroRows < 2) sheet.setRowHeight(HERO_TOP + 1, 10);

	// Chart helper data
	sheet.getRange(HELPER_ROW, 1, HELPER_ROWS, NUM_COLS).clearContent().setBackground('#FFFFFF').setFontColor('#FFFFFF');
	var paceData = [['Month', 'Books']];
	rollingMonths.forEach(function(rm) { paceData.push([rm.label, Math.max(0, rm.books)]); });
	if (paceData.length < 2) paceData.push(['Now', 0]);
	sheet.getRange(HELPER_ROW, 1, paceData.length, 2).setValues(paceData);

	var ytdData = [['Month', 'Books', 'Target']];
	if (!ytdRows.length) ytdData.push(['—', 0, monthlyGoalNum]);
	else ytdRows.forEach(function(y) { ytdData.push([y.label, y.books, y.target]); });
	sheet.getRange(HELPER_ROW, 4, ytdData.length, 3).setValues(ytdData);

	var genrePairs = Object.keys(genreCounts)
		.map(function(g) { return [g, genreCounts[g]]; })
		.sort(function(a, b) { return b[1] - a[1]; })
		.slice(0, 8);
	if (!genrePairs.length) genrePairs = [['No Data', 1]];
	var genreRows = [['Genre', 'Count']].concat(genrePairs);
	sheet.getRange(HELPER_ROW, 8, genreRows.length, 2).setValues(genreRows);

	var statusRows = [['Status', 'Count']];
	Object.keys(statusCounts).forEach(function(k) { statusRows.push([k, Math.max(0, statusCounts[k])]); });
	sheet.getRange(HELPER_ROW, 11, statusRows.length, 2).setValues(statusRows);
	SpreadsheetApp.flush();

	// Analytics header — match Library dark header styling
	sheet.setRowHeight(ANALYTICS_HEADER - 1, 12);
	sheet.getRange(ANALYTICS_HEADER, 1, 1, NUM_COLS).merge()
		.setValue('   ANALYTICS')
		.setBackground(t.headerBg).setFontColor(t.headerText || '#FFFFFF')
		.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold')
		.setHorizontalAlignment('left').setVerticalAlignment('middle');
	sheet.setRowHeight(ANALYTICS_HEADER, 32);
	for (var crow = CHART_TOP; crow < CHART_TOP + 30; crow++) sheet.setRowHeight(crow, 22);
	sheet.setRowHeight(CHART_TOP + 14, 12);

	// Charts — distinct palette per chart for visual diversity
	var DIVERSE_PALETTE = ['#6366F1', '#EC4899', '#10B981', '#F59E0B', '#3B82F6', '#8B5CF6', '#EF4444', '#06B6D4', '#A855F7', '#14B8A6', '#F97316', '#0EA5E9'];
	var CHART_W = Math.max(620, (COVER_W * 6) - 24), CHART_H = 320;
	try {
		var paceChart = sheet.newChart()
			.setChartType(Charts.ChartType.COLUMN)
			.addRange(sheet.getRange(HELPER_ROW, 1, paceData.length, 2))
			.setNumHeaders(1)
			.setPosition(CHART_TOP, 1, 8, 4)
			.setOption('title', 'Reading Pace  ·  last 12 months')
			.setOption('width', CHART_W)
			.setOption('height', CHART_H)
			.setOption('backgroundColor', '#FFFFFF')
			.setOption('legend', { position: 'none' })
			.setOption('colors', ['#6366F1'])
			.setOption('hAxis', { textStyle: { fontName: 'Montserrat', fontSize: 11, color: '#475569' } })
			.setOption('vAxis', { textStyle: { fontName: 'Montserrat', fontSize: 11, color: '#475569' }, minValue: 0, format: '0' })
			.build();
		sheet.insertChart(paceChart);

		var ytdChart = sheet.newChart()
			.setChartType(Charts.ChartType.COMBO)
			.addRange(sheet.getRange(HELPER_ROW, 4, ytdData.length, 3))
			.setNumHeaders(1)
			.setPosition(CHART_TOP, 7, 8, 4)
			.setOption('title', thisYear + ' Goal Progress  ·  cumulative vs target')
			.setOption('width', CHART_W)
			.setOption('height', CHART_H)
			.setOption('backgroundColor', '#FFFFFF')
			.setOption('legend', { position: 'top', textStyle: { fontName: 'Montserrat', fontSize: 11 } })
			.setOption('seriesType', 'bars')
			.setOption('series', { 0: { type: 'bars', color: '#EC4899' }, 1: { type: 'line', color: '#94A3B8', lineWidth: 2, pointSize: 4 } })
			.setOption('hAxis', { textStyle: { fontName: 'Montserrat', fontSize: 11, color: '#475569' } })
			.setOption('vAxis', { textStyle: { fontName: 'Montserrat', fontSize: 11, color: '#475569' }, minValue: 0, format: '0' })
			.build();
		sheet.insertChart(ytdChart);

		var genreChart = sheet.newChart()
			.setChartType(Charts.ChartType.PIE)
			.addRange(sheet.getRange(HELPER_ROW, 8, genreRows.length, 2))
			.setNumHeaders(1)
			.setPosition(CHART_TOP + 15, 1, 8, 4)
			.setOption('title', 'Genre Mix')
			.setOption('width', CHART_W)
			.setOption('height', CHART_H)
			.setOption('backgroundColor', '#FFFFFF')
			.setOption('pieHole', 0.45)
			.setOption('legend', { position: 'right', textStyle: { fontName: 'Montserrat', fontSize: 12 } })
			.setOption('colors', DIVERSE_PALETTE)
			.build();
		sheet.insertChart(genreChart);

		var statusChart = sheet.newChart()
			.setChartType(Charts.ChartType.PIE)
			.addRange(sheet.getRange(HELPER_ROW, 11, statusRows.length, 2))
			.setNumHeaders(1)
			.setPosition(CHART_TOP + 15, 7, 8, 4)
			.setOption('title', 'Library Status')
			.setOption('width', CHART_W)
			.setOption('height', CHART_H)
			.setOption('backgroundColor', '#FFFFFF')
			.setOption('pieHole', 0.55)
			.setOption('legend', { position: 'right', textStyle: { fontName: 'Montserrat', fontSize: 12 } })
			.setOption('colors', ['#3B82F6', '#10B981', '#F59E0B', '#94A3B8'])
			.build();
		sheet.insertChart(statusChart);
	} catch(chartErr) {
		_log('warn', '_dbLiteInitMyYearSheet charts', chartErr);
	}

	// Compact utilities — match Library dark header
	sheet.setRowHeight(UTIL_HEADER - 1, 12);
	sheet.getRange(UTIL_HEADER, 1, 1, NUM_COLS).merge()
		.setValue('   GOALS  ·  CHALLENGES  ·  SHELVES')
		.setBackground(t.headerBg).setFontColor(t.headerText || '#FFFFFF')
		.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold')
		.setHorizontalAlignment('left').setVerticalAlignment('middle');
	sheet.setRowHeight(UTIL_HEADER, 32);
	var UTIL_ROWS = 8;
	sheet.setRowHeight(UTIL_TOP, 32);         // panel title row
	for (var ur = UTIL_TOP + 1; ur < UTIL_TOP + UTIL_ROWS; ur++) {
		sheet.setRowHeight(ur, 34);
		sheet.getRange(ur, 1, 1, NUM_COLS).setBackground('#FFFFFF');
	}
	sheet.getRange(UTIL_TOP, 1, 1, NUM_COLS).setBackground('#FFFFFF');

	var chalSheet = ss.getSheetByName(SHEET_CHALLENGES);
	var challenges = chalSheet ? _sheetToObjects(chalSheet, CHALLENGE_HEADERS) : [];
	var shelfSheet = ss.getSheetByName(SHEET_SHELVES);
	var shelves = shelfSheet ? _sheetToObjects(shelfSheet, SHELF_HEADERS) : [];

	// Panel titles with colored accent stripe, matching KPI cards
	var PANEL_ACCENTS = [
		{ accent:'#6366F1', tint:'#EEF2FF', label:'YEAR GOAL' },
		{ accent:'#EC4899', tint:'#FDF2F8', label:'CHALLENGES' },
		{ accent:'#10B981', tint:'#ECFDF5', label:'SHELVES' }
	];
	var panelCols = [1, 5, 9];
	PANEL_ACCENTS.forEach(function(p, i) {
		var col = panelCols[i];
		sheet.getRange(UTIL_TOP, col, 1, 4).merge()
			.setValue('  ' + p.label)
			.setBackground(p.tint).setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold').setFontColor(p.accent)
			.setHorizontalAlignment('left').setVerticalAlignment('middle');
	});

	// YEAR GOAL panel — big number + graphical progress bar using Unicode blocks
	var bars = Math.round(goalPct / 5); // 0..20 blocks
	var goalBarFull = new Array(Math.max(0, bars) + 1).join('█') + new Array(Math.max(0, 20 - bars) + 1).join('░');
	sheet.getRange(UTIL_TOP + 1, 1, 1, 4).merge()
		.setValue('  ' + goalDone + ' / ' + yearlyGoal)
		.setFontFamily('Montserrat').setFontSize(26).setFontWeight('bold').setFontColor('#0F172A')
		.setBackground('#FFFFFF').setVerticalAlignment('middle');
	sheet.getRange(UTIL_TOP + 2, 1, 1, 4).merge()
		.setValue('  ' + goalBarFull)
		.setFontFamily('Roboto Mono').setFontSize(14).setFontColor('#6366F1')
		.setBackground('#FFFFFF').setVerticalAlignment('middle');
	sheet.getRange(UTIL_TOP + 3, 1, 1, 4).merge()
		.setValue('  ' + goalPct + '% of year goal')
		.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold').setFontColor('#6366F1')
		.setBackground('#FFFFFF').setVerticalAlignment('middle');
	sheet.getRange(UTIL_TOP + 4, 1, 1, 4).merge()
		.setValue('  Streak:  ' + streak + ' days')
		.setFontFamily('Montserrat').setFontSize(11).setFontColor('#334155')
		.setBackground('#FFFFFF').setVerticalAlignment('middle');
	sheet.getRange(UTIL_TOP + 5, 1, 1, 4).merge()
		.setValue('  Pace:     ' + (avgDays === '—' ? '—' : avgDays + ' days / book'))
		.setFontFamily('Montserrat').setFontSize(11).setFontColor('#334155')
		.setBackground('#FFFFFF').setVerticalAlignment('middle');
	sheet.getRange(UTIL_TOP + 6, 1, 1, 4).merge()
		.setValue('  Top genre:  ' + topGenre)
		.setFontFamily('Montserrat').setFontSize(11).setFontColor('#334155')
		.setBackground('#FFFFFF').setVerticalAlignment('middle');
	sheet.getRange(UTIL_TOP + 7, 1, 1, 4).merge()
		.setBackground('#FFFFFF');

	// CHALLENGES panel — stacked: bold name on top, mini bar + count below, readable font
	for (var ci = 0; ci < 3; ci++) {
		var rowR = UTIL_TOP + 1 + (ci * 2);
		var rowR2 = rowR + 1;
		if (ci < challenges.length) {
			var ch = challenges[ci];
			var curV = Number(ch.Current) || 0;
			var tarV = Math.max(1, Number(ch.Target) || 1);
			var pct = Math.min(100, Math.round(curV / tarV * 100));
			var b = Math.round(pct / 5); // 0..20 blocks
			var mini = new Array(Math.max(0, b) + 1).join('█') + new Array(Math.max(0, 20 - b) + 1).join('░');
			var nm = String(ch.Name || '').slice(0, 30);
			sheet.getRange(rowR, 5, 1, 4).merge()
				.setValue('  ' + nm + '     ' + curV + ' / ' + tarV)
				.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold').setFontColor('#0F172A')
				.setBackground('#FFFFFF').setVerticalAlignment('middle');
			sheet.getRange(rowR2, 5, 1, 4).merge()
				.setValue('  ' + mini + '   ' + pct + '%')
				.setFontFamily('Roboto Mono').setFontSize(12).setFontColor('#EC4899')
				.setBackground('#FFFFFF').setVerticalAlignment('middle');
		} else if (ci === 0) {
			sheet.getRange(rowR, 5, 2, 4).merge()
				.setValue('  No challenges yet.')
				.setFontFamily('Montserrat').setFontSize(11).setFontColor('#94A3B8')
				.setBackground('#FFFFFF').setVerticalAlignment('middle');
		} else {
			sheet.getRange(rowR, 5, 2, 4).merge().setBackground('#FFFFFF');
		}
	}
	sheet.getRange(UTIL_TOP + 7, 5, 1, 4).merge().setBackground('#FFFFFF');

	// SHELVES panel — bold name left + book count right in an emerald chip
	var shelfCountMap = {};
	allBooks.forEach(function(b) {
		var key = (b.genre || '').toLowerCase();
		shelfCountMap[key] = (shelfCountMap[key] || 0) + 1;
	});
	for (var si = 0; si < UTIL_ROWS - 1; si++) {
		var rowS = UTIL_TOP + 1 + si;
		if (si < shelves.length) {
			var sh = shelves[si];
			var nm2 = String(sh.Name || '').slice(0, 26);
			var cnt = shelfCountMap[nm2.toLowerCase()] || 0;
			sheet.getRange(rowS, 9, 1, 3).merge()
				.setValue('  ●   ' + nm2)
				.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold').setFontColor('#0F172A')
				.setBackground('#FFFFFF').setVerticalAlignment('middle');
			sheet.getRange(rowS, 12, 1, 1)
				.setValue(cnt + ' books  ')
				.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold').setFontColor('#10B981')
				.setBackground('#FFFFFF').setHorizontalAlignment('right').setVerticalAlignment('middle');
		} else if (si === 0) {
			sheet.getRange(rowS, 9, 1, 4).merge()
				.setValue('  No shelves yet.')
				.setFontFamily('Montserrat').setFontSize(11).setFontColor('#94A3B8')
				.setBackground('#FFFFFF').setVerticalAlignment('middle');
		} else {
			sheet.getRange(rowS, 9, 1, 4).merge().setBackground('#FFFFFF');
		}
	}

	// Outer borders + colored accent stripes on each panel
	PANEL_ACCENTS.forEach(function(p, i) {
		var col = panelCols[i];
		sheet.getRange(UTIL_TOP, col, UTIL_ROWS, 4)
			.setBorder(true, true, true, true, false, false, '#E2E8F0', SpreadsheetApp.BorderStyle.SOLID);
		sheet.getRange(UTIL_TOP, col, UTIL_ROWS, 1)
			.setBorder(null, true, null, null, null, null, p.accent, SpreadsheetApp.BorderStyle.SOLID_THICK);
	});

	// Full library wall — shows ALL books (not just overflow) to match web app Library tab
	sheet.setRowHeight(FULL_HEADER - 1, 12);
	sheet.getRange(FULL_HEADER, 1, 1, NUM_COLS).merge()
		.setValue('   FULL LIBRARY  ·  ' + allBooks.length + ' books  ·  hover any cover for details')
		.setBackground(t.headerBg).setFontColor(t.headerText || '#FFFFFF')
		.setFontFamily('Montserrat').setFontSize(11).setFontWeight('bold')
		.setHorizontalAlignment('left').setVerticalAlignment('middle');
	sheet.setRowHeight(FULL_HEADER, 32);

	// Start the lower library wall after the hero selection so it feels like a
	// broader catalog, not a repeated copy of the same first row of covers.
	var fullBooks = coverBooks.length > heroCount
		? coverBooks.slice(heroCount).concat(coverBooks.slice(0, heroCount))
		: coverBooks;
	if (!fullBooks.length) return;
	var totalCoverRows = Math.ceil(fullBooks.length / COV_PER_ROW);
	for (var rh = 0; rh < totalCoverRows; rh++) sheet.setRowHeight(FULL_TOP + rh, COVER_H);
	for (var fi = 0; fi < fullBooks.length; fi++) {
		var fullRow = FULL_TOP + Math.floor(fi / COV_PER_ROW);
		var fullCol = 1 + (fi % COV_PER_ROW);
		_setCoverCell(sheet.getRange(fullRow, fullCol), fullBooks[fi]);
	}
}

function _dbLiteApplyValidations(sheet) {
	var startRow = LIBRARY_DATA_ROW;
	var dataRows = 5000;
	var statusCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Status');
	var genreCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Genre');
	var ratingCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Rating');
	var formatCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Format');
	var favColV   = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Favorite');

	function list(col, values, allowInvalid) {
		sheet.getRange(startRow, col, dataRows, 1).setDataValidation(
			SpreadsheetApp.newDataValidation()
				.requireValueInList(values, true)
				.setAllowInvalid(!!allowInvalid).build()
		);
	}

	list(statusCol, ['Reading', 'Finished', 'Want to Read', 'DNF']);
	list(genreCol,  ['Romance','Fantasy','Mystery','Thriller','SciFi','Historical',
		            'Memoir','Biography','Self-Help','Nonfiction','Fiction',
		            'Horror','YA','Poetry','Classics','Literary','Graphic','Other']);
	list(ratingCol, ['★','★★','★★★','★★★★','★★★★★'], true);
	list(formatCol, ['Paperback', 'Hardcover', 'Ebook', 'Audiobook'], true);
	list(favColV,   ['♥', ''], true);
}

function _dbLiteApplyPillFormatting(sheet) {
	var startRow  = LIBRARY_DATA_ROW;
	var dataRows  = 5000;
	var statusCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Status');
	var genreCol  = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Genre');
	var ratingCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Rating');
	var formatCol = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Format');
	var favColP   = LIBRARY_DATA_COL + LIBRARY_HEADERS.indexOf('Favorite');
	var rules     = [];

	function pill(col, val, bg, fg) {
		rules.push(SpreadsheetApp.newConditionalFormatRule()
			.whenTextEqualTo(val).setBackground(bg).setFontColor(fg || '#FFFFFF').setBold(true)
			.setRanges([sheet.getRange(startRow, col, dataRows, 1)]).build());
	}
	function textOnly(col, val, fg) {
		rules.push(SpreadsheetApp.newConditionalFormatRule()
			.whenTextEqualTo(val).setFontColor(fg).setBold(true)
			.setRanges([sheet.getRange(startRow, col, dataRows, 1)]).build());
	}

	// Status
	pill(statusCol, 'Reading',      '#BFDBFE', '#1E3A8A');
	pill(statusCol, 'Finished',     '#BBF7D0', '#14532D');
	pill(statusCol, 'Want to Read', '#FED7AA', '#7C2D12');
	pill(statusCol, 'DNF',          '#E5E7EB', '#374151');

	// Genre
	pill(genreCol, 'Romance',    '#DB2777', '#FFFFFF');
	pill(genreCol, 'Fantasy',    '#7C3AED', '#FFFFFF');
	pill(genreCol, 'Mystery',    '#1D4ED8', '#FFFFFF');
	pill(genreCol, 'Thriller',   '#DC2626', '#FFFFFF');
	pill(genreCol, 'SciFi',      '#0891B2', '#FFFFFF');
	pill(genreCol, 'Historical', '#B45309', '#FFFFFF');
	pill(genreCol, 'Memoir',     '#059669', '#FFFFFF');
	pill(genreCol, 'Biography',  '#2563EB', '#FFFFFF');
	pill(genreCol, 'Self-Help',  '#EA580C', '#FFFFFF');
	pill(genreCol, 'Nonfiction', '#475569', '#FFFFFF');
	pill(genreCol, 'Fiction',    '#0F766E', '#FFFFFF');
	pill(genreCol, 'Horror',     '#1E293B', '#FFFFFF');
	pill(genreCol, 'YA',         '#BE185D', '#FFFFFF');
	pill(genreCol, 'Poetry',     '#6D28D9', '#FFFFFF');
	pill(genreCol, 'Classics',   '#92400E', '#FFFFFF');
	pill(genreCol, 'Literary',   '#334155', '#FFFFFF');
	pill(genreCol, 'Graphic',    '#6366F1', '#FFFFFF');
	pill(genreCol, 'Other',      '#64748B', '#FFFFFF');

	// Rating — gold text, no background
	textOnly(ratingCol, '★',     '#F59E0B');
	textOnly(ratingCol, '★★',    '#F59E0B');
	textOnly(ratingCol, '★★★',   '#F59E0B');
	textOnly(ratingCol, '★★★★',  '#F59E0B');
	textOnly(ratingCol, '★★★★★', '#F59E0B');

	// Format
	pill(formatCol, 'Paperback',  '#EDE9FE', '#5B21B6');
	pill(formatCol, 'Hardcover',  '#DBEAFE', '#1E40AF');
	pill(formatCol, 'Ebook',      '#CFFAFE', '#155E75');
	pill(formatCol, 'Audiobook',  '#FDF4FF', '#6B21A8');

	// Favorite — deep red heart, no background
	textOnly(favColP, '♥', '#DC2626');

	sheet.setConditionalFormatRules(rules);
}

/* =====================================================================
 *  ── STANDALONE RUNTIME ─────────────────────────────────────────────
 *  Everything below this line is the web-app engine that used to live
 *  in Code.gs. Paste this file on its own into Apps Script and you're
 *  done — no second .gs file required.
 * ===================================================================== */

/* =====================================================================
 *  PRODUCT VARIANT — set this before distributing each product template.
 *  'index'  → Product 1: Romantic  (pink/red)
 *  'index2' → Product 2: Horizon   (blue)
 *  'index3' → Product 3: Blossom   (mauve)
 * ===================================================================== */
var PRODUCT_VARIANT = 'index';

// ── View → default theme mapping (set once on first open per deployment) ──
var _VIEW_THEME_MAP = { 'index': 'romantic', 'index2': 'horizon', 'index3': 'blossom' };

// ── Serve the UI ────────────────────────────────────────────────────────
function doGet(e) {
	var view = _resolveWebAppView(e);
	// Seed the product-specific default theme once on first access so the sheet
	// initialises with the right palette for this product variant.
	var props = PropertiesService.getScriptProperties();
	if (!props.getProperty('PRODUCT_THEME_SEEDED')) {
		props.setProperty('PRODUCT_DEFAULT_THEME', _VIEW_THEME_MAP[view] || 'romantic');
		props.setProperty('PRODUCT_THEME_SEEDED', '1');
	}
	if (props.getProperty('SHEETS_INITIALIZED') !== '1') {
		try { _dbLiteInitializeSheets(); props.setProperty('SHEETS_INITIALIZED', '1'); } catch(eSeed) {}
	}
	var title = _buildJourneyTitle();
	var output = HtmlService.createHtmlOutputFromFile(view)
		.setTitle(title)
		.addMetaTag('viewport', 'width=device-width, initial-scale=1');
	try {
		var xfMode = HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.SAMEORIGIN;
		if (xfMode != null) output.setXFrameOptionsMode(xfMode);
	} catch (e) {}
	return output;
}

function _resolveWebAppView(e) {
	var params = (e && e.parameter) || {};
	var requested = String(params.view || params.theme || params.variant || '').toLowerCase().trim();
	var views = {
		'1': 'index',
		'index': 'index',
		'index.html': 'index',
		'2': 'index2',
		'index2': 'index2',
		'index2.html': 'index2',
		'3': 'index3',
		'index3': 'index3',
		'index3.html': 'index3'
	};
	return views[requested] || PRODUCT_VARIANT || 'index';
}

function _buildJourneyTitle() {
	var sheet = _ss().getSheetByName(SHEET_PROFILE);
	var row = _getProfileDataRow(sheet);
	if (sheet && row >= 2) {
		var name = String(sheet.getRange(row, 1).getValue() || '').trim();
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
var HIDDEN_HEADER_ROW = 2;
var HIDDEN_DATA_ROW   = 3;

// ── Library layout constants ─────────────────────────────────────────────
// Col A = auto row-number formula. Data starts at column B (LIBRARY_DATA_COL).
// Rows 1-7 = banner/image (never touched by code). Row 8 = headers. Row 9+ = data.
var LIBRARY_DATA_COL      = 2;   // Column B
var LIBRARY_HEADER_ROW    = 8;
var LIBRARY_DATA_ROW      = 9;
var LIBRARY_VISIBLE_COUNT = 11;  // Title → Favorite (cols B–L)

var LIBRARY_HEADERS = [
	// ── Visible columns (B–L, 11 total) ──────────────────────────────────
	'Title','Author','Series','Status','Genre','Rating','Format','Pages',
	'DateStarted','DateFinished','Favorite',
	// ── Hidden columns (M onward — webapp internals) ──────────────────────
	'SeriesNumber',
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
	// Product 1 (index.html) — romantic / pink-red
	romantic:  { header: '#C81464', headerText: '#FFFFFF', accent: '#DC2626', border: '#FCA5A5' },
	spicy:     { header: '#8F001F', headerText: '#FFE4E8', accent: '#FF4D4D', border: '#5A001A' },
	dreamy:    { header: '#7C3AED', headerText: '#FFFFFF', accent: '#A78BFA', border: '#DDD6FE' },
	velvet:    { header: '#7E22CE', headerText: '#F5F3FF', accent: '#D946EF', border: '#3B1F6E' },
	champagne: { header: '#7A5500', headerText: '#FFFFFF', accent: '#FACC15', border: '#FDE68A' },
	bunny:     { header: '#8A0050', headerText: '#FFFFFF', accent: '#FF007F', border: '#FCE7F3' },
	// Product 2 (index2.html) — horizon / blue
	horizon:   { header: '#0369A1', headerText: '#FFFFFF', accent: '#22D3EE', border: '#BAE6FD' },
	arctic:    { header: '#0C2A4A', headerText: '#E0F2FE', accent: '#38BDF8', border: '#1E3A5F' },
	sahara:    { header: '#B45309', headerText: '#FFFFFF', accent: '#F59E0B', border: '#FDE68A' },
	ember:     { header: '#122010', headerText: '#E8DFC8', accent: '#D4A030', border: '#1E3420' },
	volcano:   { header: '#B91C1C', headerText: '#FFFFFF', accent: '#F59E0B', border: '#FCA5A5' },
	dusk:      { header: '#1A0525', headerText: '#FFF0F5', accent: '#FFCA0A', border: '#6A1F5C' },
	// Product 3 (index3.html) — blossom / mauve
	blossom:   { header: '#C85888', headerText: '#FFFFFF', accent: '#F8D4A0', border: '#F7C9D5' },
	lavenderhaze:{ header: '#7858B0', headerText: '#FFFFFF', accent: '#C4B5FD', border: '#DDD6FE' },
	sorbet:    { header: '#904428', headerText: '#FFFFFF', accent: '#FB923C', border: '#FED7AA' },
	cloud:     { header: '#1C6094', headerText: '#FFFFFF', accent: '#BAE6FD', border: '#DBEAFE' },
	meadow:    { header: '#3A4E30', headerText: '#FFFFFF', accent: '#A3E635', border: '#D9F99D' },
	sherbet:   { header: '#0D5044', headerText: '#FFFFFF', accent: '#7048A0', border: '#B8E4D8' },
	// Legacy / fallback themes
	obsidian:  { header: '#111827', headerText: '#F9FAFB', accent: '#6366F1', border: '#1F2937' },
	pearl:     { header: '#2A3A52', headerText: '#FFFFFF', accent: '#D1D5DB', border: '#E5E7EB' },
	onyx:      { header: '#1E293B', headerText: '#F8FAFC', accent: '#64748B', border: '#334155' },
	fresh:     { header: '#065F46', headerText: '#FFFFFF', accent: '#34D399', border: '#A7F3D0' },
	midnight:  { header: '#312E81', headerText: '#E0E7FF', accent: '#818CF8', border: '#334155' },
	sunset:    { header: '#1A0525', headerText: '#FFF0F5', accent: '#FF6D00', border: '#6A1F5C' },
	petal:     { header: '#9A3060', headerText: '#FFFFFF', accent: '#F67280', border: '#F9B2D7' },
	coral:     { header: '#355C7D', headerText: '#FDF2F8', accent: '#F8B195', border: '#3D5A7A' },
	lagoon:    { header: '#0E6670', headerText: '#FFFFFF', accent: '#48CFCB', border: '#A2D5C6' },
	'mint mist':{ header: '#0E6670', headerText: '#FFFFFF', accent: '#48CFCB', border: '#A2D5C6' },
	jade:      { header: '#237227', headerText: '#F0FDF4', accent: '#CFFFE2', border: '#1A3D22' },
	'sage forest': { header: '#237227', headerText: '#F0FDF4', accent: '#CFFFE2', border: '#1A3D22' },
	opal:      { header: '#0369A1', headerText: '#FFFFFF', accent: '#67E8F9', border: '#BAE6FD' }
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
	return _VALID_THEME_KEYS[t] ? t : 'romantic';
}

function _getOrCreateSheet(name, headers) {
	var ss = _ss();
	var sheet = ss.getSheetByName(name);
	var displayRow = _displayHeaders(headers);
	if (!sheet) {
		sheet = ss.insertSheet(name);
		_ensureColumns(sheet, displayRow.length);
		// Library rows 1-8 are built entirely by _dbLiteInitLibrarySheet — never touch row 1.
		if (name !== SHEET_LIBRARY) {
			sheet.getRange(1, 1, 1, displayRow.length).setValues([displayRow]);
			sheet.setFrozenRows(1);
		}
	} else {
		_ensureColumns(sheet, displayRow.length);
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
	if (data.length < HIDDEN_DATA_ROW) return [];
	var headers = internalHeaders || data[0];
	return data.slice(HIDDEN_DATA_ROW - 1).filter(function(row) {
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
	for (var r = HIDDEN_DATA_ROW - 1; r < data.length; r++) {
		if (String(data[r][colIndex]) === String(value)) return r + 1;
	}
	return -1;
}

function _getProfileDataRow(sheet) {
	if (!sheet) return -1;
	if (sheet.getLastRow() >= HIDDEN_DATA_ROW) {
		var row3 = sheet.getRange(HIDDEN_DATA_ROW, 1, 1, PROFILE_HEADERS.length).getValues()[0];
		var hasRow3Data = row3.some(function(v) { return v !== '' && v !== null; });
		if (hasRow3Data) return HIDDEN_DATA_ROW;
	}
	if (sheet.getLastRow() >= HIDDEN_HEADER_ROW) {
		var row2 = sheet.getRange(HIDDEN_HEADER_ROW, 1, 1, PROFILE_HEADERS.length).getValues()[0];
		var displayHeaders = _displayHeaders(PROFILE_HEADERS);
		var looksLikeHeaders = row2.every(function(v, i) {
			return String(v || '') === String(displayHeaders[i] || '');
		});
		var hasRow2Data = row2.some(function(v) { return v !== '' && v !== null; });
		if (hasRow2Data && !looksLikeHeaders) return HIDDEN_HEADER_ROW;
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
	return THEME_PALETTES[String(themeName || '').toLowerCase()] || THEME_PALETTES.romantic;
}

function _getCurrentTheme() {
	var sheet = _ss().getSheetByName(SHEET_PROFILE);
	var row = _getProfileDataRow(sheet);
	if (!sheet || row < 2) return 'romantic';
	var themeCol = PROFILE_HEADERS.indexOf('Theme') + 1;
	return String(sheet.getRange(row, themeCol).getValue() || 'romantic').toLowerCase();
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
			Favorite: row.Favorite === true || String(row.Favorite).toUpperCase() === 'TRUE' || String(row.Favorite).indexOf('♥') >= 0,
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
	var profileRow2 = _getProfileDataRow(profileSheet);
	if (profileSheet && profileRow2 >= 2) {
		var pRow = profileSheet.getRange(profileRow2, 1, 1, PROFILE_HEADERS.length).getValues()[0];
		PROFILE_HEADERS.forEach(function(h, i) { profileData[h] = pRow[i]; });
	}

	var settings = {
		Theme: String(profileData.Theme || 'romantic'),
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
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
		var sheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);
		var bookId = _uuid();
		var row = _bookPayloadToRow(bookId, payload);
		// Append to first empty data row — _nextLibDataRow scans Title column since
		// getLastRow() is inflated by the 5000 pre-filled col-A formulas.
		var nextRow = _nextLibDataRow(sheet);
		sheet.getRange(nextRow, LIBRARY_DATA_COL, 1, LIBRARY_HEADERS.length).setValues([row]);
		return { BookId: bookId };
	} catch (e) { return { error: e.message }; }
	finally { lock.releaseLock(); }
}

function clientUpdateBook(bookId, updates) {
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
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
	finally { lock.releaseLock(); }
}

function clientDeleteBook(bookId) {
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
		if (!_validateId(bookId)) return;
		var sheet = _ss().getSheetByName(SHEET_LIBRARY);
		if (!sheet) return;
		var rowIdx = _findLibRowByBookId(sheet, bookId);
		if (rowIdx >= LIBRARY_DATA_ROW) sheet.deleteRow(rowIdx);
	} catch (e) { return { error: e.message }; }
	finally { lock.releaseLock(); }
}

function _bookPayloadToRow(bookId, p) {
	function _cap(v, n) { return String(v || '').slice(0, n); }
	// Combine Series + SeriesNumber into one visible column (e.g. "ACOTAR #1")
	var _serName = String(p.Series || p.series || '').trim();
	var _serNum = p.SeriesOrder || p.SeriesNumber || p.seriesNumber || '';
	// Strip trailing " #N" if the caller already pre-combined the string
	var _serNameClean = _serName.replace(/\s+#\d+$/, '');
	var _serDisplay = _serNameClean + (_serNum && _serNameClean ? ' #' + _serNum : _serNameClean ? '' : '');
	// Order must match LIBRARY_HEADERS exactly
	return [
		_cap(p.Title, 500),
		_cap(p.Author, 300),
		_cap(_serDisplay, 200),                        // Series (visible, combined)
		p.Status ? _uiStatusToSheet(p.Status) : 'Want to Read',
		_cap(p.Genres || p.Genre, 200),
		_numToStars(p.Rating),
		_cap(p.Format, 100),
		Number(p.PageCount || p.Pages) || 0,
		p.DateStarted || '',
		p.DateFinished || '',
		p.Favorite === true || p.Favorite === 'true', // Favorite (visible)
		// Hidden columns below
		p.SeriesOrder || p.SeriesNumber || '',         // SeriesNumber
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
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
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
	finally { lock.releaseLock(); }
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
		if (rowIdx >= HIDDEN_DATA_ROW) sheet.deleteRow(rowIdx);
	} catch (e) {}
}
function clientRenameShelf(shelfId, newName) {
	try {
		if (!_validateId(shelfId)) return;
		var sheet = _ss().getSheetByName(SHEET_SHELVES); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, shelfId);
		if (rowIdx >= HIDDEN_DATA_ROW) sheet.getRange(rowIdx, 2).setValue(String(newName || '').slice(0, 200));
	} catch (e) {}
}
function clientUpdateShelf(shelfId, updates) {
	try {
		if (!updates || !_validateId(shelfId)) return;
		var sheet = _ss().getSheetByName(SHEET_SHELVES); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, shelfId);
		if (rowIdx < HIDDEN_DATA_ROW) return;
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
		var rowIdx = _findRowByCol(sheet, 0, challengeId); if (rowIdx < HIDDEN_DATA_ROW) return;
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
		if (rowIdx >= HIDDEN_DATA_ROW) sheet.deleteRow(rowIdx);
	} catch (e) {}
}
function clientSyncChallenges(challengesArray) {
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
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
		if (sheet.getLastRow() >= HIDDEN_DATA_ROW) {
			sheet.getRange(HIDDEN_DATA_ROW, 1, sheet.getLastRow() - HIDDEN_DATA_ROW + 1, CHALLENGE_HEADERS.length).clearContent();
		}
		sheet.getRange(HIDDEN_DATA_ROW, 1, rows.length, CHALLENGE_HEADERS.length).setValues(rows);
	} catch (e) {}
	finally { lock.releaseLock(); }
}

// ── Settings / Profile ──────────────────────────────────────────────────
function clientSetSetting(key, value) {
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
		var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
		if (_getProfileDataRow(sheet) < HIDDEN_DATA_ROW) _dbLiteEnsureProfileDefaults(sheet);
		var colIdx = PROFILE_HEADERS.indexOf(key);
		if (colIdx < 0) return;
		var safeValue = (key === 'Theme') ? _validateTheme(value) : value;
		var row = _getProfileDataRow(sheet);
		sheet.getRange(row < HIDDEN_DATA_ROW ? HIDDEN_DATA_ROW : row, colIdx + 1).setValue(safeValue);
	} catch (e) {
	} finally {
		lock.releaseLock();
	}
}
function clientSetSettings(settingsObj) {
	try {
		if (!settingsObj || typeof settingsObj !== 'object') return;
		Object.keys(settingsObj).forEach(function(k) { clientSetSetting(k, settingsObj[k]); });
	} catch (e) {}
}

function clientSaveProfile(profileData) {
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
		var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
		if (_getProfileDataRow(sheet) < HIDDEN_DATA_ROW) _dbLiteEnsureProfileDefaults(sheet);
		var row = _getProfileDataRow(sheet);
		if (row < HIDDEN_DATA_ROW) row = HIDDEN_DATA_ROW;
		var mapping = { 'name':'Name', 'motto':'Motto', 'photoData':'PhotoData' };
		Object.keys(mapping).forEach(function(k) {
			if (profileData[k] !== undefined) {
				var colIdx = PROFILE_HEADERS.indexOf(mapping[k]);
				if (colIdx >= 0) {
					var val = (k === 'photoData')
						? String(profileData[k] || '').slice(0, 49000)
						: String(profileData[k] || '').slice(0, 500);
					sheet.getRange(row, colIdx + 1).setValue(val);
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
	finally { lock.releaseLock(); }
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
		if (_getProfileDataRow(sheet) < HIDDEN_DATA_ROW) _dbLiteEnsureProfileDefaults(sheet);
		var colIdx = PROFILE_HEADERS.indexOf('ReadingOrder');
		var row = _getProfileDataRow(sheet);
		if (colIdx >= 0) sheet.getRange(row < HIDDEN_DATA_ROW ? HIDDEN_DATA_ROW : row, colIdx + 1).setValue(JSON.stringify(orderArray));
	} catch (e) {}
}
function clientSaveRecentIds(idsArray) {
	try {
		if (!Array.isArray(idsArray)) return;
		var sheet = _getOrCreateSheet(SHEET_PROFILE, PROFILE_HEADERS);
		if (_getProfileDataRow(sheet) < HIDDEN_DATA_ROW) _dbLiteEnsureProfileDefaults(sheet);
		var colIdx = PROFILE_HEADERS.indexOf('RecentIds');
		var row = _getProfileDataRow(sheet);
		if (colIdx >= 0) sheet.getRange(row < HIDDEN_DATA_ROW ? HIDDEN_DATA_ROW : row, colIdx + 1).setValue(JSON.stringify(idsArray));
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
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
		var sheet = _getOrCreateSheet(SHEET_AUDIOBOOKS, AUDIOBOOK_HEADERS);
		var existingRow = _findRowByCol(sheet, 0, audioData.id);
		var row = [
			audioData.id || _uuid(), audioData.title || '', audioData.author || '',
			audioData.duration || '', audioData.cover || 'AUDIO', audioData.coverUrl || '',
			Number(audioData.chapterCount) || 0, audioData.audiobookId || '',
			Number(audioData.chapterIndex) || 0, Number(audioData.currentTime) || 0,
			Number(audioData.speed) || 1, Number(audioData.totalListeningMins) || 0
		];
		if (existingRow >= HIDDEN_DATA_ROW) {
			sheet.getRange(existingRow, 1, 1, AUDIOBOOK_HEADERS.length).setValues([row]);
		} else {
			sheet.appendRow(row);
		}
	} catch (e) { return { error: e.message }; }
	finally { lock.releaseLock(); }
}
function clientSaveAudioPosition(audioId, chapterIndex, currentTime, speed, totalListeningMins) {
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
		if (!_validateId(audioId)) return;
		var sheet = _ss().getSheetByName(SHEET_AUDIOBOOKS); if (!sheet) return;
		var rowIdx = _findRowByCol(sheet, 0, String(audioId)); if (rowIdx < HIDDEN_DATA_ROW) return;
		var row = sheet.getRange(rowIdx, 1, 1, AUDIOBOOK_HEADERS.length).getValues()[0];
		row[AUDIOBOOK_HEADERS.indexOf('CurrentChapterIndex')] = Number(chapterIndex) || 0;
		row[AUDIOBOOK_HEADERS.indexOf('CurrentTime')]         = Number(currentTime) || 0;
		row[AUDIOBOOK_HEADERS.indexOf('PlaybackSpeed')]       = Number(speed) || 1;
		if (totalListeningMins !== undefined && totalListeningMins !== null) {
			row[AUDIOBOOK_HEADERS.indexOf('TotalListeningMins')] = Number(totalListeningMins) || 0;
		}
		sheet.getRange(rowIdx, 1, 1, AUDIOBOOK_HEADERS.length).setValues([row]);
	} catch (e) {}
	finally { lock.releaseLock(); }
}

// ── Demo Data ───────────────────────────────────────────────────────────
function _clearSheetDataRows(sheet, headers) {
	if (!sheet || sheet.getLastRow() < HIDDEN_DATA_ROW) return;
	sheet.getRange(HIDDEN_DATA_ROW, 1, sheet.getLastRow() - HIDDEN_DATA_ROW + 1, headers.length).clearContent();
}

function _seedDemoData() {
	// Respect the "user has cleared demo data" flag so we never resurrect
	// sample books after a user has explicitly emptied their library.
	if (PropertiesService.getScriptProperties().getProperty('DEMO_CLEARED') === '1') return;

	var libSheet = _ss().getSheetByName(SHEET_LIBRARY);
	if (!libSheet) libSheet = _getOrCreateSheet(SHEET_LIBRARY, LIBRARY_HEADERS);

	// Inspect any existing rows. We normally do not reseed over existing data,
	// but we do allow a one-time safe upgrade from the original 26-book demo set
	// to the expanded 72-book demo catalog.
	var existingTitles = [];
	var lastRow = libSheet.getLastRow();
	if (lastRow >= LIBRARY_DATA_ROW) {
		var titleVals = libSheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, lastRow - LIBRARY_DATA_ROW + 1, 1).getValues();
		for (var i = 0; i < titleVals.length; i++) {
			var existingTitle = String(titleVals[i][0] || '').trim();
			if (existingTitle) existingTitles.push(existingTitle);
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
		{ t:'Circe', a:'Madeline Miller', g:'Fantasy', isbn:'9780316556347', pg:393, r:5, stat:'Finished', da:_weeksAgo(10,0), df:_monthDate(2,14), fmt:'Hardcover', ser:'Greek Myths', sn:2 },
		{ t:'The Silent Patient', a:'Alex Michaelides', g:'Thriller', isbn:'9781250301697', pg:325, r:4, stat:'Finished', da:_weeksAgo(9,4), df:_monthDate(1,3), fmt:'Paperback' },
		{ t:'Atomic Habits', a:'James Clear', g:'Self-Help', isbn:'9780735211292', pg:306, r:5, stat:'Finished', da:_weeksAgo(8,1), df:_monthDate(5,22), fmt:'Audiobook' },
		{ t:'The Song of Achilles', a:'Madeline Miller', g:'Fantasy', isbn:'9780062060624', pg:352, r:5, stat:'Finished', da:_weeksAgo(0,0), df:_monthDate(2,27), fmt:'Paperback', ser:'Greek Myths', sn:1 },
		{ t:'Where the Crawdads Sing', a:'Delia Owens', g:'Mystery', isbn:'9780735224292', pg:368, r:5, stat:'Finished', da:_weeksAgo(6,0), df:_monthDate(1,27), fmt:'Hardcover' },
		{ t:'Project Hail Mary', a:'Andy Weir', g:'SciFi', isbn:'9780593135204', pg:476, r:5, stat:'Finished', da:_weeksAgo(4,5), df:_monthDate(1,20), fmt:'Ebook' },
		{ t:'The Guest List', a:'Lucy Foley', g:'Mystery', isbn:'9780062868930', pg:312, r:4, stat:'Finished', da:_weeksAgo(8,5), df:_monthDate(4,24), fmt:'Paperback', ser:'Foley Mysteries', sn:2 },
		{ t:'Educated', a:'Tara Westover', g:'Memoir', isbn:'9780399590504', pg:334, r:5, stat:'Finished', da:_weeksAgo(3,3), df:_monthDate(1,8), fmt:'Hardcover' },
		{ t:'The Invisible Life of Addie LaRue', a:'V.E. Schwab', g:'Fantasy', isbn:'9780765387561', pg:448, r:5, stat:'Finished', da:_weeksAgo(1,4), df:_monthDate(2,25), fmt:'Paperback', ser:'Addie LaRue', sn:1 },
		{ t:'The Vanishing Half', a:'Brit Bennett', g:'Fiction', isbn:'9780525536291', pg:343, r:4, stat:'Finished', da:_weeksAgo(4,1), df:_monthDate(1,14), fmt:'Hardcover' },
		{ t:'Verity', a:'Colleen Hoover', g:'Thriller', isbn:'9781538724736', pg:374, r:5, stat:'Finished', da:_weeksAgo(7,3), df:_monthDate(4,3), fmt:'Paperback' },
		{ t:'Book Lovers', a:'Emily Henry', g:'Romance', isbn:'9780593334836', pg:368, r:5, stat:'Finished', da:_weeksAgo(3,0), df:_monthDate(2,4), fmt:'Paperback' },
		{ t:'The Spanish Love Deception', a:'Elena Armas', g:'Romance', isbn:'9781982177010', pg:358, r:4, stat:'Finished', da:_weeksAgo(5,2), df:_monthDate(1,10), fmt:'Ebook' },
		{ t:'A Court of Thorns and Roses', a:'Sarah J. Maas', g:'Fantasy', isbn:'9781635575569', pg:419, r:5, stat:'Finished', da:_weeksAgo(6,4), df:_monthDate(3,19), fmt:'Paperback', ser:'ACOTAR', sn:1 },
		{ t:'The Thursday Murder Club', a:'Richard Osman', g:'Mystery', isbn:'9781984880963', pg:369, r:4, stat:'Finished', da:_weeksAgo(2,5), df:_monthDate(2,20), fmt:'Hardcover', ser:'Thursday Murder Club', sn:1 },
		{ t:'The Four Winds', a:'Kristin Hannah', g:'Historical', isbn:'9781250178602', pg:454, r:5, stat:'Finished', da:_weeksAgo(2,2), df:_monthDate(5,8), fmt:'Hardcover' },
		{ t:'Normal People', a:'Sally Rooney', g:'Romance', isbn:'9781984822185', pg:266, r:4, stat:'Finished', da:_weeksAgo(9,4), df:_monthDate(5,15), fmt:'Paperback' },
		{ t:'The House in the Cerulean Sea', a:'TJ Klune', g:'Fantasy', isbn:'9781250217288', pg:396, r:5, stat:'Finished', da:_weeksAgo(6,0), df:_monthDate(4,29), fmt:'Paperback', ser:'Cerulean Chronicles', sn:1 },
		{ t:'Malibu Rising', a:'Taylor Jenkins Reid', g:'Fiction', isbn:'9780593158203', pg:369, r:5, stat:'Finished', da:_weeksAgo(0,4), df:_monthDate(3,5), fmt:'Hardcover' },
		{ t:'The Love Hypothesis', a:'Ali Hazelwood', g:'Romance', isbn:'9780593336823', pg:357, r:4, stat:'Finished', da:_weeksAgo(2,2), df:_monthDate(1,25), fmt:'Paperback' },
		{ t:'Daisy Jones & The Six', a:'Taylor Jenkins Reid', g:'Fiction', isbn:'9781524798628', pg:368, r:5, stat:'Finished', da:_weeksAgo(1,1), df:_monthDate(3,12), fmt:'Audiobook' },
		{ t:'The Atlas Six', a:'Olivie Blake', g:'Fantasy', isbn:'9781250854513', pg:374, r:4, stat:'Finished', da:_weeksAgo(3,3), df:_monthDate(3,26), fmt:'Paperback', ser:'The Atlas Six', sn:1 },
		{ t:'Red, White & Royal Blue', a:'Casey McQuiston', g:'Romance', isbn:'9781250316776', pg:352, r:5, stat:'Finished', da:_weeksAgo(0,2), df:_monthDate(2,20), fmt:'Paperback' },
		{ t:'It Ends With Us', a:'Colleen Hoover', g:'Romance', isbn:'9781501110375', pg:376, r:5, stat:'Reading', da:_weeksAgo(0,0), ds:'2026-03-22', cp:169, fmt:'Paperback' },
		{ t:'The Midnight Library', a:'Matt Haig', g:'Fiction', isbn:'9780525559474', pg:304, r:4, stat:'Reading', da:_weeksAgo(0,2), ds:'2026-03-01', cp:249, fmt:'Hardcover' },
		{ t:'Lessons in Chemistry', a:'Bonnie Garmus', g:'Fiction', isbn:'9780385547345', pg:400, r:0, stat:'Want to Read', da:_weeksAgo(10,1), fmt:'Hardcover' },
		{ t:'The Seven Husbands of Evelyn Hugo', a:'Taylor Jenkins Reid', g:'Fiction', isbn:'9781501156717', pg:400, r:0, stat:'Want to Read', da:_weeksAgo(9,2), fmt:'Paperback' },
		{ t:'Tomorrow, and Tomorrow, and Tomorrow', a:'Gabrielle Zevin', g:'Fiction', isbn:'9780593321201', pg:416, r:0, stat:'Want to Read', da:_weeksAgo(8,3), fmt:'Hardcover' },
		{ t:'Happy Place', a:'Emily Henry', g:'Romance', isbn:'9780593441282', pg:400, r:0, stat:'Want to Read', da:_weeksAgo(7,2), fmt:'Paperback' },
		{ t:'Fourth Wing', a:'Rebecca Yarros', g:'Fantasy', isbn:'9781649374042', pg:528, r:0, stat:'Reading', da:_weeksAgo(2,1), ds:'2026-04-08', cp:221, fmt:'Hardcover', ser:'The Empyrean', sn:1 },
		{ t:'Demon Copperhead', a:'Barbara Kingsolver', g:'Fiction', isbn:'9780063251922', pg:560, r:0, stat:'Want to Read', da:_weeksAgo(6,0), fmt:'Hardcover' },
		{ t:'The Covenant of Water', a:'Abraham Verghese', g:'Historical', isbn:'9780802162175', pg:736, r:0, stat:'Want to Read', da:_weeksAgo(5,4), fmt:'Hardcover' },
		{ t:'Hell Bent', a:'Leigh Bardugo', g:'Fantasy', isbn:'9781250313072', pg:496, r:0, stat:'Want to Read', da:_weeksAgo(7,5), fmt:'Hardcover', ser:'Alex Stern', sn:2 },
		{ t:'Babel', a:'R.F. Kuang', g:'Fantasy', isbn:'9780063021426', pg:560, r:0, stat:'Want to Read', da:_weeksAgo(11,5), fmt:'Hardcover' },
		{ t:'Gone Girl', a:'Gillian Flynn', g:'Thriller', isbn:'9780307588364', pg:432, r:0, stat:'Want to Read', da:_weeksAgo(4,0), fmt:'Paperback' },
		{ t:'The Women', a:'Kristin Hannah', g:'Historical', isbn:'9781250178619', pg:480, r:0, stat:'Want to Read', da:_weeksAgo(1,5), fmt:'Hardcover' },
		{ t:'All the Light We Cannot See', a:'Anthony Doerr', g:'Historical', isbn:'9781501156700', pg:544, r:0, stat:'Want to Read', da:_weeksAgo(10,4), fmt:'Paperback' },
		{ t:'The Hunger Games', a:'Suzanne Collins', g:'SciFi', isbn:'9780439023481', pg:374, r:0, stat:'Want to Read', da:_weeksAgo(9,0), fmt:'Paperback', ser:'The Hunger Games', sn:1 },
		{ t:'Pachinko', a:'Min Jin Lee', g:'Historical', isbn:'9781455563920', pg:496, r:0, stat:'Want to Read', da:_weeksAgo(8,4), fmt:'Paperback' },
		{ t:'The Sympathizer', a:'Viet Thanh Nguyen', g:'Fiction', isbn:'9780802123459', pg:384, r:0, stat:'Want to Read', da:_weeksAgo(6,2), fmt:'Paperback' },
		{ t:'Remarkably Bright Creatures', a:'Shelby Van Pelt', g:'Fiction', isbn:'9780063204157', pg:368, r:0, stat:'Want to Read', da:_weeksAgo(5,0), fmt:'Hardcover' },
		{ t:'Funny Story', a:'Emily Henry', g:'Romance', isbn:'9780593441299', pg:400, r:0, stat:'Reading', da:_weeksAgo(1,1), ds:'2026-04-14', cp:118, fmt:'Paperback' },
		{ t:'Iron Flame', a:'Rebecca Yarros', g:'Fantasy', isbn:'9781649374066', pg:640, r:0, stat:'Reading', da:_weeksAgo(1,3), ds:'2026-04-11', cp:302, fmt:'Hardcover', ser:'The Empyrean', sn:2 },
		{ t:'James', a:'Percival Everett', g:'Fiction', isbn:'9780385550369', pg:320, r:0, stat:'Want to Read', da:_weeksAgo(3,5), fmt:'Hardcover' },
		{ t:'Parable of the Sower', a:'Octavia E. Butler', g:'SciFi', isbn:'9781538732182', pg:368, r:0, stat:'Want to Read', da:_weeksAgo(7,0), fmt:'Paperback', ser:'Earthseed', sn:1 },
		{ t:'A Little Life', a:'Hanya Yanagihara', g:'Fiction', isbn:'9780804172707', pg:832, r:0, stat:'Want to Read', da:_weeksAgo(5,2), fmt:'Paperback' },
		{ t:'The Night Circus', a:'Erin Morgenstern', g:'Fantasy', isbn:'9780385534635', pg:512, r:0, stat:'Want to Read', da:_weeksAgo(4,3), fmt:'Paperback' },
		{ t:'The Nightingale', a:'Kristin Hannah', g:'Historical', isbn:'9781250080400', pg:608, r:0, stat:'Want to Read', da:_weeksAgo(8,1), fmt:'Paperback' },
		{ t:'Sea of Tranquility', a:'Emily St. John Mandel', g:'SciFi', isbn:'9780593321447', pg:272, r:0, stat:'Want to Read', da:_weeksAgo(2,4), fmt:'Hardcover' },
		{ t:'Station Eleven', a:'Emily St. John Mandel', g:'SciFi', isbn:'9780804172448', pg:352, r:0, stat:'Want to Read', da:_weeksAgo(11,0), fmt:'Paperback' },
		{ t:'The Hobbit', a:'J.R.R. Tolkien', g:'Fantasy', isbn:'9780547928227', pg:300, r:0, stat:'Want to Read', da:_weeksAgo(3,2), fmt:'Paperback', ser:'Middle-earth', sn:0 },
		{ t:'Dune', a:'Frank Herbert', g:'SciFi', isbn:'9780441172719', pg:896, r:0, stat:'Want to Read', da:_weeksAgo(6,5), fmt:'Paperback', ser:'Dune', sn:1 },
		{ t:'The Martian', a:'Andy Weir', g:'SciFi', isbn:'9780553418026', pg:400, r:0, stat:'Want to Read', da:_weeksAgo(10,2), fmt:'Paperback' },
		{ t:'The Book Thief', a:'Markus Zusak', g:'Historical', isbn:'9780375842207', pg:592, r:0, stat:'Want to Read', da:_weeksAgo(9,5), fmt:'Paperback' },
		{ t:'The Giver', a:'Lois Lowry', g:'SciFi', isbn:'9780544336261', pg:240, r:0, stat:'Want to Read', da:_weeksAgo(2,0), fmt:'Paperback', ser:'The Giver Quartet', sn:1 },
		{ t:'The Maid', a:'Nita Prose', g:'Mystery', isbn:'9780593356159', pg:320, r:0, stat:'Want to Read', da:_weeksAgo(7,1), fmt:'Hardcover', ser:'Molly the Maid', sn:1 },
		{ t:'The Paris Apartment', a:'Lucy Foley', g:'Mystery', isbn:'9780063003057', pg:384, r:0, stat:'Want to Read', da:_weeksAgo(4,5), fmt:'Paperback' },
		{ t:'None of This Is True', a:'Lisa Jewell', g:'Thriller', isbn:'9780593492918', pg:384, r:0, stat:'Reading', da:_weeksAgo(0,5), ds:'2026-04-18', cp:87, fmt:'Hardcover' },
		{ t:'Pineapple Street', a:'Jenny Jackson', g:'Fiction', isbn:'9780593418482', pg:320, r:0, stat:'Want to Read', da:_weeksAgo(6,3), fmt:'Hardcover' },
		{ t:'Yellowface', a:'R.F. Kuang', g:'Thriller', isbn:'9780063250833', pg:336, r:0, stat:'Want to Read', da:_weeksAgo(5,5), fmt:'Hardcover' },
		{ t:'Just for the Summer', a:'Abby Jimenez', g:'Romance', isbn:'9781538704431', pg:432, r:0, stat:'Want to Read', da:_weeksAgo(1,0), fmt:'Paperback' },
		{ t:'People We Meet on Vacation', a:'Emily Henry', g:'Romance', isbn:'9781984806758', pg:368, r:0, stat:'Want to Read', da:_weeksAgo(9,1), fmt:'Paperback' },
		{ t:'Legends & Lattes', a:'Travis Baldree', g:'Fantasy', isbn:'9781250886088', pg:304, r:0, stat:'Want to Read', da:_weeksAgo(8,0), fmt:'Paperback' },
		{ t:'The Great Gatsby', a:'F. Scott Fitzgerald', g:'Classics', isbn:'9780743273565', pg:180, r:0, stat:'Want to Read', da:_weeksAgo(7,4), fmt:'Paperback' },
		{ t:'The Priory of the Orange Tree', a:'Samantha Shannon', g:'Fantasy', isbn:'9781635570298', pg:848, r:0, stat:'Want to Read', da:_weeksAgo(3,1), fmt:'Paperback' },
		{ t:'The Fellowship of the Ring', a:'J.R.R. Tolkien', g:'Fantasy', isbn:'9780547928210', pg:432, r:0, stat:'Want to Read', da:_weeksAgo(11,4), fmt:'Paperback', ser:'The Lord of the Rings', sn:1 },
		{ t:'The Two Towers', a:'J.R.R. Tolkien', g:'Fantasy', isbn:'9780547928203', pg:352, r:0, stat:'Want to Read', da:_weeksAgo(10,3), fmt:'Paperback', ser:'The Lord of the Rings', sn:2 },
		{ t:'The Return of the King', a:'J.R.R. Tolkien', g:'Fantasy', isbn:'9780547928197', pg:416, r:0, stat:'Want to Read', da:_weeksAgo(9,3), fmt:'Paperback', ser:'The Lord of the Rings', sn:3 },
		{ t:'The Maidens', a:'Alex Michaelides', g:'Thriller', isbn:'9781250304452', pg:352, r:0, stat:'Want to Read', da:_weeksAgo(0,1), fmt:'Hardcover' },
		{ t:'The Alchemist', a:'Paulo Coelho', g:'Fiction', isbn:'9780061122415', pg:208, r:0, stat:'Want to Read', da:_weeksAgo(0,3), fmt:'Paperback' }
	];

	if (existingTitles.length) {
		var demoTitleSet = {};
		demoBooks.forEach(function(b) { demoTitleSet[b.t] = true; });
		var looksLikeOldDemo = existingTitles.length === 26 && existingTitles.every(function(t) { return !!demoTitleSet[t]; });
		if (!looksLikeOldDemo) return;

		// Replace the original demo-only library with the expanded catalog.
		libSheet.getRange(
			LIBRARY_DATA_ROW,
			LIBRARY_DATA_COL,
			lastRow - LIBRARY_DATA_ROW + 1,
			LIBRARY_HEADERS.length
		).clearContent();
	}

	var readingIds = [];
	var rows = demoBooks.map(function(b) {
		var bookId = _uuid();
		if (b.stat === 'Reading') readingIds.push(bookId);
		// Array order must match LIBRARY_HEADERS exactly:
		// Title, Author, Series, Status, Genre, Rating, Format, Pages, DateStarted, DateFinished,
		// Favorite, SeriesNumber, BookId, CoverUrl, CoverEmoji, Gradient1, Gradient2,
		// DateAdded, CurrentPage, TbrPriority, Source, SpiceLevel, Tags, Shelves, Notes,
		// Review, Quotes, ISBN, OLID, AuthorKey
		return [
			b.t,                          // Title
			b.a,                          // Author
			(b.ser || '') + (b.sn && b.ser ? ' #' + b.sn : ''), // Series (visible, combined)
			b.stat,                       // Status
			b.g,                          // Genre
			_numToStars(b.r),             // Rating (★ chip string)
			b.fmt || 'Paperback',         // Format
			b.pg,                         // Pages
			b.ds || b.da || '',           // DateStarted
			b.df || '',                   // DateFinished
			b.r === 5 ? '♥' : '',        // Favorite (visible) — ♥ for 5-star books
			// Hidden columns below
			b.sn  || '',                  // SeriesNumber
			bookId,                       // BookId
			'https://covers.openlibrary.org/b/isbn/' + b.isbn + '-L.jpg', // CoverUrl
			'BK',                         // CoverEmoji
			'',                           // Gradient1
			'',                           // Gradient2
			b.da || new Date().toISOString().slice(0, 10), // DateAdded
			b.cp || 0,                    // CurrentPage
			'',                           // TbrPriority
			'',                           // Source
			0,                            // SpiceLevel
			'',                           // Tags
			'',                           // Shelves
			'',                           // Notes
			'',                           // Review
			'',                           // Quotes
			b.isbn,                       // ISBN
			'',                           // OLID
			''                            // AuthorKey
		];
	});

	if (rows.length > 0) {
		// Write to LIBRARY_DATA_ROW (row 9), starting at LIBRARY_DATA_COL (col B = 2).
		// Col A has pre-filled =IF() formulas and is not touched here.
		libSheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, rows.length, LIBRARY_HEADERS.length).setValues(rows);
	}

	// Seed goals
	var chalSheet = _getOrCreateSheet(SHEET_CHALLENGES, CHALLENGE_HEADERS);
	if (_sheetToObjects(chalSheet, CHALLENGE_HEADERS).length === 0) {
		chalSheet.getRange(HIDDEN_DATA_ROW, 1, 3, CHALLENGE_HEADERS.length).setValues([
			[_uuid(), '50 Books Challenge', 'Books', 42, 50],
			[_uuid(), 'Read 30 Min Daily', 'Daily', 27, 30],
			[_uuid(), 'Try 5 New Authors', 'Authors', 4, 5]
		]);
	}

	// Seed shelves
	var shelfSheet = _getOrCreateSheet(SHEET_SHELVES, SHELF_HEADERS);
	if (_sheetToObjects(shelfSheet, SHELF_HEADERS).length === 0) {
		shelfSheet.getRange(HIDDEN_DATA_ROW, 1, 3, SHELF_HEADERS.length).setValues([
			[_uuid(), 'Book Club', 'Club'],
			[_uuid(), 'Comfort Reads', 'Calm'],
			[_uuid(), 'Summer TBR', 'Seasonal']
		]);
	}

	// Reading order → profile
	var profileSheet = _ss().getSheetByName(SHEET_PROFILE);
	var profileRow3 = _getProfileDataRow(profileSheet);
	if (profileSheet && profileRow3 >= 2 && readingIds.length > 0) {
		var roCol = PROFILE_HEADERS.indexOf('ReadingOrder') + 1;
		if (roCol > 0) profileSheet.getRange(profileRow3, roCol).setValue(JSON.stringify(readingIds));
	}
}

function clientClearDemoData() {
	var lock = LockService.getDocumentLock();
	try {
		lock.waitLock(10000);
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
		var profileRow = _getProfileDataRow(profileSheet);
		if (profileSheet && profileRow >= 2) {
			var resets = {
				ReadingOrder: '[]', RecentIds: '[]', SelectedFilter: 'all',
				ActiveShelf: '', SortBy: 'default', LibViewMode: 'grid',
				ChallengeBarCollapsed: false, LibToolsOpen: false
			};
			var pRow = profileSheet.getRange(profileRow, 1, 1, PROFILE_HEADERS.length).getValues()[0];
			Object.keys(resets).forEach(function(k) {
				var c = PROFILE_HEADERS.indexOf(k); if (c >= 0) pRow[c] = resets[k];
			});
			profileSheet.getRange(profileRow, 1, 1, PROFILE_HEADERS.length).setValues([pRow]);
		}
		// Rebuild the My Year dashboard so the sheet UI mirrors the now-empty web app state.
		try {
			var themeName = 'default';
			if (profileSheet && profileRow >= 2) {
				var thC = PROFILE_HEADERS.indexOf('Theme') + 1;
				if (thC > 0) themeName = String(profileSheet.getRange(profileRow, thC).getValue() || 'default');
			}
			_dbLiteInitMyYearSheet(ss, themeName);
		} catch (myErr) { _log('warn', 'clientClearDemoData myYear rebuild', myErr); }
		return { cleared: true };
	} catch (e) { return { error: e.message }; }
	finally { lock.releaseLock(); }
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
	var apiKey = PODCAST_INDEX_API_KEY || props.getProperty('PODCAST_INDEX_API_KEY');
	var apiSecret = PODCAST_INDEX_API_SECRET || props.getProperty('PODCAST_INDEX_API_SECRET');
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
	var url = 'https://librivox.org/api/feed/audiotracks?project_id=' + encodeURIComponent(projectId) + '&format=json';
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
			available: preview === 'full',
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
	return NYT_API_KEY || PropertiesService.getScriptProperties().getProperty('NYT_API_KEY') || null;
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
						description: String(b.description || '').slice(0, 120),
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
	var _nyCacheJson = JSON.stringify(cache);
	try { CacheService.getScriptCache().put('NYT_CACHE', _nyCacheJson, 86400); } catch (cE) {
		try { props.setProperty('NYT_CACHE', _nyCacheJson.slice(0, 9000)); } catch(pE) {}
	}
	props.setProperty('NYT_CACHE_DATE', new Date().toISOString().slice(0, 10));
	var _nyFeedJson = JSON.stringify({ updatedAt: new Date().toISOString().slice(0, 10), lists: currentFeed });
	try { CacheService.getScriptCache().put('NYT_FEED_CURRENT', _nyFeedJson, 86400); } catch (cE) {
		try { props.setProperty('NYT_FEED_CURRENT', _nyFeedJson.slice(0, 9000)); } catch(pE) {}
	}
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
	var raw = CacheService.getScriptCache().get('NYT_CACHE') ||
	          PropertiesService.getScriptProperties().getProperty('NYT_CACHE');
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
	var raw = CacheService.getScriptCache().get('NYT_FEED_CURRENT') ||
	          PropertiesService.getScriptProperties().getProperty('NYT_FEED_CURRENT');
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

function _scheduleInitialNYTWarmup() {
	if (!_getNytApiKey()) return;
	var triggers = ScriptApp.getProjectTriggers();
	for (var i = 0; i < triggers.length; i++) {
		if (triggers[i].getHandlerFunction() === 'runInitialNYTWarmup') return;
	}
	ScriptApp.newTrigger('runInitialNYTWarmup').timeBased().after(60 * 1000).create();
}

function runInitialNYTWarmup() {
	try { clientRefreshNYTCache(); } catch (e) { _log('warn', 'runInitialNYTWarmup', e); }
	try {
		var triggers = ScriptApp.getProjectTriggers();
		for (var i = 0; i < triggers.length; i++) {
			if (triggers[i].getHandlerFunction() === 'runInitialNYTWarmup') {
				ScriptApp.deleteTrigger(triggers[i]);
			}
		}
	} catch (cleanupErr) { _log('warn', 'runInitialNYTWarmup cleanup', cleanupErr); }
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

function resetDemoData() {
	var ui = SpreadsheetApp.getUi();
	var resp = ui.alert(
		'Reset Demo Data',
		'This will DELETE all rows in the Library sheet and re-seed fresh demo books with the correct column order.\n\nContinue?',
		ui.ButtonSet.OK_CANCEL
	);
	if (resp !== ui.Button.OK) return;

	// 1. Clear the DEMO_CLEARED flag so _seedDemoData will run again.
	PropertiesService.getScriptProperties().deleteProperty('DEMO_CLEARED');

	// 2. Wipe all data rows (row 9 onward) in the Library sheet.
	var libSheet = _ss().getSheetByName(SHEET_LIBRARY);
	if (libSheet) {
		var lastRow = libSheet.getLastRow();
		if (lastRow >= LIBRARY_DATA_ROW) {
			libSheet.getRange(LIBRARY_DATA_ROW, LIBRARY_DATA_COL, lastRow - LIBRARY_DATA_ROW + 1, LIBRARY_HEADERS.length)
				.clearContent();
		}
	}

	// 3. Re-seed and rebuild.
	_seedDemoData();
	_dbLiteInitializeSheets();

	ui.alert('Done! Demo data has been reset with the correct column order.');
}

function onOpen() {
	// First-open auto-setup: do the buyer setup automatically, but keep the
	// sheet responsive. Anything expensive gets deferred to a trigger.
	try {
		var props = PropertiesService.getScriptProperties();
		if (props.getProperty('SHEETS_INITIALIZED') !== '1') {
			// 1. Build sheet tabs, seed demo books, restyle theme.
			_dbLiteInitializeSheets();

			// 2. Install the live sheet→app sync trigger.
			try {
				var hasSyncTrigger = false;
				var existing = ScriptApp.getProjectTriggers();
				for (var i = 0; i < existing.length; i++) {
					if (existing[i].getHandlerFunction() === 'onEditSyncHandler') { hasSyncTrigger = true; break; }
				}
				if (!hasSyncTrigger) ScriptApp.newTrigger('onEditSyncHandler').forSpreadsheet(_ss()).onEdit().create();
			} catch (eSync) { _log('warn', 'onOpen', 'sync trigger: ' + eSync); }

			// 3. If NYT is enabled, install the weekly refresh trigger and queue
			//    a one-time warmup so first open is fast.
			try {
				if (_getNytApiKey()) {
					var hasNytTrigger = false;
					for (var j = 0; j < existing.length; j++) {
						if (existing[j].getHandlerFunction() === 'clientRefreshNYTCache') { hasNytTrigger = true; break; }
					}
					if (!hasNytTrigger) {
						ScriptApp.newTrigger('clientRefreshNYTCache').timeBased().everyWeeks(1)
							.onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(3).create();
					}
					_scheduleInitialNYTWarmup();
				}
			} catch (eNyt) { _log('warn', 'onOpen', 'nyt setup: ' + eNyt); }

			props.setProperty('SHEETS_INITIALIZED', '1');
		}
	} catch (e) { _log('error', 'onOpen', e); }

	try {
		var ui = SpreadsheetApp.getUi();
		var isDeployed = false;
		try { isDeployed = !!ScriptApp.getService().getUrl(); } catch(e) {}

		ui.createMenu(_buildJourneyTitle())
			.addItem(isDeployed ? '📖 Open My App' : '🚀 Set Up My App',
			         isDeployed ? '_openWebApp'    : '_setupMyApp')
			.addSeparator()
			.addItem('🎨 Refresh Styling & Colors', '_reStyleCurrentTheme')
			.addItem('🗑  Clear All Books & Data',  'clientClearDemoData')
			.addToUi();

		// Auto-popup the setup wizard on first open if not yet deployed.
		// Stored on document properties so it only fires once per buyer.
		try {
			var docProps = PropertiesService.getDocumentProperties();
			if (!isDeployed && docProps.getProperty('WELCOMED') !== '1') {
				docProps.setProperty('WELCOMED', '1');
				_setupMyApp();
			}
		} catch (eWelcome) {}
	} catch (e) {}
}

function _openWebApp() {
	var url = '';
	try { url = ScriptApp.getService().getUrl() || ''; } catch(e) {}
	if (!url) { _setupMyApp(); return; }
	var html = HtmlService.createHtmlOutput(
		'<div style="font-family:\'Google Sans\',Arial,sans-serif;padding:20px;text-align:center;">'
		+ '<div style="font-size:32px;margin-bottom:10px;">📖</div>'
		+ '<p style="margin:0 0 5px;font-size:13px;font-weight:700;color:#1f2937;">Your Reading App</p>'
		+ '<p style="margin:0 0 16px;font-size:11px;color:#9ca3af;word-break:break-all;">' + url + '</p>'
		+ '<a href="' + url + '" target="_blank" style="text-decoration:none;">'
		+ '<button style="background:#2563eb;color:#fff;border:none;border-radius:8px;'
		+ 'padding:11px 28px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit;">'
		+ 'Open My App \u2197</button></a>'
		+ '<p style="margin:12px 0 0;font-size:11px;color:#9ca3af;">Bookmark this link \u2014 it\'s yours forever</p>'
		+ '</div>'
	).setWidth(360).setHeight(190);
	SpreadsheetApp.getUi().showModalDialog(html, 'My Reading Journey');
}

function _setupMyApp() {
	var existingUrl = '';
	try { existingUrl = ScriptApp.getService().getUrl() || ''; } catch(e) {}
	if (existingUrl) { _openWebApp(); return; }
	var scriptEditorUrl = _getScriptEditorUrl();

	var html = HtmlService.createHtmlOutput(
		'<style>*{box-sizing:border-box;margin:0;padding:0}'
		+ 'body{font-family:"Google Sans",Arial,sans-serif;padding:16px 18px;font-size:13px;color:#1f2937;background:#fff}'
		+ '.hd{text-align:center;margin-bottom:14px}'
		+ '.hd .ico{font-size:32px;line-height:1;margin-bottom:6px}'
		+ '.hd h2{font-size:16px;font-weight:800;color:#111827;margin-bottom:4px}'
		+ '.hd p{color:#6b7280;font-size:11px;line-height:1.45}'
		+ '.note{margin:0 0 12px;padding:10px 11px;border-radius:10px;background:#eff6ff;border:1px solid #bfdbfe}'
		+ '.note b{display:block;color:#1d4ed8;font-size:11.5px;margin-bottom:3px}'
		+ '.note span{display:block;color:#475569;font-size:11px;line-height:1.45}'
		+ '.step{display:flex;gap:10px;padding:7px 0;border-bottom:1px solid #f3f4f6;align-items:flex-start}'
		+ '.num{flex-shrink:0;width:22px;height:22px;border-radius:50%;background:#2563eb;color:#fff;'
		+ 'display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:800}'
		+ '.st{font-weight:700;font-size:12px;margin-bottom:1px}'
		+ '.sb{color:#6b7280;font-size:11.5px;line-height:1.4}'
		+ '.sb b{color:#374151}'
		+ '.quick{display:block;width:100%;text-align:center;text-decoration:none;margin:0 0 10px;padding:10px 12px;'
		+ 'border-radius:10px;background:#e0e7ff;border:1px solid #c7d2fe;color:#3730a3;font-size:12px;font-weight:700}'
		+ '.btn{width:100%;padding:12px;background:#2563eb;color:#fff;border:none;border-radius:10px;'
		+ 'font-size:13px;font-weight:700;cursor:pointer;font-family:inherit;margin-top:12px}'
		+ '.btn:disabled{opacity:0.5;cursor:default}'
		+ '.sub{margin-top:6px;color:#6b7280;font-size:10.5px;text-align:center;line-height:1.4}'
		+ '.ok{display:none;margin-top:12px;padding:12px;background:#f0fdf4;border:1px solid #bbf7d0;'
		+ 'border-radius:8px;text-align:center}'
		+ '.ok .tick{font-size:20px;margin-bottom:5px}'
		+ '.ok p{font-size:12px;color:#166534;margin-bottom:8px;font-weight:600}'
		+ '.ok a{display:inline-block;background:#16a34a;color:#fff;text-decoration:none;'
		+ 'border-radius:6px;padding:8px 16px;font-size:12px;font-weight:700}'
		+ '.ok small{display:block;margin-top:7px;color:#6b7280;font-size:10.5px}'
		+ '#msg{min-height:14px;margin-top:7px;font-size:11px;color:#ef4444;text-align:center}'
		+ '</style>'
		+ '<div class="hd"><div class="ico">📖</div>'
		+ '<h2>Set Up Your Reading App</h2>'
		+ '<p>This is a one-time Google step. After this, you will only open your app link.</p></div>'
		+ '<div class="note"><b>Google security note</b>'
		+ '<span>If Google shows "This app isn\'t verified," click <b>Advanced</b> and continue. That warning is normal for personal Google Sheets tools.</span></div>'
		+ '<a class="quick" href="' + scriptEditorUrl + '" target="_blank">Open Setup Page ↗</a>'
		+ '<div class="step"><div class="num">1</div><div>'
		+ '<div class="st">Open the setup page</div>'
		+ '<div class="sb">Use the button above. If needed, you can also open it from the Extensions menu.</div></div></div>'
		+ '<div class="step"><div class="num">2</div><div>'
		+ '<div class="st">Create your app link</div>'
		+ '<div class="sb">Click <b>Deploy</b> &rarr; <b>New deployment</b>, then choose <b>Web app</b>.</div></div></div>'
		+ '<div class="step"><div class="num">3</div><div>'
		+ '<div class="st">Use these settings</div>'
		+ '<div class="sb"><b>Execute as:</b> Me &nbsp;&bull;&nbsp; <b>Who has access:</b> Anyone</div></div></div>'
		+ '<div class="step"><div class="num">4</div><div>'
		+ '<div class="st">Finish setup, then come back here</div>'
		+ '<div class="sb">Click <b>Deploy</b>, approve Google once, then return to this tab and press the button below.</div></div></div>'
		+ '<button class="btn" id="doneBtn" onclick="checkUrl()">Find My App Link</button>'
		+ '<div class="sub">Your NYT badges may take about a minute to appear the first time while the bestseller cache warms up.</div>'
		+ '<div id="msg"></div>'
		+ '<div class="ok" id="ok">'
		+ '<div class="tick">\uD83C\uDF89</div>'
		+ '<p>Your reading app is ready.</p>'
		+ '<a id="appLink" href="#" target="_blank">Open My Reading App \u2197</a>'
		+ '<small>Bookmark this link \u2014 it\'s yours forever.</small></div>'
		+ '<script>'
		+ 'function checkUrl(){'
		+ 'var btn=document.getElementById("doneBtn");'
		+ 'var msg=document.getElementById("msg");'
		+ 'btn.disabled=true;msg.textContent="Checking\u2026";'
		+ 'google.script.run'
		+ '.withSuccessHandler(function(u){'
		+ 'if(u){'
		+ 'document.getElementById("appLink").href=u;'
		+ 'document.getElementById("ok").style.display="block";'
		+ 'try{window.open(u,"_blank");}catch(e){}'
		+ 'btn.style.display="none";msg.textContent="";'
		+ '}else{'
		+ 'msg.textContent="Not seeing your app yet. Finish the setup step in the Google tab, then try again.";'
		+ 'btn.disabled=false;'
		+ '}})'
		+ '.withFailureHandler(function(){'
		+ 'msg.textContent="Could not find your app link yet. Try again in a few seconds.";'
		+ 'btn.disabled=false;})'
		+ '._checkDeployment();}'
		+ '<\/script>'
	).setWidth(440).setHeight(560);
	SpreadsheetApp.getUi().showModalDialog(html, 'Set Up My Reading App');
}

function _getScriptEditorUrl() {
	try {
		var sid = ScriptApp.getScriptId();
		if (!sid) return 'https://script.google.com/home';
		return 'https://script.google.com/home/projects/' + sid + '/edit';
	} catch (e) {
		return 'https://script.google.com/home';
	}
}

function _checkDeployment() {
	try { return ScriptApp.getService().getUrl() || ''; } catch(e) { return ''; }
}

function _reStyleCurrentTheme() {
	_dbLiteInitializeSheets();
	try { SpreadsheetApp.getUi().alert('Sheet styling refreshed.'); } catch (e) {}
}
