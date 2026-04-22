// ── My Reading Journey — Demo Web App ──────────────────────────────────
// This is the ONLY file needed alongside indexdemo.html.
// No Google Sheets. No database. Just serves the demo as a web page.
//
// How to deploy:
//   1. Go to script.google.com → New project
//   2. Replace the default Code.gs contents with this file's contents
//   3. Add a new HTML file: File → New → HTML file → name it "indexdemo"
//   4. Paste the full contents of indexdemo.html into that file
//   5. Deploy → New deployment → Web app
//        Execute as: Me
//        Who has access: Anyone
//   6. Copy the /exec URL and paste it on your Etsy listing
// ───────────────────────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutputFromFile('indexdemo')
    .setTitle('My Reading Journey — Live Demo')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Server-side API proxies (bypass browser CORS restrictions) ────────────────

function clientSearchBooks(query) {
  var url = 'https://openlibrary.org/search.json?q=' + encodeURIComponent(query) +
    '&limit=15&fields=key,title,author_name,author_key,first_publish_year,isbn,cover_i,number_of_pages_median,subject';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    var data = JSON.parse(resp.getContentText());
    return (data.docs || []).map(function(doc) {
      var isbn = '';
      if (doc.isbn && doc.isbn.length) {
        isbn = doc.isbn[0];
        for (var i = 0; i < doc.isbn.length; i++) {
          if (doc.isbn[i].length === 13) { isbn = doc.isbn[i]; break; }
        }
      }
      return {
        title: doc.title || '',
        author: (doc.author_name && doc.author_name.length) ? doc.author_name[0] : 'Unknown Author',
        authorKey: (doc.author_key && doc.author_key.length) ? doc.author_key[0] : '',
        year: doc.first_publish_year || '',
        isbn: isbn,
        olid: doc.key ? doc.key.replace('/works/', '') : '',
        coverId: doc.cover_i || null,
        pageCount: doc.number_of_pages_median || '',
        subjects: doc.subject ? doc.subject.slice(0, 10) : []
      };
    });
  } catch(e) {
    return [];
  }
}

function clientSearchAudiobook(query) {
  var url = 'https://librivox.org/api/feed/audiobooks?title=' +
    encodeURIComponent(query) + '&format=json&limit=10&extended=1';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    var data = JSON.parse(resp.getContentText());
    return (data.books || []).map(function(b) {
      var a = (b.authors || [])[0] || {};
      return {
        audiobookId: b.id,
        title: b.title || '',
        author: ((a.first_name || '') + ' ' + (a.last_name || '')).trim() || 'Unknown Author',
        totalTime: b.totaltimesecs ? _formatAudioDuration(parseInt(b.totaltimesecs, 10)) : '',
        numSections: parseInt(b.num_sections || 0, 10),
        coverUrl: b.url_image || ''
      };
    });
  } catch(e) {
    return [];
  }
}

function clientGetAudiobookChapters(audiobookId) {
  var url = 'https://librivox.org/api/feed/audiotracks?project_id=' +
    encodeURIComponent(audiobookId) + '&format=json';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    var data = JSON.parse(resp.getContentText());
    return (data.sections || []).map(function(section, idx) {
      return {
        chapterIndex: idx,
        title: section.title || ('Chapter ' + (idx + 1)),
        duration: section.playtime || '',
        url: section.listen_url || '',
        reader: section.readers && section.readers.length > 0 ? section.readers[0].display_name : ''
      };
    });
  } catch(e) {
    return [];
  }
}

function clientGetArchiveAudioFiles(identifier) {
  var url = 'https://archive.org/metadata/' + encodeURIComponent(identifier);
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return [];
    var data = JSON.parse(resp.getContentText());
    return (data.files || [])
      .filter(function(f) { return /\.(mp3|ogg|flac|opus)$/i.test(f.name || ''); })
      .sort(function(a, b) { return String(a.name || '').localeCompare(String(b.name || '')); })
      .map(function(f, i) {
        return {
          chapterIndex: i,
          title: f.title || f.name || ('Chapter ' + (i + 1)),
          duration: f.length || '',
          url: 'https://archive.org/download/' + identifier + '/' + encodeURIComponent(f.name || '')
        };
      });
  } catch(e) {
    return [];
  }
}

function clientSearchPodcastDiscussions(query) {
  var apiKey = 'Y2R54KZNJ2HMPKTRMKMT';
  var apiSecret = 'JMuRh9Ec#^69cb3wm22tQbkyTPwXqtFfr8dVy9zU';
  var ts = Math.floor(Date.now() / 1000).toString();
  var hashBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_1,
    apiKey + apiSecret + ts
  );
  var hashHex = hashBytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
  var url = 'https://api.podcastindex.org/api/1.0/search/byterm?q=' +
    encodeURIComponent(query + ' book') + '&max=10';
  try {
    var resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'X-Auth-Key': apiKey,
        'X-Auth-Date': ts,
        'Authorization': hashHex,
        'User-Agent': 'PageVault/1.0'
      }
    });
    if (resp.getResponseCode() !== 200) return [];
    var json = JSON.parse(resp.getContentText());
    return (json.feeds || []).slice(0, 10).map(function(f) {
      return {
        id: f.id,
        title: f.title || '',
        author: f.author || '',
        description: f.description ? String(f.description).slice(0, 140) : '',
        coverUrl: f.image || f.artwork || '',
        podcastLink: f.link || f.url || ''
      };
    });
  } catch(e) {
    return [];
  }
}

function clientCheckFreeEbook(isbn) {
  if (!isbn) return { available: false };
  var url = 'https://openlibrary.org/api/books?bibkeys=ISBN:' + encodeURIComponent(isbn) +
    '&jscmd=viewapi&format=json';
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return { available: false };
    var data = JSON.parse(resp.getContentText());
    var key = 'ISBN:' + isbn;
    var entry = data[key];
    if (!entry) return { available: false };
    return {
      available: entry.preview === 'full' || entry.preview === 'noview',
      url: entry.preview_url || '',
      preview: entry.preview || ''
    };
  } catch(e) {
    return { available: false };
  }
}

// Helper: format seconds → "Xh Ym"
function _formatAudioDuration(secs) {
  if (!secs || isNaN(secs)) return '';
  var h = Math.floor(secs / 3600);
  var m = Math.floor((secs % 3600) / 60);
  return h > 0 ? (h + 'h ' + m + 'm') : (m + 'm');
}
