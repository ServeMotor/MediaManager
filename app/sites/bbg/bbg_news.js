function fetchBloombergSitemap() {
  var url = "https://www.bloomberg.com/sitemaps/news/latest.xml";

  var options = {
    'muteHttpExceptions': true, // Prevents script failure if direct request is blocked
    'headers': {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept': 'text/xml,application/xml,application/xhtml+xml,text/html'
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    // If direct request is blocked (HTTP status non-200), fetch via proxy
    if (responseCode !== 200) {
      Logger.log("Direct request returned HTTP " + responseCode + ". Fetching via proxy...");
      return fetchSitemapViaProxy(url);
    }

    parseAndWriteSitemap(response.getContentText());

  } catch (e) {
    Logger.log("Script error: " + e.toString());
    SpreadsheetApp.getUi().alert("Error: " + e.toString());
  }
}

// XML parser specifically for Google News Sitemap structure
function parseAndWriteSitemap(xmlText) {
  var document = XmlService.parse(xmlText);
  var root = document.getRootElement(); // <urlset>

  var sitemapNamespace = root.getNamespace();
  var newsNamespace = XmlService.getNamespace("news", "http://www.google.com/schemas/sitemap-news/0.9");

  var entries = root.getChildren('url', sitemapNamespace);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Reset current sheet layout
  sheet.clear();
  sheet.appendRow(["URL", "Title", "Publication Date", "Language", "Publication Name"]);
  sheet.getRange("1:1").setFontWeight("bold");

  var rows = [];

  for (var i = 0; i < entries.length; i++) {
    var locElement = entries[i].getChild('loc', sitemapNamespace);
    var loc = locElement ? locElement.getText() : '';

    var newsElement = entries[i].getChild('news', newsNamespace);
    if (newsElement) {
      var titleElem = newsElement.getChild('title', newsNamespace);
      var pubDateElem = newsElement.getChild('publication_date', newsNamespace);
      var publicationElem = newsElement.getChild('publication', newsNamespace);

      var title = titleElem ? titleElem.getText() : '';
      var pubDate = pubDateElem ? pubDateElem.getText() : '';
      
      var language = '';
      var pubName = '';

      if (publicationElem) {
        var langElem = publicationElem.getChild('language', newsNamespace);
        var nameElem = publicationElem.getChild('name', newsNamespace);
        language = langElem ? langElem.getText() : '';
        pubName = nameElem ? nameElem.getText() : '';
      }

      rows.push([loc, title, pubDate, language, pubName]);
    }
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 5).setValues(rows);
    Logger.log("Successfully imported " + rows.length + " sitemap entries.");
  } else {
    Logger.log("No news entries found in sitemap.");
  }
}

// Fallback proxy fetcher for XML files when direct request fails
function fetchSitemapViaProxy(targetUrl) {
  var proxyUrl = "https://api.allorigins.win/raw?url=" + encodeURIComponent(targetUrl);
  var response = UrlFetchApp.fetch(proxyUrl, { 'muteHttpExceptions': true });

  if (response.getResponseCode() === 200) {
    parseAndWriteSitemap(response.getContentText());
  } else {
    SpreadsheetApp.getUi().alert("Failed to retrieve sitemap directly or through proxy.");
  }
}
