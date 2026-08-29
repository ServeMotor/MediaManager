function ftHeadlines() {
  var url = "https://www.ft.com/sitemaps/news.xml";
  
  try {
    var response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    
    if (response.getResponseCode() !== 200) {
      Logger.log("Failed to fetch XML. HTTP Status: " + response.getResponseCode());
      return;
    }

    var xml = response.getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    
    // Define namespaces once outside the loop
    var sitemapNamespace = root.getNamespace();
    var newsNamespace = XmlService.getNamespace("news", "http://www.google.com/schemas/sitemap-news/0.9");
    
    var entries = root.getChildren('url', sitemapNamespace);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Initialize output array with headers
    var outputData = [["URL", "Title", "Publication Date"]];

    for (var i = 0; i < entries.length; i++) {
      var locElem = entries[i].getChild('loc', sitemapNamespace);
      var loc = locElem ? locElem.getText() : '';

      var newsElement = entries[i].getChild('news', newsNamespace);
      if (newsElement) {
        var titleElem = newsElement.getChild('title', newsNamespace);
        var pubDateElem = newsElement.getChild('publication_date', newsNamespace);
        var langElem = newsElement.getChild('language', newsNamespace);

        var title = titleElem ? titleElem.getText() : '';
        var pubDate = pubDateElem ? pubDateElem.getText() : '';
        var language = langElem ? langElem.getText() : '';

        // Filter for English or blank language tags
        if (language === 'en' || language === '') {
          outputData.push([loc, title, pubDate]);
        }
      }
    }

    // Batch update the sheet in a single call
    sheet.clearContents();
    if (outputData.length > 0) {
      sheet.getRange(1, 1, outputData.length, 3).setValues(outputData);
      sheet.getRange("1:1").setFontWeight("bold");
      Logger.log("Successfully imported " + (outputData.length - 1) + " FT headlines.");
    }

  } catch (e) {
    Logger.log("Error parsing FT sitemap: " + e.toString());
  }
}