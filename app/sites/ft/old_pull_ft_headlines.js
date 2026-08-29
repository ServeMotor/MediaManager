function ftHeadlines() {
  function fetchXML() {
    var url = "https://www.ft.com/sitemaps/news.xml";
    var response = UrlFetchApp.fetch(url);
    var xml = response.getContentText();
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    var entries = root.getChildren('url', root.getNamespace());
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear();
    Logger.log("Total entries: " + entries.length);
    for (var i = 0; i < entries.length; i++) {
      var loc = entries[i].getChild('loc', root.getNamespace()).getText();
      var newsNamespace = XmlService.getNamespace("news", "http://www.google.com/schemas/sitemap-news/0.9");
      var newsElement = entries[i].getChild('news', newsNamespace);
      if (newsElement) {
        var title = newsElement.getChild('title', newsNamespace).getText();
        var publicationDate = newsElement.getChild('publication_date', newsNamespace).getText();
        var languageElement = newsElement.getChild('language', newsNamespace);
        var language = languageElement ? languageElement.getText() : '';
        Logger.log("Entry " + i + ": loc=" + loc + ", title=" + title + ", publicationDate=" + publicationDate + ", language=" + language);
        if (language === 'en' || language === '') { // Filter for English or blank language
          sheet.appendRow([loc, title, publicationDate]);
        }
      } else {
        Logger.log("No news element found for entry " + i);
      }
    }
  }
  fetchXML(); // Call the fetchXML function
}