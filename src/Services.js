class SheetRepository {
  constructor() {
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    this.dbSheet = this.ss.getSheetByName('DB') || this.ss.insertSheet('DB');
    this.configSheet = this.ss.getSheetByName('Config');
  }
  loadConfig() { 
    const data = this.configSheet.getDataRange().getValues();
    const res = {};
    for(let i=1;i<data.length;i++) if(data[i][2]) res[String(data[i][2])] = data[i][3];
    return res;
  }
  getKeywords() {
    const last = this.configSheet.getLastRow();
    return last < 2 ? [] : this.configSheet.getRange(2, 1, last-1, 1).getValues().flat().filter(String);
  }
  saveArticles(articles) {
    if (!articles.length) return;
    const rows = articles.map(a => [new Date(), a.keyword, a.title, a.link, a.date, a.source, false, '', '', false, '', '']);
    this.dbSheet.getRange(this.dbSheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
  }
  getUnsentArticles() {
    const last = this.dbSheet.getLastRow();
    if (last < 2) return [];
    const data = this.dbSheet.getRange(2, 1, last - 1, 11).getValues();
    return data.filter(r => !r[6]).map((r, i) => ({ rowIndex: i + 2, keyword: r[1], title: r[2], link: r[3], date: r[4], source: r[5] }));
  }
  markAsSent(articles) { articles.forEach(a => this.dbSheet.getRange(a.rowIndex, 7).setValue(true)); }
}

class UrlCacheService {
  constructor(repo){
    this.cache = CacheService.getScriptCache();
    const last = repo.dbSheet.getLastRow();
    const urls = last < 2 ? [] : repo.dbSheet.getRange(Math.max(2, last-500), 4, Math.min(last-1, 501), 1).getValues().flat();
    this.mem = new Set(urls.map(String));
  }
  exists(u){ return this.mem.has(u) || this.cache.get(this._hash(u)) != null; }
  add(u){ this.cache.put(this._hash(u), '1', 21600); }
  _hash(u){ return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, u).toString(); }
}

class RssService {
  fetch(k,r,l){
    try {
      const u = `https://news.google.com/rss/search?q=${encodeURIComponent(k)}&hl=${l}&gl=${r}&ceid=${r}:${l}`;
      const xml = UrlFetchApp.fetch(u, {muteHttpExceptions:true}).getContentText();
      const items = XmlService.parse(xml).getRootElement().getChild('channel').getChildren('item');
      return items.map(i => ({ title: i.getChild('title').getText(), link: i.getChild('link').getText(), date: Utilities.formatDate(new Date(i.getChild('pubDate').getText()), 'Asia/Tokyo', 'MM/dd'), source: i.getChild('source')?.getText() || 'News' }));
    } catch(e){ return []; }
  }
}

class LoggerService {
  constructor() { this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Logs'); }
  info(c,m){ this.sheet.appendRow([new Date(), 'INFO', c, m]); }
  error(c,m){ this.sheet.appendRow([new Date(), 'ERROR', c, m]); }
}
