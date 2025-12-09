/**
 * ==================================================
 * Enterprise News Aggregator for GAS
 * Version: 6.0.0 (Ultimate Edition)
 * ==================================================
 */

const NEWS_AGENT = Object.freeze({ NAME: 'News Agent Pro', VERSION: '6.0.0' });

const CONFIG = Object.freeze({
  SHEET_NAMES: Object.freeze({ CONFIG: 'Config', DB: 'DB', LOGS: 'Logs' }),
  CACHE: Object.freeze({ TTL_SECONDS: 21600, PREFIX: 'news_url_' }),
  RETRY: Object.freeze({ MAX_ATTEMPTS: 3, BASE_DELAY_MS: 1000 }),
  UI: Object.freeze({ TOAST_SECONDS: 3 }),
  LOCK: Object.freeze({ ACQUIRE_TIMEOUT_MS: 30000 }),
  LOG: Object.freeze({ MAX_ROWS: 2000 }),
  DB_COLUMNS: Object.freeze({
    // 1-based index
    TIMESTAMP: 1,
    KEYWORD: 2,
    TITLE: 3,
    LINK: 4,
    DATE: 5,
    SOURCE: 6,
    SENT_FLAG: 7,
    RATING: 8,
    COMMENT: 9,
    IS_READ: 10,
    CAMPAIGN_TAG: 11,
    SUMMARY: 12, // ★AI要約用に追加
    FIRST_DATA_ROW: 2
  })
});

// ==================================================
//  UI Entry Points
// ==================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚡ News Agent')
    .addItem('📰 ニュースリーダー (Sidebar)', 'showNewsSidebar')
    .addItem('📊 分析 (Analytics)', 'showAnalyticsDialog')
    .addSeparator()
    .addItem('📥 今すぐ収集 (Crawl)', 'manualCrawl')
    .addItem('📮 今すぐ配信 (Send Mail)', 'manualSend')
    .addToUi();
}

function showNewsSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('📰 News Feed Pro')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAnalyticsDialog() {
  const html = HtmlService.createTemplateFromFile('AnalyticsDialog')
    .evaluate()
    .setWidth(900)
    .setHeight(700)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 Analytics Dashboard');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================================================
//  Client APIs (Sidebar & AI Hooks)
// ==================================================

/**
 * 記事のステータス更新 (評価/コメント/既読/要約)
 * Slack連携やAI要約保存もここで行う
 */
function updateArticleStatus(link, data) {
  const repo = new SheetRepository();
  repo.updateStatus(link, data); 

  // ★Slack連携: 評価が4以上の時に通知
  if (data.rating && data.rating >= 4) {
    try {
      const conf = repo.loadConfig();
      if (conf.SlackWebhookUrl) {
        // 記事タイトル等の詳細は今回は省略(リンクのみ通知)
        // 厳密にやるならDBから記事情報を引く
        const slack = new SlackService(conf.SlackWebhookUrl);
        slack.postMessage({ 
          title: 'Article Rated High', 
          link: link, 
          source: 'News Agent', 
          rating: data.rating,
          comment: data.comment 
        }, 'High Rating Alert ★' + data.rating);
      }
    } catch(e) {
      console.warn('Slack Post Failed', e);
    }
  }

  return { success: true };
}

/**
 * サイドバー用データ取得
 */
function getNewsFeedData() {
  const repo = new SheetRepository();
  const sheet = repo.dbSheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  // 直近200件取得 (全12カラム)
  const limit = 200;
  const startRow = Math.max(2, lastRow - limit + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 12).getValues();

  const feed = values.map(r => ({
    timestamp: r[0], keyword: r[1], title: r[2], link: r[3], date: r[4], source: r[5],
    rating: r[7], comment: r[8], isRead: r[9], campaignTag: r[10], summary: r[11]
  })).reverse();

  return JSON.stringify(feed);
}

/**
 * [AI] 記事要約を実行
 */
function fetchSummary(link, title) {
  const repo = new SheetRepository();
  const ai = new AIService(repo);
  
  // AIに要約させる
  const summary = ai.summarize(`タイトル: ${title}\nURL: ${link}`);
  
  // 結果を保存
  repo.updateStatus(link, { summary: summary });
  
  return summary;
}

// ==================================================
//  Core Tasks
// ==================================================

function manualCrawl() {
  executeSafely('Manual Crawl', app => {
    const count = app.crawl();
    SpreadsheetApp.getActiveSpreadsheet().toast(`収集完了: ${count}件`, NEWS_AGENT.NAME);
  });
}

function manualSend() {
  executeSafely('Manual Send', app => app.forceSendMail());
}

function executeSafely(context, taskFunction) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CONFIG.LOCK.ACQUIRE_TIMEOUT_MS)) return;
  try {
    const logger = new LoggerService();
    const app = new NewsApp(logger);
    taskFunction(app);
  } catch (e) { console.error(e); } finally { lock.releaseLock(); }
}

// ==================================================
//  News Application Class
// ==================================================

class NewsApp {
  constructor(logger) {
    this.repo = new SheetRepository();
    this.logger = logger || new LoggerService();
    this.rss = new RssService();
    this.mailer = new MailService(this.logger);
    // CampaignManagerが存在しない場合のエラー回避 (別途ファイルが必要)
    try {
      this.campaign = new CampaignManager(this.repo, this.logger);
    } catch (e) {
      this.campaign = null;
      console.warn('CampaignManager not found. Campaign features disabled.');
    }
  }

  crawl() {
    // 1. 通常キーワード
    const settings = this.repo.loadConfig();
    const normalKeywords = this.repo.getKeywords();
    const regions = (settings.Region || 'JP').split(',');
    const lang = settings.Language || 'ja';
    
    let newItems = [];
    const urlCache = new UrlCacheService(this.repo);

    this._fetchKeywords(normalKeywords, regions, lang, urlCache, newItems, '');

    // 2. キャンペーンキーワード
    if (this.campaign) {
      const activeCampaigns = this.campaign.getActiveCampaigns();
      activeCampaigns.forEach(camp => {
        this._fetchKeywords(camp.keywords, regions, lang, urlCache, newItems, camp.name);
      });
    }

    if (newItems.length > 0) {
      this.repo.saveArticles(newItems);
    }
    return newItems.length;
  }

  _fetchKeywords(keywords, regions, lang, cache, resultList, campaignTag) {
    keywords.forEach(kw => {
      regions.forEach(region => {
        try {
          const items = this.rss.fetch(kw.trim(), region.trim(), lang);
          items.forEach(item => {
            if (!cache.exists(item.link)) {
              item.keyword = kw;
              item.campaignTag = campaignTag;
              resultList.push(item);
              cache.add(item.link);
            }
          });
          Utilities.sleep(1000);
        } catch(e) {}
      });
    });
  }

  forceSendMail() {
    const unsent = this.repo.getUnsentArticles();
    if (unsent.length === 0) return 0;

    let campaignSentIndices = new Set();
    
    // 1. キャンペーンメール
    if (this.campaign) {
      campaignSentIndices = this.campaign.processCampaignEmails(unsent);
    }

    // 2. 通常メール
    const normalUnsent = unsent.filter(a => !campaignSentIndices.has(a.rowIndex));
    if (normalUnsent.length > 0) {
      const settings = this.repo.loadConfig();
      this.mailer.sendDailyReport(normalUnsent, settings);
      this.repo.markAsSent(normalUnsent);
    }
    return unsent.length;
  }
}

// ==================================================
//  Repository Class (Fixed & Enhanced)
// ==================================================

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

  // ★エラーの原因だったメソッドを復活
  getAllRatings() {
    const last = this.dbSheet.getLastRow();
    if (last < 2) return [];
    // Link(4) と Rating(8) を含む範囲を取得
    // 取得範囲: 2行目, 4列目(Link) から最後まで、列数は Rating(8) - Link(4) + 1 = 5列分
    const data = this.dbSheet.getRange(2, CONFIG.DB_COLUMNS.LINK, last - 1, 5).getValues();
    
    // [Link, Date, Source, Sent, Rating] の形になるので、LinkとRatingを抽出
    return data.map(row => [row[0], row[4]]); // [Link, Rating]
  }

  saveArticles(articles) {
    if (!articles.length) return;
    const now = new Date();
    // 12カラム分 (Summary含む)
    const rows = articles.map(a => [
      now, a.keyword, a.title, a.link, a.date, a.source, false, 
      '', '', false, a.campaignTag || '', '' 
    ]);
    const start = this.dbSheet.getLastRow() + 1;
    this.dbSheet.getRange(start, 1, rows.length, 12).setValues(rows);
  }

  updateStatus(link, data) {
    // 直近500件から検索
    const last = this.dbSheet.getLastRow();
    const start = Math.max(2, last - 500);
    const links = this.dbSheet.getRange(start, 4, last - start + 1, 1).getValues().flat();
    const idx = links.indexOf(link);
    
    if (idx !== -1) {
      const row = start + idx;
      if (data.rating !== undefined) this.dbSheet.getRange(row, 8).setValue(data.rating);
      if (data.comment !== undefined) this.dbSheet.getRange(row, 9).setValue(data.comment);
      if (data.isRead !== undefined) this.dbSheet.getRange(row, 10).setValue(data.isRead);
      if (data.summary !== undefined) this.dbSheet.getRange(row, 12).setValue(data.summary); // ★AI要約
    }
  }

  getUnsentArticles() {
    const last = this.dbSheet.getLastRow();
    if (last < 2) return [];
    const data = this.dbSheet.getRange(2, 1, last - 1, 11).getValues();
    const rows = [];
    data.forEach((r, i) => {
      if (!r[6]) { 
        rows.push({
          rowIndex: i + 2, timestamp: r[0], keyword: r[1], title: r[2], link: r[3],
          date: r[4], source: r[5], campaignTag: r[10]
        });
      }
    });
    return rows;
  }
  
  markAsSent(articles) {
    articles.forEach(a => this.dbSheet.getRange(a.rowIndex, 7).setValue(true));
  }
}

// ==================================================
//  Helper Services (Logger, Cache, RSS, Mail)
// ==================================================

class LoggerService {
  constructor() { this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Logs') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('Logs'); }
  info(c,m){this.sheet.appendRow([new Date(), 'INFO', c, m]);}
  error(c,m){this.sheet.appendRow([new Date(), 'ERROR', c, m]);}
}
class UrlCacheService {
  constructor(repo){this.repo=repo; this.cache=CacheService.getScriptCache(); try{this.mem=this._getAllUrls();}catch(e){this.mem=new Set();}}
  exists(u){return this.mem.has(u)||this.cache.get(this._hash(u))!=null;}
  add(u){this.mem.add(u); this.cache.put(this._hash(u),'1',21600);}
  _hash(u){return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,u).toString();}
  _getAllUrls(){
    const last=this.repo.dbSheet.getLastRow(); if(last<2)return new Set();
    const v=this.repo.dbSheet.getRange(Math.max(2,last-2000),4,Math.min(last-1,2001),1).getValues();
    return new Set(v.flat().map(String));
  }
}
class RssService {
  fetch(k,r,l){
    try {
      const u=`https://news.google.com/rss/search?q=${encodeURIComponent(k)}&hl=${l}&gl=${r}&ceid=${r}:${l}`;
      const xml=UrlFetchApp.fetch(u,{muteHttpExceptions:true}).getContentText();
      const items=XmlService.parse(xml).getRootElement().getChild('channel').getChildren('item');
      return items.map(i=>({title:i.getChild('title').getText(),link:i.getChild('link').getText(),date:Utilities.formatDate(new Date(i.getChild('pubDate').getText()),'Asia/Tokyo','MM/dd'),source:i.getChild('source')?.getText()||'News'}));
    }catch(e){return[];}
  }
}
class MailService {
  constructor(l){this.logger=l;}
  sendDailyReport(articles, settings){
    const to = settings.MailTo || Session.getActiveUser().getEmail();
    let html = '<h3>Daily News</h3>';
    articles.forEach(a=>html+=`<p><a href="${a.link}">${a.title}</a><br><small>${a.source}</small></p>`);
    GmailApp.sendEmail(to, 'News Briefing', '', {htmlBody:html});
  }
}
