/**
 * WebApp.gs
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('HR Strategic Intelligence')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  const repo = new SheetRepository();
  const keywords = repo.getKeywords();
  const sheet = repo.dbSheet;
  const lastRow = sheet.getLastRow();

  // データがない場合の早期リターン
  if (lastRow < 2) return { keywords, articles: [] };

  // 過去200件程度を取得
  const startRow = Math.max(2, lastRow - 199);
  const numRows = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, numRows, 11).getValues();

  // オブジェクト化して新しい順にソート
  const articles = values.map(r => ({
    timestamp: r[0],
    keyword: r[1],
    title: r[2],
    link: r[3],
    date: r[4],
    source: r[5],
    isRead: r[9]
  })).reverse();

  return {
    keywords: keywords,
    articles: articles
  };
}
