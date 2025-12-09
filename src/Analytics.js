/**
 * ==================================================
 * Analytics Module
 * ==================================================
 */

function getAnalyticsData() {
  const repo = new SheetRepository();
  // 全データのRatingとURLを取得
  const data = repo.getAllRatings(); // [[link, rating], ...]

  const keywordCounts = {};
  const dailyCounts = {};
  const domainStats = {}; // { domain: {sum: 0, count: 0} }

  data.forEach(row => {
    // Analyticsデータとしてはシートから取り直したほうが早いが、
    // ここでは構造簡略化のためrepo依存にするか、簡易的に実装。
    // ※本来は Code.gs の getAnalyticsData ロジックを拡張すべき
  });
  
  // Analytics用に、もう一度シート全データを取得する（Keyword等が必要なため）
  const sheet = repo.dbSheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return JSON.stringify({ error: 'No data' });

  const limit = 2000;
  const startRow = Math.max(2, lastRow - limit + 1);
  const values = sheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();

  values.forEach(row => {
    // 0:Time, 1:Kw, 2:Title, 3:Link, 4:Date, 5:Src, 6:Sent, 7:Rating
    const ts = new Date(row[0]);
    const kw = String(row[1]);
    const link = String(row[3]);
    const rating = Number(row[7]);

    // Keywords
    keywordCounts[kw] = (keywordCounts[kw] || 0) + 1;
    // Daily
    const dateKey = (ts.getMonth()+1) + '/' + ts.getDate();
    dailyCounts[dateKey] = (dailyCounts[dateKey] || 0) + 1;

    // Domain Rating
    if (rating > 0) {
      try {
        const domain = new URL(link).hostname;
        if (!domainStats[domain]) domainStats[domain] = { sum: 0, count: 0 };
        domainStats[domain].sum += rating;
        domainStats[domain].count += 1;
      } catch(e) {}
    }
  });

  // 整形
  const keywordRanking = Object.keys(keywordCounts).map(k => ({name: k, count: keywordCounts[k]})).sort((a,b)=>b.count-a.count);
  
  // ドメイン評価ランキング（平均点が高い順）
  const domainRanking = Object.keys(domainStats)
    .map(d => ({
      domain: d,
      avg: (domainStats[d].sum / domainStats[d].count).toFixed(2),
      count: domainStats[d].count
    }))
    .sort((a, b) => b.avg - a.avg)
    .slice(0, 10); // Top 10

  return JSON.stringify({
    total: values.length,
    keywords: keywordRanking,
    daily: dailyCounts,
    domains: domainRanking // 追加
  });
}
