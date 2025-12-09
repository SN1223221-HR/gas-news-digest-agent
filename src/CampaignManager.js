/**
 * ==================================================
 * Campaign Manager Module
 * 期間限定の収集・配信ロジックを担当
 * ==================================================
 */

class CampaignManager {
  constructor(repo, logger) {
    this.repo = repo;
    this.logger = logger;
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    // シートがない場合は作成（エラー防止）
    this.sheet = this.ss.getSheetByName('Campaigns') || this.ss.insertSheet('Campaigns');
  }

  /**
   * 現在有効なキャンペーン設定を取得
   */
  getActiveCampaigns() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow < 2) return [];

    // Columns: Name, Keywords, Start, End, Email
    const data = this.sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const active = [];

    data.forEach(row => {
      const [name, kws, start, end, email] = row;
      if (!name || !kws) return;

      const startDate = new Date(start);
      const endDate = new Date(end);

      // 日付範囲チェック
      if (today >= startDate && today <= endDate) {
        active.push({
          name: String(name),
          keywords: String(kws).split(',').map(k => k.trim()),
          email: String(email)
        });
      }
    });

    return active;
  }

  /**
   * キャンペーンタグ付きの記事を抽出してメール送信
   * (通常配信から除外するために送信済みIDリストを返す)
   * @param {Array} unsentArticles 
   * @returns {Set} 送信済みとなった記事の行インデックス
   */
  processCampaignEmails(unsentArticles) {
    const activeCampaigns = this.getActiveCampaigns();
    const sentRowIndices = new Set();
    if (activeCampaigns.length === 0 || unsentArticles.length === 0) return sentRowIndices;

    const sentArticles = [];

    activeCampaigns.forEach(campaign => {
      // このキャンペーンに該当する記事をフィルタリング
      const targets = unsentArticles.filter(a => a.campaignTag === campaign.name);
      
      if (targets.length > 0) {
        try {
          // キャンペーン専用メール送信
          this._sendMail(targets, campaign);
          
          targets.forEach(a => {
            sentRowIndices.add(a.rowIndex);
            sentArticles.push(a);
          });
          
          this.logger.info('Campaign', `Sent ${targets.length} articles for "${campaign.name}"`);
        } catch (e) {
          this.logger.error('CampaignMail', e.message);
        }
      }
    });

    // DBのステータス更新
    if (sentArticles.length > 0) {
      this.repo.markAsSent(sentArticles);
    }

    return sentRowIndices;
  }

  _sendMail(articles, campaign) {
    const recipient = campaign.email || Session.getActiveUser().getEmail();
    const subject = `【特設】${campaign.name} ニュースレポート (${articles.length}件)`;
    
    let body = `<h3>${campaign.name} キャンペーン収集結果</h3>`;
    body += `<p>期間: ${new Date().toLocaleDateString()} の収集分</p><hr>`;
    
    articles.forEach(a => {
      body += `<p><b>${a.title}</b><br>`;
      body += `<a href="${a.link}">${a.source}</a> - ${a.date}</p>`;
    });

    GmailApp.sendEmail(recipient, subject, '', { htmlBody: body });
  }
}
