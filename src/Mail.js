/**
 * Mail.gs
 * 配信メールの構築と送信を担う
 */
class MailService {
  constructor() {
    this.repo = new SheetRepository();
  }

  /**
   * ニュースダイジェストメールの送信
   */
  sendBriefing(grouped, settings) {
    const to = settings.MailTo || Session.getActiveUser().getEmail();
    const webAppUrl = ScriptApp.getService().getUrl();
    const kws = Object.keys(grouped);
    
    // 文字化け（豆腐）対策：特殊文字を避け、標準的なフォントスタックを指定
    let html = `
      <div style="font-family: 'Helvetica Neue', Arial, 'Hiragino Kaku Gothic ProN', 'BIZ UDPGothic', sans-serif; max-width: 600px; margin: auto; border: 1px solid #eee; border-radius: 10px; padding: 25px; color: #333;">
        <h2 style="color: #1a73e8; margin-bottom: 10px;">Daily News Digest</h2>
        <p style="font-size: 14px; color: #666; line-height: 1.5;">
          最新のトピックが更新されました。詳細はダッシュボードから確認できます。
        </p>
        
        <div style="margin: 30px 0; text-align: center;">
          <a href="${webAppUrl}" style="background-color: #1a73e8; color: white; padding: 14px 28px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px;">Webダッシュボードを開く</a>
        </div>

        <div style="border-top: 2px solid #f0f0f0; padding-top: 20px;">
          <p style="font-weight: bold; font-size: 12px; color: #aaa; text-transform: uppercase; letter-spacing: 1px;">Todays Highlights</p>
    `;

    kws.forEach(kw => {
      html += `
        <div style="margin-top: 20px;">
          <span style="background: #e8f0fe; color: #1a73e8; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: bold;"># ${kw}</span>
          <ul style="padding-left: 0; list-style: none;">
      `;
      grouped[kw].forEach(a => {
        html += `
          <li style="margin: 10px 0;">
            <a href="${a.link}" style="text-decoration: none; color: #333; font-size: 14px; font-weight: 500;">${a.title}</a>
            <div style="font-size: 11px; color: #999; margin-top: 2px;">${a.source}</div>
          </li>`;
      });
      html += `</ul></div>`;
    });

    html += `</div></div>`;

    const subject = `【News】${kws.slice(0, 3).join(', ')}...`;
    
    GmailApp.sendEmail(to, subject, "HTMLメールを表示できる環境でご覧ください。", {
      htmlBody: html,
      name: "News Curator"
    });
  }
}
