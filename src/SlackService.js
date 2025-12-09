/**
 * ==================================================
 * Slack Integration Module
 * 高評価記事の通知機能を提供
 * ==================================================
 */
class SlackService {
  constructor(webhookUrl) {
    this.url = webhookUrl;
  }

  postMessage(article, note) {
    if (!this.url) return;
    
    const payload = {
      text: `${note} <${article.link}|${article.title}>`,
      blocks: [
        {
          type: "header",
          text: { type: "plain_text", text: "★ High Rating News", emoji: true }
        },
        {
          type: "section",
          text: { 
            type: "mrkdwn", 
            text: `*<${article.link}|${article.title}>*\nSource: ${article.source} | Rating: ${'★'.repeat(article.rating)}` 
          }
        }
      ]
    };
    
    if (article.comment) {
      payload.blocks.push({
        type: "context",
        elements: [{ type: "mrkdwn", text: `📝 Comment: ${article.comment}` }]
      });
    }

    try {
      UrlFetchApp.fetch(this.url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload)
      });
    } catch(e) {
      console.error('Slack Post Error', e);
    }
  }
}
