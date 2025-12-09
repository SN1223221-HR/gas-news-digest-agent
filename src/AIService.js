/**
 * ==================================================
 * AI Service Module
 * 記事の要約機能を提供
 * ==================================================
 */
class AIService {
  constructor(repo) {
    this.repo = repo;
    const conf = this.repo.loadConfig();
    this.apiKey = conf.GeminiApiKey || ''; 
  }

  /**
   * 記事を要約する
   */
  summarize(textOrUrl) {
    if (!this.apiKey) return '⚠ Configシートに "GeminiApiKey" を設定してください。';

    const prompt = `以下の記事タイトルとURLから内容を推測し、ビジネスマン向けに3行の箇条書きで要約してください。\n\n${textOrUrl}`;

    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${this.apiKey}`;
      const payload = {
        contents: [{ parts: [{ text: prompt }] }]
      };

      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const json = JSON.parse(response.getContentText());
      
      if (json.error) return `AI Error: ${json.error.message}`;
      if (!json.candidates || json.candidates.length === 0) return 'AIからの応答がありませんでした。';
      
      return json.candidates[0].content.parts[0].text;
    } catch (e) {
      return 'Summary failed: ' + e.message;
    }
  }
}
