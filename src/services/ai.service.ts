import { GoogleGenerativeAI } from "@google/generative-ai";

/**
 * AI Service: Handles interaction with Gemini API
 */
export class AIService {
  private model;

  constructor() {
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY!);
    this.model = genAI.getGenerativeModel({ model: "gemini-pro" });
  }

  /**
   * Summarize article content into 3 concise bullet points.
   */
  async summarize(title: string, url: string): Promise<string> {
    const prompt = `
      Task: Summarize the following article for a busy executive.
      Output: 3 bullet points in Japanese.
      
      Article Title: ${title}
      Article URL: ${url}
    `;

    try {
      const result = await this.model.generateContent(prompt);
      const response = await result.response;
      return response.text();
    } catch (error) {
      console.error("AI Summary Failed:", error);
      throw new Error("Failed to generate summary");
    }
  }
}

export const aiService = new AIService();
