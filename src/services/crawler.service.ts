import Parser from "rss-parser";
import { prisma } from "@/lib/prisma";
import { aiService } from "./ai.service";

const parser = new Parser();

interface CrawlResult {
  added: number;
  errors: number;
}

/**
 * Crawler Service: Fetches RSS feeds and persists to Database
 */
export class CrawlerService {
  
  /**
   * Main entry point for the crawling job
   */
  async execute(keywords: string[], regions = ["JP"]): Promise<CrawlResult> {
    let addedCount = 0;
    let errorCount = 0;

    for (const keyword of keywords) {
      for (const region of regions) {
        try {
          const feedUrl = this.buildGoogleNewsUrl(keyword, region);
          const feed = await parser.parseURL(feedUrl);

          // Process items in parallel
          await Promise.all(
            feed.items.map(async (item) => {
              if (!item.link || !item.title) return;

              // Idempotency check happens at DB level via @unique constraint on URL
              try {
                await prisma.article.upsert({
                  where: { url: item.link },
                  update: {}, // Do nothing if exists
                  create: {
                    title: item.title,
                    url: item.link,
                    source: item.source?.title || "News",
                    publishedAt: item.isoDate ? new Date(item.isoDate) : new Date(),
                    keyword: keyword,
                  },
                });
                addedCount++;
              } catch (e) {
                // Ignore unique constraint violations (duplicates)
              }
            })
          );
        } catch (e) {
          console.error(`Failed to crawl ${keyword} in ${region}:`, e);
          errorCount++;
        }
      }
    }

    return { added: addedCount, errors: errorCount };
  }

  private buildGoogleNewsUrl(keyword: string, region: string): string {
    const lang = region === "JP" ? "ja" : "en";
    return `https://news.google.com/rss/search?q=${encodeURIComponent(
      keyword
    )}&hl=${lang}&gl=${region}&ceid=${region}:${lang}`;
  }
}

export const crawlerService = new CrawlerService();
