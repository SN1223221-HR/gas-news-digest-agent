import { NextResponse } from "next/server";
import { crawlerService } from "@/services/crawler.service";

export const dynamic = 'force-dynamic'; // Disable caching

export async function GET(request: Request) {
  // Simple authentication for Cron
  const authHeader = request.headers.get('authorization');
  if (authHeader !== `Bearer ${process.env.CRON_SECRET}`) {
    return new NextResponse('Unauthorized', { status: 401 });
  }

  // Load keywords from config or DB
  const keywords = ["Generative AI", "TypeScript", "Startup"]; 
  
  const result = await crawlerService.execute(keywords);

  return NextResponse.json({
    success: true,
    message: `Crawl complete. Added: ${result.added}, Errors: ${result.errors}`,
    timestamp: new Date().toISOString()
  });
}
