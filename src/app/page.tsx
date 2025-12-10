import { prisma } from "@/lib/prisma";
import { ArticleCard } from "@/components/ArticleCard";
import { AnalyticsDashboard } from "@/components/AnalyticsDashboard";

export default async function Dashboard() {
  // Fetch unread articles
  const articles = await prisma.article.findMany({
    where: { isRead: false },
    orderBy: { publishedAt: "desc" },
    take: 50,
  });

  return (
    <main className="min-h-screen bg-gray-50 p-8">
      <header className="mb-8 flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">News Agent Pro</h1>
          <p className="text-gray-500 text-sm">Powered by Gemini & Next.js</p>
        </div>
        <div className="flex gap-4">
           {/* Analytics Trigger Button could go here */}
        </div>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        
        {/* Main Feed */}
        <div className="lg:col-span-2 space-y-4">
          <h2 className="text-xl font-semibold mb-4">📥 Inbox ({articles.length})</h2>
          {articles.map((article) => (
            <ArticleCard key={article.id} article={article} />
          ))}
          {articles.length === 0 && (
            <div className="text-center py-20 text-gray-400">
              No unread articles. Good job! 🎉
            </div>
          )}
        </div>

        {/* Sidebar / Analytics Preview */}
        <div className="space-y-6">
          <AnalyticsDashboard />
        </div>
      </div>
    </main>
  );
}
