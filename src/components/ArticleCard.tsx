"use client";

import { useState, useTransition } from "react";
import { Article } from "@prisma/client";
import { updateRating, generateSummary, toggleReadStatus, saveComment } from "@/app/actions";
import { StarIcon, ArchiveBoxIcon, SparklesIcon } from "@heroicons/react/24/outline";
import { StarIcon as StarIconSolid } from "@heroicons/react/24/solid";

export function ArticleCard({ article }: { article: Article }) {
  const [isPending, startTransition] = useTransition();
  const [summary, setSummary] = useState(article.summary);

  const handleRate = (rating: number) => {
    startTransition(() => updateRating(article.id, rating));
  };

  const handleSummary = async () => {
    const text = await generateSummary(article.id);
    setSummary(text);
  };

  return (
    <div className="bg-white p-5 rounded-xl shadow-sm border border-gray-100 hover:shadow-md transition-shadow">
      <div className="flex justify-between items-start mb-2">
        <span className="text-xs font-bold text-blue-600 bg-blue-50 px-2 py-1 rounded">
          {article.keyword}
        </span>
        <span className="text-xs text-gray-400">
          {new Date(article.publishedAt).toLocaleDateString()}
        </span>
      </div>

      <h3 className="text-lg font-bold text-gray-900 leading-tight mb-2">
        <a href={article.url} target="_blank" rel="noopener noreferrer" className="hover:text-blue-600 hover:underline">
          {article.title}
        </a>
      </h3>

      <div className="text-xs text-gray-500 mb-4">{article.source}</div>

      {/* AI Summary Section */}
      {summary ? (
        <div className="bg-gray-50 p-3 rounded-lg text-sm text-gray-700 mb-4 border border-gray-100">
          <div className="flex items-center gap-1 text-purple-600 font-bold mb-1 text-xs">
            <SparklesIcon className="w-3 h-3" /> AI Summary
          </div>
          <div className="whitespace-pre-wrap">{summary}</div>
        </div>
      ) : (
        <button 
          onClick={handleSummary}
          disabled={isPending}
          className="text-xs flex items-center gap-1 text-purple-600 hover:bg-purple-50 px-2 py-1 rounded mb-4 transition-colors"
        >
          <SparklesIcon className="w-3 h-3" /> Generate Summary
        </button>
      )}

      {/* Action Bar */}
      <div className="flex items-center justify-between pt-4 border-t border-gray-50">
        
        {/* Rating */}
        <div className="flex gap-1">
          {[1, 2, 3, 4, 5].map((star) => (
            <button key={star} onClick={() => handleRate(star)}>
              {star <= article.rating ? (
                <StarIconSolid className="w-5 h-5 text-yellow-400" />
              ) : (
                <StarIcon className="w-5 h-5 text-gray-300 hover:text-yellow-400" />
              )}
            </button>
          ))}
        </div>

        {/* Archive Button */}
        <button
          onClick={() => startTransition(() => toggleReadStatus(article.id, article.isRead))}
          className="flex items-center gap-1 text-sm text-gray-500 hover:text-green-600 px-3 py-1 rounded-full hover:bg-green-50 transition-colors"
        >
          <ArchiveBoxIcon className="w-4 h-4" />
          {article.isRead ? "Unarchive" : "Mark as Read"}
        </button>
      </div>
    </div>
  );
}
