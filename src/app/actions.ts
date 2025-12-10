"use server";

import { prisma } from "@/lib/prisma";
import { aiService } from "@/services/ai.service";
import { revalidatePath } from "next/cache";

/**
 * Updates the rating and optionally sends to Slack (omitted for brevity)
 */
export async function updateRating(id: string, rating: number) {
  await prisma.article.update({
    where: { id },
    data: { rating },
  });
  revalidatePath("/");
}

/**
 * Generates summary using Gemini and saves to DB
 */
export async function generateSummary(id: string) {
  const article = await prisma.article.findUnique({ where: { id } });
  if (!article) throw new Error("Article not found");

  const summary = await aiService.summarize(article.title, article.url);

  await prisma.article.update({
    where: { id },
    data: { summary },
  });
  
  revalidatePath("/");
  return summary;
}

export async function toggleReadStatus(id: string, currentStatus: boolean) {
  await prisma.article.update({
    where: { id },
    data: { isRead: !currentStatus },
  });
  revalidatePath("/");
}

export async function saveComment(id: string, comment: string) {
  await prisma.article.update({
    where: { id },
    data: { comment },
  });
  revalidatePath("/");
}
