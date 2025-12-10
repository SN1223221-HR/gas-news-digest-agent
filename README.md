# News Agent Pro (Next.js Edition)

[![TypeScript](https://img.shields.io/badge/TypeScript-5.0-blue?style=flat-square&logo=typescript)](https://www.typescriptlang.org/)
[![Next.js](https://img.shields.io/badge/Next.js-14-black?style=flat-square&logo=next.js)](https://nextjs.org/)
[![Supabase](https://img.shields.io/badge/Supabase-PostgreSQL-green?style=flat-square&logo=supabase)](https://supabase.com/)
[![Prisma](https://img.shields.io/badge/Prisma-ORM-2D3748?style=flat-square&logo=prisma)](https://www.prisma.io/)
[![Gemini AI](https://img.shields.io/badge/AI-Gemini_Pro-8E75B2?style=flat-square&logo=google)](https://deepmind.google/technologies/gemini/)

**News Agent Pro** は、プロフェッショナル向けに再設計された、モダンなニュースアグリゲーション＆ナレッジ管理プラットフォームです。

元々は Google Apps Script (GAS) で運用されていたツールを、パフォーマンス、拡張性、開発者体験 (DX) の観点からフルスクラッチで再構築しました。

---

## 🏗 アーキテクチャと技術選定

本プロジェクトでは、エンドツーエンドの型安全性とスケーラビリティを確保するため、**T3 Stack** の思想（Next.js + TypeScript + Prisma + Tailwind）ベースの構成を採用しています。

| カテゴリ | 採用技術 | 選定理由 |
| :--- | :--- | :--- |
| **Framework** | **Next.js 14 (App Router)** | Server Actions によるバックエンドロジックの統合と、React Server Components (RSC) によるパフォーマンス最適化のため。 |
| **Language** | **TypeScript** | 厳格な型システムにより、ランタイムエラーを抑制し、保守性を高めるため。 |
| **Database** | **Supabase (PostgreSQL)** | スプレッドシートでは限界のある大量データの永続化と、複雑なクエリに対応するため。 |
| **ORM** | **Prisma** | DBスキーマをコード管理 (Infrastructure as Code) し、型安全なクエリ発行を行うため。 |
| **AI Engine** | **Google Gemini Pro** | 高速かつ低コストに記事の要約生成を行うため。 |
| **Job Scheduler** | **Vercel Cron** | サーバーレス環境下で、定期的なクローリングジョブを実行するため。 |
| **Styling** | **Tailwind CSS** | ユーティリティファーストなスタイリングにより、UI構築速度とメンテナンス性を向上させるため。 |

### システム構成図
```mermaid
graph TD
    Cron[Vercel Cron] -->|Trigger| API[API Route /api/cron]
    API -->|Invoke| Crawler[Crawler Service]
    Crawler -->|Fetch RSS| GoogleNews[Google News]
    Crawler -->|Upsert Data| DB[(Supabase / PostgreSQL)]
    
    User[User / Client] -->|View Dashboard| Next[Next.js App Router]
    Next -->|Query via Prisma| DB
    
    User -->|Request Summary| Action[Server Action]
    Action -->|Call| AI[Gemini API]
    AI -->|Return Text| Action
    Action -->|Save| DB
```

### 🚀 主な機能

**・自動収集 (Crawler)**: 指定したキーワード・国（リージョン）に基づいてGoogle Newsを定期巡回し、PostgreSQLに保存します。

**・✨ AI要約 (Smart Summary)**: 忙しいビジネスマンのために、Gemini Pro が記事の内容を「3行の箇条書き」で即座に要約します。

・Inbox Zero ワークフロー: 「未読（Inbox）」と「アーカイブ」を明確に分離し、情報収集のノイズを減らします。

・分析ダッシュボード: 記事の評価傾向や、よく読むドメイン（情報ソース）のランキングを可視化します。

・ドメイン駆動設計 (DDD): ビジネスロジックを src/services 層に集約し、UIやAPI層と分離することで高い保守性を実現しています。

### 💡 GAS版からの移行背景
**なぜ Google Apps Script から移行したのか？**

**1.パフォーマンスの限界突破**: スプレッドシートDBは2,000行を超えると急激に動作が重くなりますが、PostgreSQLは数百万件のデータでもミリ秒単位で応答可能です。

**2.開発者体験 (DX) の向上**: VS Code + Git によるバージョン管理、Copilotによるコーディング支援、CI/CDパイプラインの構築が可能になりました。

**3.UI/UXの刷新**: GASの制限されたHTML環境から脱却し、ReactによるリッチでインタラクティブなUIを実現しました。

### 🛠 セットアップ手順

**前提条件**

・Node.js 18以上

・PostgreSQL データベース (Supabase 推奨)

・Google Gemini API Key


**インストール**

1.リポジトリのクローン
```
git clone [https://github.com/yourusername/news-agent-pro.git](https://github.com/yourusername/news-agent-pro.git)
cd news-agent-pro
```
2.依存パッケージのインストール
```
npm install
```
3.環境変数の設定 ルートディレクトリに .env ファイルを作成します。
```
# Database (Supabase Transaction Pooler)
DATABASE_URL="postgres://postgres.[ref]:[password]@aws-0-[region][.pooler.supabase.com:6543/postgres?pgbouncer=true](https://.pooler.supabase.com:6543/postgres?pgbouncer=true)"
# Database (Direct Connection for Migrations)
DIRECT_URL="postgres://postgres.[ref]:[password]@aws-0-[region][.supabase.com:5432/postgres](https://.supabase.com:5432/postgres)"

# AI
GEMINI_API_KEY="your_gemini_api_key"

# Security (for Cron)
CRON_SECRET="your_random_secret_string"
```

4.データベースのマイグレーション PrismaスキーマをDBに反映します。
```
npx prisma db push
```

5.開発サーバーの起動
```
npm run dev
```
ブラウザで http://localhost:3000 を開いてください。

📜 ライセンス
本プロジェクトは MIT License の下で公開されています。詳細は LICENSE ファイルをご確認ください。
