# News Agent Pro (v6.0.0 Ultimate Edition)

Google Apps Script (GAS) で動作する、サーバーレスかつ高機能なニュースアグリゲーター＆ナレッジベース構築ツールです。
Google News RSSから記事を収集し、スプレッドシートをデータベースとして活用。専用のサイドバーUIで閲覧、評価、AI要約、分析までをワンストップで行えます。

## 🚀 Key Features

### 1. 収集 & 配信 (Core)
* **マルチリージョン・クロール**: 指定したキーワード×複数国（JP, US, GB等）でGoogle Newsを収集。
* **重複排除**: 直近の収集URLをキャッシュし、重複記事を自動ブロック。
* **メール配信**: 設定時刻に未読記事をHTMLメールで配信。

### 2. 閲覧 & ナレッジ管理 (UI)
* **Sidebar Reader**: スプレッドシート上に専用アプリのようなリーダーを表示。
* **Inbox / Archive**: 「既読」管理により、未読記事のみに集中できるワークフロー。
* **Personalization**: 記事に対し 1〜5 の**星評価**と**コメント**を付与可能。

### 3. インテリジェンス (AI & Integration)
* **✨ AI Summary**: Gemini API を利用し、ワンクリックで記事内容を3行要約。
* **Slack連携**: ★4以上の高評価をつけた記事は、自動的にSlackチャンネルへ通知。

### 4. 分析 & マーケティング (Analytics)
* **Trend Dashboard**: キーワード出現推移、ドメイン別のお気に入りランキングをグラフ化。
* **Campaign Manager**: 期間限定のキーワードと配信先を設定可能（例：「競合調査期間」だけ特定のアドレスに別送）。

---

## 📂 File Structure

プロジェクトは以下のモジュール構成になっています。

| ファイル名 | 役割 |
| --- | --- |
| `Code.gs` | アプリケーションのコアロジック、APIエンドポイント |
| `SheetRepository.gs` | スプレッドシートへのRead/Write責務を分離 |
| `Sidebar.html` | リーダーUIの骨格 |
| `SidebarCSS.html` | UIスタイル定義 |
| `SidebarJS.html` | UIのクライアントサイドロジック |
| `Analytics.gs` | 集計・分析ロジック |
| `AnalyticsDialog.html` | 分析ダッシュボード (Google Charts) |
| `CampaignManager.gs` | 期間限定キャンペーンの管理 |
| `AIService.gs` | Gemini API との通信 |
| `SlackService.gs` | Slack Webhook との通信 |

---

## 🛠 Setup Guide

### 1. スプレッドシートの準備
以下の名前でシートを作成し、1行目にヘッダーを設定してください。

#### Sheet: `DB` (記事データベース)
全12カラムが必要です。
| A | B | C | D | E | F | G | H | I | J | K | L |
|---|---|---|---|---|---|---|---|---|---|---|---|
| Timestamp | Keyword | Title | Link | Date | Source | Sent | Rating | Comment | IsRead | CampaignTag | Summary |

#### Sheet: `Config` (設定)
| A | B | C | D |
|---|---|---|---|
| **SearchKeywords** | (任意) | **SettingKey** | **SettingValue** |
| (キーワード1) | ... | DeliveryHours | 7, 12, 19 |
| (キーワード2) | ... | Region | JP, US |
| ... | ... | Language | ja |
| | | MailTo | your@email.com |
| | | **GeminiApiKey** | (Your API Key) |
| | | **SlackWebhookUrl** | (Your Webhook URL) |

#### Sheet: `Campaigns` (キャンペーン管理)
| A | B | C | D | E |
|---|---|---|---|---|
| CampaignName | Keywords | StartDate | EndDate | TargetEmail |
| 春の競合調査 | KeywordA, KeywordB | 2025/04/01 | 2025/04/30 | team@example.com |

#### Sheet: `Logs`
(自動生成されます)

### 2. APIキーの設定
* **Gemini API Key**: [Google AI Studio](https://aistudio.google.com/) から取得し、Configシートの `GeminiApiKey` に入力。
* **Slack Webhook**: Slackアプリ設定から Incoming Webhook URL を発行し、Configシートの `SlackWebhookUrl` に入力。

### 3. トリガーの設定
GASエディタの「トリガー」メニューから以下を設定してください。

* `crawlTask`: 1〜4時間ごとの定期実行（収集）
* `checkAndSendMailTask`: 1時間ごとの定期実行（メール配信チェック）

---

## 📖 Usage

1.  **スプレッドシートを開く**:
    * メニューバーに `⚡ News Agent` が表示されます。
2.  **ニュースを読む**:
    * `📰 ニュースリーダー (Sidebar)` をクリック。
    * サイドバーで記事をチェック、★評価、コメント入力、AI要約を実行。
    * 読み終わったら `Mark as Read` でアーカイブ。
3.  **分析する**:
    * `📊 分析 (Analytics)` をクリック。
    * トレンドや、自分がよく高評価をつけるドメイン（お気に入りソース）を確認。

---

## ⚠️ Notes & Limits

* **Google News RSS仕様**: 取得できる記事は直近のものに限られます。
* **Gemini API**: 無料枠の範囲内であれば課金は発生しませんが、レートリミットに注意してください。
* **シート行数制限**: `DB` シートが数千行を超えると動作が重くなる可能性があります。適宜 `Logs` シートの掃除や古いデータのアーカイブを行ってください。

---

## License
MIT License
