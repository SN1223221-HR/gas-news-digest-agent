package main

import (
	"fmt"
	"log"
	"sync"
	"time"
)

// --- Models ---

type Article struct {
	Timestamp   time.Time
	Keyword     string
	Title       string
	Link        string
	PublishedAt time.Time
	Source      string
	SentFlag    bool
	Summary     string
}

type Config struct {
	Regions  []string
	Language string
	Limits   int
}

// --- Interfaces (Repository / Service) ---
// スプレッドシートやDB、RSS取得のロジックを抽象化します
type Repository interface {
	LoadConfig() Config
	GetKeywords() []string
	SaveArticles(articles []Article) error
	GetUnsentArticles() ([]Article, error)
	MarkAsSent(articles []Article) error
	URLExists(url string) bool
}

type RssService interface {
	Fetch(keyword, region, lang string) ([]Article, error)
}

type MailService interface {
	SendBriefing(groupedArticles map[string][]Article, config Config) error
}

// --- Core Application Logic ---

type NewsApp struct {
	repo   Repository
	rss    RssService
	mailer MailService
	logger *log.Logger
}

func NewNewsApp(repo Repository, rss RssService, mailer MailService) *NewsApp {
	return &NewsApp{
		repo:   repo,
		rss:    rss,
		mailer: mailer,
		logger: log.Default(),
	}
}

// Crawl はキーワードに基づいて記事を収集します
func (app *NewsApp) Crawl() (int, error) {
	config := app.repo.LoadConfig()
	keywords := app.repo.GetKeywords()
	var newItems []Article

	for _, kw := range keywords {
		for _, region := range config.Regions {
			// GASの Utilities.sleep(500) に相当
			time.Sleep(500 * time.Millisecond)

			items, err := app.rss.Fetch(kw, region, config.Language)
			if err != nil {
				app.logger.Printf("Fetch error for %s: %v", kw, err)
				continue
			}

			for _, item := range items {
				if !app.repo.URLExists(item.Link) {
					item.Keyword = kw
					newItems = append(newItems, item)
				}
			}
		}
	}

	if len(newItems) > 0 {
		if err := app.repo.SaveArticles(newItems); err != nil {
			return 0, err
		}
	}

	return len(newItems), nil
}

// ForceSendMail は未送信の記事をグループ化して送信します
func (app *NewsApp) ForceSendMail() (int, error) {
	unsent, err := app.repo.GetUnsentArticles()
	if err != nil || len(unsent) == 0 {
		return 0, err
	}

	config := app.repo.LoadConfig()
	grouped := make(map[string][]Article)

	for _, art := range unsent {
		kw := art.Keyword
		if kw == "" {
			kw = "その他"
		}
		
		if len(grouped[kw]) < config.Limits {
			grouped[kw] = append(grouped[kw], art)
		}
	}

	if err := app.mailer.SendBriefing(grouped, config); err != nil {
		return 0, err
	}

	if err := app.repo.MarkAsSent(unsent); err != nil {
		return 0, err
	}

	return len(unsent), nil
}

// --- Entry Point ---

func main() {
	// ここで各インターフェースの実装（Google Sheets APIを使う実装など）を注入します
	// 実行イメージ:
	// app := NewNewsApp(sheetRepo, rssService, mailService)
	// count, _ := app.Crawl()
	fmt.Println("News Agent Go version initialized.")
}
