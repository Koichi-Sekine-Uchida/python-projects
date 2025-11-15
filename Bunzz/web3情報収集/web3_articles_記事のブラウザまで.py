import feedparser
import schedule
import time
import re
import requests
import webbrowser  # ブラウザを開くためのモジュール

# Web3関連ニュースのRSSフィード
RSS_FEED_URL = "https://www.coindesk.com/arc/outboundfeeds/rss/"

# 関心のあるキーワード
KEYWORDS = ["Web3", "Ethereum", "DeFi", "NFT", "Layer 2", "Smart Contract"]

def fetch_web3_articles():
    print("\n最新のWeb3記事を取得中...")

    # RSSフィードを直接取得して確認
    try:
        response = requests.get(RSS_FEED_URL, timeout=10)
        if response.status_code != 200:
            print(f"⚠️ RSSフィードにアクセスできません (HTTP {response.status_code})")
            return []
        
        # 取得したRSSの最初の500文字をデバッグ表示
        print("\n📡 取得したRSSデータの一部:\n", response.text[:500])

    except requests.RequestException as e:
        print(f"❌ ネットワークエラー: {e}")
        return []

    # `feedparser` で解析
    try:
        feed = feedparser.parse(response.text)
        if not feed.entries:
            print("⚠️ RSSフィードの取得に失敗しました。URLを確認してください。")
            return []
    except Exception as e:
        print(f"❌ フィードの解析中にエラーが発生しました: {e}")
        return []
    
    recommended_articles = []
    
    for entry in feed.entries[:10]:  # 最新10記事を取得
        title = entry.title
        url = entry.link
        
        # キーワードフィルタリング（正規表現を使用）
        if any(re.search(rf"\b{re.escape(keyword)}\b", title, re.IGNORECASE) for keyword in KEYWORDS):
            recommended_articles.append((title, url))
    
    if recommended_articles:
        print("\n📌 おすすめの記事:")
        for idx, (title, url) in enumerate(recommended_articles, 1):
            print(f"{idx}. {title}\n   {url}\n")

        # ここでURLを開く
        for title, url in recommended_articles:
            webbrowser.open(url)  # ブラウザで開く
    else:
        print("🔍 今回はおすすめの記事が見つかりませんでした。\n")

    return recommended_articles

def schedule_fetch(time_str="08:00"):
    """指定した時間にfetch_web3_articlesを実行する"""
    parts = time_str.split(":")
    if len(parts) == 2:
        time_str = f"{int(parts[0]):02}:{parts[1]}"
    schedule.every().day.at(time_str).do(fetch_web3_articles)

if __name__ == "__main__":
    user_time = input("記事取得をスケジュールする時間を入力してください（例: 08:00）: ") or "08:00"
    schedule_fetch(user_time)
    
    # 初回実行時に記事を取得し、ブラウザで開く
    recommended_articles = fetch_web3_articles()

    try:
        while True:
            schedule.run_pending()
            time.sleep(60)  # 1分ごとにチェック
    except KeyboardInterrupt:
        print("\n🛑 プログラムを手動で終了しました。")
