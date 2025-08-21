### Outlook ビジネス文面リライト アドイン（最小構成）

#### 概要
- **機能**: 作成中のメール本文をワンクリックでビジネス向けにリライト。
- **サーバ**: Node.js + Express。OpenAI API を使用（`OPENAI_MODEL` 既定値は `gpt-5`）。
- **実行**: Docker コンテナ。

#### セットアップ
1) `.env` を作成
```
cp .env.sample .env
```
必須: `OPENAI_API_KEY`（OpenAI の API キー）
任意: `PROMPT_FILE`（外部プロンプトのパス。既定: `/app/10_Assets/system_rewrite_ja.txt`）

2) ビルド & 起動
```
docker compose up -d --build
```

3) マニフェストをサイドロード
- `http://localhost:3000/manifest.xml` を Outlook にインポート
  - Windows クライアント: 「アドインを追加 > マイ アドイン > 追加」

4) 使い方
- メール作成画面でリボンの「ビジネスに整える」をクリック

#### 備考
- 開発用の最小構成です。Outlook on the web では HTTPS が推奨/必須のため、必要に応じてリバースプロキシ等で TLS を付与してください。
- マニフェストは `/manifest.xml` をサーバ側で動的生成し、`PUBLIC_BASE_URL` を反映します。
- プロンプトは `10_Assets/system_rewrite_ja.txt` にあります。変更したい場合は同ファイルを編集するか、`PROMPT_FILE` で差し替えてください。


