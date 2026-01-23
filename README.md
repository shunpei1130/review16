# 診断データ解析ツール（静的サイト）

Excel（.xlsx）を**ブラウザ内（ローカル）**で読み込み、以下を行う静的Webツールです。

- シートのテーブル探索（検索 / ソート / ページング）
- 列プロファイル（型推定、欠損、ユニーク数、上位値、min/median/max）
- 可視化（ヒストグラム、カテゴリ件数、散布図、日時×数値、相関、Sankey）
- CSV / JSON エクスポート（PIIマスク反映）

> **重要**: このツール自体はサーバ不要の静的サイトです。  
> ただし、UI/解析ライブラリをCDNから読み込むため、初回はネット接続が必要です（CDNをローカル化すればオフラインでも動きます）。

---

## 使い方

### 1) 起動
- `index.html` をダブルクリックで開く（またはブラウザで開く）

### 2) Excelを読み込む
- 画面上部の「Excelファイル」から `.xlsx` を選択
- 読み込みが完了すると、各タブが利用できます

---

## タブ説明

### ダッシュボード
- シート数、総行数などのKPI
- シート一覧（クリックで探索へ）
- 代表シート（diagnosis / referral系）がある場合、自動インサイトと簡易グラフを表示

### 紹介深掘り
- `referral_events` をベースに、期間フィルタ付きで **share / visit / complete** を可視化
- shareのプラットフォーム内訳（`line` / `copy` / `twitter`）
- referrer（紹介者）Leaderboard：shares / visitors / completes / visit→complete率 / time-to-complete
- referrer詳細：招待ユーザー（edge）一覧、日次推移、time-to-complete分布
- `sankey_visits(_acyclic)` / `sankey_completes(_acyclic)` があれば、ネットワークのSankey表示

### シート探索
- DataTablesでテーブル表示（検索/並び替え/ページング）
- 上部のチェックで **JSON列 / Unnamed列** の表示切替

### プロファイル
- 列ごとに型推定、欠損数、ユニーク数、上位値、min/median/max を表示

### 可視化
- シートと列を選んでPlotlyで描画
- `相関ヒートマップ` は数値列を最大12列で表示
- `Sankey` は `source/target/value` 列があるシートで動作

### エクスポート
- 選択中のシートを CSV / JSON で保存
- **PIIマスクONの場合、出力もマスクされます**

---

## セキュリティ/プライバシー
- Excelは `FileReader` で読み込み、解析はブラウザ内で完結します（このページが自動送信することはありません）
- ただし、CDNからライブラリを読み込みます（ネットワーク接続が発生します）

---

## カスタムしたい場合
- `app.js` の `shouldHideColumn` / `applyMaskForDisplay` を調整すると、
  非表示列やマスク対象を増やせます。
