# 7点照合チェッカー

照明器具の技術資料画像（承認図・器具銘板・取扱説明書など）から指定項目を自動抽出し、7点照合を行う Streamlit アプリ。

## セットアップ

### 前提条件

- Python 3.13+
- [uv](https://docs.astral.sh/uv/)
- Gemini API キー

### インストール

```bash
uv sync
```

### 環境変数

`.env` ファイルをプロジェクトルートに作成し、Gemini API キーを設定する。

```
GOOGLE_API_KEY=your_api_key_here
```

## 起動

```bash
uv run streamlit run app.py
```

## 使い方

1. **項目リストをアップロード** — CSV または Excel（`項目確認ドラフト_TEST.xlsx` 等）をサイドバーからアップロード
2. **資料画像をアップロード** — PNG / JPEG / PDF（複数可）をアップロード
3. **「抽出実行」をクリック** — Gemini API で画像分類 → 値抽出 → BBox 検出を実行
4. **結果を確認** — 画像別タブで BBox 付き画像と抽出値、照合マトリクスを確認

### 正誤判定（自動）

プロジェクトルートに `７点照合正解データ_TEST.xlsx` が配置されていれば、抽出完了後に自動で正誤判定が実行される。

- 抽出結果から品番を自動検出し、正解データとマッチ
- 画像分類結果（承認図・器具銘板等）と正解データの技術文書列で照合
- 判定結果: OK / NG / 部分一致 / 対象外
- 横並び比較表 Excel（`正誤結果_{品番}_{日時}.xlsx`）がディレクトリに自動生成される
- UI 上のダウンロードボタンからも取得可能

## ファイル構成

```
app.py              - Streamlit メインアプリ
classifier.py       - 画像分類（5カテゴリ自動判定）
extractor.py        - Gemini Agentic Vision による値抽出 + BBox 取得
comparison.py       - 正解データ比較・正答率計算・Excel 出力
prompt.py           - プロンプトテンプレート
ui_components.py    - UI 描画（BBox 描画・マトリクス・正答率サマリー等）
utils.py            - ユーティリティ（ファイル読み込み・色割り当て・API クライアント）
```

## 技術スタック

- **Streamlit** — Web UI
- **Gemini API** (`google-genai`) — 画像分類・値抽出・BBox 検出
- **openpyxl** — Excel 読み書き
- **Pillow** — 画像処理・BBox 描画
- **PyMuPDF** — PDF → 画像変換
