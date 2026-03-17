# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install dependencies
uv sync

# Run the application
uv run streamlit run app.py

# Run a single Python file (for syntax checks etc.)
uv run python -c "import app"
```

No test suite exists. Manual testing via Streamlit UI.

## Environment

- Python 3.13+ (managed by uv)
- `.env` file required with `GEMINI_API_KEY` or `GOOGLE_API_KEY`

## Architecture

照明器具の技術資料画像から項目を自動抽出し、正解データと照合する Streamlit アプリ。

### Processing Pipeline

```
Phase 0: 画像分類
  ファイル名パターン → (不明なら) Gemini API
  6カテゴリ: カタログ, 組立図, 承認図, 器具銘板, 取扱説明書, 外装ラベル

Phase 1: カタログ抽出
  カタログ画像から品番・仕様値を抽出 → 他資料のヒントに使う

Phase 2: 技術資料抽出 (並列3ワーカー)
  カタログ値をヒントとして埋め込んだプロンプトで5資料を読み取り
  各画像: Agentic Vision(値抽出) → BBox座標検出 の2段階

Phase 3: 結果処理
  同一文書種別のマージ → 正解データと比較 → 意味的再判定 → 差分コメント生成 → Excel出力
```

### Module Roles

| Module | Role |
|--------|------|
| `app.py` | Streamlit メイン。パイプライン制御・セッション管理 |
| `classifier.py` | 画像→6カテゴリ分類。ファイル名優先、APIフォールバック |
| `extractor.py` | Gemini Agentic Vision で値抽出 + BBox検出。tenacityでリトライ |
| `prompt.py` | 動的プロンプト生成。カタログ用 / ヒント付き検証用 / フォールバック用 / BBox用 |
| `comparison.py` | 正解比較ロジック・マージ・正規化・Excel出力・LLMコンフリクト解決 |
| `semantic_judge.py` | Gemini APIで意味的一致判定 + 差分コメント生成。mismatch再判定をバッチ処理 |
| `ui_components.py` | BBox描画・マトリクス・正答率等のStreamlit UI描画 |
| `utils.py` | ファイル読み込み(CSV/Excel/PDF)・色割当・APIクライアント生成 |

### Key Design Rules

- **正解データは常にマスターExcel(`正解データ/７点照合正解データ_v2.xlsx`)から取得**。カタログ値は正解ではなくヒント
- **項目リスト = 何を読み取るか（主役）、カタログ値 = 参考ヒント（補助）**
- **品番**: カタログからメイン品番1つだけ取得。子品番・派生品番は不要
- 複数ページの同一文書種別は **1エントリにマージ** → コンフリクト時はカタログヒントで解決
- 比較ロジックは **汎用的**: 数値セット・範囲・接頭辞除去・重複率で判定。項目固有のハードコード禁止

### Comparison Flow (determine_match_status)

```
完全一致 → 数値セット一致 → 装飾除去一致 → 部分文字列 → 数値重複率(≥50%) → NG
```

### Output Files

すべて `output/` 配下に出力:

- `output/evidence/{timestamp}/` — BBox付き画像 + 原本
- `output/読み取り結果/読み取り結果_{timestamp}.xlsx` — 全抽出結果
- `output/正誤判定結果/正誤結果_{品番}_{timestamp}.xlsx` — 横並び比較表（OK緑/NG赤）
- `output/正誤判定結果/正誤結果_一括_{timestamp}.xlsx` — 複数品番統合版

### AB分類（要件3: 計画中）

正誤判定後に、各項目をA/Bに分類して別シートで出力する後処理（別スクリプト）:
- **A**: カタログ（①SPICE）に記載あり → 正解データとのOK/NG判定結果をそのまま使用
- **B**: カタログに記載なし → どの資料に何が書いてあるかを記録 + AIで差分コメント生成
- 判定基準: `catalog_data_merged` にその項目の値があるか否か（ルールベース、LLM不要）

## Language

コード内コメント・UI・ドメイン用語はすべて日本語。ユーザーへの応答も日本語で行うこと。
