"""意味的一致判定モジュール

プログラム比較でOKにならなかった項目について、
Gemini APIで「意味的に同じか？」を判断する。
"""

import json
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed

from google import genai
from tenacity import retry, stop_after_attempt, wait_exponential

logger = logging.getLogger(__name__)

MODEL = "gemini-3-flash-preview"

# 1バッチあたりの最大件数
_BATCH_SIZE = 20

# 意味比較プロンプトテンプレート
_SEMANTIC_PROMPT_TEMPLATE = """\
あなたは照明器具の技術仕様値を比較する専門家です。

以下の各項目について、「正解値」と「抽出値」が意味的に同じ情報を表しているかを判定してください。

## 判定基準
- OK: 表記の違い・省略・語順・接頭辞の有無・単位表記の揺れであり、同じ情報を指している場合
- NG: 数値が明らかに異なる、または情報の内容自体が異なる場合

### OKの例
- 「AC100V」↔「100V」
- 「電源ユニット別売」↔「別売」
- 「LED(白色)」↔「LED」（詳細の有無だが同じもの）
- 「ダークグレーメタリック」↔「ダークグレーメタリック仕上」
- 「直下近接限度30cm、埋込高さ185mm以上」↔「被照射物近接限度距離30cm以上、施工時埋込高さ185mm必要、断熱施工不可...」（要点が同じ）
- 「100~242V」↔「AC100V, AC200V, AC242V」
- 「0.5kg」↔「500g」

### NGの例
- 「100V」vs「200V」（異なる値）
- 「3000K」vs「4000K」（異なる色温度）
- 「5200lm」vs「6900lm」（異なる光束）
- 「NEL4500UNLA9」vs「NEL4600UNLA9」（異なる品番）

## 判定対象
{table}

JSON配列で出力してください:
[{{"index": 1, "ok": true}}, {{"index": 2, "ok": false}}]
"""


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=2, min=4, max=30),
    reraise=True,
)
def _call_gemini(client: genai.Client, prompt: str) -> str:
    """Gemini APIを呼び出してテキストレスポンスを返す。"""
    response = client.models.generate_content(
        model=MODEL,
        contents=[prompt],
    )
    text = ""
    if response.candidates:
        for part in response.candidates[0].content.parts:
            if hasattr(part, "text") and part.text:
                text += part.text
    return text


def _extract_json_from_text(text: str) -> list[dict] | None:
    """テキストからJSON配列を抽出する。"""
    cleaned = text.strip()

    # コードブロック内のJSONを抽出
    if "```json" in cleaned:
        cleaned = cleaned.split("```json", 1)[1].split("```", 1)[0].strip()
    elif "```" in cleaned:
        cleaned = cleaned.split("```", 1)[1].split("```", 1)[0].strip()

    if cleaned.startswith("[") or cleaned.startswith("{"):
        try:
            parsed = json.loads(cleaned)
            if isinstance(parsed, dict):
                return [parsed]
            return parsed
        except json.JSONDecodeError:
            pass
    return None


def _build_table(items: list[dict]) -> str:
    """判定対象テーブルをMarkdown形式で構築する。

    Args:
        items: [{"index": int, "field": str, "correct": str, "extracted": str}]
    """
    lines = ["| No. | 項目名 | 正解値 | 抽出値 |", "|-----|--------|--------|--------|"]
    for item in items:
        lines.append(
            f"| {item['index']} | {item['field']} | {item['correct']} | {item['extracted']} |"
        )
    return "\n".join(lines)


def _process_batch(
    client: genai.Client,
    batch: list[dict],
) -> dict[int, bool]:
    """1バッチ分の意味比較をGeminiで実行する。

    Args:
        client: Gemini APIクライアント
        batch: [{"index": int, "field": str, "correct": str, "extracted": str}]

    Returns:
        {index: ok_or_not} の辞書
    """
    table = _build_table(batch)
    prompt = _SEMANTIC_PROMPT_TEMPLATE.format(table=table)

    try:
        response_text = _call_gemini(client, prompt)
        results = _extract_json_from_text(response_text)
        if not results:
            logger.warning("意味比較: JSONの抽出に失敗しました。レスポンス: %s", response_text[:200])
            return {}

        judgments: dict[int, bool] = {}
        for item in results:
            idx = item.get("index")
            ok = item.get("ok")
            if idx is not None and ok is not None:
                judgments[int(idx)] = bool(ok)
        return judgments

    except Exception:
        logger.exception("意味比較: Gemini API呼び出しに失敗しました")
        return {}


def refine_with_semantic_judge(
    client: genai.Client,
    comparison_rows: list[dict],
    field_names: list[str],
) -> list[dict]:
    """プログラム比較で不一致・部分一致だった項目をGeminiで意味的に再判定する。

    Args:
        client: google.genai.Client インスタンス
        comparison_rows: compare_by_doc_type() の返り値。
            各要素は {"品番": str, "技術文書": str, "items": {項目名: {"extracted": str, "correct": str, "status": str}}}
        field_names: 項目名リスト

    Returns:
        同じ構造の comparison_rows（statusが更新済み）
    """
    # 1. 対象項目を収集
    targets: list[dict] = []
    # 各ターゲットの位置情報（後で更新するため）
    target_locations: list[tuple[int, str]] = []  # (row_index, field_name)

    for row_idx, row_data in enumerate(comparison_rows):
        items = row_data.get("items", {})
        for field_name in field_names:
            item_info = items.get(field_name, {})
            status = item_info.get("status", "")
            if status in ("partial", "mismatch"):
                extracted = item_info.get("extracted", "")
                correct = item_info.get("correct", "")
                # 抽出値・正解値の両方が空でない場合のみ対象
                if extracted and correct and extracted != "-" and correct != "-":
                    targets.append({
                        "index": len(targets) + 1,
                        "field": field_name,
                        "correct": correct,
                        "extracted": extracted,
                    })
                    target_locations.append((row_idx, field_name))

    # 2. 対象がなければそのまま返す
    if not targets:
        logger.info("意味比較: 対象項目なし（全てOKまたはN/A）")
        return comparison_rows

    logger.info("意味比較: %d 件の項目をGeminiで再判定します", len(targets))

    # 3. バッチに分割
    batches: list[list[dict]] = []
    for i in range(0, len(targets), _BATCH_SIZE):
        batches.append(targets[i : i + _BATCH_SIZE])

    # 4. 並列でバッチ処理
    all_judgments: dict[int, bool] = {}
    with ThreadPoolExecutor(max_workers=3) as executor:
        future_to_batch = {
            executor.submit(_process_batch, client, batch): batch
            for batch in batches
        }
        for future in as_completed(future_to_batch):
            batch_judgments = future.result()
            all_judgments.update(batch_judgments)

    # 5. 結果を反映
    updated_count = 0
    for target, (row_idx, field_name) in zip(targets, target_locations):
        idx = target["index"]
        if idx in all_judgments and all_judgments[idx] is True:
            old_status = comparison_rows[row_idx]["items"][field_name]["status"]
            comparison_rows[row_idx]["items"][field_name]["status"] = "match"
            updated_count += 1
            logger.debug(
                "意味比較: [%s] %s: %s → match（旧: %s）",
                comparison_rows[row_idx].get("技術文書", ""),
                field_name,
                target["extracted"],
                old_status,
            )

    logger.info(
        "意味比較: %d / %d 件を match に更新しました", updated_count, len(targets)
    )

    return comparison_rows


# ---------------------------------------------------------------------------
# 差分コメント生成
# ---------------------------------------------------------------------------

_DIFF_COMMENT_PROMPT_TEMPLATE = """\
あなたは照明器具の技術仕様値を比較する専門家です。

以下の各項目について、「正解値」と「抽出値」の具体的な差分を日本語で簡潔に説明してください。

## ルール
- 完全一致の項目は含まれていません。すべて何らかの差分があります。
- 判定がOKの場合: 意味的には同じだが表記が異なる点を説明（例: 「接頭辞ACが付加されています」「単位表記がkg→gに異なりますが値は同じです」）
- 判定がNGの場合: 何が違うのか具体的に説明（例: 「色温度が3000Kではなく4000Kになっています」「200V, 242Vが余分に読み取られています」「(白色)が読み取れていません」）
- 1項目につき1〜2文で簡潔に。

## 判定対象
{table}

JSON配列で出力:
[{{"index": 1, "comment": "差分の説明"}}, ...]
"""


def _build_diff_table(items: list[dict]) -> str:
    """差分コメント用テーブルをMarkdown形式で構築する。"""
    lines = [
        "| No. | 項目名 | 正解値 | 抽出値 | 判定 |",
        "|-----|--------|--------|--------|------|",
    ]
    for item in items:
        status_label = "OK" if item["status"] == "match" else "NG"
        lines.append(
            f"| {item['index']} | {item['field']} | {item['correct']} | {item['extracted']} | {status_label} |"
        )
    return "\n".join(lines)


def _process_diff_batch(
    client: genai.Client,
    batch: list[dict],
) -> dict[int, str]:
    """1バッチ分の差分コメントをGeminiで生成する。"""
    table = _build_diff_table(batch)
    prompt = _DIFF_COMMENT_PROMPT_TEMPLATE.format(table=table)

    try:
        response_text = _call_gemini(client, prompt)
        results = _extract_json_from_text(response_text)
        if not results:
            logger.warning("差分コメント: JSONの抽出に失敗しました。レスポンス: %s", response_text[:200])
            return {}

        comments: dict[int, str] = {}
        for item in results:
            idx = item.get("index")
            comment = item.get("comment")
            if idx is not None and comment:
                comments[int(idx)] = str(comment)
        return comments

    except Exception:
        logger.exception("差分コメント: Gemini API呼び出しに失敗しました")
        return {}


def generate_diff_comments(
    client: genai.Client,
    comparison_rows: list[dict],
    field_names: list[str],
) -> list[dict]:
    """完全一致でない項目について、Geminiで差分コメントを生成する。

    comparison_rows の各アイテムに "diff_comment" キーを追加する。

    Args:
        client: google.genai.Client インスタンス
        comparison_rows: compare_by_doc_type() の返り値
        field_names: 項目名リスト

    Returns:
        同じ構造の comparison_rows（diff_commentが追加済み）
    """
    from comparison import normalize_for_comparison

    # 1. 対象項目を収集（完全一致でない & N/Aでないもの）
    targets: list[dict] = []
    target_locations: list[tuple[int, str]] = []

    for row_idx, row_data in enumerate(comparison_rows):
        items = row_data.get("items", {})
        for field_name in field_names:
            item_info = items.get(field_name, {})
            status = item_info.get("status", "")

            # N/Aはスキップ
            if status == "na":
                item_info.setdefault("diff_comment", "")
                continue

            extracted = item_info.get("extracted", "")
            correct = item_info.get("correct", "")

            # 正規化後の完全一致チェック
            ext_norm = normalize_for_comparison(extracted).lower()
            cor_norm = normalize_for_comparison(correct).lower()

            if ext_norm == cor_norm:
                # 完全一致 → コメント不要
                item_info.setdefault("diff_comment", "")
                continue

            # 差分あり → Geminiに投げる対象
            if extracted and correct and extracted != "-" and correct != "-":
                targets.append({
                    "index": len(targets) + 1,
                    "field": field_name,
                    "correct": correct,
                    "extracted": extracted,
                    "status": status,
                })
                target_locations.append((row_idx, field_name))
            else:
                item_info.setdefault("diff_comment", "")

    # 2. 対象がなければそのまま返す
    if not targets:
        logger.info("差分コメント: 対象項目なし（全て完全一致またはN/A）")
        return comparison_rows

    logger.info("差分コメント: %d 件の項目についてGeminiで差分コメントを生成します", len(targets))

    # 3. バッチに分割
    batches: list[list[dict]] = []
    for i in range(0, len(targets), _BATCH_SIZE):
        batches.append(targets[i : i + _BATCH_SIZE])

    # 4. 並列でバッチ処理
    all_comments: dict[int, str] = {}
    with ThreadPoolExecutor(max_workers=3) as executor:
        future_to_batch = {
            executor.submit(_process_diff_batch, client, batch): batch
            for batch in batches
        }
        for future in as_completed(future_to_batch):
            batch_comments = future.result()
            all_comments.update(batch_comments)

    # 5. 結果を反映
    comment_count = 0
    for target, (row_idx, field_name) in zip(targets, target_locations):
        idx = target["index"]
        comment = all_comments.get(idx, "")
        comparison_rows[row_idx]["items"][field_name]["diff_comment"] = comment
        if comment:
            comment_count += 1

    logger.info("差分コメント: %d / %d 件のコメントを生成しました", comment_count, len(targets))

    return comparison_rows
