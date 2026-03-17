"""Agentic Vision抽出 + BBox位置取得"""

import json
import time

from google import genai
from google.genai import types
from tenacity import retry, stop_after_attempt, wait_exponential

from prompt import build_bbox_prompt

MODEL = "gemini-3-flash-preview"


def extract_json_from_response(response) -> list[dict] | None:
    """GeminiレスポンスからJSON配列を抽出する。"""
    text_parts: list[str] = []
    if response.candidates:
        for part in response.candidates[0].content.parts:
            if hasattr(part, "text") and part.text:
                text_parts.append(part.text)

    for text in reversed(text_parts):
        cleaned = text.strip()
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
                continue
    return None


def get_code_blocks(response) -> list[str]:
    """レスポンスから executable_code パートを全て抽出する。"""
    blocks: list[str] = []
    if not response.candidates:
        return blocks
    for part in response.candidates[0].content.parts:
        if hasattr(part, "executable_code") and part.executable_code:
            blocks.append(part.executable_code.code)
    return blocks


def get_response_text(response) -> str:
    """レスポンスのテキストパートを結合して返す。"""
    texts: list[str] = []
    if not response.candidates:
        return ""
    for part in response.candidates[0].content.parts:
        if hasattr(part, "text") and part.text:
            texts.append(part.text)
    return "\n".join(texts)


@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=2, min=4, max=60),
    reraise=True,
)
def _call_gemini_extract(
    client: genai.Client, image_data: bytes, mime_type: str, extraction_prompt: str
):
    """Step1: Agentic Vision で値を抽出する。"""
    return client.models.generate_content(
        model=MODEL,
        contents=[
            types.Part.from_bytes(data=image_data, mime_type=mime_type),
            extraction_prompt,
        ],
        config=types.GenerateContentConfig(
            tools=[types.Tool(code_execution=types.ToolCodeExecution)],
        ),
    )


@retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=2, min=4, max=60),
    reraise=True,
)
def _call_gemini_bbox(
    client: genai.Client, image_data: bytes, mime_type: str, prompt: str
):
    """Step2: 抽出済み値の位置をバウンディングボックスで取得する。"""
    return client.models.generate_content(
        model=MODEL,
        contents=[
            types.Part.from_bytes(data=image_data, mime_type=mime_type),
            prompt,
        ],
    )


def extract_from_image_bytes(
    client: genai.Client,
    image_data: bytes,
    mime_type: str,
    extraction_prompt: str,
    fields: list[dict],
    status_fn=None,
) -> dict:
    """画像バイトデータから2ステップで抽出する。

    Step1: Agentic Vision で値を正確に抽出
    Step2: 抽出した値の画像内位置（BBox）を取得
    """
    # --- Step 1: 値の抽出 ---
    if status_fn:
        status_fn("Step 1/2: Agentic Vision で値を抽出中...")
    start = time.perf_counter()
    response1 = _call_gemini_extract(client, image_data, mime_type, extraction_prompt)
    elapsed1 = time.perf_counter() - start

    tokens_in_1 = getattr(response1.usage_metadata, "prompt_token_count", 0) or 0
    tokens_out_1 = getattr(response1.usage_metadata, "candidates_token_count", 0) or 0
    code_blocks = get_code_blocks(response1)
    data = extract_json_from_response(response1)
    response_text = get_response_text(response1)

    if not data:
        return {
            "data": [],
            "bboxes": [],
            "code_blocks": code_blocks,
            "response_text": response_text,
            "tokens_in": tokens_in_1,
            "tokens_out": tokens_out_1,
            "elapsed": elapsed1,
        }

    # --- Step 2: BBox取得 ---
    if status_fn:
        status_fn("Step 2/2: 読み取り位置（BBox）を特定中...")
    bbox_prompt = build_bbox_prompt(data, fields)
    start2 = time.perf_counter()
    response2 = _call_gemini_bbox(client, image_data, mime_type, bbox_prompt)
    elapsed2 = time.perf_counter() - start2

    tokens_in_2 = getattr(response2.usage_metadata, "prompt_token_count", 0) or 0
    tokens_out_2 = getattr(response2.usage_metadata, "candidates_token_count", 0) or 0
    bbox_data = extract_json_from_response(response2)

    return {
        "data": data,
        "bboxes": bbox_data or [],
        "code_blocks": code_blocks,
        "response_text": response_text,
        "bbox_response_text": get_response_text(response2),
        "tokens_in": tokens_in_1 + tokens_in_2,
        "tokens_out": tokens_out_1 + tokens_out_2,
        "elapsed": elapsed1 + elapsed2,
    }


def build_draw_regions(
    extracted_data: list[dict],
    bbox_data: list[dict],
    img_w: int,
    img_h: int,
    fields: list[dict],
) -> list[dict]:
    """抽出結果とBBoxレスポンスを組み合わせて描画用リストを生成する。

    Returns:
        [{"bbox": (x1,y1,x2,y2), "label": "品番", "display_label": "品番", "value": "..."}, ...]
    """
    field_names = [f["name"] for f in fields]
    regions: list[dict] = []

    for item_idx, item in enumerate(extracted_data):
        bbox_item = bbox_data[item_idx] if item_idx < len(bbox_data) else {}
        # 最初の項目を品番的な識別子として使用
        hinban = item.get(field_names[0], "?") if field_names else "?"

        for name in field_names:
            val = item.get(name, "-")
            if not val or val == "-":
                continue

            bbox_key = f"{name}_bbox"
            raw_bbox = bbox_item.get(bbox_key)
            if not raw_bbox:
                continue

            # 文字列が返ってくるケースに対応
            if isinstance(raw_bbox, str):
                raw_bbox = raw_bbox.strip()
                if raw_bbox == "null":
                    continue
                try:
                    raw_bbox = json.loads(raw_bbox)
                except (json.JSONDecodeError, ValueError):
                    continue

            if not isinstance(raw_bbox, list) or len(raw_bbox) != 4:
                continue

            try:
                ymin, xmin, ymax, xmax = [int(v) for v in raw_bbox]
            except (ValueError, TypeError):
                continue

            # 正規化座標(0-1000) → ピクセル座標
            x1 = int(xmin * img_w / 1000)
            y1 = int(ymin * img_h / 1000)
            x2 = int(xmax * img_w / 1000)
            y2 = int(ymax * img_h / 1000)

            # 妥当性チェック
            x1, y1 = max(0, x1), max(0, y1)
            x2, y2 = min(img_w, x2), min(img_h, y2)
            if x2 <= x1 or y2 <= y1:
                continue

            label_text = name
            if len(extracted_data) > 1:
                label_text = f"{name}({hinban})"

            regions.append(
                {
                    "bbox": (x1, y1, x2, y2),
                    "label": name,
                    "display_label": label_text,
                    "value": val,
                }
            )

    return regions
