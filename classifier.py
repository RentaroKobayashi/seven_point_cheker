"""画像分類ロジック（6分類に自動判定）"""

import json
import re
import unicodedata

from google import genai
from google.genai import types
from tenacity import retry, stop_after_attempt, wait_exponential

MODEL = "gemini-3-flash-preview"

IMAGE_CATEGORIES = ["カタログ", "組立図", "承認図", "器具銘板", "取扱説明書", "外装ラベル"]

# ファイル名キーワード → カテゴリのマッピング（先にマッチした方が優先）
_FILENAME_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"カタログ", re.IGNORECASE), "カタログ"),
    (re.compile(r"catalog", re.IGNORECASE), "カタログ"),
    (re.compile(r"承認図", re.IGNORECASE), "承認図"),
    (re.compile(r"組立図|組み立て図", re.IGNORECASE), "組立図"),
    (re.compile(r"器具銘板|銘板", re.IGNORECASE), "器具銘板"),
    (re.compile(r"取扱説明|取説|説明書", re.IGNORECASE), "取扱説明書"),
    (re.compile(r"外装ラベル|外装表示|外装", re.IGNORECASE), "外装ラベル"),
    (re.compile(r"図面", re.IGNORECASE), "組立図"),
]


def classify_by_filename(filename: str) -> dict | None:
    """ファイル名からカテゴリを判定する。判定できなければ None。

    macOS等でファイル名がNFD形式（濁点が結合文字）になる場合に備え、
    NFC正規化してからパターンマッチを行う。
    """
    normalized = unicodedata.normalize("NFC", filename)
    for pattern, category in _FILENAME_PATTERNS:
        if pattern.search(normalized):
            return {"category": category, "confidence": "filename"}
    return None

CLASSIFICATION_PROMPT = f"""あなたは照明器具の技術資料を分類する専門家です。

この画像を以下の6カテゴリのいずれかに分類してください:

1. **カタログ** - 製品のカタログ情報・スペック一覧。品番・電圧・光源・色温度など複数の仕様が表形式やリスト形式でまとまっている。他の技術資料の照合基準となるマスターデータ
2. **組立図** - 器具の組立方法や部品配置を示す図面
3. **承認図** - 製品の仕様・寸法などの承認用図面
4. **器具銘板** - 製品に貼付される銘板（型番・定格など記載）
5. **取扱説明書** - 使用方法・注意事項などの説明資料
6. **外装ラベル** - 外箱に貼付されるラベル（品番・仕様など記載）

### カタログの判定基準
- 多数の仕様項目が体系的にまとまっている
- 製品の全体像がわかる情報量がある
- 他の資料（図面・銘板等）の個別情報とは異なり、総合的な製品情報が記載されている

JSON形式で出力してください:
```json
{{"category": "カテゴリ名", "confidence": "high|medium|low"}}
```

- confidence: 判定の確信度
  - high: 明確にそのカテゴリと判断できる
  - medium: おそらくそのカテゴリだが別の可能性もある
  - low: 判断が難しい
"""


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=2, min=4, max=30),
    reraise=True,
)
def classify_image(
    client: genai.Client, image_data: bytes, mime_type: str
) -> dict:
    """画像を5分類に自動判定する。

    Returns:
        {"category": "承認図", "confidence": "high|medium|low"}
    """
    response = client.models.generate_content(
        model=MODEL,
        contents=[
            types.Part.from_bytes(data=image_data, mime_type=mime_type),
            CLASSIFICATION_PROMPT,
        ],
    )

    # レスポンスからJSON抽出
    text = ""
    if response.candidates:
        for part in response.candidates[0].content.parts:
            if hasattr(part, "text") and part.text:
                text += part.text

    # JSONブロック抽出
    cleaned = text.strip()
    if "```json" in cleaned:
        cleaned = cleaned.split("```json", 1)[1].split("```", 1)[0].strip()
    elif "```" in cleaned:
        cleaned = cleaned.split("```", 1)[1].split("```", 1)[0].strip()

    try:
        result = json.loads(cleaned)
        # カテゴリが有効かチェック
        if result.get("category") not in IMAGE_CATEGORIES:
            result["category"] = "取扱説明書"  # デフォルト
            result["confidence"] = "low"
        return result
    except (json.JSONDecodeError, KeyError):
        return {"category": "取扱説明書", "confidence": "low"}
