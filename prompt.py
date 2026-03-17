"""プロンプトテンプレート + 動的生成関数

プロンプト設計の基本原則:
- 項目リスト = 「何を読み取るか」の定義（主役）
- カタログ値 = 「こういう値があるはず」というヒント（補助）
- 品番 = カタログから確定的に取得（メイン品番のみ、子品番は不要）
- 正解データ = 常にマスターExcelから取得（プロンプトとは無関係）
"""

import json

# ---------------------------------------------------------------------------
# カタログ抽出プロンプト
# カタログからメイン品番と各項目の値を読み取る。
# ここで読み取った値は、他の5資料を読み取る際の「ヒント」になる。
# ---------------------------------------------------------------------------
CATALOG_EXTRACTION_PROMPT_TEMPLATE = """
<role>
あなたは照明器具のカタログ情報を正確に読み取る専門家です。
</role>

<task>
このカタログ画像から、対象品番の情報を読み取ってください。

{anchor_section}

### 読み取り手順（必ずこの順序で実行）

**Step 1: 対象品番の行を特定する**
画像内で対象品番を探してください。カタログには似た品番が多数並んでいることがあります。
- 対象品番と上記の製品特徴（色温度・光束・タイプ等）を手がかりに、正しい行を見つけること
- 似た品番が並んでいる場合は、**必ずコード実行で対象行をクロップ・ズームし、品番を1文字ずつ照合して完全一致を確認**してからStep 2に進むこと
- 品番が1文字でも違えば別製品。絶対に隣の行と間違えないこと

**Step 2: 特定した行からのみ値を読み取る**
Step 1で特定した品番の行・セクション・ブロックの中からのみ、各項目の値を読み取ってください。
隣接する別品番の行から値を取得してはいけません。

**重要: 品番の取得ルール**
- 「品名・品番」欄に記載されているメインの製品品番のみを品番として抽出すること
- 子品番・派生品番の一覧は品番ではない。取得しないこと
- 1つのメイン品番につき1エントリを出力すること

ここで読み取った値は、他の技術資料（承認図・銘板等）を読み取る際のヒントとして使われます。
正確さが重要です。ただし画像にない情報を無理に推測しないこと。
</task>

<code_execution>
**必ず以下の手順でコード実行すること（省略不可）:**

1. まず画像全体を表示して対象品番の位置を目視で確認する
2. 対象品番がある行をクロップ・ズームして品番を1文字ずつ確認する
3. 品番が完全一致することを確認してから、その行の各項目値を読み取る

```python
from PIL import Image
img = Image.open("/tmp/input_image.jpg")
# 1. 対象品番の行をクロップ
region = img.crop((x1, y1, x2, y2))
# 2. 3倍にズームして品番を確認
zoomed = region.resize((region.width * 3, region.height * 3), Image.LANCZOS)
zoomed.show()
```

カタログの値はヒントの基準になるため、1文字の読み間違いも許されません。
似た品番が多数ある場合は特に注意し、必ずズームで確認すること。
</code_execution>

{items_section}

<rules>
- 画像に記載がない項目は "-" とする
- 推測・補完はしない。実際に見える値のみ抽出する
- 単位は記載通りに含める（例: 100V, 3000K, 0.5kg）
- 読めないが何か書いてある場合は "不鮮明" とする
- 品番は必ずメインの製品品番1つだけを取得する
- **別品番の行・セクションから値を取らないこと**
- 対象品番と製品特徴（色温度・光束等）の両方が一致する行から値を取ること
</rules>

{output_section}
"""

# ---------------------------------------------------------------------------
# 技術資料の読み取りプロンプト（カタログ値付き項目テーブル）
# 各項目にカタログ値を埋め込んで「何を探すか + 参考値」をセットで渡す。
# ---------------------------------------------------------------------------
VERIFICATION_PROMPT_TEMPLATE = """
<role>
あなたは照明器具の技術資料を正確に読み取る専門家です。
</role>

<task>
この技術資料（{doc_type}）から、対象品番の情報を読み取ってください。

**対象品番: {target_hinban}**

### 読み取り手順（必ずこの順序で実行）

**Step 0: コード実行で品番の位置を特定する（必須）**
まず**必ずコード実行**で画像内の品番「{target_hinban}」を探してください。
1. 画像全体を表示して品番の位置を目視で確認する
2. 品番が記載されている行・セクションをクロップ・ズームする
3. 品番を1文字ずつ照合し、対象品番と完全一致することを確認する

品番が確認できたら、**そのセクションの範囲を確定**してから次に進むこと。
品番が見つからない場合でも、資料のタイトルや型番表記から判断してください。
{doc_type_note}

**Step 1: 対象品番のセクションからのみ値を読み取る**
Step 0で確定したセクション・表・仕様欄の中からのみ各項目を読み取ってください。
**別の品番や別製品の情報は絶対に取得しないこと。**

**Step 2: カタログ値と照らし合わせる**
各項目にはカタログから読み取った参考値を載せています。
この資料の記載とカタログ値が異なっていても、**この資料に書いてある値をそのまま報告**してください。

{items_with_catalog}

### 表記の違いについて
技術資料ごとに書式・省略・単位表記は異なります。以下は同じ情報の例です:
- 範囲表記 ↔ 個別列挙（例: 「100~242V」と「AC100V, AC200V」）
- 省略・付加（例: 「NTS32711S」と「NTS(H)32711S」）
- 単位・接頭辞の有無（例: 「3000K」と「色温度 3000」）
- 正式名称 ↔ 略称（例: 「電源ユニット一体型」と「電源内蔵」）
</task>

<code_execution>
**Step 0で必ずコード実行すること（省略不可）:**

```python
from PIL import Image
img = Image.open("/tmp/input_image.jpg")
# 1. 画像全体を表示して品番の位置を確認
img.show()
# 2. 対象品番のセクションをクロップ・ズーム
region = img.crop((x1, y1, x2, y2))
zoomed = region.resize((region.width * 3, region.height * 3), Image.LANCZOS)
zoomed.show()
```

品番の特定後も、文字が小さい・不鮮明な場合は該当箇所をクロップ・ズームして確認すること。
</code_execution>

<rules>
- 画像に記載がない項目は "-" とする
- 推測・補完はしない。この資料に実際に見える値のみ報告する
- 単位は記載通りに含める
- 読めないが何か書いてある場合は "不鮮明" とする
- カタログと異なる値が書いてあっても、この資料の表記をそのまま報告する
- **対象品番「{target_hinban}」に関する情報のみ取得すること。別品番の値は取らない**
- **対象品番1つ分の値のみ出力する。複数品番の値を「/」「／」で併記しない**
- 項目名と完全一致する表記がなくても、同義の表現を探すこと（例: 光効率=発光効率=ランプ効率=lm/W）
</rules>

{output_section}
"""

# ---------------------------------------------------------------------------
# 技術資料の読み取りプロンプト（カタログなし・フォールバック）
# 項目リストのみで抽出する。
# ---------------------------------------------------------------------------
EXTRACTION_PROMPT_TEMPLATE = """
<role>
あなたは照明器具の技術資料を正確に読み取る専門家です。
</role>

<task>
この画像から、以下の項目を読み取ってください。
該当する情報が画像にまったく存在しない場合は、空の配列 [] を返してください。
{anchor_section}
</task>

<code_execution>
画像内の文字が小さい・不鮮明な場合は、Pythonコードで該当箇所をクロップ・ズームし、
文字が確実に判読できることを確認してから値を確定してください。

```python
from PIL import Image
img = Image.open("/tmp/input_image.jpg")
region = img.crop((x1, y1, x2, y2))
region.show()
zoomed = region.resize((region.width * 3, region.height * 3), Image.LANCZOS)
zoomed.show()
```

特に銘板・仕様表の小さい数値は必ずズームして確認すること。
読めない状態で推測しないこと。
</code_execution>

{items_section}

<rules>
- 画像に記載がない項目は "-" とする
- 推測・補完はしない。実際に見える値のみ抽出する
- 単位は記載通りに含める
- 読めないが何か書いてある場合は "不鮮明" とする
- 複数の値が列挙されている項目は、代表的な1つだけを取得する（一覧を丸ごと取得しないこと）
- **複数品番が記載されている場合、対象品番1つ分の値のみ出力する。複数品番の値を「/」「／」で併記しない**
- 項目名と完全一致する表記がなくても、同義の表現を探すこと（例: 光効率=発光効率=ランプ効率=lm/W）
</rules>

{output_section}
"""

# ---------------------------------------------------------------------------
# BBox位置特定プロンプト（変更なし）
# ---------------------------------------------------------------------------
BBOX_PROMPT_TEMPLATE = """
この画像の中から、以下のテキスト/値が書かれている位置をバウンディングボックスで示してください。

## 対象の値
{items_text}

## 出力形式
各項目について、画像内でその値が記載されている領域の座標を [ymin, xmin, ymax, xmax] の形式で返してください。
座標は 0〜1000 の正規化座標です（画像の左上が [0, 0]、右下が [1000, 1000]）。

**重要**: 値が記載されている**テキスト部分そのもの**を囲むバウンディングボックスを返してください。
該当する値が画像に存在しない場合（"-" の場合）は null を返してください。

JSON形式で出力してください:
```json
{json_template}
```
"""


# ===================================================================
# ビルダー関数
# ===================================================================

def _build_items_section(fields: list[dict]) -> str:
    """項目リストから <items> セクションを構築する（共通処理）。"""
    lines = ["<items>", "| # | 項目 | 説明 |", "|---|------|------|"]
    for i, field in enumerate(fields, 1):
        desc = field.get("description", "")
        lines.append(f"| {i} | {field['name']} | {desc} |")

    hints = [f for f in fields if f.get("hint")]
    if hints:
        lines.append("")
        lines.append("### 読み取りヒント")
        for f in hints:
            lines.append(f"- **{f['name']}**: {f['hint']}")

    lines.append("</items>")
    return "\n".join(lines)


def _build_output_section(fields: list[dict], value_template: str = "読み取った") -> str:
    """<output> セクションを構築する（共通処理）。"""
    sample_obj = {f["name"]: f"({value_template}{f['name']})" for f in fields}
    sample_json = json.dumps([sample_obj], ensure_ascii=False, indent=2)

    return f"""<output>
JSON配列を ```json ``` ブロックで出力してください。
```json
{sample_json}
```
項目が画像にない場合は "-" を入れてください。
</output>"""


def build_catalog_extraction_prompt(fields: list[dict]) -> str:
    """カタログ情報から項目を抽出するプロンプトを組み立てる。

    項目リストの具体値ヒント（品番・色温度・光束等）をアンカーとして使用し、
    カタログ内の正しい行・セクションを特定させる。
    """
    # 品番ヒントと、その他の特徴ヒントを収集
    hinban_hint = ""
    other_hints: list[str] = []
    for f in fields:
        if f["name"] == "品番" and f.get("hint"):
            hinban_hint = f["hint"]
        elif f.get("hint"):
            other_hints.append(f"- {f['name']}: {f['hint']}")

    if hinban_hint:
        anchor_section = (
            f"**対象品番: {hinban_hint}**\n\n"
            f"この画像には複数の品番・製品が掲載されている可能性があります。\n"
            f"必ず品番「{hinban_hint}」の行・セクションを特定し、"
            f"その品番の情報のみを読み取ってください。\n\n"
            f"**重要: 似た品番が多数並んでいる場合、1文字ずつ照合して正しい行を見つけること。**\n"
            f"必ずコード実行で対象行をクロップ・ズームし、品番が完全一致することを確認してから値を読み取ること。"
        )
        if other_hints:
            anchor_section += (
                "\n\n### 対象製品の特徴（行の特定に使用）\n"
                + "\n".join(other_hints)
                + "\n\nこれらの特徴と一致する行を探してください。"
            )
    else:
        anchor_section = (
            "この画像から対象製品のメイン品番を特定し、"
            "その品番の情報を読み取ってください。"
        )

    return CATALOG_EXTRACTION_PROMPT_TEMPLATE.format(
        anchor_section=anchor_section,
        items_section=_build_items_section(fields),
        output_section=_build_output_section(fields),
    )


# 文書タイプ別の注意喚起テキスト
_DOC_TYPE_NOTES: dict[str, str] = {
    "器具銘板": (
        "\n\n**【銘板の注意】** 1枚の銘板に複数品番が並んでいることがあります。"
        "対象品番「{target_hinban}」の行のデータのみを読み取り、"
        "他の品番の行の値は無視してください。"
    ),
    "取扱説明書": (
        "\n\n**【取扱説明書の注意】** 複数モデル共通の説明書の場合、"
        "仕様表内に複数モデルの値が並んでいます。"
        "対象品番「{target_hinban}」の列・行のみから値を読み取り、"
        "他のモデルの値は無視してください。"
    ),
    "外装ラベル": (
        "\n\n**【外装ラベルの注意】** ラベルに複数品番が列記されていても、"
        "対象品番「{target_hinban}」の情報のみを読み取ってください。"
    ),
}


def build_verification_prompt(
    catalog_data: dict, fields: list[dict], doc_type: str
) -> str:
    """カタログ値を項目テーブルに埋め込んで技術資料読み取りプロンプトを組み立てる。

    品番アンカー方式: カタログから取得した品番を最優先アンカーとして使い、
    その品番のセクションからのみ値を読み取るようプロンプトを構築する。

    Args:
        catalog_data: カタログから抽出した値 {"品番": "NTS32711S", "電圧": "100V", ...}
        fields: 項目リスト
        doc_type: 技術文書の種別（"承認図", "器具銘板" 等）
    """
    # 品番をアンカーとして取得（項目リストのヒント → カタログ値 → 不明）
    # 項目リストのヒントは正解データ由来で確実に正しいため最優先
    target_hinban = ""
    for f in fields:
        if f["name"] == "品番" and f.get("hint"):
            target_hinban = f["hint"]
            break
    if not target_hinban or target_hinban == "-":
        target_hinban = catalog_data.get("品番", "")
    if not target_hinban or target_hinban == "-":
        target_hinban = "（不明）"

    # 文書タイプ別の注意書きを生成
    doc_type_note = _DOC_TYPE_NOTES.get(doc_type, "")
    if doc_type_note:
        doc_type_note = doc_type_note.format(target_hinban=target_hinban)

    # 項目 + カタログ値を統合したテーブルを構築
    lines = [
        "<items>",
        "| # | 項目 | 説明 | カタログ値（参考） |",
        "|---|------|------|-------------------|",
    ]
    for i, field in enumerate(fields, 1):
        name = field["name"]
        desc = field.get("description", "")
        cat_val = catalog_data.get(name, "-")
        lines.append(f"| {i} | {name} | {desc} | {cat_val} |")

    # 読み取りヒントがあれば追加
    hints = [f for f in fields if f.get("hint")]
    if hints:
        lines.append("")
        lines.append("### 読み取りヒント（項目リストの具体値）")
        for f in hints:
            lines.append(f"- **{f['name']}**: {f['hint']}")

    lines.append("</items>")
    items_with_catalog = "\n".join(lines)

    return VERIFICATION_PROMPT_TEMPLATE.format(
        doc_type=doc_type,
        target_hinban=target_hinban,
        doc_type_note=doc_type_note,
        items_with_catalog=items_with_catalog,
        output_section=_build_output_section(
            fields, value_template="この資料に記載されている"
        ),
    )


def build_extraction_prompt(fields: list[dict], target_hinban: str = "") -> str:
    """項目リストのみで抽出するプロンプトを組み立てる（カタログなしフォールバック）。

    Args:
        fields: 項目リスト
        target_hinban: 対象品番（わかっていればアンカーとして使用）
    """
    if target_hinban and target_hinban != "-":
        anchor_section = (
            f"\n**対象品番: {target_hinban}**\n"
            f"この画像に複数の品番が記載されている場合、"
            f"品番「{target_hinban}」の情報のみを読み取ってください。"
        )
    else:
        anchor_section = ""

    return EXTRACTION_PROMPT_TEMPLATE.format(
        anchor_section=anchor_section,
        items_section=_build_items_section(fields),
        output_section=_build_output_section(fields),
    )


def build_bbox_prompt(extracted_data: list[dict], fields: list[dict]) -> str:
    """抽出結果からBBox取得用のプロンプトを組み立てる。"""
    field_names = [f["name"] for f in fields]
    items_lines: list[str] = []
    json_template_items: list[dict] = []

    for i, item in enumerate(extracted_data):
        hinban = item.get("品番", item.get(field_names[0], "-") if field_names else "-")
        items_lines.append(f"\n### 製品{i + 1}: {hinban}")
        template_item: dict = {}
        for name in field_names:
            val = item.get(name, "-")
            if val and val != "-":
                items_lines.append(f"- **{name}**: 「{val}」")
                template_item[f"{name}_bbox"] = [0, 0, 0, 0]
            else:
                template_item[f"{name}_bbox"] = None
        json_template_items.append(template_item)

    items_text = "\n".join(items_lines)
    json_template = json.dumps(json_template_items, ensure_ascii=False, indent=2)

    return BBOX_PROMPT_TEMPLATE.format(
        items_text=items_text,
        json_template=json_template,
    )
