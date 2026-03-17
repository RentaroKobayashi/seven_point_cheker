"""UI描画関数群（BBox描画、マトリクス、正答率サマリー等）"""

import os

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

from classifier import IMAGE_CATEGORIES


# ---------------------------------------------------------------------------
# フォント
# ---------------------------------------------------------------------------
def _load_font(size: int = 14) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    """macOS/Linux対応のフォント読み込み"""
    font_paths = [
        "/Library/Fonts/Arial Unicode.ttf",
        "/System/Library/Fonts/ヒラギノ角ゴシック W4.ttc",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for fp in font_paths:
        try:
            return ImageFont.truetype(fp, size)
        except OSError:
            continue
    return ImageFont.load_default()


# ---------------------------------------------------------------------------
# BBox描画
# ---------------------------------------------------------------------------
def draw_bboxes(
    image: Image.Image,
    regions: list[dict],
    label_colors: dict[str, str],
    fields: list[dict],
) -> Image.Image:
    """元画像にバウンディングボックス（枠線+項目番号）を描画して返す。"""
    annotated = image.copy()
    draw = ImageDraw.Draw(annotated)
    font_small = _load_font(11)

    field_to_num = {f["name"]: i for i, f in enumerate(fields, 1)}

    for region in regions:
        bbox = region["bbox"]
        field = region.get("label", "")
        color = label_colors.get(field, "#BDBDBD")
        num = field_to_num.get(field, 0)

        x1, y1, x2, y2 = bbox
        draw.rectangle([x1, y1, x2, y2], outline=color, width=2)

        # 左上に項目番号バッジ
        num_text = str(num)
        num_bbox = draw.textbbox((0, 0), num_text, font=font_small)
        num_w = num_bbox[2] - num_bbox[0] + 6
        num_h = num_bbox[3] - num_bbox[1] + 4
        draw.rectangle([x1, y1, x1 + num_w, y1 + num_h], fill=color)
        draw.text((x1 + 3, y1 + 1), num_text, fill="white", font=font_small)

    return annotated


# ---------------------------------------------------------------------------
# 画像分類結果表示
# ---------------------------------------------------------------------------
def render_classification_results(classification: dict):
    """画像分類結果を5列で表示する。"""
    if not classification:
        return

    st.subheader("画像分類結果")

    # カテゴリ別にグループ化
    grouped: dict[str, list[str]] = {cat: [] for cat in IMAGE_CATEGORIES}
    for img_name, info in classification.items():
        cat = info.get("category", "取扱説明書")
        grouped.setdefault(cat, []).append(img_name)

    cols = st.columns(len(IMAGE_CATEGORIES))
    for col, cat in zip(cols, IMAGE_CATEGORIES):
        with col:
            imgs = grouped.get(cat, [])
            st.markdown(f"**{cat}** ({len(imgs)})")
            for name in imgs:
                conf = classification[name].get("confidence", "")
                conf_emoji = {"high": "", "medium": " [?]", "low": " [??]"}.get(
                    conf, ""
                )
                st.markdown(f"- {name}{conf_emoji}")


# ---------------------------------------------------------------------------
# 画像別結果表示
# ---------------------------------------------------------------------------
def render_image_result(
    img_name: str,
    result: dict,
    fields: list[dict],
    label_colors: dict[str, str],
):
    """1画像分の結果をレンダリングする。"""
    # エラーケース
    if "error" in result:
        st.error(f"エラー: {result['error']}")
        if result.get("original_image"):
            st.image(result["original_image"], caption=img_name, use_container_width=True)
        return

    # メトリクス行
    cols_metric = st.columns(4)
    cols_metric[0].metric("入力トークン", f"{result.get('tokens_in', 0):,}")
    cols_metric[1].metric("出力トークン", f"{result.get('tokens_out', 0):,}")
    cols_metric[2].metric("処理時間", f"{result.get('elapsed', 0):.1f}秒")
    cols_metric[3].metric("検出BBox数", len(result.get("draw_regions", [])))

    # メインコンテンツ: 画像 + テーブル
    col_img, col_table = st.columns([3, 2])

    with col_img:
        if result.get("draw_regions"):
            st.image(
                result["annotated_image"],
                caption=f"{img_name} — 読み取り位置",
                use_container_width=True,
            )
        elif result.get("original_image"):
            st.info("バウンディングボックスを取得できませんでした。元画像を表示します。")
            st.image(result["original_image"], caption=img_name, use_container_width=True)

    with col_table:
        if result.get("data"):
            st.dataframe(result["data"], use_container_width=True, hide_index=True)
        else:
            st.warning("抽出データがありません。")

    # 凡例
    if result.get("draw_regions"):
        field_values: dict[str, list[str]] = {}
        for region in result["draw_regions"]:
            field = region.get("label", "")
            val = region.get("value", "")
            hinban = region.get("display_label", field)
            entry = (
                val
                if len(result.get("data", [])) <= 1
                else f'{val} ({hinban.split("(", 1)[-1].rstrip(")")})'
            )
            field_values.setdefault(field, []).append(entry)

        legend_lines: list[str] = []
        for num, f in enumerate(fields, 1):
            name = f["name"]
            if name not in field_values:
                continue
            color = label_colors.get(name, "#BDBDBD")
            values_str = " / ".join(field_values[name])
            legend_lines.append(
                f'<span style="display:inline-block;background:{color};color:#fff;'
                f'border-radius:3px;padding:1px 6px;margin:2px;font-size:13px;">'
                f"{num}</span> "
                f"<b>{name}</b>: {values_str}"
            )
        st.markdown("<br>".join(legend_lines), unsafe_allow_html=True)

    # Expander: コード実行ログ
    if result.get("code_blocks"):
        with st.expander("Agentic Vision コード実行ログ", expanded=False):
            for i, block in enumerate(result["code_blocks"], 1):
                st.markdown(f"**code_execution #{i}**")
                st.code(block, language="python")

    # Expander: BBox生レスポンス
    if result.get("bbox_response_text"):
        with st.expander("BBox検出 生レスポンス", expanded=False):
            st.text(result["bbox_response_text"])

    # Expander: 抽出生レスポンス
    if result.get("response_text"):
        with st.expander("抽出 生レスポンス", expanded=False):
            st.text(result["response_text"])


# ---------------------------------------------------------------------------
# 照合マトリクス
# ---------------------------------------------------------------------------
def render_comparison_matrix(
    processed_images: list[str],
    results: dict,
    fields: list[dict],
):
    """全画像の結果を横並び比較するマトリクスを表示する。"""
    field_names = [f["name"] for f in fields]
    rows: list[dict[str, str]] = []

    for img_name in processed_images:
        result = results.get(img_name, {})
        data = result.get("data", [])
        source = os.path.splitext(img_name)[0]

        if not data:
            row = {"ソース": source}
            for name in field_names:
                row[name] = "-"
            rows.append(row)
        else:
            for item in data:
                row = {"ソース": source}
                for name in field_names:
                    row[name] = item.get(name, "-") or "-"
                rows.append(row)

    if not rows:
        st.info("表示するデータがありません。")
        return

    df = pd.DataFrame(rows)

    def highlight_mismatch(col):
        if col.name == "ソース":
            return [""] * len(col)
        values = col[col != "-"].unique()
        if len(values) <= 1:
            return [""] * len(col)
        mode_val = col[col != "-"].mode()
        mode_val = mode_val.iloc[0] if len(mode_val) > 0 else None
        return [
            "background-color: #FFCDD2" if v != mode_val and v != "-" else ""
            for v in col
        ]

    styled = df.style.apply(highlight_mismatch, axis=0)
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# 正解データ比較マトリクス
# ---------------------------------------------------------------------------
def render_correct_comparison(comparison: dict, fields: list[dict]):
    """正解データ比較マトリクス（色分け付き）を表示する。"""
    field_names = [f["name"] for f in fields]

    rows: list[dict] = []
    for img_name, img_data in comparison.items():
        source = os.path.splitext(img_name)[0]
        for item_key, item_data in img_data.items():
            row = {"ソース": source, "品番": item_key}
            for name in field_names:
                if name in item_data:
                    info = item_data[name]
                    status = info["status"]
                    ext = info["extracted"]
                    cor = info["correct"]

                    icon = {"match": "OK", "mismatch": "NG", "na": "-"}.get(
                        status, "-"
                    )
                    if status == "na":
                        row[f"{name}_display"] = "-"
                    else:
                        row[f"{name}_display"] = f"{icon} {ext}"
                    row[f"{name}_correct"] = cor
                    row[f"{name}_status"] = status
                else:
                    row[f"{name}_display"] = "-"
                    row[f"{name}_correct"] = ""
                    row[f"{name}_status"] = "na"
            rows.append(row)

    if not rows:
        st.info("比較データがありません。")
        return

    # 表示用DataFrame構築
    display_data: list[dict] = []
    for row in rows:
        display_row = {"ソース": row["ソース"], "品番": row["品番"]}
        for name in field_names:
            display_row[name] = row.get(f"{name}_display", "-")
        display_data.append(display_row)

    df = pd.DataFrame(display_data)

    def color_cells(col):
        if col.name in ("ソース", "品番"):
            return [""] * len(col)
        styles = []
        for idx in range(len(col)):
            field_name = col.name
            status = rows[idx].get(f"{field_name}_status", "na")
            if status == "match":
                styles.append("background-color: #D5F5E3; color: #1E8449")
            elif status == "mismatch":
                styles.append("background-color: #FADBD8; color: #C0392B")
            else:
                styles.append("color: #AAB7B8")
        return styles

    styled = df.style.apply(color_cells, axis=0)
    st.dataframe(styled, use_container_width=True, hide_index=True)

    # 正解値の参照表
    with st.expander("正解データ一覧", expanded=False):
        correct_rows: list[dict] = []
        for row in rows:
            cr = {"ソース": row["ソース"], "品番": row["品番"]}
            for name in field_names:
                cr[name] = row.get(f"{name}_correct", "")
            correct_rows.append(cr)
        st.dataframe(pd.DataFrame(correct_rows), use_container_width=True, hide_index=True)


# ---------------------------------------------------------------------------
# 正答率サマリー
# ---------------------------------------------------------------------------
def render_accuracy_summary(accuracy: dict, fields: list[dict]):
    """正答率サマリーを表示する。"""
    field_names = [f["name"] for f in fields]
    total = accuracy.get("total", {})
    by_field = accuracy.get("by_field", {})
    by_source = accuracy.get("by_source", {})

    # 総合メトリクス
    st.subheader("正答率サマリー")
    cols = st.columns(3)
    cols[0].metric("総合正答率", f"{total.get('accuracy', 0) * 100:.1f}%")
    cols[1].metric("完全一致", f"{total.get('correct', 0)}")
    cols[2].metric("比較項目数", f"{total.get('total', 0)}")

    # 項目別正答率
    st.markdown("#### 項目別正答率")
    field_data = []
    for name in field_names:
        stats = by_field.get(name, {})
        if stats.get("total", 0) > 0:
            field_data.append(
                {
                    "項目": name,
                    "正答率": f"{stats.get('accuracy', 0) * 100:.1f}%",
                    "完全一致": stats.get("correct", 0),
                    "合計": stats.get("total", 0),
                }
            )
    if field_data:
        st.dataframe(pd.DataFrame(field_data), use_container_width=True, hide_index=True)

    # 資料別正答率
    st.markdown("#### 資料別正答率")
    source_data = []
    for src, stats in by_source.items():
        source_name = os.path.splitext(src)[0]
        if stats.get("total", 0) > 0:
            source_data.append(
                {
                    "資料": source_name,
                    "正答率": f"{stats.get('accuracy', 0) * 100:.1f}%",
                    "完全一致": stats.get("correct", 0),
                    "合計": stats.get("total", 0),
                }
            )
    if source_data:
        st.dataframe(pd.DataFrame(source_data), use_container_width=True, hide_index=True)
