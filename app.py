"""7点照合チェッカー Streamlit アプリ

照明器具の技術資料画像から動的に指定された項目を抽出し、
画像分類・BBox可視化・照合マトリクス・正解データ比較を行う。
複数品番の一括照合にも対応。

Usage:
    uv run streamlit run app.py
"""

import io
import logging
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

import pandas as pd
import streamlit as st
from PIL import Image

# ターミナルにも処理状況を出力
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

from classifier import IMAGE_CATEGORIES, classify_by_filename
from comparison import (
    compare_by_doc_type,
    export_comparison_excel,
    export_extraction_excel,
    export_multi_product_comparison_excel,
)
from semantic_judge import generate_diff_comments, refine_with_semantic_judge
from extractor import build_draw_regions, extract_from_image_bytes
from prompt import (
    build_catalog_extraction_prompt,
    build_extraction_prompt,
    build_verification_prompt,
)
from ui_components import (
    draw_bboxes,
    render_classification_results,
    render_comparison_matrix,
    render_image_result,
)
from utils import (
    assign_colors,
    expand_uploaded_files,
    extract_product_from_filename,
    extract_product_from_items_filename,
    get_genai_client,
    group_images_by_product,
    load_all_correct_data_from_master,
    load_items_from_file,
)

CORRECT_DATA_DIR = "correct_data"
MASTER_EXCEL_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "正解データ", "７点照合正解データ_v2.xlsx"
)

# 出力ディレクトリ（プロジェクトルート直下の output/ に集約）
OUTPUT_BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
OUTPUT_EXTRACTION_DIR = os.path.join(OUTPUT_BASE_DIR, "読み取り結果")
OUTPUT_COMPARISON_DIR = os.path.join(OUTPUT_BASE_DIR, "正誤判定結果")
OUTPUT_EVIDENCE_DIR = os.path.join(OUTPUT_BASE_DIR, "evidence")


def main():
    st.set_page_config(
        page_title="7点照合チェッカー",
        layout="wide",
    )
    st.title("7点照合チェッカー")

    # --- session_state 初期化 ---
    if "product_fields" not in st.session_state:
        st.session_state.product_fields = {}  # {品番: fields} のdict
    if "default_fields" not in st.session_state:
        st.session_state.default_fields = []  # 品番不明の場合のフォールバック
    if "label_colors" not in st.session_state:
        st.session_state.label_colors = {}
    if "product_results" not in st.session_state:
        st.session_state.product_results = {}
    if "multi_comparison_excel_bytes" not in st.session_state:
        st.session_state.multi_comparison_excel_bytes = None
    if "multi_comparison_excel_filename" not in st.session_state:
        st.session_state.multi_comparison_excel_filename = ""

    # --- サイドバー ---
    with st.sidebar:
        st.header("設定")

        # ① 項目リスト アップロード（CSV or Excel、複数可）
        st.subheader("1. 項目リスト")
        items_files = st.file_uploader(
            "項目リストをアップロード（CSV / Excel、複数可）",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
            help="ファイル名が「品番_項目リスト.xlsx」の場合、自動で品番と紐づけます",
        )
        if items_files:
            logger.info("項目リスト %d 件を読み込み中...", len(items_files))
            product_fields = {}
            default_fields = []
            all_fields_flat = []  # 色割り当て用の全項目統合リスト
            for items_file in items_files:
                logger.info("  読み込み: %s", items_file.name)
                fields = load_items_from_file(items_file)
                if not fields:
                    logger.warning("  → 項目なし（スキップ）")
                    continue
                product_id = extract_product_from_items_filename(items_file.name)
                if product_id:
                    product_fields[product_id] = fields
                    logger.info("  → 品番 %s: %d 項目", product_id, len(fields))
                    st.success(f"📋 {product_id}: {len(fields)} 項目を読み込みました")
                else:
                    default_fields = fields
                    logger.info("  → デフォルト: %d 項目", len(fields))
                    st.success(f"📋 デフォルト: {len(fields)} 項目を読み込みました")
                all_fields_flat.extend(fields)
            logger.info("項目リスト読み込み完了: 品番別=%d件, デフォルト=%d項目",
                        len(product_fields), len(default_fields))

            st.session_state.product_fields = product_fields
            st.session_state.default_fields = default_fields
            # 全項目の色を統合割り当て（重複項目名は1つにまとめる）
            seen = set()
            unique_fields = []
            for f in all_fields_flat:
                if f["name"] not in seen:
                    seen.add(f["name"])
                    unique_fields.append(f)
            st.session_state.label_colors = assign_colors(unique_fields)

            # 項目リストの内容を品番ごとに表示
            for pid, pfields in product_fields.items():
                with st.expander(f"{pid} の項目 ({len(pfields)}件)"):
                    for f in pfields:
                        doc_info = ""
                        if f.get("doc_types"):
                            doc_info = f' [{", ".join(f["doc_types"])}]'
                        st.markdown(
                            f"- **{f['name']}**"
                            + (f" — {f['description']}" if f.get("description") else "")
                            + doc_info
                        )
            if default_fields:
                with st.expander(f"デフォルト項目 ({len(default_fields)}件)"):
                    for f in default_fields:
                        st.markdown(f"- **{f['name']}**")

        # ② 資料画像アップロード（ZIP対応）
        st.subheader("2. 資料画像 / PDF / ZIP")
        uploaded_files = st.file_uploader(
            "資料画像・PDF・ZIPをアップロード（複数可）",
            type=["png", "jpg", "jpeg", "pdf", "zip"],
            accept_multiple_files=True,
            help="PDFは各ページを画像に変換。ZIPはフォルダ名を品番として展開します",
        )

        # API設定
        st.subheader("3. API設定")
        api_interval = st.slider("API呼び出し間隔（秒）", min_value=1, max_value=10, value=3)

        # ③ 実行ボタン
        has_items = bool(st.session_state.product_fields) or bool(st.session_state.default_fields)
        has_images = bool(uploaded_files)
        run_disabled = not (has_items and has_images)

        logger.info("ボタン状態: has_items=%s (品番別=%d, デフォルト=%d), has_images=%s (%d件), disabled=%s",
                     has_items, len(st.session_state.product_fields),
                     len(st.session_state.default_fields),
                     has_images, len(uploaded_files) if uploaded_files else 0,
                     run_disabled)

        if not has_items:
            st.warning("項目リストをアップロードしてください")
        if not has_images:
            st.warning("画像をアップロードしてください")

        run_button = st.button(
            "抽出実行",
            type="primary",
            disabled=run_disabled,
        )

        # トークン集計
        if st.session_state.product_results:
            st.divider()
            st.subheader("処理済み")
            total_in = 0
            total_out = 0
            for prod_data in st.session_state.product_results.values():
                for r in prod_data.get("results", {}).values():
                    total_in += r.get("tokens_in", 0)
                    total_out += r.get("tokens_out", 0)
            st.metric("入力トークン合計", f"{total_in:,}")
            st.metric("出力トークン合計", f"{total_out:,}")

    # --- 抽出実行 ---
    if run_button:
        logger.info("抽出実行ボタンが押されました")
        logger.info("  uploaded_files: %s", bool(uploaded_files))
        logger.info("  product_fields: %s", list(st.session_state.product_fields.keys()))
        logger.info("  default_fields: %d 項目", len(st.session_state.default_fields))

    if run_button and uploaded_files and (st.session_state.product_fields or st.session_state.default_fields):
        logger.info("=== 抽出処理を開始 ===")
        product_fields_map = st.session_state.product_fields
        default_fields = st.session_state.default_fields
        label_colors = st.session_state.label_colors

        client = get_genai_client()

        # PDF→画像展開・ZIP展開
        expanded = expand_uploaded_files(uploaded_files)
        progress = st.progress(0, text="処理を開始します...")
        status_container = st.empty()

        if not expanded:
            st.warning("有効な画像がありません。")
        else:
            status_container.info(f"{len(expanded)} 画像を処理します（PDF/ZIP展開含む）")

        # 品番ごとにグルーピング
        product_groups = group_images_by_product(expanded)
        logger.info("初期グルーピング結果: %s",
                     {pid: len(imgs) for pid, imgs in product_groups.items()})

        # ZIPフォルダ名が連番等で品番と一致しない場合、ファイル名から品番を再検出
        product_groups = _remap_product_groups(product_groups, product_fields_map)

        product_ids = list(product_groups.keys())
        num_products = len(product_ids)
        logger.info("最終グルーピング結果: %d品番 → %s",
                     num_products,
                     {pid: len(imgs) for pid, imgs in product_groups.items()})

        if num_products > 1:
            display_ids = [pid for pid in product_ids if pid != "_default"]
            status_container.info(
                f"{num_products} 品番を検出しました: {', '.join(display_ids)}"
            )

        # 正解データを事前にロード（全品番）
        all_correct: dict | None = None
        if os.path.isfile(MASTER_EXCEL_PATH):
            try:
                all_correct = load_all_correct_data_from_master(MASTER_EXCEL_PATH)
            except Exception as e:
                st.warning(f"正解データの読み込みでエラー: {e}")

        all_product_results: dict = {}
        all_comparisons: dict[str, list[dict]] = {}  # 複数品番Excel出力用
        project_dir = os.path.dirname(os.path.abspath(__file__))
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        # ============================================================
        # 品番ごとに Phase 0〜3 を実行
        # ============================================================
        for prod_idx, product_id in enumerate(product_ids):
            # 品番に対応する項目リストを取得（なければデフォルト）
            fields = product_fields_map.get(product_id, default_fields)
            logger.info("[%d/%d] 品番=%s, fields=%d項目 (マッチ=%s)",
                        prod_idx + 1, num_products, product_id, len(fields),
                        product_id in product_fields_map)
            if not fields:
                logger.warning("品番 %s に対応する項目リストなし → スキップ", product_id)
                st.warning(f"品番 {product_id} に対応する項目リストがありません。スキップします。")
                continue
            label_colors = assign_colors(fields)  # 品番ごとに色を再割り当て

            prod_images = product_groups[product_id]
            prod_base = prod_idx / num_products
            prod_range = 1.0 / num_products

            display_product = product_id if product_id != "_default" else "未分類"
            prod_prefix = f"[{display_product}] " if num_products > 1 else ""

            if num_products > 1:
                status_container.info(
                    f"[{prod_idx + 1}/{num_products}] 品番: {display_product} を処理中..."
                )

            classification_results: dict = {}
            extraction_results: dict = {}
            uploaded_image_data: dict = {}

            # ============================================================
            # Phase 0: 画像分類（ファイル名で判定）
            # ============================================================
            unknown_count = 0
            logger.info("  Phase 0 開始: %d件を分類中...", len(prod_images))
            for img_info in prod_images:
                name = img_info["name"]
                uploaded_image_data[name] = {
                    "data": img_info["data"],
                    "mime_type": img_info["mime_type"],
                }
                fn_result = classify_by_filename(name)
                if fn_result:
                    classification_results[name] = fn_result
                    logger.info("    分類: %s → %s", name, fn_result["category"])
                else:
                    classification_results[name] = {
                        "category": "取扱説明書",
                        "confidence": "low",
                    }
                    unknown_count += 1
                    logger.warning("    分類失敗: %s → デフォルト(取扱説明書)", name)

            # カテゴリ別の件数をログ出力
            cat_counts = {}
            for info in classification_results.values():
                c = info.get("category", "?")
                cat_counts[c] = cat_counts.get(c, 0) + 1
            logger.info("  Phase 0 完了: %d件分類, 不明=%d件, 内訳=%s",
                        len(prod_images), unknown_count, cat_counts)
            if unknown_count > 0:
                st.warning(
                    f"{prod_prefix}{unknown_count} 件のファイルがファイル名で分類できませんでした。"
                    "下部の手動修正UIでカテゴリを設定してください。"
                )
            progress.progress(
                prod_base + 0.2 * prod_range,
                text=f"{prod_prefix}分類完了（{len(prod_images)} 件）",
            )

            # カタログ画像と技術資料を分離
            catalog_images = [
                e for e in prod_images
                if classification_results.get(e["name"], {}).get("category") == "カタログ"
            ]
            other_images = [
                e for e in prod_images
                if classification_results.get(e["name"], {}).get("category") != "カタログ"
            ]

            # ============================================================
            # Phase 1: カタログを先行抽出（正解データ取得）
            # ============================================================
            catalog_data_merged: dict = {}

            if catalog_images:
                logger.info("  Phase 1: カタログ %d件を抽出中...", len(catalog_images))
                status_container.info(
                    f"{prod_prefix}カタログ情報 {len(catalog_images)} 件を抽出中..."
                )
                catalog_prompt = build_catalog_extraction_prompt(fields)

                def process_catalog_image(img_info: dict) -> tuple[str, dict]:
                    img_name = img_info["name"]
                    try:
                        result = extract_from_image_bytes(
                            client,
                            img_info["data"],
                            img_info["mime_type"],
                            catalog_prompt,
                            fields,
                        )
                        pil_image = Image.open(io.BytesIO(img_info["data"]))
                        img_w, img_h = pil_image.size
                        draw_regions = build_draw_regions(
                            result["data"], result["bboxes"], img_w, img_h, fields
                        )
                        annotated = draw_bboxes(
                            pil_image, draw_regions, label_colors, fields
                        )
                        return img_name, {
                            "data": result["data"],
                            "bboxes": result["bboxes"],
                            "draw_regions": draw_regions,
                            "annotated_image": annotated,
                            "original_image": pil_image,
                            "code_blocks": result["code_blocks"],
                            "response_text": result["response_text"],
                            "bbox_response_text": result.get("bbox_response_text", ""),
                            "tokens_in": result["tokens_in"],
                            "tokens_out": result["tokens_out"],
                            "elapsed": result["elapsed"],
                        }
                    except Exception as e:
                        pil_image = Image.open(io.BytesIO(img_info["data"]))
                        return img_name, {
                            "data": [], "bboxes": [], "draw_regions": [],
                            "annotated_image": None, "original_image": pil_image,
                            "code_blocks": [], "response_text": "",
                            "bbox_response_text": "", "tokens_in": 0,
                            "tokens_out": 0, "elapsed": 0, "error": str(e),
                        }

                cat_done = 0
                with ThreadPoolExecutor(
                    max_workers=min(3, len(catalog_images))
                ) as executor:
                    cat_futures = {
                        executor.submit(process_catalog_image, img): img["name"]
                        for img in catalog_images
                    }
                    for future in as_completed(cat_futures):
                        fname = cat_futures[future]
                        try:
                            name, cat_res = future.result()
                            extraction_results[name] = cat_res
                            for item in cat_res.get("data", []):
                                catalog_data_merged.update(
                                    {k: v for k, v in item.items() if v and v != "-"}
                                )
                        except Exception as e:
                            st.warning(f"{prod_prefix}カタログ {fname} の抽出でエラー: {e}")
                        cat_done += 1
                        progress.progress(
                            prod_base + (0.2 + cat_done / len(catalog_images) * 0.2) * prod_range,
                            text=f"{prod_prefix}カタログ抽出 [{cat_done}/{len(catalog_images)}] {fname}",
                        )

                progress.progress(
                    prod_base + 0.4 * prod_range,
                    text=f"{prod_prefix}カタログ抽出完了",
                )

            # ============================================================
            # Phase 2: 残りの技術資料を3並列で抽出
            # ============================================================
            if other_images:
                use_verification = bool(catalog_data_merged)

                def process_doc_image(img_info: dict) -> tuple[str, dict]:
                    img_name = img_info["name"]
                    image_data = img_info["data"]
                    mime_type = img_info["mime_type"]
                    doc_type = classification_results.get(
                        img_name, {}
                    ).get("category", "")

                    if use_verification:
                        prompt = build_verification_prompt(
                            catalog_data_merged, fields, doc_type
                        )
                    else:
                        # 品番ヒントがあればフォールバックでもアンカーとして使用
                        hinban_hint = ""
                        for f in fields:
                            if f["name"] == "品番" and f.get("hint"):
                                hinban_hint = f["hint"]
                                break
                        prompt = build_extraction_prompt(fields, target_hinban=hinban_hint)

                    try:
                        result = extract_from_image_bytes(
                            client, image_data, mime_type, prompt, fields,
                        )
                        pil_image = Image.open(io.BytesIO(image_data))
                        img_w, img_h = pil_image.size
                        draw_regions = build_draw_regions(
                            result["data"], result["bboxes"], img_w, img_h, fields
                        )
                        annotated = draw_bboxes(
                            pil_image, draw_regions, label_colors, fields
                        )
                        return img_name, {
                            "data": result["data"],
                            "bboxes": result["bboxes"],
                            "draw_regions": draw_regions,
                            "annotated_image": annotated,
                            "original_image": pil_image,
                            "code_blocks": result["code_blocks"],
                            "response_text": result["response_text"],
                            "bbox_response_text": result.get("bbox_response_text", ""),
                            "tokens_in": result["tokens_in"],
                            "tokens_out": result["tokens_out"],
                            "elapsed": result["elapsed"],
                        }
                    except Exception as e:
                        pil_image = Image.open(io.BytesIO(image_data))
                        return img_name, {
                            "data": [], "bboxes": [], "draw_regions": [],
                            "annotated_image": None, "original_image": pil_image,
                            "code_blocks": [], "response_text": "",
                            "bbox_response_text": "", "tokens_in": 0,
                            "tokens_out": 0, "elapsed": 0, "error": str(e),
                        }

                logger.info("  Phase 2: 技術資料 %d件を%s中（3並列）...",
                            len(other_images), "検証" if use_verification else "抽出")
                status_container.info(
                    f"{prod_prefix}技術資料 {len(other_images)} 件を"
                    f"{'検証' if use_verification else '抽出'}中（3並列）..."
                )
                doc_done = 0
                with ThreadPoolExecutor(
                    max_workers=min(3, len(other_images))
                ) as executor:
                    doc_futures = {
                        executor.submit(process_doc_image, img): img["name"]
                        for img in other_images
                    }
                    for future in as_completed(doc_futures):
                        fname = doc_futures[future]
                        try:
                            name, ext_res = future.result()
                            extraction_results[name] = ext_res
                        except Exception as e:
                            st.warning(f"{prod_prefix}{fname} の処理中にエラー: {e}")
                        doc_done += 1
                        progress.progress(
                            prod_base + (0.4 + doc_done / len(other_images) * 0.5) * prod_range,
                            text=f"{prod_prefix}抽出 [{doc_done}/{len(other_images)}] {fname}",
                        )

            # --- Phase 3: エビデンス画像をディレクトリに保存 ---
            logger.info("  Phase 3: エビデンス画像を保存中...")
            status_container.info(f"{prod_prefix}エビデンス画像を保存中...")
            try:
                evidence_dir = os.path.join(OUTPUT_EVIDENCE_DIR, timestamp)
                os.makedirs(evidence_dir, exist_ok=True)
                used_evidence_names: dict[str, int] = {}  # ファイル名衝突防止
                for img_name, result in extraction_results.items():
                    cls_info = classification_results.get(img_name, {})
                    doc_type = cls_info.get("category", "不明")
                    hinban = ""
                    for item in result.get("data", []):
                        for f in fields:
                            v = item.get(f["name"], "")
                            if v and v != "-":
                                hinban = v
                                break
                        if hinban:
                            break
                    base = hinban or os.path.splitext(img_name)[0]

                    # PDFのページ番号を保持（複数ページの上書き防止）
                    page_suffix = ""
                    if hinban:
                        # hinban取得時はimg_nameからページ番号を抽出して付与
                        page_match = re.search(r"_p(\d+)", img_name)
                        if page_match:
                            page_suffix = f"_p{page_match.group(1)}"

                    safe_name = f"{base}_{doc_type}{page_suffix}"

                    # それでも重複する場合は連番付与
                    if safe_name in used_evidence_names:
                        used_evidence_names[safe_name] += 1
                        safe_name = f"{safe_name}_{used_evidence_names[safe_name]}"
                    else:
                        used_evidence_names[safe_name] = 0

                    # BBox描画済み画像: draw_regionsがある場合のみ保存
                    if result.get("draw_regions") and result.get("annotated_image"):
                        result["annotated_image"].save(
                            os.path.join(evidence_dir, f"{safe_name}_BBox.png")
                        )
                    if result.get("original_image"):
                        result["original_image"].save(
                            os.path.join(evidence_dir, f"{safe_name}_原本.png")
                        )
                evidence_count = len(used_evidence_names)
                bbox_count = sum(
                    1 for r in extraction_results.values() if r.get("draw_regions")
                )
                logger.info(
                    "  エビデンス保存完了: %d画像 (BBox付き=%d, 原本のみ=%d)",
                    evidence_count, bbox_count, evidence_count - bbox_count,
                )
            except Exception as e:
                st.warning(f"{prod_prefix}エビデンス画像の保存でエラー: {e}")

            # --- Phase 3a: 読み取り結果Excelをディレクトリに保存 ---
            status_container.info(f"{prod_prefix}読み取り結果Excelを出力中...")
            try:
                os.makedirs(OUTPUT_EXTRACTION_DIR, exist_ok=True)
                suffix = f"_{product_id}" if product_id != "_default" and num_products > 1 else ""
                extraction_excel_path = os.path.join(
                    OUTPUT_EXTRACTION_DIR, f"読み取り結果{suffix}_{timestamp}.xlsx"
                )
                export_extraction_excel(
                    extraction_results, classification_results, fields, extraction_excel_path
                )
            except Exception as e:
                st.warning(f"{prod_prefix}読み取り結果Excelの出力でエラー: {e}")

            # --- Phase 3b: 正誤判定Excelをディレクトリに保存 ---
            doc_comparison: list = []
            detected_hinban = ""
            comparison_excel_bytes = None
            comparison_excel_filename = ""

            if all_correct and classification_results:
                status_container.info(f"{prod_prefix}正解データと比較中...")
                try:
                    detected_product = _detect_product_number(
                        extraction_results, fields, all_correct
                    )

                    # 品番検出失敗時、グルーピングの product_id をフォールバック
                    if (not detected_product or detected_product not in all_correct) \
                            and product_id != "_default" and product_id in all_correct:
                        logger.info(
                            "  品番自動検出失敗 → グルーピングキー '%s' をフォールバック使用",
                            product_id,
                        )
                        detected_product = product_id

                    if detected_product and detected_product in all_correct:
                        non_catalog_results = {
                            name: res for name, res in extraction_results.items()
                            if classification_results.get(name, {}).get("category") != "カタログ"
                        }

                        doc_comparison = compare_by_doc_type(
                            non_catalog_results,
                            classification_results,
                            all_correct[detected_product],
                            fields,
                            catalog_data=catalog_data_merged or None,
                            client=client,
                        )

                        # Gemini意味比較で mismatch を再判定
                        doc_comparison = refine_with_semantic_judge(
                            client,
                            doc_comparison,
                            [f["name"] for f in fields],
                        )

                        # 差分コメントを生成（完全一致以外の項目）
                        doc_comparison = generate_diff_comments(
                            client,
                            doc_comparison,
                            [f["name"] for f in fields],
                        )

                        detected_hinban = detected_product
                except Exception as e:
                    st.warning(f"{prod_prefix}正誤判定でエラー: {e}")

            if doc_comparison:
                if not detected_hinban:
                    detected_hinban = product_id if product_id != "_default" else "unknown"
                os.makedirs(OUTPUT_COMPARISON_DIR, exist_ok=True)
                comparison_filename = f"正誤結果_{detected_hinban}_{timestamp}.xlsx"
                comparison_path = os.path.join(OUTPUT_COMPARISON_DIR, comparison_filename)
                comparison_excel_bytes = export_comparison_excel(
                    doc_comparison, fields, comparison_path,
                    catalog_data=catalog_data_merged or None,
                )
                comparison_excel_filename = comparison_filename
                all_comparisons[detected_hinban] = doc_comparison

            # 品番ごとの結果を保存
            all_product_results[product_id] = {
                "classification": classification_results,
                "results": extraction_results,
                "uploaded_images": uploaded_image_data,
                "catalog_data": catalog_data_merged,
                "doc_type_comparison": doc_comparison,
                "comparison_excel_bytes": comparison_excel_bytes,
                "comparison_excel_filename": comparison_excel_filename,
                "detected_hinban": detected_hinban,
                "fields": fields,  # この品番で使った項目リスト
            }

            progress.progress(
                (prod_idx + 1) / num_products,
                text=f"{prod_prefix}処理完了",
            )

        # session_state に保存
        st.session_state.product_results = all_product_results

        # 複数品番の統合Excel出力
        if len(all_comparisons) > 1:
            try:
                os.makedirs(OUTPUT_COMPARISON_DIR, exist_ok=True)
                multi_filename = f"正誤結果_一括_{timestamp}.xlsx"
                multi_path = os.path.join(OUTPUT_COMPARISON_DIR, multi_filename)
                # 品番ごとの fields マップとカタログデータを構築
                comparison_fields_map = {}
                comparison_catalog_map = {}
                for hinban in all_comparisons:
                    comparison_fields_map[hinban] = product_fields_map.get(
                        hinban, default_fields
                    )
                    # 品番に対応する品番結果からカタログデータを取得
                    for _pid, pdata in all_product_results.items():
                        if pdata.get("detected_hinban") == hinban:
                            comparison_catalog_map[hinban] = pdata.get("catalog_data") or {}
                            break
                multi_bytes = export_multi_product_comparison_excel(
                    all_comparisons, comparison_fields_map, multi_path,
                    all_catalog_data=comparison_catalog_map or None,
                )
                st.session_state.multi_comparison_excel_bytes = multi_bytes
                st.session_state.multi_comparison_excel_filename = multi_filename
            except Exception as e:
                st.warning(f"統合Excel出力でエラー: {e}")
        else:
            st.session_state.multi_comparison_excel_bytes = None
            st.session_state.multi_comparison_excel_filename = ""

        progress.progress(1.0, text="全ての処理が完了しました")
        status_container.success("抽出完了")
        st.rerun()

    # --- 結果表示 ---
    if not st.session_state.product_results:
        st.info("サイドバーで項目リストと画像をアップロードし「抽出実行」をクリックしてください。")
        return

    product_fields_map = st.session_state.product_fields
    default_fields = st.session_state.default_fields
    label_colors = st.session_state.label_colors
    product_results = st.session_state.product_results
    product_ids = list(product_results.keys())

    # 複数品番統合Excelダウンロード
    if st.session_state.multi_comparison_excel_bytes:
        st.download_button(
            label="全品番の正誤結果Excelをダウンロード",
            data=st.session_state.multi_comparison_excel_bytes,
            file_name=st.session_state.multi_comparison_excel_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    if len(product_ids) == 1:
        pid = product_ids[0]
        pfields = product_fields_map.get(pid, default_fields)
        # 結果に保存された fields があればそちらを優先
        if not pfields and product_results[pid].get("fields"):
            pfields = product_results[pid]["fields"]
        plabel_colors = assign_colors(pfields) if pfields else label_colors
        _render_product_results(pid, product_results[pid], pfields, plabel_colors)
    else:
        # 品番が複数の場合: 品番ごとのタブを最上位に追加
        display_names = [
            pid if pid != "_default" else "未分類" for pid in product_ids
        ]
        product_tabs = st.tabs(display_names)
        for tab, pid in zip(product_tabs, product_ids):
            with tab:
                pfields = product_fields_map.get(pid, default_fields)
                if not pfields and product_results[pid].get("fields"):
                    pfields = product_results[pid]["fields"]
                plabel_colors = assign_colors(pfields) if pfields else label_colors
                _render_product_results(pid, product_results[pid], pfields, plabel_colors)


def _render_product_results(
    product_id: str,
    prod_data: dict,
    fields: list[dict],
    label_colors: dict[str, str],
):
    """1品番分の結果を表示する。"""
    results = prod_data.get("results", {})
    classification = prod_data.get("classification", {})

    if not results:
        st.info("結果がありません。")
        return

    processed_images = list(results.keys())

    # 画像分類結果
    if classification:
        render_classification_results(classification)

        # 手動修正UI（confidence=low の場合）
        low_conf = [
            name
            for name, info in classification.items()
            if info.get("confidence") == "low"
        ]
        if low_conf:
            st.markdown("---")
            st.markdown("**分類の確信度が低い画像があります。必要に応じて修正してください:**")
            for img_name in low_conf:
                new_cat = st.selectbox(
                    f"{img_name} のカテゴリ",
                    IMAGE_CATEGORIES,
                    index=IMAGE_CATEGORIES.index(
                        classification[img_name].get("category", "取扱説明書")
                    ),
                    key=f"cat_{product_id}_{img_name}",
                )
                if new_cat != classification[img_name].get("category"):
                    classification[img_name]["category"] = new_cat
                    classification[img_name]["confidence"] = "manual"

        st.divider()

    # 画像別タブ
    if processed_images:
        img_tabs = st.tabs(processed_images)
        for tab, img_name in zip(img_tabs, processed_images):
            with tab:
                result = results[img_name]
                render_image_result(img_name, result, fields, label_colors)

    # 照合マトリクス
    st.divider()
    st.subheader("照合マトリクス")
    render_comparison_matrix(processed_images, results, fields)

    # 技術文書タイプ別比較（マスターExcel経由）
    doc_type_comparison = prod_data.get("doc_type_comparison", [])
    if doc_type_comparison:
        st.divider()
        st.subheader("正誤結果（技術文書別）")

        # ダウンロードボタン
        comp_bytes = prod_data.get("comparison_excel_bytes")
        comp_filename = prod_data.get("comparison_excel_filename", "")
        if comp_bytes:
            st.download_button(
                label="正誤結果Excelをダウンロード",
                data=comp_bytes,
                file_name=comp_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key=f"dl_{product_id}",
            )

        # 比較結果テーブル表示
        field_names = [f["name"] for f in fields]
        status_icon = {"match": "OK", "mismatch": "NG", "na": "-"}
        table_rows = []
        for row in doc_type_comparison:
            table_row = {"品番": row["品番"], "技術文書": row["技術文書"]}
            for name in field_names:
                item = row["items"].get(name, {})
                s = item.get("status", "na")
                table_row[f"{name}_正解"] = item.get("correct", "")
                table_row[f"{name}_抽出"] = item.get("extracted", "")
                table_row[f"{name}_判定"] = status_icon.get(s, "-")
                table_row[f"{name}_差分"] = item.get("diff_comment", "")
            table_rows.append(table_row)

        if table_rows:
            df = pd.DataFrame(table_rows)

            def color_verdict(col):
                if not col.name.endswith("_判定"):
                    return [""] * len(col)
                return [
                    "background-color: #D5F5E3; color: #1E8449"
                    if v == "OK"
                    else "background-color: #FADBD8; color: #C0392B"
                    if v == "NG"
                    else "color: #AAB7B8"
                    for v in col
                ]

            styled = df.style.apply(color_verdict, axis=0)
            st.dataframe(styled, use_container_width=True, hide_index=True)


def _remap_product_groups(
    product_groups: dict[str, list[dict]],
    product_fields_map: dict[str, list[dict]],
) -> dict[str, list[dict]]:
    """ZIPフォルダ名が品番と一致しないグループをファイル名から再マッピングする。

    ZIPのフォルダが連番（2, 3, 4...）等の場合、中のファイル名から品番を推定し
    項目リストの品番と紐づける。
    """
    known_products = set(product_fields_map.keys())
    if not known_products:
        return product_groups

    remapped: dict[str, list[dict]] = {}

    for group_key, images in product_groups.items():
        # 既に項目リストの品番と一致している → そのまま
        if group_key in known_products:
            remapped.setdefault(group_key, []).extend(images)
            logger.info("  グループ '%s' → 品番一致（そのまま）", group_key)
            continue

        # ファイル名から品番を推定
        detected_products: dict[str, int] = {}
        for img in images:
            fname = img["name"]
            # 方法1: extract_product_from_filename で推定
            product = extract_product_from_filename(fname)
            if product and product in known_products:
                detected_products[product] = detected_products.get(product, 0) + 1
                continue
            # 方法2: ファイル名に既知の品番が部分一致するか
            for known in known_products:
                if known in fname:
                    detected_products[known] = detected_products.get(known, 0) + 1
                    break

        if detected_products:
            # 最も多く検出された品番に割り当て
            best_product = max(detected_products, key=detected_products.get)
            remapped.setdefault(best_product, []).extend(images)
            logger.info("  グループ '%s' → ファイル名から '%s' に再マッピング（%d/%d件一致）",
                         group_key, best_product,
                         detected_products[best_product], len(images))
        else:
            # ファイル名からも推定できない → 元のグループキーで保持
            remapped.setdefault(group_key, []).extend(images)
            logger.info("  グループ '%s' → 品番推定不可（そのまま）", group_key)

    return remapped


def _detect_product_number(
    results: dict, fields: list[dict], all_correct: dict
) -> str | None:
    """抽出結果から品番を自動検出し、正解データの品番とマッチする。

    品番フィールドを優先的にチェックし、正規化マッチング（スペース・括弧除去）で
    表記揺れにも対応する。
    """
    known_products = set(all_correct.keys())
    if not known_products:
        return None

    def _normalize(s: str) -> str:
        """品番の正規化: スペース・括弧除去、大文字化"""
        s = re.sub(r"[\s\u3000]+", "", s)
        s = re.sub(r"[()（）]", "", s)
        return s.upper()

    # 正規化品番 → 元品番 のマップ
    normalized_map: dict[str, str] = {_normalize(p): p for p in known_products}

    # 正解データの「品番」フィールドの値のみを子品番として収集
    child_to_parent: dict[str, str] = {}
    for parent, docs in all_correct.items():
        for _doc_type, values in docs.items():
            hinban_val = values.get("品番", "")
            if hinban_val and hinban_val.strip() and hinban_val.strip() != "-":
                child_to_parent[_normalize(hinban_val.strip())] = parent

    # Step 1: 品番フィールドを優先チェック
    for _img_name, result in results.items():
        for item in result.get("data", []):
            hinban = item.get("品番", "")
            if not hinban or str(hinban).strip() == "-":
                continue
            v = str(hinban).strip()
            norm_v = _normalize(v)
            if v in known_products:
                return v
            if norm_v in normalized_map:
                return normalized_map[norm_v]
            if norm_v in child_to_parent:
                return child_to_parent[norm_v]

    # Step 2: 全フィールドで親品番のみチェック（汎用値でのマッチを防ぐ）
    for _img_name, result in results.items():
        for item in result.get("data", []):
            for _key, val in item.items():
                v = str(val).strip() if val else ""
                if not v or v == "-":
                    continue
                if v in known_products:
                    return v
                norm_v = _normalize(v)
                if norm_v in normalized_map:
                    return normalized_map[norm_v]

    return None


def _find_correct_data_file() -> str | None:
    """correct_data/ ディレクトリからExcelファイルを探す。"""
    if not os.path.isdir(CORRECT_DATA_DIR):
        return None
    for fname in os.listdir(CORRECT_DATA_DIR):
        if fname.endswith((".xlsx", ".xls")) and not fname.startswith("~$"):
            return os.path.join(CORRECT_DATA_DIR, fname)
    return None


if __name__ == "__main__":
    main()
