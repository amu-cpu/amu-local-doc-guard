from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from uuid import uuid4
from xml.etree import ElementTree

import fitz
from PIL import Image, ImageDraw, ImageFilter, ImageOps


OUTPUT_ROOT = Path("output")
RULES_PATH = Path(__file__).resolve().parent.parent / "rules_sensitive.json"
SAFE_TEMP_ROOT = Path(os.environ.get("DESENSITIZER_TEMP", r"F:\Tools\Temp\desensitizer"))
ALLOWED_EXTENSIONS = {".pdf", ".docx", ".doc"}
ALLOWED_IMAGE_METHODS = {"black_box", "mosaic", "blur", "gray_mosaic"}
DETAIL_IMAGE_DPI = 200
IMAGE_ZOOM = DETAIL_IMAGE_DPI / 72
OUTPUT_PREFIX = "脱敏后_"


SENSITIVE_PATTERNS = [
    ("phone", re.compile(r"(?<!\d)1[3-9]\d{9}(?!\d)"), "regex:phone"),
    (
        "id_card",
        re.compile(
            r"(?<![0-9Xx])\d{6}(?:18|19|20)\d{2}"
            r"(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx](?![0-9Xx])"
        ),
        "regex:id_card",
    ),
    ("email", re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"), "regex:email"),
    ("bank_card", re.compile(r"(?<!\d)(?:\d[ -]?){16,19}(?!\d)"), "regex:bank_card"),
    (
        "unified_social_credit_code",
        re.compile(r"(?<![0-9A-Z])[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}(?![0-9A-Z])"),
        "regex:unified_social_credit_code",
    ),
]

PROJECT_CODE_RE = re.compile(r"(?<!\d)\d{4}-\d{6}-\d{2}-\d{2}-\d{6}(?!\d)")
AMOUNT_RE = re.compile(r"(?:人民币|￥|¥)?\s*\d[\d,，]*(?:\.\d+)?\s*(?:万元|亿元|元|万|亿)")
PERCENT_RE = re.compile(r"(?<![\d.])\d+(?:\.\d+)?\s*(?:%|％)")
METRIC_RE = re.compile(r"(?<![\d.])\d+(?:\.\d+)?\s*(?:年|个月|月|倍)")
NUMBER_RE = re.compile(r"[-+]?\d[\d,，]*(?:\.\d+)?\s*(?:万元|亿元|元|万|亿|%|％|年|个月|月|倍)?")
COMPANY_RE = re.compile(
    r"[\u4e00-\u9fffA-Za-z0-9（）()·\-]{2,60}"
    r"(?:股份有限公司|有限责任公司|有限公司|集团|合作社|研究院|中心|工厂|厂)"
)
PROJECT_NAME_RE = re.compile(r"[\u4e00-\u9fffA-Za-z0-9（）()·\- ]{4,90}项目")
GENERIC_TITLE_WORDS = ["可行性研究报告", "商业计划书", "可研报告", "项目建议书", "融资计划书"]


class ProcessingError(Exception):
    pass


def load_rules() -> dict:
    with RULES_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


RULES = load_rules()
SUBJECT_KEYWORDS = RULES.get("subject_keywords", [])
PROJECT_KEYWORDS = RULES.get("project_keywords", [])
FINANCIAL_KEYWORDS = RULES.get("financial_keywords", [])
BUSINESS_KEYWORDS = RULES.get("business_keywords", [])
TABLE_VALUE_LABELS = RULES.get("table_value_labels", [])
FINANCIAL_CONTEXT_WORDS = RULES.get("financial_context_words", [])
BLACK_BOX_IMAGE_TYPES = set(RULES.get("black_box_image_types", []))
ALL_LABELS = list(dict.fromkeys(SUBJECT_KEYWORDS + PROJECT_KEYWORDS + FINANCIAL_KEYWORDS + BUSINESS_KEYWORDS + TABLE_VALUE_LABELS))


def process_uploaded_file(
    original_name: str,
    save_func,
    output_root: Path = OUTPUT_ROOT,
    image_method: str | None = None,
    expected_pages: int | None = None,
) -> dict:
    original_name = sanitize_filename(original_name)
    suffix = Path(original_name).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise ProcessingError("仅支持 PDF、DOCX、DOC 文件。")

    safe_stem = sanitize_filename(Path(original_name).stem) or "document"
    output_stem = prefixed_output_stem(safe_stem)
    job_id = f"{safe_stem}_{uuid4().hex[:8]}"
    job_dir = (output_root / job_id).resolve()
    input_dir = job_dir / "input"
    work_dir = job_dir / "work"
    input_dir.mkdir(parents=True, exist_ok=False)
    work_dir.mkdir(parents=True, exist_ok=True)

    input_path = input_dir / original_name
    save_func(input_path)
    source_pdf = convert_to_source_pdf(input_path, work_dir)
    converted_page_count = get_pdf_page_count(source_pdf)
    docx_declared_page_count = read_docx_declared_page_count(input_path) if suffix == ".docx" else None
    page_count_check = build_page_count_check(
        suffix=suffix,
        expected_pages=expected_pages,
        docx_declared_page_count=docx_declared_page_count,
        converted_page_count=converted_page_count,
    )

    meta = {
        "job_id": job_id,
        "original_name": original_name,
        "safe_stem": safe_stem,
        "output_stem": output_stem,
        "source_pdf": str(source_pdf.relative_to(job_dir)),
        "image_method": normalize_image_method(image_method or RULES.get("default_image_method", "mosaic")),
        "source_type": suffix.lstrip(".").upper(),
        "conversion_method": conversion_method_for_suffix(suffix),
        "converted_page_count": converted_page_count,
        "expected_page_count": normalize_expected_pages(expected_pages),
        "docx_declared_page_count": docx_declared_page_count,
        "page_count_check": page_count_check,
    }
    write_json(job_dir / "job_meta.json", meta)
    return reprocess_job(job_dir, [])


def add_manual_redaction(
    job_id: str,
    page_number: int,
    image_rect: tuple[float, float, float, float],
    image_size: tuple[float, float],
    method: str,
    output_root: Path = OUTPUT_ROOT,
) -> dict:
    job_dir = resolve_job_dir(job_id, output_root)
    report = read_report(job_dir)

    page_info = next((page for page in report.get("pages", []) if page["page"] == page_number), None)
    if not page_info:
        raise ProcessingError("没有找到对应页码。")

    x, y, width, height = image_rect
    image_width, image_height = image_size
    if width < 4 or height < 4:
        raise ProcessingError("框选区域太小。")
    if image_width <= 0 or image_height <= 0:
        raise ProcessingError("图片尺寸无效。")

    x0 = max(0.0, x) / image_width * float(page_info["pdf_width"])
    y0 = max(0.0, y) / image_height * float(page_info["pdf_height"])
    x1 = min(image_width, x + width) / image_width * float(page_info["pdf_width"])
    y1 = min(image_height, y + height) / image_height * float(page_info["pdf_height"])
    if x1 <= x0 or y1 <= y0:
        raise ProcessingError("框选区域无效。")

    manual_redactions = get_manual_redactions(report)
    manual_item = {
        "id": f"manual-{uuid4().hex[:10]}",
        "job_id": job_id,
        "sensitive_type": "manual",
        "page": page_number,
        "page_index": page_number - 1,
        "rect": round_rect([x0, y0, x1, y1]),
        "preview_rect": {
            "x": round(float(x), 2),
            "y": round(float(y), 2),
            "width": round(float(width), 2),
            "height": round(float(height), 2),
        },
        "preview_image_size": {
            "width": round(float(image_width), 2),
            "height": round(float(image_height), 2),
        },
        "method": normalize_image_method(method),
        "mask_type": normalize_image_method(method),
        "pdf_method": "mosaic_image_pdf",
        "manual": True,
        "source": "manual",
        "rule": "manual:preview_selection",
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
    }
    manual_redactions.append(manual_item)
    update_report_manual_redactions(job_dir, report, manual_redactions)
    apply_image_masks(job_dir, [page_info], [manual_item])
    return read_report(job_dir)


def undo_last_manual_redaction(job_id: str, page_number: int | None = None, output_root: Path = OUTPUT_ROOT) -> dict:
    job_dir = resolve_job_dir(job_id, output_root)
    report = read_report(job_dir)
    manual_redactions = get_manual_redactions(report)
    if page_number:
        for index in range(len(manual_redactions) - 1, -1, -1):
            if int(manual_redactions[index].get("page", 0)) == int(page_number):
                manual_redactions.pop(index)
                break
    elif manual_redactions:
        page_number = int(manual_redactions[-1].get("page", 0))
        manual_redactions.pop()
    update_report_manual_redactions(job_dir, report, manual_redactions)
    if page_number:
        rerender_report_page(job_dir, report, int(page_number))
    return read_report(job_dir)


def clear_manual_redactions(job_id: str, page_number: int | None = None, output_root: Path = OUTPUT_ROOT) -> dict:
    job_dir = resolve_job_dir(job_id, output_root)
    report = read_report(job_dir)
    if page_number:
        manual_redactions = [
            item for item in get_manual_redactions(report) if int(item.get("page", 0)) != int(page_number)
        ]
    else:
        manual_redactions = []
    update_report_manual_redactions(job_dir, report, manual_redactions)
    if page_number:
        rerender_report_page(job_dir, report, int(page_number))
    return read_report(job_dir)


def reprocess_existing_job(
    job_id: str,
    image_method: str | None = None,
    output_root: Path = OUTPUT_ROOT,
) -> dict:
    job_dir = resolve_job_dir(job_id, output_root)
    report = read_report(job_dir)
    return reprocess_job(job_dir, get_manual_redactions(report), image_method=image_method)


def reprocess_job(
    job_dir: Path,
    manual_redactions: list[dict],
    image_method: str | None = None,
) -> dict:
    meta = read_json(job_dir / "job_meta.json")
    source_pdf = job_dir / meta["source_pdf"]
    output_stem = meta.get("output_stem") or prefixed_output_stem(meta["safe_stem"])
    selected_image_method = normalize_image_method(image_method or meta.get("image_method") or RULES.get("default_image_method", "mosaic"))
    meta["output_stem"] = output_stem
    meta["image_method"] = selected_image_method
    write_json(job_dir / "job_meta.json", meta)

    pages_dir = job_dir / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)

    output_pdf = job_dir / f"{output_stem}.pdf"
    report_path = job_dir / f"{output_stem}_处理报告.json"
    zip_path = job_dir / f"{output_stem}_分页详情图.zip"

    auto_redactions, seeds = detect_auto_redactions(source_pdf, selected_image_method)
    redactions = dedupe_redactions(auto_redactions + manual_redactions)
    pages = render_pdf_pages(source_pdf, pages_dir, output_stem)
    apply_image_masks(job_dir, pages, redactions)
    manual_count = sum(1 for item in redactions if item.get("manual"))
    auto_count = sum(1 for item in redactions if not item.get("manual"))

    report = {
        "job_id": meta["job_id"],
        "original_file": meta["original_name"],
        "created_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "local_only": True,
        "outputs": {
            "pdf": output_pdf.name,
            "zip": zip_path.name,
            "report": report_path.name,
            "images_dir": pages_dir.name,
        },
        "summary": {
            "page_count": len(pages),
            "auto_redactions": auto_count,
            "manual_redactions": manual_count,
            "detail_image_method": selected_image_method,
            "pdf_method": "mosaic_image_pdf",
            "sensitive_seed_count": len(seeds),
            "ocr_enabled": False,
        },
        "pdf_output_type": "mosaic_image_pdf",
        "black_pdf_enabled": False,
        "manual_masks_count": manual_count,
        "auto_masks_count": auto_count,
        "total_pages": len(pages),
        "source_file_type": meta.get("source_type"),
        "convert_method": meta.get("conversion_method"),
        "page_count_warning": (meta.get("page_count_check") or {}).get("page_count_warning", ""),
        "output_files": {
            "脱敏 PDF": output_pdf.name,
            "分页详情图 ZIP": zip_path.name,
            "处理报告 JSON": report_path.name,
        },
        "page_count_check": meta.get("page_count_check", {}),
        "conversion": {
            "source_type": meta.get("source_type"),
            "conversion_method": meta.get("conversion_method"),
            "converted_page_count": meta.get("converted_page_count"),
            "expected_page_count": meta.get("expected_page_count"),
            "docx_declared_page_count": meta.get("docx_declared_page_count"),
        },
        "sensitive_seeds": seeds,
        "manual_masks": [manual_mask_record(item) for item in redactions if item.get("manual")],
        "pages": pages,
        "redactions": redactions,
    }
    write_json(report_path, report)
    return report


def export_job_output(
    job_id: str,
    kind: str,
    output_root: Path = OUTPUT_ROOT,
    progress_callback=None,
) -> tuple[dict, str]:
    job_dir = resolve_job_dir(job_id, output_root)
    report = read_report(job_dir)
    outputs = report.get("outputs", {})
    if kind not in {"pdf", "zip"}:
        raise ProcessingError("导出类型无效。")

    report = regenerate_masked_pages(job_dir, report, progress_callback=progress_callback)
    if kind == "pdf":
        progress_callback and progress_callback("正在生成脱敏 PDF", None, report["summary"]["page_count"])
        output_name = outputs.get("pdf") or f"{report_output_stem(job_dir)}.pdf"
        create_mosaic_pdf_from_images(job_dir, report["pages"], job_dir / output_name)
    else:
        progress_callback and progress_callback("正在打包 ZIP", None, report["summary"]["page_count"])
        output_name = outputs.get("zip") or f"{report_output_stem(job_dir)}_分页详情图.zip"
        create_images_zip(report["pages"], job_dir / output_name)

    write_report(job_dir, report)
    return report, output_name


def regenerate_masked_pages(job_dir: Path, report: dict, progress_callback=None) -> dict:
    meta = read_json(job_dir / "job_meta.json")
    source_pdf = job_dir / meta["source_pdf"]
    output_stem = meta.get("output_stem") or prefixed_output_stem(meta["safe_stem"])
    pages_dir = job_dir / "pages"
    progress_callback and progress_callback("正在生成分页详情图", 0, int(report.get("summary", {}).get("page_count") or 0))
    pages = render_pdf_pages(source_pdf, pages_dir, output_stem, progress_callback=progress_callback)
    report["pages"] = pages
    report.setdefault("summary", {})["page_count"] = len(pages)
    report["total_pages"] = len(pages)
    apply_image_masks(job_dir, pages, report.get("redactions", []))
    return report


def update_report_manual_redactions(job_dir: Path, report: dict, manual_redactions: list[dict]):
    auto_redactions = [item for item in report.get("redactions", []) if not item.get("manual")]
    redactions = dedupe_redactions(auto_redactions + manual_redactions)
    manual_count = sum(1 for item in redactions if item.get("manual"))
    auto_count = sum(1 for item in redactions if not item.get("manual"))
    report["redactions"] = redactions
    report["manual_masks"] = [manual_mask_record(item) for item in redactions if item.get("manual")]
    report["manual_masks_count"] = manual_count
    report["auto_masks_count"] = auto_count
    report["total_pages"] = int(report.get("summary", {}).get("page_count") or len(report.get("pages", [])))
    report.setdefault("summary", {})["manual_redactions"] = manual_count
    report.setdefault("summary", {})["auto_redactions"] = auto_count
    report.setdefault("summary", {})["pdf_method"] = "mosaic_image_pdf"
    report["pdf_output_type"] = "mosaic_image_pdf"
    report["black_pdf_enabled"] = False
    write_report(job_dir, report)


def rerender_report_page(job_dir: Path, report: dict, page_number: int):
    page_info = next((page for page in report.get("pages", []) if int(page.get("page", 0)) == page_number), None)
    if not page_info:
        return
    meta = read_json(job_dir / "job_meta.json")
    source_pdf = job_dir / meta["source_pdf"]
    render_single_pdf_page(source_pdf, job_dir / page_info["image"], page_number - 1)
    page_redactions = [item for item in report.get("redactions", []) if int(item.get("page", 0)) == page_number]
    apply_image_masks(job_dir, [page_info], page_redactions)


def detect_auto_redactions(pdf_path: Path, image_method: str) -> tuple[list[dict], list[dict]]:
    doc = fitz.open(pdf_path)
    pages: list[dict] = []
    redactions: list[dict] = []
    try:
        for page_index in range(doc.page_count):
            page = doc[page_index]
            pages.append(
                {
                    "page_number": page_index + 1,
                    "rect": page.rect,
                    "lines": list(iter_text_lines(page)),
                }
            )

        seeds = extract_sensitive_seeds(pages, image_method)
        for page_data in pages:
            redactions.extend(detect_page_redactions(page_data, seeds, image_method))
    finally:
        doc.close()
    return dedupe_redactions(redactions), seeds


def detect_page_redactions(page_data: dict, seeds: list[dict], image_method: str) -> list[dict]:
    redactions: list[dict] = []
    page_number = page_data["page_number"]
    page_rect = page_data["rect"]
    lines = page_data["lines"]

    for line in lines:
        text = line["text"]
        boxes = line["boxes"]
        if not text.strip():
            continue

        for sensitive_type, pattern, rule in SENSITIVE_PATTERNS:
            for match in pattern.finditer(text):
                value = match.group(0)
                if sensitive_type == "bank_card" and len(re.sub(r"\D", "", value)) < 16:
                    continue
                if sensitive_type == "bank_card" and is_inside_project_code(text, match.start(), match.end()):
                    continue
                add_span_redaction(redactions, page_number, boxes, match.start(), match.end(), sensitive_type, rule)

        for match in PROJECT_CODE_RE.finditer(text):
            add_span_redaction(redactions, page_number, boxes, match.start(), match.end(), "project_code", "regex:project_code")

        for seed in seeds:
            seed_text = seed["text"]
            if len(seed_text) < 2:
                continue
            for match in re.finditer(re.escape(seed_text), text, flags=re.IGNORECASE):
                add_span_redaction(
                    redactions,
                    page_number,
                    boxes,
                    match.start(),
                    match.end(),
                    seed["sensitive_type"],
                    f"seed:{seed['source']}",
                    method=seed["method"],
                )

        detect_label_values(redactions, page_number, line, lines, image_method)
        detect_conservative_financial_values(redactions, page_number, line, image_method)
        detect_header_footer_title(redactions, page_number, line, page_rect, seeds, image_method)

    return redactions


def detect_label_values(
    redactions: list[dict],
    page_number: int,
    line: dict,
    lines: list[dict],
    image_method: str,
):
    text = line["text"]
    boxes = line["boxes"]
    for label in ALL_LABELS:
        for match in re.finditer(re.escape(label), text, flags=re.IGNORECASE):
            category = category_for_label(label)
            sensitive_type = sensitive_type_for_category(category, label)
            method = method_for_type(sensitive_type, image_method)
            start = skip_separators(text, match.end())
            end = find_value_end(text, start)

            if end > start:
                if category == "financial":
                    segment = text[start:end]
                    for number_match in NUMBER_RE.finditer(segment):
                        if not has_digit(number_match.group(0)):
                            continue
                        add_span_redaction(
                            redactions,
                            page_number,
                            boxes,
                            start + number_match.start(),
                            start + number_match.end(),
                            sensitive_type,
                            f"financial:{label}",
                            method=method,
                        )
                else:
                    add_span_redaction(
                        redactions,
                        page_number,
                        boxes,
                        start,
                        end,
                        sensitive_type,
                        f"label:{label}",
                        method=method,
                    )

            if label in TABLE_VALUE_LABELS and should_search_adjacent_value(line, label):
                for value_line in find_adjacent_value_lines(lines, line):
                    add_rect_redaction(
                        redactions,
                        page_number,
                        value_line["rect"],
                        sensitive_type,
                        f"table_adjacent:{label}",
                        method=method,
                    )


def detect_conservative_financial_values(
    redactions: list[dict],
    page_number: int,
    line: dict,
    image_method: str,
):
    if not RULES.get("conservative_finance", True):
        return

    text = line["text"]
    boxes = line["boxes"]
    has_finance_context = any(word.lower() in text.lower() for word in FINANCIAL_CONTEXT_WORDS + FINANCIAL_KEYWORDS)
    method = method_for_type("financial_value", image_method)

    for pattern, rule in [
        (AMOUNT_RE, "finance_amount"),
        (PERCENT_RE, "finance_percent"),
        (METRIC_RE, "finance_metric"),
    ]:
        if pattern is not AMOUNT_RE and not has_finance_context:
            continue
        for match in pattern.finditer(text):
            if pattern is METRIC_RE and not has_finance_context:
                continue
            add_span_redaction(
                redactions,
                page_number,
                boxes,
                match.start(),
                match.end(),
                "financial_value",
                f"conservative:{rule}",
                method=method,
            )

    if has_finance_context:
        for match in NUMBER_RE.finditer(text):
            value = match.group(0)
            if not has_digit(value) or looks_like_year_only(value):
                continue
            add_span_redaction(
                redactions,
                page_number,
                boxes,
                match.start(),
                match.end(),
                "financial_value",
                "conservative:finance_context_number",
                method=method,
            )


def detect_header_footer_title(
    redactions: list[dict],
    page_number: int,
    line: dict,
    page_rect,
    seeds: list[dict],
    image_method: str,
):
    text = line["text"].strip()
    if not text:
        return

    rect = line["rect"]
    top_limit = page_rect.height * 0.10
    bottom_limit = page_rect.height * 0.90
    in_header_footer = rect.y1 <= top_limit or rect.y0 >= bottom_limit
    title_like = rect.y0 <= page_rect.height * 0.30 and line.get("font_size", 0) >= 14
    sensitive_title = title_like and ("项目" in text or any(word in text for word in ["有限公司", "集团"])) and any(
        word in text for word in GENERIC_TITLE_WORDS + ["项目"]
    )

    if not in_header_footer and not sensitive_title:
        return

    for seed in seeds:
        if seed["sensitive_type"] not in {"entity_seed", "project_seed", "address_seed", "label_value"}:
            continue
        for match in re.finditer(re.escape(seed["text"]), text, flags=re.IGNORECASE):
            add_span_redaction(
                redactions,
                page_number,
                line["boxes"],
                match.start(),
                match.end(),
                seed["sensitive_type"],
                "header_footer_title:seed",
                method=seed["method"],
            )

    for pattern, sensitive_type, source in [
        (COMPANY_RE, "entity_seed", "header_footer_title:company"),
        (PROJECT_NAME_RE, "project_seed", "header_footer_title:project"),
    ]:
        for match in pattern.finditer(text):
            value = clean_seed_text(match.group(0))
            if not value:
                continue
            start = match.start() + match.group(0).find(value)
            add_span_redaction(
                redactions,
                page_number,
                line["boxes"],
                max(match.start(), start),
                match.end(),
                sensitive_type,
                source,
                method=method_for_type(sensitive_type, image_method),
            )


def extract_sensitive_seeds(pages: list[dict], image_method: str) -> list[dict]:
    seeds: dict[str, dict] = {}
    scan_pages = pages[:5]

    def add_seed(text: str, sensitive_type: str, source: str):
        cleaned = clean_seed_text(text)
        if not cleaned or len(cleaned) < 2 or len(cleaned) > 120:
            return
        if cleaned in ALL_LABELS or cleaned in GENERIC_TITLE_WORDS:
            return
        if cleaned not in seeds:
            seeds[cleaned] = {
                "text": cleaned,
                "sensitive_type": sensitive_type,
                "source": source,
                "method": method_for_type(sensitive_type, image_method),
            }

    for page_data in scan_pages:
        lines = page_data["lines"]
        for line in lines:
            text = line["text"].strip()
            if not text:
                continue

            for match in COMPANY_RE.finditer(text):
                add_seed(match.group(0), "entity_seed", "company_suffix")
            for match in PROJECT_NAME_RE.finditer(text):
                add_seed(strip_generic_title_words(match.group(0)), "project_seed", "project_name")
            for match in PROJECT_CODE_RE.finditer(text):
                add_seed(match.group(0), "project_code", "project_code")

            for label in ALL_LABELS:
                for match in re.finditer(re.escape(label), text, flags=re.IGNORECASE):
                    category = category_for_label(label)
                    sensitive_type = sensitive_type_for_category(category, label)
                    start = skip_separators(text, match.end())
                    end = find_value_end(text, start)
                    if end > start:
                        add_seed(text[start:end], sensitive_type, f"label:{label}")
                    if label in TABLE_VALUE_LABELS and should_search_adjacent_value(line, label):
                        for value_line in find_adjacent_value_lines(lines, line):
                            add_seed(value_line["text"], sensitive_type, f"table_adjacent:{label}")

            has_finance_context = any(word.lower() in text.lower() for word in FINANCIAL_CONTEXT_WORDS + FINANCIAL_KEYWORDS)
            for match in AMOUNT_RE.finditer(text):
                add_seed(match.group(0), "financial_value", "amount")
            if has_finance_context:
                for pattern, source in [(PERCENT_RE, "percent"), (METRIC_RE, "metric")]:
                    for match in pattern.finditer(text):
                        add_seed(match.group(0), "financial_value", source)

    return sorted(seeds.values(), key=lambda item: (item["sensitive_type"], item["text"]))


def iter_text_lines(page):
    data = page.get_text("rawdict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            chars = []
            boxes = []
            sizes = []
            for span in line.get("spans", []):
                if span.get("size"):
                    sizes.append(float(span["size"]))
                for char in span.get("chars", []):
                    value = char.get("c", "")
                    if not value:
                        continue
                    chars.append(value)
                    boxes.append(fitz.Rect(char["bbox"]))
            text = "".join(chars)
            if text.strip() and boxes:
                rect = boxes[0]
                for box in boxes[1:]:
                    rect |= box
                yield {
                    "text": text,
                    "boxes": boxes,
                    "rect": rect,
                    "font_size": max(sizes) if sizes else rect.height,
                }


def find_adjacent_value_lines(lines: list[dict], label_line: dict) -> list[dict]:
    label_rect = label_line["rect"]
    label_center_y = (label_rect.y0 + label_rect.y1) / 2
    right_candidates = []
    for candidate in lines:
        if candidate is label_line or is_label_only(candidate["text"]):
            continue
        rect = candidate["rect"]
        if rect.x0 <= label_rect.x1 - 2:
            continue
        vertical_overlap = max(0, min(label_rect.y1, rect.y1) - max(label_rect.y0, rect.y0))
        same_row = vertical_overlap >= min(label_rect.height, rect.height) * 0.35
        close_center = abs(((rect.y0 + rect.y1) / 2) - label_center_y) <= max(8, label_rect.height * 0.9)
        if same_row or close_center:
            right_candidates.append(candidate)

    if right_candidates:
        right_candidates.sort(key=lambda item: (item["rect"].x0, item["rect"].y0))
        first = right_candidates[0]
        same_cell = [first]
        for candidate in right_candidates[1:]:
            if candidate["rect"].x0 - first["rect"].x1 > 160:
                break
            if candidate["rect"].x0 < first["rect"].x1 + 90:
                same_cell.append(candidate)
        return same_cell[:3]

    below_candidates = []
    for candidate in lines:
        if candidate is label_line or is_label_only(candidate["text"]):
            continue
        rect = candidate["rect"]
        vertical_gap = rect.y0 - label_rect.y1
        if 0 <= vertical_gap <= max(36, label_rect.height * 2.2):
            horizontal_overlap = max(0, min(label_rect.x1, rect.x1) - max(label_rect.x0, rect.x0))
            close_x = abs(rect.x0 - label_rect.x0) <= 80 or horizontal_overlap > 0
            if close_x:
                below_candidates.append(candidate)
    below_candidates.sort(key=lambda item: (item["rect"].y0, item["rect"].x0))
    return below_candidates[:2]


def should_search_adjacent_value(line: dict, label: str) -> bool:
    text = line["text"].strip()
    compact_text = compact_label_text(text)
    compact_label = compact_label_text(label)
    if compact_text == compact_label:
        return True
    index = text.lower().find(label.lower())
    if index < 0:
        return False
    after = text[index + len(label) :]
    after = re.sub(r"^[\s：:;；,，、\-\|/\\（）()【】\[\]]+", "", after)
    return len(after.strip()) <= 1


def add_span_redaction(
    redactions: list[dict],
    page_number: int,
    boxes: list,
    start: int,
    end: int,
    sensitive_type: str,
    rule: str,
    method: str | None = None,
):
    start, end = trim_span(boxes, start, end)
    if end <= start:
        return
    rect = boxes[start]
    for box in boxes[start + 1 : end]:
        rect |= box
    add_rect_redaction(redactions, page_number, rect, sensitive_type, rule, method=method)


def add_rect_redaction(
    redactions: list[dict],
    page_number: int,
    rect,
    sensitive_type: str,
    rule: str,
    method: str | None = None,
    manual: bool = False,
):
    expanded = fitz.Rect(rect.x0 - 1, rect.y0 - 1, rect.x1 + 1, rect.y1 + 1)
    if expanded.width < 2 or expanded.height < 2:
        return
    chosen_method = normalize_image_method(method or method_for_type(sensitive_type, RULES.get("default_image_method", "mosaic")))
    if not manual:
        chosen_method = visual_method_for_rect(sensitive_type, rule, expanded, chosen_method)
    redactions.append(
        {
            "id": f"{'manual' if manual else 'auto'}-{uuid4().hex[:10]}",
            "sensitive_type": sensitive_type,
            "page": page_number,
            "rect": round_rect([expanded.x0, expanded.y0, expanded.x1, expanded.y1]),
            "method": chosen_method,
            "pdf_method": "mosaic_image_pdf",
            "manual": manual,
            "rule": rule,
        }
    )


def trim_span(boxes: list, start: int, end: int) -> tuple[int, int]:
    start = max(0, min(start, len(boxes)))
    end = max(0, min(end, len(boxes)))
    return start, end


def skip_separators(text: str, start: int) -> int:
    while start < len(text) and text[start] in " \t　:：:-—|/\\（）()[]【】":
        start += 1
    return start


def find_value_end(text: str, start: int) -> int:
    if start >= len(text):
        return start
    stops = [len(text)]
    for marker in ["；", ";", "。", "\n"]:
        index = text.find(marker, start)
        if index > start:
            stops.append(index)
    for token in ALL_LABELS:
        index = text.find(token, start)
        if index > start:
            stops.append(index)
    end = min(stops)
    while end > start and text[end - 1].isspace():
        end -= 1
    return end


def apply_pdf_redactions(source_pdf: Path, output_pdf: Path, redactions: list[dict]):
    doc = fitz.open(source_pdf)
    temp_path = safe_temp_path(".pdf")
    try:
        by_page: dict[int, list[dict]] = defaultdict(list)
        for item in redactions:
            by_page[int(item["page"])].append(item)

        for page_index in range(doc.page_count):
            page = doc[page_index]
            page_items = by_page.get(page_index + 1, [])
            for item in page_items:
                page.add_redact_annot(fitz.Rect(item["rect"]), fill=(0, 0, 0))
            if page_items:
                page.apply_redactions()

        doc.save(temp_path, garbage=4, deflate=True)
    finally:
        doc.close()
    os.replace(temp_path, output_pdf)


def render_pdf_pages(pdf_path: Path, pages_dir: Path, output_stem: str, progress_callback=None) -> list[dict]:
    for old_image in pages_dir.glob("*.png"):
        old_image.unlink()

    doc = fitz.open(pdf_path)
    pages: list[dict] = []
    matrix = fitz.Matrix(IMAGE_ZOOM, IMAGE_ZOOM)
    try:
        for page_index in range(doc.page_count):
            page = doc[page_index]
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            image_name = f"{output_stem}_page_{page_index + 1:03d}.png"
            image_path = pages_dir / image_name
            temp_image_path = safe_temp_path(".png")
            pixmap.save(temp_image_path)
            os.replace(temp_image_path, image_path)
            pages.append(
                {
                    "page": page_index + 1,
                    "image": f"pages/{image_name}",
                    "image_width": pixmap.width,
                    "image_height": pixmap.height,
                    "pdf_width": round(float(page.rect.width), 2),
                    "pdf_height": round(float(page.rect.height), 2),
                }
            )
            progress_callback and progress_callback(f"正在处理第 {page_index + 1} / {doc.page_count} 页", page_index + 1, doc.page_count)
    finally:
        doc.close()
    return pages


def render_single_pdf_page(pdf_path: Path, image_path: Path, page_index: int):
    doc = fitz.open(pdf_path)
    matrix = fitz.Matrix(IMAGE_ZOOM, IMAGE_ZOOM)
    temp_image_path = safe_temp_path(".png")
    try:
        page = doc[page_index]
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        pixmap.save(temp_image_path)
        os.replace(temp_image_path, image_path)
    finally:
        doc.close()
        if temp_image_path.exists():
            temp_image_path.unlink()


def apply_image_masks(job_dir: Path, pages: list[dict], redactions: list[dict]):
    by_page: dict[int, list[dict]] = defaultdict(list)
    for item in redactions:
        by_page[int(item["page"])].append(item)

    for page in pages:
        masks = by_page.get(int(page["page"]), [])
        if not masks:
            continue
        image_path = job_dir / page["image"]
        with Image.open(image_path) as source:
            image = source.convert("RGB")
        draw = ImageDraw.Draw(image)
        for item in masks:
            rect = pdf_rect_to_image_rect(item["rect"], page)
            if rect[2] <= rect[0] or rect[3] <= rect[1]:
                continue
            method = normalize_image_method(item.get("method", "mosaic"))
            if method == "black_box":
                draw.rectangle(rect, fill=(0, 0, 0))
            elif method == "blur":
                blur_region(image, rect)
            elif method == "gray_mosaic":
                pixelate_region(image, rect, gray=True)
            else:
                pixelate_region(image, rect, gray=False)
        image.save(image_path)


def pdf_rect_to_image_rect(rect: list[float], page: dict) -> tuple[int, int, int, int]:
    scale_x = float(page["image_width"]) / float(page["pdf_width"])
    scale_y = float(page["image_height"]) / float(page["pdf_height"])
    x0 = int(max(0, round(rect[0] * scale_x)))
    y0 = int(max(0, round(rect[1] * scale_y)))
    x1 = int(min(page["image_width"], round(rect[2] * scale_x)))
    y1 = int(min(page["image_height"], round(rect[3] * scale_y)))
    return x0, y0, x1, y1


def pixelate_region(image: Image.Image, rect: tuple[int, int, int, int], gray: bool):
    x0, y0, x1, y1 = rect
    if x1 - x0 < 2 or y1 - y0 < 2:
        return
    region = image.crop(rect)
    small_size = (max(1, region.width // 14), max(1, region.height // 14))
    resampling = getattr(Image, "Resampling", Image)
    region = region.resize(small_size, resampling.BILINEAR)
    region = region.resize((x1 - x0, y1 - y0), resampling.NEAREST)
    if gray:
        region = ImageOps.grayscale(region).convert("RGB")
    image.paste(region, (x0, y0))


def blur_region(image: Image.Image, rect: tuple[int, int, int, int]):
    x0, y0, x1, y1 = rect
    if x1 - x0 < 2 or y1 - y0 < 2:
        return
    region = image.crop(rect)
    radius = max(4, min(18, min(region.width, region.height) // 5))
    image.paste(region.filter(ImageFilter.GaussianBlur(radius=radius)), (x0, y0))


def create_images_zip(pages: list[dict], zip_path: Path):
    temp_path = zip_path.with_suffix(".tmp.zip")
    if temp_path.exists():
        temp_path.unlink()
    with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for page in pages:
            image_path = zip_path.parent / page["image"]
            archive.write(image_path, image_path.name)
    os.replace(temp_path, zip_path)


def create_mosaic_pdf_from_images(job_dir: Path, pages: list[dict], output_pdf: Path):
    if not pages:
        raise ProcessingError("没有可用于生成脱敏 PDF 的分页图片。")

    doc = fitz.open()
    temp_path = safe_temp_path(".pdf")
    try:
        for page_info in pages:
            width = float(page_info["pdf_width"])
            height = float(page_info["pdf_height"])
            image_path = job_dir / page_info["image"]
            page = doc.new_page(width=width, height=height)
            page.insert_image(fitz.Rect(0, 0, width, height), stream=image_path.read_bytes())
        doc.save(temp_path, garbage=4, deflate=True)
        os.replace(temp_path, output_pdf)
    finally:
        doc.close()
        if temp_path.exists():
            temp_path.unlink()


def safe_temp_path(suffix: str) -> Path:
    SAFE_TEMP_ROOT.mkdir(parents=True, exist_ok=True)
    return SAFE_TEMP_ROOT / f"{uuid4().hex}{suffix}"


def get_pdf_page_count(pdf_path: Path) -> int:
    doc = fitz.open(pdf_path)
    try:
        return int(doc.page_count)
    finally:
        doc.close()


def read_docx_declared_page_count(input_path: Path) -> int | None:
    if input_path.suffix.lower() != ".docx":
        return None
    try:
        with zipfile.ZipFile(input_path) as archive:
            with archive.open("docProps/app.xml") as file:
                root = ElementTree.parse(file).getroot()
    except (KeyError, zipfile.BadZipFile, ElementTree.ParseError, OSError):
        return None

    for element in root.iter():
        tag = element.tag.rsplit("}", 1)[-1]
        if tag == "Pages" and element.text:
            try:
                value = int(element.text.strip())
            except ValueError:
                return None
            return value if value > 0 else None
    return None


def build_page_count_check(
    suffix: str,
    expected_pages: int | None,
    docx_declared_page_count: int | None,
    converted_page_count: int,
) -> dict:
    normalized_expected = normalize_expected_pages(expected_pages)
    source_type = suffix.lstrip(".").upper()
    conversion_method = conversion_method_for_suffix(suffix)
    check = {
        "source_type": source_type,
        "conversion_method": conversion_method,
        "converted_page_count": converted_page_count,
        "expected_page_count": normalized_expected,
        "docx_declared_page_count": docx_declared_page_count,
        "page_count_anomaly": False,
        "page_count_warning": "",
        "conversion_warning": "",
    }

    if suffix == ".pdf":
        check["conversion_summary"] = "PDF 直接处理，未进行 Word 排版转换。"
        return check

    check["conversion_warning"] = "自动转换可能导致页数、行距、表格和分页变化，仅建议用于初步预览，不建议作为最终详情图来源。"
    compare_pages = normalized_expected or docx_declared_page_count
    if not compare_pages or converted_page_count <= 0:
        check["conversion_summary"] = f"已通过 LibreOffice 自动转换，转换后 {converted_page_count} 页。"
        return check

    difference = abs(converted_page_count - compare_pages)
    ratio = difference / max(compare_pages, 1)
    if difference >= 3 and ratio >= 0.08:
        check["page_count_anomaly"] = True
        if normalized_expected:
            check["page_count_warning"] = (
                f"页数从 {normalized_expected} 变为 {converted_page_count}，疑似自动转换导致排版重排。"
                "建议先用 WPS/Word 导出 PDF 后再处理。"
            )
        else:
            check["page_count_warning"] = (
                f"DOCX 属性页数为 {docx_declared_page_count}，自动转换后为 {converted_page_count}，"
                "疑似自动转换导致排版重排。建议先用 WPS/Word 导出 PDF 后再处理。"
            )
    check["conversion_summary"] = f"已通过 LibreOffice 自动转换，转换后 {converted_page_count} 页。"
    return check


def normalize_expected_pages(expected_pages: int | None) -> int | None:
    try:
        value = int(expected_pages) if expected_pages is not None else 0
    except (TypeError, ValueError):
        return None
    return value if value > 0 else None


def conversion_method_for_suffix(suffix: str) -> str:
    return "direct_pdf" if suffix == ".pdf" else "libreoffice_auto"


def convert_to_source_pdf(input_path: Path, work_dir: Path) -> Path:
    target = work_dir / "source.pdf"
    if input_path.suffix.lower() == ".pdf":
        shutil.copy2(input_path, target)
        return target

    soffice = find_soffice()
    if not soffice:
        raise ProcessingError("未找到 LibreOffice。处理 DOC/DOCX 前请安装 LibreOffice，并确保 soffice 可执行。")

    before = set(work_dir.glob("*.pdf"))
    command = [
        str(soffice),
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(work_dir),
        str(input_path),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=180, check=False)
    if result.returncode != 0:
        message = (result.stderr or result.stdout or "").strip()
        raise ProcessingError(f"LibreOffice 转 PDF 失败：{message}")

    converted = work_dir / f"{input_path.stem}.pdf"
    if not converted.exists():
        candidates = [path for path in work_dir.glob("*.pdf") if path not in before]
        if not candidates:
            raise ProcessingError("LibreOffice 未生成 PDF 文件。")
        converted = max(candidates, key=lambda path: path.stat().st_mtime)

    if converted.resolve() != target.resolve():
        shutil.copy2(converted, target)
    return target


def find_soffice() -> Path | None:
    candidates = [
        os.environ.get("LIBREOFFICE_PATH"),
        r"F:\Tools\LibreOffice\program\soffice.com",
        r"F:\Tools\LibreOffice\program\soffice.exe",
        shutil.which("soffice.com"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        r"C:\Program Files\LibreOffice\program\soffice.com",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.com",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for candidate in candidates:
        if not candidate:
            continue
        path = Path(candidate)
        if path.exists():
            return path
    return None


def dedupe_redactions(redactions: list[dict]) -> list[dict]:
    seen = set()
    unique = []
    for item in redactions:
        rect = round_rect(item["rect"])
        key = (item["page"], item["sensitive_type"], tuple(rect), item.get("manual", False), item.get("rule"), item.get("method"))
        if key in seen:
            continue
        seen.add(key)
        copy = dict(item)
        copy["rect"] = rect
        unique.append(copy)
    unique.sort(key=lambda item: (item["page"], item["rect"][1], item["rect"][0], item.get("manual", False)))
    return unique


def get_manual_redactions(report: dict) -> list[dict]:
    return [item for item in report.get("redactions", []) if item.get("manual")]


def read_report(job_dir: Path) -> dict:
    meta = read_json(job_dir / "job_meta.json")
    output_stem = meta.get("output_stem") or prefixed_output_stem(meta["safe_stem"])
    candidates = [
        job_dir / f"{output_stem}_处理报告.json",
        job_dir / f"{meta['safe_stem']}_report.json",
    ]
    for path in candidates:
        if path.exists():
            return read_json(path)
    raise ProcessingError("没有找到处理报告，请先处理文件。")


def report_path_for_job(job_dir: Path) -> Path:
    meta = read_json(job_dir / "job_meta.json")
    return job_dir / f"{report_output_stem(job_dir)}_处理报告.json"


def report_output_stem(job_dir: Path) -> str:
    meta = read_json(job_dir / "job_meta.json")
    return meta.get("output_stem") or prefixed_output_stem(meta["safe_stem"])


def write_report(job_dir: Path, report: dict):
    write_json(report_path_for_job(job_dir), report)


def manual_mask_record(item: dict) -> dict:
    return {
        "id": item.get("id"),
        "job_id": item.get("job_id"),
        "page_index": int(item.get("page_index", int(item.get("page", 1)) - 1)),
        "page": int(item.get("page", 1)),
        "rect": item.get("preview_rect") or item.get("rect"),
        "pdf_rect": item.get("rect"),
        "preview_image_size": item.get("preview_image_size"),
        "mask_type": normalize_image_method(item.get("mask_type") or item.get("method")),
        "source": "manual",
        "created_at": item.get("created_at"),
    }


def method_for_type(sensitive_type: str, default_method: str) -> str:
    if sensitive_type in BLACK_BOX_IMAGE_TYPES:
        return "black_box"
    return normalize_image_method(default_method)


def visual_method_for_rect(sensitive_type: str, rule: str, rect, method: str) -> str:
    if method == "black_box" or sensitive_type in BLACK_BOX_IMAGE_TYPES:
        return "black_box"
    is_large_table_value = rule.startswith("table_adjacent:") and rect.width >= 140
    is_long_named_value = sensitive_type in {"address_seed", "project_seed"} and rect.width >= 220
    if is_large_table_value or is_long_named_value:
        return "gray_mosaic"
    return method


def normalize_image_method(method: str | None) -> str:
    if method in ALLOWED_IMAGE_METHODS:
        return str(method)
    return "mosaic"


def category_for_label(label: str) -> str:
    if label in SUBJECT_KEYWORDS:
        return "subject"
    if label in PROJECT_KEYWORDS:
        return "project"
    if label in FINANCIAL_KEYWORDS:
        return "financial"
    if label in BUSINESS_KEYWORDS:
        return "business"
    return "label"


def sensitive_type_for_category(category: str, label: str) -> str:
    if "代码" in label:
        return "project_code"
    if category == "financial":
        return "financial_value"
    if category == "subject":
        return "entity_seed"
    if category == "project":
        return "project_seed" if "地址" not in label and "地点" not in label and "位置" not in label else "address_seed"
    if category == "business":
        return "business_value"
    return "label_value"


def clean_seed_text(text: str) -> str:
    value = re.sub(r"\s+", " ", text or "").strip()
    value = re.sub(r"^[：:;；,，、\-\s/\\|（）()【】\[\]]+", "", value)
    value = re.sub(r"[：:;；,，、\-\s/\\|（）()【】\[\]]+$", "", value)
    for label in sorted(ALL_LABELS, key=len, reverse=True):
        if value.startswith(label):
            value = value[len(label) :]
            value = re.sub(r"^[：:;；,，、\-\s/\\|（）()【】\[\]]+", "", value)
    return strip_generic_title_words(value)


def strip_generic_title_words(text: str) -> str:
    value = text.strip()
    for word in GENERIC_TITLE_WORDS:
        if value.endswith(word):
            value = value[: -len(word)].strip()
    return value


def is_label_only(text: str) -> bool:
    stripped = compact_label_text(text)
    return stripped in {compact_label_text(label) for label in ALL_LABELS}


def compact_label_text(text: str) -> str:
    return re.sub(r"[\s：:;；,，、（）()【】\[\]\-—|/\\]", "", text)


def has_digit(text: str) -> bool:
    return bool(re.search(r"\d", text))


def looks_like_year_only(text: str) -> bool:
    stripped = text.strip()
    return bool(re.fullmatch(r"(?:19|20)\d{2}\s*(?:年)?", stripped))


def is_inside_project_code(text: str, start: int, end: int) -> bool:
    for match in PROJECT_CODE_RE.finditer(text):
        if match.start() <= start and end <= match.end():
            return True
    return False


def prefixed_output_stem(safe_stem: str) -> str:
    return safe_stem if safe_stem.startswith(OUTPUT_PREFIX) else f"{OUTPUT_PREFIX}{safe_stem}"


def round_rect(rect: list[float] | tuple[float, float, float, float]) -> list[float]:
    return [round(float(value), 2) for value in rect]


def sanitize_filename(name: str) -> str:
    name = Path(name).name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.strip(" .")
    return name or "document"


def resolve_job_dir(job_id: str, output_root: Path) -> Path:
    if "/" in job_id or "\\" in job_id or not job_id.strip():
        raise ProcessingError("任务 ID 无效。")
    root = output_root.resolve()
    job_dir = (root / job_id).resolve()
    if root not in job_dir.parents:
        raise ProcessingError("任务路径无效。")
    if not job_dir.exists():
        raise ProcessingError("任务不存在。")
    return job_dir


def read_json(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def write_json(path: Path, data: dict):
    with path.open("w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)
