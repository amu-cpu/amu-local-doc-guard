from __future__ import annotations

import argparse
import sys
from pathlib import Path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="从原始 Excel 路径生成模型版和客户交付值版，并默认把两个 Excel 输出到原文件同目录。")
    parser.add_argument("--excel-path", required=True, help="原始 Excel 文件路径")
    parser.add_argument("--output-root", default=None, help="内部任务输出目录；不填则用工具默认 output")
    parser.add_argument("--keep-hidden-sheets", action="store_true", help="生成客户交付值版时保留隐藏/veryHidden 表")
    parser.add_argument("--password", default="123456", help="编辑/保护密码")
    parser.add_argument("--output-dir", default=None, help="可选：指定两个 Excel 输出目录；不填则使用原文件同目录。质检报告始终放工具 output")
    parser.add_argument("--allow-value-as-source", action="store_true", help="允许把客户交付值版作为源文件（默认拒绝）")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    project_root = Path(__file__).resolve().parent.parent
    sys.path.insert(0, str(project_root))

    from desensitizer.excel_value_version import generate_value_version_from_path
    from desensitizer.processor import OUTPUT_ROOT, ProcessingError

    excel_path_input = Path(args.excel_path).expanduser()
    excel_path = excel_path_input.resolve()
    if not excel_path.exists():
        print(f"Excel 文件不存在：{excel_path_input}", file=sys.stderr)
        return 1

    if "客户交付值版" in excel_path.name and not args.allow_value_as_source:
        print("当前输入疑似已是客户交付值版，请提供模型版或原始带公式版。", file=sys.stderr)
        return 3

    output_root = Path(args.output_root).expanduser() if args.output_root else OUTPUT_ROOT
    output_dir = Path(args.output_dir).expanduser() if args.output_dir else excel_path_input.parent
    source_mode = "value" if "客户交付值版" in excel_path.name else "model" if "模型版" in excel_path.name else "raw"

    try:
        report = generate_value_version_from_path(
            excel_path,
            output_root=output_root,
            original_name=excel_path.name,
            output_dir=output_dir,
            keep_hidden_sheets=bool(args.keep_hidden_sheets),
            protection_password=args.password,
            publish_to_desktop=False,
            allow_value_as_source=bool(args.allow_value_as_source),
        )
    except ProcessingError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    print(f"结论：{report.get('final_conclusion', '')}")
    print(f"输入类型：{source_mode}")
    print(f"模型版：{report.get('model_version_published_path', report.get('model_version_path', ''))}")
    print(f"客户交付值版：{report.get('delivery_value_published_path', report.get('delivery_value_path', ''))}")
    print(f"质检报告（工具 output）：{report.get('markdown_report_path', '')}")
    print(f"JSON 报告（工具 output）：{report.get('json_report_path', '')}")
    timing = report.get("timing_summary", {})
    if timing:
        print("耗时：")
        print(f"  源文件检查：{timing.get('source_file_check_seconds', 0)}s")
        print(f"  Excel COM 重算：{timing.get('excel_com_recalculation_seconds', 0)}s")
        print(f"  转值：{timing.get('value_conversion_seconds', 0)}s")
        print(f"  复检：{timing.get('delivery_recheck_seconds', 0)}s")
        print(f"  总耗时：{timing.get('total_seconds', 0)}s")
    print(f"误用客户交付值版作为源文件：{report.get('is_value_version_used_as_source', False)}")
    return 0 if report.get("delivery_ready") else 2


if __name__ == "__main__":
    raise SystemExit(main())
