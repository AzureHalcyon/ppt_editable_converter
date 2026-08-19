"""Command-line interface for the PowerPoint converter."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Sequence

from .converter import (
    DEFAULT_SUFFIX,
    OUTPUT_FORMATS,
    PowerPointConverter,
    build_output_path,
    discover_files,
)


def build_parser() -> argparse.ArgumentParser:
    formats = ", ".join(sorted(OUTPUT_FORMATS))
    parser = argparse.ArgumentParser(
        prog="ppt-editable-convert",
        description=(
            "通过 Microsoft PowerPoint 将 .ppt/.pptx 等演示文稿复制到新的可编辑文件中。"
        ),
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        metavar="INPUT",
        help="一个 PowerPoint 文件，或包含 PowerPoint 文件的文件夹",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default="editable_output",
        metavar="DIR",
        help="输出文件夹（默认：当前目录下的 editable_output）",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="递归扫描输入文件夹",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="允许覆盖已存在的输出文件",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="显示 PowerPoint 窗口（调试或需要手动处理提示时使用）",
    )
    parser.add_argument(
        "--suffix",
        default=DEFAULT_SUFFIX,
        help=f"输出文件名后缀（默认：{DEFAULT_SUFFIX}）",
    )
    parser.add_argument(
        "--output-format",
        choices=sorted(OUTPUT_FORMATS),
        default="pptx",
        help=f"输出格式（默认：pptx；可选：{formats}）",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="只列出将要处理的文件，不启动 PowerPoint",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    output_dir = Path(args.output_dir).expanduser().resolve()

    try:
        files = discover_files(
            args.inputs,
            recursive=args.recursive,
            exclude_dir=output_dir,
        )
    except (FileNotFoundError, ValueError) as exc:
        parser.error(str(exc))

    if not files:
        parser.error("没有找到可处理的 PowerPoint 文件")

    jobs: list[tuple[Path, Path]] = []
    seen_outputs: set[Path] = set()
    for source in files:
        destination = build_output_path(
            source,
            output_dir,
            suffix=args.suffix,
            output_format=args.output_format,
        ).resolve()
        if destination in seen_outputs:
            parser.error(f"多个输入文件产生了相同的输出路径: {destination}")
        seen_outputs.add(destination)
        jobs.append((source, destination))

    if args.dry_run:
        for source, destination in jobs:
            print(f"{source} -> {destination}")
        return 0

    failures = 0
    try:
        with PowerPointConverter(visible=args.visible) as converter:
            for source, destination in jobs:
                try:
                    result = converter.convert(
                        source,
                        destination,
                        overwrite=args.overwrite,
                        output_format=args.output_format,
                    )
                except Exception as exc:
                    failures += 1
                    print(f"失败: {source.name}: {exc}")
                else:
                    print(
                        f"完成: {source.name} ({result.slide_count} 张幻灯片) -> "
                        f"{destination}"
                    )
    except Exception as exc:
        parser.exit(1, f"无法启动转换器: {exc}\n")

    if failures:
        print(f"处理结束：成功 {len(jobs) - failures} 个，失败 {failures} 个")
        return 1
    print(f"处理结束：成功 {len(jobs)} 个")
    return 0

