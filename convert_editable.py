"""PowerPoint conversion implementation.

The conversion is performed through PowerPoint's COM automation interface:
the source presentation is opened read-only, its slides are inserted into a
new presentation, and the new presentation is saved as an Open XML file.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence


DEFAULT_SUFFIX = "-editable"

# These are the PowerPoint file extensions that can normally be opened by
# Presentations.Open and imported with Slides.InsertFromFile.  The extension
# check is deliberately case-insensitive.
SUPPORTED_EXTENSIONS = frozenset(
    {
        ".ppt",
        ".pptx",
        ".pptm",
        ".pps",
        ".ppsx",
        ".ppsm",
        ".pot",
        ".potx",
        ".potm",
    }
)

OUTPUT_FORMATS = {
    "pptx": (".pptx", 24),  # ppSaveAsOpenXMLPresentation
    "pptm": (".pptm", 25),  # ppSaveAsOpenXMLPresentationMacroEnabled
}


class ConversionError(RuntimeError):
    """Raised when PowerPoint cannot convert a presentation."""


@dataclass(frozen=True)
class ConversionResult:
    """Information about one successful conversion."""

    source: Path
    output: Path
    slide_count: int


def is_supported_file(path: Path) -> bool:
    """Return whether *path* has a supported PowerPoint extension."""

    return path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS


def _is_relative_to(path: Path, parent: Path) -> bool:
    """Compatibility helper for Path.is_relative_to."""

    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def discover_files(
    inputs: Sequence[str | Path],
    *,
    recursive: bool = False,
    exclude_dir: str | Path | None = None,
    exclude_suffix: str | None = DEFAULT_SUFFIX,
) -> list[Path]:
    """Expand files and directories into a sorted, de-duplicated file list.

    Explicitly supplied files must have a supported extension. Unsupported
    files found while scanning a directory are ignored.  ``exclude_dir`` is
    useful when the output directory is inside a recursively scanned input
    directory, preventing generated files from being picked up again.
    """

    if not inputs:
        raise ValueError("至少需要指定一个输入文件或文件夹")

    excluded = Path(exclude_dir).expanduser().resolve() if exclude_dir else None
    found: dict[Path, Path] = {}

    for raw_input in inputs:
        path = Path(raw_input).expanduser()
        if not path.exists():
            raise FileNotFoundError(f"找不到输入路径: {path}")

        if path.is_file():
            if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
                supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
                raise ValueError(
                    f"不支持的 PowerPoint 文件格式: {path.suffix or '(无扩展名)'}; "
                    f"支持: {supported}"
                )
            candidates: Iterable[Path] = (path,)
        elif path.is_dir():
            candidates = path.rglob("*") if recursive else path.iterdir()
        else:
            raise ValueError(f"输入路径不是普通文件或文件夹: {path}")

        for candidate in candidates:
            resolved = candidate.resolve()
            if excluded and _is_relative_to(resolved, excluded):
                continue
            if exclude_suffix and candidate.stem.endswith(exclude_suffix):
                continue
            if is_supported_file(candidate):
                found[resolved] = resolved

    return sorted(found.values(), key=lambda item: str(item).casefold())


def build_output_path(
    source: str | Path,
    output_dir: str | Path | None = None,
    *,
    suffix: str = DEFAULT_SUFFIX,
    output_format: str = "pptx",
) -> Path:
    """Build the output path for a source presentation.

    By default the output is placed beside the source.  The CLI passes an
    explicit output directory, which makes batch conversion predictable.
    """

    source_path = Path(source)
    format_name = output_format.lower().lstrip(".")
    try:
        extension, _ = OUTPUT_FORMATS[format_name]
    except KeyError as exc:
        supported = ", ".join(sorted(OUTPUT_FORMATS))
        raise ValueError(f"不支持的输出格式 {output_format!r}; 支持: {supported}") from exc

    destination_dir = Path(output_dir) if output_dir is not None else source_path.parent
    return destination_dir / f"{source_path.stem}{suffix}{extension}"


class PowerPointConverter:
    """Context manager around a dedicated PowerPoint COM instance."""

    def __init__(self, *, visible: bool = False) -> None:
        self.visible = visible
        self._app = None

    def __enter__(self) -> "PowerPointConverter":
        try:
            import win32com.client as win32
        except ImportError as exc:  # pragma: no cover - depends on Windows setup
            raise ConversionError(
                "未找到 pywin32。请在安装 Microsoft PowerPoint 的 Windows Python 中运行 "
                "pip install pywin32"
            ) from exc

        try:
            self._app = win32.DispatchEx("PowerPoint.Application")
            # PowerPoint starts hidden when launched through COM. Some Office
            # builds reject an explicit ``Visible = False`` assignment with
            # "Hiding the application window is not allowed", so only make
            # the affirmative request here.
            if self.visible:
                self._app.Visible = True
            # 0 = ppAlertsNone. We ask the user to opt into overwrite through
            # the CLI, so PowerPoint should not stop on an overwrite prompt.
            self._app.DisplayAlerts = 0
        except Exception as exc:  # pragma: no cover - depends on Office setup
            raise ConversionError(
                "无法启动 Microsoft PowerPoint，请确认 Office 已安装且 COM 可用"
            ) from exc
        return self

    @property
    def _powerpoint(self):
        if self._app is None:
            raise RuntimeError("PowerPointConverter 必须在 with 代码块中使用")
        return self._app

    def convert(
        self,
        source: str | Path,
        destination: str | Path,
        *,
        overwrite: bool = False,
        output_format: str = "pptx",
    ) -> ConversionResult:
        """Convert one presentation into a new editable presentation."""

        source_path = Path(source).expanduser().resolve()
        destination_path = Path(destination).expanduser().resolve()
        if not is_supported_file(source_path):
            raise ValueError(f"不支持的输入文件: {source_path}")
        if source_path == destination_path:
            raise ValueError("输出文件不能覆盖输入文件")
        if destination_path.exists() and not overwrite:
            raise FileExistsError(
                f"输出文件已存在: {destination_path}（如需覆盖请使用 --overwrite）"
            )

        format_name = output_format.lower().lstrip(".")
        try:
            _, save_type = OUTPUT_FORMATS[format_name]
        except KeyError as exc:
            supported = ", ".join(sorted(OUTPUT_FORMATS))
            raise ValueError(f"不支持的输出格式 {output_format!r}; 支持: {supported}") from exc

        destination_path.parent.mkdir(parents=True, exist_ok=True)
        source_presentation = None
        target_presentation = None
        try:
            # ReadOnly=True, Untitled=False, WithWindow=False.
            source_presentation = self._powerpoint.Presentations.Open(
                str(source_path), True, False, False
            )
            slide_count = int(source_presentation.Slides.Count)
            if slide_count < 1:
                raise ConversionError(f"源文件没有幻灯片: {source_path.name}")

            target_presentation = self._powerpoint.Presentations.Add(False)
            # InsertFromFile lets PowerPoint parse the source and writes new
            # slides into an unprotected target presentation.
            target_presentation.Slides.InsertFromFile(
                str(source_path), 0, 1, slide_count
            )
            target_presentation.SaveAs(str(destination_path), save_type)
        except ConversionError:
            raise
        except Exception as exc:  # pragma: no cover - depends on Office setup
            raise ConversionError(
                f"转换失败: {source_path.name}: {exc}"
            ) from exc
        finally:
            if target_presentation is not None:
                try:
                    target_presentation.Close()
                except Exception:
                    pass
            if source_presentation is not None:
                try:
                    source_presentation.Close()
                except Exception:
                    pass

        return ConversionResult(source_path, destination_path, slide_count)

    def close(self) -> None:
        if self._app is not None:
            try:
                self._app.Quit()
            finally:
                self._app = None

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()


# The script is intentionally self-contained. Copy this file into a folder
# containing presentations, then run ``python convert_editable.py``.
SCRIPT_DIR = Path(__file__).resolve().parent


def default_scan_dir() -> Path:
    """Choose the current folder, falling back to the script folder.

    The current folder is useful when the module is installed as a CLI entry
    point. The fallback keeps the copy-the-script workflow working even when
    the script is launched from another working directory.
    """

    current_dir = Path.cwd().resolve()
    if any(is_supported_file(path) for path in current_dir.iterdir()):
        return current_dir
    return SCRIPT_DIR


def build_parser() -> argparse.ArgumentParser:
    formats = ", ".join(sorted(OUTPUT_FORMATS))
    parser = argparse.ArgumentParser(
        prog="convert_editable.py",
        description="选择并转换当前目录中的 PowerPoint 文件。",
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        metavar="INPUT",
        help="可选的文件或文件夹；不填写时扫描本脚本所在目录",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        help="递归扫描输入文件夹",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="跳过选择菜单，转换找到的全部文件",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="允许覆盖已有的输出文件",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="显示 PowerPoint 窗口",
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
        help="只显示选择和输出路径，不启动 PowerPoint",
    )
    return parser


def _display_name(path: Path, root: Path) -> str:
    """Use a short relative name when a file is under the scan directory."""

    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def choose_files(files: Sequence[Path], root: Path) -> list[Path]:
    """Display an interactive menu and return the user's selection."""

    print("发现以下 PowerPoint 文件：")
    for index, path in enumerate(files, start=1):
        print(f"  [{index}] {_display_name(path, root)}")
    print("  [a] 全部转换")
    print("  [q] 退出")

    while True:
        try:
            answer = input("请输入编号（可用逗号选择多个）：").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print()
            return []

        if answer in {"q", "quit", "0"}:
            return []
        if answer in {"a", "all"}:
            return list(files)
        if not answer:
            print("请输入文件编号，或输入 a 全部转换、q 退出。")
            continue

        selected: list[Path] = []
        invalid = False
        for token in answer.replace("，", ",").split(","):
            token = token.strip()
            try:
                index = int(token)
            except ValueError:
                invalid = True
                break
            if not 1 <= index <= len(files):
                invalid = True
                break
            path = files[index - 1]
            if path not in selected:
                selected.append(path)

        if invalid or not selected:
            print("编号无效，请重新输入。")
            continue
        return selected


def _confirm_overwrite(path: Path) -> bool:
    try:
        answer = input(f"输出文件已存在：{path.name}，是否覆盖？[y/N] ").strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    return answer in {"y", "yes", "是"}


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    # With no positional arguments, the copied script operates on its own
    # directory. This is what makes it portable by simple copy and paste.
    scan_root = default_scan_dir() if not args.inputs else Path.cwd().resolve()
    input_paths: Sequence[str | Path] = args.inputs or [scan_root]
    try:
        files = discover_files(
            input_paths,
            recursive=args.recursive,
            exclude_suffix=args.suffix,
        )
    except (FileNotFoundError, ValueError) as exc:
        print(f"输入错误：{exc}")
        return 2

    if not files:
        print("没有找到可处理的 PowerPoint 文件。")
        return 1

    if args.all or args.inputs or args.dry_run:
        selected = files
    else:
        selected = choose_files(files, scan_root)
        if not selected:
            print("没有选择文件，程序结束。")
            return 0

    jobs = [
        (
            source,
            build_output_path(
                source,
                source.parent,
                suffix=args.suffix,
                output_format=args.output_format,
            ).resolve(),
        )
        for source in selected
    ]

    if args.dry_run:
        for source, destination in jobs:
            print(f"{source} -> {destination}")
        return 0

    pending: list[tuple[Path, Path]] = []
    skipped = 0
    interactive = not args.inputs and not args.all
    for source, destination in jobs:
        if destination.exists() and not args.overwrite:
            if interactive and _confirm_overwrite(destination):
                pending.append((source, destination))
            else:
                print(f"跳过（输出已存在）：{destination}")
                skipped += 1
        else:
            pending.append((source, destination))

    if not pending:
        print(f"没有需要转换的文件（跳过 {skipped} 个）。")
        return 0

    successes = 0
    failures = 0
    try:
        with PowerPointConverter(visible=args.visible) as converter:
            for source, destination in pending:
                try:
                    result = converter.convert(
                        source,
                        destination,
                        overwrite=args.overwrite or destination.exists(),
                        output_format=args.output_format,
                    )
                except Exception as exc:
                    failures += 1
                    print(f"失败：{source.name}：{exc}")
                else:
                    successes += 1
                    print(
                        f"完成：{source.name}（{result.slide_count} 张幻灯片） -> "
                        f"{destination.name}"
                    )
    except Exception as exc:
        print(f"无法启动 PowerPoint：{exc}")
        return 1

    print(f"处理结束：成功 {successes} 个，跳过 {skipped} 个，失败 {failures} 个。")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
