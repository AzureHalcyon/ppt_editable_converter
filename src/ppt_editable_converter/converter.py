"""PowerPoint conversion implementation.

The conversion is performed through PowerPoint's COM automation interface:
the source presentation is opened read-only, its slides are inserted into a
new presentation, and the new presentation is saved as an Open XML file.
"""

from __future__ import annotations

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
            self._app.Visible = self.visible
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

