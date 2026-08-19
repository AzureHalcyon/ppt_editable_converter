"""Reusable PowerPoint conversion helpers."""

from .converter import (
    DEFAULT_SUFFIX,
    SUPPORTED_EXTENSIONS,
    ConversionError,
    ConversionResult,
    PowerPointConverter,
    build_output_path,
    discover_files,
)

__all__ = [
    "DEFAULT_SUFFIX",
    "SUPPORTED_EXTENSIONS",
    "ConversionError",
    "ConversionResult",
    "PowerPointConverter",
    "build_output_path",
    "discover_files",
]

