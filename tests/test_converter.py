import tempfile
import unittest
from pathlib import Path

from ppt_editable_converter.converter import (
    SUPPORTED_EXTENSIONS,
    build_output_path,
    discover_files,
)


class ConverterHelpersTests(unittest.TestCase):
    def test_build_output_path_uses_pptx_by_default(self):
        result = build_output_path(
            Path("课件") / "第1章.PPT",
            Path("输出"),
        )
        self.assertEqual(result, Path("输出") / "第1章-editable.pptx")

    def test_build_output_path_supports_pptm(self):
        result = build_output_path("source.pptx", "out", output_format="pptm")
        self.assertEqual(result, Path("out") / "source-editable.pptm")

    def test_discover_files_is_case_insensitive_and_deduplicates(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "a.PPT").write_bytes(b"")
            (root / "b.pptx").write_bytes(b"")
            (root / "ignore.pdf").write_bytes(b"")
            nested = root / "nested"
            nested.mkdir()
            (nested / "c.PpSm").write_bytes(b"")

            files = discover_files([root], recursive=True)

            self.assertEqual([file.name for file in files], ["a.PPT", "b.pptx", "c.PpSm"])

    def test_supported_extensions_include_legacy_and_open_xml_formats(self):
        self.assertTrue({".ppt", ".pptx", ".pptm", ".ppsx", ".potx"}.issubset(SUPPORTED_EXTENSIONS))


if __name__ == "__main__":
    unittest.main()

