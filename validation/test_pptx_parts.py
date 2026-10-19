"""Synthetic ZIP/XML tests, not editor compatibility tests."""

from pathlib import Path
import tempfile
import unittest
import zipfile

from pptx_parts import compare, inspect, xml_inventory


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
BASE = {"[Content_Types].xml": b"<Types/>", "_rels/.rels": b"<Relationships/>",
        "ppt/presentation.xml": b"<presentation/>"}


class PackageTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)

    def package(self, name, parts, compression=zipfile.ZIP_STORED):
        path = Path(self.temp.name) / name
        with zipfile.ZipFile(path, "w", compression=compression) as archive:
            for key, value in parts.items():
                archive.writestr(key, value)
        return inspect(path)

    def test_zip_order_and_compression_do_not_count_as_part_changes(self):
        before = self.package("a.pptx", BASE)
        after = self.package("b.pptx", dict(reversed(list(BASE.items()))), zipfile.ZIP_DEFLATED)
        self.assertTrue(compare(before, after)["part_payloads_identical"])

    def test_removed_chart_and_changed_workbook_are_reported(self):
        before = self.package("a.pptx", {**BASE, "ppt/charts/chart1.xml": b"<chart/>",
                                        "ppt/embeddings/data.xlsx": b"old-workbook"})
        after = self.package("b.pptx", {**BASE, "ppt/embeddings/data.xlsx": b"new-workbook"})
        diff = compare(before, after)
        self.assertEqual(diff["removed"], ["ppt/charts/chart1.xml"])
        self.assertEqual([p["part"] for p in diff["changed"]], ["ppt/embeddings/data.xlsx"])
        self.assertFalse(diff["part_payloads_identical"])

    def test_nested_group_uses_own_id_and_keeps_parent_path(self):
        xml = f'<p:sld xmlns:p="{P}"><p:grpSp><p:nvGrpSpPr><p:cNvPr id="2"/></p:nvGrpSpPr><p:sp><p:nvSpPr><p:cNvPr id="3"/></p:nvSpPr></p:sp></p:grpSp></p:sld>'
        self.assertEqual(xml_inventory(xml.encode())["objects"], [
            {"kind": "grpSp", "id": "2", "parents": []},
            {"kind": "sp", "id": "3", "parents": ["2"]}])
        strict = xml.replace(P, "http://purl.oclc.org/ooxml/presentationml/main")
        self.assertEqual(xml_inventory(strict.encode())["objects"][1]["parents"], ["2"])

    def test_malformed_xml_marks_incomplete_without_losing_byte_inventory(self):
        report = self.package("a.pptx", {**BASE, "ppt/slides/slide1.xml": b"<broken"})
        self.assertFalse(report["inspection_complete"])
        self.assertIn("sha256", report["parts"]["ppt/slides/slide1.xml"])

    def test_non_pptx_and_duplicate_parts_are_rejected(self):
        with self.assertRaises(ValueError):
            self.package("other.zip", {"unrelated": b"data"})
        path = Path(self.temp.name) / "duplicate.pptx"
        with zipfile.ZipFile(path, "w") as archive:
            for key, value in BASE.items():
                archive.writestr(key, value)
            with self.assertWarns(UserWarning):
                archive.writestr("ppt/presentation.xml", b"<second/>")
        with self.assertRaises(ValueError):
            inspect(path)

    def test_dtd_is_rejected_in_utf8_and_utf16(self):
        xml = '<!DOCTYPE root [<!ENTITY example "content">]><root>&example;</root>'
        for encoding in ("utf-8", "utf-16"):
            with self.assertRaises(ValueError):
                xml_inventory(xml.encode(encoding))


if __name__ == "__main__":
    unittest.main()
