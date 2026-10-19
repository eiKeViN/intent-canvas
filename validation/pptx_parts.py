#!/usr/bin/env python3
"""Read-only PPTX package inventory/diff. This is not a fidelity validator."""

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile


MAX_PART_BYTES = 64 * 1024 * 1024
MAX_PACKAGE_BYTES = 512 * 1024 * 1024
MAX_PARTS = 20000
PML = {
    "http://schemas.openxmlformats.org/presentationml/2006/main",
    "http://purl.oclc.org/ooxml/presentationml/main",
}


def split_tag(tag):
    if tag.startswith("{"):
        return tag[1:].split("}", 1)
    return "", tag


def xml_inventory(data):
    # No DTDs are required for this inspection; do not expand entity content.
    class NoDTD(ET.TreeBuilder):
        def doctype(self, name, public_id, system_id):
            raise ValueError("DTD declarations are outside this probe's scope")

    root = ET.fromstring(data, parser=ET.XMLParser(target=NoDTD()))
    tags = Counter(element.tag for element in root.iter())
    objects = []

    def visit(element, parents):
        uri, name = split_tag(element.tag)
        next_parents = parents
        if uri in PML and name in {"sp", "grpSp", "pic", "graphicFrame", "cxnSp"}:
            # cNvPr belongs to this object's non-visual property container,
            # never to a descendant group child.
            props = next((p for child in element for p in child
                          if p.tag == "{" + uri + "}cNvPr"), None)
            object_id = props.get("id") if props is not None else None
            objects.append({"kind": name, "id": object_id, "parents": parents})
            next_parents = parents + [object_id]
        for child in element:
            visit(child, next_parents)

    visit(root, [])
    return {"root": root.tag, "elements": dict(sorted(tags.items())),
            "objects": objects}


def inspect(path):
    parts = {}
    warnings = []
    with zipfile.ZipFile(path) as archive:
        entries = [entry for entry in archive.infolist() if not entry.is_dir()]
        names = [entry.filename for entry in entries]
        if len(names) != len(set(names)):
            raise ValueError("Duplicate ZIP part names are ambiguous")
        if len(entries) > MAX_PARTS:
            raise ValueError("Package exceeds the probe's part-count limit")
        if sum(entry.file_size for entry in entries) > MAX_PACKAGE_BYTES:
            raise ValueError("Package exceeds the probe's uncompressed size limit")
        required = {"[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml"}
        if not required.issubset(names):
            raise ValueError("Expected an unencrypted PPTX with standard part paths")
        for entry in sorted(entries, key=lambda item: item.filename):
            if entry.file_size > MAX_PART_BYTES:
                raise ValueError("Part exceeds the probe's size limit: " + entry.filename)
            data = archive.read(entry)
            item = {"bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}
            if entry.filename.endswith((".xml", ".rels")):
                try:
                    item["xml"] = xml_inventory(data)
                except (ET.ParseError, ValueError, RecursionError) as error:
                    item["xml_error"] = str(error)
                    warnings.append(entry.filename + ": XML inspection incomplete")
            parts[entry.filename] = item
    return {"schema": "intent-canvas-pptx-parts/1", "parts": parts,
            "inspection_complete": not warnings, "warnings": warnings,
            "limits": "Part bytes and selected XML structure only; no rendering, "
                      "relationship resolution, embedded workbook parsing or editability proof."}


def compare(before, after):
    left, right = before["parts"], after["parts"]
    changed = sorted(name for name in left.keys() & right.keys()
                     if left[name]["sha256"] != right[name]["sha256"])
    added = sorted(right.keys() - left.keys())
    removed = sorted(left.keys() - right.keys())
    return {
        "schema": "intent-canvas-pptx-diff/1",
        "part_payloads_identical": not (changed or added or removed),
        "inspection_complete": before["inspection_complete"] and after["inspection_complete"],
        "added": added, "removed": removed,
        "changed": [{"part": name, "before": left[name], "after": right[name]}
                    for name in changed],
        "warnings": {"before": before["warnings"], "after": after["warnings"]},
        "interpretation": "ZIP ordering, timestamps and compression are ignored. "
                          "Changed XML bytes may be harmless serialization changes. "
                          "Identical payloads do not prove host API editability or visual fidelity.",
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path, nargs="?")
    parser.add_argument("--output", type=Path, help="Create a new JSON file; never overwrite")
    args = parser.parse_args()
    try:
        result = inspect(args.before)
        if args.after:
            result = compare(result, inspect(args.after))
        rendered = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
        if args.output:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            with args.output.open("x", encoding="utf-8") as handle:
                handle.write(rendered)
        else:
            print(rendered, end="")
        # A detected difference is evidence to review, not a failed compatibility test.
        return 0
    except (OSError, ValueError, zipfile.BadZipFile, RuntimeError) as error:
        print("Inspection failed: " + str(error), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
