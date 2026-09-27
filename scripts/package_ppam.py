"""Turn a macro-enabled PowerPoint source deck into a Ribbon add-in."""

from __future__ import annotations

import argparse
import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships"
RIBBON_REL = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
ADDIN_TYPE = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
ICONS = {
    "PPT2FigScopeIcon": "scope.png",
    "PPT2FigSelectedIcon": "selected.png",
    "PPT2FigShapesIcon": "shapes.png",
    "PPT2FigAllIcon": "all.png",
    "PPT2FigPNGIcon": "png.png",
    "PPT2FigJPGIcon": "jpg.png",
    "PPT2FigTIFFIcon": "tiff.png",
    "PPT2FigPDFIcon": "pdf.png",
    "PPT2FigEPSIcon": "eps.png",
    "PPT2FigSettingsIcon": "settings.png",
}


def package(source: Path, ribbon: Path, destination: Path) -> None:
    if destination.exists():
        raise FileExistsError(destination)
    with ZipFile(source) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}
    if "ppt/vbaProject.bin" not in members:
        raise ValueError("Source presentation has no VBA project")

    types = ET.fromstring(members["[Content_Types].xml"])
    main = next(
        (item for item in types.findall(f"{{{CONTENT_TYPES}}}Override")
         if item.attrib.get("PartName") == "/ppt/presentation.xml"),
        None,
    )
    if main is None:
        raise ValueError("Missing PowerPoint main part")
    main.set("ContentType", ADDIN_TYPE)
    members["[Content_Types].xml"] = ET.tostring(types, encoding="utf-8", xml_declaration=True)

    relations = ET.fromstring(members["_rels/.rels"])
    if any(item.attrib.get("Type") == RIBBON_REL for item in relations):
        raise ValueError("Source already has a Ribbon customization")
    identifiers = {item.attrib.get("Id") for item in relations}
    next_id = next(i for i in range(1, 1000) if f"rId{i}" not in identifiers)
    ET.SubElement(relations, f"{{{RELATIONSHIPS}}}Relationship", {
        "Id": f"rId{next_id}", "Type": RIBBON_REL,
        "Target": "customUI/customUI.xml",
    })
    members["_rels/.rels"] = ET.tostring(relations, encoding="utf-8", xml_declaration=True)
    members["customUI/customUI.xml"] = ribbon.read_bytes()
    icon_dir = ribbon.parent / "icons"
    icon_relations = ET.Element(f"{{{RELATIONSHIPS}}}Relationships")
    for relation_id, filename in ICONS.items():
        members[f"customUI/images/{filename}"] = (icon_dir / filename).read_bytes()
        ET.SubElement(icon_relations, f"{{{RELATIONSHIPS}}}Relationship", {
            "Id": relation_id, "Type": IMAGE_REL, "Target": f"images/{filename}",
        })
    members["customUI/_rels/customUI.xml.rels"] = ET.tostring(
        icon_relations, encoding="utf-8", xml_declaration=True,
    )
    if not any(item.attrib.get("Extension") == "png" for item in types.findall(f"{{{CONTENT_TYPES}}}Default")):
        ET.SubElement(types, f"{{{CONTENT_TYPES}}}Default", {
            "Extension": "png", "ContentType": "image/png",
        })
    members["[Content_Types].xml"] = ET.tostring(types, encoding="utf-8", xml_declaration=True)

    ET.fromstring(members["customUI/customUI.xml"])
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(destination, "w", ZIP_DEFLATED, compresslevel=6) as archive:
        for name, data in members.items():
            archive.writestr(name, data)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    parser.add_argument("--ribbon", type=Path, default=Path(__file__).resolve().parents[1] / "office/ribbon.xml")
    args = parser.parse_args()
    package(args.source, args.ribbon, args.destination)
    print(args.destination)


if __name__ == "__main__":
    main()
