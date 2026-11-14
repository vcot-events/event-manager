from zipfile import ZipFile
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def _xml(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def _element_text(element) -> str:
    """Extract text while preserving tabs, breaks, and paragraph edges."""
    parts = []
    for node in element.iter():
        tag = node.tag
        if tag == _xml(W_NS, "t"):
            parts.append(node.text or "")
        elif tag in (_xml(W_NS, "tab"), _xml(W_NS, "br")):
            parts.append("\t" if tag.endswith("tab") else "\n")
    return "".join(parts)


def docx_body_blocks(path) -> list[dict]:
    """Return a flat list of paragraph/table blocks in Word document order."""
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    body = root.find(_xml(W_NS, "body"))
    blocks = []
    for child in body:
        tag = child.tag
        if tag == _xml(W_NS, "p"):
            text = _element_text(child).strip()
            if text:
                blocks.append({"type": "paragraph", "text": text})
        elif tag == _xml(W_NS, "tbl"):
            rows = []
            for row in child.findall(_xml(W_NS, "tr")):
                cells = [_element_text(cell).strip() for cell in row.findall(_xml(W_NS, "tc"))]
                rows.append(cells)
            blocks.append({"type": "table", "rows": rows})
    return blocks
