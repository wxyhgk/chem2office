from __future__ import annotations

from pathlib import Path
import tempfile


def cdx_page_aspect(cdx_data: bytes) -> float | None:
    from pycdxml.cdxml_converter.chemdraw_io import read_cdx

    with tempfile.TemporaryDirectory(prefix="retain_cdx_geometry_") as tmp:
        path = Path(tmp) / "input.cdx"
        path.write_bytes(cdx_data)
        with path.open("rb") as f:
            doc = read_cdx(f)

    page = doc.cdxml.getroot().find("page")
    if page is None:
        return None
    box = _parse_box(page.get("BoundingBox"))
    if box is None:
        return None
    width = box[2] - box[0]
    height = box[3] - box[1]
    if width <= 0 or height <= 0:
        return None
    return width / height


def cdx_content_aspect(cdx_data: bytes) -> float | None:
    from pycdxml.cdxml_converter.chemdraw_io import read_cdx

    with tempfile.TemporaryDirectory(prefix="retain_cdx_geometry_") as tmp:
        path = Path(tmp) / "input.cdx"
        path.write_bytes(cdx_data)
        with path.open("rb") as f:
            doc = read_cdx(f)

    page = doc.cdxml.getroot().find("page")
    if page is None:
        return None
    boxes = [_parse_box(fragment.get("BoundingBox")) for fragment in page.findall("fragment")]
    boxes = [box for box in boxes if box is not None]
    if not boxes:
        # Not every producer writes fragment boxes (Structura's CDXML doesn't);
        # the atom positions bound the drawing just as well.
        points = [_parse_point(node.get("p")) for node in page.iter("n")]
        points = [point for point in points if point is not None]
        if not points:
            return None
        xs = [x for x, _ in points]
        ys = [y for _, y in points]
        boxes = [(min(xs), min(ys), max(xs), max(ys))]
    min_x = min(box[0] for box in boxes)
    min_y = min(box[1] for box in boxes)
    max_x = max(box[2] for box in boxes)
    max_y = max(box[3] for box in boxes)
    width = max_x - min_x
    height = max_y - min_y
    if width <= 0 or height <= 0:
        return None
    return width / height


def _parse_point(text: str | None) -> tuple[float, float] | None:
    if not text:
        return None
    try:
        values = [float(part) for part in text.split()]
    except ValueError:
        return None
    return (values[0], values[1]) if len(values) >= 2 else None


def _parse_box(text: str | None) -> tuple[float, float, float, float] | None:
    if not text:
        return None
    try:
        values = [float(part) for part in text.split()]
    except ValueError:
        return None
    if len(values) != 4:
        return None
    return values[0], values[1], values[2], values[3]
