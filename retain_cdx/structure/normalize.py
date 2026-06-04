from __future__ import annotations

from pathlib import Path
import tempfile


def normalize_cdx_viewbox(
    cdx_data: bytes,
    *,
    margin: float = 18,
    target_aspect: float | None = None,
    target_width: float | None = None,
    target_height: float | None = None,
    fill_ratio: float = 0.8,
) -> bytes:
    """Normalize CDX page bounds so the molecule fills its ChemDraw canvas.

    Some CDX files store a small fragment inside a much larger page. Excel can
    show a fitted preview PNG, but ChemDraw may re-render the editable OLE
    against the large internal page after the object is opened and closed. This
    helper shifts the fragment near the page origin and shrinks the document
    bounding boxes around the actual structure.
    """
    from pycdxml.cdxml_converter.chemdraw_io import read_cdx, write_cdx_file

    with tempfile.TemporaryDirectory(prefix="retain_cdx_normalize_") as tmp:
        tmpdir = Path(tmp)
        input_path = tmpdir / "input.cdx"
        output_path = tmpdir / "output.cdx"
        input_path.write_bytes(cdx_data)

        with input_path.open("rb") as f:
            doc = read_cdx(f)

        root = doc.cdxml.getroot()
        page = root.find("page")
        if page is None:
            return cdx_data

        fragments = page.findall("fragment")
        boxes = [_parse_box(frag.get("BoundingBox")) for frag in fragments]
        boxes = [box for box in boxes if box is not None]
        if not boxes:
            return cdx_data

        min_x = min(box[0] for box in boxes)
        min_y = min(box[1] for box in boxes)
        max_x = max(box[2] for box in boxes)
        max_y = max(box[3] for box in boxes)

        content_width = max_x - min_x
        content_height = max_y - min_y
        if content_width <= 0 or content_height <= 0:
            return cdx_data

        width = target_width or (content_width + 2 * margin)
        height = target_height or (content_height + 2 * margin)

        if target_aspect and target_aspect > 0 and width > 0 and height > 0:
            current_aspect = width / height
            if current_aspect < target_aspect:
                width = height * target_aspect
            elif current_aspect > target_aspect:
                height = width / target_aspect

        scale = min(
            (width * fill_ratio) / content_width,
            (height * fill_ratio) / content_height,
        )
        if scale <= 0:
            return cdx_data

        dx = ((width - content_width * scale) / 2) - min_x * scale
        dy = ((height - content_height * scale) / 2) - min_y * scale
        for element in page.iter():
            _transform_point_attr(element, "p", scale, dx, dy)
            _transform_box_attr(element, "BoundingBox", scale, dx, dy)

        box_text = f"0 0 {_fmt(width)} {_fmt(height)}"
        page.attrib["BoundingBox"] = box_text
        root.attrib["BoundingBox"] = box_text
        root.attrib["WindowSize"] = f"{_fmt(width)} {_fmt(height)}"

        write_cdx_file(doc, output_path)
        return output_path.read_bytes()


def _parse_nums(text: str | None) -> list[float] | None:
    if not text:
        return None
    try:
        return [float(part) for part in text.split()]
    except ValueError:
        return None


def _parse_box(text: str | None) -> tuple[float, float, float, float] | None:
    values = _parse_nums(text)
    if values is None or len(values) != 4:
        return None
    return values[0], values[1], values[2], values[3]


def _transform_point_attr(element, attr: str, scale: float, dx: float, dy: float) -> None:
    values = _parse_nums(element.attrib.get(attr))
    if values is None or len(values) != 2:
        return
    element.attrib[attr] = f"{_fmt(values[0] * scale + dx)} {_fmt(values[1] * scale + dy)}"


def _transform_box_attr(element, attr: str, scale: float, dx: float, dy: float) -> None:
    values = _parse_nums(element.attrib.get(attr))
    if values is None or len(values) != 4:
        return
    element.attrib[attr] = (
        f"{_fmt(values[0] * scale + dx)} {_fmt(values[1] * scale + dy)} "
        f"{_fmt(values[2] * scale + dx)} {_fmt(values[3] * scale + dy)}"
    )


def _fmt(value: float) -> str:
    return str(round(value, 2))
