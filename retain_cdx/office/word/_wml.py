"""WordprocessingML surgery: 把 ChemDraw OLE 对象放进段落。

Word 侧的落地模型（详见 `docs/word_ole_research.md`）：

  * 对象**随文字排版**：一个 ``w:r`` 里放一个 ``w:object``，没有锚点、没有坐标，
    位置就是它在文字流里的位置
  * 显示层只有 VML：``v:shape`` 的 style 给宽高（pt），``v:imagedata`` 指预览图
  * ``o:OLEObject`` 指容器；``w:dxaOrig``/``w:dyaOrig`` 是原始尺寸（twips）

一个对象 = ``w:object``（``v:shape`` + ``o:OLEObject``）+ document rels 两条
（oleObject / image）+ 两个部件（.bin / .emf|.png）。

三个必须文档内唯一的东西都在这里分配：VML shape id（``_x0000_i1025``）、
``ObjectID``（``_1853394678``）、以及 ``_x0000_t75`` 这个 shapetype 的定义
——Word 只在第一次用到时定义它，后面的对象只引用。
"""

from __future__ import annotations

import copy
import re
from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET

from retain_cdx.office._opc import (
    REL_IMAGE,
    REL_OLE,
    ensure_ole_content_types,
    main_part_path,
    next_rid,
    read_rels,
    rels_path_for,
    write_rels,
)

NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_V = "urn:schemas-microsoft-com:vml"
NS_O = "urn:schemas-microsoft-com:office:office"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_XML = "http://www.w3.org/XML/1998/namespace"

NS = {"w": NS_W, "r": NS_R, "v": NS_V, "o": NS_O}

SHAPETYPE_ID = "_x0000_t75"
EMU_PER_POINT = 12700
TWIPS_PER_POINT = 20

_W = f"{{{NS_W}}}"
_SHAPE_ID = re.compile(r"_x0000_[a-z](\d+)")
_OBJECT_ID = re.compile(r'ObjectID="_(\d+)"')
_ALIGNMENTS = ("left", "center", "right")


def _register_namespaces() -> None:
    for prefix, uri in (("w", NS_W), ("r", NS_R), ("v", NS_V), ("o", NS_O), ("w14", NS_W14)):
        ET.register_namespace(prefix, uri)


@dataclass(frozen=True)
class OleObjectTemplate:
    """A native Word+ChemDraw ``w:object`` (shapetype, shape, OLEObject)."""

    object_xml: str
    object_id: int

    @classmethod
    def from_package(cls, package_dir: Path) -> "OleObjectTemplate":
        document = main_part_path(package_dir)
        root = ET.parse(document).getroot()
        element = root.find(".//w:object", NS)
        if element is None or element.find("o:OLEObject", NS) is None:
            raise ValueError(f"No ChemDraw OLE object in {document}")
        ole = element.find("o:OLEObject", NS)
        match = re.fullmatch(r"_(\d+)", ole.get("ObjectID") or "")
        _register_namespaces()
        return cls(
            object_xml=ET.tostring(element, encoding="unicode"),
            object_id=int(match.group(1)) if match else 0,
        )


@dataclass(frozen=True)
class ObjectSize:
    """Displayed size of one object, in points."""

    width_pt: float
    height_pt: float

    @classmethod
    def from_emu(cls, width: int, height: int) -> "ObjectSize":
        return cls(width / EMU_PER_POINT, height / EMU_PER_POINT)


class WordDocument:
    """``document.xml`` held open across a batch of inserts.

    Parsing once per batch (rather than per object) keeps the id allocators in
    memory, so a batch can never hand the same shape id or ObjectID out twice.
    """

    def __init__(self, package_dir: Path):
        self.package_dir = package_dir
        self.part = main_part_path(package_dir)
        self._declarations = _root_namespace_declarations(self.part)
        self.tree = ET.parse(self.part)
        self.root = self.tree.getroot()
        body = self.root.find("w:body", NS)
        if body is None:
            raise ValueError(f"Document has no w:body: {self.part}")
        self.body = body
        self.rels_path = rels_path_for(self.part)
        self.rels = read_rels(self.rels_path)

        # Shape ids and ObjectIDs are document-wide, and headers/footers carry
        # VML shapes too — scan every WordprocessingML part, not just the body.
        texts = [
            path.read_text(encoding="utf-8", errors="replace")
            for path in self.part.parent.glob("*.xml")
        ]
        shape_ids = [int(m) for text in texts for m in _SHAPE_ID.findall(text)]
        object_ids = [int(m) for text in texts for m in _OBJECT_ID.findall(text)]
        self._next_shape = max(shape_ids, default=1024) + 1
        self._object_ids = set(object_ids)
        self._has_shapetype = any(
            shapetype.get("id") == SHAPETYPE_ID
            for shapetype in self.root.iter(f"{{{NS_V}}}shapetype")
        )

    # -- locating ----------------------------------------------------------

    def body_paragraphs(self) -> list[ET.Element]:
        """Top-level paragraphs, the same list python-docx calls ``paragraphs``."""
        return self.body.findall("w:p", NS)

    def append_paragraph(self, run: ET.Element, *, align: str | None) -> None:
        """Add a new paragraph holding *run* at the end of the body.

        The body's last child is normally ``w:sectPr`` (the final section's
        page setup), which must stay last — so the paragraph goes before it.
        """
        check_alignment(align)
        paragraph = ET.Element(f"{_W}p")
        if align is not None and align != "left":
            ppr = ET.SubElement(paragraph, f"{_W}pPr")
            ET.SubElement(ppr, f"{_W}jc", {f"{_W}val": align})
        paragraph.append(run)
        children = list(self.body)
        if children and children[-1].tag == f"{_W}sectPr":
            self.body.insert(len(children) - 1, paragraph)
        else:
            self.body.append(paragraph)

    def append_to_paragraph(self, index: int, run: ET.Element) -> None:
        paragraphs = self.body_paragraphs()
        if not 0 <= index < len(paragraphs):
            raise ValueError(
                f"Paragraph {index} out of range (document has {len(paragraphs)})"
            )
        paragraphs[index].append(run)

    def count_placeholder(self, placeholder: str) -> int:
        """How many times *placeholder* occurs, counted the way it is replaced."""
        if not placeholder:
            raise ValueError("placeholder must not be empty")
        return sum(
            "".join(t.text or "" for t, _, _ in _text_segments(paragraph)).count(placeholder)
            for paragraph in self.root.iter(f"{_W}p")
        )

    def replace_placeholder(self, placeholder: str, run: ET.Element) -> None:
        """Put *run* where the first occurrence of *placeholder* is.

        Word splits text into runs wherever formatting, spell-check or editing
        history changes, so ``{{mol}}`` typed by hand often spans several runs.
        The match is made against the paragraph's concatenated text and mapped
        back onto the runs.
        """
        if not placeholder:
            raise ValueError("placeholder must not be empty")
        for paragraph in self.root.iter(f"{_W}p"):
            segments = list(_text_segments(paragraph))
            text = "".join(t.text or "" for t, _, _ in segments)
            start = text.find(placeholder)
            if start != -1:
                _splice(segments, start, start + len(placeholder), run)
                return
        raise ValueError(f"Placeholder {placeholder!r} not found in the document")

    def text_width_twips(self) -> int:
        """Width between the margins of the last section, in twips."""
        sect = self.body.find("w:sectPr", NS)
        size = None if sect is None else sect.find("w:pgSz", NS)
        margins = None if sect is None else sect.find("w:pgMar", NS)
        page = int(size.get(f"{_W}w")) if size is not None and size.get(f"{_W}w") else 11906
        left = right = gutter = 1440
        if margins is not None:
            left = int(margins.get(f"{_W}left") or 1440)
            right = int(margins.get(f"{_W}right") or 1440)
            gutter = int(margins.get(f"{_W}gutter") or 0)
        else:
            gutter = 0
        return max(1, page - left - right - gutter)

    def insert_block(self, element: ET.Element, *, placeholder: str | None = None) -> None:
        """Put a block (table) at the end of the body, or where *placeholder* is.

        A placeholder must be the whole text of a top-level paragraph; that
        paragraph is replaced. A table is always followed by a paragraph —
        two tables in a row would merge, and Word wants one before sectPr.
        """
        if placeholder is None:
            children = list(self.body)
            index = len(children) - 1 if children and children[-1].tag == f"{_W}sectPr" else len(children)
            # A trailing empty paragraph (a blank document has nothing else)
            # becomes the paragraph after the table instead of a gap above it.
            if index > 0 and _is_empty_paragraph(children[index - 1]):
                index -= 1
        else:
            index = self.find_placeholder_paragraph(placeholder)
            self.body.remove(list(self.body)[index])
        self.body.insert(index, element)
        if element.tag == f"{_W}tbl":
            following = list(self.body)[index + 1 : index + 2]
            if not following or following[0].tag != f"{_W}p":
                self.body.insert(index + 1, ET.Element(f"{_W}p"))

    def find_placeholder_paragraph(self, placeholder: str) -> int:
        """Body index of the top-level paragraph whose whole text is *placeholder*."""
        for index, child in enumerate(self.body):
            if child.tag != f"{_W}p":
                continue
            text = "".join(t.text or "" for t, _, _ in _text_segments(child))
            if text.strip() == placeholder:
                return index
            if placeholder in text:
                raise ValueError(
                    f"Placeholder {placeholder!r} shares its paragraph with other "
                    "text; a table needs a paragraph of its own"
                )
        raise ValueError(f"Placeholder {placeholder!r} not found in the document")

    # -- building ----------------------------------------------------------

    def build_object_run(
        self,
        *,
        template: OleObjectTemplate,
        container: bytes,
        preview_bytes: bytes,
        preview_ext: str,
        size: ObjectSize,
    ) -> ET.Element:
        """Write the object's two parts and return the ``w:r`` that shows it."""
        word_dir = self.part.parent
        embeddings_dir = word_dir / "embeddings"
        media_dir = word_dir / "media"
        for directory in (embeddings_dir, media_dir):
            directory.mkdir(parents=True, exist_ok=True)

        bin_name = _numbered_part_name(embeddings_dir, "oleObject", "bin")
        (embeddings_dir / bin_name).write_bytes(container)
        image_name = _numbered_part_name(media_dir, "image", preview_ext)
        (media_dir / image_name).write_bytes(preview_bytes)

        image_rid = next_rid(self.rels)
        self.rels.append((image_rid, REL_IMAGE, f"media/{image_name}"))
        ole_rid = next_rid(self.rels)
        self.rels.append((ole_rid, REL_OLE, f"embeddings/{bin_name}"))
        ensure_ole_content_types(self.package_dir, preview_ext)

        element = ET.fromstring(template.object_xml)
        shape_id = f"_x0000_i{self._next_shape}"
        self._next_shape += 1
        object_id = self._allocate_object_id(template.object_id)

        # anchorId 是 Word 给每个对象的身份标记，克隆来的那份不能复用。
        element.attrib.pop(f"{{{NS_W14}}}anchorId", None)
        element.set(f"{_W}dxaOrig", str(round(size.width_pt * TWIPS_PER_POINT)))
        element.set(f"{_W}dyaOrig", str(round(size.height_pt * TWIPS_PER_POINT)))

        for shapetype in element.findall("v:shapetype", NS):
            if self._has_shapetype:
                element.remove(shapetype)
        self._has_shapetype = True

        shape = element.find("v:shape", NS)
        if shape is None:
            raise ValueError("Template object has no v:shape")
        shape.set("id", shape_id)
        shape.set("style", f"width:{_pt(size.width_pt)};height:{_pt(size.height_pt)}")
        imagedata = shape.find("v:imagedata", NS)
        if imagedata is not None:
            imagedata.set(f"{{{NS_R}}}id", image_rid)

        ole = element.find("o:OLEObject", NS)
        ole.set("ShapeID", shape_id)
        ole.set("ObjectID", f"_{object_id}")
        ole.set(f"{{{NS_R}}}id", ole_rid)

        run = ET.Element(f"{_W}r")
        run.append(element)
        return run

    def _allocate_object_id(self, base: int) -> int:
        candidate = max([base, *self._object_ids]) + 1
        while candidate in self._object_ids:
            candidate += 1
        self._object_ids.add(candidate)
        return candidate

    # -- saving ------------------------------------------------------------

    def save(self) -> None:
        write_rels(self.rels_path, self.rels)
        _write_document(self.tree, self.part, self._declarations)


@dataclass(frozen=True)
class TableCell:
    """One structure-and-label cell of a compound table."""

    run: ET.Element
    label: str


CELL_MARGIN_TWIPS = 108  # Word's own default left/right cell margin
_VALIGNS = ("top", "center", "bottom")


def build_compound_table(
    cells: list[TableCell],
    *,
    columns: int,
    column_twips: int,
    vertical_align: str = "bottom",
    label_space_after: int = 160,
) -> ET.Element:
    """A borderless table: per group a structure row, then a label row.

    Rows of a group are kept together across page breaks (``cantSplit`` on
    each row, ``keepNext`` on the structure row's paragraphs) so a number is
    never separated from its structure.
    """
    if vertical_align not in _VALIGNS:
        raise ValueError(f"vertical_align must be one of {_VALIGNS}, got {vertical_align!r}")

    table = ET.Element(f"{_W}tbl")
    tbl_pr = ET.SubElement(table, f"{_W}tblPr")
    _sub(tbl_pr, "tblW", w=str(column_twips * columns), type="dxa")
    _sub(tbl_pr, "jc", val="center")
    _sub(tbl_pr, "tblLayout", type="fixed")
    margins = _sub(tbl_pr, "tblCellMar")
    _sub(margins, "left", w=str(CELL_MARGIN_TWIPS), type="dxa")
    _sub(margins, "right", w=str(CELL_MARGIN_TWIPS), type="dxa")
    grid = ET.SubElement(table, f"{_W}tblGrid")
    for _ in range(columns):
        _sub(grid, "gridCol", w=str(column_twips))

    for start in range(0, len(cells), columns):
        group = cells[start : start + columns]
        structure_row = _table_row(table)
        label_row = _table_row(table)
        for column in range(columns):
            cell = group[column] if column < len(group) else None
            paragraph = _table_cell(structure_row, column_twips, vertical_align)
            ppr = paragraph.find("w:pPr", NS)
            ppr.insert(0, ET.Element(f"{_W}keepNext"))
            if cell is not None:
                paragraph.append(cell.run)

            # The gap between groups sits under the labels.
            paragraph = _table_cell(label_row, column_twips, "top", space_after=label_space_after)
            if cell is not None:
                run = ET.SubElement(paragraph, f"{_W}r")
                _set_text(ET.SubElement(run, f"{_W}t"), cell.label)
    return table


def _table_row(table: ET.Element) -> ET.Element:
    row = ET.SubElement(table, f"{_W}tr")
    tr_pr = ET.SubElement(row, f"{_W}trPr")
    ET.SubElement(tr_pr, f"{_W}cantSplit")
    return row


def _table_cell(
    row: ET.Element, width_twips: int, vertical_align: str, *, space_after: int = 0
) -> ET.Element:
    """Add a cell holding one centred paragraph; returns the paragraph."""
    cell = ET.SubElement(row, f"{_W}tc")
    tc_pr = ET.SubElement(cell, f"{_W}tcPr")
    _sub(tc_pr, "tcW", w=str(width_twips), type="dxa")
    _sub(tc_pr, "vAlign", val=vertical_align)
    paragraph = ET.SubElement(cell, f"{_W}p")
    ppr = ET.SubElement(paragraph, f"{_W}pPr")
    # Body text styles often indent the first line or snap lines to the
    # document grid (Chinese templates do both); either would push a centred
    # structure off centre or pad it vertically.
    _sub(ppr, "snapToGrid", val="0")
    _sub(ppr, "spacing", before="0", after=str(space_after))
    _sub(ppr, "ind", left="0", right="0", firstLine="0", firstLineChars="0")
    _sub(ppr, "jc", val="center")
    return paragraph


def _is_empty_paragraph(element: ET.Element) -> bool:
    return element.tag == f"{_W}p" and all(
        child.tag == f"{_W}pPr" for child in element
    )


def _sub(parent: ET.Element, tag: str, **attrs: str) -> ET.Element:
    return ET.SubElement(parent, f"{_W}{tag}", {f"{_W}{k}": v for k, v in attrs.items()})


def make_blank(package_dir: Path) -> None:
    """Turn an extracted template into an empty document.

    Keeps the styles, theme and page setup; drops the body content, the
    embedded objects with their relationships, and the author fields.
    """
    document = main_part_path(package_dir)
    tree = ET.parse(document)
    body = tree.getroot().find("w:body", NS)
    for child in list(body):
        if child.tag != f"{_W}sectPr":
            body.remove(child)
    body.insert(0, ET.Element(f"{_W}p"))
    _write_document(tree, document, _root_namespace_declarations(document))

    rels_path = rels_path_for(document)
    kept = []
    for rid, rel_type, target in read_rels(rels_path):
        if rel_type in (REL_OLE, REL_IMAGE):
            (document.parent / target).unlink(missing_ok=True)
        else:
            kept.append((rid, rel_type, target))
    write_rels(rels_path, kept)

    core = package_dir / "docProps" / "core.xml"
    if core.exists():
        text = core.read_text(encoding="utf-8")
        text = re.sub(r"(<dc:creator>)[^<]*(</dc:creator>)", r"\1\2", text)
        text = re.sub(r"(<cp:lastModifiedBy>)[^<]*(</cp:lastModifiedBy>)", r"\1\2", text)
        core.write_text(text, encoding="utf-8")


def check_alignment(align: str | None) -> None:
    if align is not None and align not in _ALIGNMENTS:
        raise ValueError(f"align must be one of {_ALIGNMENTS}, got {align!r}")


def _numbered_part_name(directory: Path, stem: str, extension: str) -> str:
    """``oleObject2.bin`` after ``oleObject1.bin`` — the way Word numbers parts."""
    index = 1
    while (directory / f"{stem}{index}.{extension}").exists():
        index += 1
    return f"{stem}{index}.{extension}"


def _pt(value: float) -> str:
    return f"{value:.2f}".rstrip("0").rstrip(".") + "pt"


# -- placeholder splicing ---------------------------------------------------


def _text_segments(paragraph: ET.Element):
    """Yield ``(w:t, run, run_parent)`` for the paragraph's own text, in order.

    Runs may sit directly in the paragraph or inside ``w:hyperlink`` /
    ``w:ins`` / ``w:smartTag``; nested paragraphs (text boxes) are their own
    paragraphs and are skipped here.
    """

    def walk(parent: ET.Element):
        for child in parent:
            if child.tag == f"{_W}p":
                continue
            if child.tag == f"{_W}r":
                for t in child.findall("w:t", NS):
                    yield t, child, parent
            else:
                yield from walk(child)

    yield from walk(paragraph)


def _splice(segments, start: int, end: int, run: ET.Element) -> None:
    """Replace text ``[start, end)`` spread over *segments* with *run*."""
    offsets = []
    position = 0
    for t, _, _ in segments:
        offsets.append(position)
        position += len(t.text or "")

    first = max(i for i, offset in enumerate(offsets) if offset <= start)
    last = max(i for i, offset in enumerate(offsets) if offset < end)
    first_t, first_run, first_parent = segments[first]
    last_t, _, _ = segments[last]
    head = (first_t.text or "")[: start - offsets[first]]
    tail = (last_t.text or "")[end - offsets[last] :]

    # Trim first, while every w:t still sits in the run we recorded for it.
    # A tail in a later run stays there, keeping that run's formatting.
    for t, owner, _ in segments[first + 1 : last]:
        owner.remove(t)
    _set_text(first_t, head)
    if last != first:
        _set_text(last_t, tail)

    # Split the first run right after its w:t so the object lands between the
    # text before the placeholder and whatever followed it.
    after = _split_run_after(first_run, first_t)
    if tail and last == first:
        if after is None:
            after = _empty_run_like(first_run)
        tail_t = ET.Element(f"{_W}t")
        _set_text(tail_t, tail)
        rpr = after.find("w:rPr", NS)
        after.insert(1 if rpr is not None else 0, tail_t)

    index = list(first_parent).index(first_run)
    first_parent.insert(index + 1, run)
    if after is not None:
        first_parent.insert(index + 2, after)


def _split_run_after(run: ET.Element, t: ET.Element) -> ET.Element | None:
    children = list(run)
    following = children[children.index(t) + 1 :]
    if not following:
        return None
    new_run = _empty_run_like(run)
    for child in following:
        run.remove(child)
        new_run.append(child)
    return new_run


def _empty_run_like(run: ET.Element) -> ET.Element:
    """A run with *run*'s formatting (``w:rPr``) and no content."""
    new_run = ET.Element(f"{_W}r")
    rpr = run.find("w:rPr", NS)
    if rpr is not None:
        new_run.append(copy.deepcopy(rpr))
    return new_run


def _set_text(t: ET.Element, text: str) -> None:
    t.text = text
    # Without xml:space="preserve" Word trims the spaces next to the object.
    t.set(f"{{{NS_XML}}}space", "preserve")


# -- serialization -----------------------------------------------------------


def _root_namespace_declarations(path: Path) -> dict[str, str]:
    """``prefix → uri`` for every ``xmlns:*`` on the root element."""
    text = path.read_text(encoding="utf-8")
    start = text.find("<w:document")
    end = text.find(">", start)
    if start == -1 or end == -1:
        return {}
    return dict(re.findall(r'xmlns:([\w.-]+)="([^"]*)"', text[start:end]))


def _write_document(tree: ET.ElementTree, path: Path, declarations: dict[str, str]) -> None:
    """Serialize ``document.xml`` without losing the namespaces MCE names.

    Word's root declares two dozen prefixes (``w15``, ``w16cid``…) purely so
    ``mc:Ignorable`` can list them; no element uses most of them, so
    ElementTree drops their declarations and leaves ``Ignorable`` naming
    undeclared prefixes — the same defect that made PowerPoint and Excel
    offer repairs. The original declarations are put back on the root.
    """
    _register_namespaces()
    for prefix, uri in declarations.items():
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:  # ElementTree reserves ns0, ns1…
            pass
    tree.write(path, encoding="UTF-8", xml_declaration=True)

    text = path.read_text(encoding="utf-8")
    start = text.find("<w:document")
    end = text.find(">", start)
    if start == -1 or end == -1:
        return
    root_tag = text[start:end]
    missing = "".join(
        f' xmlns:{prefix}="{uri}"'
        for prefix, uri in declarations.items()
        if f"xmlns:{prefix}=" not in root_tag
    )
    if missing:
        path.write_text(text[:end] + missing + text[end:], encoding="utf-8")


__all__ = [
    "TableCell",
    "build_compound_table",
    "check_alignment",
    "make_blank",
    "ObjectSize",
    "OleObjectTemplate",
    "WordDocument",
]
