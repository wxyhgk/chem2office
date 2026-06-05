from __future__ import annotations

from dataclasses import dataclass

from retain_cdx.office.excel._utils import EMU_PER_PIXEL


@dataclass(frozen=True)
class ShellSize:
    width_px: int
    height_px: int

    def fit(
        self,
        *,
        width_px: int | None,
        height_px: int | None,
        preserve_aspect: bool,
    ) -> "ShellSize":
        if width_px is None and height_px is None:
            return self
        if not preserve_aspect:
            return ShellSize(width_px or self.width_px, height_px or self.height_px)

        aspect = self.width_px / self.height_px
        if width_px is not None and height_px is not None:
            scale = min(width_px / self.width_px, height_px / self.height_px)
            return ShellSize(
                max(1, round(self.width_px * scale)),
                max(1, round(self.height_px * scale)),
            )
        if width_px is not None:
            return ShellSize(width_px, max(1, round(width_px / aspect)))
        height = height_px or self.height_px
        return ShellSize(max(1, round(height * aspect)), height)


@dataclass(frozen=True)
class AnchorCorner:
    col: int
    col_off_px: int
    row: int
    row_off_px: int

    def to_xdr_xml(self) -> str:
        return (
            f"<xdr:col>{self.col}</xdr:col>"
            f"<xdr:colOff>{self.col_off_px * EMU_PER_PIXEL}</xdr:colOff>"
            f"<xdr:row>{self.row}</xdr:row>"
            f"<xdr:rowOff>{self.row_off_px * EMU_PER_PIXEL}</xdr:rowOff>"
        )


@dataclass(frozen=True)
class TwoCellAnchor:
    from_: AnchorCorner
    to: AnchorCorner

    def to_vml_anchor(self) -> str:
        return (
            f"{self.from_.col}, {self.from_.col_off_px}, "
            f"{self.from_.row}, {self.from_.row_off_px}, "
            f"{self.to.col}, {self.to.col_off_px}, "
            f"{self.to.row}, {self.to.row_off_px}"
        )
