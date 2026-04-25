from __future__ import annotations

from docx.oxml.ns import qn


def _set_rfonts(run_properties, ascii_font: str, east_asia_font: str) -> None:
    fonts = run_properties.get_or_add_rFonts()
    fonts.set(qn("w:ascii"), ascii_font)
    fonts.set(qn("w:hAnsi"), ascii_font)
    fonts.set(qn("w:eastAsia"), east_asia_font)


def set_run_mixed_fonts(
    run,
    ascii_font: str = "Times New Roman",
    east_asia_font: str = "宋体",
) -> None:
    """Set western and East Asian fonts on a python-docx run."""
    run.font.name = ascii_font
    _set_rfonts(run._element.get_or_add_rPr(), ascii_font, east_asia_font)


def set_style_mixed_fonts(
    style,
    ascii_font: str = "Times New Roman",
    east_asia_font: str = "宋体",
) -> None:
    """Set western and East Asian fonts on a python-docx paragraph style."""
    style.font.name = ascii_font
    _set_rfonts(style._element.get_or_add_rPr(), ascii_font, east_asia_font)
