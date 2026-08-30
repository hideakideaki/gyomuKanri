from __future__ import annotations

from .models import RichTextRun


def ole_color_to_hex(color: int) -> str:
    color = int(color) & 0xFFFFFF
    red = color & 0xFF
    green = (color >> 8) & 0xFF
    blue = (color >> 16) & 0xFF
    return f"#{red:02X}{green:02X}{blue:02X}"


def hex_to_ole_color(color: str) -> int:
    value = color.lstrip("#")
    if len(value) != 6:
        raise ValueError(f"色は#RRGGBB形式で指定してください: {color}")
    red, green, blue = int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)
    return red | (green << 8) | (blue << 16)


def compress_runs(styles: list[tuple[str, bool, bool, bool]]) -> list[RichTextRun]:
    if not styles:
        return []
    runs: list[RichTextRun] = []
    start = 1
    previous = styles[0]
    for index, style in enumerate(styles[1:], start=2):
        if style != previous:
            runs.append(RichTextRun(start=start, length=index - start, font_color=previous[0], bold=previous[1], italic=previous[2], underline=previous[3]))
            start = index
            previous = style
    runs.append(RichTextRun(start=start, length=len(styles) - start + 1, font_color=previous[0], bold=previous[1], italic=previous[2], underline=previous[3]))
    return runs
