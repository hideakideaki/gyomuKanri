from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw


def main() -> int:
    source = Path(sys.argv[1])
    tag = sys.argv[2]
    output = Path(sys.argv[3])
    files = sorted(source.glob(f"{tag}_*.png"))
    thumb_width, thumb_height, label_height = 420, 250, 28
    columns = 2
    rows = (len(files) + columns - 1) // columns
    canvas = Image.new("RGB", (columns * thumb_width, rows * (thumb_height + label_height)), "white")
    draw = ImageDraw.Draw(canvas)
    for index, file in enumerate(files):
        image = Image.open(file).convert("RGB")
        image.thumbnail((thumb_width - 12, thumb_height - 12))
        left = (index % columns) * thumb_width
        top = (index // columns) * (thumb_height + label_height)
        canvas.paste(image, (left + 6, top + 6))
        draw.text((left + 6, top + thumb_height + 3), file.stem, fill="black")
    output.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
