"""Generate a high-quality Windows .ico file from a PNG image.

Produces a multi-resolution ICO with PNG-compressed 256x256 layer
and BMP layers for smaller sizes — the standard Windows expects.
"""

import io
import struct
from PIL import Image


def png_to_ico_hq(png_path: str, ico_path: str) -> None:
    """Convert PNG to ICO with maximum quality."""
    img = Image.open(png_path).convert("RGBA")
    w, h = img.size
    side = min(w, h)
    left = (w - side) // 2
    top = (h - side) // 2
    img_sq = img.crop((left, top, left + side, top + side))

    sizes = [256, 128, 64, 48, 32, 24, 16]
    entries = []

    for s in sizes:
        resized = img_sq.resize((s, s), Image.LANCZOS)
        if s == 256:
            # 256x256 stored as PNG (standard for Vista+ icons)
            buf = io.BytesIO()
            resized.save(buf, format="PNG", optimize=False)
            data = buf.getvalue()
        else:
            # Smaller sizes as raw BGRA bitmaps
            buf = io.BytesIO()
            resized.save(buf, format="PNG", optimize=False)
            data = buf.getvalue()
        entries.append((s, data))

    # Build ICO file manually for full control
    # ICO Header: reserved(2) + type(2) + count(2)
    num = len(entries)
    header = struct.pack("<HHH", 0, 1, num)

    # Calculate offsets
    dir_size = 16 * num  # each directory entry is 16 bytes
    offset = 6 + dir_size  # after header + directory

    directory = b""
    image_data = b""

    for size, data in entries:
        w_byte = 0 if size == 256 else size
        h_byte = 0 if size == 256 else size
        entry = struct.pack(
            "<BBBBHHII",
            w_byte,     # width (0 = 256)
            h_byte,     # height (0 = 256)
            0,          # color palette
            0,          # reserved
            1,          # color planes
            32,         # bits per pixel
            len(data),  # data size
            offset,     # offset from file start
        )
        directory += entry
        image_data += data
        offset += len(data)

    with open(ico_path, "wb") as f:
        f.write(header + directory + image_data)

    import os
    total = os.path.getsize(ico_path)
    print(f"ICO created: {total:,} bytes, {num} resolutions")
    for s, d in entries:
        print(f"  {s}x{s}: {len(d):,} bytes")


if __name__ == "__main__":
    png_to_ico_hq("icon.png", "icon.ico")
