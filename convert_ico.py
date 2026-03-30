from PIL import Image
import struct
import io
import os

def buat_ico_manual(input_png, output_ico):
    """
    Buat .ico dengan ukuran besar (hingga 1024x1024)
    dengan menyimpan PNG langsung di dalam .ico (format modern Windows Vista+)
    """

    ukuran_list = [16, 24, 32, 48, 64, 128, 256, 512, 1024]

    img_asli = Image.open(input_png).convert("RGBA")

    entries = []  # (width, height, png_bytes)

    for size in ukuran_list:
        img_resize = img_asli.resize((size, size), Image.LANCZOS)

        buffer = io.BytesIO()
        img_resize.save(buffer, format="PNG")
        png_bytes = buffer.getvalue()

        entries.append((size, size, png_bytes))

    # ── Tulis file .ico secara manual ──
    num_images = len(entries)

    # Header ICO: 6 bytes
    # Directory: 16 bytes per image
    header_size = 6 + 16 * num_images

    ico_data = b""

    # ICO Header
    ico_data += struct.pack("<HHH", 0, 1, num_images)

    # Hitung offset awal data image
    offset = header_size

    # Directory entries
    for (w, h, png_bytes) in entries:
        size_byte = 0 if w >= 256 else w   # 0 = 256 atau lebih (konvensi ICO)
        ico_data += struct.pack(
            "<BBBBHHII",
            size_byte,   # width  (0 = 256+)
            size_byte,   # height (0 = 256+)
            0,           # color count (0 = tidak pakai palette)
            0,           # reserved
            1,           # color planes
            32,          # bits per pixel
            len(png_bytes),
            offset
        )
        offset += len(png_bytes)

    # Image data
    for (_, _, png_bytes) in entries:
        ico_data += png_bytes

    with open(output_ico, "wb") as f:
        f.write(ico_data)

    print(f"✅ Berhasil! {output_ico} dibuat dengan {num_images} ukuran:")
    for (w, h, png_bytes) in entries:
        print(f"   {w}x{h}  →  {len(png_bytes):,} bytes")
    print(f"   Total ukuran file: {len(ico_data):,} bytes")


# ── Jalankan ──
buat_ico_manual("logo_app.png", "logo_app.ico")