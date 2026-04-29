"""Prepare logo assets: transparent PNG + tightly cropped versions for the deck."""
from PIL import Image, ImageChops
import os

src = r"d:\Hanova\Hanova Logo.jpeg"
out_dir = r"d:\Hanova\branding"
os.makedirs(out_dir, exist_ok=True)

im = Image.open(src).convert("RGBA")

# Make near-white background transparent
datas = im.getdata()
new = []
for px in datas:
    r, g, b, a = px
    # treat near-white as transparent
    if r > 240 and g > 240 and b > 240:
        new.append((255, 255, 255, 0))
    else:
        new.append((r, g, b, 255))
im.putdata(new)

# Trim transparent borders
bbox = im.getbbox()
if bbox:
    im = im.crop(bbox)

# Save full transparent
full = os.path.join(out_dir, "hanova-logo-transparent.png")
im.save(full, "PNG")

print("Saved:", full, "size:", im.size)

# ---- Crop just the H mark (top portion, above the wordmark) ----
# The mark sits in roughly the top ~50% of the cropped image.
W, H = im.size
mark = im.crop((int(W * 0.20), 0, int(W * 0.80), int(H * 0.52)))
# trim again
bbox2 = mark.getbbox()
if bbox2:
    mark = mark.crop(bbox2)
mark_path = os.path.join(out_dir, "hanova-mark.png")
mark.save(mark_path, "PNG")
print("Saved:", mark_path, "size:", mark.size)

# ---- Light-on-dark version: composite logo onto a white rounded card ----
# Simpler: a white-tinted version where dark navy becomes white, teal stays.
inv = Image.new("RGBA", im.size, (0, 0, 0, 0))
px_src = im.load()
px_dst = inv.load()
for y in range(im.size[1]):
    for x in range(im.size[0]):
        r, g, b, a = px_src[x, y]
        if a == 0:
            continue
        # If pixel is dark navy-ish, recolor to white. Else keep (teal/orange).
        if r < 90 and g < 90 and b < 130:
            px_dst[x, y] = (255, 255, 255, a)
        else:
            px_dst[x, y] = (r, g, b, a)
inv_path = os.path.join(out_dir, "hanova-logo-onDark.png")
inv.save(inv_path, "PNG")
print("Saved:", inv_path, "size:", inv.size)

# Mark on dark
mark_inv = inv.crop((int(W * 0.20), 0, int(W * 0.80), int(H * 0.52)))
b3 = mark_inv.getbbox()
if b3:
    mark_inv = mark_inv.crop(b3)
mark_inv_path = os.path.join(out_dir, "hanova-mark-onDark.png")
mark_inv.save(mark_inv_path, "PNG")
print("Saved:", mark_inv_path, "size:", mark_inv.size)
