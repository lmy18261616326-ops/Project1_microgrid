from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

pages_dir = Path(r"D:\PV_MPPT\PV_ESS_MPC_Paper_Reproduction\beginner_from_zero\docs\qa_fallback_v2\pages")
out_dir = pages_dir.parent / "contact_sheets"
out_dir.mkdir(parents=True, exist_ok=True)
pages = sorted(pages_dir.glob("page-*.png"))
font_path = Path(r"C:\Windows\Fonts\arial.ttf")
font = ImageFont.truetype(str(font_path), 22) if font_path.exists() else ImageFont.load_default()

for sheet_no, start in enumerate(range(0, len(pages), 4), 1):
    group = pages[start:start + 4]
    thumbs = []
    for path in group:
        image = Image.open(path).convert("RGB")
        image.thumbnail((700, 900))
        thumbs.append((path, image.copy()))
    sheet = Image.new("RGB", (1460, 1900), "#d9dde3")
    draw = ImageDraw.Draw(sheet)
    for idx, (path, image) in enumerate(thumbs):
        x = 20 + (idx % 2) * 720
        y = 45 + (idx // 2) * 930
        sheet.paste(image, (x + (700-image.width)//2, y))
        draw.text((x, 12 + (idx // 2) * 930), path.stem, fill="black", font=font)
    sheet.save(out_dir / f"sheet_{sheet_no:02d}.jpg", quality=88)

print(len(pages), len(list(out_dir.glob('sheet_*.jpg'))))
