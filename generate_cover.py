#!/usr/bin/env python3
"""Generate 1280×720 cover image for BRRRR Deal Analyzer Gumroad listing."""

from PIL import Image, ImageDraw, ImageFont
import os

W, H = 1280, 720
OUT = "output/cover.png"

# Create image with solid dark navy
img = Image.new("RGB", (W, H), color="#1B2A4A")
draw = ImageDraw.Draw(img)

# Gradient overlay (manual per-pixel)
for y in range(H):
    r = int(27 + (46 - 27) * y / H)
    g = int(42 + (74 - 42) * y / H)
    b = int(74 + (120 - 74) * y / H)
    for x in range(W):
        img.putpixel((x, y), (r, g, b))

# Re-draw for clarity after putpixel
draw = ImageDraw.Draw(img)

# Accent stripe at bottom
draw.rectangle([0, H - 100, W, H], fill="#27AE60")

# Fonts
bold_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 52)
sub_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 26)
tag_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 20)
feat_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 18)
price_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 32)
small_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 14)

# Header bar
draw.rectangle([0, 0, W, 110], fill="#0F1A30")

# Title
title = "BRRRR Deal Analyzer"
bbox = draw.textbbox((0, 0), title, font=bold_font)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) // 2, 30), title, fill="white", font=bold_font)

# Subtitle
sub = "Analyze BRRRR, Flip & Buy-and-Hold Deals in 60 Seconds"
bbox = draw.textbbox((0, 0), sub, font=sub_font)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) // 2, 110), sub, fill="#D6E4F0", font=sub_font)

# Feature boxes (3 columns x 2 rows)
features = [
    "💰 BRRRR Cash-Out",
    "🔨 Flip Analyzer",
    "📈 30-Year Projection",
    "🎯 Auto Deal Scoring",
    "🔄 Refi Scenarios",
    "✅ Buy Box Rubric",
]

box_w, box_h = 340, 60
gap = 30
start_x = (W - (box_w * 3 + gap * 2)) // 2
start_y = 200

for i, feat in enumerate(features):
    col = i % 3
    row = i // 3
    x = start_x + col * (box_w + gap)
    y = start_y + row * (box_h + 20)
    draw.rounded_rectangle([x, y, x + box_w, y + box_h], radius=10, fill="#27AE60", outline="white", width=2)
    bbox = draw.textbbox((0, 0), feat, font=feat_font)
    fw = bbox[2] - bbox[0]
    fh = bbox[3] - bbox[1]
    draw.text((x + (box_w - fw) // 2, y + (box_h - fh) // 2), feat, fill="white", font=feat_font)

# Key formula banner
formula = "ARV × Refi LTV − Purchase × 1.06 − Rehab − Holding × 6mo = Cash Pulled Out"
form_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
draw.rounded_rectangle([60, 440, W - 60, 500], radius=12, fill="#F39C12")
bbox = draw.textbbox((0, 0), formula, font=form_font)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) // 2, 452), formula, fill="#1B2A4A", font=form_font)

# USP text
usp = "★ The refinance seasoning math most competitors miss"
bbox = draw.textbbox((0, 0), usp, font=sub_font)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) // 2, 530), usp, fill="#F1C40F", font=sub_font)

# Price tag
price = "$19"
draw.rounded_rectangle([W//2 - 90, H - 155, W//2 + 90, H - 100], radius=35, fill="white")
bbox = draw.textbbox((0, 0), price, font=price_font)
pw = bbox[2] - bbox[0]
draw.text((W//2 - pw//2, H - 152), price, fill="#1B2A4A", font=price_font)

# Bottom tagline
tag = "Google Sheet Template | Instant Deal Analysis | Investor-Proven"
bbox = draw.textbbox((0, 0), tag, font=tag_font)
tw = bbox[2] - bbox[0]
draw.text(((W - tw) // 2, H - 72), tag, fill="white", font=tag_font)

img.save(OUT)
print(f"✅ Cover image saved: {OUT} ({os.path.getsize(OUT)} bytes)")
