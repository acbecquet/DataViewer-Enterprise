"""
Image and PDF Translator
Translates visible text in JPG, PNG, and PDF files using Claude's vision API
(for raster content) and PyMuPDF (for selectable text in PDFs).

Approach
--------
PNG / JPG
  Claude's vision API identifies every text region, returns the translation and
  an approximate bounding-box (as % of image dimensions).  Pillow then paints a
  background-colored rectangle over each original-text region and draws the
  translated text in its place.

PDF – text-based (selectable text)
  PyMuPDF extracts each text block with its exact rectangle and font info.
  The original text is whited-out with a filled rectangle, then the translated
  text is inserted at the same location using the nearest available font.

PDF – image-based (scanned / no selectable text)
  Each page is rasterised to a high-resolution PNG, processed with the same
  vision pipeline as PNG/JPG files, then compiled back into a new PDF.
"""

import os
import sys
import time
import re
import json
import base64
import textwrap
from pathlib import Path

try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import fitz          # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

from anthropic import Anthropic

# ── Font resolution ────────────────────────────────────────────────────────────

# Fonts that cover Latin + CJK (Microsoft YaHei ships with Windows and handles
# Chinese, Japanese, Korean, plus all Latin scripts).
_FONT_CANDIDATES = {
    'cjk':    ['C:/Windows/Fonts/msyh.ttc',
               'C:/Windows/Fonts/simsun.ttc',
               'C:/Windows/Fonts/malgun.ttf'],
    'latin':  ['C:/Windows/Fonts/calibri.ttf',
               'C:/Windows/Fonts/arial.ttf',
               'C:/Windows/Fonts/msyh.ttc'],   # msyh also handles Latin well
    'bold':   ['C:/Windows/Fonts/calibrib.ttf',
               'C:/Windows/Fonts/arialbd.ttf'],
}

CJK_LANGUAGES = {"Chinese (Simplified)", "Chinese (Traditional)", "Japanese", "Korean"}

# PyMuPDF built-in font names used when inserting translated text into PDFs.
_PDF_FONT_LATIN = "helv"        # Helvetica – good Latin coverage
_PDF_FONT_CJK   = "china-s"    # Simplified Chinese (PyMuPDF built-in)
_PDF_FONT_CJK_T = "china-t"    # Traditional Chinese
_PDF_FONT_JP    = "japan"       # Japanese
_PDF_FONT_KO    = "korea"       # Korean


def _find_font(lang_key: str) -> str | None:
    """Return the first existing font file path for the given language key."""
    for path in _FONT_CANDIDATES.get(lang_key, []):
        if os.path.exists(path):
            return path
    return None


def _get_pil_font(target_language: str, size: int) -> ImageFont.FreeTypeFont | None:
    """Return a PIL ImageFont appropriate for the target language."""
    lang_key = 'cjk' if target_language in CJK_LANGUAGES else 'latin'
    path = _find_font(lang_key)
    if path:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    # Fallback: try CJK font (handles everything)
    cjk = _find_font('cjk')
    if cjk:
        try:
            return ImageFont.truetype(cjk, size)
        except Exception:
            pass
    return None


def _pdf_fontname(target_language: str) -> str:
    """Return a PyMuPDF built-in font name appropriate for the target language."""
    if target_language == "Chinese (Simplified)":
        return _PDF_FONT_CJK
    elif target_language == "Chinese (Traditional)":
        return _PDF_FONT_CJK_T
    elif target_language == "Japanese":
        return _PDF_FONT_JP
    elif target_language == "Korean":
        return _PDF_FONT_KO
    else:
        return _PDF_FONT_LATIN


# ── Shared translation helpers ─────────────────────────────────────────────────

def needs_translation(text: str, source_language: str, target_language: str) -> bool:
    if source_language == target_language:
        return False
    stripped = text.strip()
    if not stripped:
        return False
    return any(
        c.isalpha() or '\u4e00' <= c <= '\u9fff' or '\u3040' <= c <= '\u30ff'
        for c in stripped
    )


def translate_text(text: str, client: Anthropic, source_language: str,
                   target_language: str, max_retries: int = 2) -> str:
    context = (
        "You are translating internal business documentation from a product development company.\n"
        "This documentation contains technical specifications, sample requirements, and testing procedures.\n"
        "Note: SDR refers to Spectrum Dynamic Research, which is the testing department.\n"
        "HC refers to the R&D department.\n\n"
        f"Please translate the following {source_language} text to {target_language}. "
        f"Provide only the {target_language} translation, nothing else:\n\n"
    )
    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=2000,
                messages=[{"role": "user", "content": context + text}]
            )
            if message.stop_reason == 'refusal':
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                return text
            if message and hasattr(message, 'content') and message.content:
                block = message.content[0]
                if hasattr(block, 'text') and block.text.strip():
                    return block.text.strip()
            if attempt < max_retries - 1:
                time.sleep(1)
        except Exception as e:
            print(f"  Exception (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                time.sleep(1)
    print(f"  ERROR: Translation failed for: {text[:80]}...")
    return text


# ── Vision pipeline (PNG / JPG) ────────────────────────────────────────────────

_VISION_PROMPT = """\
You are an expert OCR and translation assistant. Carefully examine the image and \
identify every piece of visible text.

For each distinct text region return:
  - "original"    : the exact text as it appears in the image
  - "translated"  : the text translated into {target_language}
  - "left_pct"    : left edge of the text bounding box, as % of image width  (0–100)
  - "top_pct"     : top edge, as % of image height (0–100)
  - "right_pct"   : right edge, as % of image width  (0–100)
  - "bottom_pct"  : bottom edge, as % of image height (0–100)
  - "text_color"  : "dark" if the text appears dark (black / navy / dark-grey), \
"light" if light (white / yellow / light-grey)

Rules:
• Include ALL text — headings, body, labels, captions, numbers with units, legends, etc.
• right_pct > left_pct and bottom_pct > top_pct always.
• Bounding boxes should tightly enclose the text.
• If the source text is already in {target_language} set translated == original.
• SDR = Spectrum Dynamic Research (testing dept), HC = R&D department.
• Source language: {source_language}

Return ONLY a JSON object — no markdown, no explanation:
{{
  "text_regions": [
    {{
      "original": "...",
      "translated": "...",
      "left_pct": 0.0,
      "top_pct": 0.0,
      "right_pct": 100.0,
      "bottom_pct": 5.0,
      "text_color": "dark"
    }}
  ]
}}
"""


def _encode_image_b64(image_path: str) -> tuple[str, str]:
    """Return (base64_data, media_type) for an image file."""
    ext = Path(image_path).suffix.lower()
    media_map = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
                 '.png': 'image/png', '.gif': 'image/gif', '.webp': 'image/webp'}
    media_type = media_map.get(ext, 'image/png')
    with open(image_path, 'rb') as f:
        return base64.standard_b64encode(f.read()).decode('utf-8'), media_type


def _encode_pil_b64(pil_image: Image.Image) -> tuple[str, str]:
    """Encode a PIL image to base64 PNG."""
    import io
    buf = io.BytesIO()
    pil_image.save(buf, format='PNG')
    return base64.standard_b64encode(buf.getvalue()).decode('utf-8'), 'image/png'


def _call_vision_api(client: Anthropic, b64_data: str, media_type: str,
                     source_language: str, target_language: str,
                     max_retries: int = 2) -> list[dict]:
    """
    Send an image to Claude's vision API.
    Returns a list of text-region dicts (may be empty on failure).
    """
    prompt = _VISION_PROMPT.format(
        source_language=source_language,
        target_language=target_language
    )

    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=4096,
                messages=[{
                    "role": "user",
                    "content": [
                        {
                            "type": "image",
                            "source": {
                                "type": "base64",
                                "media_type": media_type,
                                "data": b64_data,
                            },
                        },
                        {"type": "text", "text": prompt},
                    ],
                }]
            )

            if message.stop_reason == 'refusal':
                print(f"  Vision API refusal on attempt {attempt + 1}")
                if attempt < max_retries - 1:
                    time.sleep(2)
                    continue
                return []

            raw = message.content[0].text.strip() if message.content else ""

            # Strip any markdown code fences Claude might add
            raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.MULTILINE)
            raw = re.sub(r'\s*```$', '', raw, flags=re.MULTILINE)

            data = json.loads(raw)
            return data.get("text_regions", [])

        except json.JSONDecodeError as e:
            print(f"  Warning: Could not parse vision response as JSON (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                time.sleep(1)
        except Exception as e:
            print(f"  Vision API error (attempt {attempt + 1}): {e}")
            if attempt < max_retries - 1:
                time.sleep(1)

    return []


def _get_bg_color(img: Image.Image, x1: int, y1: int, x2: int, y2: int) -> tuple:
    """
    Estimate the background color of a region by sampling edge pixels.
    Returns an RGB tuple.
    """
    if img.mode not in ('RGB', 'RGBA'):
        img_rgb = img.convert('RGB')
    else:
        img_rgb = img

    w, h = img_rgb.size
    x1, y1 = max(0, x1), max(0, y1)
    x2, y2 = min(w, x2), min(h, y2)

    if x2 <= x1 or y2 <= y1:
        return (255, 255, 255)

    # Sample pixels along the outer edge of the region
    step_x = max(1, (x2 - x1) // 20)
    step_y = max(1, (y2 - y1) // 20)
    pixels = []

    for x in range(x1, x2, step_x):
        pixels.append(img_rgb.getpixel((x, y1)))
        pixels.append(img_rgb.getpixel((x, min(h - 1, y2 - 1))))
    for y in range(y1, y2, step_y):
        pixels.append(img_rgb.getpixel((x1, y)))
        pixels.append(img_rgb.getpixel((min(w - 1, x2 - 1), y)))

    if not pixels:
        return (255, 255, 255)

    # Bucket colours (32-level quantisation) and return the most frequent bucket
    from collections import Counter
    buckets = [(r // 32, g // 32, b // 32) for r, g, b in (p[:3] for p in pixels)]
    best = Counter(buckets).most_common(1)[0][0]
    return (best[0] * 32 + 16, best[1] * 32 + 16, best[2] * 32 + 16)


def _is_light(color: tuple) -> bool:
    """Return True if the colour is perceived as light (high luminance)."""
    r, g, b = color[:3]
    # Perceived luminance (sRGB weighting)
    return (0.299 * r + 0.587 * g + 0.114 * b) > 128


def _fit_text(draw: ImageDraw.ImageDraw, text: str, max_w: int, max_h: int,
              target_language: str, max_size: int = 72) -> tuple:
    """
    Find the largest font size at which `text` fits within max_w × max_h.
    Returns (ImageFont, actual_size).
    """
    for size in range(min(max_size, max_h), 5, -1):
        font = _get_pil_font(target_language, size)
        if font is None:
            continue
        # Wrap text to fit width
        wrapped = _wrap_text(draw, text, font, max_w)
        bbox = draw.multiline_textbbox((0, 0), wrapped, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        if tw <= max_w and th <= max_h:
            return font, size, wrapped
    # Absolute fallback – smallest readable size
    font = _get_pil_font(target_language, 8)
    if font is None:
        font = ImageFont.load_default()
    return font, 8, text


def _wrap_text(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.FreeTypeFont,
               max_width: int) -> str:
    """Wrap text so that no line exceeds max_width pixels."""
    words = text.split()
    if not words:
        return text

    lines = []
    current = ""
    for word in words:
        test = (current + " " + word).strip()
        bbox = draw.textbbox((0, 0), test, font=font)
        if bbox[2] - bbox[0] <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return "\n".join(lines) if lines else text


def _overlay_regions(img: Image.Image, regions: list, target_language: str) -> Image.Image:
    """
    Paint translated text over each detected region on the image.
    Returns a modified copy of the image.
    """
    if not regions:
        return img

    result = img.convert('RGBA').copy()
    overlay = Image.new('RGBA', result.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)
    w, h = result.size

    for region in regions:
        try:
            # Convert % coordinates to pixels
            x1 = int(region['left_pct']   / 100 * w)
            y1 = int(region['top_pct']    / 100 * h)
            x2 = int(region['right_pct']  / 100 * w)
            y2 = int(region['bottom_pct'] / 100 * h)

            # Guard against degenerate boxes
            if x2 <= x1 or y2 <= y1 or (x2 - x1) < 4 or (y2 - y1) < 4:
                continue

            translated = region.get('translated', '').strip()
            if not translated:
                continue

            # Background colour for this region
            bg_color = _get_bg_color(img, x1, y1, x2, y2)
            text_color_hint = region.get('text_color', 'dark')

            # If the hint says the text was light, the background is dark (and vice-versa)
            if text_color_hint == 'light':
                text_rgb = (240, 240, 240)
            else:
                text_rgb = (20, 20, 20)

            # Paint background rectangle (solid, no alpha so it fully covers)
            draw.rectangle([x1, y1, x2, y2], fill=bg_color + (255,))

            # Find the best font + size and wrap text
            pad = 2
            box_w = max(1, x2 - x1 - pad * 2)
            box_h = max(1, y2 - y1 - pad * 2)
            font, size, wrapped = _fit_text(draw, translated, box_w, box_h, target_language)

            draw.multiline_text((x1 + pad, y1 + pad), wrapped, font=font,
                                fill=text_rgb + (255,))

        except Exception as e:
            print(f"  Warning: Could not overlay region '{region.get('original', '')[:40]}': {e}")
            continue

    result = Image.alpha_composite(result, overlay)
    return result.convert('RGB')


# ── PNG / JPG translation ──────────────────────────────────────────────────────

def translate_png_jpg(input_path: str, output_path: str,
                      source_language: str, target_language: str,
                      client: Anthropic) -> None:
    """Translate visible text in a PNG or JPG image and save the result."""
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow is required for image translation (pip install Pillow)")

    print(f"Loading image: {input_path}")
    print(f"Translating: {source_language} → {target_language}")

    img = Image.open(input_path)
    if img.mode not in ('RGB', 'RGBA'):
        img = img.convert('RGB')

    print("  Sending image to Claude vision API for text detection and translation...")
    b64_data, media_type = _encode_image_b64(input_path)
    regions = _call_vision_api(client, b64_data, media_type, source_language, target_language)

    if not regions:
        print("  No text regions detected or translation failed — saving original.")
        img.save(output_path)
        return

    # Filter out regions where original == translated (already in target language)
    regions_to_draw = [r for r in regions
                       if r.get('translated', '').strip() != r.get('original', '').strip()
                       and r.get('translated', '').strip()]

    print(f"  Detected {len(regions)} text region(s), translating {len(regions_to_draw)}...")

    result = _overlay_regions(img, regions_to_draw, target_language)

    # Preserve original format where possible
    fmt = 'JPEG' if output_path.lower().endswith(('.jpg', '.jpeg')) else 'PNG'
    if fmt == 'JPEG':
        result.save(output_path, 'JPEG', quality=95)
    else:
        result.save(output_path, 'PNG')

    print(f"Saved translated image: {output_path}")


# ── PDF translation ────────────────────────────────────────────────────────────

def _page_has_selectable_text(page: 'fitz.Page', min_chars: int = 10) -> bool:
    """Return True if a PDF page contains enough selectable text to translate directly."""
    return len(page.get_text("text").strip()) >= min_chars


def _color_int_to_float(color_int: int) -> tuple:
    """Convert a PyMuPDF integer colour (0xRRGGBB) to a (r, g, b) float tuple 0–1."""
    r = ((color_int >> 16) & 0xFF) / 255.0
    g = ((color_int >> 8)  & 0xFF) / 255.0
    b = (color_int         & 0xFF) / 255.0
    return (r, g, b)


def _translate_pdf_text_page(page: 'fitz.Page', source_language: str,
                              target_language: str, client: Anthropic) -> int:
    """
    Translate all text blocks on a text-based PDF page in-place.
    Returns the number of blocks translated.
    """
    blocks_raw = page.get_text("dict", flags=11)["blocks"]
    translated_count = 0
    pdf_fontname = _pdf_fontname(target_language)

    # Collect blocks to translate first, then apply (to avoid modifying while iterating)
    to_process = []
    for block in blocks_raw:
        if block.get("type") != 0:
            continue
        # Gather the full block text and its font properties from the first span
        block_text = ""
        first_span = None
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                block_text += span.get("text", "")
                if first_span is None and span.get("text", "").strip():
                    first_span = span

        block_text = block_text.strip()
        if not block_text or not first_span:
            continue
        if not needs_translation(block_text, source_language, target_language):
            continue

        to_process.append((fitz.Rect(block["bbox"]), block_text, first_span))

    if not to_process:
        return 0

    # Translate all blocks
    translated_blocks = []
    for rect, text, span in to_process:
        translated = translate_text(text, client, source_language, target_language)
        if translated != text:
            translated_blocks.append((rect, translated, span))
            translated_count += 1

    if not translated_blocks:
        return 0

    # White-out original text
    for rect, _, _ in translated_blocks:
        # Slightly expand rect to ensure full coverage
        expanded = rect + (-1, -1, 1, 1)
        page.draw_rect(expanded, color=(1, 1, 1), fill=(1, 1, 1))

    # Insert translated text
    for rect, translated, span in translated_blocks:
        font_size = span.get("size", 11)
        color_int = span.get("color", 0)
        color = _color_int_to_float(color_int)

        # Try to fit text at original size, then reduce if needed
        inserted = False
        for size in [font_size, font_size * 0.85, font_size * 0.7, font_size * 0.55, 7]:
            size = max(6.0, size)
            try:
                overflow = page.insert_textbox(
                    rect,
                    translated,
                    fontsize=size,
                    fontname=pdf_fontname,
                    color=color,
                    align=0,
                )
                if overflow >= 0:     # text fit
                    inserted = True
                    break
            except Exception as e:
                # Some CJK font names may not be available; fall back to helv
                try:
                    overflow = page.insert_textbox(
                        rect, translated, fontsize=size,
                        fontname=_PDF_FONT_LATIN, color=color, align=0
                    )
                    if overflow >= 0:
                        inserted = True
                        break
                except Exception:
                    pass
                break

        if not inserted:
            # Last resort: insert at smallest size, allow slight overflow
            try:
                page.insert_textbox(rect, translated, fontsize=7,
                                    fontname=_PDF_FONT_LATIN, color=color, align=0)
            except Exception as e:
                print(f"  Warning: Could not insert translated text: {e}")

    return translated_count


def _translate_pdf_image_page(page: 'fitz.Page', page_num: int,
                               source_language: str, target_language: str,
                               client: Anthropic) -> 'Image.Image | None':
    """
    Rasterise an image-based PDF page and translate text via vision API.
    Returns a PIL Image of the translated page, or None on failure.
    """
    if not PIL_AVAILABLE:
        return None

    # Render at 2× for better OCR accuracy
    mat = fitz.Matrix(2.0, 2.0)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    print(f"    Page {page_num}: image-based, using vision API...")
    b64_data, media_type = _encode_pil_b64(img)
    regions = _call_vision_api(client, b64_data, media_type, source_language, target_language)

    if not regions:
        return img

    regions_to_draw = [r for r in regions
                       if r.get('translated', '').strip() != r.get('original', '').strip()
                       and r.get('translated', '').strip()]

    return _overlay_regions(img, regions_to_draw, target_language)


def translate_pdf(input_path: str, output_path: str,
                  source_language: str, target_language: str,
                  client: Anthropic) -> None:
    """
    Translate all text in a PDF file.

    Text-based pages are handled with PyMuPDF redaction + text insertion.
    Image-based pages (scanned) are rasterised, processed with the vision API,
    and compiled back into a new PDF.
    """
    if not FITZ_AVAILABLE:
        raise RuntimeError(
            "PyMuPDF is required for PDF translation (pip install pymupdf)"
        )

    print(f"Loading PDF: {input_path}")
    print(f"Translating: {source_language} → {target_language}")

    doc = fitz.open(input_path)
    n_pages = len(doc)

    # Track which pages are image-based so we can compile a hybrid output
    image_pages: dict[int, 'Image.Image'] = {}     # page_index → PIL image

    for page_idx in range(n_pages):
        page = doc[page_idx]
        page_num = page_idx + 1
        print(f"\n  Processing page {page_num}/{n_pages}...")

        if _page_has_selectable_text(page):
            count = _translate_pdf_text_page(page, source_language, target_language, client)
            print(f"    Translated {count} text block(s)")
        else:
            translated_img = _translate_pdf_image_page(
                page, page_num, source_language, target_language, client
            )
            if translated_img is not None:
                image_pages[page_idx] = translated_img
                print(f"    Image page translated via vision API")
            else:
                print(f"    Image page — no text found, keeping original")

    # For pages processed as images: replace page content with the translated image
    if image_pages:
        for page_idx, pil_img in image_pages.items():
            page = doc[page_idx]
            pw = page.rect.width
            ph = page.rect.height

            # Save PIL image to a temp buffer and insert as a full-page image
            import io
            buf = io.BytesIO()
            pil_img.save(buf, format='PNG')
            buf.seek(0)

            # Clear the page and insert the translated image
            page.clean_contents()
            page.insert_image(page.rect, stream=buf.read())

    print(f"\nSaving translated PDF: {output_path}")
    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    print("PDF translation complete!")


# ── Main entry point ───────────────────────────────────────────────────────────

def translate_image_file(input_path: str, output_path: str,
                         source_language: str, target_language: str,
                         client: Anthropic) -> None:
    """
    Translate visible text in an image or PDF file.

    Supported input formats: .jpg, .jpeg, .png, .pdf

    Args:
        input_path:       Path to the input file
        output_path:      Path to save the translated file
        source_language:  Source language display name (e.g. 'English')
        target_language:  Target language display name (e.g. 'Chinese (Simplified)')
        client:           Anthropic API client instance
    """
    ext = Path(input_path).suffix.lower()

    if ext in ('.jpg', '.jpeg', '.png'):
        translate_png_jpg(input_path, output_path, source_language, target_language, client)
    elif ext == '.pdf':
        translate_pdf(input_path, output_path, source_language, target_language, client)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Supported: .jpg, .jpeg, .png, .pdf")


def generate_output_path(input_path: str) -> str:
    """Generate an output path by inserting '_translated' before the extension."""
    p = Path(input_path)
    return str(p.parent / f"{p.stem}_translated{p.suffix}")


# ── CLI entry point ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    from tkinter import Tk, filedialog

    root = Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    input_file = filedialog.askopenfilename(
        title="Select image or PDF to translate",
        filetypes=[
            ("Supported files", "*.jpg *.jpeg *.png *.pdf"),
            ("Image files", "*.jpg *.jpeg *.png"),
            ("PDF files", "*.pdf"),
            ("All files", "*.*"),
        ]
    )
    root.destroy()

    if not input_file:
        print("No file selected. Exiting.")
        sys.exit(0)

    print("\nSelect source language:")
    print("1. English\n2. Chinese (Simplified)")
    src = "Chinese (Simplified)" if input("Enter choice (1 or 2): ").strip() == '2' else "English"

    print("\nSelect target language:")
    print("1. English\n2. Chinese (Simplified)")
    tgt = "Chinese (Simplified)" if input("Enter choice (1 or 2): ").strip() == '2' else "English"

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        api_key = input("Enter Anthropic API key: ").strip()

    cli_client = Anthropic(api_key=api_key)
    output_file = generate_output_path(input_file)

    print(f"\nInput:  {input_file}")
    print(f"Output: {output_file}\n")

    translate_image_file(input_file, output_file, src, tgt, cli_client)
