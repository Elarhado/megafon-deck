#!/usr/bin/env python3
"""Render a pptx to PDF + per-slide PNGs + contact sheet.
Usage: render.py deck.pptx [--pdf out.pdf] [--png-dir DIR] [--dpi 80]
Exit 3 when a required tool is missing (message says how to install)."""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402

C.require("PIL")

from PIL import Image, ImageDraw, ImageFont  # noqa: E402

MAC_SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
MAC_POWERPOINT = "/Applications/Microsoft PowerPoint.app"


def die(msg, code=3):
    print(msg, file=sys.stderr)
    sys.exit(code)


def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    return MAC_SOFFICE if os.path.exists(MAC_SOFFICE) else None


def pdf_libreoffice(soffice, pptx, pdf):
    with tempfile.TemporaryDirectory() as tmp:
        profile = "file://" + os.path.join(tmp, "profile")
        cmd = [soffice, "-env:UserInstallation=" + profile, "--headless", "--convert-to", "pdf",
               "--outdir", tmp, pptx]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        out = os.path.join(tmp, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
        if not os.path.exists(out):
            die("LibreOffice failed to convert:\n%s%s" % (r.stdout, r.stderr), 4)
        shutil.move(out, pdf)


def _as_str(p):
    """Escape a path for an AppleScript string literal."""
    return p.replace("\\", "\\\\").replace('"', '\\"')


def pdf_powerpoint(pptx, pdf):
    script = ('tell application "Microsoft PowerPoint"\n'
              '  open POSIX file "%s"\n'
              '  save active presentation in POSIX file "%s" as save as PDF\n'
              '  close active presentation saving no\n'
              'end tell' % (_as_str(os.path.abspath(pptx)), _as_str(os.path.abspath(pdf))))
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=300)
    if r.returncode != 0 or not os.path.exists(pdf):
        die("PowerPoint export failed:\n" + r.stderr, 4)


def make_pdf(pptx, pdf):
    so = find_soffice()
    if so:
        return pdf_libreoffice(so, pptx, pdf)
    if sys.platform == "darwin" and os.path.exists(MAC_POWERPOINT):
        return pdf_powerpoint(pptx, pdf)
    die("No PDF converter found. Install LibreOffice (macOS: `brew install --cask libreoffice`; "
        "Linux: `sudo apt install libreoffice-impress`) or export the PDF manually from PowerPoint "
        "(File > Export > PDF) and pass it with --pdf.")


def pdf_to_pngs(pdf, outdir, dpi):
    os.makedirs(outdir, exist_ok=True)
    for old in glob.glob(os.path.join(outdir, "slide-*.png")):
        os.remove(old)
    if shutil.which("pdftoppm"):
        prefix = os.path.join(outdir, "slide")
        subprocess.run(["pdftoppm", "-r", str(dpi), "-png", pdf, prefix], check=True)
    else:
        try:
            import fitz
        except ImportError:
            die("Neither pdftoppm nor PyMuPDF found. Install one: `python3 -m pip install --user pymupdf` "
                "(or poppler: `brew install poppler`).")
        doc = fitz.open(pdf)
        for i, page in enumerate(doc, 1):
            page.get_pixmap(dpi=dpi).save(os.path.join(outdir, "slide-%02d.png" % i))
    files = sorted(glob.glob(os.path.join(outdir, "slide-*.png")))
    # normalise names to slide-NN.png
    out = []
    for i, f in enumerate(files, 1):
        new = os.path.join(outdir, "slide-%02d.png" % i)
        if f != new:
            os.replace(f, new)
        out.append(new)
    return out


def contact_sheet(pngs, path, per_row=4, thumb_w=480):
    ims = [Image.open(p).convert("RGB") for p in pngs]
    scale = thumb_w / ims[0].width
    th = int(ims[0].height * scale)
    rows = (len(ims) + per_row - 1) // per_row
    gap, lab = 12, 26
    sheet = Image.new("RGB", (per_row * (thumb_w + gap) + gap, rows * (th + lab + gap) + gap), (200, 200, 200))
    fp = C.find_font(True)
    font = ImageFont.truetype(fp, 20) if fp else ImageFont.load_default()
    d = ImageDraw.Draw(sheet)
    for i, im in enumerate(ims):
        x, y = gap + (i % per_row) * (thumb_w + gap), gap + (i // per_row) * (th + lab + gap)
        d.text((x + 2, y + 1), str(i + 1), fill=(0, 0, 0), font=font)
        sheet.paste(im.resize((thumb_w, th), Image.LANCZOS), (x, y + lab))
    sheet.save(path)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Render pptx to PDF and PNGs")
    ap.add_argument("pptx")
    ap.add_argument("--pdf")
    ap.add_argument("--png-dir")
    ap.add_argument("--dpi", type=int, default=80)
    a = ap.parse_args(argv)
    stem = os.path.splitext(os.path.abspath(a.pptx))[0]
    pdf = a.pdf or stem + ".pdf"
    png_dir = a.png_dir or stem + "_png"
    make_pdf(a.pptx, pdf)
    print("pdf: " + pdf)
    pngs = pdf_to_pngs(pdf, png_dir, a.dpi)
    sheet = os.path.join(png_dir, "sheet.png")
    contact_sheet(pngs, sheet)
    print("png: %d slides in %s\ncontact sheet: %s" % (len(pngs), png_dir, sheet))
    return 0


if __name__ == "__main__":
    sys.exit(main())
