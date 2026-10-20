#!/usr/bin/env python3
"""
render.py —— 把 PPTX 渲染成 PDF 和逐页 PNG，用于目检。

    python render.py deck.pptx [--out 输出目录] [--dpi 110]

转换后端按顺序尝试：
  1. LibreOffice（soffice / libreoffice，跨平台）
  2. macOS 上的 Microsoft PowerPoint（AppleScript）
  3. Windows 上的 Microsoft PowerPoint（PowerShell COM）
PNG 由 pdftoppm（poppler）或 PyMuPDF 生成；两者都没有时只输出 PDF。
注意：LibreOffice 对中文字体的替换可能与 PowerPoint/WPS 略有差异，以 PowerPoint 渲染为准。
"""
import argparse
import glob
import os
import platform
import shutil
import subprocess
import sys
import tempfile


def to_pdf(pptx, out_dir):
    pptx = os.path.abspath(pptx)
    base = os.path.splitext(os.path.basename(pptx))[0]
    pdf = os.path.join(out_dir, base + ".pdf")

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice and platform.system() == "Darwin" and os.path.exists("/Applications/LibreOffice.app"):
        soffice = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx],
                       check=False, capture_output=True, timeout=300)
        if os.path.exists(pdf):
            return pdf, "LibreOffice"

    if platform.system() == "Darwin" and os.path.exists("/Applications/Microsoft PowerPoint.app"):
        # PowerPoint 在沙盒中运行，先复制到临时目录再导出
        tmp = tempfile.mkdtemp(prefix="unicom_render_")
        src = os.path.join(tmp, "deck.pptx")
        dst = os.path.join(tmp, "deck.pdf")
        shutil.copy(pptx, src)
        script = f'''
        tell application "Microsoft PowerPoint"
            open POSIX file "{src}"
            set p to active presentation
            save p in POSIX file "{dst}" as save as PDF
            close p saving no
        end tell'''
        subprocess.run(["osascript", "-e", script], check=False, capture_output=True, timeout=300)
        if os.path.exists(dst):
            shutil.move(dst, pdf)
            return pdf, "PowerPoint (macOS)"

    if platform.system() == "Windows":
        ps = f'''
        $pp = New-Object -ComObject PowerPoint.Application
        $p = $pp.Presentations.Open("{pptx}", $true, $false, $false)
        $p.SaveAs("{pdf}", 32)
        $p.Close(); $pp.Quit()'''
        subprocess.run(["powershell", "-NoProfile", "-Command", ps], check=False, capture_output=True, timeout=300)
        if os.path.exists(pdf):
            return pdf, "PowerPoint (Windows)"

    return None, None


def to_png(pdf, out_dir, dpi):
    for old in glob.glob(os.path.join(out_dir, "slide-*.png")):
        os.remove(old)
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf, os.path.join(out_dir, "slide")], check=True)
    else:
        try:
            import fitz  # PyMuPDF
        except ImportError:
            return []
        doc = fitz.open(pdf)
        for i, page in enumerate(doc, 1):
            page.get_pixmap(dpi=dpi).save(os.path.join(out_dir, f"slide-{i:02d}.png"))
    for f in glob.glob(os.path.join(out_dir, "slide-*.png")):        # 统一为 slide-01.png
        n = os.path.basename(f)[6:-4]
        if n.isdigit() and len(n) < 2:
            os.rename(f, os.path.join(out_dir, f"slide-{int(n):02d}.png"))
    return sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None, help="输出目录（默认：与 pptx 同目录下的 <文件名>_render）")
    ap.add_argument("--dpi", type=int, default=110)
    a = ap.parse_args()
    out = a.out or os.path.splitext(os.path.abspath(a.pptx))[0] + "_render"
    os.makedirs(out, exist_ok=True)
    pdf, backend = to_pdf(a.pptx, out)
    if not pdf:
        print("未找到可用的转换工具：请安装 LibreOffice，或在 macOS/Windows 上安装 Microsoft PowerPoint。")
        sys.exit(2)
    pngs = to_png(pdf, out, a.dpi)
    print(f"渲染后端：{backend}\nPDF：{pdf}")
    if pngs:
        print("PNG：\n" + "\n".join(pngs))
    else:
        print("未找到 pdftoppm / PyMuPDF，仅生成 PDF（brew install poppler 或 pip install pymupdf）。")


if __name__ == "__main__":
    main()
