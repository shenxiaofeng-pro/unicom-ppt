#!/usr/bin/env python3
"""
render.py —— 把 PPTX 渲染成 PDF 和逐页 PNG，用于目检。

    python render.py deck.pptx [--out 输出目录] [--dpi 110] [--pages 3 | 2-4 | 1,3,5-6]

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


LAST_ERROR = []          # 最近一次转换失败的原因（找到了工具但转换没成功）


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
        # PowerPoint 是沙盒应用，每访问一个新目录都会弹窗请求授权，弹窗未处理时所有自动化都会失败（-9074）。
        # 所以副本固定放在输出目录下、使用固定文件名：每个目录只需授权一次。
        # 导出前先关掉以前残留的渲染副本（只关自己的，不碰用户打开的文件）。
        work = os.path.abspath(out_dir)
        src = os.path.join(work, "_unicom_render_src.pptx")
        dst = os.path.join(work, "_unicom_render_out.pdf")
        for f in (src, dst):
            if os.path.exists(f):
                os.remove(f)
        shutil.copy(pptx, src)
        script = f'''
        tell application "Microsoft PowerPoint"
            repeat with i from (count of presentations) to 1 by -1
                set fn to full name of presentation i
                if fn contains "_unicom_render_src" or fn contains "unicom_render_" then close presentation i saving no
            end repeat
            open POSIX file "{src}"
            set target to missing value
            repeat with i from 1 to (count of presentations)
                if full name of presentation i is "{src}" then set target to presentation i
            end repeat
            if target is missing value then error "PowerPoint 未能打开文件"
            try
                save target in POSIX file "{dst}" as save as PDF
            on error errMsg
                close target saving no
                error errMsg
            end try
            close target saving no
        end tell'''
        err = ""
        for attempt in range(2):                     # PowerPoint 偶尔卡住，失败后重试一次
            try:
                r = subprocess.run(["osascript", "-e", script], check=False, capture_output=True, timeout=180)
                err = r.stderr.decode(errors="ignore").strip()
            except subprocess.TimeoutExpired:
                err = "PowerPoint 响应超时"
            if os.path.exists(dst):
                shutil.move(dst, pdf)
                os.remove(src)
                return pdf, "PowerPoint (macOS)"
        if os.path.exists(src):
            os.remove(src)
        LAST_ERROR.append(f"PowerPoint (macOS) 转换失败：{err or '未知原因'}。"
                          f"请检查 PowerPoint 是否弹窗：如果是请求访问「{work}」的授权，请点「允许」（每个目录只需一次）；"
                          "其他弹窗（登录、更新、修复提示）关闭后重试。")

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


def to_png(pdf, out_dir, dpi, pages=None):
    """pages: 只渲染这些页码（1 起），None 表示全部。"""
    for old in glob.glob(os.path.join(out_dir, "slide-*.png")):
        os.remove(old)
    if shutil.which("pdftoppm"):
        for a, b in _ranges(pages) if pages else [(None, None)]:
            rng = ["-f", str(a), "-l", str(b)] if a else []
            subprocess.run(["pdftoppm", "-png", "-r", str(dpi), *rng, pdf, os.path.join(out_dir, "slide")],
                           check=True)
    else:
        try:
            import fitz  # PyMuPDF
        except ImportError:
            return []
        doc = fitz.open(pdf)
        for i, page in enumerate(doc, 1):
            if pages is None or i in pages:
                page.get_pixmap(dpi=dpi).save(os.path.join(out_dir, f"slide-{i:02d}.png"))
    for f in glob.glob(os.path.join(out_dir, "slide-*.png")):        # 统一为 slide-01.png
        n = os.path.basename(f)[6:-4]
        if n.isdigit() and len(n) < 2:
            os.rename(f, os.path.join(out_dir, f"slide-{int(n):02d}.png"))
    return sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))


def _ranges(pages):
    pages = sorted(set(pages))
    out, start, prev = [], pages[0], pages[0]
    for p in pages[1:]:
        if p != prev + 1:
            out.append((start, prev))
            start = p
        prev = p
    out.append((start, prev))
    return out


def _parse_pages(text):
    """'3' / '2-4' / '1,3,5-6' → [页码]"""
    out = []
    for part in text.split(","):
        if "-" in part:
            a, b = part.split("-")
            out += range(int(a), int(b) + 1)
        else:
            out.append(int(part))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None, help="输出目录（默认：与 pptx 同目录下的 <文件名>_render）")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--pages", default=None, help="只渲染指定页，如 3、2-4、1,3,5-6（逐页校验时使用）")
    a = ap.parse_args()
    out = a.out or os.path.splitext(os.path.abspath(a.pptx))[0] + "_render"
    os.makedirs(out, exist_ok=True)
    pdf, backend = to_pdf(a.pptx, out)
    if not pdf:
        print(LAST_ERROR[-1] if LAST_ERROR else
              "未找到可用的转换工具：请安装 LibreOffice，或在 macOS/Windows 上安装 Microsoft PowerPoint。")
        sys.exit(2)
    pngs = to_png(pdf, out, a.dpi, pages=_parse_pages(a.pages) if a.pages else None)
    print(f"渲染后端：{backend}\nPDF：{pdf}")
    if pngs:
        print("PNG：\n" + "\n".join(pngs))
    else:
        print("未找到 pdftoppm / PyMuPDF，仅生成 PDF（brew install poppler 或 pip install pymupdf）。")


if __name__ == "__main__":
    main()
