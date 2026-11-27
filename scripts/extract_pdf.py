#!/usr/bin/env python3
"""Extract text, per-page text, and figure/table/equation references from a PDF.

Pure standard library. Uses poppler's pdftotext / pdftoppm when available and falls
back to PyMuPDF / pypdf for text. Outputs a manifest.json describing what was found.

Usage:
  python extract_pdf.py --pdf paper.pdf --out workdir [--dpi 120] [--no-render]
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys


def run(cmd, **kw):
    kw.setdefault("capture_output", True)
    kw.setdefault("text", True)
    kw.setdefault("errors", "replace")
    return subprocess.run(cmd, **kw)


def page_texts(fulltext):
    """Split pdftotext output into per-page strings (form-feed separated)."""
    pages = fulltext.split("\f")
    if len(pages) > 1 and pages[-1].strip() == "":
        pages = pages[:-1]
    return pages


def detect_refs(pages):
    fig = re.compile(r'(?i)\b(fig(?:ure)?)\s*\.?\s*(\d+[a-z]?)')
    tab = re.compile(r'(?i)\b(table)\s*\.?\s*(\d+[a-z]?)')
    eqn = re.compile(r'(?i)\b(eq(?:uation)?)\s*\.?\s*\(?(\d+[a-z]?)\)?')
    figures, tables, equations = [], [], []
    for i, txt in enumerate(pages, 1):
        for m in fig.finditer(txt):
            figures.append({"id": "Fig. %s" % m.group(2), "kind": "figure",
                            "page": i, "caption_snippet": snippet(txt, m.start())})
        for m in tab.finditer(txt):
            tables.append({"id": "Table %s" % m.group(2), "kind": "table",
                           "page": i, "caption_snippet": snippet(txt, m.start())})
        for m in eqn.finditer(txt):
            equations.append({"id": "Eq. %s" % m.group(2), "kind": "equation",
                              "page": i, "caption_snippet": snippet(txt, m.start())})
    # de-duplicate by id, keep first occurrence
    def dedup(items):
        seen, out = set(), []
        for it in items:
            if it["id"] not in seen:
                seen.add(it["id"]); out.append(it)
        return out
    return dedup(figures), dedup(tables), dedup(equations)


def snippet(txt, pos):
    s = txt[max(0, pos - 60):pos + 160].replace("\n", " ").strip()
    return re.sub(r"\s+", " ", s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--dpi", type=int, default=120)
    ap.add_argument("--no-render", action="store_true")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)
    pages_dir = os.path.join(args.out, "pages")
    images_dir = os.path.join(args.out, "page_images")
    os.makedirs(pages_dir, exist_ok=True)

    manifest = {"pdf": os.path.abspath(args.pdf), "page_count": None,
                "fulltext": None, "pages": [], "figures": [], "tables": [],
                "equations": [], "rendered_pages": [], "tool": None,
                "status": "ok", "warnings": []}

    pdftotext = shutil.which("pdftotext")
    pdftoppm = shutil.which("pdftoppm")
    fulltext = None

    if pdftotext:
        fulltext_path = os.path.join(args.out, "fulltext.txt")
        r = run([pdftotext, "-layout", args.pdf, fulltext_path])
        if r.returncode == 0:
            fulltext = open(fulltext_path, encoding="utf-8", errors="replace").read()
            manifest["fulltext"] = "fulltext.txt"
            manifest["tool"] = "pdftotext"
        else:
            manifest["warnings"].append("pdftotext failed: %s" % r.stderr.strip()[:300])

    if fulltext is None:
        # fallback: PyMuPDF / pypdf
        try:
            import fitz  # PyMuPDF
            doc = fitz.open(args.pdf)
            fulltext = "\f".join(page.get_text() for page in doc)
            manifest["tool"] = "pymupdf"
            manifest["warnings"].append("pdftotext unavailable; used PyMuPDF")
        except Exception:
            try:
                from pypdf import PdfReader
                fulltext = "\f".join((p.extract_text() or "") for p in PdfReader(args.pdf).pages)
                manifest["tool"] = "pypdf"
                manifest["warnings"].append("pdftotext unavailable; used pypdf")
            except Exception as e:
                fulltext = None
                manifest["status"] = "partial"
                manifest["warnings"].append("text extraction unavailable: %s" % e)

    if fulltext is not None:
        pages = page_texts(fulltext)
        manifest["page_count"] = len(pages)
        for i, txt in enumerate(pages, 1):
            pname = "page-%03d.txt" % i
            open(os.path.join(pages_dir, pname), "w", encoding="utf-8").write(txt)
            manifest["pages"].append("pages/" + pname)
        manifest["figures"], manifest["tables"], manifest["equations"] = detect_refs(pages)

    # render pages
    if pdftoppm and not args.no_render:
        os.makedirs(images_dir, exist_ok=True)
        prefix = os.path.join(images_dir, "page")
        r = run([pdftoppm, "-png", "-r", str(args.dpi), args.pdf, prefix])
        if r.returncode == 0:
            imgs = sorted(f for f in os.listdir(images_dir) if f.endswith(".png"))
            manifest["rendered_pages"] = ["page_images/" + f for f in imgs]
        else:
            manifest["warnings"].append("pdftoppm failed: %s" % r.stderr.strip()[:200])

    json.dump(manifest, open(os.path.join(args.out, "manifest.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print(json.dumps({"status": manifest["status"], "page_count": manifest["page_count"],
                      "figures": len(manifest["figures"]), "tables": len(manifest["tables"]),
                      "equations": len(manifest["equations"]), "rendered": len(manifest["rendered_pages"]),
                      "tool": manifest["tool"], "manifest": os.path.join(args.out, "manifest.json")},
                     ensure_ascii=False))


if __name__ == "__main__":
    main()
