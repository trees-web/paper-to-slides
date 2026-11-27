#!/usr/bin/env python3
"""Render a PPTX/PPT to PDF and per-page PNG previews for inspection.

Uses LibreOffice (soffice) for the PDF and poppler pdftoppm for PNG previews.
Records what succeeded and what did not in render_manifest.json.

Usage:
  python render_deck.py --deck deck.pptx --out workdir [--dpi 120] [--no-png]
"""
import argparse
import json
import os
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path

SOFFICE_CANDIDATES = ["soffice", "libreoffice",
                      r"C:\Program Files\LibreOffice\program\soffice.exe",
                      r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                      "/usr/bin/soffice", "/usr/bin/libreoffice"]


def which_soffice():
    p = shutil.which("soffice") or shutil.which("libreoffice")
    if p:
        return p
    for c in SOFFICE_CANDIDATES:
        if os.path.exists(c):
            return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--dpi", type=int, default=120)
    ap.add_argument("--no-png", action="store_true")
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    manifest = {"deck": os.path.abspath(args.deck), "pdf": None, "pngs": [],
                "soffice": None, "status": "ok", "warnings": []}

    soffice = which_soffice()
    if not soffice:
        manifest["status"] = "failed"
        manifest["warnings"].append("LibreOffice (soffice) not found; cannot render PPTX")
        json.dump(manifest, open(os.path.join(args.out, "render_manifest.json"), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)
        print(json.dumps(manifest, ensure_ascii=False))
        return 1

    manifest["soffice"] = soffice
    if os.name == "nt":
        # mkdtemp uses mode 0o700, which replaces inherited Windows ACLs.
        # Keep the workspace ACL so the sandboxed LibreOffice child can write.
        profile = str(Path(args.out).resolve() / ("lo-profile-" + uuid.uuid4().hex))
        os.mkdir(profile, mode=0o777)
    else:
        profile = tempfile.mkdtemp(prefix="lo-profile-")
    try:
        r = subprocess.run([soffice, "-env:UserInstallation=" + Path(profile).resolve().as_uri(),
                            "--headless", "--norestore", "--nofirststartwizard", "--convert-to", "pdf",
                            "--outdir", args.out, args.deck],
                           capture_output=True, text=True, errors="replace", timeout=600)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    pdf_path = os.path.join(args.out, os.path.splitext(os.path.basename(args.deck))[0] + ".pdf")
    if r.returncode == 0 and os.path.exists(pdf_path):
        manifest["pdf"] = os.path.basename(pdf_path)
    else:
        manifest["status"] = "partial"
        manifest["warnings"].append("soffice PDF export failed: %s" % (r.stderr or r.stdout).strip()[:400])

    if manifest["pdf"] and not args.no_png:
        pdftoppm = shutil.which("pdftoppm")
        if pdftoppm:
            png_dir = os.path.join(args.out, "previews")
            os.makedirs(png_dir, exist_ok=True)
            rr = subprocess.run([pdftoppm, "-png", "-r", str(args.dpi), pdf_path,
                                 os.path.join(png_dir, "preview")],
                                capture_output=True, text=True, errors="replace", timeout=600)
            if rr.returncode == 0:
                pngs = sorted(f for f in os.listdir(png_dir) if f.endswith(".png"))
                manifest["pngs"] = ["previews/" + f for f in pngs]
            else:
                manifest["warnings"].append("pdftoppm failed: %s" % rr.stderr.strip()[:200])
        else:
            manifest["warnings"].append("pdftoppm not found; no PNG previews")

    json.dump(manifest, open(os.path.join(args.out, "render_manifest.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print(json.dumps(manifest, ensure_ascii=False))


if __name__ == "__main__":
    main()
