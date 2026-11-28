#!/usr/bin/env python3
"""Build an editable PPTX deck from a JSON page-outline, preserving an optional template.

Pure standard library (no python-pptx / lxml required). Two paths:

* python-pptx (preferred when importable) — highest-fidelity template editing.
* self-contained OOXML writer — copies the template package verbatim and appends new
  slides/notesSlides that reference the template's layouts; each slide picks the layout
  that best matches its page type and reuses that layout's placeholder geometry, so the
  deck follows the template's title/body/picture positions, theme colors and fonts.

Usage:
  python build_deck.py --outline outline.json --out deck.pptx [--template template.pptx]
"""
import argparse
import json
import os
import re
import zipfile
from xml.sax.saxutils import escape

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

RT_SLIDEMASTER = NS_R + "/slideMaster"
RT_SLIDELAYOUT = NS_R + "/slideLayout"
RT_SLIDE = NS_R + "/slide"
RT_NOTESSLIDE = NS_R + "/notesSlide"
RT_NOTESMASTER = NS_R + "/notesMaster"
RT_THEME = NS_R + "/theme"
RT_IMAGE = NS_R + "/image"
RT_OFFICE = NS_R + "/officeDocument"

EMU_PER_INCH = 914400
SLIDE_W = int(13.333 * EMU_PER_INCH)
SLIDE_H = int(7.5 * EMU_PER_INCH)
NOTES_W = int(6.857 * EMU_PER_INCH)
NOTES_H = int(9.143 * EMU_PER_INCH)

CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
CT_NOTESSLIDE = "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"
CT_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"
CT_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
CT_NOTESMASTER = "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"


def esc(s):
    return escape(str(s), {'"': "&quot;"})


def _xfrm(x, y, cx, cy):
    return ('<a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>' % (x, y, cx, cy))


def _text_shape(shape_id, name, ph_type, x, y, cx, cy, paragraphs, idx=None):
    pxml = []
    for text, size, bold, bullet in paragraphs:
        ppr = '<a:pPr><a:buChar char="\u2022"/></a:pPr>' if bullet else ""
        rpr = '<a:rPr lang="zh-CN" sz="%d" dirty="0"%s/>' % (size, ' b="1"' if bold else "")
        pxml.append('<a:p>%s<a:r>%s<a:t>%s</a:t></a:r></a:p>' % (ppr, rpr, esc(text)))
    body = ('<a:bodyPr wrap="square"><a:normAutofit fontScale="100000" lnSpcReduction="10000"/></a:bodyPr>'
            '<a:lstStyle/>' + "".join(pxml))
    if ph_type:
        ph = '<p:ph type="%s"%s/>' % (ph_type, ' idx="%s"' % idx if idx else "")
    else:
        ph = "<p:ph/>"
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr>%s</p:nvPr></p:nvSpPr><p:spPr>%s</p:spPr><p:txBody>%s</p:txBody></p:sp>'
            % (shape_id, esc(name), ph, _xfrm(x, y, cx, cy), body))


def _picture(shape_id, name, rel_id, x, y, cx, cy):
    return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            '<p:spPr>%s<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
            % (shape_id, esc(name), rel_id, _xfrm(x, y, cx, cy)))


def _notes_xml(notes_text):
    body = ('<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="1400"/><a:t>%s</a:t></a:r></a:p>'
            % esc((notes_text or "").strip()))
    sp = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
          '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody>%s</p:txBody></p:sp>' % body)
    spTree = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
              '<p:grpSpPr/>%s</p:spTree>' % sp)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:notes xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld>%s</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
            % (NS_A, NS_R, NS_P, spTree))


def _rel(id_, reltype, target):
    return '<Relationship Id="%s" Type="%s" Target="%s"/>' % (id_, reltype, target)


def _rels_xml(items):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="%s">%s</Relationships>' % (REL_NS, "".join(items)))


# --- minimal from-scratch design parts ---
MIN_THEME = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<a:theme xmlns:a="%s" name="Office Theme"><a:themeElements>'
             '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
             '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2>'
             '<a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>'
             '<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>'
             '<a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5>'
             '<a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>'
             '<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>'
             '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/>'
             '</a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
             '<a:fmtScheme name="Office">'
             '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
             '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
             '<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
             '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'
             '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>'
             '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>'
             '<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
             '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
             '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>' % NS_A)


def _min_master():
    title = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
             '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr>%s</p:spPr>'
             '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>'
             % _xfrm(int(0.06 * SLIDE_W), int(0.05 * SLIDE_H), int(0.88 * SLIDE_W), int(0.18 * SLIDE_H)))
    body = ('<p:sp><p:nvSpPr><p:cNvPr id="3" name="Text Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>%s</p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>'
            % _xfrm(int(0.07 * SLIDE_W), int(0.27 * SLIDE_H), int(0.86 * SLIDE_W), int(0.66 * SLIDE_H)))
    spTree = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
              '<p:grpSpPr/>%s%s</p:spTree>' % (title, body))
    cSld = ('<p:cSld><p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>%s</p:cSld>' % spTree)
    tx = ('<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400" b="1"/></a:lvl1pPr></p:titleStyle>'
          '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:bodyStyle>'
          '<p:otherStyle><a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr></p:otherStyle></p:txStyles>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sldMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">%s'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3"'
            ' accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>%s</p:sldMaster>'
            % (NS_A, NS_R, NS_P, cSld, tx))


def _min_layout():
    title = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
             '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
             '<a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>')
    body = ('<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>')
    spTree = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
              '<p:grpSpPr/>%s%s</p:spTree>' % (title, body))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sldLayout xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld>%s</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>' % (NS_A, NS_R, NS_P, spTree))


def _min_notes_master():
    body = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:endParaRPr lang="zh-CN"/></a:p></p:txBody></p:sp>')
    spTree = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
              '<p:grpSpPr/>%s</p:spTree>' % body)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:notesMaster xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld>%s</p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3"'
            ' accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/></p:notesMaster>'
            % (NS_A, NS_R, NS_P, spTree))


def _min_presentation_xml():
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:presentation xmlns:a="%s" xmlns:r="%s" xmlns:p="%s">'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            '<p:notesMasterIdLst><p:notesMasterId r:id="rId2"/></p:notesMasterIdLst>'
            '<p:sldIdLst></p:sldIdLst>'
            '<p:sldSz cx="%d" cy="%d" type="screen16x9"/><p:notesSz cx="%d" cy="%d"/>'
            '<p:defaultTextStyle><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1800"/></a:lvl1pPr></p:defaultTextStyle>'
            '</p:presentation>' % (NS_A, NS_R, NS_P, SLIDE_W, SLIDE_H, NOTES_W, NOTES_H))


def _min_presentation_rels():
    return _rels_xml([
        _rel("rId1", RT_SLIDEMASTER, "slideMasters/slideMaster1.xml"),
        _rel("rId2", RT_NOTESMASTER, "notesMasters/notesMaster1.xml"),
    ])


def _min_content_types():
    defaults = ('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="png" ContentType="image/png"/>'
                '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                '<Default Extension="jpg" ContentType="image/jpeg"/>'
                '<Default Extension="gif" ContentType="image/gif"/>')
    overrides = ('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                 '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="%s"/>'
                 '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="%s"/>'
                 '<Override PartName="/ppt/theme/theme1.xml" ContentType="%s"/>'
                 '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="%s"/>'
                 % (CT_MASTER, CT_LAYOUT, CT_THEME, CT_NOTESMASTER))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="%s">%s%s</Types>' % (CT_NS, defaults, overrides))


def _root_rels():
    return _rels_xml([_rel("rId1", RT_OFFICE, "ppt/presentation.xml")])


# --- template handling ---
def _read_zip(path):
    z = zipfile.ZipFile(path)
    return {n: z.read(n) for n in z.namelist()}


def _parse_layout_geo(layout_xml):
    """Return {'title': (x,y,cx,cy), 'body:1': (...), 'body:2': (...), 'pic': (...)} in EMU."""
    geo = {}
    for m in re.finditer(r'<p:sp>(.*?)</p:sp>', layout_xml, re.S):
        sp = m.group(1)
        phm = re.search(r'<p:ph type="(\w+)"[^>]*?(?:idx="(\d+)")?', sp)
        if not phm:
            continue
        phtype = phm.group(1)
        idx = phm.group(2)
        xf = re.search(r'<a:off x="(\d+)" y="(\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>', sp)
        if not xf:
            continue
        key = phtype if not idx else "%s:%s" % (phtype, idx)
        geo[key] = tuple(int(v) for v in xf.groups())
    return geo


def _discover_layouts(parts):
    layouts = []
    for name in parts:
        m = re.match(r'ppt/slideLayouts/slideLayout(\d+)\.xml$', name)
        if not m:
            continue
        xml = parts[name].decode("utf-8", "replace")
        types = set(re.findall(r'<p:ph type="(\w+)"', xml))
        geo = _parse_layout_geo(xml)
        layouts.append({"path": name, "num": int(m.group(1)), "types": types, "geo": geo})
    layouts.sort(key=lambda l: l["num"])
    return layouts


def _layout_for_type(layouts, stype):
    """Choose the layout whose placeholder signature best matches the slide type."""
    if not layouts:
        return None
    # desired signatures, most specific first
    prefs = {
        "title":       [{"title", "ctrTitle"}, {"title", "subTitle"}, {"title"}],
        "section":     [{"title"}, {"title", "body"}],
        "outline":     [{"title", "body"}],
        "content":     [{"title", "body"}],
        "two_content": [{"title", "body"}, {"title", "body:1", "body:2"}],
        "image":       [{"title", "pic"}, {"title", "body"}, {"title"}],
        "thanks":      [{"title"}],
    }
    candidates = prefs.get(stype, [{"title", "body"}])
    # scoring: +2 per exact match, +1 for title, +1 for body; prefer layouts whose types
    # cover the wanted set; fall back to first layout.
    best, best_score = None, -1
    for lay in layouts:
        for want in candidates:
            score = 0
            for t in want:
                if t in lay["types"]:
                    score += 2
            if "title" in lay["types"] and "title" in want:
                score += 1
            if score > best_score:
                best, best_score = lay, score
    return best or layouts[0]


def _slide_geometry(stype, has_img, has_bullets, layout):
    """Final placement boxes; prefer template layout geometry, else hardcoded 16:9."""
    geo = (layout or {}).get("geo") or {}
    W, H = SLIDE_W, SLIDE_H
    out = {
        "title": geo.get("title") or geo.get("ctrTitle") or (int(0.06 * W), int(0.05 * H), int(0.88 * W), int(0.18 * H)),
    }
    # body positions (idx 1 and 2 if present)
    b1 = geo.get("body") or geo.get("body:1") or (int(0.07 * W), int(0.27 * H), int(0.86 * W), int(0.66 * H))
    b2 = geo.get("body:2")
    pic = geo.get("pic")
    out["body"] = b1
    out["body2"] = b2
    out["pic"] = pic
    if stype in ("title", "section", "thanks"):
        out["show_title_only"] = True
    else:
        out["show_title_only"] = False
    # image layouts
    if stype == "image" or (has_img and stype != "two_content"):
        if pic:
            out["img"] = pic
            out["body"] = b1 if has_bullets else None
        else:
            if has_bullets:
                out["img"] = (int(0.60 * W), int(0.27 * H), int(0.35 * W), int(0.66 * H))
                out["body"] = (int(0.07 * W), int(0.27 * H), int(0.50 * W), int(0.66 * H))
            else:
                out["img"] = (int(0.10 * W), int(0.30 * H), int(0.80 * W), int(0.62 * H))
    elif has_img and stype == "two_content":
        out["img"] = (int(0.60 * W), int(0.27 * H), int(0.35 * W), int(0.66 * H))
        out["body"] = b1 if has_bullets else None
        out["body2"] = None
    else:
        out["img"] = None
    return out


def _slide_xml(slide, has_img, layout):
    stype = slide.get("type") or "content"
    bullets = slide.get("bullets") or []
    has_bullets = bool(bullets)
    geo = _slide_geometry(stype, has_img, has_bullets, layout)
    shapes = []
    sid = 2
    shapes.append(_text_shape(sid, "Title 1", "title", *geo["title"],
                              [(slide.get("title") or "", 2800, True, False)]))
    sid += 1
    if not geo.get("show_title_only") and has_bullets:
        if stype == "two_content" and geo.get("body2"):
            half = len(bullets) // 2 or 1
            left = bullets[:half]
            right = bullets[half:]
            shapes.append(_text_shape(sid, "Content Placeholder 2", "body", *geo["body"],
                                      [(b, 1800, False, True) for b in left], idx=1))
            sid += 1
            shapes.append(_text_shape(sid, "Content Placeholder 3", "body", *geo["body2"],
                                      [(b, 1800, False, True) for b in right], idx=2))
            sid += 1
        elif geo.get("body"):
            shapes.append(_text_shape(sid, "Content Placeholder 2", "body", *geo["body"],
                                      [(b, 1800, False, True) for b in bullets], idx=1))
            sid += 1
    if has_img and geo.get("img"):
        shapes.append(_picture(sid, "Picture %d" % sid, "rId2", *geo["img"]))
        sid += 1
    spTree = ('<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
              '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
              '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
              + "".join(shapes) + "</p:spTree>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="%s" xmlns:r="%s" xmlns:p="%s"><p:cSld>%s</p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
            % (NS_A, NS_R, NS_P, spTree))


def _next_slide_id(pres_xml):
    ids = [int(m) for m in re.findall(r'<p:sldId[^>]*id="(\d+)"', pres_xml)]
    return (max(ids) + 1) if ids else 256


def _next_slide_part_number(parts):
    nums = [int(m.group(1)) for n in parts if (m := re.search(r'ppt/slides/slide(\d+)\.xml$', n))]
    return (max(nums) + 1) if nums else 1


def _next_rid(rels_xml):
    ids = [int(m) for m in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return (max(ids) + 1) if ids else 1


def _insert_sld_ids(pres_xml, pairs):
    adds = "".join('<p:sldId id="%d" r:id="rId%d"/>' % (sid, rid) for sid, rid in pairs)
    if "<p:sldIdLst>" in pres_xml and "</p:sldIdLst>" in pres_xml:
        pres_xml = pres_xml.replace("</p:sldIdLst>", adds + "</p:sldIdLst>", 1)
    else:
        pres_xml = pres_xml.replace("</p:presentation>", '<p:sldIdLst>%s</p:sldIdLst></p:presentation>' % adds, 1)
    return pres_xml


def _insert_rels(rels_xml, rels):
    adds = "".join(rels)
    if "</Relationships>" in rels_xml:
        return rels_xml.replace("</Relationships>", adds + "</Relationships>", 1)
    return rels_xml


def _insert_content_types(ct_xml, overrides, defaults):
    if "</Types>" not in ct_xml:
        return ct_xml
    inner = "".join(overrides) + "".join(defaults)
    return ct_xml.replace("</Types>", inner + "</Types>", 1)


# --- main pure-stdlib writer ---
def build_oo_xml(outline, template, out):
    slides = outline.get("slides") or []
    if not slides:
        raise ValueError("outline has no slides")

    if template and os.path.exists(template):
        parts = _read_zip(template)
        pres_xml = parts.get("ppt/presentation.xml", b"").decode("utf-8", "replace")
        pres_rels = parts.get("ppt/_rels/presentation.xml.rels", b"").decode("utf-8", "replace")
        ct_xml = parts.get("[Content_Types].xml", b"").decode("utf-8", "replace")
        layouts = _discover_layouts(parts)
    else:
        parts = {
            "ppt/presentation.xml": _min_presentation_xml().encode("utf-8"),
            "ppt/_rels/presentation.xml.rels": _min_presentation_rels().encode("utf-8"),
            "[Content_Types].xml": _min_content_types().encode("utf-8"),
            "_rels/.rels": _root_rels().encode("utf-8"),
            "ppt/slideMasters/slideMaster1.xml": _min_master().encode("utf-8"),
            "ppt/slideMasters/_rels/slideMaster1.xml.rels": _rels_xml([
                _rel("rId1", RT_SLIDELAYOUT, "../slideLayouts/slideLayout1.xml"),
                _rel("rId2", RT_THEME, "../theme/theme1.xml")]).encode("utf-8"),
            "ppt/slideLayouts/slideLayout1.xml": _min_layout().encode("utf-8"),
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels": _rels_xml([
                _rel("rId1", RT_SLIDEMASTER, "../slideMasters/slideMaster1.xml")]).encode("utf-8"),
            "ppt/theme/theme1.xml": MIN_THEME.encode("utf-8"),
            "ppt/notesMasters/notesMaster1.xml": _min_notes_master().encode("utf-8"),
            "ppt/notesMasters/_rels/notesMaster1.xml.rels": _rels_xml([
                _rel("rId1", RT_THEME, "../theme/theme1.xml")]).encode("utf-8"),
        }
        pres_xml = parts["ppt/presentation.xml"].decode("utf-8")
        pres_rels = parts["ppt/_rels/presentation.xml.rels"].decode("utf-8")
        ct_xml = parts["[Content_Types].xml"].decode("utf-8")
        layouts = []

    slide_id = _next_slide_id(pres_xml)
    rid = _next_rid(pres_rels)
    first_num = _next_slide_part_number(parts)
    pairs = []
    ct_overrides = []
    ct_defaults = []

    for i, slide in enumerate(slides, 1):
        num = first_num + i - 1
        stype = slide.get("type") or "content"
        layout = _layout_for_type(layouts, stype)
        layout_name = layout["path"] if layout else "ppt/slideLayouts/slideLayout1.xml"
        img = slide.get("image")
        has_img = bool(img and os.path.exists(img))
        media_name = None
        if has_img:
            ext = os.path.splitext(img)[1].lower().lstrip(".") or "png"
            if ext not in ("png", "jpg", "jpeg", "gif"):
                ext = "png"
            media_name = "ppt/media/image%d.%s" % (num, ext)
            parts[media_name] = open(img, "rb").read()
            ct_defaults.append('<Default Extension="%s" ContentType="image/%s"/>' % (ext, "jpeg" if ext in ("jpg", "jpeg") else ext))

        parts["ppt/slides/slide%d.xml" % num] = _slide_xml(slide, has_img, layout).encode("utf-8")
        srels = [_rel("rId1", RT_SLIDELAYOUT, "../slideLayouts/" + os.path.basename(layout_name))]
        if has_img:
            srels.append(_rel("rId2", RT_IMAGE, "../media/" + os.path.basename(media_name)))
        srels.append(_rel("rId3", RT_NOTESSLIDE, "../notesSlides/notesSlide%d.xml" % num))
        parts["ppt/slides/_rels/slide%d.xml.rels" % num] = _rels_xml(srels).encode("utf-8")

        parts["ppt/notesSlides/notesSlide%d.xml" % num] = _notes_xml(slide.get("notes") or "").encode("utf-8")
        parts["ppt/notesSlides/_rels/notesSlide%d.xml.rels" % num] = _rels_xml([
            _rel("rId1", RT_SLIDE, "../slides/slide%d.xml" % num)]).encode("utf-8")

        ct_overrides.append('<Override PartName="/ppt/slides/slide%d.xml" ContentType="%s"/>' % (num, CT_SLIDE))
        ct_overrides.append('<Override PartName="/ppt/notesSlides/notesSlide%d.xml" ContentType="%s"/>' % (num, CT_NOTESSLIDE))

        pairs.append((slide_id, rid))
        slide_id += 1
        rid += 1

    pres_xml = _insert_sld_ids(pres_xml, pairs)
    slide_rels = [_rel("rId%d" % rid, RT_SLIDE, "slides/slide%d.xml" % (first_num + idx - 1)) for idx, (_sid, rid) in enumerate(pairs, 1)]
    pres_rels = _insert_rels(pres_rels, slide_rels)
    ct_xml = _insert_content_types(ct_xml, ct_overrides, ct_defaults)

    parts["ppt/presentation.xml"] = pres_xml.encode("utf-8")
    parts["ppt/_rels/presentation.xml.rels"] = pres_rels.encode("utf-8")
    parts["[Content_Types].xml"] = ct_xml.encode("utf-8")
    if "_rels/.rels" not in parts:
        parts["_rels/.rels"] = _root_rels().encode("utf-8")

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(parts):
            z.writestr(name, parts[name])
    return out


# --- python-pptx path (higher fidelity when available) ---
def build_with_pptx(outline, template, out):
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.enum.shapes import PP_PLACEHOLDER
    prs = Presentation(template) if template and os.path.exists(template) else Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    layouts = list(prs.slide_layouts)

    def pick(stype):
        prefs = {
            "title": [{"title", "ctrTitle"}, {"title", "subtitle"}, {"title"}],
            "section": [{"title"}, {"title", "body"}],
            "outline": [{"title", "body"}],
            "content": [{"title", "body"}],
            "two_content": [{"title", "body"}, {"title", "body", "picture"}],
            "image": [{"title", "picture"}, {"title", "body"}, {"title"}],
            "thanks": [{"title"}],
        }
        best, best_score = None, -1
        for lay in layouts:
            types = {ph.placeholder_format.type for ph in lay.placeholders}
            for want in prefs.get(stype, [{"title", "body"}]):
                score = sum(2 for t in want if _ppt_type(t) in types)
                if "title" in types:
                    score += 1
                if score > best_score:
                    best, best_score = lay, score
        return best or (layouts[0] if layouts else None)

    def _ppt_type(name):
        return {"title": PP_PLACEHOLDER.TITLE, "ctrTitle": PP_PLACEHOLDER.CENTER_TITLE,
                "subtitle": PP_PLACEHOLDER.SUBTITLE, "body": PP_PLACEHOLDER.BODY,
                "picture": PP_PLACEHOLDER.PICTURE}.get(name, name)

    for idx, slide in enumerate(outline.get("slides") or []):
        stype = slide.get("type") or "content"
        lay = pick(stype)
        if lay is None:
            continue
        s = prs.slides.add_slide(lay)
        try:
            s.shapes.title.text = slide.get("title") or ""
        except Exception:
            pass
        bullets = slide.get("bullets") or []
        body_phs = [ph for ph in s.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.BODY]
        if body_phs:
            if stype == "two_content" and len(body_phs) >= 2:
                half = len(bullets) // 2 or 1
                for ph, chunk in zip(body_phs[:2], (bullets[:half], bullets[half:])):
                    tf = ph.text_frame
                    tf.clear()
                    for j, b in enumerate(chunk):
                        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                        p.text = b
                        p.level = 0
            else:
                ph = body_phs[0]
                tf = ph.text_frame
                tf.clear()
                for j, b in enumerate(bullets):
                    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                    p.text = b
                    p.level = 0
        img = slide.get("image")
        if img and os.path.exists(img):
            pic_phs = [ph for ph in s.placeholders if ph.placeholder_format.type == PP_PLACEHOLDER.PICTURE]
            try:
                if pic_phs:
                    pic_phs[0].insert_picture(img)
                else:
                    s.shapes.add_picture(img, Emu(int(0.60 * SLIDE_W)), Emu(int(0.27 * SLIDE_H)), width=Emu(int(0.35 * SLIDE_W)))
            except Exception:
                pass
        notes = slide.get("notes") or ""
        if notes:
            s.notes_slide.notes_text_frame.text = notes
    prs.save(out)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--outline", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--template", default=None)
    args = ap.parse_args()
    outline = json.load(open(args.outline, encoding="utf-8"))
    try:
        import pptx  # noqa: F401
        build_with_pptx(outline, args.template, args.out)
        print("built_with=python-pptx")
    except Exception:
        build_oo_xml(outline, args.template, args.out)
        print("built_with=pure-ooxml")


if __name__ == "__main__":
    main()
