#!/usr/bin/env python3
"""Strip think-cell/collab leftovers from a pptx built by cloning template slides.

Removes per slide/layout/master:
  - hidden think-cell OLE graphicFrames (progId TCLayout.*, hidden ghost objects)
  - p:custDataLst (think-cell tag references)
  - rels of type oleObject / vmlDrawing / tags, plus rels that became unused
Removes package-level:
  - ppt/tags/*, ppt/changesInfos/*, ppt/revisionInfo.xml, ppt/drawings/vmlDrawing*.vml,
    ppt/embeddings/* that are no longer referenced, and their content-type overrides
Repacks with [Content_Types].xml first and no directory entries.
"""
import sys, os, re, glob, shutil, zipfile, posixpath
from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}
RID = '{%s}id' % NS['r']

def strip_part(xml_path):
    """Remove think-cell OLE frames + custDataLst from one slide/layout/master."""
    tree = etree.parse(xml_path)
    root = tree.getroot()
    removed = 0
    for gf in root.findall('.//p:graphicFrame', NS):
        gd = gf.find('.//a:graphicData', NS)
        if gd is None:
            continue
        if gd.get('uri', '').endswith('/ole'):
            txt = etree.tostring(gd, encoding='unicode')
            if 'TCLayout' in txt or 'think-cell' in txt:
                gf.getparent().remove(gf)
                removed += 1
    for cd in root.findall('.//p:custDataLst', NS):
        cd.getparent().remove(cd)
        removed += 1
    if removed:
        tree.write(xml_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    return removed

def used_rids(xml_path):
    txt = open(xml_path, encoding='utf-8').read()
    return set(re.findall(r'r:(?:id|embed|link|pict|dm|lo|qs|cs)="(rId\d+)"', txt))

DROP_TYPES = ('/oleObject', '/vmlDrawing', '/tags', '/revisionInfo', '/changesInfo', '/changesInfos')
KEEP_UNREF = ('/slideLayout', '/slideMaster', '/notesSlide', '/notesMaster', '/handoutMaster',
              '/theme', '/tableStyles', '/presProps', '/viewProps', '/commentAuthors', '/slide')

def fix_rels(rels_path, part_xml):
    """Drop think-cell rel types; drop other rels no longer referenced from the part."""
    if not os.path.exists(rels_path):
        return []
    tree = etree.parse(rels_path)
    root = tree.getroot()
    used = used_rids(part_xml) if part_xml else set()
    dropped = []
    for r in list(root):
        typ, rid = r.get('Type', ''), r.get('Id', '')
        tgt = r.get('Target', '')
        if any(typ.endswith(t) for t in DROP_TYPES):
            root.remove(r); dropped.append(tgt); continue
        if r.get('TargetMode') == 'External':
            continue
        if any(typ.endswith(t) for t in KEEP_UNREF):
            continue
        if rid not in used:  # e.g. EMF fallback image of a removed OLE frame
            root.remove(r); dropped.append(tgt)
    tree.write(rels_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    return dropped

def main(src, dst):
    work = dst + '.unpacked'
    if os.path.exists(work):
        shutil.rmtree(work)
    zipfile.ZipFile(src).extractall(work)

    total = 0
    for part in glob.glob(work + '/ppt/slides/slide*.xml') + \
                glob.glob(work + '/ppt/slideLayouts/slideLayout*.xml') + \
                glob.glob(work + '/ppt/slideMasters/slideMaster*.xml') + \
                glob.glob(work + '/ppt/notesMasters/notesMaster*.xml') + \
                glob.glob(work + '/ppt/handoutMasters/handoutMaster*.xml'):
        total += strip_part(part)
        d = os.path.join(os.path.dirname(part), '_rels', os.path.basename(part) + '.rels')
        fix_rels(d, part)
    # presentation.xml itself may carry a p:custDataLst (think-cell tags) — strip it too
    total += strip_part(work + '/ppt/presentation.xml')
    # presentation-level rels: drop revisionInfo/changesInfos/tags refs
    fix_rels(work + '/ppt/_rels/presentation.xml.rels', work + '/ppt/presentation.xml')
    print(f'{src}: removed {total} think-cell elements')

    # delete now-orphaned parts
    victims = []
    for pat in ('ppt/tags/*', 'ppt/changesInfos/*', 'ppt/revisionInfo.xml',
                'ppt/drawings/vmlDrawing*.vml'):
        victims += glob.glob(os.path.join(work, pat))
    # embeddings/media: keep only ones still referenced by some rels file
    referenced = set()
    for rels in glob.glob(work + '/**/_rels/*.rels', recursive=True):
        base = posixpath.dirname(posixpath.dirname(os.path.relpath(rels, work).replace(os.sep, '/')))
        for tgt in re.findall(r'Target="([^"]+)"', open(rels, encoding='utf-8').read()):
            if tgt.startswith('http'):
                continue
            referenced.add(posixpath.normpath(posixpath.join(base, tgt)))
    for pat in ('ppt/embeddings/*', 'ppt/media/*'):
        for f in glob.glob(os.path.join(work, pat)):
            relp = os.path.relpath(f, work).replace(os.sep, '/')
            if relp not in referenced:
                victims.append(f)
    gone = set()
    for v in victims:
        relp = os.path.relpath(v, work).replace(os.sep, '/')
        if relp in referenced and 'tags/' not in relp and 'changesInfos' not in relp and 'revisionInfo' not in relp and '.vml' not in relp:
            continue
        os.remove(v)
        gone.add('/' + relp)
    # drop empty rels dirs' stale rels for deleted vml? (none reference them anymore)
    # content types: remove overrides for deleted parts
    ct_path = os.path.join(work, '[Content_Types].xml')
    ct = etree.parse(ct_path)
    for o in list(ct.getroot()):
        pn = o.get('PartName')
        if pn and (pn in gone or not os.path.exists(os.path.join(work, pn.lstrip('/')))):
            if pn.startswith('/ppt/') and o.tag.endswith('Override'):
                ct.getroot().remove(o)
    ct.write(ct_path, xml_declaration=True, encoding='UTF-8', standalone=True)
    print(f'  deleted {len(gone)} orphan parts')

    # repack: [Content_Types].xml first, no directory entries
    if os.path.exists(dst):
        os.remove(dst)
    files = []
    for r, _, fs in os.walk(work):
        for f in fs:
            full = os.path.join(r, f)
            files.append(os.path.relpath(full, work).replace(os.sep, '/'))
    files.sort(key=lambda n: (n != '[Content_Types].xml', not n.startswith('_rels/'), n))
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
        for n in files:
            z.write(os.path.join(work, n), n)
    print(f'  wrote {dst} ({os.path.getsize(dst)} bytes, {len(files)} entries)')

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
