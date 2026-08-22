#!/usr/bin/env python3
"""Fix DrawingML child-element ordering that PowerPoint strict-parses (repair prompt):
  1. <a:p>: endParaRPr must be the last child (runs inserted after it are illegal)
  2. CT_TextCharacterProperties (rPr/defRPr/endParaRPr): children in canonical order
  3. <a:tr>: extLst must come after all <a:tc>
Operates on an unpacked pptx dir in place.
"""
import sys, glob, os
from lxml import etree

A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
def q(t): return '{%s}%s' % (A, t)

# canonical child order of CT_TextCharacterProperties
RPR_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill',
             'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill',
             'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst']
RPR_KEY = {q(t): i for i, t in enumerate(RPR_ORDER)}

def fix_part(path):
    tree = etree.parse(path)
    root = tree.getroot()
    changed = 0
    # 1. a:p — endParaRPr last, pPr first
    for p in root.iter(q('p')):
        kids = list(p)
        eprs = [k for k in kids if k.tag == q('endParaRPr')]
        pprs = [k for k in kids if k.tag == q('pPr')]
        bad = False
        if eprs and kids and kids[-1] is not eprs[-1]:
            bad = True
        if pprs and kids and kids[0] is not pprs[0]:
            bad = True
        if bad:
            for e in eprs:
                p.remove(e)
            for e in pprs:
                p.remove(e)
            for e in reversed(pprs):
                p.insert(0, e)
            for e in eprs:
                p.append(e)          # keep only order; duplicates stay (schema allows 1; dupes unseen)
            changed += 1
    # 2. rPr / defRPr / endParaRPr child order
    for tag in ('rPr', 'defRPr', 'endParaRPr'):
        for r in root.iter(q(tag)):
            kids = list(r)
            keyed = [RPR_KEY.get(k.tag, 99) for k in kids]
            if keyed != sorted(keyed):
                order = sorted(range(len(kids)), key=lambda i: (keyed[i], i))
                for k in kids:
                    r.remove(k)
                for i in order:
                    r.append(kids[i])
                changed += 1
    # 3. a:tr — tc first, extLst last
    for tr in root.iter(q('tr')):
        kids = list(tr)
        tcs = [k for k in kids if k.tag == q('tc')]
        rest = [k for k in kids if k.tag != q('tc')]
        if rest and kids and (kids[-len(rest):] != rest or any(k.tag != q('tc') for k in kids[:len(tcs)])):
            for k in kids:
                tr.remove(k)
            for k in tcs:
                tr.append(k)
            for k in rest:
                tr.append(k)
            changed += 1
    if changed:
        tree.write(path, xml_declaration=True, encoding='UTF-8', standalone=True)
    return changed

if __name__ == '__main__':
    base = sys.argv[1]
    total = 0
    for pat in ('ppt/slides/slide*.xml', 'ppt/slideLayouts/slideLayout*.xml',
                'ppt/slideMasters/slideMaster*.xml', 'ppt/notesSlides/notesSlide*.xml',
                'ppt/notesMasters/notesMaster*.xml', 'ppt/handoutMasters/handoutMaster*.xml'):
        for f in glob.glob(os.path.join(base, pat)):
            c = fix_part(f)
            if c:
                print(f'  {os.path.basename(f)}: {c} fixes')
                total += c
    print('total fixes:', total)
