#!/usr/bin/env python3
"""Regenerate the site CV page and the Squarespace code block from Lindsey's CV docx.

Usage: python3 tools/build-cv.py /path/to/LDPrat_CV_XXX.docx
Reads hyperlinks and italics per paragraph from the docx itself.
Rewrites: dewittcv/index.html; in the vault media kit: cv-squarespace.html, and syncs cv.html.
"""
import re, sys, zipfile, html as H, os, shutil

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MEDIAKIT = "/Users/ledewitt/Vault/Lindsey's Brain/Lindsey-Media-Kit"

SECTIONS = [
    ("CURRENT POSITIONS", "current-positions", "Positions"),
    ("PEER-REVIEWED PUBLICATIONS", "peer-reviewed-publications", "Publications"),
    ("OTHER PUBLICATIONS (RECENT SELECTED)", "other-publications", "Other Writing"),
    ("AWARDS", "awards", "Awards"),
    ("RESEARCH LEADERSHIP", "research-leadership", "Research Leadership"),
    ("CONFERENCES AND SYMPOSIA ORGANIZED (RECENT SELECTED)", "conferences-organized", "Conferences Organized"),
    ("SPEAKING ENGAGEMENTS (RECENT SELECTED)", "speaking-engagements", "Speaking"),
    ("EDUCATION", "education", "Education"),
    ("ACADEMIC EMPLOYMENT", "academic-employment", "Academic Employment"),
    ("INVITED ACADEMIC LECTURES", "invited-academic-lectures", "Invited Lectures"),
    ("INVITED WORKSHOPS, SEMINARS, FORUMS, ROUNDTABLES", "invited-workshops", "Workshops & Forums"),
    ("PANELS AND ROUNDTABLES ORGANIZED", "panels-organized", "Panels Organized"),
    ("REFEREED CONFERENCE PRESENTATIONS (Selected)", "refereed-presentations", "Conference Papers"),
    ("ACADEMIC OUTREACH PRESENTATIONS", "academic-outreach", "Outreach"),
    ("TEACHING", "teaching", "Teaching"),
    ("SCHOLARSHIPS & FELLOWSHIPS", "scholarships-fellowships", "Fellowships"),
    ("GRANTS", "grants", "Grants"),
    ("SELECTED PROFESSIONAL EXPERIENCE AND TRAINING", "professional-training", "Professional Training"),
    ("IN THE MEDIA", "in-the-media", "Media"),
    ("LANGUAGES", "languages", "Languages"),
    ("PROFESSIONAL MEMBERSHIPS", "memberships", "Memberships"),
]
HEADER_IDS = {h: (i, t) for h, i, t in SECTIONS}
SUBHEADS = {"Peer-Reviewed Articles, Chapters & Proceedings", "Peer-Reviewed Articles & Chapters",
            "Monographs", "Translations (Japanese to English)", "2026", "2025", "Discussant",
            "Ghent University", "The Hebrew University of Jerusalem", "Freie Universität Berlin",
            "University of Tokyo", "Kyushu University (IMAP in Japanese Humanities)", "UCLA",
            "University of Washington"}
PREFIX = re.compile(r'^(\(under review\)|\(in press\)|\(translation in progress\)|'
                    r'(?:Spring|Summer|Fall|Winter) \d{4}|'
                    r'(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{4}|'
                    r'\d{4}–\d{1,4}|\d{4}–|\d{4})\s+(.+)$', re.S)

def esc(t): return t.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def run_seg(run):
    text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', run))
    if '<w:tab/>' in run: text = ' ' + text
    rpr = re.search(r'<w:rPr>(.*?)</w:rPr>', run, re.S)
    ital = bool(rpr and re.search(r'<w:i(?:\s+w:val="(?:1|true)")?\s*/>', rpr.group(1)))
    return [H.unescape(text), ital, None]

def parse_paragraphs(docx_path):
    z = zipfile.ZipFile(docx_path)
    rels = z.read("word/_rels/document.xml.rels").decode("utf8")
    relmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    doc = z.read("word/document.xml").decode("utf8")
    paras = []
    for pm in re.finditer(r'<w:p[ >].*?</w:p>', doc, re.S):
        p, segs, pos = pm.group(0), [], 0
        for hm in re.finditer(r'<w:hyperlink[^>]*?r:id="([^"]+)"[^>]*>(.*?)</w:hyperlink>', p, re.S):
            for rm in re.finditer(r'<w:r[ >].*?</w:r>', p[pos:hm.start()], re.S):
                segs.append(run_seg(rm.group(0)))
            url = relmap.get(hm.group(1), '')
            text = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', hm.group(2)))
            ital = bool(re.search(r'<w:i(?:\s[^>]*)?/>', hm.group(2)))
            if text:
                segs.append([H.unescape(text), ital, url if url.startswith('http') else None])
            pos = hm.end()
        for rm in re.finditer(r'<w:r[ >].*?</w:r>', p[pos:], re.S):
            segs.append(run_seg(rm.group(0)))
        segs = [s for s in segs if s[0]]
        if segs: paras.append(segs)
    return paras

def render(segs):
    merged = []
    for t, i, u in segs:
        if merged and merged[-1][1] == i and merged[-1][2] == u:
            merged[-1][0] += t
        else:
            merged.append([t, i, u])
    out = ''
    for t, i, u in merged:
        h = esc(t)
        if i: h = f'<em>{h}</em>'
        if u: h = f'<a href="{esc(u)}">{h}</a>'
        out += h
    return re.sub(r'  +', ' ', out).strip()

def plain(segs): return re.sub(r'\s+', ' ', ''.join(s[0] for s in segs)).strip()

ORDER = ["current-positions", "education", "academic-employment", "languages",
         "research-leadership", "peer-reviewed-publications", "other-publications",
         "awards", "conferences-organized", "speaking-engagements",
         "invited-academic-lectures", "invited-workshops", "panels-organized",
         "refereed-presentations", "academic-outreach", "teaching",
         "scholarships-fellowships", "grants", "professional-training",
         "in-the-media", "memberships"]

def build(docx_path):
    cur, secs = None, {}
    for segs in parse_paragraphs(docx_path):
        txt = plain(segs)
        if txt in HEADER_IDS:
            sid, label = HEADER_IDS[txt]
            cur = sid
            secs[sid] = {"title": txt, "label": label, "rows": []}
        elif cur is None:
            continue
        elif txt in SUBHEADS:
            secs[cur]["rows"].append(f'<h3>{esc(txt)}</h3>')
        else:
            h = render(segs)
            m = PREFIX.match(h)
            if m: h = f'<strong>{m.group(1)}</strong> {m.group(2)}'
            secs[cur]["rows"].append(f'<p class="item">{h}</p>')
    order = [i for i in ORDER if i in secs] + [i for i in secs if i not in ORDER]
    out, toc = [], []
    for sid in order:
        d = secs[sid]
        out.append(f'<section id="{sid}"><h2>{esc(d["title"])}</h2>')
        out.extend(d["rows"])
        out.append('</section>')
        toc.append(f'<a href="#{sid}">{esc(d["label"])}</a>')
    body = '\n'.join(out)
    # site ruling 2026-08-09: under-review entry names no journal
    body = body.replace('\u201d <em>ACM Journal on Responsible Computing</em>.', '\u201d')
    tochtml = '<nav class="toc" aria-label="Sections">' + ''.join(toc) + '</nav>'
    assert body.count('<a ') == body.count('</a>'), "unbalanced anchors"
    assert not re.search(r'href="[^"]*<', body), "tag inside href"
    return body, tochtml

if __name__ == '__main__':
    body, tochtml = build(sys.argv[1])
    for path in [os.path.join(REPO, 'dewittcv/index.html'), os.path.join(MEDIAKIT, 'cv-squarespace.html')]:
        s = open(path).read()
        s = re.sub(r'<nav class="toc"[^>]*>.*?</nav>', lambda m: tochtml, s, count=1, flags=re.S)
        head_end = s.index('</header>', s.index('class="masthead"')) + len('</header>')
        s = s[:head_end] + '\n' + body + '\n' + s[s.index('<footer class="updated">'):]
        open(path, 'w').write(s)
        print('spliced:', path)
    shutil.copy(os.path.join(REPO, 'dewittcv/index.html'), os.path.join(MEDIAKIT, 'cv.html'))
    print('synced cv.html | sections:', body.count('<section'), '| items:', body.count('<p class="item">'), '| links:', body.count('<a href'))
