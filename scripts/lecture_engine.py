# -*- coding: utf-8 -*-
"""
lecture_engine.py — Build student + teacher revision lectures (.docx) by MERGING
the original XML of source Word files, so every figure (image) and formula (OMML)
is preserved exactly rather than retyped.

Driven by a JSON config (see references/config_schema.md). Usage:

    python lecture_engine.py --config config.json [--open]

It produces the student file and (if a solution_file is configured) the teacher file.
"""
import zipfile, re, os, shutil, html, json, sys, argparse

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
IMG_T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
OLE_T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'

# ---------------------------------------------------------------- low level
def body_children(doc):
    """Split <w:body> into a list of (tag, raw_xml) top-level elements, exactly
    as written (no re-serialization → perfect fidelity for figures/formulas)."""
    bstart = doc.index('<w:body'); bstart = doc.index('>', bstart) + 1
    bend = doc.rindex('</w:body>')
    inner = doc[bstart:bend]
    out = []; i = 0; n = len(inner)
    while i < n:
        if inner[i] != '<': i += 1; continue
        m = re.match(r'<(w:[a-zA-Z]+)', inner[i:])
        if not m: i += 1; continue
        tag = m.group(1)
        sc = re.match(r'<' + tag + r'[^>]*/>', inner[i:])
        if sc:
            out.append((tag, inner[i:i + sc.end()])); i += sc.end(); continue
        depth = 0; placed = False
        for mo in re.compile(r'<(/?)' + tag + r'(\s|>|/)').finditer(inner, i):
            if mo.group(1) == '':
                seg = inner[mo.start():inner.index('>', mo.start()) + 1]
                if seg.endswith('/>'): continue
                depth += 1
            else:
                depth -= 1
                if depth == 0:
                    end = inner.index('>', mo.start()) + 1
                    out.append((tag, inner[i:end])); i = end; placed = True; break
        if not placed: break
    return out

def load(fn):
    z = zipfile.ZipFile(fn)
    return z, z.read('word/document.xml').decode('utf-8')

def parse_rels(z):
    rels = z.read('word/_rels/document.xml.rels').decode('utf-8')
    d = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels):
        s = m.group(0)
        rid = re.search(r'Id="([^"]+)"', s).group(1)
        tgt = re.search(r'Target="([^"]+)"', s).group(1)
        mode = re.search(r'TargetMode="([^"]+)"', s)
        d[rid] = (tgt, mode.group(1) if mode else 'Internal')
    return d

def childmap(doc):
    return [(t, s) for t, s in body_children(doc) if t in ('w:p', 'w:tbl')]

def firsttext(s):
    m = re.search(r'<w:t[^>]*>(.*?)</w:t>', s)
    return m.group(1) if m else ''

def fulltext(s):
    # all run text in document order, tags stripped — robust to leading empty/space runs
    return re.sub(r'<[^>]+>', '', ''.join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', s, re.S)))

def is_start(s):
    # a problem starts with a number followed by the full-width period ．(or ASCII .)
    return bool(re.match(r'^\s*\d+[．.]', fulltext(s)))

def build_probs(PT, segments):
    """segments: [{'kao':k,'start':a,'end':b,'type':'choice'|'solve'}].
    Returns {(kao,num): {'idxs':[...], 'type':...}} keyed by detected number."""
    probs = {}
    for seg in segments:
        kao, a, b = seg['kao'], seg['start'], seg['end']
        typ = seg.get('type', 'solve')
        starts = [i for i in range(a, b + 1) if PT[i][0] == 'w:p' and is_start(PT[i][1])]
        for k, st in enumerate(starts):
            num = int(re.match(r'^\s*(\d+)', fulltext(PT[st][1])).group(1))
            end = (starts[k + 1] - 1) if k + 1 < len(starts) else b
            probs[(kao, num)] = {'idxs': list(range(st, end + 1)), 'type': typ}
    return probs

# ---------------------------------------------------------------- OMML formulas
def _mrun(t, it=False):
    return '<m:r><m:rPr><m:sty m:val="%s"/></m:rPr><m:t>%s</m:t></m:r>' % (
        'i' if it else 'p', html.escape(t))

def _render_tokens(tokens):
    """tokens: a string, or a list whose items are strings / dicts.
    dict forms: {"text":s} {"it":s} {"frac":[num,den]} {"sup":[b,e]} {"sub":[b,e]} {"sqrt":x}"""
    if isinstance(tokens, str):
        return _mrun(tokens)
    out = []
    for tok in tokens:
        if isinstance(tok, str):
            out.append(_mrun(tok))
        elif 'text' in tok:
            out.append(_mrun(tok['text']))
        elif 'it' in tok:
            out.append(_mrun(tok['it'], it=True))
        elif 'frac' in tok:
            n, d = tok['frac']
            out.append('<m:f><m:num>%s</m:num><m:den>%s</m:den></m:f>' % (
                _render_tokens(n), _render_tokens(d)))
        elif 'sup' in tok:
            b, e = tok['sup']
            out.append('<m:sSup><m:e>%s</m:e><m:sup>%s</m:sup></m:sSup>' % (
                _render_tokens(b), _render_tokens(e)))
        elif 'sub' in tok:
            b, e = tok['sub']
            out.append('<m:sSub><m:e>%s</m:e><m:sub>%s</m:sub></m:sSub>' % (
                _render_tokens(b), _render_tokens(e)))
        elif 'sqrt' in tok:
            out.append('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>%s</m:e></m:rad>' %
                       _render_tokens(tok['sqrt']))
    return ''.join(out)

# ---------------------------------------------------------------- main builder
def build(config, mode, outname, base_dir):
    teacher = (mode == 'teacher')
    def rp(p): return p if os.path.isabs(p) else os.path.join(base_dir, p)

    zk, dock = load(rp(config['knowledge_file'])); rels_k = parse_rels(zk); PTK = childmap(dock)
    zq, docq = load(rp(config['question_file'])); rels_q = parse_rels(zq); PTQ = childmap(docq)
    PROB_Q = build_probs(PTQ, config['question_segments'])
    sol = config.get('solution_file')
    if sol:
        zj, docj = load(rp(sol)); rels_j = parse_rels(zj)
        PTJ = childmap(docj); PROB_J = build_probs(PTJ, config.get('solution_segments', config['question_segments']))
        base_doc = docj
    else:
        zj = zq; PTJ = PTQ; rels_j = rels_q; PROB_J = PROB_Q; base_doc = docq

    BUILD = '_pkg_' + mode + '_' + str(os.getpid())
    if os.path.exists(BUILD): shutil.rmtree(BUILD)
    os.makedirs(BUILD)
    media = {}; rel_entries = []; exts = set(); cnt = [100]
    def nr():
        cnt[0] += 1; return 'rId%d' % cnt[0]
    def remap(s, zz, rr):
        for old in set(re.findall(r'r:(?:embed|id|link)="([^"]+)"', s)):
            if old not in rr: continue
            tgt, m2 = rr[old]
            if m2 == 'External':
                r = nr(); rel_entries.append((r, tgt, IMG_T, 'External')); s = s.replace('"%s"' % old, '"%s"' % r); continue
            try: data = zz.read('word/' + tgt.lstrip('/'))
            except KeyError: continue
            ext = tgt.rsplit('.', 1)[-1].lower(); idx = len(media) + 1
            if 'embeddings' in tgt or ext == 'bin':
                nm = 'embeddings/oleObject%d.bin' % idx; media[nm] = data
                r = nr(); rel_entries.append((r, nm, OLE_T, 'Internal'))
            else:
                nm = 'media/image%d.%s' % (idx, ext); media[nm] = data; exts.add(ext)
                r = nr(); rel_entries.append((r, nm, IMG_T, 'Internal'))
            s = s.replace('"%s"' % old, '"%s"' % r)
        return s
    def clean(s, dg=False):
        s = re.sub(r'<w:bookmarkStart\b[^>]*/>', '', s); s = re.sub(r'<w:bookmarkEnd\b[^>]*/>', '', s)
        s = re.sub(r'<w:proofErr\b[^>]*/>', '', s); s = re.sub(r'<w:numPr>.*?</w:numPr>', '', s, flags=re.S)
        s = re.sub(r'<w:pStyle w:val="Heading\d"\s*/>', '', s); s = re.sub(r'<w:outlineLvl\b[^>]*/>', '', s)
        if dg: s = re.sub(r'<w:color w:val="00B050"\s*/>', '', s)
        return s
    def renum(s, k):
        # replace the leading "数字．" (which may be split across several runs) with 例K
        lead = re.match(r'^\s*\d+[．.]', fulltext(s))
        if not lead:
            return s
        rem = len(lead.group(0))
        out = []; last = 0; inserted = False
        for mt in re.finditer(r'(<w:t[^>]*>)(.*?)(</w:t>)', s, re.S):
            out.append(s[last:mt.start()])
            o, txt, c = mt.group(1), mt.group(2), mt.group(3)
            if rem > 0:
                cut = min(rem, len(txt)); txt = txt[cut:]; rem -= cut
            if not inserted:
                txt = '例%d　' % k + txt; inserted = True
            out.append(o + txt + c); last = mt.end()
        out.append(s[last:])
        return ''.join(out)

    parts = []; P = parts.append
    HEI = '<w:rFonts w:ascii="SimHei" w:eastAsia="SimHei" w:hAnsi="SimHei"/>'
    SONG = '<w:rFonts w:ascii="SimSun" w:eastAsia="SimSun" w:hAnsi="SimSun"/>'
    def run(t, font=SONG, sz=24, b=False, color=None, it=False):
        rp_ = '<w:rPr>' + font + ('<w:b/>' if b else '') + ('<w:i/>' if it else '')
        if color: rp_ += '<w:color w:val="%s"/>' % color
        rp_ += '<w:sz w:val="%d"/><w:szCs w:val="%d"/></w:rPr>' % (sz, sz)
        return '<w:r>' + rp_ + '<w:t xml:space="preserve">' + html.escape(t) + '</w:t></w:r>'
    def para(rs, jc=None, line=360, before=0, after=0, style=None, outline=None, pbb=False):
        p = '<w:p><w:pPr>'
        if style: p += '<w:pStyle w:val="%s"/>' % style
        if pbb: p += '<w:pageBreakBefore/>'
        sp = '<w:spacing'
        if before: sp += ' w:before="%d"' % before
        if after: sp += ' w:after="%d"' % after
        sp += ' w:line="%d" w:lineRule="auto"/>' % line; p += sp
        if jc: p += '<w:jc w:val="%s"/>' % jc
        if outline is not None: p += '<w:outlineLvl w:val="%d"/>' % outline
        p += '</w:pPr>' + ''.join(rs) + '</w:p>'; return p
    def blank(n=1):
        for _ in range(n): P('<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:p>')
    def ruled(n=1, line=480):
        for _ in range(n):
            P('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="2" w:color="BFBFBF"/></w:pBdr>'
              '<w:spacing w:before="60" w:line="%d" w:lineRule="auto"/></w:pPr></w:p>' % line)
    def sep():
        P('<w:p><w:pPr><w:pBdr><w:bottom w:val="dotted" w:sz="6" w:space="4" w:color="9DC3E6"/></w:pBdr>'
          '<w:spacing w:before="100" w:after="100" w:line="240" w:lineRule="auto"/></w:pPr></w:p>')
    def h1(t): P(para([run(t, font=HEI, sz=32, b=True)], jc='center', before=240, after=180, style='Heading1', outline=0, pbb=True))
    def label(t, m='◆'): P(para([run(m + ' ' + t, font=HEI, sz=26, b=True, color='1F3864')], before=140, after=70))
    def omath(inner): P('<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
                        '<m:oMathPara><m:oMath>%s</m:oMath></m:oMathPara></w:p>' % inner)

    title = config.get('title', '复习讲义')
    subtitle = config.get('subtitle', '')
    # ---- cover ----
    blank(3)
    P(para([run(title, font=HEI, sz=52, b=True)], jc='center', after=120))
    P('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="6" w:color="1F3864"/></w:pBdr>'
      '<w:spacing w:line="360" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr></w:p>')
    blank(1)
    if subtitle:
        P(para([run(subtitle, font=HEI, sz=36, b=True, color='1F3864')], jc='center', after=120))
    blank(2)
    if teacher:
        P(para([run('（教师版 · 含答案与详解）', font=HEI, sz=30, b=True, color='C00000')], jc='center', after=120))
    else:
        P(para([run('（学生版 · 知识 + 例题）', font=HEI, sz=30, b=True, color='1F3864')], jc='center', after=120))
    blank(3)
    P(para([run('★ 期末复习 · 刷题讲义 ★', font=SONG, sz=28, color='808080')], jc='center'))
    blank(5)
    fields = (['任课教师：', '班　级：', '日　期：'] if teacher else ['班　级：', '姓　名：', '日　期：'])
    for f in fields:
        P(para([run(f, font=SONG, sz=30), run('______________________', font=SONG, sz=30)], jc='center', after=120))
    blank(3)
    P(para([run('━' * 20, font=SONG, sz=24, color='1F3864')], jc='center'))
    # ---- TOC ----
    P(para([run('目　录', font=HEI, sz=32, b=True)], jc='center', before=120, after=200, pbb=True))
    P('<w:p><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
      '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
      '<w:r><w:rPr>' + SONG + '<w:sz w:val="24"/></w:rPr>'
      '<w:t>请在 Word/WPS 中右键此处选择“更新域”，或按 F9 生成目录与页码。</w:t></w:r>'
      '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')

    # ---- sections ----
    def emit_problems(prob_list):
        if not prob_list: return
        label('对应例题训练', '◆')
        for k, pr in enumerate(prob_list, 1):
            key = (pr['kao'], pr['num'])
            if teacher and key in PROB_J:
                rec = PROB_J[key]; PT = PTJ; zz = zj; rr = rels_j
            elif key in PROB_Q:
                rec = PROB_Q[key]; PT = PTQ; zz = zq; rr = rels_q
            else:
                continue
            for j, i in enumerate(rec['idxs']):
                s = clean(PT[i][1], dg=True)
                if j == 0: s = renum(s, k)
                P(remap(s, zz, rr))
            if not teacher:
                if rec['type'] == 'choice':
                    P(para([run('【答案】', font=HEI, sz=24, b=True, color='1F3864'),
                            run('  ' + '_' * 44, font=SONG, sz=24)], before=40, after=60))
                else:
                    P(para([run('【答题区】', font=HEI, sz=24, b=True, color='1F3864')], before=40, after=40))
                    ruled(config.get('answer_lines', 5))
            sep()

    for sec in config['sections']:
        h1(sec['title'])
        if sec.get('knowledge'):
            label('核心知识', '◆')
            for rng in sec['knowledge']:
                for i in range(rng[0], rng[1] + 1):
                    P(remap(clean(PTK[i][1]), zk, rels_k))
        if sec.get('formulas'):
            label('重点公式', '◆')
            for f in sec['formulas']:
                omath(_render_tokens(f))
        label('笔记区', '★'); ruled(config.get('note_lines', 2))
        emit_problems(sec.get('problems', []))

    # ---- section properties (A4 portrait, 2cm/2.5cm margins) ----
    sectpr = ('<w:sectPr><w:footerReference w:type="default" r:id="rId10"/>'
              '<w:footerReference w:type="first" r:id="rId11"/>'
              '<w:pgSz w:w="11906" w:h="16838"/>'
              '<w:pgMar w:top="1134" w:right="1417" w:bottom="1134" w:left="1417" '
              'w:header="851" w:footer="850" w:gutter="0"/>'
              '<w:cols w:space="425"/><w:titlePg/>'
              '<w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>')
    body = ''.join(parts) + sectpr
    ro = base_doc[base_doc.index('<w:document'):base_doc.index('>', base_doc.index('<w:document')) + 1]
    for need, uri in [('xmlns:asvg', 'http://schemas.microsoft.com/office/drawing/2016/SVG/main'),
                      ('xmlns:aink', 'http://schemas.microsoft.com/office/drawing/2016/ink'),
                      ('xmlns:a14', 'http://schemas.microsoft.com/office/drawing/2010/main'),
                      ('xmlns:am3d', 'http://schemas.microsoft.com/office/drawing/2017/model3d')]:
        if need not in ro: ro = ro[:-1] + ' %s="%s"' % (need, uri) + '>'
    document_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + ro + '<w:body>' + body + '</w:body></w:document>'

    def wr(path, data):
        full = os.path.join(BUILD, path); os.makedirs(os.path.dirname(full), exist_ok=True)
        open(full, 'wb').write(data.encode('utf-8') if isinstance(data, str) else data)
    wr('word/document.xml', document_xml)
    for nm, data in media.items(): wr('word/' + nm, data)
    wr('word/styles.xml', zj.read('word/styles.xml'))
    wr('word/theme/theme1.xml', zj.read('word/theme/theme1.xml'))
    wr('word/fontTable.xml', zj.read('word/fontTable.xml'))
    wr('word/settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       '<w:settings xmlns:w="%s"><w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/>'
       '<w:updateFields w:val="true"/><w:compat><w:compatSetting w:name="compatibilityMode" '
       'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>' % W)
    ftxt = '教师版' if teacher else '学生版'
    wr('word/footer1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       '<w:ftr xmlns:w="%s"><w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/><w:jc w:val="center"/></w:pPr>'
       '<w:r><w:rPr>%s<w:sz w:val="18"/><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve">— 第 </w:t></w:r>'
       '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
       '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
       '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
       '<w:r><w:rPr><w:sz w:val="18"/><w:color w:val="808080"/></w:rPr><w:t>1</w:t></w:r>'
       '<w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
       '<w:r><w:rPr>%s<w:sz w:val="18"/><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve"> 页 · %s —</w:t></w:r>'
       '</w:p></w:ftr>' % (W, SONG, SONG, ftxt))
    wr('word/footer2.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       '<w:ftr xmlns:w="%s"><w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:p></w:ftr>' % W)
    fixed = [('rId1', 'styles.xml', 'styles'), ('rId2', 'settings.xml', 'settings'),
             ('rId3', 'fontTable.xml', 'fontTable'), ('rId4', 'theme/theme1.xml', 'theme'),
             ('rId10', 'footer1.xml', 'footer'), ('rId11', 'footer2.xml', 'footer')]
    rx = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for rid, tgt, t in fixed:
        rx.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/%s" Target="%s"/>' % (rid, t, tgt))
    for rid, tgt, typ, m2 in rel_entries:
        if m2 == 'External':
            rx.append('<Relationship Id="%s" Type="%s" Target="%s" TargetMode="External"/>' % (rid, typ, tgt))
        else:
            rx.append('<Relationship Id="%s" Type="%s" Target="%s"/>' % (rid, typ, tgt))
    rx.append('</Relationships>'); wr('word/_rels/document.xml.rels', '\n'.join(rx))
    wr('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
       '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
       '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
       '</Relationships>')
    wr('docProps/core.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
       '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
       'xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>%s（%s）</dc:title>'
       '<dc:creator>lecture-builder</dc:creator></cp:coreProperties>' % (html.escape(title), ftxt))
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>']
    cm = {'png': 'image/png', 'jpeg': 'image/jpeg', 'jpg': 'image/jpeg', 'gif': 'image/gif',
          'tiff': 'image/tiff', 'wmf': 'image/x-wmf', 'emf': 'image/x-emf', 'svg': 'image/svg+xml', 'bmp': 'image/bmp'}
    for e in sorted(exts):
        ct.append('<Default Extension="%s" ContentType="%s"/>' % (e, cm.get(e, 'application/octet-stream')))
    if any(n.startswith('embeddings/') for n in media):
        ct.append('<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>')
    for pn, c in [('/word/document.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'),
                  ('/word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'),
                  ('/word/settings.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'),
                  ('/word/fontTable.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'),
                  ('/word/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'),
                  ('/word/footer1.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'),
                  ('/word/footer2.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'),
                  ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml')]:
        ct.append('<Override PartName="%s" ContentType="%s"/>' % (pn, c))
    ct.append('</Types>'); wr('[Content_Types].xml', '\n'.join(ct))

    outpath = rp(outname)
    if os.path.exists(outpath):
        try: os.remove(outpath)
        except PermissionError:
            raise SystemExit('无法写入 %s —— 该文件正在 Word/WPS 中打开，请先关闭后重试。' % outpath)
    zf = zipfile.ZipFile(outpath, 'w', zipfile.ZIP_DEFLATED)
    zf.writestr('[Content_Types].xml', open(os.path.join(BUILD, '[Content_Types].xml'), 'rb').read())
    for root, _, files in os.walk(BUILD):
        for f in files:
            arc = os.path.relpath(os.path.join(root, f), BUILD).replace('\\', '/')
            if arc == '[Content_Types].xml': continue
            zf.writestr(arc, open(os.path.join(root, f), 'rb').read())
    zf.close(); shutil.rmtree(BUILD)
    print('WROTE %s  | media=%d  paragraphs=%d' % (outpath, len(media), body.count('<w:p>') + body.count('<w:p ')))
    return outpath

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--config', required=True)
    ap.add_argument('--open', action='store_true', help='open results in default app (WPS/Word) when done')
    args = ap.parse_args()
    with open(args.config, encoding='utf-8') as f:
        cfg = json.load(f)
    base = os.path.dirname(os.path.abspath(args.config))
    outs = build(cfg, 'student', cfg.get('output_student', 'lecture.docx'), base)
    results = [outs]
    if cfg.get('solution_file'):
        outt = build(cfg, 'teacher', cfg.get('output_teacher', 'lecture_教师版.docx'), base)
        results.append(outt)
    else:
        print('注意：未提供 solution_file，仅生成学生版（教师版需要含答案/详解的题目文件）。')
    if args.open and sys.platform.startswith('win'):
        for r in results:
            try: os.startfile(r)
            except Exception as e: print('open failed:', e)
    print('DONE:', ' , '.join(results))

if __name__ == '__main__':
    main()
