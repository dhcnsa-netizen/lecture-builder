# -*- coding: utf-8 -*-
"""
extract_outline.py — Dump a per-paragraph index of a .docx so you can design the
lecture config: knowledge ranges, question segments, and problem classification.

    python extract_outline.py "知识点.docx"  [> outline.txt]

Each line is:  <pt_index>  <TAG>  | <text first 100 chars> [flags]
flags:  〖IMG〗 has figure   〖TBL〗 table   〖MATH〗 has formula   ★START problem-start
The pt_index is exactly what you put in config knowledge ranges / segment start-end.

IMPORTANT: this uses the SAME parser as lecture_engine (imported below), so the pt
indices here are guaranteed identical to what the build engine sees.
"""
import zipfile, re, sys, os, io

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import lecture_engine as E  # single source of truth for body_children / childmap / is_start

WT = re.compile(r'<w:t[^>]*>(.*?)</w:t>', re.S)
MT = re.compile(r'<m:t[^>]*>(.*?)</m:t>', re.S)

def text_of(s):
    t = ''.join(WT.findall(s)); t = re.sub(r'<[^>]+>', '', t)
    m = ''.join(MT.findall(s)); m = re.sub(r'<[^>]+>', '', m)
    return t, m

def main():
    path = sys.argv[1]
    out = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    z = zipfile.ZipFile(path)
    doc = z.read('word/document.xml').decode('utf-8')
    pt = E.childmap(doc)
    out.write('FILE: %s   pt_count=%d\n' % (path, len(pt)))
    for i, (tag, s) in enumerate(pt):
        if tag == 'w:tbl':
            out.write('%03d  TBL | 〖TBL〗\n' % i); continue
        t, m = text_of(s)
        flags = ''
        if '<a:blip' in s: flags += '〖IMG〗'
        if m: flags += '〖MATH:%s〗' % m[:30]
        start = ' ★START' if E.is_start(s) else ''
        out.write('%03d  p   | %s %s%s\n' % (i, t[:100], flags, start))
    out.flush()

if __name__ == '__main__':
    main()
