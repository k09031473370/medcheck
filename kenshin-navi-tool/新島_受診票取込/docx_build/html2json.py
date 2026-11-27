# -*- coding: utf-8 -*-
# 受診票取込_説明書.html を、Word 作成用の JSON (ブロックの並び) に直す。
# 使い方: python3 html2json.py 説明書.html out.json
import sys, json, re
import lxml.html

def norm(t):
    return re.sub(r'\s+', ' ', t or '')

def inl(el, st=None, out=None, top=True):
    """要素の中身を run の並びにする。入れ子の ul / ol は含めない。"""
    st = dict(st or {}); out = [] if out is None else out
    if el.text: out.append(dict(st, t=norm(el.text)))
    for c in el:
        tag = c.tag
        if tag in ('ul', 'ol'):
            pass
        elif tag == 'br':
            out.append({'br': True})
        else:
            s2 = dict(st)
            cls = c.get('class') or ''
            if tag == 'b': s2['b'] = True
            elif tag == 'code': s2['code'] = True
            elif tag == 'span' and 'btn' in cls: s2['btn'] = True
            elif tag == 'span' and 'small' in cls: s2['small'] = True
            inl(c, s2, out, False)
        if c.tail: out.append(dict(st, t=norm(c.tail)))
    if top:
        # 前後の空白を落とし、空の run を消す
        out = [r for r in out if r.get('br') or r.get('t')]
        while out and not out[0].get('br') and not out[0]['t'].strip(): out.pop(0)
        while out and not out[-1].get('br') and not out[-1]['t'].strip(): out.pop()
        if out and 't' in out[0]: out[0]['t'] = out[0]['t'].lstrip()
        if out and 't' in out[-1]: out[-1]['t'] = out[-1]['t'].rstrip()
        for i, r in enumerate(out):
            if r.get('br'):
                if i > 0 and 't' in out[i-1]: out[i-1]['t'] = out[i-1]['t'].rstrip()
                if i+1 < len(out) and 't' in out[i+1]: out[i+1]['t'] = out[i+1]['t'].lstrip()
        out = [r for r in out if r.get('br') or r.get('t')]
    return out

def lst(el):
    items = []
    for li in el.findall('li'):
        sub = []
        for c in li:
            if c.tag in ('ul', 'ol'):
                sub = [inl(x) for x in c.findall('li')]
        items.append({'runs': inl(li), 'sub': sub})
    return {'type': el.tag, 'items': items}

def main(src, dst):
    doc = lxml.html.parse(src)
    main_el = doc.getroot().find('.//main')
    blocks = []
    for el in main_el:
        tag = el.tag; cls = el.get('class') or ''
        if tag in ('h1', 'h2', 'h3'):
            blocks.append({'type': tag, 'runs': inl(el)})
        elif tag == 'p':
            kind = 'lead' if 'lead' in cls else ('small' if 'small' in cls else 'p')
            blocks.append({'type': kind, 'runs': inl(el)})
        elif tag == 'ul' and 'flow' in cls:
            blocks.append({'type': 'flow', 'steps': [norm(li.text_content()).strip() for li in el.findall('li') if 'arrow' not in (li.get('class') or '')]})
        elif tag in ('ul', 'ol'):
            blocks.append(lst(el))
        elif tag == 'div' and 'toc' in cls:
            blocks.append({'type': 'toc', 'items': [norm(a.text_content()).strip() for a in el.findall('a')]})
        elif tag == 'div' and 'tablewrap' in cls:
            rows = []
            for tr in el.iter('tr'):
                rows.append({'head': tr.find('th') is not None, 'cells': [inl(c) for c in tr if c.tag in ('th', 'td')]})
            blocks.append({'type': 'table', 'rows': rows})
        elif tag == 'div' and ('note' in cls or 'warn' in cls):
            kind = 'warn' if 'warn' in cls else 'note'
            ul = el.find('ul')
            if ul is not None:
                blocks.append({'type': 'callout', 'kind': kind, 'items': [inl(li) for li in ul.findall('li')]})
            else:
                blocks.append({'type': 'callout', 'kind': kind, 'runs': inl(el)})
        else:
            raise SystemExit('知らない要素: %s class=%s' % (tag, cls))
    json.dump(blocks, open(dst, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('blocks', len(blocks))

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
