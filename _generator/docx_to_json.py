"""Reads an "IELTS Speaking N" handout (.docx) and writes its content as JSON.

Usage: python -I docx_to_json.py <handout.docx> <out.json>

The handout layout this expects (IELTS_Speaking_1.docx):
  title lines, "Hướng dẫn sử dụng" (Heading 1), "Phần B ..." (Heading 1),
  per unit: "Unit N  Name" (Heading 2), "A. Lý thuyết", vocabulary / expansion tables,
  "Câu mẫu", three 1x1 boxes (Lỗi cần tránh, Chọn cho đúng, Âm cần luyện),
  "B. Bài tập", Bài tập 1-3 (Heading 3);
  "Phụ lục: Ngân hàng câu hỏi Part 1" (Heading 1) with one table per topic;
  "Đáp án" (Heading 1) with the same units and exercises in the same order.
Bold and italic runs are kept as <b> and <i>.
"""
import html
import json
import re
import sys

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph


def runs_of(p):
    """[(text, bold, italic)] with neighbouring runs of the same format merged."""
    out = []
    for r in p.runs:
        if not r.text:
            continue
        k = (bool(r.bold), bool(r.italic))
        if out and (out[-1][1], out[-1][2]) == k:
            out[-1] = (out[-1][0] + r.text, *k)
        else:
            out.append((r.text, *k))
    return out


def strip_prefix(runs, prefix_re):
    """Removes a leading pattern (matched on the plain text) from a runs list."""
    text = ''.join(t for t, _, _ in runs)
    m = re.match(prefix_re, text)
    if not m:
        return runs
    n = m.end()
    out = []
    for t, b, i in runs:
        if n >= len(t):
            n -= len(t)
            continue
        out.append((t[n:], b, i))
        n = 0
    return out


def strip_suffix(runs, suffix_re):
    text = ''.join(t for t, _, _ in runs)
    m = re.search(suffix_re, text)
    if not m:
        return runs
    keep = m.start()
    out = []
    for t, b, i in runs:
        if keep <= 0:
            break
        out.append((t[:keep], b, i))
        keep -= len(t)
    return out


def to_html(runs):
    s = ''
    for t, b, i in runs:
        e = html.escape(t, quote=False)
        if i:
            e = f'<i>{e}</i>'
        if b:
            e = f'<b>{e}</b>'
        s += e
    return re.sub(r'\s+', ' ', s).strip()


def plain(runs):
    return re.sub(r'\s+', ' ', ''.join(t for t, _, _ in runs)).strip()


def level_of(text):
    m = re.match(r'\s*(●+)', text)
    return len(m.group(1)) if m else 0


DOT = r'\s*●+\s*'
SAY = r'\s*Nói to\s*☐\s*$'


def cell_paras(cell):
    return [runs_of(p) for p in cell.paragraphs if p.text.strip()]


def main(src, dst):
    d = docx.Document(src)
    blocks = []
    for el in d.element.body.iterchildren():
        tag = el.tag.split('}')[1]
        if tag == 'p':
            p = Paragraph(el, d)
            if not p.text.strip():
                continue
            style = p.style.name if p.style is not None else ''
            blocks.append(('p', style, p))
        elif tag == 'tbl':
            blocks.append(('t', '', Table(el, d)))

    data = {'cover': [], 'guide': [], 'units': [], 'bankIntro': [], 'bank': []}
    section = 'cover'  # cover, guide, units, bank, key
    unit = None
    part = None  # theory heading or exercise key
    key_unit = None
    key_ex = None
    key_idx = 0
    pending = None  # current exercise item (questions) or answer item (key)

    def unit_by_name(text):
        m = re.match(r'Unit\s+(\d+)', text)
        return next(u for u in data['units'] if u['num'] == int(m.group(1)))

    for kind, style, obj in blocks:
        if kind == 'p':
            text = obj.text.strip()
            runs = runs_of(obj)
        if kind == 'p' and style == 'Heading 1':
            if text.startswith('Hướng dẫn'):
                section = 'guide'
                data['guideTitle'] = text
            elif text.startswith('Phần B'):
                section = 'units'
                data['unitsTitle'] = re.sub(r'\s+', ' ', text)
            elif text.startswith('Phụ lục'):
                section = 'bank'
                data['bankTitle'] = text
            elif text.startswith('Đáp án'):
                section = 'key'
            continue

        if section == 'cover':
            if kind == 'p' and text != 'Mục lục':
                data['cover'].append(text)
            continue

        if section == 'guide':
            if kind == 'p':
                if style == 'List Paragraph':
                    data['guide'].append({'type': 'li', 'html': to_html(runs)})
                elif all(b for _, b, _ in runs) or len(text) < 60 and not text.endswith('.'):
                    data['guide'].append({'type': 'h', 'html': to_html(runs)})
                else:
                    data['guide'].append({'type': 'p', 'html': to_html(runs)})
            continue

        if section == 'units':
            if kind == 'p' and style == 'Heading 2':
                m = re.match(r'Unit\s+(\d+)\s+(.*)', text)
                unit = {'num': int(m.group(1)), 'name': m.group(2).strip(), 'purpose': [],
                        'vocab': [], 'expandFormula': '', 'expand': [], 'samples': [],
                        'mistakes': [], 'choose': [], 'sounds': {'lines': [], 'tip': ''},
                        'ex': {'1': {'intro': '', 'items': []}, '2': {'intro': '', 'items': []},
                               '3': {'intro': '', 'items': []}}}
                data['units'].append(unit)
                part = None
                continue
            if kind == 'p' and style == 'Heading 3':
                m = re.match(r'Bài tập (\d)', text)
                if m:
                    part = 'ex' + m.group(1)
                    unit['ex'][m.group(1)]['title'] = text
                else:
                    part = 'theory' if text.startswith('A.') else 'exhead'
                continue
            if kind == 't':
                rows = obj.rows
                if len(rows) == 1 and len(obj.columns) == 1:
                    paras = cell_paras(obj.cell(0, 0))
                    head = plain(paras[0])
                    body = paras[1:]
                    if head.startswith('Lỗi'):
                        for j in range(0, len(body), 3):
                            unit['mistakes'].append({
                                'bad': to_html(strip_prefix(body[j], r'\s*✗\s*')),
                                'good': to_html(strip_prefix(body[j + 1], r'\s*✓\s*')),
                                'note': to_html(body[j + 2])})
                    elif head.startswith('Chọn'):
                        for j in range(0, len(body), 2):
                            unit['choose'].append({'h': plain(body[j]), 'html': to_html(body[j + 1])})
                    elif head.startswith('Âm'):
                        unit['sounds']['lines'] = [to_html(r) for r in body[:-1]]
                        unit['sounds']['tip'] = to_html(body[-1])
                    else:
                        raise SystemExit(f'Unknown box: {head}')
                elif plain(runs_of(obj.cell(0, 0).paragraphs[0])) == 'Cụm từ':
                    for r in rows[1:]:
                        c = [cell_paras(x)[0] for x in r.cells]
                        unit['vocab'].append({'level': level_of(plain(c[0])),
                                              'phrase': plain(strip_prefix(c[0], DOT)),
                                              'meaning': plain(c[1]), 'example': to_html(c[2])})
                else:
                    for r in rows[1:]:
                        c = [cell_paras(x)[0] for x in r.cells]
                        q, short = plain(c[0]).split(' → ', 1)
                        unit['expand'].append({'q': q, 'short': short, 'tech': plain(c[1]),
                                               'answer': to_html(c[2])})
                continue
            # paragraphs inside a unit
            if part == 'theory':
                if text in ('Mục đích', 'Từ vựng hay cho chủ đề', 'Từ câu ngắn đến câu trả lời đầy đủ',
                            'Câu mẫu'):
                    unit['_sub'] = text
                elif unit.get('_sub') == 'Mục đích':
                    unit['purpose'].append(to_html(runs))
                elif unit.get('_sub') == 'Từ câu ngắn đến câu trả lời đầy đủ':
                    unit['expandFormula'] = to_html(runs)
                elif unit.get('_sub') == 'Câu mẫu':
                    q, a = text.split(' → ', 1)
                    unit['samples'].append({'q': q, 'a': a})
                else:
                    raise SystemExit(f'Unplaced theory text: {text}')
            elif part and part.startswith('ex'):
                ex = unit['ex'][part[2]]
                if style == 'List Paragraph':
                    r2 = strip_suffix(strip_prefix(runs, DOT), SAY)
                    item = {'level': level_of(text)}
                    if part == 'ex1':
                        item['cue'] = to_html(r2)
                    else:
                        item['q'] = plain(strip_prefix(r2, r'\s*Examiner:\s*'))
                    ex['items'].append(item)
                elif not ex['items']:
                    ex['intro'] = to_html(runs)
                else:
                    item = ex['items'][-1]
                    for label, k in (('Gợi ý:', 'hint'), ('Câu ngắn:', 'short'),
                                     ('Mở rộng:', 'expand'), ('Dùng:', 'use')):
                        if text.startswith(label):
                            item[k] = to_html(strip_prefix(runs, r'\s*' + label + r'\s*'))
                            break
                    else:
                        raise SystemExit(f'Unknown exercise line: {text}')
            else:
                raise SystemExit(f'Unplaced unit text: {text}')
            continue

        if section == 'bank':
            if kind == 'p' and style == 'Heading 3':
                data['bank'].append({'topic': text, 'rows': []})
            elif kind == 'p':
                data['bankIntro'].append(to_html(runs))
            else:
                for r in obj.rows[1:]:
                    data['bank'][-1]['rows'].append([to_html(cell_paras(x)[0]) if cell_paras(x) else ''
                                                     for x in r.cells])
            continue

        if section == 'key':
            if kind == 'p' and re.match(r'Unit\s+\d+', text) and style != 'List Paragraph':
                key_unit = unit_by_name(text)
                continue
            if kind == 'p' and style == 'Heading 3':
                key_ex = key_unit['ex'][re.match(r'Bài tập (\d)', text).group(1)]
                key_idx = 0
                continue
            if style == 'List Paragraph':
                pending = key_ex['items'][key_idx]
                key_idx += 1
                lv = level_of(text)
                if lv != pending['level']:
                    raise SystemExit(f'Level mismatch in key: {text}')
                r2 = strip_prefix(strip_prefix(runs, DOT), r'\s*Câu mẫu:\s*')
                pending['answer'] = to_html(r2)
            elif text.startswith('Cũng chấp nhận:'):
                pending['alt'] = to_html(strip_prefix(runs, r'\s*Cũng chấp nhận:\s*'))
            else:
                pending['note'] = to_html(runs)

    for u in data['units']:
        u.pop('_sub', None)
        for n, ex in u['ex'].items():
            for it in ex['items']:
                if 'answer' not in it:
                    raise SystemExit(f'No answer: Unit {u["num"]} Bài tập {n}: {it}')
    with open(dst, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    counts = {u['num']: [len(u['ex'][k]['items']) for k in '123'] for u in data['units']}
    print('units', counts, 'bank', [(b['topic'], len(b['rows'])) for b in data['bank']])


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
