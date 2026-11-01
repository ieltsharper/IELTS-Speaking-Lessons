"""Reads an IELTS Speaking workbook (.docx) and writes its content as JSON.

Usage: python -I docx_to_json.py <workbook.docx> <out.json>

Layout this expects (IELTS_Intensive_Speaking_Lesson_1_Workbook v2.docx):
  cover lines, "Hướng dẫn sử dụng" (Heading 1), "Phần A/B ..." (Heading 1),
  per unit: "Unit N  Name" (Heading 2), "A. Lý thuyết" with sub-headings, paragraphs, tables,
  "Câu mẫu" and 1x1 boxes (Lỗi cần tránh, Chọn cho đúng, Âm cần luyện); "B. Bài tập" with
  "Bài tập N: ..." (Heading 3). Exercise items are List Paragraphs starting with ●, followed by
  either labelled lines (Gợi ý:, Câu ngắn:, ...) or A/B/C options;
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
    s = re.sub(r'</b><b>|</i><i>', '', s)
    return re.sub(r'\s+', ' ', s).strip()


def plain(runs):
    return re.sub(r'\s+', ' ', ''.join(t for t, _, _ in runs)).strip()


def level_of(text):
    m = re.match(r'\s*(●+)', text)
    return len(m.group(1)) if m else 0


DOT = r'\s*●+\s*'
SAY = r'\s*Nói to\s*☐\s*$'
LABELS = (('Gợi ý:', 'hint'), ('Câu ngắn:', 'short'), ('Mở rộng:', 'expand'), ('Dùng:', 'use'),
          ('Câu nghe như học thuộc:', 'rote'), ('Đặc điểm cá nhân:', 'trait'))
OPTION = re.compile(r'([ABC])\s{2,}(.*?)(?=\s{3,}[ABC]\s{2,}|$)')


def cell_paras(cell):
    return [runs_of(p) for p in cell.paragraphs if p.text.strip()]


def cell_html(cell):
    return '<br>'.join(to_html(r) for r in cell_paras(cell))


def main(src, dst):
    d = docx.Document(src)
    blocks = []
    for el in d.element.body.iterchildren():
        tag = el.tag.split('}')[1]
        if tag == 'p':
            p = Paragraph(el, d)
            if p.text.strip():
                blocks.append(('p', p.style.name if p.style is not None else '', p))
        elif tag == 'tbl':
            blocks.append(('t', '', Table(el, d)))

    data = {'cover': [], 'guide': [], 'parts': [], 'units': [], 'bankIntro': [], 'bank': []}
    section = 'cover'
    unit = ex = None
    part = None
    key_unit = key_ex = pending = None
    key_idx = 0

    def unit_by_num(text):
        n = int(re.match(r'Unit\s+(\d+)', text).group(1))
        return next(u for u in data['units'] if u['num'] == n)

    for kind, style, obj in blocks:
        text = obj.text.strip() if kind == 'p' else ''
        runs = runs_of(obj) if kind == 'p' else []

        if kind == 'p' and style == 'Heading 1':
            if text.startswith('Hướng dẫn'):
                section, data['guideTitle'] = 'guide', text
            elif text.startswith('Phần'):
                section = 'units'
                data['parts'].append({'title': re.sub(r'\s+', ' ', text), 'units': []})
            elif text.startswith('Phụ lục'):
                section, data['bankTitle'] = 'bank', text
            elif text.startswith('Đáp án'):
                section = 'key'
            continue

        if section == 'cover':
            if text != 'Mục lục':
                data['cover'].append(text)
            continue

        if section == 'guide':
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
                unit = {'num': int(m.group(1)), 'name': m.group(2).strip(), 'theory': [], 'exercises': []}
                data['units'].append(unit)
                data['parts'][-1]['units'].append(unit['num'])
                part = None
                continue
            if kind == 'p' and style == 'Heading 3':
                if re.match(r'Bài tập \d', text):
                    part = 'ex'
                    ex = {'title': text, 'intro': '', 'items': []}
                    unit['exercises'].append(ex)
                else:
                    part = 'theory' if text.startswith('A.') else 'exhead'
                continue
            if part == 'theory':
                th = unit['theory']
                if kind == 't':
                    if len(obj.rows) == 1 and len(obj.columns) == 1:
                        paras = cell_paras(obj.cell(0, 0))
                        head, body = plain(paras[0]), paras[1:]
                        if head.startswith('Lỗi'):
                            th.append({'type': 'mistakes', 'items': [
                                {'bad': to_html(strip_prefix(body[j], r'\s*✗\s*')),
                                 'good': to_html(strip_prefix(body[j + 1], r'\s*✓\s*')),
                                 'note': to_html(body[j + 2])} for j in range(0, len(body), 3)]})
                        elif head.startswith('Chọn'):
                            th.append({'type': 'choose', 'items': [
                                {'h': to_html(body[j]), 'html': to_html(body[j + 1])}
                                for j in range(0, len(body), 2)]})
                        elif head.startswith('Âm'):
                            th.append({'type': 'sounds', 'lines': [to_html(r) for r in body[:-1]],
                                       'tip': to_html(body[-1])})
                        else:
                            raise SystemExit(f'Unknown box: {head}')
                    else:
                        head = [plain(cell_paras(c)[0]) for c in obj.rows[0].cells]
                        rows = []
                        for r in obj.rows[1:]:
                            cells = [cell_html(c) for c in r.cells]
                            first = plain(cell_paras(r.cells[0])[0]) if cell_paras(r.cells[0]) else ''
                            lv = level_of(first)
                            if lv:
                                cells[0] = re.sub(r'^●+\s*', '', cells[0])
                            rows.append({'level': lv, 'cells': cells})
                        th.append({'type': 'table', 'head': head, 'rows': rows})
                elif style == 'List Paragraph':
                    if not th or th[-1]['type'] != 'samples':
                        th.append({'type': 'samples', 'items': []})
                    th[-1]['items'].append(to_html(runs))
                elif all(b for _, b, _ in runs) or (len(text) < 70 and not text.endswith('.')):
                    th.append({'type': 'sub', 'text': text})
                else:
                    th.append({'type': 'p', 'html': to_html(runs)})
                continue
            if part == 'ex':
                if style == 'List Paragraph':
                    r2 = strip_suffix(strip_prefix(runs, DOT), SAY)
                    m = re.match(r'\s*Examiner(?:\s*\((Q\d+)\))?:\s*', plain(r2))
                    item = {'level': level_of(text)}
                    if m:
                        item['q'] = plain(r2)[m.end():]
                        if m.group(1):
                            item['qnum'] = m.group(1)
                    else:
                        item['cue'] = to_html(r2)
                    ex['items'].append(item)
                elif not ex['items']:
                    ex['intro'] = to_html(runs)
                else:
                    item = ex['items'][-1]
                    if re.match(r'[ABC]\s{2,}', text):
                        item.setdefault('options', []).extend(
                            {'letter': a, 'text': re.sub(r'\s+', ' ', t).strip()}
                            for a, t in OPTION.findall(text))
                        continue
                    for label, k in LABELS:
                        if text.startswith(label):
                            item[k] = to_html(strip_prefix(runs, r'\s*' + re.escape(label) + r'\s*'))
                            break
                    else:
                        raise SystemExit(f'Unknown exercise line: {text}')
                continue
            raise SystemExit(f'Unplaced unit text: {text}')

        if section == 'bank':
            if kind == 'p' and style == 'Heading 3':
                data['bank'].append({'topic': text, 'head': [], 'rows': []})
            elif kind == 'p':
                data['bankIntro'].append(to_html(runs))
            else:
                data['bank'][-1]['head'] = [plain(cell_paras(c)[0]) for c in obj.rows[0].cells]
                for r in obj.rows[1:]:
                    data['bank'][-1]['rows'].append([cell_html(c) for c in r.cells])
            continue

        if section == 'key':
            if style != 'List Paragraph' and style != 'Heading 3' and re.match(r'Unit\s+\d+', text):
                key_unit = unit_by_num(text)
                continue
            if style == 'Heading 3':
                n = int(re.match(r'Bài tập (\d)', text).group(1))
                key_ex = key_unit['exercises'][n - 1]
                key_idx = 0
                continue
            if style == 'List Paragraph':
                pending = key_ex['items'][key_idx]
                key_idx += 1
                if level_of(text) != pending['level']:
                    raise SystemExit(f'Level mismatch in key: {text}')
                r2 = strip_prefix(runs, DOT)
                m = re.match(r'\s*([ABC])\s{2,}', ''.join(t for t, _, _ in r2))
                if 'options' in pending:
                    if not m:
                        raise SystemExit(f'Option key expected: {text}')
                    pending['correct'] = m.group(1)
                    pending['note'] = to_html(strip_prefix(r2, r'\s*[ABC]\s{2,}'))
                else:
                    pending['answer'] = to_html(strip_prefix(r2, r'\s*Câu mẫu:\s*'))
                    pending['model'] = plain(r2).startswith('Câu mẫu:')
            elif text.startswith('Cũng chấp nhận:'):
                pending['alt'] = to_html(strip_prefix(runs, r'\s*Cũng chấp nhận:\s*'))
            else:
                pending['note'] = to_html(runs)

    for u in data['units']:
        for n, e in enumerate(u['exercises'], 1):
            for it in e['items']:
                if 'answer' not in it and 'correct' not in it:
                    raise SystemExit(f'No answer: Unit {u["num"]} Bài tập {n}: {it}')
                if 'options' in it and len(it['options']) != 3:
                    raise SystemExit(f'Options: Unit {u["num"]} Bài tập {n}: {it}')
                # Read-aloud items: a plain sentence, nothing to write.
                it['kind'] = ('choice' if 'options' in it
                              else 'read' if 'cue' in it and 'hint' not in it else 'write')
    with open(dst, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print('units', {u['num']: [(len(e['items']), sorted({i['kind'] for i in e['items']}))
                               for e in u['exercises']] for u in data['units']})
    print('bank', [(b['topic'], len(b['rows'])) for b in data['bank']])


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
