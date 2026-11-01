"""Fills template.html with a lesson's JSON.

Usage: python -I build_html.py <lesson.json> <out.html> <localStorage key>
"""
import json
import pathlib
import sys

src, dst, store = sys.argv[1], sys.argv[2], sys.argv[3]
data = json.loads(pathlib.Path(src).read_text(encoding='utf-8'))
template = (pathlib.Path(__file__).parent / 'template.html').read_text(encoding='utf-8')
payload = json.dumps(data, ensure_ascii=False).replace('</', '<\\/')
page = (template.replace('__TITLE__', data['cover'][0])
        .replace('__STORE__', store)
        .replace('__DATA__', payload))
pathlib.Path(dst).write_text(page, encoding='utf-8', newline='\n')
print('wrote', dst, len(page), 'bytes')
