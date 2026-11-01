"""Fills template.html with a lesson's JSON and lists the audio clips the page needs.

Usage: python -I build_html.py <lesson.json> <out.html> <localStorage key>

Every 🔊 button plays audio/<id>.mp3. The page looks the id up by the exact text its script
derives (see speakBtn calls in template.html); speech_texts() below derives the same texts.
The clips themselves are made by make_audio.py from _generator/audio-needed.json.
"""
import hashlib
import html
import json
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).parent
QUESTION_VOICE = 'en-GB-RyanNeural'  # the examiner
ANSWER_VOICE = 'en-GB-SoniaNeural'


def text_of(h):
    return html.unescape(re.sub(r'<[^>]+>', '', h))


def sample_speech(h):  # same as sampleSpeech() in the template
    t = text_of(h)
    i = t.find(' → ')
    if i >= 0:
        t = t[i + 3:]
    t = re.sub(r'\s*\([^)]*\)\s*$', '', t)
    return re.sub(r'\s*/[^/\s]+/\s*', ' ', t)


def spoken(t):
    """What the voice reads: stress marks such as aPARTment or PARents read as plain words."""
    t = re.sub(r'\b(?=\w*[a-z])(?=\w*[A-Z]{2})\w+\b', lambda m: m.group(0).lower(), t)
    t = re.sub(r'\b([A-Z]{2,})([a-z]+)\b', lambda m: m.group(0).capitalize(), t)
    return re.sub(r'\s+', ' ', t).strip()


def speech_texts(data):
    """[(text the page looks up, voice)] for every 🔊 button the page can show."""
    out = []
    for u in data['units']:
        for b in u['theory']:
            if b['type'] == 'table':
                vocab = b['head'][0] == 'Cụm từ'
                if vocab or re.search(r'câu trả lời mở rộng', b['head'][-1], re.I):
                    out += [(text_of(r['cells'][-1]), ANSWER_VOICE) for r in b['rows']]
            elif b['type'] == 'samples':
                out += [(sample_speech(x), ANSWER_VOICE) for x in b['items']]
        for ex in u['exercises']:
            for it in ex['items']:
                if 'q' in it:
                    out.append((it['q'], QUESTION_VOICE))
                if it['kind'] == 'read':
                    out.append((text_of(it['cue']), ANSWER_VOICE))
                if 'answer' in it:
                    out.append((text_of(it['answer']), ANSWER_VOICE))
                if it.get('alt'):
                    out.append((text_of(it['alt']), ANSWER_VOICE))
    return [(t, v) for t, v in out if t.strip()]


def clip_id(text, voice):
    return hashlib.sha1(f'{voice}|{spoken(text)}'.encode('utf-8')).hexdigest()[:12]


def main(src, dst, store):
    data = json.loads(pathlib.Path(src).read_text(encoding='utf-8'))
    edits_path = pathlib.Path(src).with_suffix('.edits.json')
    if edits_path.exists():  # the teacher's changes on top of the .docx
        edits = json.loads(edits_path.read_text(encoding='utf-8'))
        drop = edits.get('dropGuide', [])
        if not all(any(b['html'] == h for b in data['guide']) for h in drop):
            raise SystemExit('dropGuide: a block was not found')
        data['guide'] = [b for b in data['guide'] if b['html'] not in drop]
        raw = json.dumps(data, ensure_ascii=False)
        for old, new in edits.get('replace', []):
            o, n = json.dumps(old, ensure_ascii=False)[1:-1], json.dumps(new, ensure_ascii=False)[1:-1]
            if o not in raw:
                raise SystemExit(f'replace: not found: {old[:60]}')
            raw = raw.replace(o, n)
        data = json.loads(raw)
    for k in ('bank', 'bankIntro', 'bankTitle'):  # the question bank is not shown (teacher's choice)
        data.pop(k, None)
    audio, needed = {}, {}
    for text, voice in speech_texts(data):
        cid = clip_id(text, voice)
        if audio.get(text, cid) != cid:
            continue  # same text as a question and an answer: keep the first voice
        audio[text] = cid
        needed[cid] = {'voice': voice, 'text': spoken(text)}
    needed_path = HERE / 'audio-needed.json'
    old = json.loads(needed_path.read_text(encoding='utf-8')) if needed_path.exists() else {}
    old.update(needed)
    needed_path.write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding='utf-8')

    def js(v):
        return json.dumps(v, ensure_ascii=False).replace('</', '<\\/')

    page = ((HERE / 'template.html').read_text(encoding='utf-8')
            .replace('__TITLE__', data['cover'][0])
            .replace('__STORE__', store)
            .replace('__AUDIO_DIR__', 'audio/')
            .replace('__AUDIO__', js(audio))
            .replace('__DATA__', js(data)))
    pathlib.Path(dst).write_text(page, encoding='utf-8', newline='\n')
    print('wrote', dst, len(page), 'bytes;', len(audio), 'clips')


if __name__ == '__main__':
    main(*sys.argv[1:4])
