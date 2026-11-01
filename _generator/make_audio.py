"""Makes the missing clips listed in audio-needed.json as audio/<id>.mp3 (Microsoft neural voices).

Usage (needs the edge-tts package): python make_audio.py
"""
import asyncio
import json
import pathlib

import edge_tts

HERE = pathlib.Path(__file__).parent
OUT = HERE.parent / 'audio'


async def make(cid, clip, sem):
    path = OUT / f'{cid}.mp3'
    async with sem:
        for attempt in range(4):
            try:
                await edge_tts.Communicate(clip['text'], clip['voice'], rate='-5%').save(str(path))
                return
            except Exception as e:  # network hiccups: try again
                if attempt == 3:
                    raise SystemExit(f'{cid}: {e}')
                await asyncio.sleep(2 + attempt * 3)


async def main():
    OUT.mkdir(exist_ok=True)
    needed = json.loads((HERE / 'audio-needed.json').read_text(encoding='utf-8'))
    todo = {k: v for k, v in needed.items() if not (OUT / f'{k}.mp3').exists()}
    sem = asyncio.Semaphore(6)
    await asyncio.gather(*(make(k, v, sem) for k, v in todo.items()))
    print('made', len(todo), 'clips;', len(needed), 'listed')


asyncio.run(main())
