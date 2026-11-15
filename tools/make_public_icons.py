"""Public edition icons: one set per colour preset, from the Focus mark (no dark square).
  <hex>/any-192|512, favicon-64 : the mark in that colour on transparent      -> computer, browser tab
  <hex>/apple-180, maskable-*   : white mark on a tile of that colour            -> iPhone/iPad, Android
  mono-192|512                  : white mark on transparent (Android tints it)   -> themed icons
  mark.png                      : white mark, tight, used as a CSS mask so the header mark takes any colour
Also writes manifest-<hex>.webmanifest for each preset (same app id) — the app links the one for the chosen colour.
usage: python make_public_icons.py <source.webp> <out_dir> <base manifest.webmanifest>"""
import json, os, sys
import numpy as np
from PIL import Image, ImageFilter

ACCENTS = ['#EE2E24', '#2F6FDE', '#1D9A55', '#7B4FD6', '#F07A1A', '#0E9F9A', '#D9417A', '#4A4A4F']
src, out, base_manifest = sys.argv[1], sys.argv[2], sys.argv[3]
a = np.asarray(Image.open(src).convert('RGBA')).astype(int)
solid = (a[..., 0] > 200) & (a[..., 1] < 90) & (a[..., 3] > 200)
m = Image.fromarray((solid * 255).astype(np.uint8), 'L').filter(ImageFilter.MedianFilter(5))
ys, xs = np.where(np.asarray(m) > 127)
cx, cy = (xs.min() + xs.max()) / 2, (ys.min() + ys.max()) / 2
rmax = np.sqrt((xs - cx) ** 2 + (ys - cy) ** 2).max()

def placed(frac, size=2048):
    k = (frac * size) / rmax
    mm = m.resize((int(m.size[0] * k), int(m.size[1] * k)), Image.LANCZOS)
    c = Image.new('L', (size, size), 0); c.paste(mm, (int(size / 2 - cx * k), int(size / 2 - cy * k))); return c
def colour(mask, hexc, bg=None, s=512):
    rgb = tuple(int(hexc[i:i + 2], 16) for i in (1, 3, 5))
    al = mask.resize((s, s), Image.LANCZOS)
    if bg is None:
        im = Image.new('RGBA', (s, s), rgb + (0,)); im.putalpha(al); return im
    tile = Image.new('RGBA', (s, s), tuple(int(bg[i:i + 2], 16) for i in (1, 3, 5)) + (255,))
    tile.alpha_composite(Image.merge('RGBA', [Image.new('L', (s, s), v) for v in rgb] + [al])); return tile.convert('RGB')

big, safe, mono = placed(0.46), placed(0.33), placed(0.30)
os.makedirs(out, exist_ok=True)
base = json.load(open(base_manifest, encoding='utf-8'))
for hexc in ACCENTS:
    d = os.path.join(out, hexc[1:].lower()); os.makedirs(d, exist_ok=True)
    for s in (512, 192): colour(big, hexc, s=s).save(f'{d}/any-{s}.png', optimize=True)
    colour(big, hexc, s=64).save(f'{d}/favicon-64.png', optimize=True)
    for s in (512, 192): colour(safe, '#FFFFFF', bg=hexc, s=s).save(f'{d}/maskable-{s}.png', optimize=True)
    colour(safe, '#FFFFFF', bg=hexc, s=180).save(f'{d}/apple-180.png', optimize=True)
    man = dict(base); h = hexc[1:].lower()
    man['theme_color'] = hexc
    man['icons'] = [
        {'src': f'icons/{h}/any-192.png', 'sizes': '192x192', 'type': 'image/png', 'purpose': 'any'},
        {'src': f'icons/{h}/any-512.png', 'sizes': '512x512', 'type': 'image/png', 'purpose': 'any'},
        {'src': f'icons/{h}/maskable-192.png', 'sizes': '192x192', 'type': 'image/png', 'purpose': 'maskable'},
        {'src': f'icons/{h}/maskable-512.png', 'sizes': '512x512', 'type': 'image/png', 'purpose': 'maskable'},
        {'src': 'icons/mono-192.png', 'sizes': '192x192', 'type': 'image/png', 'purpose': 'monochrome'},
        {'src': 'icons/mono-512.png', 'sizes': '512x512', 'type': 'image/png', 'purpose': 'monochrome'},
    ]
    man['shortcuts'] = [dict(sc, icons=[{'src': f'icons/{h}/any-192.png', 'sizes': '192x192'}]) for sc in base.get('shortcuts', [])]
    json.dump(man, open(os.path.join(os.path.dirname(out), f'manifest-{h}.webmanifest'), 'w', encoding='utf-8', newline='\n'), ensure_ascii=False, indent=2)
for s in (512, 192):
    w = Image.new('RGBA', (s, s), (255, 255, 255, 0)); w.putalpha(mono.resize((s, s), Image.LANCZOS)); w.save(f'{out}/mono-{s}.png', optimize=True)
tight = m.crop(m.getbbox()); side = max(tight.size)
sq = Image.new('L', (side, side), 0); sq.paste(tight, ((side - tight.size[0]) // 2, (side - tight.size[1]) // 2))
mk = Image.new('RGBA', (256, 256), (255, 255, 255, 0)); mk.putalpha(sq.resize((256, 256), Image.LANCZOS)); mk.save(f'{out}/mark.png', optimize=True)
print('icons for', len(ACCENTS), 'colours')
