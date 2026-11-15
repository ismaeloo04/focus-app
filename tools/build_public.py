"""Builds the public edition of Focus (anyone can use it; nothing of Graficoo or of the owner inside)
from this folder's index.html into ../Focus Public.

    python tools/build_public.py [version]

The public edition shares the code with the personal one; only the edition config differs:
  - own localStorage key (both apps live on ismaeloo04.github.io, so keys must differ or they share data)
  - own manifest id (/focus/) and cache prefix (focuspub-), so it installs beside the personal app
  - no Graficoo link, first-run setup with job templates, colour themes on
"""
import json, os, re, shutil, sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(os.path.dirname(HERE), 'Focus Public')
VER = sys.argv[1] if len(sys.argv) > 1 else '1'
os.makedirs(os.path.join(OUT, 'icons'), exist_ok=True)

cfg = {'edition': 'public', 'storageKey': 'focuspub.v1', 'graficoo': False, 'themes': True, 'onboarding': True, 'brandSub': ''}
html = open(os.path.join(HERE, 'index.html'), encoding='utf-8').read()
def sub(pattern, repl, text, count=1):
    new, n = re.subn(pattern, repl, text)
    assert n == count, (pattern, n)
    return new
html = sub(r"<script>\n'use strict';\n", "<script>window.FOCUS_CFG = " + json.dumps(cfg) + ";</script>\n<script>\n'use strict';\n", html)
html = sub(r"const G_DEFAULT_URL = '[^']*';", "const G_DEFAULT_URL = '';", html)
html = sub(r"const G_APP = '[^']*';", "const G_APP = '';", html)
html = sub(r'<meta name="description" content="[^"]*">', '<meta name="description" content="مهامك ومشاريعك وعاداتك وملاحظاتك مع تايمر بومودورو للتركيز">', html)
html = sub(r'<small id="brandSub">GRAFICOO</small>', '<small id="brandSub" hidden></small>', html)
assert 'script.google.com/macros/s/AKfycbx' not in html
open(os.path.join(OUT, 'index.html'), 'w', encoding='utf-8', newline='\n').write(html)

man = json.load(open(os.path.join(HERE, 'manifest.webmanifest'), encoding='utf-8'))
man.update({'id': '/focus/', 'name': 'Focus', 'short_name': 'Focus', 'description': 'مهامك ومشاريعك وعاداتك وملاحظاتك مع تايمر بومودورو للتركيز'})
json.dump(man, open(os.path.join(OUT, 'manifest.webmanifest'), 'w', encoding='utf-8', newline='\n'), ensure_ascii=False, indent=2)

sw = open(os.path.join(HERE, 'sw.js'), encoding='utf-8').read()
sw = sub(r"const CACHE = '[^']*';", f"const CACHE = 'focuspub-v{VER}';", sw)
sw = sub(r"k\.startsWith\('focus-'\)", "k.startsWith('focuspub-')", sw)
open(os.path.join(OUT, 'sw.js'), 'w', encoding='utf-8', newline='\n').write(sw)

for f in ['any-192.png', 'any-512.png', 'maskable-192.png', 'maskable-512.png', 'mono-192.png', 'mono-512.png', 'apple-180.png', 'favicon-64.png']:
    shutil.copyfile(os.path.join(HERE, 'icons', f), os.path.join(OUT, 'icons', f))

os.makedirs(os.path.join(OUT, 'apps-script'), exist_ok=True)
gs = open(os.path.join(HERE, 'apps-script', 'Code.gs'), encoding='utf-8').read().replace('Graficoo Focus', 'Focus')
assert "const SYNC_KEY = 'CHANGE-ME'" in gs
open(os.path.join(OUT, 'apps-script', 'Code.gs'), 'w', encoding='utf-8', newline='\n').write(gs)
shutil.copyfile(os.path.join(HERE, 'apps-script', 'appsscript.json'), os.path.join(OUT, 'apps-script', 'appsscript.json'))
os.makedirs(os.path.join(OUT, 'widgets'), exist_ok=True)
wd = open(os.path.join(HERE, 'widgets', 'ios-scriptable-widget.js'), encoding='utf-8').read().replace('Graficoo Focus', 'Focus')
open(os.path.join(OUT, 'widgets', 'ios-scriptable-widget.js'), 'w', encoding='utf-8', newline='\n').write(wd)
open(os.path.join(OUT, '.nojekyll'), 'w').close()

for f in ['index.html', 'manifest.webmanifest', 'sw.js', 'apps-script/Code.gs', 'widgets/ios-scriptable-widget.js']:
    t = open(os.path.join(OUT, f), encoding='utf-8').read()
    assert 'AKfycbx' not in t and 'AKfycby' not in t and 'graficoo-app' not in t, f
print('built', OUT)
