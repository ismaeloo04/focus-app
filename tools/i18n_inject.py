"""Writes tools/i18n_en.json into index.html (the /*i18n-skip*/const EN = …;/*/i18n-skip*/ block)
and reports any wrapped phrase that has no English yet.  Usage: python tools/i18n_inject.py"""
import json, os, re, subprocess, sys
sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
p = os.path.join(HERE, 'index.html')
d = json.load(open(os.path.join(HERE, 'tools', 'i18n_en.json'), encoding='utf-8'))
s = open(p, encoding='utf-8').read()
new, n = re.subn(r'/\*i18n-skip\*/const EN = .*?;/\*/i18n-skip\*/', lambda m: '/*i18n-skip*/const EN = ' + json.dumps(d, ensure_ascii=False) + ';/*/i18n-skip*/', s, flags=re.S)
assert n == 1, n
open(p, 'w', encoding='utf-8', newline='\n').write(new)
subprocess.run([sys.executable, os.path.join(HERE, 'tools', 'i18n_wrap.py'), 'scan'], check=True)
keys = json.load(open(os.path.join(HERE, 'tools', 'i18n_keys.json'), encoding='utf-8'))
missing = [k for k in keys if k not in d]
print('dictionary', len(d), 'missing', missing)
