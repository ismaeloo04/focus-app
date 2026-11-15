"""Finds the fixed Arabic UI text in index.html's main script and wraps it in L('…') so it can be shown in English.

    python tools/i18n_wrap.py scan   -> writes tools/i18n_keys.json (every Arabic phrase, in order)
    python tools/i18n_wrap.py apply  -> rewrites index.html with L() around each phrase

Only code is touched: comments and regex literals are skipped, and the user's own data is never in the source.
In plain strings the whole string becomes the key; in template literals each run of Arabic text between
markup/interpolations becomes ${L('…')}.
"""
import json, os, re, sys

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(HERE, 'index.html')
AR = re.compile(r'[؀-ۿ]')
# a run of display text: Arabic plus the punctuation / spaces / digits that sit inside sentences
RUN = re.compile(r'[؀-ۿ«»—…ـ،؟][؀-ۿ0-9«»—…ـ،؟!?.:()\-–/+ ‏%]*[؀-ۿ«»—…)؟!?.]|[؀-ۿ]')
SKIP = {'ى'}  # keyboard shortcut character, not text

def unescape(raw):
    out, i = [], 0
    while i < len(raw):
        c = raw[i]
        if c == '\\' and i + 1 < len(raw):
            n = raw[i + 1]; out.append({'n': '\n', 't': '\t'}.get(n, n)); i += 2
        else:
            out.append(c); i += 1
    return ''.join(out)

def lex(code, i, end_char=None, depth=0):
    """Walks JS from i; returns (pieces, i) where pieces is a list of ('code'|'str'|'tpl'..., text)."""
    out = []
    buf = []
    def flush():
        if buf: out.append(('code', ''.join(buf))); buf.clear()
    last_sig = '('
    while i < len(code):
        c = code[i]
        if end_char == '}' and c == '}' and depth == 0:
            flush(); return out, i
        if c == '{': depth += 1
        elif c == '}': depth -= 1
        if code.startswith('//', i):
            j = code.find('\n', i); j = len(code) if j < 0 else j
            buf.append(code[i:j]); i = j; continue
        if code.startswith('/*', i):
            j = code.find('*/', i) + 2; buf.append(code[i:j]); i = j; continue
        if c in '\'"':
            j = i + 1
            while code[j] != c:
                j += 2 if code[j] == '\\' else 1
            flush(); out.append(('str', code[i:j + 1])); i = j + 1; last_sig = 'x'; continue
        if c == '`':
            flush(); parts, i = lex_tpl(code, i + 1); out.append(('tpl', parts)); last_sig = 'x'; continue
        if c == '/' and (last_sig in '(,=:[!&|?{};+-*%<>~^' or re.search(r'(return|typeof|case)\s*$', ''.join(buf[-12:]))):
            j = i + 1; cls = False
            while True:
                ch = code[j]
                if ch == '\\': j += 2; continue
                if ch == '[': cls = True
                elif ch == ']': cls = False
                elif ch == '/' and not cls: break
                j += 1
            j += 1
            while j < len(code) and code[j].isalpha(): j += 1
            buf.append(code[i:j]); i = j; last_sig = 'x'; continue
        buf.append(c)
        if not c.isspace(): last_sig = c if not (c.isalnum() or c in '_$') else 'x'
        i += 1
    flush(); return out, i

def lex_tpl(code, i):
    parts, buf = [], []
    while True:
        c = code[i]
        if c == '\\': buf.append(code[i:i + 2]); i += 2; continue
        if c == '`':
            parts.append(('static', ''.join(buf))); return parts, i + 1
        if code.startswith('${', i):
            parts.append(('static', ''.join(buf))); buf = []
            inner, j = lex(code, i + 2, '}')
            parts.append(('expr', inner)); i = j + 1; continue
        buf.append(c); i += 1

KEYS = []
def key(k):
    if k not in KEYS: KEYS.append(k)

def emit(pieces):
    s = []
    for kind, val in pieces:
        if kind == 'code': s.append(val)
        elif kind == 'str':
            raw = val[1:-1]
            v = unescape(raw)
            if ''.join(s).rstrip().endswith('L('):   # already wrapped
                key(v); s.append(val); continue
            if AR.search(v) and v not in SKIP:
                key(v); s.append('L(' + val + ')')
            else: s.append(val)
        elif kind == 'tpl':
            t = ['`']
            for pk, pv in val:
                if pk == 'static':
                    def wrap(m):
                        txt = m.group(0)
                        if txt in SKIP: return txt
                        v = unescape(txt)
                        key(v)
                        return '${L(' + json.dumps(v, ensure_ascii=False).replace('`', '\\`') + ')}'
                    t.append(RUN.sub(wrap, pv))
                else:
                    t.append('${' + emit(pv) + '}')
            t.append('`'); s.append(''.join(t))
    return ''.join(s)

def main(mode):
    html = open(SRC, encoding='utf-8').read()
    start = html.index("'use strict';")
    end = html.index('</script>', start)
    code = html[start:end]
    a, b = code.find('/*i18n-skip*/'), code.find('/*/i18n-skip*/')
    skipped = ''
    if a >= 0:   # the English dictionary itself is never wrapped
        skipped = code[a:b + len('/*/i18n-skip*/')]
        code = code[:a] + '/*@@SKIP@@*/' + code[b + len('/*/i18n-skip*/'):]
    pieces, _ = lex(code, 0)
    new = emit(pieces).replace('/*@@SKIP@@*/', skipped)
    if mode != 'scan':
        en = json.load(open(os.path.join(HERE, 'tools', 'i18n_en.json'), encoding='utf-8')) if os.path.exists(os.path.join(HERE, 'tools', 'i18n_en.json')) else {}
        missing = [k for k in KEYS if k not in en]
        if missing: print('NOT IN DICTIONARY (add to tools/i18n_en.json):', missing)
    if mode == 'scan':
        json.dump(KEYS, open(os.path.join(HERE, 'tools', 'i18n_keys.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
        print(len(KEYS), 'phrases')
    else:
        open(SRC, 'w', encoding='utf-8', newline='\n').write(html[:start] + new + html[end:])
        print('wrapped', len(KEYS), 'phrases')

main(sys.argv[1] if len(sys.argv) > 1 else 'scan')
