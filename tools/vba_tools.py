"""VBA ソースのビルド (行番号付け) と規約チェックを行う開発用スクリプト。

src/      編集用。行番号なし、LF。
modules/  VBE に貼り付ける用。行番号付き (10 刻み)、CRLF。build で生成する。

使い方:
    python tools/vba_tools.py build   # src/*.bas -> modules/*.bas
    python tools/vba_tools.py lint    # src/*.bas を GitHub/VBA/CLAUDE.md の規約で確かめる

行番号を付けない行:
    宣言 (Dim / Const / Static)、コメント、空行、ラベル (ErrHandler: など)、継続行、
    Else / ElseIf / Case / End ... / Loop / Next / Wend、On Error、LogSuccess、
    最初のラベル (CleanUp: / ErrHandler:) より後ろ
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / 'src'
OUT_DIR = ROOT / 'modules'

PROC_START = re.compile(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property (Get|Let|Set)) (\w+)', re.I)
PROC_END = re.compile(r'^End (Sub|Function|Property)\b', re.I)
LABEL = re.compile(r'^[A-Za-z_]\w*:\s*$')
NO_NUMBER = re.compile(
    r'^(Dim |Const |Static |Else\b|ElseIf\b|Case\b|End (If|Select|With)\b|'
    r'Loop\b|Next\b|Wend\b|On Error\b|LogSuccess\b|\')', re.I)
BODY_INDENT = 4     # src の本文の字下げ
NUM_WIDTH = 6       # modules での行番号欄の幅 (本文は 2 桁右へずれる)
LINE_STEP = 10
SHARED_MODULE = 'ErrorLogger.bas'  # GitHub/VBA の共有モジュール (コピーして使う)

FULLWIDTH_FORBIDDEN = re.compile('[（）：！？　０-９]')
NON_ASCII = re.compile(r'[^\x00-\x7F]')
CONST_DEF = re.compile(r'^\s*(Public |Private )?Const (\w+)', re.I)
UPPER_TOKEN = re.compile(r'\b([A-Z][A-Z0-9]*_[A-Z0-9_]+)\b')
STRING_LIT = re.compile(r'"(?:[^"]|"")*"')
HEX_LIT = re.compile(r'&H([0-9A-Fa-f]{4,8})(?!&)')


def code_part(line):
    """行末コメントを除いたコード部分 (文字列リテラルの中の ' は無視)。"""
    out, in_str = [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
        if ch == "'" and not in_str:
            break
        out.append(ch)
    return ''.join(out)


def build_file(src_path, out_path):
    lines = src_path.read_text(encoding='utf-8').replace('\r\n', '\n').split('\n')
    out = []
    in_proc = after_label = prev_cont = False
    n = 0
    for line in lines:
        if not in_proc and PROC_START.match(line):
            in_proc, after_label, n = True, False, 0
            out.append(line)
            prev_cont = code_part(line).rstrip().endswith(' _')
            continue
        if in_proc and PROC_END.match(line):
            in_proc = False
            out.append(line)
            prev_cont = False
            continue
        if not in_proc:
            out.append(line)
            continue
        stripped = line.strip()
        if LABEL.match(line):
            after_label = True
            out.append(line)
            prev_cont = False
            continue
        if not stripped:
            out.append('')
            prev_cont = False
            continue
        if not line.startswith(' ' * BODY_INDENT):
            raise SystemExit(f'{src_path.name}: 本文の字下げが {BODY_INDENT} 未満: {line!r}')
        body = line[BODY_INDENT:]
        numberable = not after_label and not prev_cont and not NO_NUMBER.match(stripped)
        if numberable:
            n += LINE_STEP
            out.append(str(n).ljust(NUM_WIDTH) + body)
        else:
            out.append(' ' * NUM_WIDTH + body)
        prev_cont = code_part(line).rstrip().endswith(' _')
    out_path.write_bytes('\r\n'.join(out).encode('utf-8'))


def build():
    OUT_DIR.mkdir(exist_ok=True)
    for src in sorted(SRC_DIR.glob('*.bas')):
        if src.name == SHARED_MODULE:
            # 共有モジュールは手を加えず、改行だけ CRLF にする
            text = src.read_text(encoding='utf-8').replace('\r\n', '\n').replace('\n', '\r\n')
            (OUT_DIR / src.name).write_bytes(text.encode('utf-8'))
        else:
            build_file(src, OUT_DIR / src.name)
        print(f'built: modules/{src.name}')


def lint():
    problems = []
    consts = set()
    sources = {}
    for p in sorted(SRC_DIR.glob('*.bas')):
        lines = p.read_text(encoding='utf-8').replace('\r\n', '\n').split('\n')
        sources[p.name] = lines
        for line in lines:
            m = CONST_DEF.match(line)
            if m:
                consts.add(m.group(2))
    for name, lines in sources.items():
        if name == SHARED_MODULE:
            continue  # 共有モジュールは GitHub/VBA 側で管理する
        if not any(l.strip() == 'Option Explicit' for l in lines):
            problems.append(f'{name}: Option Explicit がない')
        if not any(l.startswith('Private Const MODULE_NAME') for l in lines):
            problems.append(f'{name}: MODULE_NAME 定数がない')
        proc = None
        in_const = False
        for i, line in enumerate(lines, 1):
            code = code_part(line)
            if FULLWIDTH_FORBIDDEN.search(code):
                problems.append(f'{name}:{i}: 全角の禁止文字 {FULLWIDTH_FORBIDDEN.findall(code)}')
            if '\t' in line:
                problems.append(f'{name}:{i}: タブ文字')
            is_const = bool(CONST_DEF.match(code)) or in_const
            in_const = is_const and code.rstrip().endswith(' _')
            if not is_const and not code.startswith('Attribute'):
                for lit in STRING_LIT.findall(code):
                    if NON_ASCII.search(lit):
                        problems.append(f'{name}:{i}: 定数以外に日本語の文字列 {lit[:40]}')
            for tok in UPPER_TOKEN.findall(STRING_LIT.sub('""', code)):
                if tok not in consts:
                    problems.append(f'{name}:{i}: 未定義の定数 {tok}')
            m = PROC_START.match(line)
            if m:
                proc, start = m.group(5), i
                flags = {'on_error': False, 'log_error': False, 'log_success': False}
                continue
            if proc:
                if 'On Error GoTo ErrHandler' in code:
                    flags['on_error'] = True
                if re.search(r'\bLogError\b', code):
                    flags['log_error'] = True
                if re.search(r'\bLogSuccess\b', code):
                    flags['log_success'] = True
                if PROC_END.match(line):
                    for key, label in (('on_error', 'On Error GoTo ErrHandler'), ('log_error', 'LogError'),
                                       ('log_success', 'LogSuccess')):
                        if not flags[key]:
                            problems.append(f'{name}:{start}: {proc} に {label} がない')
                    proc = None
    for pr in problems:
        print(pr)
    print(f'{len(problems)} problem(s)')
    return 1 if problems else 0


def main():
    if len(sys.argv) != 2 or sys.argv[1] not in ('build', 'lint'):
        print(__doc__)
        sys.exit(2)
    if sys.argv[1] == 'build':
        build()
    else:
        sys.exit(lint())


if __name__ == '__main__':
    main()
