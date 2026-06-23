#!/usr/bin/env python3
"""Comprehensive fix for Chinese quotation marks in generate_doc.py"""
import re

LQ = '\u300c'  # 「
RQ = '\u300d'  # 」

with open(r'C:\Users\NewtN\下载\generate_doc.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace pairs of ASCII " used as Chinese quotes within Python strings
# Pattern: Chinese_char " text " Chinese_char_or_punctuation
pattern1 = re.compile(r'([\u4e00-\u9fff\u3001\uff0c])"([\u4e00-\u9fff\u2215\u3001\u4e00-\u9fff\-]+)"([\u4e00-\u9fff\u3001\uff0c\u3002\uff1f\uff01\uff0e\u300a\u300b])')
content = pattern1.sub(lambda m: m.group(1) + LQ + m.group(2) + RQ + m.group(3), content)

# Also handle 、 or punctuation boundaries
pattern2 = re.compile(r'"([\u4e00-\u9fff\u2215\u3001\-]+)"([\u4e00-\u9fff\u3001\uff0c\u3002])')
content = pattern2.sub(lambda m: LQ + m.group(1) + RQ + m.group(2), content)

# Also handle Chinese char " text " at end or start
pattern3 = re.compile(r'([\u4e00-\u9fff])"([\u4e00-\u9fff\u2215\u3001\-]+)"')
content = pattern3.sub(lambda m: m.group(1) + LQ + m.group(2) + RQ, content)

with open(r'C:\Users\NewtN\下载\generate_doc.py', 'w', encoding='utf-8') as f:
    f.write(content)

# Verify syntax
try:
    compile(content, 'generate_doc.py', 'exec')
    print('SUCCESS: Syntax is valid!')
except SyntaxError as e:
    print(f'Still failing at line {e.lineno}: {e.msg}')
    lines = content.split('\n')
    if e.lineno:
        for j in range(max(0, e.lineno-3), min(len(lines), e.lineno+2)):
            snippet = lines[j][:150]
            # Count ASCII double quotes
            dq = snippet.count('"')
            print(f'  L{j+1} (dq={dq}): {snippet}')
