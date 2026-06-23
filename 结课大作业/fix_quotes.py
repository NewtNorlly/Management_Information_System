#!/usr/bin/env python3
"""Fix Chinese quotation marks in generate_doc.py"""
import sys

with open(r'C:\Users\NewtN\下载\generate_doc.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace left/right double quotation marks with corner brackets
content = content.replace('\u201c', '\u300c')  # " -> 「
content = content.replace('\u201d', '\u300d')  # " -> 」

# Also replace any ASCII double-quotes used as Chinese quotes inside strings
# These are harder to detect automatically, so handle known patterns
# Replace leftover patterns where ASCII " appears as Chinese quote inside strings
import re
# Find: inside Python strings, 提供"提交" should keep the quotes as 「」
# But we already replaced \u201c\u201d, so any remaining " inside Chinese text
# is actually ASCII " (0x22) that would break the string
# We can't easily distinguish these from actual Python string delimiters
# So let's just check if the file compiles now

try:
    compile(content, 'generate_doc.py', 'exec')
    print('Syntax OK after U+201C/U+201D replacement')
except SyntaxError as e:
    print(f'Still has syntax error: {e}')
    print(f'Line {e.lineno}: {e.msg}')
    # Try to find the problematic line
    lines = content.split('\n')
    if e.lineno:
        lineno = e.lineno
        for i in range(max(0, lineno-3), min(len(lines), lineno+2)):
            line = lines[i]
            # Show the line, highlighting quote issues
            print(f'  Line {i+1}: has {line.count(chr(0x22))} ASCII double-quotes')
            if i+1 == lineno:
                # Find position
                print(f'  >>> {line[:80]}...')

with open(r'C:\Users\NewtN\下载\generate_doc.py', 'w', encoding='utf-8') as f:
    f.write(content)
print('File updated.')
