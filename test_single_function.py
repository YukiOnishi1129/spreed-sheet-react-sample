#!/usr/bin/env python3
import os
import re

# 1つの関数だけテスト
test_function = 'SERIESSUM'

test_dir = 'src/data/individualTests'
for filename in os.listdir(test_dir):
    if filename.endswith('.ts'):
        filepath = os.path.join(test_dir, filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 関数を探す - より柔軟なパターン
        pattern = rf"name:\s*['\"]({test_function})['\"]"
        match = re.search(pattern, content)
        
        if match:
            print(f"Found {test_function} in {filename}")
            # 関数のブロック全体を取得
            start = content.rfind('{', 0, match.start())
            brace_count = 1
            end = start + 1
            while brace_count > 0 and end < len(content):
                if content[end] == '{':
                    brace_count += 1
                elif content[end] == '}':
                    brace_count -= 1
                end += 1
            
            block = content[start:end]
            print("Block content:")
            print(block[:500] + '...' if len(block) > 500 else block)
            
            # expectedValuesがあるか確認
            if 'expectedValues' in block:
                print("\nHas expectedValues")
                expected_match = re.search(r'expectedValues:\s*({[^}]+})', block)
                if expected_match:
                    print(f"expectedValues: {expected_match.group(1)}")
            else:
                print("\nNo expectedValues found")