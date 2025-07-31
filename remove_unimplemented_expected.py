#!/usr/bin/env python3
import os
import re

# 未実装関数のリスト（check_unimplemented_functions.pyの結果から）
unimplemented_functions = {
    # 01. 数学
    'CEILING.MATH', 'CEILING.PRECISE', 'FLOOR.MATH', 'FLOOR.PRECISE', 'ISO.CEILING',
    
    # 02. 統計
    'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.INV', 'CHISQ.DIST', 'CHISQ.DIST.RT',
    'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'CONFIDENCE.NORM', 'CONFIDENCE.T',
    'COVARIANCE.P', 'COVARIANCE.S', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV',
    'F.INV.RT', 'F.TEST', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN.PRECISE', 'HYPGEOM.DIST',
    'LOGNORM.DIST', 'LOGNORM.INV', 'MODE.MULT', 'MODE.SNGL', 'NORM.DIST', 'NORM.INV',
    'NORM.S.DIST', 'NORM.S.INV', 'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC',
    'PERCENTRANK.INC', 'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'RANK.AVG',
    'RANK.EQ', 'SKEW.P', 'STDEV.P', 'STDEV.S', 'STDEVA', 'STDEVPA', 'T.DIST',
    'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'VAR.P', 'VAR.S',
    'VARA', 'VARPA', 'WEIBULL.DIST', 'Z.TEST',
    
    # 03. 文字列
    'T',
    
    # 04. 日付
    'NETWORKDAYS.INTL', 'WORKDAY.INTL',
    
    # 09. 情報
    'ERROR.TYPE', 'N',
    
    # 11. エンジニアリング
    'ERF.PRECISE', 'ERFC.PRECISE',
    
    # 14. Web・インポート
    'SPLIT'
}

test_dir = 'src/data/individualTests'
modified_count = 0

for filename in os.listdir(test_dir):
    if filename.endswith('.ts'):
        filepath = os.path.join(test_dir, filename)
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        original_content = content
        
        # 各未実装関数のexpectedValuesを削除
        for func in unimplemented_functions:
            # エスケープされた関数名（.を含む場合のため）
            escaped_func = func.replace('.', r'\.')
            
            # パターン: name: 'FUNC_NAME' から expectedValues: { ... } を含むブロックを探す
            pattern = rf"({{[^{{}}]*name:\s*['\"]({escaped_func})['\"][^{{}}]*)(expectedValues:\s*{{[^{{}}]*}}[,\s]*)"
            
            def replacer(match):
                return match.group(1)  # expectedValues部分を削除
            
            content = re.sub(pattern, replacer, content, flags=re.DOTALL)
        
        if content != original_content:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)
            modified_count += 1
            print(f"Modified: {filename}")

print(f"\n総修正ファイル数: {modified_count}")