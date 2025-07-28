# エラーが発生している関数一覧（59個）

最終更新: 2025-07-28

## 概要
E2Eテストで確認されたエラー関数の一覧です。全506関数中59個でエラーが発生しています。

## カテゴリ別エラー関数リスト

### 01. 数学・三角関数（1個）
1. **SUMPRODUCT** - 計算結果が異常（5040を返す、期待値は56）

### 03. テキスト関数（4個）
2. **TEXTJOIN** - #VALUE!エラー
3. **TEXTAFTER** - #NAME?エラー（未実装の可能性）
4. **UNICHAR** - #NAME?エラー
5. **FIXED** - #NAME?エラー

### 04. 日付・時刻関数（1個）
6. **WEEKDAY** - 日付の日の部分を返している（期待値は曜日番号）

### 05. 論理関数（5個）
7. **IF** - #NAME?エラー
8. **NOT** - #NAME?エラー
9. **XOR** - #NAME?エラー
10. **IFERROR** - #NAME?エラー
11. **IFNA** - #NAME?エラー

### 06. 検索・参照関数（4個）
12. **XLOOKUP** - #NAME?エラー
13. **OFFSET** - #NAME?エラー
14. **INDIRECT** - #NAME?エラー
15. **FORMULATEXT** - #NAME?エラー

### 07. 財務関数（18個）
16. **NPV** - #VALUE!エラー
17. **DDB** - #VALUE!エラー
18. **VDB** - #VALUE!エラー
19. **MDURATION** - #VALUE!エラー
20. **DURATION** - #VALUE!エラー
21. **PRICE** - #VALUE!エラー
22. **YIELD** - #VALUE!エラー
23. **COUPDAYBS** - #VALUE!エラー
24. **COUPDAYSNC** - #VALUE!エラー
25. **COUPNCD** - #VALUE!エラー
26. **COUPNUM** - #VALUE!エラー
27. **COUPPCD** - #VALUE!エラー
28. **AMORDEGRC** - #VALUE!エラー
29. **DOLLARDE** - #VALUE!エラー
30. **DOLLARFR** - #VALUE!エラー
31. **EFFECT** - #VALUE!エラー
32. **PRICEMAT** - #VALUE!エラー
33. **YIELDMAT** - #VALUE!エラー

### 08. 行列関数（3個）
34. **MINVERSE** - #VALUE!エラー
35. **MMULT** - #VALUE!エラー
36. **MUNIT** - #VALUE!エラー

### 09. 情報関数（8個）
37. **ISERROR** - #NAME?エラー
38. **ISNA** - #NAME?エラー
39. **ISTEXT** - #NAME?エラー
40. **ISNUMBER** - #NAME?エラー
41. **N** - #NAME?エラー
42. **ERROR.TYPE** - #NAME?エラー
43. **ISFORMULA** - #NAME?エラー
44. **ISBETWEEN** - #NAME?エラー

### 10. データベース関数（8個）
45. **DSUM** - #VALUE!エラー
46. **DAVERAGE** - #VALUE!エラー
47. **DCOUNT** - #VALUE!エラー
48. **DMAX** - #VALUE!エラー
49. **DMIN** - #VALUE!エラー
50. **DPRODUCT** - #VALUE!エラー
51. **DSTDEV** - #VALUE!エラー
52. **DVAR** - #VALUE!エラー

### 11. エンジニアリング関数（5個）
53. **BIN2DEC** - #NAME?エラー
54. **COMPLEX** - #NAME?エラー
55. **OCT2BIN** - #NAME?エラー
56. **IMABS** - #NAME?エラー
57. **BITAND** - #NAME?エラー

### 12. 動的配列関数（2個）
58. **ARRAYTOTEXT** - #NAME?エラー
59. **TOROW** - #NAME?エラー

## エラーの種類別分類

### #NAME?エラー（34個）
関数が認識されていない、またはパターンマッチングに失敗している
- テキスト: TEXTAFTER, UNICHAR, FIXED
- 論理: IF, NOT, XOR, IFERROR, IFNA
- 検索・参照: XLOOKUP, OFFSET, INDIRECT, FORMULATEXT
- 情報: ISERROR, ISNA, ISTEXT, ISNUMBER, N, ERROR.TYPE, ISFORMULA, ISBETWEEN
- エンジニアリング: BIN2DEC, COMPLEX, OCT2BIN, IMABS, BITAND
- 動的配列: ARRAYTOTEXT, TOROW

### #VALUE!エラー（24個）
関数は認識されているが、計算中にエラーが発生
- テキスト: TEXTJOIN
- 財務: NPV, DDB, VDB, MDURATION, DURATION, PRICE, YIELD, COUPDAYBS, COUPDAYSNC, COUPNCD, COUPNUM, COUPPCD, AMORDEGRC, DOLLARDE, DOLLARFR, EFFECT, PRICEMAT, YIELDMAT
- 行列: MINVERSE, MMULT, MUNIT
- データベース: DSUM, DAVERAGE, DCOUNT, DMAX, DMIN, DPRODUCT, DSTDEV, DVAR

### その他のエラー（1個）
- SUMPRODUCT: 計算結果が異常（5040を返す）
- WEEKDAY: 誤った値を返す（日付の日の部分）

## 修正優先順位

### 高優先度（基本的な関数）
1. IF, NOT, XOR - 基本的な論理関数
2. SUMPRODUCT - 計算エラー
3. WEEKDAY - 日付関数の基本

### 中優先度（よく使われる関数）
4. TEXTJOIN - テキスト結合
5. XLOOKUP, OFFSET, INDIRECT - 検索・参照
6. NPV, DDB - 財務計算
7. 情報関数群（ISERROR, ISNA等）

### 低優先度（特殊な関数）
8. データベース関数
9. エンジニアリング関数
10. 動的配列関数