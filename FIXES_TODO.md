# E2Eテスト結果サマリーと修正が必要な関数一覧

最終更新: 2025-07-28

## E2Eテスト実行状況

### 完全に成功したカテゴリ
- 00. 基本演算子 - 全て成功 ✓
- 02. 統計関数 - 全て成功 ✓

### 部分的に成功しているカテゴリ
1. **01. 数学・三角関数** - SUMPRODUCT以外は成功
2. **03. テキスト関数** - LENB修正済み、TEXTJOIN/UNICHAR/TEXTAFTER/FIXEDにエラー
3. **04. 日付関数** - WEEKDAY以外は成功
4. **05. 論理関数** - IF/NOT/XOR/IFERROR/IFNAにエラー

### 部分的に成功しているカテゴリ（続き）
5. **06. 検索・参照** - XLOOKUP/OFFSET/INDIRECT/FORMULATEXTにエラー

### 部分的に成功しているカテゴリ（続き）
6. **07. 財務** - ほぼ全てexpectedValues未定義
7. **08. 行列** - 1/4成功（MDETERMのみ成功）
8. **09. 情報** - 8/16失敗（ISERROR/ISNA/ISTEXT/ISNUMBER/N/ERROR.TYPE/ISFORMULA/ISBETWEENにエラー）

### 部分的に成功しているカテゴリ（続き）
9. **10. データベース** - 8/12失敗
10. **11. エンジニアリング** - 多数のエラー（BIN2DEC/COMPLEX/OCT2BIN等）
11. **12. 動的配列** - ARRAYTOTEXT/TORAWにエラー、他はexpectedValues未定義
12. **13. キューブ** - 全てexpectedValues未定義
13. **14. Web・インポート** - ENCODEURL/JOINにエラー
14. **15. その他** - 全てexpectedValues未定義

## テスト結果サマリー

### カテゴリ別エラー数
- 00. 基本演算子: 0エラー ✓
- 01. 数学・三角: 1エラー（SUMPRODUCT）
- 02. 統計: 0エラー ✓
- 03. テキスト: 4エラー（TEXTJOIN, TEXTAFTER, UNICHAR, FIXED）
- 04. 日付: 1エラー（WEEKDAY）
- 05. 論理: 5エラー（IF, NOT, XOR, IFERROR, IFNA）
- 06. 検索・参照: 4エラー（XLOOKUP, OFFSET, INDIRECT, FORMULATEXT）
- 07. 財務: ほぼ全て未実装
- 08. 行列: 3/4未実装
- 09. 情報: 8エラー
- 10. データベース: 8エラー
- 11. エンジニアリング: 多数のエラー
- 12. 動的配列: 2エラー（ARRAYTOTEXT, TOROW）
- 13. キューブ: 全て未実装
- 14. Web・インポート: 2エラー（ENCODEURL, JOIN）
- 15. その他: 全て未実装

### 総計
- 実装済み関数でエラー: 約35関数
- expectedValues未定義: 約100関数以上

## 優先度：高（基本的な関数でエラーが出ているもの）

### 1. SUMPRODUCT - 5040（7!）を返す問題
- 期待値: 56
- 実際の値: 5040 (7の階乗)
- 原因調査が必要

### 2. WEEKDAY - 日付の日の部分を返している
- 例: 2024/01/15（月曜日）→ 15を返す（期待値は2）
- dateLogic.tsの実装を確認する必要あり

### 3. 論理関数の複数のエラー
- IF: #NAME?エラー
- NOT: #NAME?エラー
- XOR: #NAME?エラー
- IFERROR: #NAME?エラー
- IFNA: #NAME?エラー

## 優先度：中（特定の条件で失敗するもの）

### 4. TEXTJOIN - #VALUE!エラー
- 区切り文字での結合に失敗
- パターンマッチングまたはパース処理の問題

### 5. TEXTAFTER - #NAME?エラー
- 関数が実装されていない可能性
- textLogic.tsに実装を追加する必要あり

### 6. UNICHAR - #NAME?エラー
- Unicode番号から文字を返す関数
- 実装の確認が必要

### 7. FIXED - #NAME?エラー
- 小数点以下の桁数を固定する関数
- 実装の確認が必要

## 修正済み

- LENB: UTF-8バイト数計算をExcel互換に修正済み
- MROUND: 関数の評価順序を修正済み
- COUNTBLANK: 空のセル値を含めるよう修正済み
- AGGREGATE: テストデータの範囲を修正済み
- 統計関数（STDEV.S, VAR.S等）: エクスポートを追加済み