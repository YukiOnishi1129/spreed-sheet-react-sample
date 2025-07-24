// テスト用ダミーデータ
export interface TestCategory {
  id: string;
  name: string;
  description: string;
  color: string;
  data: (string | number | null)[][];
  formulas: { row: number; col: number; formula: string }[];
  expectedResults?: { row: number; col: number; value: string | number }[];
}

export const testCategories: TestCategory[] = [
  {
    id: 'math',
    name: '数学・三角関数',
    description: '基本的な計算、集計、条件付き集計などのテスト',
    color: 'orange',
    data: [
      ['項目', '1月', '2月', '3月', '合計', '平均'],
      ['売上A', 100000, 150000, 120000, null, null],
      ['売上B', 80000, 90000, 110000, null, null],
      ['売上C', 120000, 100000, 130000, null, null],
      ['合計', null, null, null, null, null],
      ['', '', '', '', '', ''],
      ['商品', '単価', '数量', '金額', '税込'],
      ['商品A', 1000, 5, null, null],
      ['商品B', 2500, 3, null, null],
      ['商品C', 800, 10, null, null],
    ],
    formulas: [
      // 行の合計
      { row: 1, col: 4, formula: '=SUM(B2:D2)' },
      { row: 2, col: 4, formula: '=SUM(B3:D3)' },
      { row: 3, col: 4, formula: '=SUM(B4:D4)' },
      // 列の合計
      { row: 4, col: 1, formula: '=SUM(B2:B4)' },
      { row: 4, col: 2, formula: '=SUM(C2:C4)' },
      { row: 4, col: 3, formula: '=SUM(D2:D4)' },
      // 平均
      { row: 1, col: 5, formula: '=AVERAGE(B2:D2)' },
      { row: 2, col: 5, formula: '=AVERAGE(B3:D3)' },
      { row: 3, col: 5, formula: '=AVERAGE(B4:D4)' },
      // 金額計算
      { row: 7, col: 3, formula: '=B8*C8' },
      { row: 8, col: 3, formula: '=B9*C9' },
      { row: 9, col: 3, formula: '=B10*C10' },
      // 税込計算（10%）
      { row: 7, col: 4, formula: '=D8*1.1' },
      { row: 8, col: 4, formula: '=D9*1.1' },
      { row: 9, col: 4, formula: '=D10*1.1' },
    ],
    expectedResults: [
      { row: 1, col: 4, value: 370000 },
      { row: 2, col: 4, value: 280000 },
      { row: 3, col: 4, value: 350000 },
      { row: 4, col: 1, value: 300000 },
      { row: 4, col: 2, value: 340000 },
      { row: 4, col: 3, value: 360000 },
      { row: 1, col: 5, value: 123333.33 },
      { row: 7, col: 3, value: 5000 },
      { row: 8, col: 3, value: 7500 },
      { row: 9, col: 3, value: 8000 },
      { row: 7, col: 4, value: 5500 },
      { row: 8, col: 4, value: 8250 },
      { row: 9, col: 4, value: 8800 },
    ]
  },
  {
    id: 'logical',
    name: '論理関数',
    description: 'IF、AND、OR、比較演算子のテスト',
    color: 'purple',
    data: [
      ['点数', '判定', '合否', 'ランク'],
      [85, null, null, null],
      [92, null, null, null],
      [78, null, null, null],
      [65, null, null, null],
      [40, null, null, null],
      ['', '', '', ''],
      ['売上', '目標', '達成', 'ボーナス'],
      [120000, 100000, null, null],
      [85000, 100000, null, null],
      [150000, 100000, null, null],
    ],
    formulas: [
      // 判定（80点以上で合格）
      { row: 1, col: 1, formula: '=IF(A2>=80,"合格","不合格")' },
      { row: 2, col: 1, formula: '=IF(A3>=80,"合格","不合格")' },
      { row: 3, col: 1, formula: '=IF(A4>=80,"合格","不合格")' },
      { row: 4, col: 1, formula: '=IF(A5>=80,"合格","不合格")' },
      { row: 5, col: 1, formula: '=IF(A6>=80,"合格","不合格")' },
      // 合否（60点以上で可）
      { row: 1, col: 2, formula: '=IF(A2>=60,"可","不可")' },
      { row: 2, col: 2, formula: '=IF(A3>=60,"可","不可")' },
      { row: 3, col: 2, formula: '=IF(A4>=60,"可","不可")' },
      { row: 4, col: 2, formula: '=IF(A5>=60,"可","不可")' },
      { row: 5, col: 2, formula: '=IF(A6>=60,"可","不可")' },
      // ランク付け（IFS関数）
      { row: 1, col: 3, formula: '=IFS(A2>=90,"A",A2>=80,"B",A2>=70,"C",A2>=60,"D",TRUE,"F")' },
      { row: 2, col: 3, formula: '=IFS(A3>=90,"A",A3>=80,"B",A3>=70,"C",A3>=60,"D",TRUE,"F")' },
      { row: 3, col: 3, formula: '=IFS(A4>=90,"A",A4>=80,"B",A4>=70,"C",A4>=60,"D",TRUE,"F")' },
      { row: 4, col: 3, formula: '=IFS(A5>=90,"A",A5>=80,"B",A5>=70,"C",A5>=60,"D",TRUE,"F")' },
      { row: 5, col: 3, formula: '=IFS(A6>=90,"A",A6>=80,"B",A6>=70,"C",A6>=60,"D",TRUE,"F")' },
      // 目標達成判定
      { row: 8, col: 2, formula: '=IF(A9>=B9,"達成","未達成")' },
      { row: 9, col: 2, formula: '=IF(A10>=B10,"達成","未達成")' },
      { row: 10, col: 2, formula: '=IF(A11>=B11,"達成","未達成")' },
      // ボーナス計算（達成時は売上の10%）
      { row: 8, col: 3, formula: '=IF(A9>=B9,A9*0.1,0)' },
      { row: 9, col: 3, formula: '=IF(A10>=B10,A10*0.1,0)' },
      { row: 10, col: 3, formula: '=IF(A11>=B11,A11*0.1,0)' },
    ],
    expectedResults: [
      { row: 1, col: 1, value: '合格' },
      { row: 2, col: 1, value: '合格' },
      { row: 3, col: 1, value: '不合格' },
      { row: 4, col: 1, value: '不合格' },
      { row: 5, col: 1, value: '不合格' },
      { row: 1, col: 3, value: 'B' },
      { row: 2, col: 3, value: 'A' },
      { row: 3, col: 3, value: 'C' },
      { row: 4, col: 3, value: 'D' },
      { row: 5, col: 3, value: 'F' },
      { row: 8, col: 2, value: '達成' },
      { row: 9, col: 2, value: '未達成' },
      { row: 10, col: 2, value: '達成' },
      { row: 8, col: 3, value: 12000 },
      { row: 9, col: 3, value: 0 },
      { row: 10, col: 3, value: 15000 },
    ]
  },
  {
    id: 'text',
    name: 'テキスト関数',
    description: '文字列操作、結合、抽出のテスト',
    color: 'green',
    data: [
      ['姓', '名', 'フルネーム', 'イニシャル'],
      ['山田', '太郎', null, null],
      ['佐藤', '花子', null, null],
      ['鈴木', '一郎', null, null],
      ['', '', '', ''],
      ['メール', 'ユーザー名', 'ドメイン', '大文字'],
      ['taro.yamada@example.com', null, null, null],
      ['hanako.sato@test.co.jp', null, null, null],
      ['', '', '', ''],
      ['コード', '先頭3文字', '4文字目以降', '長さ'],
      ['ABC12345', null, null, null],
      ['XYZ789', null, null, null],
    ],
    formulas: [
      // フルネーム結合
      { row: 1, col: 2, formula: '=CONCATENATE(A2," ",B2)' },
      { row: 2, col: 2, formula: '=CONCATENATE(A3," ",B3)' },
      { row: 3, col: 2, formula: '=CONCATENATE(A4," ",B4)' },
      // イニシャル（左端1文字）
      { row: 1, col: 3, formula: '=CONCATENATE(LEFT(A2,1),".",LEFT(B2,1),".")' },
      { row: 2, col: 3, formula: '=CONCATENATE(LEFT(A3,1),".",LEFT(B3,1),".")' },
      { row: 3, col: 3, formula: '=CONCATENATE(LEFT(A4,1),".",LEFT(B4,1),".")' },
      // メールアドレスからユーザー名抽出
      { row: 6, col: 1, formula: '=LEFT(A7,FIND("@",A7)-1)' },
      { row: 7, col: 1, formula: '=LEFT(A8,FIND("@",A8)-1)' },
      // ドメイン抽出
      { row: 6, col: 2, formula: '=MID(A7,FIND("@",A7)+1,LEN(A7))' },
      { row: 7, col: 2, formula: '=MID(A8,FIND("@",A8)+1,LEN(A8))' },
      // 大文字変換
      { row: 6, col: 3, formula: '=UPPER(A7)' },
      { row: 7, col: 3, formula: '=UPPER(A8)' },
      // 文字列の切り出し
      { row: 10, col: 1, formula: '=LEFT(A11,3)' },
      { row: 11, col: 1, formula: '=LEFT(A12,3)' },
      { row: 10, col: 2, formula: '=MID(A11,4,LEN(A11))' },
      { row: 11, col: 2, formula: '=MID(A12,4,LEN(A12))' },
      // 文字列の長さ
      { row: 10, col: 3, formula: '=LEN(A11)' },
      { row: 11, col: 3, formula: '=LEN(A12)' },
    ],
    expectedResults: [
      { row: 1, col: 2, value: '山田 太郎' },
      { row: 2, col: 2, value: '佐藤 花子' },
      { row: 3, col: 2, value: '鈴木 一郎' },
      { row: 1, col: 3, value: '山.太.' },
      { row: 2, col: 3, value: '佐.花.' },
      { row: 3, col: 3, value: '鈴.一.' },
      { row: 6, col: 1, value: 'taro.yamada' },
      { row: 7, col: 1, value: 'hanako.sato' },
      { row: 6, col: 2, value: 'example.com' },
      { row: 7, col: 2, value: 'test.co.jp' },
      { row: 10, col: 1, value: 'ABC' },
      { row: 11, col: 1, value: 'XYZ' },
      { row: 10, col: 3, value: 8 },
      { row: 11, col: 3, value: 6 },
    ]
  },
  {
    id: 'datetime',
    name: '日付・時刻関数',
    description: '日付計算、期間計算のテスト',
    color: 'blue',
    data: [
      ['開始日', '終了日', '日数', '営業日数', '年齢'],
      ['2024-01-01', '2024-12-31', null, null, null],
      ['2024-03-15', '2024-06-20', null, null, null],
      ['', '', '', '', ''],
      ['生年月日', '年齢（年）', '年齢（月）', '年齢（日）'],
      ['1990-05-15', null, null, null],
      ['2000-12-25', null, null, null],
      ['', '', '', '', ''],
      ['日付', '年', '月', '日', '曜日'],
      ['2024-07-15', null, null, null, null],
      ['2024-12-25', null, null, null, null],
    ],
    formulas: [
      // 日数計算
      { row: 1, col: 2, formula: '=DAYS(B2,A2)' },
      { row: 2, col: 2, formula: '=DAYS(B3,A3)' },
      // 営業日数計算
      { row: 1, col: 3, formula: '=NETWORKDAYS(A2,B2)' },
      { row: 2, col: 3, formula: '=NETWORKDAYS(A3,B3)' },
      // 年齢計算（DATEDIFを使用）
      { row: 5, col: 1, formula: '=DATEDIF(A6,TODAY(),"Y")' },
      { row: 6, col: 1, formula: '=DATEDIF(A7,TODAY(),"Y")' },
      { row: 5, col: 2, formula: '=DATEDIF(A6,TODAY(),"M")' },
      { row: 6, col: 2, formula: '=DATEDIF(A7,TODAY(),"M")' },
      { row: 5, col: 3, formula: '=DATEDIF(A6,TODAY(),"D")' },
      { row: 6, col: 3, formula: '=DATEDIF(A7,TODAY(),"D")' },
      // 日付から年月日を抽出
      { row: 9, col: 1, formula: '=YEAR(A10)' },
      { row: 10, col: 1, formula: '=YEAR(A11)' },
      { row: 9, col: 2, formula: '=MONTH(A10)' },
      { row: 10, col: 2, formula: '=MONTH(A11)' },
      { row: 9, col: 3, formula: '=DAY(A10)' },
      { row: 10, col: 3, formula: '=DAY(A11)' },
      // 曜日（1=日曜日）
      { row: 9, col: 4, formula: '=WEEKDAY(A10)' },
      { row: 10, col: 4, formula: '=WEEKDAY(A11)' },
    ],
    expectedResults: [
      { row: 1, col: 2, value: 365 },
      { row: 2, col: 2, value: 97 },
      { row: 9, col: 1, value: 2024 },
      { row: 10, col: 1, value: 2024 },
      { row: 9, col: 2, value: 7 },
      { row: 10, col: 2, value: 12 },
      { row: 9, col: 3, value: 15 },
      { row: 10, col: 3, value: 25 },
      { row: 9, col: 4, value: 2 }, // 月曜日
      { row: 10, col: 4, value: 4 }, // 水曜日
    ]
  },
  {
    id: 'statistical',
    name: '統計関数',
    description: '統計計算、集計のテスト',
    color: 'red',
    data: [
      ['データ', '個数', '最大', '最小', '中央値', '標準偏差'],
      [85, null, null, null, null, null],
      [92, null, null, null, null, null],
      [78, null, null, null, null, null],
      [88, null, null, null, null, null],
      [95, null, null, null, null, null],
      [82, null, null, null, null, null],
      [null, null, null, null, null, null],
      [90, null, null, null, null, null],
      ['', '', '', '', '', ''],
      ['カテゴリ', '値', '', '条件カウント', '条件平均'],
      ['A', 100, '', 'カテゴリA:', null],
      ['B', 150, '', 'カテゴリB:', null],
      ['A', 120, '', 'カテゴリC:', null],
      ['C', 80, '', '', ''],
      ['B', 130, '', '', ''],
      ['A', 110, '', '', ''],
      ['C', 90, '', '', ''],
    ],
    formulas: [
      // データの個数（空白を除く）
      { row: 1, col: 1, formula: '=COUNTA(A2:A9)' },
      // 最大値
      { row: 1, col: 2, formula: '=MAX(A2:A9)' },
      // 最小値
      { row: 1, col: 3, formula: '=MIN(A2:A9)' },
      // 中央値
      { row: 1, col: 4, formula: '=MEDIAN(A2:A9)' },
      // 標準偏差
      { row: 1, col: 5, formula: '=STDEV(A2:A9)' },
      // 条件付きカウント
      { row: 11, col: 4, formula: '=COUNTIF(A12:A18,"A")' },
      { row: 12, col: 4, formula: '=COUNTIF(A12:A18,"B")' },
      { row: 13, col: 4, formula: '=COUNTIF(A12:A18,"C")' },
      // 条件付き平均
      { row: 11, col: 3, formula: '=AVERAGEIF(A12:A18,"A",B12:B18)' },
      { row: 12, col: 3, formula: '=AVERAGEIF(A12:A18,"B",B12:B18)' },
      { row: 13, col: 3, formula: '=AVERAGEIF(A12:A18,"C",B12:B18)' },
    ],
    expectedResults: [
      { row: 1, col: 1, value: 7 },
      { row: 1, col: 2, value: 95 },
      { row: 1, col: 3, value: 78 },
      { row: 1, col: 4, value: 88 },
      { row: 11, col: 4, value: 3 },
      { row: 12, col: 4, value: 2 },
      { row: 13, col: 4, value: 2 },
      { row: 11, col: 3, value: 110 },
      { row: 12, col: 3, value: 140 },
      { row: 13, col: 3, value: 85 },
    ]
  },
  {
    id: 'lookup',
    name: '検索・参照関数',
    description: 'VLOOKUP、HLOOKUP、INDEX、MATCHのテスト',
    color: 'cyan',
    data: [
      ['商品コード', '商品名', '単価', '在庫'],
      ['A001', 'ノートPC', 120000, 5],
      ['A002', 'マウス', 3000, 50],
      ['A003', 'キーボード', 8000, 20],
      ['A004', 'モニター', 45000, 8],
      ['', '', '', ''],
      ['検索コード', '商品名', '単価', '在庫'],
      ['A002', null, null, null],
      ['A004', null, null, null],
      ['A001', null, null, null],
      ['', '', '', ''],
      ['順位', '1位', '2位', '3位'],
      ['商品名', null, null, null],
    ],
    formulas: [
      // VLOOKUP検索
      { row: 7, col: 1, formula: '=VLOOKUP(A8,$A$2:$D$5,2,FALSE)' },
      { row: 8, col: 1, formula: '=VLOOKUP(A9,$A$2:$D$5,2,FALSE)' },
      { row: 9, col: 1, formula: '=VLOOKUP(A10,$A$2:$D$5,2,FALSE)' },
      // 単価検索
      { row: 7, col: 2, formula: '=VLOOKUP(A8,$A$2:$D$5,3,FALSE)' },
      { row: 8, col: 2, formula: '=VLOOKUP(A9,$A$2:$D$5,3,FALSE)' },
      { row: 9, col: 2, formula: '=VLOOKUP(A10,$A$2:$D$5,3,FALSE)' },
      // 在庫検索
      { row: 7, col: 3, formula: '=VLOOKUP(A8,$A$2:$D$5,4,FALSE)' },
      { row: 8, col: 3, formula: '=VLOOKUP(A9,$A$2:$D$5,4,FALSE)' },
      { row: 9, col: 3, formula: '=VLOOKUP(A10,$A$2:$D$5,4,FALSE)' },
      // INDEX+MATCHで単価順位の商品名を取得
      { row: 13, col: 1, formula: '=INDEX($B$2:$B$5,MATCH(LARGE($C$2:$C$5,1),$C$2:$C$5,0))' },
      { row: 13, col: 2, formula: '=INDEX($B$2:$B$5,MATCH(LARGE($C$2:$C$5,2),$C$2:$C$5,0))' },
      { row: 13, col: 3, formula: '=INDEX($B$2:$B$5,MATCH(LARGE($C$2:$C$5,3),$C$2:$C$5,0))' },
    ],
    expectedResults: [
      { row: 7, col: 1, value: 'マウス' },
      { row: 8, col: 1, value: 'モニター' },
      { row: 9, col: 1, value: 'ノートPC' },
      { row: 7, col: 2, value: 3000 },
      { row: 8, col: 2, value: 45000 },
      { row: 9, col: 2, value: 120000 },
      { row: 7, col: 3, value: 50 },
      { row: 8, col: 3, value: 8 },
      { row: 9, col: 3, value: 5 },
      { row: 13, col: 1, value: 'ノートPC' },
      { row: 13, col: 2, value: 'モニター' },
      { row: 13, col: 3, value: 'キーボード' },
    ]
  }
];