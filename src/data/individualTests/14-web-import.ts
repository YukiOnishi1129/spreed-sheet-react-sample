import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const webImportTests: IndividualFunctionTest[] = [
  {
    name: 'ENCODEURL',
    category: '14. Web・インポート',
    description: 'URLエンコード',
    data: [
      ['テキスト', 'エンコード結果'],
      ['Hello World', '=ENCODEURL(A2)'],
      ['こんにちは', '=ENCODEURL(A3)']
    ],
    expectedValues: { 'B2': 'Hello%20World' }
  },
  {
    name: 'SPLIT',
    category: '14. Web・インポート',
    description: 'テキストを分割',
    data: [
      ['テキスト', '区切り文字', '分割結果'],
      ['apple,banana,orange', ',', '=SPLIT(A2,B2)'],
      ['one-two-three', '-', '=SPLIT(A3,B3)']
    ],
    },
  {
    name: 'JOIN',
    category: '14. Web・インポート',
    description: 'テキストを結合',
    data: [
      ['区切り文字', '値1', '値2', '値3', '結合結果'],
      [',', 'apple', 'banana', 'orange', '=JOIN(A2,B2:D2)'],
      ['-', 'one', 'two', 'three', '=JOIN(A3,B3:D3)']
    ],
    expectedValues: { 'E2': 'apple,banana,orange', 'E3': 'one-two-three' }
  },
  {
    name: 'QUERY',
    category: '14. Web・インポート',
    description: 'データのクエリ',
    data: [
      ['名前', '年齢', '部署', '', 'クエリ結果'],
      ['田中', 25, '営業', '', '=QUERY(A2:C4,"SELECT A, B WHERE B > 25")'],
      ['佐藤', 30, '技術', '', ''],
      ['鈴木', 28, '人事', '', '']
    ],
    expectedValues: { 'E2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'FLATTEN',
    category: '14. Web・インポート',
    description: '多次元配列を1次元に',
    data: [
      ['配列1', '', '配列2', '', 'フラット化'],
      [1, 2, 5, 6, '=FLATTEN(A2:D3)'],
      [3, 4, 7, 8, '']
    ],
    expectedValues: { 'E2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'ARRAYFORMULA',
    category: '14. Web・インポート',
    description: '配列数式',
    data: [
      ['値', '結果'],
      [1, '=ARRAYFORMULA(A2:A5*2)'],
      [2, ''],
      [3, ''],
      [4, '']
    ],
    expectedValues: { 'B2': 2, 'B3': 4, 'B4': 6, 'B5': 8 }
  },
  {
    name: 'REGEXEXTRACT',
    category: '14. Web・インポート',
    description: '正規表現で抽出',
    data: [
      ['テキスト', 'パターン', '抽出結果'],
      ['abc123def', '[0-9]+', '=REGEXEXTRACT(A2,B2)'],
      ['test@example.com', '[^@]+', '=REGEXEXTRACT(A3,B3)']
    ],
    expectedValues: { 'C2': '123', 'C3': 'test' }
  },
  {
    name: 'REGEXMATCH',
    category: '14. Web・インポート',
    description: '正規表現でマッチ',
    data: [
      ['テキスト', 'パターン', 'マッチ'],
      ['abc123', '[0-9]+', '=REGEXMATCH(A2,B2)'],
      ['abcdef', '[0-9]+', '=REGEXMATCH(A3,B3)']
    ],
    expectedValues: { 'C2': true, 'C3': false }
  },
  {
    name: 'REGEXREPLACE',
    category: '14. Web・インポート',
    description: '正規表現で置換',
    data: [
      ['テキスト', 'パターン', '置換文字', '結果'],
      ['abc123def', '[0-9]+', 'XXX', '=REGEXREPLACE(A2,B2,C2)'],
      ['hello world', ' ', '_', '=REGEXREPLACE(A3,B3,C3)']
    ],
    expectedValues: { 'D2': 'abcXXXdef', 'D3': 'hello_world' }
  },
  {
    name: 'SORTN',
    category: '14. Web・インポート',
    description: '上位N件をソート',
    data: [
      ['値', 'ソート結果'],
      [85, '=SORTN(A2:A6,3)'],
      [92, ''],
      [78, ''],
      [95, ''],
      [88, '']
    ],
    expectedValues: { 'C2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'WEBSERVICE',
    category: '14. Web・インポート',
    description: 'Webサービスからデータ取得',
    data: [
      ['URL', '取得結果'],
      ['https://api.example.com/data', '=WEBSERVICE(A2)']
    ],
    expectedValues: { 'B2': '#N/A - Web functions not available' }
  },
  {
    name: 'FILTERXML',
    category: '14. Web・インポート',
    description: 'XMLからデータ抽出',
    data: [
      ['XMLデータ', 'XPath', '抽出結果'],
      ['<root><item>Value1</item><item>Value2</item></root>', '//item[1]', '=FILTERXML(A2,B2)']
    ],
    expectedValues: { 'C2': 'Value1' }
  },
  {
    name: 'SPARKLINE',
    category: '14. Web・インポート',
    description: 'ミニグラフを作成',
    data: [
      ['データ', '', '', '', 'スパークライン'],
      [1, 3, 2, 5, '=SPARKLINE(A2:D2)'],
      [4, 2, 6, 3, '=SPARKLINE(A3:D3,{"charttype","column"})']
    ],
    expectedValues: { 'B2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'IMPORTDATA',
    category: '14. Web・インポート',
    description: 'URLからデータをインポート',
    data: [
      ['URL', 'インポート結果'],
      ['https://example.com/data.csv', '=IMPORTDATA(A2)']
    ],
    expectedValues: { 'B2': '#N/A - Import functions not available' }
  },
  {
    name: 'IMPORTFEED',
    category: '14. Web・インポート',
    description: 'RSSやAtomフィードをインポート',
    data: [
      ['フィードURL', 'クエリ', 'ヘッダー', 'アイテム数', 'インポート結果'],
      ['https://example.com/feed.rss', 'items title', true, 5, '=IMPORTFEED(A2,B2,C2,D2)']
    ],
    expectedValues: { 'E2': '#N/A - Import functions not available' }
  },
  {
    name: 'IMPORTHTML',
    category: '14. Web・インポート',
    description: 'HTMLテーブルやリストをインポート',
    data: [
      ['URL', 'クエリ', 'インデックス', 'インポート結果'],
      ['https://example.com/page.html', 'table', 1, '=IMPORTHTML(A2,B2,C2)']
    ],
    expectedValues: { 'D2': '#N/A - Import functions not available' }
  },
  {
    name: 'IMPORTXML',
    category: '14. Web・インポート',
    description: 'XMLデータをインポート',
    data: [
      ['URL', 'XPathクエリ', 'インポート結果'],
      ['https://example.com/data.xml', '//item/title', '=IMPORTXML(A2,B2)']
    ],
    expectedValues: { 'C2': '#N/A - Import functions not available' }
  },
  {
    name: 'IMPORTRANGE',
    category: '14. Web・インポート',
    description: '他のスプレッドシートから範囲をインポート',
    data: [
      ['スプレッドシートURL', '範囲', 'インポート結果'],
      ['https://docs.google.com/spreadsheets/d/abc123', 'Sheet1!A1:C10', '=IMPORTRANGE(A2,B2)']
    ],
    expectedValues: { 'C2': '#N/A - Import functions not available' }
  },
  {
    name: 'IMAGE',
    category: '14. Web・インポート',
    description: '画像を挿入',
    data: [
      ['画像URL', 'モード', '高さ', '幅', '画像'],
      ['https://example.com/image.png', 1, 100, 100, '=IMAGE(A2,B2,C2,D2)']
    ],
    expectedValues: { 'B2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'GOOGLEFINANCE',
    category: '14. Web・インポート',
    description: '金融情報を取得',
    data: [
      ['銘柄', '属性', '開始日', '終了日', '間隔', '結果'],
      ['GOOG', 'price', '', '', '', '=GOOGLEFINANCE(A2,B2)'],
      ['AAPL', 'volume', '2024/1/1', '2024/1/31', 'DAILY', '=GOOGLEFINANCE(A3,B3,C3,D3,E3)']
    ],
    expectedValues: { 'C2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'GOOGLETRANSLATE',
    category: '14. Web・インポート',
    description: 'テキストを翻訳',
    data: [
      ['テキスト', '元言語', '翻訳先言語', '翻訳結果'],
      ['Hello', 'en', 'ja', '=GOOGLETRANSLATE(A2,B2,C2)'],
      ['こんにちは', 'ja', 'en', '=GOOGLETRANSLATE(A3,B3,C3)']
    ],
    expectedValues: { 'C2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'DETECTLANGUAGE',
    category: '14. Web・インポート',
    description: '言語を検出',
    data: [
      ['テキスト', '検出言語'],
      ['Hello World', '=DETECTLANGUAGE(A2)'],
      ['こんにちは世界', '=DETECTLANGUAGE(A3)'],
      ['Bonjour le monde', '=DETECTLANGUAGE(A4)']
    ],
    expectedValues: { 'B2': 'en', 'B3': 'ja', 'B4': 'fr' }
  },
  {
    name: 'TO_DATE',
    category: '14. Web・インポート',
    description: '値を日付に変換',
    data: [
      ['値', '日付'],
      [44926, '=TO_DATE(A2)'],
      ['2023/1/1', '=TO_DATE(A3)']
    ],
    expectedValues: { 'B2': '#N/A - Google Sheets specific functions not supported' }
  },
  {
    name: 'TO_PERCENT',
    category: '14. Web・インポート',
    description: '値をパーセントに変換',
    data: [
      ['値', 'パーセント'],
      [0.25, '=TO_PERCENT(A2)'],
      [0.5, '=TO_PERCENT(A3)']
    ],
    expectedValues: { 'B2': '25%', 'B3': '50%' }
  },
  {
    name: 'TO_DOLLARS',
    category: '14. Web・インポート',
    description: '値をドル表記に変換',
    data: [
      ['値', 'ドル表記'],
      [1234.56, '=TO_DOLLARS(A2)'],
      [9876.54, '=TO_DOLLARS(A3)']
    ],
    expectedValues: { 'B2': '$1,234.56', 'B3': '$9,876.54' }
  },
  {
    name: 'TO_TEXT',
    category: '14. Web・インポート',
    description: '値をテキストに変換',
    data: [
      ['値', 'テキスト'],
      [123, '=TO_TEXT(A2)'],
      ['=TRUE()', '=TO_TEXT(A3)']
    ],
    expectedValues: { 'B2': '123', 'B3': 'TRUE' }
  }
];
