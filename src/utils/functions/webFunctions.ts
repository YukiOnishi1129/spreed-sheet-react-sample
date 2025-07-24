// Web関数の定義
import { type FunctionDefinition } from '../types';
import { COLOR_SCHEMES } from '../colorSchemes';

export const WEB_FUNCTIONS: Record<string, FunctionDefinition> = {
  WEBSERVICE: {
    name: 'WEBSERVICE',
    syntax: 'WEBSERVICE(url)',
    params: [
      { name: 'url', desc: 'WebサービスのURL' }
    ],
    description: 'WebサービスからデータをHTTPで取得します',
    category: 'web',
    examples: ['=WEBSERVICE("http://api.example.com/data")', '=WEBSERVICE("https://httpbin.org/json")'],
    colorScheme: COLOR_SCHEMES.web
  },

  FILTERXML: {
    name: 'FILTERXML',
    syntax: 'FILTERXML(xml, xpath)',
    params: [
      { name: 'xml', desc: 'XMLデータ' },
      { name: 'xpath', desc: 'XPath式' }
    ],
    description: 'XMLデータから指定したXPathに一致するデータを抽出します',
    category: 'web',
    examples: ['=FILTERXML(A1,"//book/title")', '=FILTERXML(WEBSERVICE("http://api.example.com/xml"),"//item/@id")'],
    colorScheme: COLOR_SCHEMES.web
  },

  ENCODEURL: {
    name: 'ENCODEURL',
    syntax: 'ENCODEURL(text)',
    params: [
      { name: 'text', desc: 'エンコードする文字列' }
    ],
    description: '文字列をURLエンコードします',
    category: 'web',
    examples: ['=ENCODEURL("Hello World")', '=ENCODEURL("データ分析")'],
    colorScheme: COLOR_SCHEMES.web
  },

  // Webデータ取得関数（主にGoogle Sheets由来）
  IMPORTDATA: {
    name: 'IMPORTDATA',
    syntax: 'IMPORTDATA(url)',
    params: [
      { name: 'url', desc: 'CSVまたはTSVファイルのURL' }
    ],
    description: 'WebからCSVまたはTSVデータをインポートします',
    category: 'web',
    examples: ['=IMPORTDATA("https://example.com/data.csv")', '=IMPORTDATA("https://docs.google.com/spreadsheets/d/abc123/export?format=csv")'],
    colorScheme: COLOR_SCHEMES.web
  },

  IMPORTHTML: {
    name: 'IMPORTHTML',
    syntax: 'IMPORTHTML(url, query, index)',
    params: [
      { name: 'url', desc: 'WebページのURL' },
      { name: 'query', desc: '"table"または"list"' },
      { name: 'index', desc: 'テーブルまたはリストのインデックス' }
    ],
    description: 'Webページからテーブルまたはリストをインポートします',
    category: 'web',
    examples: ['=IMPORTHTML("https://example.com","table",1)', '=IMPORTHTML("https://ja.wikipedia.org/wiki/日本","table",2)'],
    colorScheme: COLOR_SCHEMES.web
  },

  IMPORTXML: {
    name: 'IMPORTXML',
    syntax: 'IMPORTXML(url, xpath_query)',
    params: [
      { name: 'url', desc: 'XMLまたはHTMLページのURL' },
      { name: 'xpath_query', desc: 'XPath式' }
    ],
    description: 'WebページからXPathクエリに一致するデータをインポートします',
    category: 'web',
    examples: ['=IMPORTXML("https://example.com","//div[@class=\'price\']")', '=IMPORTXML("https://www.example.com/rss.xml","//item/title")'],
    colorScheme: COLOR_SCHEMES.web
  },

  IMPORTFEED: {
    name: 'IMPORTFEED',
    syntax: 'IMPORTFEED(url, [query], [headers], [num_items])',
    params: [
      { name: 'url', desc: 'RSSまたはAtomフィードのURL' },
      { name: 'query', desc: 'クエリタイプ', optional: true },
      { name: 'headers', desc: 'ヘッダーを含むかどうか', optional: true },
      { name: 'num_items', desc: '取得するアイテム数', optional: true }
    ],
    description: 'RSSまたはAtomフィードからデータをインポートします',
    category: 'web',
    examples: ['=IMPORTFEED("https://news.yahoo.co.jp/pickup/rss.xml")', '=IMPORTFEED("https://example.com/feed.xml","items",TRUE,10)'],
    colorScheme: COLOR_SCHEMES.web
  },

  IMPORTRANGE: {
    name: 'IMPORTRANGE',
    syntax: 'IMPORTRANGE(spreadsheet_url, range_string)',
    params: [
      { name: 'spreadsheet_url', desc: '他のスプレッドシートのURL' },
      { name: 'range_string', desc: '範囲指定文字列' }
    ],
    description: '他のスプレッドシートから指定範囲のデータをインポートします',
    category: 'web',
    examples: ['=IMPORTRANGE("https://docs.google.com/spreadsheets/d/abc123","Sheet1!A1:C10")', '=IMPORTRANGE("spreadsheet_key","Data!B2:D20")'],
    colorScheme: COLOR_SCHEMES.web
  },

  IMAGE: {
    name: 'IMAGE',
    syntax: 'IMAGE(url, [mode], [height], [width])',
    params: [
      { name: 'url', desc: '画像のURL' },
      { name: 'mode', desc: '表示モード（1-4）', optional: true },
      { name: 'height', desc: '高さ（ピクセル）', optional: true },
      { name: 'width', desc: '幅（ピクセル）', optional: true }
    ],
    description: 'セルに画像を挿入します',
    category: 'web',
    examples: ['=IMAGE("https://www.google.com/images/logo.png")', '=IMAGE("https://example.com/image.jpg",4,100,200)'],
    colorScheme: COLOR_SCHEMES.web
  }
};