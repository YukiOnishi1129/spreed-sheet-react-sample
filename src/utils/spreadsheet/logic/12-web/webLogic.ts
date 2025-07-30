// Web関数の実装

import type { CustomFormula, FormulaContext, FormulaResult } from '../shared/types';
import { FormulaError } from '../shared/types';
import { getCellValue } from '../shared/utils';

// WEBSERVICE関数の実装（Webサービスからデータ取得）
export const WEBSERVICE: CustomFormula = {
  name: 'WEBSERVICE',
  pattern: /WEBSERVICE\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, urlRef] = matches;
    
    try {
      let url = getCellValue(urlRef.trim(), context)?.toString() ?? urlRef.trim();
      
      // 引用符を除去
      if (url.startsWith('"') && url.endsWith('"')) {
        url = url.slice(1, -1);
      }
      
      // URLの検証
      try {
        new URL(url);
      } catch {
        return FormulaError.VALUE;
      }
      
      // 実際のWeb APIコールはセキュリティ上の理由でブラウザ環境では制限される
      // ここでは警告メッセージを返す
      return '#N/A - Web service calls require external access';
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// FILTERXML関数の実装（XMLからデータ抽出）
export const FILTERXML: CustomFormula = {
  name: 'FILTERXML',
  pattern: /FILTERXML\(([^,]+),\s*([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, xmlRef, xpathRef] = matches;
    
    try {
      let xml = getCellValue(xmlRef.trim(), context)?.toString() ?? xmlRef.trim();
      let xpath = getCellValue(xpathRef.trim(), context)?.toString() ?? xpathRef.trim();
      
      // 引用符を除去
      if (xml.startsWith('"') && xml.endsWith('"')) {
        xml = xml.slice(1, -1);
      }
      if (xpath.startsWith('"') && xpath.endsWith('"')) {
        xpath = xpath.slice(1, -1);
      }
      
      // 簡易的なXML解析（実装例）
      // 実際のXPath評価は複雑なので、基本的なケースのみ対応
      if (!xml.includes('<') || !xml.includes('>')) {
        return FormulaError.VALUE;
      }
      
      // 簡単なタグ抽出の例
      const tagMatch = xpath.match(/\/\/(\w+)(?:\[(\d+)\])?$/);
      if (tagMatch) {
        const [, tagName, index] = tagMatch;
        const regex = new RegExp(`<${tagName}[^>]*>([^<]*)</${tagName}>`, 'gi');
        const matches = Array.from(xml.matchAll(regex));
        
        if (matches.length === 0) {
          return FormulaError.NA;
        }
        
        if (index) {
          const idx = parseInt(index) - 1;
          return idx < matches.length ? matches[idx][1] : FormulaError.NA;
        }
        
        // 最初の一致を返す
        return matches[0][1];
      }
      
      return FormulaError.VALUE;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ENCODEURL関数の実装（URLエンコード）
export const ENCODEURL: CustomFormula = {
  name: 'ENCODEURL',
  pattern: /ENCODEURL\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, textRef] = matches;
    
    try {
      let text = getCellValue(textRef.trim(), context)?.toString() ?? textRef.trim();
      
      // 引用符を除去
      if (text.startsWith('"') && text.endsWith('"')) {
        text = text.slice(1, -1);
      }
      
      // URLエンコード
      return encodeURIComponent(text);
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// HYPERLINK関数の実装（ハイパーリンクの作成）
export const HYPERLINK: CustomFormula = {
  name: 'HYPERLINK',
  pattern: /HYPERLINK\(([^,)]+)(?:,\s*([^)]+))?\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, linkLocationRef, linkLabelRef] = matches;
    
    try {
      let linkLocation = getCellValue(linkLocationRef.trim(), context)?.toString() ?? linkLocationRef.trim();
      let linkLabel = linkLabelRef ? 
        (getCellValue(linkLabelRef.trim(), context)?.toString() ?? linkLabelRef.trim()) :
        null;
      
      // 引用符を除去
      if (linkLocation.startsWith('"') && linkLocation.endsWith('"')) {
        linkLocation = linkLocation.slice(1, -1);
      }
      if (linkLabel && linkLabel.startsWith('"') && linkLabel.endsWith('"')) {
        linkLabel = linkLabel.slice(1, -1);
      }
      
      // linkLabelが指定されていない場合は、linkLocationを表示
      if (linkLabel === null || linkLabel === undefined) {
        return linkLocation;
      }
      
      // linkLabelが指定されている場合は、それを返す
      return linkLabel;
    } catch {
      return FormulaError.VALUE;
    }
  }
};

// ISURL関数の実装（URLかどうかの検証）
export const ISURL: CustomFormula = {
  name: 'ISURL',
  pattern: /ISURL\(([^)]+)\)/i,
  calculate: (matches: RegExpMatchArray, context: FormulaContext): FormulaResult => {
    const [, valueRef] = matches;
    
    try {
      const value = getCellValue(valueRef.trim(), context);
      
      // 数値の場合はfalseを返す
      if (typeof value === 'number') {
        return false;
      }
      
      let text = value?.toString() ?? valueRef.trim();
      
      // 引用符を除去
      if (text.startsWith('"') && text.endsWith('"')) {
        text = text.slice(1, -1);
      }
      
      // 空文字列の場合はfalse
      if (text === '') {
        return false;
      }
      
      // URLパターンの検証
      try {
        const url = new URL(text);
        // 有効なプロトコルかどうかチェック
        const validProtocols = ['http:', 'https:', 'ftp:', 'ftps:', 'data:'];
        return validProtocols.includes(url.protocol);
      } catch {
        return false;
      }
    } catch {
      return false;
    }
  }
};