import { useState, useEffect, useMemo } from 'react';
import type { IndividualFunctionTest } from '../data/individualFunctionTests';
import {
  mathFunctionTests,
  statisticalFunctionTests,
  textFunctionTests,
  dateFunctionTests,
  logicalFunctionTests,
  lookupFunctionTests,
  financialFunctionTests,
  allIndividualFunctionTests
} from '../data/individualFunctionTests';

interface FunctionSelectorModalProps {
  isOpen: boolean;
  onClose: () => void;
  onSelect: (func: IndividualFunctionTest) => void;
  selectedFunction: IndividualFunctionTest | null;
}

export function FunctionSelectorModal({ 
  isOpen, 
  onClose, 
  onSelect, 
  selectedFunction 
}: FunctionSelectorModalProps) {
  const [selectedCategory, setSelectedCategory] = useState<string>('all');
  const [searchQuery, setSearchQuery] = useState<string>('');

  // カテゴリー別の関数マップ
  const functionCategories = useMemo(() => ({
    '数学・三角関数': mathFunctionTests,
    '統計関数': statisticalFunctionTests,
    'テキスト関数': textFunctionTests,
    '日付・時刻関数': dateFunctionTests,
    '論理関数': logicalFunctionTests,
    '検索・参照関数': lookupFunctionTests,
    '財務関数': financialFunctionTests
  }), []);

  // フィルター済み関数リスト
  const filteredFunctions = useMemo(() => {
    let functions: IndividualFunctionTest[] = [];
    
    if (selectedCategory === 'all') {
      functions = allIndividualFunctionTests;
    } else if (selectedCategory in functionCategories) {
      functions = functionCategories[selectedCategory as keyof typeof functionCategories];
    }
    
    if (searchQuery) {
      const query = searchQuery.toLowerCase();
      functions = functions.filter(func => 
        func.name.toLowerCase().includes(query) || 
        func.description.toLowerCase().includes(query)
      );
    }
    
    return functions;
  }, [selectedCategory, functionCategories, searchQuery]);

  // ESCキーでモーダルを閉じる
  useEffect(() => {
    const handleEsc = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isOpen) {
        onClose();
      }
    };
    window.addEventListener('keydown', handleEsc);
    return () => window.removeEventListener('keydown', handleEsc);
  }, [isOpen, onClose]);

  // モーダルが開いたときに検索をリセット
  useEffect(() => {
    if (isOpen) {
      setSearchQuery('');
      setSelectedCategory('all');
    }
  }, [isOpen]);

  if (!isOpen) return null;

  return (
    <>
      {/* オーバーレイ */}
      <div 
        className="fixed inset-0 bg-black bg-opacity-50 z-40 transition-opacity"
        onClick={onClose}
      />
      
      {/* モーダル本体 */}
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
        <div 
          className="bg-white rounded-xl shadow-2xl max-w-4xl w-full max-h-[85vh] flex flex-col"
          onClick={(e) => e.stopPropagation()}
        >
          {/* ヘッダー */}
          <div className="p-6 border-b">
            <div className="flex items-center justify-between mb-4">
              <h2 className="text-xl font-semibold text-gray-900">関数を選択</h2>
              <button
                onClick={onClose}
                className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
              >
                <svg className="w-5 h-5 text-gray-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>
            
            {/* 検索バー */}
            <div className="relative">
              <input
                type="text"
                placeholder="関数名または説明で検索..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-10 pr-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                autoFocus
              />
              <svg className="w-5 h-5 text-gray-400 absolute left-3 top-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
              </svg>
            </div>
          </div>
          
          {/* カテゴリータブ */}
          <div className="px-6 pt-4 border-b overflow-x-auto">
            <div className="flex gap-2 pb-3">
              <button
                onClick={() => setSelectedCategory('all')}
                className={`px-4 py-2 text-sm font-medium rounded-lg whitespace-nowrap transition-colors ${
                  selectedCategory === 'all'
                    ? 'bg-blue-500 text-white'
                    : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                }`}
              >
                すべて
              </button>
              {Object.keys(functionCategories).map(category => (
                <button
                  key={category}
                  onClick={() => setSelectedCategory(category)}
                  className={`px-4 py-2 text-sm font-medium rounded-lg whitespace-nowrap transition-colors ${
                    selectedCategory === category
                      ? 'bg-blue-500 text-white'
                      : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                  }`}
                >
                  {category}
                </button>
              ))}
            </div>
          </div>
          
          {/* 関数グリッド */}
          <div className="flex-1 overflow-y-auto p-6">
            {filteredFunctions.length > 0 ? (
              <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                {filteredFunctions.map(func => (
                  <button
                    key={func.name}
                    onClick={() => {
                      onSelect(func);
                      onClose();
                    }}
                    className={`p-4 text-left rounded-lg border-2 transition-all hover:shadow-lg transform hover:scale-105 ${
                      selectedFunction?.name === func.name
                        ? 'border-blue-500 bg-blue-50 shadow-md'
                        : 'border-gray-200 hover:border-gray-300 bg-white'
                    }`}
                  >
                    <div className="font-medium text-gray-900 mb-1">{func.name}</div>
                    <div className="text-xs text-gray-600 line-clamp-2">{func.description}</div>
                  </button>
                ))}
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center h-64 text-gray-500">
                <svg className="w-16 h-16 mb-4 text-gray-300" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9.172 16.172a4 4 0 015.656 0M9 10h.01M15 10h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
                <p className="text-lg">
                  {searchQuery 
                    ? `"${searchQuery}" に一致する関数が見つかりません`
                    : '関数を選択してください'
                  }
                </p>
              </div>
            )}
          </div>
          
          {/* フッター */}
          <div className="p-4 border-t bg-gray-50 rounded-b-xl">
            <div className="flex justify-between items-center">
              <div className="text-sm text-gray-600">
                {filteredFunctions.length} 個の関数
              </div>
              <button
                onClick={onClose}
                className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
              >
                キャンセル
              </button>
            </div>
          </div>
        </div>
      </div>
    </>
  );
}