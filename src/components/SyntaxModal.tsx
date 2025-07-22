import React from 'react';

interface SyntaxModalProps {
  isOpen: boolean;
  onClose: () => void;
  functionName: string;
  syntax: string;
  syntaxDetail?: string;
  description: string;
  examples?: string[];
}

// 個別の関数構文を解析する関数
const parseFunctionSyntax = (syntax: string, syntaxDetail?: string) => {
  // 複数関数の場合（+ で区切られている場合）
  if (syntax.includes(' + ')) {
    const functions = syntax.split(' + ').map(s => s.trim());
    const details = syntaxDetail?.split(' + ') ?? [];
    
    return functions.map((func, index) => ({
      syntax: func,
      detail: details[index]?.trim() ?? ''
    }));
  }
  
  // 単一関数の場合
  return [{
    syntax: syntax,
    detail: syntaxDetail ?? ''
  }];
};

const SyntaxModal: React.FC<SyntaxModalProps> = ({
  isOpen,
  onClose,
  functionName,
  syntax,
  syntaxDetail,
  description,
  examples
}) => {
  if (!isOpen) return null;

  const parsedFunctions = parseFunctionSyntax(syntax, syntaxDetail);
  const isMultipleFunctions = parsedFunctions.length > 1;

  return (
    <div className="fixed inset-0 bg-black/60 flex justify-center items-center z-[10000] backdrop-blur-sm p-4">
      <div className="bg-white/95 backdrop-blur-sm rounded-3xl shadow-2xl border border-white/50 max-w-3xl w-full max-h-[90vh] sm:max-h-[80vh] overflow-y-auto">
        {/* ヘッダー */}
        <div className="sticky top-0 bg-gradient-to-r from-blue-600 to-indigo-600 text-white p-4 sm:p-6 rounded-t-3xl">
          <div className="flex items-center justify-between">
            <h2 className="text-lg sm:text-2xl font-bold flex items-center gap-2 sm:gap-3">
              📚 {functionName} 関数の詳細
            </h2>
            <button
              onClick={onClose}
              className="w-8 h-8 rounded-full bg-white/20 hover:bg-white/30 transition-colors flex items-center justify-center text-xl"
              aria-label="閉じる"
            >
              ×
            </button>
          </div>
        </div>

        {/* コンテンツ */}
        <div className="p-4 sm:p-6 lg:p-8 space-y-6 sm:space-y-8">
          {/* 基本構文 - 複数関数の場合は個別に表示 */}
          <div className="bg-gradient-to-br from-pink-50 to-purple-50 rounded-2xl p-4 sm:p-6 border border-pink-200">
            <h3 className="text-lg sm:text-xl font-bold text-pink-700 mb-3 sm:mb-4 flex items-center gap-2">
              🔧 {isMultipleFunctions ? '使用する関数' : '基本構文'}
            </h3>
            {isMultipleFunctions ? (
              <div className="space-y-4">
                {parsedFunctions.map((func, index) => (
                  <div key={index} className="bg-white/70 border border-pink-300 rounded-xl p-4">
                    <code className="block text-pink-600 font-mono text-lg font-semibold mb-2">
                      {func.syntax}
                    </code>
                    {func.detail && (
                      <div className="text-pink-800 text-sm leading-relaxed font-mono whitespace-pre-line">
                        {func.detail}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            ) : (
              <div>
                <code className="block bg-white/70 border border-pink-300 px-4 py-3 rounded-xl text-pink-600 font-mono text-lg font-semibold mb-2">
                  {syntax}
                </code>
                {syntaxDetail && (
                  <div className="text-pink-800 text-sm leading-relaxed font-mono bg-white/70 border border-pink-300 px-4 py-3 rounded-xl mt-2 whitespace-pre-line">
                    {syntaxDetail}
                  </div>
                )}
              </div>
            )}
          </div>

          {/* 関数の説明 */}
          <div className="bg-gradient-to-br from-green-50 to-emerald-50 rounded-2xl p-6 border border-green-200">
            <h3 className="text-xl font-bold text-green-700 mb-4 flex items-center gap-2">
              💡 機能説明
            </h3>
            <p className="text-green-800 leading-relaxed text-lg">
              {description}
            </p>
          </div>

          {/* 使用例 */}
          {examples && examples.length > 0 && (
            <div className="bg-gradient-to-br from-orange-50 to-yellow-50 rounded-2xl p-6 border border-orange-200">
              <h3 className="text-xl font-bold text-orange-700 mb-4 flex items-center gap-2">
                🎯 使用例
              </h3>
              <div className="space-y-3">
                {examples.map((example, index) => (
                  <code 
                    key={index} 
                    className="block bg-white/70 border border-orange-300 px-4 py-3 rounded-xl text-orange-600 font-mono hover:shadow-md transition-shadow cursor-pointer"
                    onClick={() => navigator.clipboard.writeText(example)}
                    title="クリックでコピー"
                  >
                    {example}
                  </code>
                ))}
              </div>
              <p className="text-sm text-orange-600 mt-3 text-center">
                💡 使用例をクリックするとクリップボードにコピーされます
              </p>
            </div>
          )}
        </div>

        {/* フッター */}
        <div className="sticky bottom-0 bg-gradient-to-r from-gray-50 to-gray-100 p-6 rounded-b-3xl border-t border-gray-200">
          <div className="flex justify-end">
            <button
              onClick={onClose}
              className="px-8 py-3 bg-gradient-to-r from-gray-600 to-gray-700 text-white rounded-xl font-semibold hover:from-gray-700 hover:to-gray-800 transition-all duration-200 shadow-lg hover:shadow-xl transform hover:-translate-y-0.5"
            >
              閉じる
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default SyntaxModal;