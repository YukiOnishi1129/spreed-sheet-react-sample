import { Routes, Route, Link } from 'react-router-dom';
import ChatGPTSpreadsheet from './components/ChatGPTSpreadsheet';
import DemoSpreadsheet from './components/DemoSpreadsheet';

function HomePage() {
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-100 relative overflow-hidden">
      {/* Background decorative elements */}
      <div className="absolute inset-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-40 -right-40 w-80 h-80 bg-blue-400/10 rounded-full blur-3xl"></div>
        <div className="absolute top-1/2 -left-32 w-64 h-64 bg-indigo-400/10 rounded-full blur-3xl"></div>
        <div className="absolute bottom-20 right-20 w-48 h-48 bg-purple-400/10 rounded-full blur-2xl"></div>
      </div>
      
      <div className="relative z-10">
        <div className="max-w-7xl mx-auto px-6 py-12 lg:py-16">
          {/* Header Section */}
          <div className="text-center mb-16">
            <div className="inline-flex items-center px-6 py-2 bg-white/60 backdrop-blur-sm rounded-full border border-white/20 shadow-lg mb-6">
              <span className="text-sm font-medium text-indigo-600 bg-gradient-to-r from-indigo-600 to-purple-600 bg-clip-text text-transparent">
                ✨ Powered by ChatGPT
              </span>
            </div>
            <h1 className="text-5xl lg:text-6xl font-extrabold bg-gradient-to-r from-gray-900 via-blue-900 to-indigo-900 bg-clip-text text-transparent mb-6 leading-tight">
              Excel関数学習
              <br />
              <span className="text-4xl lg:text-5xl">デモンストレーション</span>
            </h1>
            <p className="text-xl text-gray-600 max-w-2xl mx-auto leading-relaxed">
              自然言語でExcel関数を検索・学習できる次世代ツール
            </p>
            
            {/* Navigation */}
            <div className="mt-8 flex gap-4 justify-center">
              <Link 
                to="/demo" 
                className="px-6 py-3 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition-colors shadow-lg"
              >
                関数デモモード
              </Link>
            </div>
          </div>
          
          <ChatGPTSpreadsheet />
        </div>
      </div>
    </div>
  );
}

function App() {
  return (
    <Routes>
      <Route path="/" element={<HomePage />} />
      <Route path="/demo" element={<DemoSpreadsheet />} />
    </Routes>
  );
}

export default App
