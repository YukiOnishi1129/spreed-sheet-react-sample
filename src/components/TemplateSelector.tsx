import React, { useState } from 'react';
import { TEMPLATE_CATEGORIES, type FunctionTemplate } from '../types/templates';

interface TemplateSelectorProps {
  onTemplateSelect: (template: FunctionTemplate) => void;
  onClose: () => void;
}

const TemplateSelector: React.FC<TemplateSelectorProps> = ({ onTemplateSelect, onClose }) => {
  const [selectedCategory, setSelectedCategory] = useState<string | null>(null);

  const handleTemplateClick = (template: FunctionTemplate) => {
    onTemplateSelect(template);
    onClose();
  };

  const getDifficultyColor = (difficulty: string) => {
    switch (difficulty) {
      case 'beginner': return '#4CAF50';
      case 'intermediate': return '#FF9800';
      case 'advanced': return '#F44336';
      default: return '#757575';
    }
  };

  const getDifficultyLabel = (difficulty: string) => {
    switch (difficulty) {
      case 'beginner': return '初級';
      case 'intermediate': return '中級';
      case 'advanced': return '上級';
      default: return '';
    }
  };

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(0,0,0,0.5)',
      display: 'flex',
      justifyContent: 'center',
      alignItems: 'center',
      zIndex: 1000
    }}>
      <div style={{
        backgroundColor: 'white',
        borderRadius: '12px',
        padding: '24px',
        maxWidth: '800px',
        maxHeight: '80vh',
        width: '90%',
        overflowY: 'auto',
        boxShadow: '0 10px 30px rgba(0,0,0,0.3)'
      }}>
        <div style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          marginBottom: '24px',
          borderBottom: '2px solid #f0f0f0',
          paddingBottom: '16px'
        }}>
          <h2 style={{ margin: 0, color: '#333', fontSize: '24px' }}>
            📚 関数テンプレート集
          </h2>
          <button
            onClick={onClose}
            style={{
              background: 'none',
              border: 'none',
              fontSize: '24px',
              cursor: 'pointer',
              color: '#999',
              padding: '4px'
            }}
          >
            ✕
          </button>
        </div>

        {!selectedCategory ? (
          // カテゴリ選択画面
          <div>
            <p style={{ 
              color: '#666', 
              marginBottom: '20px',
              fontSize: '16px'
            }}>
              用途に合わせてカテゴリをお選びください
            </p>
            <div style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))',
              gap: '16px'
            }}>
              {TEMPLATE_CATEGORIES.map((category) => (
                <div
                  key={category.id}
                  onClick={() => setSelectedCategory(category.id)}
                  style={{
                    border: '2px solid #e0e0e0',
                    borderRadius: '8px',
                    padding: '20px',
                    textAlign: 'center',
                    cursor: 'pointer',
                    transition: 'all 0.2s ease',
                    backgroundColor: '#fafafa'
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.borderColor = '#007bff';
                    e.currentTarget.style.backgroundColor = '#f0f8ff';
                    e.currentTarget.style.transform = 'translateY(-2px)';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.borderColor = '#e0e0e0';
                    e.currentTarget.style.backgroundColor = '#fafafa';
                    e.currentTarget.style.transform = 'translateY(0)';
                  }}
                >
                  <div style={{ fontSize: '48px', marginBottom: '12px' }}>
                    {category.icon}
                  </div>
                  <h3 style={{ 
                    margin: '0 0 8px 0', 
                    color: '#333',
                    fontSize: '18px'
                  }}>
                    {category.name}
                  </h3>
                  <p style={{ 
                    margin: 0, 
                    color: '#666',
                    fontSize: '14px',
                    lineHeight: '1.4'
                  }}>
                    {category.description}
                  </p>
                  <div style={{
                    marginTop: '12px',
                    fontSize: '12px',
                    color: '#999'
                  }}>
                    {category.templates.length}個のテンプレート
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          // テンプレート詳細画面
          <div>
            <button
              onClick={() => setSelectedCategory(null)}
              style={{
                background: 'none',
                border: '1px solid #ddd',
                borderRadius: '4px',
                padding: '8px 16px',
                cursor: 'pointer',
                marginBottom: '20px',
                color: '#666'
              }}
            >
              ← カテゴリ一覧に戻る
            </button>
            
            {TEMPLATE_CATEGORIES.find(cat => cat.id === selectedCategory)?.templates.map((template) => (
              <div
                key={template.id}
                style={{
                  border: '1px solid #e0e0e0',
                  borderRadius: '8px',
                  padding: '20px',
                  marginBottom: '16px',
                  backgroundColor: '#fff',
                  cursor: 'pointer',
                  transition: 'all 0.2s ease'
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.borderColor = '#007bff';
                  e.currentTarget.style.backgroundColor = '#f8f9ff';
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.borderColor = '#e0e0e0';
                  e.currentTarget.style.backgroundColor = '#fff';
                }}
                onClick={() => handleTemplateClick(template)}
              >
                <div style={{ display: 'flex', alignItems: 'flex-start', gap: '16px' }}>
                  <div style={{ fontSize: '32px', flexShrink: 0 }}>
                    {template.icon}
                  </div>
                  <div style={{ flex: 1 }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '12px', marginBottom: '8px' }}>
                      <h3 style={{ margin: 0, color: '#333', fontSize: '20px' }}>
                        {template.title}
                      </h3>
                      <span style={{
                        backgroundColor: getDifficultyColor(template.difficulty),
                        color: 'white',
                        padding: '2px 8px',
                        borderRadius: '12px',
                        fontSize: '12px',
                        fontWeight: 'bold'
                      }}>
                        {getDifficultyLabel(template.difficulty)}
                      </span>
                    </div>
                    <p style={{ 
                      margin: '0 0 12px 0', 
                      color: '#666',
                      fontSize: '14px',
                      lineHeight: '1.5'
                    }}>
                      {template.description}
                    </p>
                    <div style={{ display: 'flex', gap: '8px', marginBottom: '12px' }}>
                      {template.functions.map((func) => (
                        <span
                          key={func}
                          style={{
                            backgroundColor: '#e3f2fd',
                            color: '#1976d2',
                            padding: '4px 8px',
                            borderRadius: '4px',
                            fontSize: '12px',
                            fontWeight: 'bold'
                          }}
                        >
                          {func}
                        </span>
                      ))}
                    </div>
                    <div style={{ display: 'flex', gap: '6px' }}>
                      {template.tags.map((tag) => (
                        <span
                          key={tag}
                          style={{
                            backgroundColor: '#f5f5f5',
                            color: '#666',
                            padding: '2px 6px',
                            borderRadius: '3px',
                            fontSize: '11px'
                          }}
                        >
                          #{tag}
                        </span>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

export default TemplateSelector;