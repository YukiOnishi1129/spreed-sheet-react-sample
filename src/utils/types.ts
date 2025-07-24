// 型定義
export interface FunctionParam {
  name: string;
  desc: string;
  optional?: boolean;
}

export interface FunctionDefinition {
  name: string;
  syntax: string;
  params: FunctionParam[];
  description: string;
  category: string;
  examples?: string[];
  colorScheme?: {
    bg: string;
    fc: string;
  };
}