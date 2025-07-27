import type { IndividualFunctionTest } from '../../types/spreadsheet';

export const basicOperatorsTests: IndividualFunctionTest[] = [
  {
    name: 'ADD_OPERATOR',
    category: '00. 基本演算子',
    description: '加算演算子（+）',
    data: [
      ['値1', '値2', '結果'],
      [10, 20, '=A2+B2'],
      [5.5, 4.5, '=A3+B3'],
      [-10, 5, '=A4+B4']
    ],
    expectedValues: { 'C2': 30, 'C3': 10, 'C4': -5 }
  },
  {
    name: 'SUBTRACT_OPERATOR',
    category: '00. 基本演算子',
    description: '減算演算子（-）',
    data: [
      ['値1', '値2', '結果'],
      [30, 10, '=A2-B2'],
      [5.5, 4.5, '=A3-B3'],
      [-10, 5, '=A4-B4']
    ],
    expectedValues: { 'C2': 20, 'C3': 1, 'C4': -15 }
  },
  {
    name: 'MULTIPLY_OPERATOR',
    category: '00. 基本演算子',
    description: '乗算演算子（*）',
    data: [
      ['値1', '値2', '結果'],
      [10, 3, '=A2*B2'],
      [2.5, 4, '=A3*B3'],
      [-5, 3, '=A4*B4']
    ],
    expectedValues: { 'C2': 30, 'C3': 10, 'C4': -15 }
  },
  {
    name: 'DIVIDE_OPERATOR',
    category: '00. 基本演算子',
    description: '除算演算子（/）',
    data: [
      ['値1', '値2', '結果'],
      [20, 4, '=A2/B2'],
      [15, 3, '=A3/B3'],
      [-10, 2, '=A4/B4']
    ],
    expectedValues: { 'C2': 5, 'C3': 5, 'C4': -5 }
  },
  {
    name: 'GREATER_THAN',
    category: '00. 基本演算子',
    description: '大なり演算子（>）',
    data: [
      ['値1', '値2', '結果'],
      [10, 5, '=A2>B2'],
      [5, 10, '=A3>B3'],
      [5, 5, '=A4>B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': false }
  },
  {
    name: 'LESS_THAN',
    category: '00. 基本演算子',
    description: '小なり演算子（<）',
    data: [
      ['値1', '値2', '結果'],
      [5, 10, '=A2<B2'],
      [10, 5, '=A3<B3'],
      [5, 5, '=A4<B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': false }
  },
  {
    name: 'GREATER_THAN_OR_EQUAL',
    category: '00. 基本演算子',
    description: '以上演算子（>=）',
    data: [
      ['値1', '値2', '結果'],
      [10, 5, '=A2>=B2'],
      [5, 10, '=A3>=B3'],
      [5, 5, '=A4>=B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  },
  {
    name: 'LESS_THAN_OR_EQUAL',
    category: '00. 基本演算子',
    description: '以下演算子（<=）',
    data: [
      ['値1', '値2', '結果'],
      [5, 10, '=A2<=B2'],
      [10, 5, '=A3<=B3'],
      [5, 5, '=A4<=B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  },
  {
    name: 'EQUAL',
    category: '00. 基本演算子',
    description: '等価演算子（=）',
    data: [
      ['値1', '値2', '結果'],
      [10, 10, '=A2=B2'],
      [5, 10, '=A3=B3'],
      ['ABC', 'ABC', '=A4=B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  },
  {
    name: 'NOT_EQUAL',
    category: '00. 基本演算子',
    description: '不等価演算子（<>）',
    data: [
      ['値1', '値2', '結果'],
      [10, 5, '=A2<>B2'],
      [10, 10, '=A3<>B3'],
      ['ABC', 'DEF', '=A4<>B4']
    ],
    expectedValues: { 'C2': true, 'C3': false, 'C4': true }
  }
];
