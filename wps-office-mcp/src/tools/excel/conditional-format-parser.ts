/**
 * Input: 条件格式的字符串描述（condition / format）
 * Output: wps-com.ps1 addConditionalFormat 所需的结构化参数
 * Pos: wps_excel_set_conditional_format 的参数解析层
 */

export interface ParsedCondition {
  operator: string;
  value1: string;
  value2?: string;
}

export interface ParsedFormat {
  backgroundColor?: string;
  fontColor?: string;
  bold?: boolean;
}

const COLOR_MAP: Record<string, { fill: string; font: string }> = {
  red: { fill: '#FFC7CE', font: '#9C0006' },
  green: { fill: '#C6EFCE', font: '#006100' },
  yellow: { fill: '#FFEB9C', font: '#9C5700' },
  orange: { fill: '#FCD5B4', font: '#C65911' },
  blue: { fill: '#DDEBF7', font: '#1F4E78' },
  purple: { fill: '#E4DFEC', font: '#7030A0' },
  gray: { fill: '#EDEDED', font: '#595959' },
  grey: { fill: '#EDEDED', font: '#595959' },
};

/** 非数字的比较值按文本处理，需要写成 ="文本" 形式 */
function toFormulaValue(raw: string): string {
  const v = raw.trim().replace(/^["']|["']$/g, '');
  return v !== '' && !isNaN(Number(v)) ? v : `="${v.replace(/"/g, '""')}"`;
}

const OPERATOR_ALIASES: Record<string, string> = {
  greater: 'greater', greaterthan: 'greater', gt: 'greater', '>': 'greater',
  less: 'less', lessthan: 'less', lt: 'less', '<': 'less',
  greaterequal: 'greaterEqual', greaterthanorequal: 'greaterEqual', gte: 'greaterEqual', ge: 'greaterEqual', '>=': 'greaterEqual',
  lessequal: 'lessEqual', lessthanorequal: 'lessEqual', lte: 'lessEqual', le: 'lessEqual', '<=': 'lessEqual',
  equal: 'equal', equals: 'equal', eq: 'equal', '=': 'equal',
  notequal: 'notEqual', ne: 'notEqual', '<>': 'notEqual', '!=': 'notEqual',
  between: 'between', notbetween: 'notBetween',
};

/** 对象形式：{ operator: "greaterThan", value1: 3000, value2? } */
function parseConditionObject(obj: Record<string, unknown>): ParsedCondition {
  const rawOp = String(obj.operator ?? 'equal').replace(/[\s_-]/g, '').toLowerCase();
  const operator = OPERATOR_ALIASES[rawOp];
  const v1 = obj.value1 ?? obj.value;
  if (!operator || v1 === undefined || v1 === null) {
    throw new Error(`无法解析条件对象: ${JSON.stringify(obj)}，需要 operator 与 value1`);
  }
  const result: ParsedCondition = { operator, value1: toFormulaValue(String(v1)) };
  if (operator === 'between' || operator === 'notBetween') {
    if (obj.value2 === undefined || obj.value2 === null) throw new Error(`${operator} 需要 value2`);
    result.value2 = toFormulaValue(String(obj.value2));
  }
  return result;
}

/**
 * 支持：>100  <0  >=5  <=5  =0  <>0  !=0  between(1,10)  notBetween(1,10)
 * 不带运算符的纯值视为等于；也接受 { operator, value1, value2 } 对象
 */
export function parseCondition(condition: unknown): ParsedCondition {
  if (condition && typeof condition === 'object') return parseConditionObject(condition as Record<string, unknown>);
  const c = String(condition ?? '').trim();
  const range = c.match(/^(between|notBetween)\s*\(\s*([^,]+?)\s*,\s*([^)]+?)\s*\)$/i);
  if (range) {
    return {
      operator: range[1].toLowerCase() === 'between' ? 'between' : 'notBetween',
      value1: toFormulaValue(range[2]),
      value2: toFormulaValue(range[3]),
    };
  }
  const cmp = c.match(/^(>=|<=|<>|!=|>|<|=)\s*(.+)$/);
  const opMap: Record<string, string> = {
    '>': 'greater', '<': 'less', '>=': 'greaterEqual', '<=': 'lessEqual',
    '=': 'equal', '<>': 'notEqual', '!=': 'notEqual',
  };
  if (cmp) return { operator: opMap[cmp[1]], value1: toFormulaValue(cmp[2]) };
  if (c) return { operator: 'equal', value1: toFormulaValue(c) };
  throw new Error(`无法解析条件: "${condition}"，示例: ">100"、"=0"、"between(1,10)"`);
}

/**
 * 支持逗号/空格/+ 组合：red_fill、green_font、bold、"red_fill,bold"
 * 也接受十六进制颜色：fill:#FF0000、font:#0000FF
 * 以及对象形式：{ bgColor | backgroundColor | fill, fontColor | color, bold }
 */
export function parseFormat(format: unknown): ParsedFormat {
  const result: ParsedFormat = {};
  if (format && typeof format === 'object') {
    const f = format as Record<string, unknown>;
    const bg = f.bgColor ?? f.backgroundColor ?? f.fill;
    const fg = f.fontColor ?? f.color;
    if (bg) result.backgroundColor = String(bg);
    if (fg) result.fontColor = String(fg);
    if (f.bold === true) result.bold = true;
    if (!result.bold && !result.backgroundColor && !result.fontColor) {
      throw new Error(`格式对象为空或无法识别: ${JSON.stringify(format)}`);
    }
    return result;
  }
  const tokens = String(format ?? '').toLowerCase().split(/[\s,;+|]+/).filter(Boolean);
  for (const t of tokens) {
    const hex = t.match(/^(fill|bg|font):(#[0-9a-f]{6})$/);
    const named = t.match(/^([a-z]+)_(fill|bg|background|font|text)$/);
    if (t === 'bold') {
      result.bold = true;
    } else if (hex) {
      if (hex[1] === 'font') result.fontColor = hex[2];
      else result.backgroundColor = hex[2];
    } else if (named && COLOR_MAP[named[1]]) {
      if (named[2] === 'font' || named[2] === 'text') result.fontColor = COLOR_MAP[named[1]].font;
      else result.backgroundColor = COLOR_MAP[named[1]].fill;
    } else {
      throw new Error(
        `无法解析格式: "${t}"，支持 bold、<颜色>_fill、<颜色>_font（颜色: ${Object.keys(COLOR_MAP).join('/')}）或 fill:#RRGGBB / font:#RRGGBB`
      );
    }
  }
  if (!result.bold && !result.backgroundColor && !result.fontColor) {
    throw new Error(`格式为空: "${format}"`);
  }
  return result;
}
