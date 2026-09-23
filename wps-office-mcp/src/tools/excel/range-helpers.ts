/**
 * Input: 行号/列号/列字母
 * Output: wps-com.ps1 与 macOS 加载项 hideRows/showRows/hideColumns/showColumns 所需的 rows/columns 数组
 * Pos: excel 行列类工具的参数转换层
 */

/** 列字母转列号：A -> 1，AA -> 27 */
export function columnToNumber(col: string | number): number {
  if (typeof col === 'number') return col;
  const s = String(col).trim().toUpperCase();
  if (/^\d+$/.test(s)) return Number(s);
  if (!/^[A-Z]+$/.test(s)) throw new Error(`无效列: "${col}"`);
  return [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
}

/** 列号转列字母：1 -> A，27 -> AA */
export function numberToColumn(n: number): string {
  if (!Number.isInteger(n) || n < 1) throw new Error(`无效列号: ${n}`);
  let s = '';
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  return s;
}

/** 从 start 起连续 count 个整数 */
export function seq(start: number, count: number): number[] {
  const c = Math.max(1, Math.floor(count || 1));
  return Array.from({ length: c }, (_, i) => start + i);
}

/** 闭区间 [start, end] 的整数，end 缺省时等于 start */
export function between(start: number, end?: number): number[] {
  const e = end ?? start;
  if (e < start) throw new Error(`结束 ${e} 小于开始 ${start}`);
  return seq(start, e - start + 1);
}

/**
 * 列标识转为相对区域的字段序号（AutoFilter/Subtotal 的 Field 从 1 开始）
 * 纯数字视为已是字段序号；列字母按区域起始列换算，如 range="B1:E9" 时 "C" -> 2
 */
export function fieldIndex(range: string, col: string | number): number {
  const s = String(col).trim();
  if (/^\d+$/.test(s)) return Number(s);
  const start = range.replace(/^.*!/, '').replace(/\$/g, '').match(/^([A-Za-z]+)/);
  if (!start) throw new Error(`无法从范围 "${range}" 解析起始列`);
  const idx = columnToNumber(s) - columnToNumber(start[1]) + 1;
  const width = range.includes(':') ? columnToNumber(range.replace(/\$/g, '').split(':')[1].match(/^([A-Za-z]+)/)?.[1] ?? start[1]) - columnToNumber(start[1]) + 1 : 1;
  if (idx < 1 || idx > width) throw new Error(`列 ${s} 不在范围 ${range} 内`);
  return idx;
}

/** 列区间转列字母数组，接受字母或列号 */
export function columnLetters(start: string | number, end?: string | number): string[] {
  const s = columnToNumber(start);
  return between(s, end === undefined ? s : columnToNumber(end)).map(numberToColumn);
}
