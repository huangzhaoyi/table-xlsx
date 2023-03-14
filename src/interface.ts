export type DefaultValueType = Record<string, any>
export type DataIndex = string | number | undefined | readonly (string | number)[];

export interface MergesArrType {
  s: {
    c: number,
    r: number
  },
  e: {
    c: number,
    r: number
  }
}

export interface MergesObjType {
  [key: string]: {
    colSpan?: number,
    rowSpan?: number,
  };
}

export interface ColumnType {
  key?: string,
  title?: string,
  dataIndex?: string,
  mergesObj?: MergesObjType,
  render?: (value: any, row: DefaultValueType, rowIndex: number) => any,
  onTxBodyCell?: (row: DefaultValueType, rowIndex: number) => { style: CellStyleType },
  children?: ColumnType[],

  [key: string]: any,
}

export interface DataType {
  [key: string]: any,
}

export interface TableType {
  sheetName?: string,
  dataSource?: DataType[],
  columns?: ColumnType[],
}

export interface CellStyleType {
  fontSize?: number,
  fontBold?: boolean,
  fontName?: string,
  fontColorRgb?: string,
  fillFgColorRgb?: string,
  borderStyle?: string,
  borderColorRgb?: string,
  alignmentHorizontal?: string,
  alignmentVertical?: string,
}

export interface CellStyleModel {
  /**
   * 原始值（数字、字符串、日期对象、布尔值）
   */
  v?: string;
  /**
   * 类型：b布尔值，e错误，n数字，d日期，s文本，z存根
   */
  t?: string;
  /**
   * 与单元格关联的数字格式字符串
   */
  z?: string;
  /**
   * 数字格式化文本
   */
  w?: string;
  /**
   * 单元格公式编码为A1样式字符串（如果适用）
   */
  f?: string;
  /**
   * 单元格超链接/工具提示（更多信息）
   */
  l?: string;
  /**
   * 单元格注释（更多信息）
   */
  c?: string;
  /**
   *  富文本编码（如果适用）
   */
  r?: string;
  /**
   * 富文本的HTML呈现（如果适用）
   */
  h?: string;
  /**
   * 单元格的样式/主题（如果适用）
    */
  s?: CellStyle;
}

export interface CellStyle {
  fill?: CellFillStyle;
  font?: CellFontStyle;
  border?: CellBorderStyle;
  alignment: CellAlignmentStyle
}

/**
 *  alignmentHorizontal = 'center', // left center right
 *   alignmentVertical = 'center', // top center bottom
 *   alignmentWrapText = false, // true false
 *   alignmentReadingOrder = 2, // for right-to-left
 *   alignmentTextRotation = 0, // Number from 0 to 180 or 255
 */
export interface CellAlignmentStyle {
  horizontal?: 'left'| 'center' |'right',
  vertical?: 'top' |'center' |'bottom',
  wrapText: boolean,
  readingOrder?:number,
  textRotation?: number,
}

export interface CellFillStyle {
  fgColor?: CellRgb;
}

export interface CellFontStyle {
  /**
   * 字体大小
   */
  sz?: number;
  /**
   * 字体颜色
   */
  color?: CellRgb
  /**
   * 字体名称
   */
  name?: string,
  /**
   * 是否加粗
   */
  bold?: boolean
}

export interface BorderStyle {
  /**
   * 例：'thin'
   */
  style: string;
  color: CellRgb;
}

export interface CellBorderStyle {
  top?: BorderStyle,
  left?: BorderStyle,
  bottom?: BorderStyle,
  right?: BorderStyle,
  diagonal?: BorderStyle
}

export interface CellRgb {
  /**
   * 不需要带#
   */
  rgb: string;
}

export interface SheetType {
  '!ref'?: string,
  '!cols'?: { wpx?: number }[],
  '!rows'?: { hpx?: number }[],
  '!merges'?: MergesArrType[],

  [key: string]: {
    t?: string,
    v?: string | number,
    s?: DefaultValueType
  } | undefined | string
    | { wpx?: number }[]
    | { hpx?: number }[]
    | MergesArrType[],

}

export interface HeaderCellType {
  title?: string,
  merges?: MergesArrType,
  txHeaderCellStyle?: CellStyleType,
}
