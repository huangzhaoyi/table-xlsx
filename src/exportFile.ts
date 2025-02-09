import {
  CellStyleModel,
  CellStyleType,
  ColumnType,
  DataType,
  DefaultValueType,
  HeaderCellType,
  MergesArrType,
  SheetType
} from './interface';

import {sameType} from './utils/base';
import {flattenColumns, formatToWpx, getHeader2dArray} from './utils/columnsUtils';
import {getPathValue, getRenderValue} from './utils/valueUtils';
import {getStyles} from './utils/cellStylesUtils';

const XLSX = require('@pengchen/xlsx');

const ROW_HPX = 25, ROW_HPT = 25;
/**
 * 导出
 * @param fileName
 * @param sheetNames
 * @param columns
 * @param dataSource
 * @param showHeader 是否显示表头
 * @param raw 是否格式化值的类型
 * @param cellStyle 单元格样式
 * @param headerCellStyle 单元格样式
 * @param bodyCellStyle 单元格样式
 * @param useRender 是否使用render返回的值
 * @param onTxBodyRow
 * @param cellStyleModels
 */
export const exportFile = (
  {
    fileName = 'table.xlsx',
    sheetNames = ['sheet1'],
    columns = [],
    dataSource = [],
    showHeader = true,
    raw = false,
    cellStyle = {},
    headerCellStyle = {},
    bodyCellStyle = {},
    useRender = true,
    onTxBodyRow,
    cellStyleModels = {}
  }: {
    fileName?: string,
    sheetNames?: (string | number)[],
    columns: ColumnType[][],
    dataSource: DataType[][],
    showHeader?: boolean,
    raw?: boolean,
    cellStyle?: CellStyleType,
    headerCellStyle?: CellStyleType,
    bodyCellStyle?: CellStyleType,
    useRender?: boolean,
    onTxBodyRow?: (row: DefaultValueType, rowIndex: number) => { style: CellStyleType },
    cellStyleModels?: Record<string, CellStyleModel>
  }
): {
  SheetNames: (string | number)[],
  Sheets: { [key: string]: SheetType }
} => {
  const Sheets: { [key: string]: SheetType } = {};
  sheetNames.forEach((sheetName: string | number, sheetIndex: number) => {
    const _columns = sameType(columns[sheetIndex], 'Array') ? columns[sheetIndex] : columns;
    const _dataSource = sameType(dataSource[sheetIndex], 'Array') ? dataSource[sheetIndex] : dataSource;
    const {sheet} = formatToSheet({
      columns: _columns,
      dataSource: _dataSource,
      useRender,
      showHeader,
      raw,
      cellStyle,
      headerCellStyle,
      bodyCellStyle,
      onTxBodyRow,
      cellStyleModels: cellStyleModels
    });
    Sheets[sheetName] = sheet;
  });
  const wb = {
    SheetNames: sheetNames,
    Sheets: Sheets
  };
  XLSX.writeFile(wb, fileName);
  return wb;
};
/**
 * 转换成sheet对象
 */
const formatToSheet = (
  {
    columns,
    dataSource,
    showHeader,
    raw,
    cellStyle,
    headerCellStyle,
    bodyCellStyle,
    useRender,
    onTxBodyRow,
    cellStyleModels
  }: {
    columns: ColumnType[],
    dataSource: DataType[],
    showHeader: boolean,
    raw: boolean,
    cellStyle?: CellStyleType,
    headerCellStyle?: CellStyleType,
    bodyCellStyle?: CellStyleType,
    useRender?: boolean,
    onTxBodyRow?: (row: DefaultValueType, rowIndex: number) => { style: CellStyleType },
    cellStyleModels?: Record<string, CellStyleModel>
  }
) => {
  const sheet: SheetType = {};
  const $cols: { wpx: number }[] = [];
  const $rows: { hpx: number, hpt: number }[] = [];
  const $merges: MergesArrType[] = [];
  //
  const {columns: flatColumns, level} = flattenColumns({columns});
  let headerLevel = level;
  if (showHeader) {
    for (let i = 0; i < headerLevel; i++) {
      $rows.push({hpx: ROW_HPX, hpt: ROW_HPT});
    }
    // 表头信息
    const headerData = getHeaderData({columns, headerLevel, cellStyle, headerCellStyle, cellStyleModels});
    Object.assign(sheet, headerData.sheet);
    $merges.push(...headerData.merges);
  } else {
    headerLevel = 0;
  }
  //
  flatColumns.forEach((col: ColumnType, colIndex: number) => {
    const key = col.dataIndex || col.key;
    $cols.push({wpx: formatToWpx(col.width)});
    const xAxis = XLSX.utils.encode_col(colIndex);
    dataSource.forEach((data: DataType, rowIndex: number) => {
      if (colIndex === 0) {
        $rows.push({hpx: data.ROW_HPX || ROW_HPX, hpt: data.ROW_HPT || ROW_HPT});
      }
      let value = getPathValue(data, key);
      if (col.render) {
        const renderResult = col.render(value, data, rowIndex);
        value = useRender ? getRenderValue(renderResult) : value;
        const merge = getMerge({renderResult, colIndex, rowIndex, headerLevel});
        if (merge) {
          $merges.push(merge);
        }
      }
      let txBodyRowStyle = {};
      let txBodyCellStyle = {};
      if (onTxBodyRow) {
        const result = onTxBodyRow(data, rowIndex);
        txBodyRowStyle = result?.style || {};
      }
      if (col.onTxBodyCell) {
        const result = col.onTxBodyCell(data, rowIndex);
        txBodyCellStyle = result?.style || {};
      }
      const keyIndex = `${xAxis}${headerLevel + rowIndex + 1}`;
      const csm = cellStyleModels ? cellStyleModels[`${colIndex}${headerLevel + rowIndex}`] : {};
      sheet[keyIndex] = Object.assign({
        t: (raw && typeof value === 'number') ? 'n' : 's',
        v: value ? value : '',
        s: getStyles({
          alignmentHorizontal: 'left',
          ...cellStyle,
          ...bodyCellStyle,
          ...txBodyRowStyle,
          ...txBodyCellStyle,
        }),
      }, csm);
    });
  });
  const xe = XLSX.utils.encode_col(Math.max(flatColumns.length - 1, 0));
  const ye = headerLevel + dataSource.length;
  // const $sourceRows = sourceSheets['!rows'];
  // const realRows = Object.assign($rows, $sourceRows);
  // console.warn('realRows:', realRows);
  sheet['!ref'] = `A1:${xe}${ye}`;
  sheet['!cols'] = $cols;
  sheet['!rows'] = $rows;
  sheet['!merges'] = $merges;
  // console.warn('结果sheet:', sheet);
  return {
    sheet
  };
};

/**
 * 获取表头数据
 */
const getHeaderData = ({
  columns,
  headerLevel,
  cellStyle,
  headerCellStyle,
  cellStyleModels
}: {
  columns: ColumnType[],
  headerLevel: number,
  cellStyle?: CellStyleType,
  headerCellStyle?: CellStyleType,
  cellStyleModels?: Record<string, CellStyleModel>
}) => {
  const sheet: SheetType = {};
  const merges: { s: { c: number, r: number }, e: { c: number, r: number } }[] = [];

  const headerArr: HeaderCellType[][] = getHeader2dArray({columns, headerLevel});
  const mergesWeakMap = new WeakMap();
  headerArr.forEach((rowsArr: HeaderCellType[], rowIndex: number) => {
    rowsArr.forEach((cols: HeaderCellType, colIndex: number) => {
      const xAxis = XLSX.utils.encode_col(colIndex);
      const style = cols?.txHeaderCellStyle || {};
      const csm = cellStyleModels ? cellStyleModels[`${colIndex}${rowIndex}`] : {};
      // https://github.com/SheetJS/sheetjs#cell-object
      sheet[`${xAxis}${rowIndex + 1}`] = Object.assign({
        t: 's',
        v: cols.title,
        s: getStyles({
          fillFgColorRgb: 'e9ebf0',
          fontBold: true,
          ...cellStyle,
          ...headerCellStyle,
          ...style,
        }),
      }, csm);
      if (cols.merges && !mergesWeakMap.get(cols.merges)) {
        mergesWeakMap.set(cols.merges, true); // Microsoft Excel 如果传了相同的merges信息，文件会损坏，做个去重
        merges.push(cols.merges);
      }
    });
  });
  return {
    sheet,
    merges
  };
};

/**
 * 获取单个合并信息
 */
const getMerge = ({
  renderResult,
  colIndex,
  rowIndex,
  headerLevel
}: {
  renderResult: {
    props: { colSpan: number, rowSpan: number }
  },
  colIndex: number,
  rowIndex: number,
  headerLevel: number
}) => {
  if (renderResult?.props) {
    const {colSpan, rowSpan} = renderResult.props;
    if ((colSpan && colSpan !== 1) || (rowSpan && rowSpan !== 1)) {
      const realRowIndex = rowIndex + headerLevel;
      const merge: MergesArrType = {
        s: {c: colIndex, r: realRowIndex},
        e: {
          c: colIndex + (colSpan || 1) - 1,
          r: realRowIndex + (rowSpan || 1) - 1
        },
      };
      return merge;
    }
  }
  return false;
};
