using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Extension
{
    public static class NPOIMethod
    {
        #region 公有方法

        #region 获取单个格的值
        public static dynamic GetCellValue(NPOI.SS.UserModel.ICell cell)
        {
            if (cell == null)
            { return null; }
            else
            {
                switch (cell.CellType)
                {
                    case NPOI.SS.UserModel.CellType.Blank:
                        return "";
                    case NPOI.SS.UserModel.CellType.Boolean:
                        return cell.BooleanCellValue;
                    case NPOI.SS.UserModel.CellType.Error:
                        return cell.ErrorCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case NPOI.SS.UserModel.CellType.Formula:
                        if (cell.CachedFormulaResultType == NPOI.SS.UserModel.CellType.Error)
                        { return "#NUM!"; }
                        else
                        { return cell.NumericCellValue; }
                    case NPOI.SS.UserModel.CellType.Numeric:
                        if (NPOI.SS.UserModel.DateUtil.IsCellDateFormatted(cell))
                        { return cell.DateCellValue; }
                        else
                        { return cell.NumericCellValue; }
                    case NPOI.SS.UserModel.CellType.String:
                        return cell.StringCellValue;
                    default:
                        return cell.ToString();
                }
            }
        }
        #endregion

        #region 获取单元格样式
        public static NPOI.SS.UserModel.ICellStyle GetCellStyle(NPOI.SS.UserModel.IWorkbook workbook, short backColorIndex, short fontColorIndex, System.Drawing.Font font, NPOI.SS.UserModel.HorizontalAlignment horizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment.General, NPOI.SS.UserModel.VerticalAlignment verticalAlignment = NPOI.SS.UserModel.VerticalAlignment.None, NPOI.SS.UserModel.BorderStyle borderLeft = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderTop = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderRight = NPOI.SS.UserModel.BorderStyle.None, NPOI.SS.UserModel.BorderStyle borderBottom = NPOI.SS.UserModel.BorderStyle.None)
        {
            // 获取样式
            NPOI.SS.UserModel.ICellStyle cellStyle = workbook.CreateCellStyle();
            // 设置单元格样式
            if (backColorIndex > 0)
            {
                cellStyle.FillForegroundColor = backColorIndex;
                cellStyle.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            }
            cellStyle.Alignment = horizontalAlignment;
            cellStyle.VerticalAlignment = verticalAlignment;
            cellStyle.BorderLeft = borderLeft;
            cellStyle.BorderTop = borderTop;
            cellStyle.BorderRight = borderRight;
            cellStyle.BorderBottom = borderBottom;
            // 设置字体样式
            if (font != null)
            {
                NPOI.SS.UserModel.IFont cellFont = workbook.CreateFont();
                if (fontColorIndex > 0)
                { cellFont.Color = fontColorIndex; }
                cellFont.FontName = font.Name;
                cellFont.FontHeightInPoints = (short)font.Size;
                cellFont.Boldweight = (short)(font.Bold ? NPOI.SS.UserModel.FontBoldWeight.Bold : NPOI.SS.UserModel.FontBoldWeight.Normal);
                cellStyle.SetFont(cellFont);
            }
            return cellStyle;
        }
        #endregion

        #region 将工作簿保存到文件
        public static void SaveToFile(NPOI.SS.UserModel.IWorkbook workbook, string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite);
            workbook.Write(fileStream);
            fileStream.Close();
            workbook.Close();
        }
        #endregion

        #region 从文件中获取工作簿
        public static NPOI.SS.UserModel.IWorkbook GetWorkbookFromExcelFile(string fileName)
        {
            NPOI.SS.UserModel.IWorkbook workbook = null;
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            if (!fileInfo.Exists)
            { return workbook; }
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            workbook = NPOI.SS.UserModel.WorkbookFactory.Create(fileStream);
            fileStream.Close();
            return workbook;
        }
        #endregion

        #region 从Sheet中获取DataTable
        public static System.Data.DataTable GetDataTableFromSheet(NPOI.SS.UserModel.ISheet sheet, int startRowIndex, int startColumnIndex, bool firstRowIsColumnHead, bool autoAddColumn, bool ignoreBlankRow)
        {
            if (sheet == null) { return null; }
            System.Data.DataTable dataTable = new System.Data.DataTable(sheet.SheetName);
            NPOI.SS.UserModel.IRow row = null;
            if (firstRowIsColumnHead)
            { row = sheet.GetRow(0); }
            else
            { row = sheet.GetRow(startRowIndex); }
            if (row != null)
            {
                for (int i = startColumnIndex; i < row.PhysicalNumberOfCells; i++)
                {
                    NPOI.SS.UserModel.ICell cell = row.GetCell(i);
                    if (cell == null)
                    { dataTable.Columns.Add(); }
                    else
                    {
                        if (firstRowIsColumnHead)
                        { dataTable.Columns.Add(cell.ToString()); }
                        else
                        { dataTable.Columns.Add(); }
                    }
                }
            }
            if (startRowIndex == 0 && firstRowIsColumnHead)
            { startRowIndex = 1; }
            for (int rowIndex = startRowIndex; rowIndex < sheet.PhysicalNumberOfRows; rowIndex++)
            {
                row = sheet.GetRow(rowIndex);
                if (row != null)
                {
                    System.Data.DataRow dataRow = dataTable.NewRow();
                    dataTable.Rows.Add(dataRow);
                    int dataTableColumnIndex = 0;
                    for (int columnIndex = startColumnIndex; columnIndex < row.PhysicalNumberOfCells; columnIndex++)
                    {
                        if (dataTableColumnIndex >= dataTable.Columns.Count)
                        {
                            if (autoAddColumn)
                            { dataTable.Columns.Add(); }
                            else
                            { break; }
                        }
                        dataRow[dataTableColumnIndex] = GetCellValue(row.GetCell(columnIndex));
                        dataTableColumnIndex++;
                    }
                }
            }
            return dataTable;
        }
        #endregion

        #region 从文件中获取DataSet
        public static System.Data.DataSet GetDataSetFromExcelFile(string fileName, bool firstRowIsColumnHead = true, bool ignoreBlankRow = true, bool ignoreHiddenSheet = true)
        {
            NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
            if (workbook == null)
            { return null; }
            System.Data.DataSet dataSet = new System.Data.DataSet();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                bool isHiddenSheet = workbook.IsSheetHidden(i) || workbook.IsSheetVeryHidden(i);
                if (ignoreHiddenSheet && isHiddenSheet)
                { continue; }
                dataSet.Tables.Add(GetDataTableFromSheet(workbook.GetSheetAt(i), 0, 0, true, true, true));
            }
            return dataSet;
        }
        #endregion

        #region 从文件中获取DataTable
        public static System.Data.DataTable GetDataTableFromExcelFile(string fileName, string sheetName, int startRowIndex, int startColumnIndex, bool firstRowIsColumnHead = true, bool autoAddColumn = false, bool ignoreBlankRow = true)
        {
            NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
            if (workbook == null)
            { return null; }
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            { return null; }
            return GetDataTableFromSheet(sheet, startRowIndex, startColumnIndex, firstRowIsColumnHead, autoAddColumn, ignoreBlankRow);
        }
        #endregion

        #region 从文件中获取DataTable
        public static System.Data.DataTable GetDataTableFromExcelFile(string fileName, string sheetName, int startRowIndex, int startColumnIndex, bool ignoreBlankRow, params string[] columnHeadNames)
        {
            NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
            if (workbook == null)
            { return null; }
            NPOI.SS.UserModel.ISheet sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            { return null; }
            System.Data.DataTable dataTable = new System.Data.DataTable(sheet.SheetName);
            for (int i = 0; i < columnHeadNames.Length; i++)
            {
                dataTable.Columns.Add(columnHeadNames[i]);
            }
            if (columnHeadNames.Length > 0)
            {
                for (int rowIndex = startRowIndex; rowIndex < sheet.PhysicalNumberOfRows; rowIndex++)
                {
                    NPOI.SS.UserModel.IRow row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        System.Data.DataRow dataRow = dataTable.NewRow();
                        dataTable.Rows.Add(dataRow);
                        int dataTableColumnIndex = 0;
                        for (int columnIndex = startColumnIndex; columnIndex < row.PhysicalNumberOfCells; columnIndex++)
                        {
                            if (dataTableColumnIndex >= dataTable.Columns.Count)
                            { break; }
                            dataRow[dataTableColumnIndex] = GetCellValue(row.GetCell(columnIndex));
                            dataTableColumnIndex++;
                        }
                    }
                }
            }
            return dataTable;
        }
        #endregion

        //#region 从文件中获取数据
        //public static System.Data.DataTable GetDataTabelFromExcelFile(string fileName, int sheetNumber, bool firstRowIsColumnHead = true, bool ignoreBlankRow = true, bool ignoreHiddenSheet = true)
        //{
        //    System.Data.DataTable dataTable = new System.Data.DataTable();
        //    NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
        //    if (workbook == null)
        //    { return dataTable; }
        //    NPOI.SS.UserModel.ISheet sheet = null;
        //    int sheetNO = 1;
        //    if (ignoreHiddenSheet)
        //    {
        //        for (int i = 0; i < workbook.NumberOfSheets; i++)
        //        {
        //            if (workbook.IsSheetHidden(i) || workbook.IsSheetVeryHidden(i))
        //            {
        //                continue;
        //            }
        //            if (sheetNO == sheetNumber)
        //            {
        //                sheet = workbook.GetSheetAt(i);
        //                break;
        //            }
        //            sheetNO++;
        //        }
        //    }
        //    else
        //    { workbook.GetSheetAt(sheetNumber - 1); }
        //    if (sheet == null)
        //    { return dataTable; }
        //    dataTable.TableName = sheet.SheetName;
        //    if (firstRowIsColumnHead)
        //    {
        //        NPOI.SS.UserModel.IRow columnHeadRow = sheet.GetRow(0);
        //        if (columnHeadRow != null)
        //        {
        //            for (int columnIndex = 0; columnIndex < columnHeadRow.LastCellNum; columnIndex++)
        //            {
        //                NPOI.SS.UserModel.ICell cell = columnHeadRow.GetCell(columnIndex);
        //                if (cell == null)
        //                { dataTable.Columns.Add(); }
        //                else
        //                { dataTable.Columns.Add(cell.ToString()); }
        //            }
        //        }
        //    }
        //    FillDataTable(sheet, firstRowIsColumnHead ? 1 : 0, 0, ignoreBlankRow, !firstRowIsColumnHead, dataTable);
        //    return dataTable;
        //}
        //#endregion

        //#region 从文件中获取数据
        //public static System.Data.DataTable GetDataTabelFromExcelFile(string fileName, bool firstRowIsColumnHead, bool ignoreBlankRow)
        //{
        //    System.Data.DataTable dataSource = new System.Data.DataTable();
        //    NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
        //    if (workbook == null)
        //    { return dataSource; }
        //    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(workbook.ActiveSheetIndex);
        //    if (sheet == null)
        //    { return dataSource; }
        //    dataSource.TableName = sheet.SheetName;
        //    if (firstRowIsColumnHead)
        //    {
        //        NPOI.SS.UserModel.IRow columnHeadRow = sheet.GetRow(0);
        //        if (columnHeadRow != null)
        //        {
        //            for (int columnIndex = 0; columnIndex < columnHeadRow.LastCellNum; columnIndex++)
        //            {
        //                NPOI.SS.UserModel.ICell cell = columnHeadRow.GetCell(columnIndex);
        //                if (cell == null)
        //                { dataSource.Columns.Add(); }
        //                else
        //                { dataSource.Columns.Add(cell.ToString()); }
        //            }
        //        }
        //    }
        //    FillDataTable(sheet, firstRowIsColumnHead ? 1 : 0, 0, ignoreBlankRow, true, dataSource);
        //    return dataSource;
        //}
        //#endregion

        //#region 从文件中获取数据
        //public static System.Data.DataTable GetDataTabelFromExcelFile(string fileName, int sheetNumber, int startColumnNumber, int startRowNumber, int columnCount, bool ignoreBlankRow, params string[] columnNames)
        //{
        //    System.Data.DataTable dataTable = new System.Data.DataTable();
        //    for (int i = 0; i < columnCount; i++)
        //    {
        //        if (i < columnNames.Length)
        //        { dataTable.Columns.Add(columnNames[i]); }
        //        else
        //        { dataTable.Columns.Add(); }
        //    }
        //    NPOI.SS.UserModel.IWorkbook workbook = GetWorkbookFromExcelFile(fileName);
        //    if (workbook == null)
        //    { return dataTable; }
        //    if (sheetNumber < 0 || sheetNumber > workbook.NumberOfSheets)
        //    { return dataTable; }
        //    NPOI.SS.UserModel.ISheet sheet = workbook.GetSheetAt(sheetNumber - 1);
        //    if (sheet == null)
        //    { return dataTable; }
        //    FillDataTable(sheet, startRowNumber < 1 ? 0 : startRowNumber - 1, startColumnNumber < 1 ? 0 : startColumnNumber - 1, ignoreBlankRow, false, dataTable);
        //    return dataTable;
        //}
        //#endregion

        #endregion


    }
}
