using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Demo
{
    public class NPOIDemo
    {
        #region 创建Excel颜色列表
        public static void CreateExcelColorFile(string fileName)
        {
            System.IO.FileInfo fileInfo = new System.IO.FileInfo(fileName);
            // 创建Workbook。
            NPOI.SS.UserModel.IWorkbook workbook;
            if (fileInfo.Extension == ".xlsx")
            { workbook = new NPOI.XSSF.UserModel.XSSFWorkbook(); }
            else if (fileInfo.Extension == ".xls")
            { workbook = new NPOI.HSSF.UserModel.HSSFWorkbook(); }
            else
            { return; }
            // 创建Sheet
            NPOI.SS.UserModel.ISheet sheet = workbook.CreateSheet("ExcelColor");
            // 设置列宽。
            sheet.SetColumnWidth(0, 15 * 256);
            sheet.SetColumnWidth(1, 15 * 256);
            sheet.SetColumnWidth(2, 20 * 256);
            sheet.SetColumnWidth(3, 20 * 256);
            sheet.SetColumnWidth(4, 15 * 256);

            //sheet.AutoSizeColumn(0);
            //sheet.AutoSizeColumn(1);
            //sheet.AutoSizeColumn(2);
            //sheet.AutoSizeColumn(3);
            //sheet.AutoSizeColumn(4);


            // 创建标题。
            int rowIndex = 0;
            NPOI.SS.UserModel.ICell cell = sheet.CreateRow(rowIndex).CreateCell(0);
            cell.SetCellValue("Excel颜色");
            cell.CellStyle = Studio.Extension.NPOIMethod.GetCellStyle(workbook, Studio.Office.Excel.ExcelColor.None.IndexNPOI, Studio.Office.Excel.ExcelColor.Black.IndexNPOI, new System.Drawing.Font("Arial", 16, System.Drawing.FontStyle.Bold));
            // 空一行
            rowIndex++;
            // 创建标题行
            rowIndex++;
            NPOI.SS.UserModel.IRow row = sheet.CreateRow(rowIndex);
            row.Height = 30 * 20;
            NPOI.SS.UserModel.ICellStyle columnHeadCellStyle = Studio.Extension.NPOIMethod.GetCellStyle(workbook, Studio.Office.Excel.ExcelColor.None.IndexNPOI, Studio.Office.Excel.ExcelColor.Black.IndexNPOI, new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.VerticalAlignment.Center);
            cell = row.CreateCell(0);
            cell.SetCellValue("索引");
            cell.CellStyle = columnHeadCellStyle;

            cell = row.CreateCell(1);
            cell.SetCellValue("索引（NPOI）");
            cell.CellStyle = columnHeadCellStyle;

            cell = row.CreateCell(2);
            cell.SetCellValue("名称");
            cell.CellStyle = columnHeadCellStyle;
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(rowIndex, rowIndex, 2, 3));

            cell = row.CreateCell(4);
            cell.SetCellValue("十六进制代码");
            cell.CellStyle = columnHeadCellStyle;

            Studio.Office.Excel.ExcelColor.KnownColors.ForEach(x =>
            {
                rowIndex++;
                NPOI.SS.UserModel.IRow contentRow = sheet.CreateRow(rowIndex);
                contentRow.Height = 30 * 20;
                NPOI.SS.UserModel.ICellStyle contentCellStyle = Studio.Extension.NPOIMethod.GetCellStyle(workbook, x.IndexNPOI, (short)(x.IsDarkColor ? 9 : 8), new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Center, NPOI.SS.UserModel.VerticalAlignment.Center);
                NPOI.SS.UserModel.ICell contentCell = contentRow.CreateCell(0);
                contentCell.SetCellValue(x.Index);
                contentCell.CellStyle = contentCellStyle;

                contentCell = contentRow.CreateCell(1);
                contentCell.SetCellValue(x.IndexNPOI + "（NPOI）");
                contentCell.CellStyle = contentCellStyle;

                contentCell = contentRow.CreateCell(2);
                contentCell.SetCellValue(x.Name);
                contentCell.CellStyle = Studio.Extension.NPOIMethod.GetCellStyle(workbook, x.IndexNPOI, (short)(x.IsDarkColor ? 9 : 8), new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Right, NPOI.SS.UserModel.VerticalAlignment.Center);

                contentCell = contentRow.CreateCell(3);
                contentCell.SetCellValue(x.Description);
                contentCell.CellStyle = Studio.Extension.NPOIMethod.GetCellStyle(workbook, x.IndexNPOI, (short)(x.IsDarkColor ? 9 : 8), new System.Drawing.Font("Consolas", 11, System.Drawing.FontStyle.Regular), NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Center);

                contentCell = contentRow.CreateCell(4);
                contentCell.SetCellValue(x.HexString);
                contentCell.CellStyle = contentCellStyle;

            });

            sheet.AutoSizeColumn(0);
            sheet.AutoSizeColumn(1);
            sheet.AutoSizeColumn(2);
            sheet.AutoSizeColumn(3);
            sheet.AutoSizeColumn(4);

            // 保存到文件。
            System.IO.FileStream fileStream = new System.IO.FileStream(fileInfo.FullName, System.IO.FileMode.Create, System.IO.FileAccess.ReadWrite);
            workbook.Write(fileStream);
            fileStream.Close();
            workbook.Close();

        }
        #endregion
    }
}
