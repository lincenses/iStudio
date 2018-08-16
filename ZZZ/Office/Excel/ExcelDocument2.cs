using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office.Excel
{
    public class ExcelDocument2 : IDisposable
    {
        #region API函数
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        #endregion

        #region 私有成员

        private Microsoft.Office.Interop.Excel.Application _Excel;

        private bool _IsDisposed = false;

        private int _ProcessID = 0;

        #endregion

        #region 实现接口
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_IsDisposed)// 如果资源未释放 这个判断主要用了防止对象被多次释放
            {
                if (disposing)
                {
                    // 释放托管资源
                }
                // 释放非托管资源
                _Excel.Application.DisplayAlerts = false;
                _Excel.Workbooks.Close();
                _Excel.Application.DisplayAlerts = true;
                _Excel.Quit();
                if (_ProcessID != 0)
                { System.Diagnostics.Process.GetProcessById(_ProcessID).Kill(); }
            }
            _IsDisposed = true; // 标识此对象已释放
        }
        #endregion

        #region 公有属性

        #region 获取对象是否已被释放
        public bool IsDisposed
        {
            get { return _IsDisposed; }
        }
        #endregion

        #region 获取句柄
        public IntPtr Hander
        {
            get { return new IntPtr(_Excel.Hwnd); }
        }
        #endregion

        #region 获取进程ID
        public int ProcessID
        {
            get { return _ProcessID; }
        }
        #endregion

        #endregion

        #region 构造函数

        #region 初始化此类的新实例
        public ExcelDocument2()
        {
            _Excel = new Microsoft.Office.Interop.Excel.Application();
            GetWindowThreadProcessId(new IntPtr(_Excel.Hwnd), out _ProcessID);
            _Excel.Workbooks.Add(Type.Missing);
            
        }

        public ExcelDocument2(string templateFileName)
        {
            _Excel = new Microsoft.Office.Interop.Excel.Application();
            GetWindowThreadProcessId(new IntPtr(_Excel.Hwnd), out _ProcessID);
            _Excel.Workbooks.Add(new System.IO.FileInfo(templateFileName).FullName);

        }

        ~ExcelDocument2()
        {
            Dispose(false);
        }
        #endregion

        #endregion

        #region 私有方法

        #region 获取单元格
        private Microsoft.Office.Interop.Excel.Range GetRange(int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, int sheetIndex = 0)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = sheetIndex == 0 ? _Excel.ActiveSheet : _Excel.Worksheets[sheetIndex];
            Microsoft.Office.Interop.Excel.Range startRange = sheet.Cells[startRowIndex, startColumnIndex];
            Microsoft.Office.Interop.Excel.Range endRange = sheet.Cells[endRowIndex, endColumnIndex];
            return sheet.get_Range(startRange, endRange);
        }
        #endregion

        #endregion

        #region 共有方法

        #region 关闭Excel
        public void Close()
        {
            Dispose();
        }
        #endregion

        #region 保存文件
        public void SaveAs(string fileName, ExcelFileFormat format = ExcelFileFormat.xlWorkbookDefault)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            _Excel.Application.DisplayAlerts = false;
            _Excel.ActiveWorkbook.SaveAs(fullName, (Microsoft.Office.Interop.Excel.XlFileFormat)format, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _Excel.Application.DisplayAlerts = true;
        }
        #endregion

        #region 保存Excel文件（PDF格式）
        public void SaveAsPDF(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            _Excel.Application.DisplayAlerts = false;
            _Excel.ActiveWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, fullName, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, false, Type.Missing);
            _Excel.Application.DisplayAlerts = true;
        }
        #endregion

        #region 获取Excel对象中的所有Sheet名称
        public string[] GetSheetNames()
        {
            if (_Excel != null)
            {
                return _Excel.Worksheets.Cast<Microsoft.Office.Interop.Excel.Worksheet>().Select<Microsoft.Office.Interop.Excel.Worksheet, string>(sheet => sheet.Name).ToArray();
            }
            else
            {
                return new string[0];
            }
        }
        #endregion

        #region 设置活动Sheet
        public void SetActiveSheet(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            sheet.Select();
        }
        #endregion

        #region 设置活动Sheet
        public void SetActiveSheet(string sheetName)
        {
            int sheetIndex = GetSheetIndex(sheetName);
            SetActiveSheet(sheetIndex);
        }
        #endregion

        #region 获取指定索引的Sheet名称
        public string GetSheetName(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            return sheet.Name;
        }
        #endregion

        #region 设置指定索引的Sheet名称
        public void SetSheetName(int index, string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            sheet.Name = sheetName;
        }
        #endregion

        #region 获取Sheet数量
        public int GetSheetCount()
        {
            return _Excel.Worksheets.Count;
        }
        #endregion

        #region 获取Sheet的索引
        public int GetSheetIndex(string sheetName)
        {
            return GetSheetNames().ToList().FindIndex(x => x.ToLower() == sheetName.ToLower()) + 1;
        }
        #endregion

        #region 添加一个Sheet
        public void AddSheet()
        {
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            _Excel.Worksheets.Add(Type.Missing, lastSheet, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
        }
        #endregion

        #region 添加一个Sheet
        public void AddSheet(string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            _Excel.Worksheets.Add(Type.Missing, lastSheet, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            //ActiveSheetName = sheetName;
        }
        #endregion

        #region 添加一个复制的Sheet
        public void AddCopySheet(int originalSheetIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            originalSheet.Copy(Type.Missing, lastSheet);
        }
        #endregion

        #region 添加一个复制的Sheet
        public void AddCopySheet(int originalSheetIndex, string newSheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet lastSheet = _Excel.Worksheets[_Excel.Worksheets.Count];
            originalSheet.Copy(Type.Missing, lastSheet);
            //ActiveSheetName = newSheetName;
        }
        #endregion

        #region 插入一个Sheet
        public void InsertSheet(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            this._Excel.Worksheets.Add(sheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
        }
        #endregion

        #region 插入一个Sheet
        public void InsertSheet(int index, string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            this._Excel.Worksheets.Add(sheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            //ActiveSheetName = sheetName;
        }
        #endregion

        #region 插入一个复制Sheet
        public void InsertCopySheet(int originalSheetIndex, int destinationSheetIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet destinationSheet = _Excel.Worksheets[destinationSheetIndex];
            originalSheet.Copy(destinationSheet, Type.Missing);
        }
        #endregion

        #region 插入一个复制Sheet
        public void InsertCopySheet(int originalSheetIndex, int destinationSheetIndex, string newSheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalSheetIndex];
            Microsoft.Office.Interop.Excel.Worksheet destinationSheet = _Excel.Worksheets[destinationSheetIndex];
            originalSheet.Copy(destinationSheet, Type.Missing);
            //ActiveSheetName = newSheetName;
        }
        #endregion

        #region 移动Sheet
        public void MoveSheet(int originalIndex, int destinationIndex)
        {
            if (originalIndex == destinationIndex)
            { return; }
            Microsoft.Office.Interop.Excel.Worksheet originalSheet = _Excel.Worksheets[originalIndex];
            Microsoft.Office.Interop.Excel.Worksheet destinationSheet = _Excel.Worksheets[destinationIndex];
            if (originalIndex < destinationIndex)
            {
                destinationSheet = _Excel.Worksheets[destinationIndex + 1];
            }
            originalSheet.Move(destinationSheet, Type.Missing);
        }
        #endregion

        #region 删除Sheet
        public void DeleteSheet(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[index];
            sheet.Delete();
        }
        #endregion

        #region 删除Sheet
        public void DeleteSheet(string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.Worksheets[GetSheetIndex(sheetName)];
            sheet.Delete();
        }
        #endregion

        #region 获取Sheet中已使用的行数
        public int GetUsedRowsCount()
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.ActiveSheet;
            return sheet.UsedRange.Rows.Count;
        }
        #endregion

        #region 获取Sheet中已使用的列数
        public int GetUsedColumnCount(int sheetIndex = 0)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.ActiveSheet;
            return sheet.UsedRange.Columns.Count;
        }
        #endregion

        #region 设置活动单元格
        public void SetActiveCell(int rowIndex, int columnIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet sheet = _Excel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range range = sheet.Cells[rowIndex, columnIndex];
            range.Select();
            return;
        }
        #endregion

        #region 获取单元格数据
        public object GetCellValue(int rowIndex, int columnIndex)
        {
            Microsoft.Office.Interop.Excel.Worksheet activeSheet = _Excel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range range = activeSheet.Cells[rowIndex, columnIndex];
            return range.Value;
        }
        #endregion

        #region 设置单元格数据
        public void SetCellValue(int rowIndex, int columnIndex, object value)
        {
            Microsoft.Office.Interop.Excel.Worksheet activeSheet = _Excel.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range range = activeSheet.Cells[rowIndex, columnIndex];
            range.Value = value;
        }
        #endregion

        #region 获取行数据
        public object[] GetRowValues(int rowIndex)
        {
            int columnCount = GetUsedColumnCount();
            Microsoft.Office.Interop.Excel.Range range = GetRange(rowIndex, 1, rowIndex, columnCount);
            object[,] allValues = range.Value;
            object[] values = new object[columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                values[i] = allValues[1, i + 1];
                if (values[i] == null)
                {
                    values[i] = "";
                }
            }
            return values;
        }
        #endregion

        #region 获取行数据
        public object[] GetRowValues(int rowIndex, int columnIndex, int columnCount)
        {
            Microsoft.Office.Interop.Excel.Range range = GetRange(rowIndex, columnIndex, rowIndex, columnIndex + columnCount - 1);
            object[,] allValues = range.Value;
            object[] values = new object[columnCount];
            for (int i = 0; i < columnCount; i++)
            {
                values[i] = allValues[1, i + 1];
                if (values[i] == null)
                {
                    values[i] = "";
                }
            }
            return values;
        }
        #endregion


        #endregion

    }
}
