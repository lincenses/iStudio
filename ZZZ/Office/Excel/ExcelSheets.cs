using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office.Excel
{
    public class ExcelSheets
    {
        private bool _IncludeHide = false;

        private Microsoft.Office.Interop.Excel.Sheets _XlWorksheets = null;
        
        private List<Microsoft.Office.Interop.Excel.Worksheet> _XlWorksheetList = new List<Microsoft.Office.Interop.Excel.Worksheet>();

        internal ExcelSheets(Microsoft.Office.Interop.Excel.Sheets xlWorksheets, bool includeHide)
        {
            _XlWorksheets = xlWorksheets;
            _IncludeHide = includeHide;
            RefreshXlWorksheetsList();
        }

        internal void RefreshXlWorksheetsList()
        {
            _XlWorksheetList = new List<Microsoft.Office.Interop.Excel.Worksheet>();
            for (int i = 1; i <= _XlWorksheets.Count; i++)
            {
                Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = _XlWorksheets[i];
                if (_IncludeHide == false && xlWorksheet.Visible != Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible)
                { continue; }
                _XlWorksheetList.Add(_XlWorksheets[i]);
            }
        }


        public int Count
        { get => _XlWorksheetList.Count; }

        public ExcelSheet this[int index]
        { get => new ExcelSheet(_XlWorksheetList[index - 1]); }

        public ExcelSheet this[string sheetName]
        {
            get
            {
                for (int i = 0; i < _XlWorksheetList.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = _XlWorksheetList[i];
                    if (string.Compare(xlWorksheet.Name, sheetName, true) == 0)
                    {
                        return new ExcelSheet(xlWorksheet);
                    }
                }
                return null;
            }
        }

        public ExcelSheet ActiveSheet
        { get => new ExcelSheet(_XlWorksheets.Application.ActiveSheet); }

        public ExcelSheet FirstSheet
        { get => new ExcelSheet(_XlWorksheetList[0]); }

        public ExcelSheet LastSheet
        { get => new ExcelSheet(_XlWorksheetList[_XlWorksheetList.Count - 1]); }

        public ExcelSheet Add()
        {
            Microsoft.Office.Interop.Excel.Worksheet xlNewWorksheet = _XlWorksheets.Add(Type.Missing, LastSheet.XlWorkSheet, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            RefreshXlWorksheetsList();
            return new ExcelSheet(xlNewWorksheet);
        }

        public ExcelSheet Add(string sheetName)
        {
            Microsoft.Office.Interop.Excel.Worksheet xlNewWorksheet = _XlWorksheets.Add(Type.Missing, LastSheet.XlWorkSheet, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            xlNewWorksheet.Name = sheetName;
            RefreshXlWorksheetsList();
            return new ExcelSheet(xlNewWorksheet);
        }

        public ExcelSheet Insert(int index)
        {
            if (index > Count)
            { return Add(); }
            Microsoft.Office.Interop.Excel.Worksheet xlbeforeWorksheet = _XlWorksheetList[index - 1];
            Microsoft.Office.Interop.Excel.Worksheet xlNewWorksheet = _XlWorksheets.Add(xlbeforeWorksheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            RefreshXlWorksheetsList();
            return new ExcelSheet(xlNewWorksheet);

            //ExcelSheet beforeExcelSheet = ExcelSheetList[index - 1];
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = XlWorksheets.Add(beforeExcelSheet.XlWorkSheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            //ExcelSheet excelSheet = new ExcelSheet(worksheet, this);
            //ExcelSheetList.Insert(index - 1, excelSheet);
            //return excelSheet;
        }

        public ExcelSheet Insert(int index, string sheetName)
        {
            if (index > Count)
            { return Add(sheetName); }
            Microsoft.Office.Interop.Excel.Worksheet xlbeforeWorksheet = _XlWorksheets[index];
            Microsoft.Office.Interop.Excel.Worksheet xlNewWorksheet = _XlWorksheets.Add(xlbeforeWorksheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            xlNewWorksheet.Name = sheetName;
            return new ExcelSheet(xlNewWorksheet);

            //if (index == Count + 1)
            //{ return Add(sheetName); }
            //ExcelSheet beforeExcelSheet = ExcelSheetList[index - 1];
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = XlWorksheets.Add(beforeExcelSheet.XlWorkSheet, Type.Missing, 1, Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);
            //worksheet.Name = sheetName;
            //ExcelSheet excelSheet = new ExcelSheet(worksheet, this);
            //ExcelSheetList.Insert(index - 1, excelSheet);
            //return excelSheet;
        }

    }
}
