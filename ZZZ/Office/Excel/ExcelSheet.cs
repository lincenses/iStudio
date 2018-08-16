using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office.Excel
{
    public class ExcelSheet
    {
        internal Microsoft.Office.Interop.Excel.Worksheet XlWorkSheet;

        internal Microsoft.Office.Interop.Excel.Sheets XlSheets;

        internal ExcelSheet(Microsoft.Office.Interop.Excel.Worksheet xlWorksheet)
        {
            XlWorkSheet = xlWorksheet;
            XlSheets = ((Microsoft.Office.Interop.Excel.Workbook)xlWorksheet.Parent).Worksheets;
        }

        public int Index { get => XlWorkSheet.Index; }

        public string Name { get => XlWorkSheet.Name; set => XlWorkSheet.Name = value; }

        public bool Activated
        {
            get => XlSheets.Application.ActiveSheet == XlWorkSheet;
            set => XlWorkSheet.Activate();
        }

        public bool Visible
        {
            get => XlWorkSheet.Visible == Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible;
            set => XlWorkSheet.Visible = value ? Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible : Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;
        }

        #region 保存文件
        public void SaveAsCSV(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            XlWorkSheet.SaveAs(fullName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        #endregion

        #region 保存Excel文件（PDF格式）
        public void SaveAsPDF(string fileName)
        {
            string fullName = new System.IO.FileInfo(fileName).FullName;
            XlWorkSheet.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, fullName, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, false, Type.Missing);
        }
        #endregion

        public void CopyTo(int index)
        {
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet = XlSheets[1];
            string xlSheetName = xlWorksheet.Name;

            if (index > XlSheets.Count)
            {
                

                //XlWorkSheet.Copy(Type.Missing, Parent.LastSheet.XlWorkSheet);
                //Microsoft.Office.Interop.Excel.Worksheet newSheet = Parent.XlWorksheets[Parent.XlWorksheets.Count];
                //Parent.ExcelSheetList.Add(new ExcelSheet(newSheet, Parent));
            }
            else
            {
                Microsoft.Office.Interop.Excel.Workbook xlWorksheets = XlWorkSheet.Parent;
                

                //ExcelSheet beforeExcelSheet = Parent[index];
                //XlWorkSheet.Copy(beforeExcelSheet.XlWorkSheet, Type.Missing);
                //Microsoft.Office.Interop.Excel.Worksheet newSheet = Parent.XlWorksheets[index + 1];
                //Parent.ExcelSheetList.Insert(index - 1, new ExcelSheet(newSheet, Parent));
            }
        }

        public void CopyTo(int index, string sheetName)
        {
            //if (index == Parent.Count + 1)
            //{
            //    XlWorkSheet.Copy(Type.Missing, Parent.LastSheet.XlWorkSheet);
            //    Microsoft.Office.Interop.Excel.Worksheet newSheet = Parent.XlWorksheets[Parent.XlWorksheets.Count];
            //    newSheet.Name = sheetName;
            //    //Parent.ExcelSheetList.Add(new ExcelSheet(newSheet, Parent));
            //}
            //else
            //{
            //    ExcelSheet beforeExcelSheet = Parent[index];
            //    XlWorkSheet.Copy(beforeExcelSheet.XlWorkSheet, Type.Missing);
            //    Microsoft.Office.Interop.Excel.Worksheet newSheet = Parent.XlWorksheets[index + 1];
            //    newSheet.Name = sheetName;
            //    //Parent.ExcelSheetList.Insert(index - 1, new ExcelSheet(newSheet, Parent));
            //}
        }

    }
}
