using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office.Excel
{
    public class ExcelSheetCollection
    {
        private ExcelDocument _Parent;

        public ExcelDocument Parent { get => _Parent; }

        internal ExcelSheetCollection(ExcelDocument parent)
        {
            _Parent = parent;
        }
    }
}
