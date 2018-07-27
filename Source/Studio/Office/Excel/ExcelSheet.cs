using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Studio.Office.Excel
{
    public class ExcelSheet
    {
        private ExcelDocument _Parent;
        private int _Index;
        private bool _Actived;

        public ExcelDocument Parent { get => _Parent; }
        public bool Actived { get => _Actived; }
        public int Index { get => _Index;}

        internal ExcelSheet(ExcelDocument parent)
        {
            _Parent = parent;
            _Index = 1;
            _Actived = false;
        }

        
    }
}
