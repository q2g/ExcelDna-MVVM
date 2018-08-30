using System.Collections.Generic;

namespace ExcelDna_MVVM.MVVM.ExcelData
{
    class WorkbookData
    {
        public string Name { get; set; }
        public List<int> hwnds { get; set; } = new List<int>();
        public List<string> sheetIds { get; set; } = new List<string>();

    }
}
