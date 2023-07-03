using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace prjFilesList.Models
{
    public class CSheetData
    {
        public string SheetName { get; set; } = "sheet";
        public List<CFileData> FileDatas { get; set; } = new List<CFileData>();

        public int FileCount { get; set; } = 0;
    }
}
