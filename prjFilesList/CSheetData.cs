using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace prjFilesList
{
    public class CSheetData
    {
        public string sheetName { get; set; } = "tab";
        public List<CFileData> fileDatas { get; set; } = new List<CFileData>();

        public int fileCount { get; set; } = 0;
    }
}
