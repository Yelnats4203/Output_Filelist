using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace prjFilesList.Models
{
    public class CFileData
    {
        [Display(Name = "異動類型")]
        public string Type { get; set; } = "type";

        [Display(Name = "程式名稱")]
        public string FileName { get; set; } = "";

        [Display(Name = "附件zip檔案解壓縮後資料夾名稱")]
        public string ZipFolderName { get; set; } = "";

        [Display(Name = "目的程式路徑")]
        public string FileFullPath { get; set; } = "";

    }
}
