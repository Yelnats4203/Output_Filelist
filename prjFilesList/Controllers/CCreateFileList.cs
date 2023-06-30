using Org.BouncyCastle.Math.EC;
using prjFilesList.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace prjFilesList.Controllers
{
    public class CCreateFileList
    {
        private string[] _tabs { get; }
        private string _folderPath { get; }
        private string _outputPath { get; }

        private string[] _columns { get; }

        private List<CTabData> datas = new List<CTabData>();
        public CCreateFileList()
        {
            _tabs = new string[4] { "ADMIN", "API", "BATCH", "SFTP" };
            _folderPath = "D:\\作業\\進管程式清單小作業\\進管\\";
            _outputPath = "S:\\進管\\";
            _columns = new string[4] {"異動類型", "程式名稱", "附件zip檔案解壓縮後資料夾名稱", "目的程式路徑" };
            createDataList();
            createListFile();
        }

        //取得檔案清單
        public void createDataList()
        {
            foreach (string tab in _tabs)
            {
                DirectoryInfo tabInfo = new DirectoryInfo(_folderPath + tab);
                FileInfo[] files = tabInfo.GetFiles("*", SearchOption.AllDirectories);
                CTabData tabData = new CTabData()
                {
                    tabName = tab,
                    fileCount = files.Count(),
                };
                if (files.Count() != 0)
                {
                    foreach (FileInfo file in files)
                    {
                        CFileData fileData = new CFileData()
                        {
                            type = "新增",
                            fileName = file.Name,
                            zipFolderName = "",
                            fileFullPath = file.FullName.Replace(_folderPath, _outputPath),   //將原檔案路徑改為輸出的目的路徑
                        };
                        tabData.fileDatas.Add(fileData);
                        //Console.WriteLine(file.FullName.Replace(_filesPath, _outputPath));
                    }
                }
                datas.Add(tabData);
            }
        }
        //匯出成excel
        public void createListFile()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            foreach (var tabData in datas)
            {
                workbook.CreateSheet(tabData.tabName);
            }
            using (var fileStream = File.Create(_folderPath + "檔案清單.xlsx"))
            {
                workbook.Write(fileStream);
            }
        }
    }



}
