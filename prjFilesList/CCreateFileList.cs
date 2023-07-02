using Org.BouncyCastle.Math.EC;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace prjFilesList
{
    public class CCreateFileList
    {
        private string[] _sheetNames { get; }
        private string _folderPath { get; }
        private string _outputPath { get; }

        private string[] _columns { get; }

        private List<CSheetData> datas = new List<CSheetData>();
        public CCreateFileList()
        {
            _sheetNames = new string[4] { "ADMIN", "API", "BATCH", "SFTP" };
            _folderPath = "D:\\作業\\進管程式清單小作業\\進管\\";
            _outputPath = "S:\\進管\\";
            _columns = new string[4] { "異動類型", "程式名稱", "附件zip檔案解壓縮後資料夾名稱", "目的程式路徑" };
            createDataList();
            createListFile();
        }

        //取得檔案清單
        public void createDataList()
        {
            foreach (string sheet in _sheetNames)
            {
                DirectoryInfo tabInfo = new DirectoryInfo(_folderPath + sheet);
                FileInfo[] rowFiles = tabInfo.GetFiles("*", SearchOption.AllDirectories);
                CSheetData sheetData = new CSheetData()
                {
                    sheetName = sheet,
                    fileCount = rowFiles.Count(),
                };
                if (rowFiles.Count() != 0)
                {
                    List<FileInfo> files = rowFiles.OrderBy(file => file.Name).ToList();

                    foreach (FileInfo file in files)
                    {
                        CFileData fileData = new CFileData()
                        {
                            type = "新增",
                            fileName = file.Name,
                            zipFolderName = "",
                            fileFullPath = file.FullName.Replace(_folderPath, _outputPath),   //將原檔案路徑改為輸出的目的路徑
                        };
                        sheetData.fileDatas.Add(fileData);
                        Console.WriteLine(file.FullName);
                    }
                }
                datas.Add(sheetData);
            }
        }
        //匯出成excel
        public void createListFile()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            foreach (var sheetData in datas)
            {
                
                var sheet = workbook.CreateSheet(sheetData.sheetName);

                //header
                var header = sheet.CreateRow(0);
                for (int i = 0; i < _columns.Length; i++)
                {
                    header.CreateCell(i).SetCellValue(_columns[i]);
                }

                //data body
                for (int i = 0; i < sheetData.fileCount; i++)
                {
                    var fileRow = sheet.CreateRow(i + 1);
                    fileRow.CreateCell(0).SetCellValue(sheetData.fileDatas[i].type);
                    fileRow.CreateCell(1).SetCellValue(sheetData.fileDatas[i].fileName);
                    fileRow.CreateCell(2).SetCellValue("");
                    fileRow.CreateCell(3).SetCellValue(sheetData.fileDatas[i].fileFullPath);
                }

                //file count
                var fileCount = sheet.CreateRow(sheet.LastRowNum + 1);
                fileCount.CreateCell(_columns.Length - 1).SetCellValue($"程式數量：{sheetData.fileCount}");

            }

            //create excel file
            using (var fileStream = File.Create(_folderPath + "(作業)檔案清單.xlsx"))
            {
                workbook.Write(fileStream);
            }
        }
    }



}
