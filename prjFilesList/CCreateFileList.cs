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
        private string[] _SheetNames { get; }
        private string _FolderPath { get; }
        private string _OutputPath { get; }

        private string[] _Columns { get; }

        private List<CSheetData> datas = new List<CSheetData>();
        public CCreateFileList()
        {
            _SheetNames = new string[4] { "ADMIN", "API", "BATCH", "SFTP" };
            _FolderPath = @"D:\作業\進管程式清單小作業\進管\";
            _OutputPath = @"S:\進管\";
            _Columns = new string[4] { "異動類型", "程式名稱", "附件zip檔案解壓縮後資料夾名稱", "目的程式路徑" };
            CreateDataList();
            CreateListFile();
        }

        //取得檔案清單
        public void CreateDataList()
        {
            foreach (string sheet in _SheetNames)
            {
                DirectoryInfo tabInfo = new DirectoryInfo(_FolderPath + sheet);
                FileInfo[] rowFiles = tabInfo.GetFiles("*", SearchOption.AllDirectories);
                CSheetData sheetData = new CSheetData()
                {
                    SheetName = sheet,
                    FileCount = rowFiles.Count(),
                };
                if (rowFiles.Count() != 0)
                {
                    List<FileInfo> files = rowFiles.OrderBy(file => file.Name).ToList();

                    foreach (FileInfo file in files)
                    {
                        CFileData fileData = new CFileData()
                        {
                            Type = "新增",
                            FileName = file.Name,
                            ZipFolderName = "",
                            FileFullPath = file.FullName.Replace(_FolderPath, _OutputPath),   //將原檔案路徑改為輸出的目的路徑
                        };
                        sheetData.FileDatas.Add(fileData);
                        Console.WriteLine(file.FullName);
                    }
                }
                datas.Add(sheetData);
            }
        }
        //匯出成excel
        public void CreateListFile()
        {
            XSSFWorkbook workbook = new XSSFWorkbook();
            foreach (var sheetData in datas)
            {
                
                var sheet = workbook.CreateSheet(sheetData.SheetName);

                //header
                var header = sheet.CreateRow(0);
                for (int i = 0; i < _Columns.Length; i++)
                {
                    header.CreateCell(i).SetCellValue(_Columns[i]);
                }

                //data body
                for (int i = 0; i < sheetData.FileCount; i++)
                {
                    var fileRow = sheet.CreateRow(i + 1);
                    fileRow.CreateCell(0).SetCellValue(sheetData.FileDatas[i].Type);
                    fileRow.CreateCell(1).SetCellValue(sheetData.FileDatas[i].FileName);
                    fileRow.CreateCell(2).SetCellValue("");
                    fileRow.CreateCell(3).SetCellValue(sheetData.FileDatas[i].FileFullPath);
                }

                //file count
                var fileCount = sheet.CreateRow(sheet.LastRowNum + 1);
                fileCount.CreateCell(_Columns.Length - 1).SetCellValue($"程式數量：{sheetData.FileCount}");

            }

            //create excel file
            using (var fileStream = File.Create(_FolderPath + "(作業)檔案清單.xlsx"))
            {
                workbook.Write(fileStream);
            }
        }
    }



}
