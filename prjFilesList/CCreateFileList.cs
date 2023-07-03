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
        private List<string> _SheetNames { get; }
        private string _FolderPath { get; }
        private string _TargetPath { get; }

        private string[] _Columns { get; }

        private List<CSheetData> Datas = new List<CSheetData>();
        public CCreateFileList(List<string> sheetNames, string inputFolderPath, string targetPath)
        {
            _SheetNames = sheetNames;
            _FolderPath = inputFolderPath;
            _TargetPath = targetPath;
            _Columns = new string[4] { "異動類型", "程式名稱", "附件zip檔案解壓縮後資料夾名稱", "目的程式路徑" };
        }

        //取得檔案清單
        public void CreateDataList()
        {
            foreach (string sheet in _SheetNames)
            {
                DirectoryInfo directoryInfo = new DirectoryInfo(_FolderPath + sheet);
                FileInfo[] rowFiles = directoryInfo.GetFiles("*", SearchOption.AllDirectories);
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
                            FileFullPath = file.FullName.Replace(_FolderPath, _TargetPath),   //將原檔案路徑改為輸出的目的路徑
                        };
                        sheetData.FileDatas.Add(fileData);
                        Console.WriteLine(file.FullName);
                    }
                }
                Datas.Add(sheetData);
            }
        }
        //匯出成excel
        public void CreateListFile()
        {
            if (Datas.Count() == 0)
            {
                Console.WriteLine("目標資料夾無檔案");
                Console.ReadLine();
            }
            XSSFWorkbook workbook = new XSSFWorkbook();
            foreach (var sheetData in Datas)
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
                    fileRow.CreateCell(2).SetCellValue(sheetData.FileDatas[i].ZipFolderName);
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
