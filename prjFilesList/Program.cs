

namespace prjFilesList
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<string> sheetNames = new List<string>{ "ADMIN", "API", "BATCH", "SFTP" };
            string folderPath = @"D:\作業\進管程式清單小作業\進管\";
            string targetPath = @"S:\進管\";
            CCreateFileList fileList = new CCreateFileList(sheetNames, folderPath, targetPath);
            fileList.CreateDataList();
            fileList.CreateListFile();
        }
    }
}