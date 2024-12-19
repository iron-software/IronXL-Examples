using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section4
    {
        public static void Run()
        {
            newXLFile.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv",delimiter:"|");
        }
    }
}