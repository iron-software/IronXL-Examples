using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateXlsxFileCSharp
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
        }
    }
}