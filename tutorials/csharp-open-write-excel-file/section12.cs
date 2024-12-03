using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section12
    {
        public void Run()
        {
            workBook.SaveAsCsv($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.csv", delimiter: "|");
        }
    }
}