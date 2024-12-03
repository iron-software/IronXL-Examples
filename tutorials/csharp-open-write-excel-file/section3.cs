using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\HelloWorld.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            string cell = workSheet["A1"].StringValue;
            Console.WriteLine(cell);
        }
    }
}