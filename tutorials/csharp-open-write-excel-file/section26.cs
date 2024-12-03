using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section26
    {
        public void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\testFile.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet2");
            var range = workSheet["A2:D2"];
            foreach (var cell in range)
            {
                Console.WriteLine(cell.Text);
            }
        }
    }
}