using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section21
    {
        public static void Run()
        {
            WorkBook workBook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            int i = 1;
            foreach (var cell in workSheet["B1:B4"])
            {
                cell.Formula = "=IF(A" + i + ">=20,\" Pass\" ,\" Fail\" )";
                i++;
            }
            workBook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
        }
    }
}