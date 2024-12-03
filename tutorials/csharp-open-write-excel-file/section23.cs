using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section23
    {
        public void Run()
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