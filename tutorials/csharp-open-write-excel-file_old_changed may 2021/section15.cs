using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section15
    {
        public static void Run()
        {
            /**
            Condition IF
            anchor-if-condition-example
            **/
            var workbook = IronXL.WorkBook.Load($@"{Directory.GetCurrentDirectory()}\Files\Sum.xlsx");
            var sheet = workbook.WorkSheets.First();
            int i = 1;
            foreach(var cell in sheet ["B1:B4"])
            {
                cell.Formula = "=IF(A" +i+ ">=20,\" Pass\" ,\" Fail\" )";
                i++;
            }
            workbook.SaveAs($@"{Directory.GetCurrentDirectory()}\Files\NewExcelFile.xlsx");
        }
    }
}