using IronXL;
namespace IronXL.Examples.Example.ExcelFormulasCsharp
{
    public static class Section1
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            workSheet ["A2"].Formula = "=SQRT(A1)";
            workSheet ["B8"].Formula = "=C9/C11";
            workSheet ["G31"].Formula = "=TAN(G30)";
        }
    }
}