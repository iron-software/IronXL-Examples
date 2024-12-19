using IronXL.Excel;
namespace IronXL.Examples.Example.ExcelFormulasCsharp
{
    public static class Section1
    {
        public static void Run()
        {
            workSheet ["A2"].Formula = "=SQRT(A1)"
            workSheet ["B8"].Formula = "=C9/C11"
            workSheet ["G31"].Formula = "=TAN(G30)"
        }
    }
}