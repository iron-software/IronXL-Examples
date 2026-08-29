using IronXL;
namespace IronXL.Examples.Overview.Quickstart
{
    public static class Section6
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Set a formula
            workSheet["A1"].Value = "=SUM(A2:A10)";
            
            // Get the calculated value
            decimal sum = workSheet["A1"].DecimalValue;
        }
    }
}