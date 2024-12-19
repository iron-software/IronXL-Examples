using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section9
    {
        public static void Run()
        {
            /**
            Use Formulas in Cells
            anchor-use-formulas-in-cells
            **/
            decimal sum = sheet ["A2:A11"].Sum();
            decimal avg = sheet ["B2:B11"].Avg();
            decimal max = sheet ["C2:C11"].Max();
            decimal min = sheet ["D2:D11"].Min();
            
            sheet ["A12"].Value = sum;
            sheet ["B12"].Value = avg;
            sheet ["C12"].Value = max;
            sheet ["D12"].Value = min;
        }
    }
}