using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section6
    {
        public void Run()
        {
            // Set a formula
            workSheet["A1"].Value = "=SUM(A2:A10)";
            
            // Get the calculated value
            decimal sum = workSheet["A1"].DecimalValue;
        }
    }
}