using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadExcelFileExample
{
    public static class Section7
    {
        public static void Run()
        {
            WorkBook wb = WorkBook.Load("Book1.xlsx");
            WorkSheet ws = wb.GetWorkSheet("Sheet1");
            
            //Traverse all rows of Excel WorkSheet
            foreach(RangeRow row in ws.Rows)
            {
                //Traverse all columns of specific Row
                foreach(Cell cell in row)
                {
                    //Get the values
                    string val = cell.StringValue;
                    int currentRow = cell.RowIndex + 1;
                    int currentCol = cell.ColumnIndex + 1;
            
                    Console.WriteLine($"Value of Row {currentRow} and Column {currentCol} is: {val}");
                }
            }
        }
    }
}