using IronXL.Styles;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.NamedTable
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Add data
            workSheet["A2:C5"].StringValue = "Text";
            
            // Configure named table
            var selectedRange = workSheet["A1:C5"];
            bool showFilter = false;
            var tableStyle = TableStyle.TableStyleDark1;
            
            // Add named table
            workSheet.AddNamedTable("table1", selectedRange, showFilter, tableStyle);
            
            workBook.SaveAs("addNamedTable.xlsx");
        }
    }
}