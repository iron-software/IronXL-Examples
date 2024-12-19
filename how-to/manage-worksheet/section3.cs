using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ManageWorksheet
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Set active for workSheet3
            workBook.SetActiveTab(2);
            
            workBook.SaveAs("setActiveTab.xlsx");
        }
    }
}