using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Set active for workSheet3
            workBook.SetActiveTab(2);
            
            workBook.SaveAs("setActiveTab.xlsx");
        }
    }
}