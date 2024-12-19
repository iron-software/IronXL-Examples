using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section12
    {
        public static void Run()
        {
            /**
            Save Workbook
            anchor-save-workbook
            **/
            workbook.SaveAs("Budget.xlsx");
        }
    }
}