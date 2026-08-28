using IronXL;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section4
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            /**
            Create Csharp WorkSheets 
            anchor-c-num-create-excel-workbook
            **/
            WorkSheet ws1 = wb.CreateWorkSheet("Sheet1");
            WorkSheet ws2 = wb.CreateWorkSheet("Sheet2");
        }
    }
}