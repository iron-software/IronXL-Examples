using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section4
    {
        public static void Run()
        {
            /**
            Create Csharp WorkSheets 
            anchor-c-num-create-excel-workbook
            **/
            WorkSheet ws1 = wb.CreateWorkSheet("Sheet1");
            WorkSheet ws2 = wb.CreateWorkSheet("Sheet2");
        }
    }
}