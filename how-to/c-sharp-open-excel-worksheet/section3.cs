using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpOpenExcelWorksheet
{
    public static class Section3
    {
        public static void Run()
        {
            /**
            Open Excel Worksheet
            anchor-open-excel-worksheet
            **/
            //by sheet index
            WorkSheet ws = wb.WorkSheets [0];
            //for the default
            WorkSheet ws = wb.DefaultWorkSheet; 
            //for the first sheet: 
            WorkSheet ws = wb.WorkSheets.First();
            //for the first or default sheet:
            WorkSheet ws = wb.WorkSheets.FirstOrDefault();
        }
    }
}