using IronXL;
namespace IronXL.Examples.HowTo.CSharpOpenExcelWorksheet
{
    public static class Section3
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            /**
            Open Excel Worksheet
            anchor-open-excel-worksheet
            **/
            //by sheet index
            WorkSheet ws = wb.WorkSheets [0];
            //for the default
            WorkSheet defaultSheet = wb.DefaultWorkSheet;
            //for the first sheet: 
            WorkSheet firstSheet = wb.WorkSheets.First();
            //for the first or default sheet:
            WorkSheet firstOrDefaultSheet = wb.WorkSheets.FirstOrDefault();
        }
    }
}