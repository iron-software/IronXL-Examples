using IronXL;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section6
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet WorkSheet = workBook.DefaultWorkSheet;
            int RowIndex = 0;
            int ColumnIndex = 0;

            /**
            Import Data by Cell Address
            anchor-import-excel-data-in-c-num
            **/
            //by cell addressing
            string val = WorkSheet ["Cell Address"].ToString();
            //by row and column indexing
            string valByIndex = WorkSheet.Rows [RowIndex].Columns [ColumnIndex].Value.ToString();
        }
    }
}