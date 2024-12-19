using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpImportExcel
{
    public static class Section6
    {
        public static void Run()
        {
            /**
            Import Data by Cell Address
            anchor-import-excel-data-in-c-num
            **/
            //by cell addressing
            string val = WorkSheet ["Cell Address"].ToString();
            //by row and column indexing
            string val = WorkSheet.Rows [RowIndex].Columns [ColumnIndex].Value.ToString();
        }
    }
}