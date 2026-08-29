using IronXL;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section9
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            /**
            Remove Worksheet from File
            anchor-remove-worksheet-from-excel-file
            **/
            wb.RemoveWorkSheet(1); // by sheet indexing
        }
    }
}