using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpEditExcelFile
{
    public static class Section9
    {
        public static void Run()
        {
            /**
            Remove Worksheet from File
            anchor-remove-worksheet-from-excel-file
            **/
            wb.RemoveWorkSheet(1); // by sheet indexing
        }
    }
}