using IronXL;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section7
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook WorkBook = IronXL.WorkBook.Load("sample.xlsx");

            /**
            Save Excel File
            anchor-save-excel-file
            **/
            WorkBook.SaveAs("Path + Filename");
        }
    }
}