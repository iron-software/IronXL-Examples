using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpCreateExcel
{
    public static class Section7
    {
        public static void Run()
        {
            /**
            Save Excel File
            anchor-save-excel-file
            **/
            WorkBook.SaveAs("Path + Filename");
        }
    }
}