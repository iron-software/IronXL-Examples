using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section2
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            /**
            Access Sheet by Name
            anchor-access-specific-worksheet
            **/
            WorkSheet ws = wb.GetWorkSheet("Sheet1"); //by sheet name
        }
    }
}