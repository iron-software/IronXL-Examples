using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section5
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook wb = IronXL.WorkBook.Load("sample.xlsx");

            WorkSheet ws = wb.WorkSheets.First();//for the first sheet:
        }
    }
}