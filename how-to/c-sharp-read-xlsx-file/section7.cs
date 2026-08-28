using IronXL;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section7
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet ws = workBook.DefaultWorkSheet;

            string c = ws ["cell address"].ToString(); //for string
            Int32 val = ws ["cell address"].Int32Value; //for integer
        }
    }
}