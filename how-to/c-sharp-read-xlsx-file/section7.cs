using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpReadXlsxFile
{
    public static class Section7
    {
        public static void Run()
        {
            string c = ws ["cell address"].ToString(); //for string
            Int32 val = ws ["cell address"].Int32Value; //for integer
        }
    }
}