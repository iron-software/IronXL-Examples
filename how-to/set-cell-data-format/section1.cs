using System.Linq;
using IronXL.Excel;
namespace ironxl.SetCellDataFormat
{
    public class Section1
    {
        public void Run()
        {
            // Create a new workbook
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Set the data format to 12300.00%
            workSheet["A1"].Value = 123;
            workSheet["A1"].FormatString = BuiltinFormats.Percent2;
            
            // Set the data format to 123.0000
            workSheet["A2"].Value = 123;
            workSheet["A2"].FormatString = "0.0000";
            
            // Set data display format to range
            DateTime dateValue = new DateTime(2020, 1, 1, 12, 12, 12);
            workSheet["A3"].Value = dateValue;
            workSheet["A4"].Value = new DateTime(2022, 3, 3, 10, 10, 10);
            workSheet["A5"].Value = new DateTime(2021, 2, 2, 11, 11, 11);
            
            IronXL.Range range = workSheet["A3:A5"];
            
            // Set the data format to 1/1/2020 12:12:12
            range.FormatString = "MM/dd/yy h:mm:ss";
            
            workBook.SaveAs("dataFormats.xlsx");
        }
    }
}