using IronXL.Excel;
namespace ironxl.CreateExcelFileNet
{
    public class Section4
    {
        public void Run()
        {
            workSheet["A1"].Value = "January";
            workSheet["B1"].Value = "February";
            workSheet["C1"].Value = "March";
            workSheet["D1"].Value = "April";
            workSheet["E1"].Value = "May";
            workSheet["F1"].Value = "June";
            workSheet["G1"].Value = "July";
            workSheet["H1"].Value = "August";
            workSheet["I1"].Value = "September";
            workSheet["J1"].Value = "October";
            workSheet["K1"].Value = "November";
            workSheet["L1"].Value = "December";
        }
    }
}