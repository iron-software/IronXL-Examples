using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CreateExcelFileNet_Old_Changed May 2021
{
    public static class Section4
    {
        public static void Run()
        {
            /**
            Set Cell Value Manually
            anchor-set-cell-values
            **/
            sheet ["A1"].Value = "January";
            sheet ["B1"].Value = "February";
            sheet ["C1"].Value = "March";
            sheet ["D1"].Value = "April";
            sheet ["E1"].Value = "May";
            sheet ["F1"].Value = "June";
            sheet ["G1"].Value = "July";
            sheet ["H1"].Value = "August";
            sheet ["I1"].Value = "September";
            sheet ["J1"].Value = "October";
            sheet ["K1"].Value = "November";
            sheet ["L1"].Value = "December";
        }
    }
}