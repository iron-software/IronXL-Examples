using System.Collections.Generic;
using IronXL.Excel;
namespace ironxl.CsharpExcelChartCreateEditTutorial
{
    public class Section5
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("pieChart.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Retrieve the chart
            List<IChart> chart = workSheet.Charts;
            
            // Remove the chart
            workSheet.RemoveChart(chart[0]);
            
            workBook.SaveAs("removedChart.xlsx");
        }
    }
}