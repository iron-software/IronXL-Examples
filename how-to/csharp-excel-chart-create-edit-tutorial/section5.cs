using IronXL;
using IronXL.Drawing.Charts;
using System.Collections.Generic;

WorkBook workBook = WorkBook.Load("pieChart.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve the chart
List<IChart> chart = workSheet.Charts;

// Remove the chart
workSheet.RemoveChart(chart[0]);

workBook.SaveAs("removedChart.xlsx");