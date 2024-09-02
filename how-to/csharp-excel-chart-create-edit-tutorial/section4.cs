using IronXL;
using IronXL.Drawing.Charts;

WorkBook workBook = WorkBook.Load("pieChart.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve the chart
IChart chart = workSheet.Charts[0];

// Edit the legend position
chart.SetLegendPosition(LegendPosition.Top);

// Edit the chart title
chart.SetTitle("Edited Chart");

workBook.SaveAs("editedChart.xlsx");