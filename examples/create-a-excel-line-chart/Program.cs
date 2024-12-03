using IronXL;
using IronXL.Drawing.Charts;

WorkBook workBook = WorkBook.Load("test.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Set the chart type and it's position  on the worksheet.
var chart = workSheet.CreateChart(ChartType.Line, 10, 10, 18, 20);

// Add the series to the chart
// The first parameter represents the address of the range for horizontal(category) axis.
// The second  parameter represents the address of the range for vertical(value) axis.
var series = chart.AddSeries("B3:B8", "A3:A8");

// Set the chart title.
series.Title = "Line Chart";

// Set the legend position.
// Can be removed by setting it to None.
chart.SetLegendPosition(LegendPosition.Bottom);

// We can change the position of the chart.
chart.Position.LeftColumnIndex = 2;
chart.Position.RightColumnIndex = chart.Position.LeftColumnIndex + 3;

// Plot all the data that was added to the chart before.
// Multiple call of this method leads to plotting multiple charts instead of modifying the existing chart.
// Yet there is no possibility to remove chart or edit it's series/position.
// We can just create new one.
chart.Plot();

workBook.SaveAs("CreateLineChart.xlsx");
