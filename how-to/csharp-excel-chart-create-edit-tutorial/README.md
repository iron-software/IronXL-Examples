# Creating and Modifying Excel Charts Using C# with IronXL

***Based on <https://ironsoftware.com/how-to/csharp-excel-chart-create-edit-tutorial/>***


Excel charts provide a powerful way to visually represent data, making it easier to understand and analyze. Excel offers a variety of chart types such as bar, line, pie, and others, each designed for specific data presentations.

IronXL supports a rich set of chart types including column, scatter, line, pie, bar, and area charts. These can be customized with different series names, legend placements, titles, and positions on a worksheet.

### Getting Started with IronXL

---

## Example: Creating Charts

IronXL enables you to create various charts easily. Here are the steps to create a chart:

1. Initiate with the `CreateChart` method, specifying the chart type and location in the worksheet.
2. Add data series using the `AddSeries` method, which can accept data ranges for both the horizontal and vertical axes.
3. Optionally, you can set the series name, chart title, and the position of the legend.
4. Use the `Plot` method to render the chart. Each invocation of this method creates a new chart on the worksheet instead of modifying an existing one.

Below is an example based on data from the [chart.xlsx](https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/chart.xlsx) file:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/data.webp" alt="Data" class="img-responsive add-shadow">
    </div>
</div>

### Creating a Column Chart

```cs
using IronXL;
using IronXL.Drawing.Charts;

WorkBook workbook = WorkBook.Load("chart.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Determine chart type and specify its position on the worksheet
IChart chart = worksheet.CreateChart(ChartType.Column, 5, 5, 20, 10);

string xAxisDataRange = "A2:A7";

// Add series to the chart
IChartSeries series = chart.AddSeries(xAxisDataRange, "B2:B7");
series.Title = worksheet["B1"].StringValue;

series = chart.AddSeries(xAxisDataRange, "C2:C7");
series.Title = worksheet["C1"].StringValue;

series = chart.AddSeries(xAxisDataRange, "D2:D7");
series.Title = worksheet["D1"].StringValue;

// Configure chart title and legend placement
chart.SetTitle("Column Chart");
chart.SetLegendPosition(LegendPosition.Bottom);

// Render the chart
chart.Plot();

workbook.SaveAs("columnChart.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/column-chart.webp" alt="Column chart" class="img-responsive add-shadow">
    </div>
</div>

### Creating a Line Chart

Switching to a line chart from a column chart is straightforward by merely changing the chart type.

```cs
using IronXL;
using IronXL.Drawing.Charts;

WorkBook workbook = WorkBook.Load("chart.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Define chart type and its location
IChart chart = worksheet.CreateChart(ChartType.Line, 5, 5, 20, 10);

// Adding data series
series = chart.AddSeries("A2:A7", "B2:B7");
series.Title = worksheet["B1"].StringValue;

series = chart.AddSeries("A2:A7", "C2:C7");
series.Title = worksheet["C1"].StringValue;

series = chart.AddSeries("A2:A7", "D2:D7");
series.Title = worksheet["D1"].StringValue;

// Set chart properties
chart.SetTitle("Line Chart");
chart.SetLegendPosition(LegendPosition.Bottom);

// Display the chart
chart.Plot();

workbook.SaveAs("lineChart.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/line-chart.webp" alt="Line chart" class="img-responsive add-shadow">
    </div>
</div>

### Creating a Pie Chart

A pie chart typically requires only one set of data.

```cs
using IronXL;
using IronXL.Drawing.Charts;

WorkBook workbook = WorkBook.Load("chart.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Define the pie chart type and position
IChart chart = worksheet.CreateChart(ChartType.Pie, 5, 5, 20, 10);

// Add a single data series
series = chart.AddSeries("A2:A7", "B2:B7");
series.Title = worksheet["B1"].StringValue;

// Configure the chart
chart.SetTitle("Pie Chart");
chart.SetLegendPosition(LegendPosition.Bottom);

// Render the pie chart
chart.Plot();

workbook.SaveAs("pieChart.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/pie-chart.webp" alt="Pie chart" class="img-responsive add-shadow">
    </div>
</div>

---

## Example: Modifying an Existing Chart

You can update various properties of an existing chart, such as the legend position and the chart title.

```cs
using IronXL;
using IronXL.Drawing.Charts;

WorkBook workbook = WorkBook.Load("pieChart.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Retrieve the existing chart
IChart chart = worksheet.Charts[0];

// Modify the legend position and the chart title
chart.SetLegendPosition(LegendPosition.Top);
chart.SetTitle("Updated Pie Chart");

workbook.SaveAs("updatedChart.xlsx");
```

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 48%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/before.webp" alt="Before" class="img-responsive add-shadow" >
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">Before Updated</p>
    </div>
    <div class="competitors__card" style="width: 49%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/create-edit-charts/after.webp" alt="After" class="img-responsive add-shadow">
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">After Updated</p>
    </div>
</div>

---

## Example: Removing a Chart from a Spreadsheet

Removal of charts is also facilitated through IronXL by fetching the chart from the `Charts` property and then passing it to the `RemoveChart` method.

```cs
using IronXL;
using IronXL.Drawing.Charts;
using System.Collections.Generic;

WorkBook workbook = WorkBook.Load("updatedChart.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Access the list of charts
List<IChart> charts = worksheet.Charts;

// Remove the specified chart
worksheet.RemoveChart(charts[0]);

workbook.SaveAs("cleanWorksheet.xlsx");
```

This comprehensive guide should provide you with detailed insights into creating, modifying, and removing charts in Excel using IronXL in a C# environment.