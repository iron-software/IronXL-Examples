# How to Programatically Generate Excel Charts in C#

This guide demonstrates how to use C# and IronXL to generate Excel charts programmatically.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Automating Excel Chart Creation in .NET</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-2-create-excel-chart-for-net">Generate Excel charts automatically</a></li>
        <li><a href="#anchor-2-create-excel-chart-for-net">Incorporate series with titles and legends</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h2>Steps to Generate an Excel Chart with C#</h2>

1. Install the Excel library needed for chart generation.
2. Open an existing Excel file into a `Workbook` object.
3. Initiate a chart using `CreateChart`.
4. Assign a title and configure the legend of the chart.
5. Execute the `Plot` method.
6. Export the modified `Workbook` back to an Excel file.

<p class="main-content__segment-title">Step-by-step Guide</p>

## 1. Install IronXL

The simplest method to install IronXL is by utilizing NuGet Package Manager within Visual Studio:

- Navigate to the Project menu
- Choose Manage NuGet Packages
- Search for IronXL.Excel
- Click Install

Alternatively, run this command at the Developer Command Prompt:

```shell
Install-Package IronXL.Excel
```

Or download directly from: <a class="js-modal-open" href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">IronXL Download</a>

<hr class="separator">

<p class="main-content__segment-title">Tutorial Overview</p>

## 2. Create an Excel Chart for .NET

Begin your project by incorporating the following data into an Excel spreadsheet, as shown below:

<div class="content-img-align-center">
  <div class="center-image-wrapper">
  <a
    href="https://ironsoftware.com/img/faq/excel/csharp-create-excel-chart-programmatically/data-to-be-used-for-charting.png"
    target="_blank"
  >
    <img
      class="img-responsive"
      src="https://ironsoftware.com/img/faq/excel/csharp-create-excel-chart-programmatically/data-to-be-used-for-charting.png"
      alt="Data to be used for charting"
    >
  </a>
  <p><strong>Figure 1</strong> – <em>Data to be used for charting</em></p>
  </div>
</div>

Incorporate the necessary namespaces to work with Excel charts in IronXL:

```cs
using IronXL;
using IronXL.Drawing.Charts;
```

Below is a snippet to create the Excel chart programmatically via IronXL:

```cs
private void button1_Click(object sender, EventArgs e)
{
    WorkBook wb = WorkBook.Load("Chart_Ex.xlsx");
    WorkSheet ws = wb.DefaultWorkSheet;

    var chart = ws.CreateChart(ChartType.Column, 10, 15, 25, 20);

    const string xAxis = "A2:A7";

    var series = chart.AddSeries(xAxis, "B2:B7");
    series.Title = ws["B1"].StringValue;

    series = chart.AddSeries(xAxis, "C2:C7");
    series.Title = ws["C1"].StringValue;

    series = chart.AddSeries(xAxis, "D2:D7");
    series.Title = ws["D1"].StringValue;

    chart.SetTitle("Column Chart");
    chart.SetLegendPosition(LegendPosition.Bottom);
    chart.Plot();
    wb.SaveAs("Exported_Column_Chart.xlsx");
}
```

Initial objects for Workbook and Worksheet are instantiated, followed by definition of the chart type and location using `CreateChart`. After adding series with respective titles and setting up the legend, the data is visualized by executing the `.Plot()` method.

<div class="content-img-align-center">
  <div class="center-image-wrapper">
    <img
      class="img-responsive"
      src="https://ironsoftware.com/img/faq/excel/csharp-create-excel-chart-programmatically/chart-output.png"
      alt="Chart output"
    >
    <p><strong>Figure 2</strong> – <em>Chart output depicted</em></p>
  </div>
</div>

<hr class="separator">

<p class="main-content__segment-title">Quick Access to Library</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore IronXL API Documentation</h3>
      <p>Discover more about merging, splitting, and manipulating cells in Excel spreadsheets through the comprehensive IronXL API Documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> IronXL API Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>