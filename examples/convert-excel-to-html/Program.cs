using IronXL;
using IronXL.Options;

WorkBook workBook = WorkBook.Load("sample.xlsx");

var options = new HtmlExportOptions()
{
    // Set row/column numbers visible in html document
    OutputRowNumbers = true,
    OutputColumnHeaders = true,

    // Set hidden rows/columns visible in html document
    OutputHiddenRows = true,
    OutputHiddenColumns = true,

    // Set leading spaces as non-breaking
    OutputLeadingSpacesAsNonBreaking = true
};

// Export workbook to the HTML file
workBook.ExportToHtml("workBook.html", options);
