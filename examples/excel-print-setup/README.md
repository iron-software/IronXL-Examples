***Based on <https://ironsoftware.com/examples/excel-print-setup/>***

IronXL provides the capability to programmatically define the **Print Setup** parameters for any Excel document. This allows developers to adjust a wide array of printing options for both physical and PDF printers.

Furthermore, you can customize document headers and footers, and even integrate dynamic "mail merge" variables:

- `&P`: page numbers
- `&N`: total page count
- `&D`: the current date
- `&T`: the current time
- `&Z&F`: full file path
- `&F`: file name
- `&A`: name of the worksheet

These variables are readily employable in the `Footer` property string. For example, to show the page number at the bottom of each printed page, use the following setup: `workSheet.Footer.Center = "Page &P of &N"`.

This capability empowers IronXL to efficiently manage printing configurations for spreadsheets.