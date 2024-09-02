IronXL provides the capability to programmatically configure **Print Setup** for any Excel file. This feature offers detailed control over numerous settings that dictate how documents are printed, whether on physical paper or as a PDF.

Additionally, users can customize document headers and footers, incorporating dynamic "mail merge" variables such as:

- `&P`: indicates page numbers
- `&N`: represents the total number of pages
- `&D`: denotes the current date
- `&T`: represents the current time
- `&Z&F`: captures the file path
- `&F`: refers to the file name
- `&A`: denotes the sheet name

These variables can be seamlessly integrated into the `Footer` property. For example, to place the page number at the bottom center of every printed page, one could use the following configuration: `workSheet.Footer.Center = "Page &P of &N"`.

This functionality positions IronXL as a comprehensive tool for managing how spreadsheets are printed.