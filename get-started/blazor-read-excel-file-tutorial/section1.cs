using IronXL.Excel;
namespace IronXL.Examples.GettingStarted.BlazorReadExcelFileTutorial
{
    public static class Section1
    {
        public static void Run()
        {
            @using IronXL;
            @using System.Data;
            
            @page "/fetchdata"
            
            <PageTitle>Excel File Viewer</PageTitle>
            
            <h1>Open Excel File to View</h1>
            
            <InputFile OnChange="@OpenExcelFileFromDisk" />
            
            <table>
                <thead>
                    <tr>
                        @foreach (DataColumn column in displayDataTable.Columns)
                        {
                            <th>
                                @column.ColumnName
                            </th>
                        }
                    </tr>
                </thead>
                <tbody>
                    @foreach (DataRow row in displayDataTable.Rows)
                    {
                        <tr>
                            @foreach (DataColumn column in displayDataTable.Columns)
                            {
                                <td>
                                    @row[column.ColumnName].ToString()
                                </td>
                            }
                        </tr>
                    }
                </tbody>
            </table>
            
            @code {
                // Create a DataTable instance
                private DataTable displayDataTable = new DataTable();
            
                // This method is triggered when a file is uploaded
                async Task OpenExcelFileFromDisk(InputFileChangeEventArgs e)
                {
                    IronXL.License.LicenseKey = "PASTE TRIAL OR LICENSE KEY";
            
                    // Load the uploaded file into a MemoryStream
                    MemoryStream ms = new MemoryStream();
            
                    await e.File.OpenReadStream().CopyToAsync(ms);
                    ms.Position = 0;
            
                    // Create an IronXL workbook from the MemoryStream
                    WorkBook loadedWorkBook = WorkBook.FromStream(ms);
                    WorkSheet loadedWorkSheet = loadedWorkBook.DefaultWorkSheet; // Or use .GetWorkSheet()
            
                    // Add header Columns to the DataTable
                    RangeRow headerRow = loadedWorkSheet.GetRow(0);
                    for (int col = 0; col < loadedWorkSheet.ColumnCount; col++)
                    {
                        displayDataTable.Columns.Add(headerRow.ElementAt(col).ToString());
                    }
            
                    // Populate the DataTable with data from the Excel sheet
                    for (int row = 1; row < loadedWorkSheet.RowCount; row++)
                    {
                        IEnumerable<string> excelRow = loadedWorkSheet.GetRow(row).ToArray().Select(c => c.ToString());
                        displayDataTable.Rows.Add(excelRow.ToArray());
                    }
                }
            }
        }
    }
}