using IronXL.Excel;
namespace IronXL.Examples.HowTo.BlazorReadExcelFileTutorial
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
                                    @row [column.ColumnName].ToString()
                                </td>
                            }
                        </tr>
                    }
                </tbody>
            </table>
            
            @code {
                // Create a DataTable
                private DataTable displayDataTable = new DataTable();
            
                // When a file is uploaded to the App using the InputFile, trigger:
                async Task OpenExcelFileFromDisk(InputFileChangeEventArgs e)
                {
                    IronXL.License.LicenseKey = "PASTE TRIAL OR LICENSE KEY";
            
                    // Open the File to a MemoryStream object
                    MemoryStream ms = new MemoryStream();
            
                    await e.File.OpenReadStream().CopyToAsync(ms);
                    ms.Position = 0;
            
                    // Define variables for IronXL
                    WorkBook loadedWorkBook = WorkBook.FromStream(ms);
                    WorkSheet loadedWorkSheet = loadedWorkBook.DefaultWorkSheet; // Or use .GetWorkSheet()
            
                    // Add header Columns to the DataTable
                    RangeRow headerRow = loadedWorkSheet.GetRow(0);
                    for (int col = 0 ; col < loadedWorkSheet.ColumnCount ; col++)
                    {
                        displayDataTable.Columns.Add(headerRow.ElementAt(col).ToString());
                    }
            
                    // Populate the DataTable
                    for (int row = 1 ; row < loadedWorkSheet.RowCount ; row++)
                    {
                        IEnumerable<string> excelRow = loadedWorkSheet.GetRow(row).ToArray().Select(c => c.ToString());
                        displayDataTable.Rows.Add(excelRow.ToArray());
                    }
                }
            }
        }
    }
}