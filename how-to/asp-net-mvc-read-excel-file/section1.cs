using IronXL.Excel;
namespace IronXL.Examples.HowTo.AspNetMvcReadExcelFile
{
    public static class Section1
    {
        public static void Run()
        {
            @{
                ViewData ["Title"] = "Home Page";
            }
            
            @using System.Data
            @model DataTable
            
            <div class="text-center">
                <h1 class="display-4">Welcome to IronXL Read Excel MVC</h1>
            </div>
            <table class="table table-dark">
                <tbody>
                    @foreach (DataRow row in Model.Rows)
                    {
                        <tr>
                            @for (int i = 0; i < Model.Columns.Count; i++)
                            {
                                <td>@row [i]</td>
                            }
                        </tr>
                    }
            
                </tbody>
            </table>
        }
    }
}