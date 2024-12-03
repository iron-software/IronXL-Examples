using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section15
    {
        public void Run()
        {
            public async Task ProcessAsync()
            {
                //Get the first worksheet
                var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
                var worksheet = workbook.GetWorkSheet("GDPByCountry");
                //Create the database connection
                using (var countryContext = new CountryContext())
                {
                    //Iterate through all the cells
                    for (var i = 2; i <= 213; i++)
                    {
                        //Get the range from A-B
                        var range = worksheet[$"A{i}:B{i}"].ToList();
                        //Create a Country entity to be saved to the database
                        var country = new Country
                        {
                            Name = (string)range[0].Value,
                            GDP = (decimal)(double)range[1].Value
                        };
                        //Add the entity
                        await countryContext.Countries.AddAsync(country);
                    }
                    //Commit changes to the database
                    await countryContext.SaveChangesAsync();
                }
            }
        }
    }
}