using IronXL.Excel;
namespace ironxl.HowToReadExcelFileCsharp
{
    public class Section16
    {
        public void Run()
        {
            var client = new Client(new Uri("https://restcountries.eu/rest/v2/"));
            List<RestCountry> countries = await client.GetAsync<List<RestCountry>>();
        }
    }
}