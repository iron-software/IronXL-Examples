using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section14
    {
        public static void Run()
        {
            var client = new Client(new Uri("https://restcountries.eu/rest/v2/"));
            List<RestCountry> countries = await client.GetAsync<List<RestCountry>>();
        }
    }
}