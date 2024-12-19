using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section15
    {
        public static void Run()
        {
            for (var i = 2; i < countries.Count; i++)
            {
                var country = countries[i];
                //Set the basic values
                workSheet[$"A{i}"].Value = country.name;
                workSheet[$"B{i}"].Value = country.population;
                workSheet[$"G{i}"].Value = country.region;
                workSheet[$"H{i}"].Value = country.numericCode;
                //Iterate through languages
                for (var x = 0; x < 3; x++)
                {
                    if (x > (country.languages.Count - 1)) break;
                    var language = country.languages[x];
                    //Get the letter for the column
                    var columnLetter = GetColumnLetter(4 + x);
                    //Set the language name
                    workSheet[$"{columnLetter}{i}"].Value = language.name;
                }
            }
        }
    }
}