using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section13
    {
        public static void Run()
        {
            for (var i = 2; i < countries.Count; i++)
            {
                var country = countries [i];
            
                //Set the basic values
                worksheet [$"A{i}"].Value = country.name;
                worksheet [$"B{i}"].Value = country.population;
                worksheet [$"G{i}"].Value = country.region;
                worksheet [$"H{i}"].Value = country.numericCode;
            
                //Iterate through languages
                for (var x = 0; x < 3; x++)
                {
                    if (x > (country.languages.Count - 1)) break;
            
                    var language = country.languages [x];
            
                    //Get the letter for the column
                    var columnLetter = GetColumnLetter(4 + x);
            
                    //Set the language name
                    worksheet [$"{columnLetter}{i}"].Value = language.name;
                }
            }
        }
    }
}