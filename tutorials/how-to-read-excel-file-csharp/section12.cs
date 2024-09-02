var resultsSheet = workBook.CreateWorkSheet("Results");
resultsSheet["A1"].Value = "Row";
resultsSheet["B1"].Value = "Valid";
resultsSheet["C1"].Value = "Phone Error";
resultsSheet["D1"].Value = "Email Error";
resultsSheet["E1"].Value = "Date Error";
for (var i = 0; i < results.Count; i++)
{
    var result = results[i];
    resultsSheet[$"A{i + 2}"].Value = result.Row;
    resultsSheet[$"B{i + 2}"].Value = result.IsValid ? "Yes" : "No";
    resultsSheet[$"C{i + 2}"].Value = result.PhoneNumberErrorMessage;
    resultsSheet[$"D{i + 2}"].Value = result.EmailErrorMessage;
    resultsSheet[$"E{i + 2}"].Value = result.DateErrorMessage;
}
workBook.SaveAs(@"Spreadsheets\\PeopleValidated.xlsx");