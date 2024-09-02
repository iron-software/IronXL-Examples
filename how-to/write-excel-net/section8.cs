using IronXL;

// Load Excel file
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Open WorkSheet of sample.xlsx
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Specify range in which we want to write the values
for (int i = 2; i <= 7; i++)
{
    // Write the Dynamic value in one row
    workSheet["B" + i].Value = "Value" + i;

    // Write the Dynamic value in another row
    workSheet["D" + i].Value = "Value" + i;
}

// Save changes
workBook.SaveAs("sample.xlsx");