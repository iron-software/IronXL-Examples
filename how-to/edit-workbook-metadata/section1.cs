using System;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.EditWorkbookMetadata
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Set author
            workBook.Metadata.Author = "Your Name";
            // Set comments
            workBook.Metadata.Comments = "Monthly report";
            // Set title
            workBook.Metadata.Title = "July";
            // Set keywords
            workBook.Metadata.Keywords = "Report";
            
            // Read the creation date of the excel file
            DateTime? creationDate = workBook.Metadata.Created;
            
            // Read the last printed date of the excel file
            DateTime? printDate = workBook.Metadata.LastPrinted;
            
            workBook.SaveAs("editedMetadata.xlsx");
        }
    }
}