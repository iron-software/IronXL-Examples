using System;
using IronXL.Excel;
namespace ironxl.CsharpExcelMergeCells
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Apply merge
            workSheet.Merge("B4:C4");
            workSheet.Merge("A1:A4");
            workSheet.Merge("A6:D9");
            
            // Retrieve merged regions
            List<IronXL.Range> retrieveMergedRegions = workSheet.GetMergedRegions();
            
            foreach (IronXL.Range mergedRegion in retrieveMergedRegions)
            {
                Console.WriteLine(mergedRegion.RangeAddressAsString);
            }
        }
    }
}