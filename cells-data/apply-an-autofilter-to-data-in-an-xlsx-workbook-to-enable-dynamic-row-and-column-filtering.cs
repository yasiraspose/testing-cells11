using System;
using Aspose.Cells;

public class AutoFilterDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("Food");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["A3"].PutValue("Transport");
        worksheet.Cells["B3"].PutValue(80);
        worksheet.Cells["A4"].PutValue("Food");
        worksheet.Cells["B4"].PutValue(150);
        worksheet.Cells["A5"].PutValue("Utilities");
        worksheet.Cells["B5"].PutValue(200);

        // Apply an auto‑filter to the data range (including the header)
        worksheet.AutoFilter.Range = "A1:B5";

        // Example filter: show only rows where Category = "Food"
        worksheet.AutoFilter.Filter(0, "Food");
        worksheet.AutoFilter.Refresh();

        // Save the workbook with the auto‑filter applied
        workbook.Save("AutoFilterDemo.xlsx");
    }
}