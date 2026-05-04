using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsEnumeratorScenarios
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for demonstration
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);
            sheet.Cells["A4"].PutValue("Charlie");
            sheet.Cells["B4"].PutValue(78);

            // -------------------------------------------------
            // Scenario 1: Enumerate all cells in the worksheet
            // -------------------------------------------------
            Console.WriteLine("=== All Cells ===");
            IEnumerator cellEnum = sheet.Cells.GetEnumerator();
            while (cellEnum.MoveNext())
            {
                Cell c = (Cell)cellEnum.Current;
                Console.WriteLine($"{c.Name}: {c.Value}");
            }

            // -------------------------------------------------
            // Scenario 2: Enumerate rows (normal order)
            // -------------------------------------------------
            Console.WriteLine("\n=== Rows (Normal Order) ===");
            IEnumerator rowEnum = sheet.Cells.Rows.GetEnumerator();
            while (rowEnum.MoveNext())
            {
                Row r = (Row)rowEnum.Current;
                Console.WriteLine($"Row {r.Index} Height={r.Height}");
            }

            // -------------------------------------------------
            // Scenario 3: Enumerate rows in reverse order using synchronized enumerator
            // -------------------------------------------------
            Console.WriteLine("\n=== Rows (Reverse Order, Synchronized) ===");
            IEnumerator revSyncEnum = sheet.Cells.Rows.GetEnumerator(true, true);
            while (revSyncEnum.MoveNext())
            {
                Row r = (Row)revSyncEnum.Current;
                Console.WriteLine($"Row {r.Index}");
            }

            // -------------------------------------------------
            // Scenario 4: Enumerate cells within a specific range
            // -------------------------------------------------
            Console.WriteLine("\n=== Range B2:C4 ===");
            AsposeRange range = sheet.Cells.CreateRange("B2:C4");
            IEnumerator rangeEnum = range.GetEnumerator();
            while (rangeEnum.MoveNext())
            {
                Cell c = (Cell)rangeEnum.Current;
                Console.WriteLine($"{c.Name}: {c.Value}");
            }

            // -------------------------------------------------
            // Scenario 5: Enumerate pivot table row fields
            // -------------------------------------------------
            // Add data for a pivot table
            sheet.Cells["D1"].PutValue("Region");
            sheet.Cells["E1"].PutValue("Sales");
            sheet.Cells["D2"].PutValue("North");
            sheet.Cells["E2"].PutValue(1500);
            sheet.Cells["D3"].PutValue("South");
            sheet.Cells["E3"].PutValue(1200);
            sheet.Cells["D4"].PutValue("East");
            sheet.Cells["E4"].PutValue(1300);
            sheet.Cells["D5"].PutValue("West");
            sheet.Cells["E5"].PutValue(1100);

            // Create pivot table
            int pivotIdx = sheet.PivotTables.Add("D1:E5", "G3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Region
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales

            // Enumerate row fields of the pivot table
            Console.WriteLine("\n=== Pivot Row Fields ===");
            PivotFieldCollection rowFields = pivot.RowFields;
            IEnumerator pivotFieldEnum = rowFields.GetEnumerator();
            while (pivotFieldEnum.MoveNext())
            {
                PivotField pf = (PivotField)pivotFieldEnum.Current;
                Console.WriteLine($"Field: {pf.Name}");
            }

            // -------------------------------------------------
            // Save the workbook to demonstrate that all changes persist
            // -------------------------------------------------
            workbook.Save("EnumeratorScenariosOutput.xlsx");
            Console.WriteLine("\nWorkbook saved as 'EnumeratorScenariosOutput.xlsx'.");
        }
    }
}