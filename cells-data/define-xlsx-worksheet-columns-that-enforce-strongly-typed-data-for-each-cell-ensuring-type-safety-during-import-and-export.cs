using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsTypedColumnsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Define column data types using data validation
            // ------------------------------------------------------------

            // Column A (index 0) – Whole numbers (int)
            CellArea intArea = new CellArea { StartRow = 1, EndRow = 100, StartColumn = 0, EndColumn = 0 };
            int intValidationIdx = sheet.Validations.Add(intArea);
            Validation intValidation = sheet.Validations[intValidationIdx];
            intValidation.Type = ValidationType.WholeNumber;
            intValidation.Operator = OperatorType.Between;
            intValidation.Formula1 = "0";      // Minimum allowed value
            intValidation.Formula2 = "1000";   // Maximum allowed value
            intValidation.ShowError = true;
            intValidation.ErrorTitle = "Invalid Integer";
            intValidation.ErrorMessage = "Please enter an integer between 0 and 1000.";

            // Column B (index 1) – Decimal numbers
            CellArea decArea = new CellArea { StartRow = 1, EndRow = 100, StartColumn = 1, EndColumn = 1 };
            int decValidationIdx = sheet.Validations.Add(decArea);
            Validation decimalValidation = sheet.Validations[decValidationIdx];
            decimalValidation.Type = ValidationType.Decimal;
            decimalValidation.Operator = OperatorType.Between;
            decimalValidation.Formula1 = "0.0";
            decimalValidation.Formula2 = "10000.0";
            decimalValidation.ShowError = true;
            decimalValidation.ErrorTitle = "Invalid Decimal";
            decimalValidation.ErrorMessage = "Please enter a decimal number between 0 and 10000.";

            // Column C (index 2) – Date values
            CellArea dateArea = new CellArea { StartRow = 1, EndRow = 100, StartColumn = 2, EndColumn = 2 };
            int dateValidationIdx = sheet.Validations.Add(dateArea);
            Validation dateValidation = sheet.Validations[dateValidationIdx];
            dateValidation.Type = ValidationType.Date;
            dateValidation.Operator = OperatorType.Between;
            // Allow dates from Jan 1, 2000 to Dec 31, 2100
            dateValidation.Formula1 = DateTime.Parse("2000-01-01").ToOADate().ToString();
            dateValidation.Formula2 = DateTime.Parse("2100-12-31").ToOADate().ToString();
            dateValidation.ShowError = true;
            dateValidation.ErrorTitle = "Invalid Date";
            dateValidation.ErrorMessage = "Please enter a date between 01/01/2000 and 12/31/2100.";

            // ------------------------------------------------------------
            // 2. Insert header row
            // ------------------------------------------------------------
            cells[0, 0].PutValue("Quantity (Int)");
            cells[0, 1].PutValue("Price (Decimal)");
            cells[0, 2].PutValue("Order Date (Date)");

            // ------------------------------------------------------------
            // 3. Import sample data (strongly typed)
            // ------------------------------------------------------------
            object[] sampleData = new object[]
            {
                10, 99.95, new DateTime(2023, 12, 15),
                25, 15.5,  new DateTime(2024, 1, 5),
                5,  250.0, new DateTime(2023, 11, 30)
            };
            // Import horizontally starting at row 1, column 0
            cells.ImportObjectArray(sampleData, 1, 0, false, 0);

            // ------------------------------------------------------------
            // 4. Export the typed range to a DataTable with type checking
            // ------------------------------------------------------------
            ExportTableOptions exportOptions = new ExportTableOptions
            {
                ExportColumnName = true,      // Use first row as column names
                CheckMixedValueType = true,   // Verify that each column has a consistent type
                PlotVisibleCells = true,
                PlotVisibleRows = true,
                PlotVisibleColumns = true
            };

            // Export rows 0..3 (header + 3 data rows) and columns 0..2
            DataTable dt = sheet.Cells.ExportDataTable(0, 0, 4, 3, exportOptions);

            // Display exported DataTable schema (for demonstration)
            Console.WriteLine("Exported DataTable column types:");
            foreach (DataColumn col in dt.Columns)
            {
                Console.WriteLine($"{col.ColumnName} : {col.DataType}");
            }

            // ------------------------------------------------------------
            // 5. Save the workbook (XLSX)
            // ------------------------------------------------------------
            workbook.Save("TypedColumnsDemo.xlsx");
        }
    }
}