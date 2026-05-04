using System;
using System.Data;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Prepare a DataTable with typed columns (string, int, decimal)
        DataTable dataTable = new DataTable();
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Age", typeof(int));
        dataTable.Columns.Add("Salary", typeof(decimal));

        // Add sample rows
        dataTable.Rows.Add("John", 30, 50000.75m);
        dataTable.Rows.Add("Jane", 28, 62000.00m);
        dataTable.Rows.Add("Bob", 35, 72000.50m);

        // Write column headers
        for (int col = 0; col < dataTable.Columns.Count; col++)
        {
            cells[0, col].PutValue(dataTable.Columns[col].ColumnName);
        }

        // Write data rows
        for (int row = 0; row < dataTable.Rows.Count; row++)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                cells[row + 1, col].PutValue(dataTable.Rows[row][col]);
            }
        }

        // Save the workbook as an XLSX file
        workbook.Save("DataGridImportDemo.xlsx");
    }
}