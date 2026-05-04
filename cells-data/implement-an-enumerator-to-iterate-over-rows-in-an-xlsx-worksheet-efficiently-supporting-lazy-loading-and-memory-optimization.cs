using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsLazyRowIteration
{
    // Enumerator that lazily iterates over rows of a worksheet.
    // It internally uses RowCollection.GetEnumerator with synchronization enabled,
    // which avoids loading the whole worksheet into memory at once.
    public class LazyRowEnumerator : IEnumerator<Row>
    {
        private readonly IEnumerator _innerEnumerator;

        // reversed: iterate from bottom to top if true.
        // sync:    keep the enumerator synchronized with any modifications.
        public LazyRowEnumerator(Worksheet worksheet, bool reversed = false, bool sync = true)
        {
            // RowCollection provides a built‑in enumerator.
            _innerEnumerator = worksheet.Cells.Rows.GetEnumerator(reversed, sync);
        }

        public Row Current => (Row)_innerEnumerator.Current;

        object IEnumerator.Current => Current;

        public bool MoveNext()
        {
            return _innerEnumerator.MoveNext();
        }

        public void Reset()
        {
            _innerEnumerator.Reset();
        }

        public void Dispose()
        {
            // No unmanaged resources to release.
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to a large XLSX file.
            const string inputPath = "LargeWorkbook.xlsx";
            const string outputPath = "ProcessedWorkbook.xlsx";

            // Load the workbook with memory‑optimized settings.
            // LoadOptions enables lazy loading of rows/cells.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                // Prefer memory‑optimized mode (available in Aspose.Cells 23+).
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook using the options defined above.
            var workbook = new Workbook(inputPath, loadOptions);
            var worksheet = workbook.Worksheets[0];

            // Optional: further reduce memory usage for the worksheet.
            worksheet.Cells.MemorySetting = MemorySetting.MemoryPreference;

            // Iterate over rows lazily.
            using (var rowEnumerator = new LazyRowEnumerator(worksheet))
            {
                while (rowEnumerator.MoveNext())
                {
                    Row row = rowEnumerator.Current;

                    // Example processing: output row index and first cell value (if any).
                    Cell firstCell = row.GetCellOrNull(0);
                    string cellValue = firstCell?.StringValue ?? "<empty>";
                    Console.WriteLine($"Row {row.Index}: First cell = {cellValue}");

                    // Additional row‑level processing can be placed here.
                }
            }

            // Save the workbook after processing.
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}