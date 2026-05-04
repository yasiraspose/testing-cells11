using System;
using System.IO;
using Aspose.Cells;

class ThreadedCommentExample
{
    static void Main()
    {
        // Create a new workbook in memory
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author to the workbook
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "John Doe",               // Author name
            "john.doe@example.com",   // User ID / email
            "EXAMPLE_PROVIDER");      // Provider (can be empty)

        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Insert a threaded comment into cell B2 using the cell name overload
        worksheet.Comments.AddThreadedComment("B2", "This is a threaded comment.", author);

        // Retrieve and display all threaded comments for the cell (optional)
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments("B2");
        foreach (ThreadedComment tc in threadedComments)
        {
            Console.WriteLine($"Comment by {tc.Author.Name}: {tc.Notes}");
        }

        // Save the workbook to a memory stream (in‑memory XLSX file)
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsx);
            // The MemoryStream now contains the XLSX data.
        }
    }
}