using System;
using Aspose.Cells;

class RetrieveThreadedCommentTimestamps
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the comments collection of the current worksheet
            CommentCollection comments = sheet.Comments;

            // Loop through each comment in the collection
            for (int i = 0; i < comments.Count; i++)
            {
                Comment comment = comments[i];

                // Determine the cell address of the comment (e.g., "A1")
                string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);

                // Retrieve all threaded comments associated with this cell
                ThreadedCommentCollection threadedComments = comments.GetThreadedComments(cellName);

                // Loop through each threaded comment and output its creation timestamp
                for (int j = 0; j < threadedComments.Count; j++)
                {
                    ThreadedComment tc = threadedComments[j];
                    DateTime createdTime = tc.CreatedTime;
                    Console.WriteLine($"Worksheet: {sheet.Name}, Cell: {cellName}, ThreadedComment #{j + 1}, CreatedTime: {createdTime}");
                }
            }
        }

        // Save the workbook if any modifications were made (optional)
        workbook.Save("output.xlsx");
    }
}