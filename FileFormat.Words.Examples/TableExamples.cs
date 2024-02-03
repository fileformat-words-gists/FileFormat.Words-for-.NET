namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word tables
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Table at the root of your project.
    /// // Check reference for more options and details.
    /// TableExamples tableExamples = new TableExamples();
    /// // Creates a word document with tables and saves word document to the specified 
    /// // directory. Check reference for more options and details.
    /// tableExamples.CreateWordDocumentWithTables();
    /// // Read tables from the specified Word Document and displays table contents.
    /// // Check reference for more options and details.
    /// tableExamples.ReadTablesInWordDocument();
    /// // Modify Images in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// tableExamples.ModifyTablesInWordDocument();
    /// </code>
    /// </example>
    public class TableExamples
    {
        private const string docsDirectory = "../../../Documents/Table";
        /// <summary>
        /// Initializes a new instance of the <see cref="TableExamples"/> class.
        /// Prepares the directory 'Documents/Table' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public TableExamples()
        {
            if (!System.IO.Directory.Exists(docsDirectory))
            {
                // If it doesn't exist, create the directory
                System.IO.Directory.CreateDirectory(docsDirectory);
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(docsDirectory)}' " +
                    $"created successfully.");
            }
            else
            {
                var files = System.IO.Directory.GetFiles(System.IO.Path.GetFullPath(docsDirectory));
                foreach (var file in files)
                {
                    System.IO.File.Delete(file);
                    System.Console.WriteLine($"File deleted: {file}");
                }
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(docsDirectory)}' " +
                    $"cleaned up.");
            }
        }
        /// <summary>
        /// Creates a new Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a>.
        /// Generates 5(rows) x 3(cols) tables with table styles defined by the Word document template.
        /// Appends each table to the body of the word document.
        /// Saves the newly created word document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Table' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordTables.docx").
        /// </param>
        public void CreateWordDocumentWithTables(string documentDirectory = docsDirectory,
            string filename = "WordTables.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new FileFormat.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new FileFormat.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Get all table styles
                var tableStyles = doc.ElementStyles.TableStyles;
                System.Console.WriteLine("Table styles loaded");

                // Create Headings Paragraph and append to the body.
                foreach (var tableStyle in tableStyles)
                {
                    var table = new FileFormat.Words.IElements.Table(5,3);
                    table.Style = tableStyle;

                    table.Column.Width = 2500;

                    var rowNumber = 0;
                    var columnNumber = 0;

                    var para = new FileFormat.Words.IElements.Paragraph();
                    para.Style = FileFormat.Words.IElements.Headings.Heading1;
                    para.AddRun(new FileFormat.Words.IElements.Run {
                        Text = $"Table With Style '{tableStyle}' : "
                    });

                    body.AppendChild(para);

                    foreach (var row in table.Rows)
                    {
                        rowNumber++;
                        foreach(var cell in row.Cells)
                        {
                            columnNumber++;
                            para = new FileFormat.Words.IElements.Paragraph();
                            para.AddRun(new FileFormat.Words.IElements.Run {
                                Text = $"Row {rowNumber} Column {columnNumber}"
                            });
                            cell.Paragraphs.Add(para);
                        }
                        columnNumber = 0;
                    }
                    body.AppendChild(table);
                    System.Console.WriteLine($"Table with style {tableStyle} created and appended");
                }

                // Save the newly created Word Document.
                doc.Save($"{documentDirectory}/{filename}");
                System.Console.WriteLine($"Word Document {filename} Created. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Words.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Loads a Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a>.
        /// Traverses tables and displays associated styles as defined by the Word document template.
        /// Traverses through each row and then traverses columns within the row.
        /// Traverses through paragrpahs within each cell and displays paragraph plain text
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Table' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordTables.docx").
        /// </param>
        public void ReadTablesInWordDocument(string documentDirectory = docsDirectory,
            string filename = "WordTables.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);

                var tables = body.Tables;

                var tableNumber = 0;
                var rowNumber = 0;
                var columnNumber = 0;
                var paraNumber = 0;

                foreach (var table in tables)
                {
                    tableNumber++;
                    System.Console.WriteLine($"Table Number : {tableNumber}");
                    System.Console.WriteLine($"..Table Style : {table.Style}");
                    foreach (var row in table.Rows)
                    {
                        rowNumber++;
                        System.Console.WriteLine($"..Row Number : {rowNumber}");
                        foreach (var cell in row.Cells)
                        {
                            columnNumber++;
                            System.Console.WriteLine($"....Column Number : {columnNumber}");
                            foreach (var para in cell.Paragraphs)
                            {
                                paraNumber++;
                                System.Console.WriteLine($"......Paragraph Number : {paraNumber}");
                                System.Console.WriteLine($"......Paragraph Text : {para.Text}");
                            }
                            paraNumber = 0;
                        }
                        columnNumber = 0;
                    }
                    rowNumber = 0;
                }
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Words.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Loads a Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a>.
        /// Modifies tables by setting column widths to 2000
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Paragraph' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordTables.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordTables.docx").
        /// </param>
        public void ModifyTablesInWordDocument(string documentDirectory = docsDirectory,
            string filename = "WordTables.docx", string filenameModified = "ModifiedWordTables.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);

                var tables = body.Tables;

                foreach (var table in tables)
                {
                    table.Column.Width = 2000;
                    doc.Update(table);
                }

                // Save the modified Word Document
                doc.Save($"{documentDirectory}/{filenameModified}");
                System.Console.WriteLine($"Word Document {filename} Modified and " +
                    $"Saved As {filenameModified}. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Words.FileFormatException("An error occurred.", ex);
            }
        }
    }
}
