using System.Linq;

namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying indentation of Word paragraphs
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Paragraph/Indent at the root of your project.
    /// // Check reference for more options and details.
    /// ParagraphIndentExamples paragraphIndentExamples = new ParagraphIndentExamples();
    /// // Creates a word document with paragraphs having various indentations and saves word document
    /// // to the specified directory. Check reference for more options and details.
    /// paragraphIndentExamples.CreateIndent();
    /// // Reads Paragraphs from the specified Word Document and displays plain text with indentation.
    /// // Check reference for more options and details.
    /// paragraphIndentExamples.ReadIndent();
    /// // Modifies Paragraph's indentation in the Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// paragraphIndentExamples.ModifyIndent();
    /// </code>
    /// </example>
    public class ParagraphIndentExamples
    {
        private const string docsDirectory = "../../../Documents/Paragraph/Indent";
        /// <summary>
        /// Initializes a new instance of the <see cref="ParagraphIndentExamples"/> class.
        /// Prepares the directory 'Documents/Paragraph/Indent' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ParagraphIndentExamples()
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
        /// Generates paragraphs with different indents including left, right, firstline and hanging.
        /// Saves the newly created Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Paragraph' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordParagraphs.docx").
        /// </param>
        public void CreateIndent(string documentDirectory = docsDirectory, string filename = "WordParagraphsIndented.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new FileFormat.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new FileFormat.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Add paragraph with left indentation
                FileFormat.Words.IElements.Paragraph para1 = new FileFormat.Words.IElements.Paragraph();
                para1.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "First paragraph with left indentation"
                });
                // Setting the Paragraph Indent 2 inches to left.
                para1.Indentation.Left = 2;

                // Add paragraph with right indentation
                FileFormat.Words.IElements.Paragraph para2 = new FileFormat.Words.IElements.Paragraph();
                para2.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "Second paragraph with right indentation"
                });
                // Setting the Paragraph Indent 2 inches to right.
                para2.Indentation.Right = 2;

                // Add paragraph with firstline indentation
                FileFormat.Words.IElements.Paragraph para3 = new FileFormat.Words.IElements.Paragraph();
                para3.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "Third paragraph with firstline indentation"
                });
                // Setting the Paragraph Indent 2 inches to firstline.
                para3.Indentation.FirstLine = 2;

                // Add paragraph with hanging indentation
                FileFormat.Words.IElements.Paragraph para4 = new FileFormat.Words.IElements.Paragraph();
                para4.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "Fourth paragraph with hanging indentation"
                });
                // Setting the Paragraph Indent 2 inches to hanging.
                para4.Indentation.Hanging = 2;

                // Append paragraphs to the document body
                body.AppendChild(para1);
                body.AppendChild(para2);
                body.AppendChild(para3);
                body.AppendChild(para4);

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
        /// Traverses paragraphs and displays its text along with indentation.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Paragraph/Indent' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordParagraphsIndented.docx").
        /// </param>
        public void ReadIndent(string documentDirectory = docsDirectory, string filename = "WordParagraphsIndented.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");

                // Initialize the body with the document
                var body = new FileFormat.Words.Body(doc);

                System.Collections.Generic.List<FileFormat.Words.IElements.Paragraph>
                    paragraphs = body.Paragraphs;

                foreach (FileFormat.Words.IElements.Paragraph paragraph in paragraphs)
                {
                    System.Console.WriteLine($"Paragraph Text : {paragraph.Text}");
                    
                    switch (paragraph.Indentation)
                    {
                        case var indent when indent.Left > 0:
                            System.Console.WriteLine($"Paragraph Left Indent: {indent.Left}");
                            break;
                        case var indent when indent.Right > 0:
                            System.Console.WriteLine($"Paragraph Right Indent: {indent.Right}");
                            break;
                        case var indent when indent.FirstLine > 0:
                            System.Console.WriteLine($"Paragraph Firstline Indent: {indent.FirstLine}");
                            break;
                        case var indent when indent.Hanging > 0:
                            System.Console.WriteLine($"Paragraph Hanging Indent: {indent.Hanging}");
                            break;
                        default:
                            System.Console.WriteLine("No Indentation.");
                            break;
                    }
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
        /// Traverses through all paragraphs within the document.
        /// Modifies paragraphs by appending the relevant indent message in italic format
        /// and modifies the indent (if found) to 0.
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Paragraph/Indent' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordParagraphsIndented.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordParagraphsIndented.docx").
        /// </param>
        public void ModifyIndent(string documentDirectory = docsDirectory,
            string filename = "WordParagraphsIndented.docx", string filenameModified = "ModifiedWordParagraphsIndented.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");

                // Initialize the body with the document
                var body = new FileFormat.Words.Body(doc);

                foreach (FileFormat.Words.IElements.Paragraph paragraph in body.Paragraphs)
                {
                    switch (paragraph.Indentation)
                    {
                        case var indent when indent.Left > 0:
                            paragraph.AddRun(new FileFormat.Words.IElements.Run
                            { Text = " (left indent set to 0)", Italic = true });
                            paragraph.Indentation.Left = 0;
                            break;
                        case var indent when indent.Right > 0:
                            paragraph.AddRun(new FileFormat.Words.IElements.Run
                            { Text = " (right indent set to 0)", Italic = true });
                            paragraph.Indentation.Right = 0;
                            break;
                        case var indent when indent.FirstLine > 0:
                            paragraph.AddRun(new FileFormat.Words.IElements.Run
                            { Text = " (FirstLine indent set to 0)", Italic = true });
                            paragraph.Indentation.FirstLine = 0;
                            break;
                        case var indent when indent.Hanging > 0:
                            paragraph.AddRun(new FileFormat.Words.IElements.Run
                            { Text = " (hanging indent set to 0)", Italic = true });
                            paragraph.Indentation.Hanging = 0;
                            break;
                        default:
                            paragraph.AddRun(new FileFormat.Words.IElements.Run
                            { Text = " (No indent and no change)", Italic = true });
                            break;
                    }
                    doc.Update(paragraph);
                }

                // Save the modified Word Document
                doc.Save($"{documentDirectory}/{filenameModified}");
                System.Console.WriteLine($"Word Document {filename} Modified and Saved As " +
                    $"{filenameModified}. Please check directory: " +
                    $"{System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Words.FileFormatException("An error occurred.", ex);
            }
        }
    }
}
