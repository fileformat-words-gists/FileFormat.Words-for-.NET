using System.Linq;

namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying alignment of Word paragraphs
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Paragraph/Alignment at the root of your project.
    /// // Check reference for more options and details.
    /// ParagraphAlignmentExamples paragraphAlignment = new ParagraphAlignmentExamples();
    /// // Creates a word document with paragraphs having different alignments and saves word document
    /// // to the specified directory. Check reference for more options and details.
    /// paragraphAlignment.CreateAlignment();
    /// // Reads Paragraphs from the specified Word Document and displays plain text alongwith alignment.
    /// // Check reference for more options and details.
    /// paragraphAlignment.ReadAlignment();
    /// // Modifies Paragraph's alignment in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// paragraphAlignment.ModifyAlignment();
    /// </code>
    /// </example>
    public class ParagraphAlignmentExamples
    {
        private const string docsDirectory = "../../../Documents/Paragraph/Alignment";
        /// <summary>
        /// Initializes a new instance of the <see cref="ParagraphAlignmentExamples"/> class.
        /// Prepares the directory 'Documents/Paragraph/Alignment' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ParagraphAlignmentExamples()
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
        /// Generates paragraphs with different alignments including left, center, right and justify.
        /// Saves the newly created Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Paragraph/Alignment' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordParagraphs.docx").
        /// </param>
        public void CreateAlignment(string documentDirectory = docsDirectory, string filename = "WordParagraphsAligned.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new FileFormat.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new FileFormat.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Add paragraph with left alignment
                var para1 = new FileFormat.Words.IElements.Paragraph();
                para1.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "First paragraph with left alignment"
                });
                // Setting the Paragraph Alignment to Center.
                para1.Alignment = FileFormat.Words.IElements.ParagraphAlignment.Left; //"Left";

                // Add paragraph with center alignment
                var para2 = new FileFormat.Words.IElements.Paragraph();
                para2.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "Second paragraph with center alignment"
                });
                // Setting the Paragraph Alignment to Center.
                para2.Alignment = FileFormat.Words.IElements.ParagraphAlignment.Center; //"Center";

                // Add paragraph with right alignment
                var para3 = new FileFormat.Words.IElements.Paragraph();
                para3.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "Third paragraph with right alignment"
                });
                // Setting the Paragraph Alignment to Center.
                para3.Alignment = FileFormat.Words.IElements.ParagraphAlignment.Right; //"Right";

                // Add paragraph with justify alignment
                var para4 = new FileFormat.Words.IElements.Paragraph();
                para4.AddRun(new FileFormat.Words.IElements.Run
                {
                    Text = "Fourth paragraph with justify alignment"
                });
                // Setting the Paragraph Alignment to Center.
                para4.Alignment = FileFormat.Words.IElements.ParagraphAlignment.Justify; //"Justify";

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
        /// Traverses paragraphs and displays its text along with alignment.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Paragraph/Alignment' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordParagraphsAligned.docx").
        /// </param>
        public void ReadAlignment(string documentDirectory = docsDirectory,
            string filename = "WordParagraphsAligned.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");

                // Initialize the body with the document
                var body = new FileFormat.Words.Body(doc);

                //System.Collections.Generic.List<FileFormat.Words.IElements.Paragraph>
                  //  paragraphs = body.Paragraphs;

                var paragraphs = body.Paragraphs;

                foreach (FileFormat.Words.IElements.Paragraph paragraph in paragraphs)
                {
                    System.Console.WriteLine($"Paragraph Text : {paragraph.Text}");
                    System.Console.WriteLine($"Paragraph Alignment : {paragraph.Alignment}");
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
        /// Modifies paragraphs by appending ' (alignment modified to justify)' with italic format
        /// and justify alignment.
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Paragraph/Alignment' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordParagraphsAligned.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordParagraphsAligned.docx").
        /// </param>
        public void ModifyAlignment(string documentDirectory = docsDirectory,
            string filename = "WordParagraphsAligned.docx",
            string filenameModified = "ModifiedWordParagraphsAligned.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");

                // Initialize the body with the document
                var body = new FileFormat.Words.Body(doc);

                //foreach (FileFormat.Words.IElements.Paragraph paragraph in body.Paragraphs)
                foreach (var paragraph in body.Paragraphs)
                {
                    paragraph.AddRun(new FileFormat.Words.IElements.Run
                    { Text = " (alignment modified to justify)", Italic=true });
                    paragraph.Alignment = FileFormat.Words.IElements.ParagraphAlignment.Justify; //"Justify";
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
