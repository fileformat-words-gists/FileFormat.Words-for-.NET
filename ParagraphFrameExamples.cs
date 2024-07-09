using System.Linq;

namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word paragraphs frames/borders
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Paragraph/Frame at the root of your project.
    /// // Check reference for more options and details.
    /// var paraFrameExamples = new FileFormat.Words.Examples.ParagraphFrameExamples();
    /// // Creates a word document with 4 different paragraphs with frames + one paragraph without frames and saves 
    /// // word document to the specified directory. Check reference for more options and details.
    /// paraFrameExamples.CreateParagraphsFrames();
    /// // Reads Paragraphs from the specified Word Document and displays plain text with borders/frames info.
    /// // Check reference for more options and details.
    /// paraFrameExamples.ReadParagraphsFrames();
    /// // Modifies border and text of framed paragraphs in the Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// paraFrameExamples.ModifyParagraphsFrames();
    /// </code>
    /// </example>
    public class ParagraphFrameExamples
    {
        private const string docsDirectory = "../../../Documents/Paragraph/Frame";
        /// <summary>
        /// Initializes a new instance of the <see cref="ParagraphFrameExamples"/> class.
        /// Prepares the directory 'Documents/Paragraph/Frame' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ParagraphFrameExamples()
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
        /// Generates 4 different paragraphs with frames/borders + one paragraph without frames/borders.
        /// Saves the newly created Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Paragraph/Frame' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordParagraphsFrame.docx").
        /// </param>
        public void CreateParagraphsFrames(string documentDirectory = docsDirectory, string filename = "WordParagraphsFrame.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new FileFormat.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new FileFormat.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Create first framed paragraph
                var para = new FileFormat.Words.IElements.Paragraph();
                para.AddRun(new FileFormat.Words.IElements.Run { Text = "This paragraph should have single width frames with blue color" });

                para.ParagraphBorder.Width = FileFormat.Words.IElements.BorderWidth.Single;
                para.ParagraphBorder.Color = FileFormat.Words.IElements.Colors.Blue;
                para.ParagraphBorder.Size = 4;
                System.Console.WriteLine("Single width blue color frame paragraph created");
                body.AppendChild(para);
                System.Console.WriteLine($"Single width blue color frame paragraph appended to word document Body");

                // Create second framed paragraph
                para = new FileFormat.Words.IElements.Paragraph();
                para.AddRun(new FileFormat.Words.IElements.Run { Text = "This paragraph should have double width frames with red color" });

                para.ParagraphBorder.Width = FileFormat.Words.IElements.BorderWidth.Double;
                para.ParagraphBorder.Color = FileFormat.Words.IElements.Colors.Red;
                para.ParagraphBorder.Size = 4;
                System.Console.WriteLine("Double width red color frame paragraph created");
                body.AppendChild(para);
                System.Console.WriteLine($"Double width red color frame paragraph appended to word document Body");

                // Create third framed paragraph
                para = new FileFormat.Words.IElements.Paragraph();
                para.AddRun(new FileFormat.Words.IElements.Run { Text = "This paragraph should have dotted width frames with purple color" });

                para.ParagraphBorder.Width = FileFormat.Words.IElements.BorderWidth.Dotted;
                para.ParagraphBorder.Color = FileFormat.Words.IElements.Colors.Purple;
                para.ParagraphBorder.Size = 4;
                System.Console.WriteLine("Dotted width purple color frame paragraph created");
                body.AppendChild(para);
                System.Console.WriteLine($"Dotted width purple color frame paragraph appended to word document Body");

                // Create fourth framed paragraph
                para = new FileFormat.Words.IElements.Paragraph();
                para.AddRun(new FileFormat.Words.IElements.Run { Text = "This paragraph should have dotdash width frames with navy color" });

                para.ParagraphBorder.Width = FileFormat.Words.IElements.BorderWidth.DotDash;
                para.ParagraphBorder.Color = FileFormat.Words.IElements.Colors.Navy;
                para.ParagraphBorder.Size = 4;
                System.Console.WriteLine("Dotdash width navy color frame paragraph created");
                body.AppendChild(para);
                System.Console.WriteLine($"Dotdash width navy color frame paragraph appended to word document Body");

                // Create normal paragraph
                para = new FileFormat.Words.IElements.Paragraph();
                para.AddRun(new FileFormat.Words.IElements.Run { Text = "This paragraph should have no frames/borders" });

                System.Console.WriteLine("Normal paragraph without frames/borders created");
                body.AppendChild(para);
                System.Console.WriteLine($"Normal paragraph without frames/borders appended to word document Body");


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
        /// Traverses paragraphs and displays thier plain text along with border/frame info.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Paragraph/Frame' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordParagraphsFrame.docx").
        /// </param>
        public void ReadParagraphsFrames(string documentDirectory = docsDirectory, string filename = "WordParagraphsFrame.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);

                // Traverse and display paragraphs with plain text
                foreach (var paragraph in body.Paragraphs)
                {
                    System.Console.WriteLine($" Paragraph Text: {paragraph.Text}");
                    if (paragraph.ParagraphBorder.Size > 0)
                    {
                        System.Console.WriteLine($" Paragraph Border Width: {paragraph.ParagraphBorder.Width}");
                        System.Console.WriteLine($" Paragraph Border Color: {paragraph.ParagraphBorder.Color}");
                        System.Console.WriteLine($" Paragraph Border Size: {paragraph.ParagraphBorder.Size}");
                    }
                    else
                    {
                        System.Console.WriteLine($" Paragraph Border Width: ");
                        System.Console.WriteLine($" Paragraph Border Color: ");
                        System.Console.WriteLine($" Paragraph Border Size: ");
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
        /// Modifies paragraphs frames to single border width and black color.
        /// Modifies the text of framed paragraph with modified frames info. 
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Paragraph' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordParagraphs.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordParagraphs.docx").
        /// </param>
        public void ModifyParagraphsFrames(string documentDirectory = docsDirectory,
            string filename = "WordParagraphsFrame.docx", string filenameModified = "ModifiedParagraphsFrame.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);

                foreach (var paragraph in body.Paragraphs)
                {
                    if (paragraph.ParagraphBorder.Size > 0)
                    {
                        paragraph.ParagraphBorder.Width = FileFormat.Words.IElements.BorderWidth.Single;
                        paragraph.ParagraphBorder.Color = FileFormat.Words.IElements.Colors.Black;
                        foreach (var run in paragraph.Runs)
                        {
                            // Modified paragraph text
                            run.Text = "Paragraph border modified to single width with black color";
                        }
                        System.Console.WriteLine("Frames/Borders changed to single width with black color");
                    }
                    // Update the paragraph in the document
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