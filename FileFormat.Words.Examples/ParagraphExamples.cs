using System.Linq;

namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word paragraphs
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Paragraph at the root of your project.
    /// // Check reference for more options and details.
    /// ParagraphExamples paragraphExamples = new ParagraphExamples();
    /// // Creates a word document with paragraphs and saves word document to the specified 
    /// // directory. Check reference for more options and details.
    /// paragraphExamples.CreateWordParagraphs();
    /// // Reads Paragraphs from the specified Word Document and displays plain text and formatting.
    /// // Check reference for more options and details.
    /// paragraphExamples.ReadWordParagraphs();
    /// // Modifies Paragraphs in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// paragraphExamples.ModifyWordParagraphs();
    /// </code>
    /// </example>
    public class ParagraphExamples
    {
        private const string docsDirectory = "../../../Documents/Paragraph";
        /// <summary>
        /// Initializes a new instance of the <see cref="ParagraphExamples"/> class.
        /// Prepares the directory 'Documents/Paragraph' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ParagraphExamples()
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
        /// Generates paragraphs with heading styles defined by the Word document template.
        /// Adds normal paragraphs under each heading paragraph, including text runs with various fonts as per the template.
        /// Saves the newly created Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Paragraph' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordParagraphs.docx").
        /// </param>
        public void CreateWordParagraphs(string documentDirectory = docsDirectory, string filename = "WordParagraphs.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new FileFormat.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new FileFormat.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Get all paragraph styles
                var paragraphStyles = doc.ElementStyles.ParagraphStyles;
                System.Console.WriteLine("Paragraph styles loaded");

                // Get all fonts defined by FontTable and Theme
                var fonts = doc.ElementStyles.TableFonts;
                var fontsTheme = doc.ElementStyles.ThemeFonts;
                System.Console.WriteLine("Fonts defined by FontsTable and Theme loaded");

                // Merge all fonts
                fonts.AddRange(fontsTheme);
                System.Console.WriteLine("All Fonts merged");

                // Create Headings Paragraph and append to the body.
                foreach (var paragraphStyle in paragraphStyles.Where(style => !style.Contains("Normal")))
                {
                    var paragraphWithStyle = new FileFormat.Words.IElements.Paragraph { Style = paragraphStyle };
                    paragraphWithStyle.AddRun(new FileFormat.Words.IElements.Run
                    { Text = $"Paragraph with {paragraphStyle} Style" });
                    System.Console.WriteLine($"Styled Paragraph with {paragraphStyle} Created");

                    body.AppendChild(paragraphWithStyle);
                    System.Console.WriteLine($"Styled Paragraph with {paragraphStyle} Appended to Word Document Body");

                    // Create Normal Paragraph and include text runs with various fonts as per the template.
                    var paragraphNormal = new FileFormat.Words.IElements.Paragraph();
                    System.Console.WriteLine("Normal Paragraph Created");
                    paragraphNormal.AddRun(new FileFormat.Words.IElements.Run
                    {
                        Text = $"Text in normal paragraph with default font and size but with bold " +
                               $"and underlined Gray Color ",
                        Color = FileFormat.Words.IElements.Colors.Gray,
                        Bold = true,
                        Underline = true
                    });
                    foreach (string font in fonts)
                    {
                        paragraphNormal.AddRun(new FileFormat.Words.IElements.Run
                        {
                            Text = $"Text in normal paragraph with font {font} and size 10 but with default " +
                                   $"color, bold, and underlines. ",
                            FontFamily = font,
                            FontSize = 10
                        });
                    }
                    System.Console.WriteLine("All Runs with all fonts Created for Normal Paragraph");
                    body.AppendChild(paragraphNormal);
                    System.Console.WriteLine($"Normal Paragraph Appended to Word Document Body");
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
        /// Traverses paragraphs and displays associated styles as defined by the Word document template.
        /// Traverses through each run (text fragment) within each paragraph and displays fragment values.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Paragraph' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordParagraphs.docx").
        /// </param>
        public void ReadWordParagraphs(string documentDirectory = docsDirectory, string filename = "WordParagraphs.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);
                var num = 0;

                System.Console.WriteLine("Paragraphs Plain Text");

                // Traverse and display paragraphs with plain text
                foreach (var paragraph in body.Paragraphs)
                {
                    num++;
                    System.Console.WriteLine($" Paragraph Number: {num}");
                    System.Console.WriteLine($" Paragraph Style: {paragraph.Style}");
                    System.Console.WriteLine($" Paragraph Text: {paragraph.Text}");
                }

                num = 0;
                var runnum = 0;
                System.Console.WriteLine("Paragraphs with formatting");

                // Traverse and display paragraphs with formatting details
                foreach (var paragraph in body.Paragraphs)
                {
                    num++;
                    System.Console.WriteLine($" Paragraph Number: {num}");
                    System.Console.WriteLine($" Paragraph Style: {paragraph.Style}");

                    // Traverse and display runs within each paragraph
                    foreach (var run in paragraph.Runs)
                    {
                        runnum++;
                        System.Console.WriteLine($"  Text fragment ({num} - {runnum}): {run.Text}");
                        System.Console.WriteLine($"  Font fragment ({num} - {runnum}): {run.FontFamily}");
                        System.Console.WriteLine($"  Color fragment ({num} - {runnum}): {run.Color}");
                        System.Console.WriteLine($"  Size fragment ({num} - {runnum}): {run.FontSize}");
                        System.Console.WriteLine($"  Bold fragment ({num} - {runnum}): {run.Bold}");
                        System.Console.WriteLine($"  Italic fragment ({num} - {runnum}): {run.Italic}");
                        System.Console.WriteLine($"  Underline fragment ({num} - {runnum}): {run.Underline}");
                    }
                    runnum = 0;
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
        /// Modifies paragraphs by prepending 'Modified Heading :' for styled paragraphs
        /// and 'Modified Run :' for each run within normal paragraphs, preserving the existing format.
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
        public void ModifyWordParagraphs(string documentDirectory = docsDirectory,
            string filename = "WordParagraphs.docx", string filenameModified = "ModifiedWordParagraphs.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);

                foreach (var paragraph in body.Paragraphs)
                {
                    foreach (var run in paragraph.Runs)
                    {
                        // Prepend 'Modified Heading :' for styled paragraphs
                        // and 'Modified Run :' for each run within normal paragraphs, preserving the existing format
                        run.Text = paragraph.Style.Contains("Heading") ? $"Modified Heading: {run.Text}" : $"Modified Run : {run.Text}";
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
