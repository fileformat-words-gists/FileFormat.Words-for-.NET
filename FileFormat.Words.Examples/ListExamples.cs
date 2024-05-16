using System;
namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Roman+Alphabetic paragraphs in DOCX
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Paragraph/List at the root of your project.
    /// // Check reference for more options and details.
    /// var listExamples = new ListExamples();
    /// // Creates a word document with two mulitlevel lists each having 3 levels and saves word
    /// // document to the specified directory. Check reference for more options and details.
    /// listExamples.CreateMultilevelLists();
    /// // Reads Paragraphs from the specified Word Document and displays plain text with prefix info at each level.
    /// // Check reference for more options and details.
    /// listExamples.ReadMultilevelLists();
    /// // Modifies list paragraphs in the Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// listExamples.ModiyMultilevelLists();
    /// </code>
    /// </example>
    public class ListExamples
    {
        private const string docsDirectory = "../../../Documents/Paragraph/List";
        /// <summary>
        /// Initializes a new instance of the <see cref="ListExamples"/> class.
        /// Prepares the directory 'Documents/Paragraph/List' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ListExamples()
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
        /// Generates roman and alphabetic paragraphs with nested levels.
        /// Saves the newly created Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Paragraph/List' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordListParagraphs.docx").
        /// </param>
        public void CreateMultilevelLists(string documentDirectory = docsDirectory,
            string filename = "WordListParagraphs.docx")
        {
            try
            {
                // Create a document with multiple multilevel list paragraphs using FileFormat.Words (https://www.nuget.org/packages/FileFormat.Words)

                // The resulting docx document should be like this: https://i.imgur.com/6PKIb56.png

                // Initialize docx document
                FileFormat.Words.Document doc = new FileFormat.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize document body
                var body = new FileFormat.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Function to create a simple paragraph
                void CreateParagraph(FileFormat.Words.Body body, string text)
                {
                    var para = new FileFormat.Words.IElements.Paragraph();
                    para.AddRun(new FileFormat.Words.IElements.Run { Text = text });
                    body.AppendChild(para);
                }

                // Function to create a list paragraph
                void CreateListParagraph(FileFormat.Words.Body body, string text, int numberingId, string numberingType, int numberingLevel)
                {
                    var para = new FileFormat.Words.IElements.Paragraph { Style = "ListParagraph" };
                    para.AddRun(new FileFormat.Words.IElements.Run { Text = text });
                    para.NumberingId = numberingId;
                    if (numberingType == "Number")
                        para.IsNumbered = true;
                    else if (numberingType == "Alphabet")
                        para.IsAlphabeticNumber = true;
                    else if (numberingType == "Roman")
                        para.IsRoman = true;
                    para.NumberingLevel = numberingLevel;
                    body.AppendChild(para);
                }

                // Create paragraphs
                CreateParagraph(body, "This document is generated by FileFormat.Words.");
                CreateParagraph(body, "Below is first multilevel list of paragraphs:");
                CreateListParagraph(body, "First number at first level", 1, "Number", 1);
                CreateListParagraph(body, "First alphabetic at second level", 1, "Alphabet", 2);
                CreateListParagraph(body, "Second alphabetic at second level", 1, "Alphabet", 2);
                CreateListParagraph(body, "First roman at third level", 1, "Roman", 3);
                CreateListParagraph(body, "Second roman at third level", 1, "Roman", 3);
                CreateListParagraph(body, "Second number at first level", 1, "Number", 1);
                CreateParagraph(body, "The first multilevel list ends here...");
                CreateParagraph(body, "Below is second multilevel list of paragraphs:");
                CreateListParagraph(body, "First number at first level", 2, "Number", 1);
                CreateListParagraph(body, "First roman at second level", 2, "Roman", 2);
                CreateListParagraph(body, "Second roman at second level", 2, "Roman", 2);
                CreateListParagraph(body, "First alphabet at third level", 2, "Alphabet", 3);
                CreateListParagraph(body, "Second alphabet at third level", 2, "Alphabet", 3);
                CreateListParagraph(body, "Second number at first level", 2, "Number", 1);
                CreateParagraph(body, "The second multilevel list ends here...");
                CreateParagraph(body, "The document ends here...");

                // Save docx document to the disk
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
        /// Traverses paragraphs and displays its text, roman/alphabetic status and level.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Paragraph/List' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordListParagraphs.docx").
        /// </param>
        public void ReadMultilevelLists(string documentDirectory = docsDirectory,
            string filename = "WordListParagraphs.docx")
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
                    System.Console.WriteLine($"Paragraph NumberingId : {paragraph.NumberingId}");
                    System.Console.WriteLine($"Paragraph Numbered? : {paragraph.IsNumbered}");
                    System.Console.WriteLine($"Paragraph Roman? : {paragraph.IsRoman}");
                    System.Console.WriteLine($"Paragraph AlphabeticNumber? : {paragraph.IsAlphabeticNumber}");
                    System.Console.WriteLine($"Paragraph Numbering Level : {paragraph.NumberingLevel}");
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
        /// If alphabetic, modifies paragraphs by appending ' (numbering type changed to numeric)' with italic format
        /// and paragraph numbering type is changed to numeric.
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Paragraph/List' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordListParagraphs.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordListParagraphs.docx").
        /// </param>
        public void ModifyMultilevelLists(string documentDirectory = docsDirectory,
            string filename = "WordListParagraphs.docx",
            string filenameModified = "ModifiedWordListParagraphs.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");

                // Initialize the body with the document
                var body = new FileFormat.Words.Body(doc);

                foreach (var paragraph in body.Paragraphs)
                {
                    if (paragraph.Style == "ListParagraph")
                    {
                        paragraph.IsNumbered = true;
                        paragraph.AddRun(new FileFormat.Words.IElements.Run
                        { Text = " (numbering type changed to numeric)", Italic = true });
                        doc.Update(paragraph);
                    }
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

