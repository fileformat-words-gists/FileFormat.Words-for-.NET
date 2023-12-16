using System;
using FileFormat.Words.Examples;

namespace FileFormat.Words.Examples.Usage
{
    class Program
    {
        static void Main(string[] args)
        {
            // Prepares directory Documents/Paragraph at the root of your project.
            // Check reference for more options and details.
            ParagraphExamples paragraphExamples = new ParagraphExamples();
            // Creates a word document with paragraphs and saves word document to the specified 
            // directory. Check reference for more options and details.
            paragraphExamples.CreateWordParagraphs();
            // Reads Paragraphs from the specified Word Document and displays plain text and formatting.
            // Check reference for more options and details.
            paragraphExamples.ReadWordParagraphs();
            // Modifies Paragraphs in the specified Word Document and saves the modified word document.
            // Check reference for more options and details.
            paragraphExamples.ModifyWordParagraphs();

            // Prepares directory Documents/Image at the root of your project.
            // Check reference for more options and details.
            ImageExamples imageExamples = new ImageExamples();
            // Reads images from the specified directory, creates and saves word document to the specified 
            // directory. Check reference for more options and details.
            imageExamples.CreateWordDocumentWithImages();
            // Read Images from the specified Word Document and displays image metadata.
            // Check reference for more options and details.
            imageExamples.ReadImagesInWordDocument();
            // Modify Images in the specified Word Document and saves the modified word document.
            // Check reference for more options and details.
            imageExamples.ModifyImagesInWordDocument();

            // Prepares directory Documents/Table at the root of your project.
            // Check reference for more options and details.
            TableExamples tableExamples = new TableExamples();
            // Creates a word document with tables and saves word document to the specified 
            // directory. Check reference for more options and details.
            tableExamples.CreateWordDocumentWithTables();
            // Read tables from the specified Word Document and displays table contents.
            // Check reference for more options and details.
            tableExamples.ReadTablesInWordDocument();
            // Modify Images in the specified Word Document and saves the modified word document.
            // Check reference for more options and details.
            tableExamples.ModifyTablesInWordDocument();
        }
    }
}
