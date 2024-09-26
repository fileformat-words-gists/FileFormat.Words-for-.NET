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
            var paragraphExamples = new ParagraphExamples();
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
            var imageExamples = new ImageExamples();
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
            var tableExamples = new TableExamples();
            // Creates a word document with tables and saves word document to the specified 
            // directory. Check reference for more options and details.
            tableExamples.CreateWordDocumentWithTables();
            // Read tables from the specified Word Document and displays table contents.
            // Check reference for more options and details.
            tableExamples.ReadTablesInWordDocument();
            // Modify Images in the specified Word Document and saves the modified word document.
            // Check reference for more options and details.
            tableExamples.ModifyTablesInWordDocument();

            // Prepares directory Documents/Paragraph/Alignment at the root of your project.
            // Check reference for more options and details.
            var paragraphAlignment = new ParagraphAlignmentExamples();
            // Creates a word document with paragraphs having different alignments and saves word document
            // to the specified directory. Check reference for more options and details.
            paragraphAlignment.CreateAlignment();
            // Reads Paragraphs from the specified Word Document and displays plain text alongwith alignment.
            // Check reference for more options and details.
            paragraphAlignment.ReadAlignment();
            // Modifies Paragraph's alignment in the specified Word Document and saves the modified word document.
            // Check reference for more options and details.
            paragraphAlignment.ModifyAlignment();

            // Prepares directory Documents/Paragraph/Indent at the root of your project.
            // Check reference for more options and details.
            var paragraphIndentExamples = new ParagraphIndentExamples();
            // Creates a word document with paragraphs having various indentations and saves word document
            // to the specified directory. Check reference for more options and details.
            paragraphIndentExamples.CreateIndent();
            // Reads Paragraphs from the specified Word Document and displays plain text with indentation.
            // Check reference for more options and details.
            paragraphIndentExamples.ReadIndent();
            // Modifies Paragraph's indentation in the Word Document and saves the modified word document.
            // Check reference for more options and details.
            paragraphIndentExamples.ModifyIndent();

            // Prepares directory Documents/Paragraph/Numbering at the root of your project.
            // Check reference for more options and details.
            var paragraphNumberExamples = new ParagraphNumberExamples();
            // Creates a word document with paragraphs having various numbering levels and saves word
            // document to the specified directory. Check reference for more options and details.
            paragraphNumberExamples.CreateNumberedParagraphs();
            // Reads Paragraphs from the specified Word Document and displays plain text with numbering info.
            // Check reference for more options and details.
            paragraphNumberExamples.ReadNumberedParagraphs();
            // Modifies Paragraph's numbering in the Word Document and saves the modified word document.
            // Check reference for more options and details.
            paragraphNumberExamples.ModifyNumberedParagraphs();

            // Prepares directory Documents/Paragraph/RomanAlphabeitc at the root of your project.
            // Check reference for more options and details.
            var paragraphRomanAlphabeticExamples = new ParagraphRomanAlphabeticExamples();
            // Creates a word document with paragraphs having roman and alphabetic levels and saves word
            // document to the specified directory. Check reference for more options and details.
            paragraphRomanAlphabeticExamples.CreateRomanAlphabeticParagraphs();
            // Reads Paragraphs from the specified Word Document and displays plain text with roman and alphabetic info.
            // Check reference for more options and details.
            paragraphRomanAlphabeticExamples.ReadRomanAlphabeticParagraphs();
            // Modifies Roman and Alphabetic paragraphs in the Word Document and saves the modified word document.
            // Check reference for more options and details.
            paragraphRomanAlphabeticExamples.ModifyRomanAlphabeticParagraphs();
           
            // Prepares directory Documents/Paragraph/List at the root of your project.
            // Check reference for more options and details.
            var listExamples = new ListExamples();
            // Creates a word document with two mulitlevel lists each having 3 levels and saves word
            // document to the specified directory. Check reference for more options and details.
            listExamples.CreateMultilevelLists();
            // Reads Paragraphs from the specified Word Document and displays plain text with prefix info at each level.
            // Check reference for more options and details.
            listExamples.ReadMultilevelLists();
            // Modifies list paragraphs in the Word Document and saves the modified word document.
            // Check reference for more options and details.
            listExamples.ModifyMultilevelLists();

            // Prepares directory Documents/Paragraph/Frame at the root of your project.
            // Check reference for more options and details.
            var paraFrameExamples = new FileFormat.Words.Examples.ParagraphFrameExamples();
            // Creates a word document with 4 different paragraphs with frames + one paragraph without frames and saves word
            // document to the specified directory. Check reference for more options and details.
            paraFrameExamples.CreateParagraphsFrames();
            // Reads Paragraphs from the specified Word Document and displays plain text with borders/frames info.
            // Check reference for more options and details.
            paraFrameExamples.ReadParagraphsFrames();
            // Modifies border and text of framed paragraphs in the Word Document and saves the modified word document.
            // Check reference for more options and details.
            paraFrameExamples.ModifyParagraphsFrames();
            
            // Prepares directory Documents/Shape at the root of your project.
            // Check reference for more options and details.
            var shapeExamples = new FileFormat.Words.Examples.ShapeExamples();
            // Creates a word document with shapes and saves word document to the specified 
            // directory. Check reference for more options and details.
            shapeExamples.CreateShapes();
            // Reads shapes from the specified Word Document and displays shape attributes.
            // Check reference for more options and details.
            shapeExamples.ReadShapes();
            // Modifies shapes in the specified Word Document and saves the modified word document.
            // Check reference for more options and details.
            shapeExamples.ModifyShapes();
        }
    }
}
