namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word tables
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Shape at the root of your project.
    /// // Check reference for more options and details.
    /// var shapeExamples = new FileFormat.Words.Examples.ShapeExamples();
    /// // Creates a word document with shapes and saves word document to the specified 
    /// // directory. Check reference for more options and details.
    /// shapeExamples.CreateShapes();
    /// // Reads shapes from the specified Word Document and displays shape attributes.
    /// // Check reference for more options and details.
    /// shapeExamples.ReadShapes();
    /// // Modifies shapes in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// shapeExamples.ModifyShapes();
    /// </code>
    /// </example>
    public class ShapeExamples
    {
        private const string docsDirectory = "../../../Documents/Shape";
        /// <summary>
        /// Initializes a new instance of the <see cref="ShapeExamples"/> class.
        /// Prepares the directory 'Documents/Table' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public ShapeExamples()
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
        public void CreateShapes(string documentDirectory = docsDirectory,
            string filename = "WordShapes.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new FileFormat.Words.Document();
                System.Console.WriteLine("Word Document with default template initialized");

                // Initialize the body with the new document
                var body = new FileFormat.Words.Body(doc);
                System.Console.WriteLine("Body of the Word Document initialized");

                // Instantiate shape element with hexagone and coordinates/size.
                var shape = new FileFormat.Words.IElements.Shape(100, 100, 400, 400,
                FileFormat.Words.IElements.ShapeType.Hexagone);
                // Add hexagone shape to the word document.
                body.AppendChild(shape);
                System.Console.WriteLine("Hexagone shape added");

                // Reinstantiate shape element with diamond and coordinates/size.
                shape = new FileFormat.Words.IElements.Shape(100, 100, 400, 400,
                FileFormat.Words.IElements.ShapeType.Diamond);
                // Add daimond shape to the word document.
                body.AppendChild(shape);
                System.Console.WriteLine("Diamond shape added");

                // Reinstantiate shape element with ellipse and coordinates/size.
                shape = new FileFormat.Words.IElements.Shape(100, 100, 400, 400,
                FileFormat.Words.IElements.ShapeType.Ellipse);
                // Add ellipse shape to the word document.
                body.AppendChild(shape);
                System.Console.WriteLine("Ellipse shape added");

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
        /// Traverses through shapes of the Word document.
        /// Reads and displays properties of the shape.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Shape' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordShapes.docx").
        /// </param>
        public void ReadShapes(string documentDirectory = docsDirectory,
            string filename = "WordShapes.docx")
        {
            try
            {
                // Load the Word Document.
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                // Initialize the body with the loaded document.
                var body = new FileFormat.Words.Body(doc);

                // Load all shapes with the document
                var shapes = body.Shapes;

                // Initialize the shape counter
                var shapeNumber = 0;

                // Traverse through each shape and display its properties
                foreach (var shape in shapes)
                {
                    shapeNumber++;
                    System.Console.WriteLine($"Shape Number : {shapeNumber}");
                    System.Console.WriteLine($"...Shape Type : {shape.Type}");
                    System.Console.WriteLine($"...X Position : {shape.X}");
                    System.Console.WriteLine($"...Y Position : {shape.Y}");
                    System.Console.WriteLine($"...Width      : {shape.Width}");
                    System.Console.WriteLine($"...Height     : {shape.Height}");
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
        /// Traverses through shapes of the Word document.
        /// Modifies shapes by setting their type to Diamond.
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Shape' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordShapes.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordShapes.docx").
        /// </param>
        public void ModifyShapes(string documentDirectory = docsDirectory,
            string filename = "WordShapes.docx", string filenameModified = "ModifiedWordShapes.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                // Initialize the body with the loaded document. 
                var body = new FileFormat.Words.Body(doc);

                // Load all shapes
                var shapes = body.Shapes;

                // Traverse through each shape, change the shape type to diamond and update document.
                foreach (var shape in shapes)
                {
                    shape.Type = FileFormat.Words.IElements.ShapeType.Diamond;
                    doc.Update(shape);
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
