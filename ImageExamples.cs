namespace FileFormat.Words.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Word images
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a> library.
    /// </summary>
    /// <example>
    /// <code>
    /// // Prepares directory Documents/Image at the root of your project.
    /// // Check reference for more options and details.
    /// ImageExamples imageExamples = new ImageExamples();
    /// // Reads images from the specified directory, creates and saves word document to the specified 
    /// // directory. Check reference for more options and details.
    /// imageExamples.CreateWordDocumentWithImages();
    /// // Read Images from the specified Word Document and displays image metadata.
    /// // Check reference for more options and details.
    /// imageExamples.ReadImagesInWordDocument();
    /// // Modify Images in the specified Word Document and saves the modified word document.
    /// // Check reference for more options and details.
    /// imageExamples.ModifyImagesInWordDocument();
    /// </code>
    /// </example>
    public class ImageExamples
    {
        private const string docsDirectory = "../../../Documents/Image";
        private const string imagesDirectory = "../../../Documents/Image/Images";
        /// <summary>
        /// Initializes a new instance of the <see cref="ImageExamples"/> class.
        /// Prepares the directory 'Documents/Image' for storing or loading Word documents
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// Prepares the directory 'Documents/Image/Images' to store images to be added
        /// to the word documents.
        /// </summary>
        public ImageExamples()
        {
            if (!System.IO.Directory.Exists(docsDirectory))
            {
                // If it doesn't exist, create the directory
                System.IO.Directory.CreateDirectory(docsDirectory);
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(docsDirectory)}' created successfully.");
            }
            else
            {
                var files = System.IO.Directory.GetFiles(System.IO.Path.GetFullPath(docsDirectory));
                foreach (var file in files)
                {
                    System.IO.File.Delete(file);
                    System.Console.WriteLine($"File deleted: {file}");
                }
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(docsDirectory)}' cleaned up.");
            }
            if (!System.IO.Directory.Exists(imagesDirectory))
            {
                // If it doesn't exist, create the directory
                System.IO.Directory.CreateDirectory(imagesDirectory);
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(docsDirectory)}' created successfully.");
                using (System.Net.WebClient webClient = new System.Net.WebClient())
                {
                    try
                    {
                        webClient.DownloadFile("https://i.imgur.com/V8aRisV.jpg",
                            $"{imagesDirectory}/image1.jpg");
                        System.Console.WriteLine($"First image downloaded...");
                        webClient.DownloadFile("https://i.imgur.com/xrbdI7n.png",
                            $"{imagesDirectory}/image2.png");
                        System.Console.WriteLine($"Second image downloaded...");
                        webClient.DownloadFile("https://i.imgur.com/bqzDqUZ.png",
                            $"{imagesDirectory}/image3.png");
                        System.Console.WriteLine($"Third image downloaded...");
                    }
                    catch (System.Exception ex)
                    {
                        System.Console.WriteLine($"Error downloading image: {ex.Message}");
                    }
                }
            }
            else
            {
                var files = System.IO.Directory.GetFiles(System.IO.Path.GetFullPath(imagesDirectory));
                foreach (var file in files)
                {
                    System.IO.File.Delete(file);
                    System.Console.WriteLine($"File deleted: {file}");
                }
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(imagesDirectory)}' cleaned up.");
                using (System.Net.WebClient webClient = new System.Net.WebClient())
                {
                    try
                    {
                        webClient.DownloadFile("https://i.imgur.com/V8aRisV.jpg",
                            $"{imagesDirectory}/image1.jpg");
                        System.Console.WriteLine($"First image downloaded...");
                        webClient.DownloadFile("https://i.imgur.com/xrbdI7n.png",
                            $"{imagesDirectory}/image2.png");
                        System.Console.WriteLine($"Second image downloaded...");
                        webClient.DownloadFile("https://i.imgur.com/bqzDqUZ.png",
                            $"{imagesDirectory}/image3.png");
                        System.Console.WriteLine($"Third image downloaded...");
                    }
                    catch (System.Exception ex)
                    {
                        System.Console.WriteLine($"Error downloading image: {ex.Message}");
                    }
                }
            }
        }
        /// <summary>
        /// Creates a new Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a>.
        /// Loads images from the specified diretory and decodes using SkiaSharp.
        /// Creates a word document, appends loaded images and then saves the word document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document will be saved (default is the 'Documents/Image' directory auto-created at the root of your project).
        /// </param>
        /// <param name="imageDirectory">
        /// The directory from where the images will be loaded (default is "Documents/Image/Images").
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file (default is "WordImages.docx").
        /// </param>
        public void CreateWordDocumentWithImages(string documentDirectory = docsDirectory,
            string imageDirectory = imagesDirectory,
            string filename = "WordImages.docx")
        {
            try
            {
                // Initialize a new word document with the default template
                var doc = new FileFormat.Words.Document();

                // Initialize the body with the new document
                var body = new FileFormat.Words.Body(doc);

                // Load images from the specified directory
                var imageFiles = System.IO.Directory.GetFiles(imageDirectory);
                foreach(var imageFile in imageFiles)
                {
                    // Decode the image with SkiaSharp
                    using (var skBMP = SkiaSharp.SKBitmap.Decode(imageFile))
                    {
                        using (var skIMG = SkiaSharp.SKImage.FromBitmap(skBMP))
                        {
                            var encoded = skIMG.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                            // Initialize the word document image element
                            var img = new FileFormat.Words.IElements.Image();
                            // Load data for the word document image element
                            img.ImageData = encoded.ToArray();
                            img.Height = 350;
                            img.Width = 300;
                            // Append image element to the word document
                            body.AppendChild(img);
                            System.Console.WriteLine($"Image {System.IO.Path.GetFullPath(imageFile)}  added to the word document.");
                        }
                    }   
                }
                // Save the newly created Word Document.
                doc.Save($"{documentDirectory}/{filename}");
                System.Console.WriteLine($"Word Document {filename} Created. " +
                    $"Please check directory: {System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Words.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Loads a Word Document with structured content using 
        /// <a href="https://www.nuget.org/packages/FileFormat.Words">FileFormat.Words</a>.
        /// Traverses images and displays image metadata.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present
        /// (default is the 'Documents/Paragraph' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to load (default is "WordImages.docx").
        /// </param>
        public void ReadImagesInWordDocument(string documentDirectory = docsDirectory, string filename = "WordImages.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);
                var num = 0;

                // Traverse images and display metadata info
                foreach (var img in body.Images)
                {
                    num++;
                    System.Console.WriteLine($" Image Number: {num}");
                    System.Console.WriteLine($" Image Data Length: {img.ImageData.Length}");
                    System.Console.WriteLine($" Image Height: {img.Height}");
                    System.Console.WriteLine($" Image Width: {img.Width}");
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
        /// Gets images from the word document. Decodes image using SkiaSharp and encode to JPG.
        /// Resize image to 250(height) and 200(width).
        /// Saves the modified Word Document.
        /// </summary>
        /// <param name="documentDirectory">
        /// The directory where the Word Document to load is present and
        /// the modified document will be saved (default is the 'Documents/Paragraph' directory auto-created at the root of your project).
        /// </param>
        /// <param name="filename">
        /// The name of the Word Document file to modify (default is "WordImages.docx").
        /// </param>
        /// <param name="filenameModified">
        /// The name of the modified Word Document (default is "ModifiedWordImages.docx").
        /// </param>
        public void ModifyImagesInWordDocument(string documentDirectory = docsDirectory,
            string filename = "WordImages.docx", string filenameModified = "ModifiedWordImages.docx")
        {
            try
            {
                // Load the Word Document
                var doc = new FileFormat.Words.Document($"{documentDirectory}/{filename}");
                var body = new FileFormat.Words.Body(doc);

                foreach (var img in body.Images)
                {
                    var skBitmap = SkiaSharp.SKBitmap.Decode(img.ImageData);
                    var skImage = SkiaSharp.SKImage.FromBitmap(skBitmap);
                    var encoded = skImage.Encode(SkiaSharp.SKEncodedImageFormat.Jpeg, 100);

                    img.ImageData = encoded.ToArray();

                    img.Height = 250;
                    img.Width = 200;

                    doc.Update(img);
                }

                // Save the modified Word Document
                doc.Save($"{documentDirectory}/{filenameModified}");
                System.Console.WriteLine($"Word Document {filename} Modified and Saved As {filenameModified}. Please check directory: {System.IO.Path.GetFullPath(documentDirectory)}");
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Words.FileFormatException("An error occurred.", ex);
            }
        }
    }
}
