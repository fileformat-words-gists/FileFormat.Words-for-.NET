# C# Word Document API Examples

**FileFormat.Words for .NET Gists** - C# code examples for using [FileFormat.Words for .NET](https://github.com/fileformat-words/FileFormat.Words-for-.NET) - A versatile API for creating, loading, and modifying MS Word documents.

## Table of Contents
- [Create Word Document Paragraphs in C#](#create-word-document-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/0f5c7fa92216dec7c8b1b07f5a8060ea)
- [Read Word Document Paragraphs in C#](#read-word-document-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/84eb759e58049ddc28c25943d2d3c121)
- [Modify Word Document Paragraphs in C#](#modify-word-document-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/53dbf77cd1168f06320f4b1a447bc4d1)
- [Create Word Document Images in C#](#create-word-document-images-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/cae0acdf7e5ef5f177402e4742aadc3d)
- [Read Word Document Images in C#](#read-word-document-images-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/ad5c621c8764eb69555f2ab339f3ea01)
- [Modify Word Document Images in C#](#modify-word-document-images-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/1c1eaf2878e5e25717561a3f3cbe43d6)
- [Create Word Document Tables in C#](#create-word-document-tables-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/c05f8f128080801fff348a41e38d0364)
- [Read Word Document Tables in C#](#read-word-document-tables-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/fb5d9fa3c0576b45140ee3be87405c79)
- [Modify Word Document Tables in C#](#modify-word-document-tables-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/20db884541540196dd00a9f313d9f77b)
- [Resources](#resources)
- [System Requirements](system-requirements)
- [License](#license)
  
## [Create Word Document Paragraphs in C#](https://gist.github.com/fileformat-words-gists/0f5c7fa92216dec7c8b1b07f5a8060ea)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates paragraphs with heading styles defined by the Word document template.
* Adds normal paragraphs under each heading paragraph.
* Includes text runs with various fonts as per the template.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/0f5c7fa92216dec7c8b1b07f5a8060ea).        

## [Read Word Document Paragraphs in C#](https://gist.github.com/fileformat-words-gists/84eb759e58049ddc28c25943d2d3c121)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses paragraphs and displays associated styles as defined by the Word document template.
* Traverses through each run (text fragment) within each paragraph and displays fragment values.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/84eb759e58049ddc28c25943d2d3c121)

## [Modify Word Document Paragraphs in C#](https://gist.github.com/fileformat-words-gists/53dbf77cd1168f06320f4b1a447bc4d1)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Modifies paragraphs by prepending 'Modified Heading :' for styled paragraphs and 'Modified Run :' for each run within normal paragraphs, preserving the existing format.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/53dbf77cd1168f06320f4b1a447bc4d1)

## [Create Word Document Images in C#](https://gist.github.com/fileformat-words-gists/cae0acdf7e5ef5f177402e4742aadc3d)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Loads images from the specified diretory and decodes using SkiaSharp.
* Creates a word document and appends loaded images to it.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/cae0acdf7e5ef5f177402e4742aadc3d)

## [Read Word Document Images in C#](https://gist.github.com/fileformat-words-gists/ad5c621c8764eb69555f2ab339f3ea01)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses images and displays image metadata.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/ad5c621c8764eb69555f2ab339f3ea01)

## [Modify Word Document Images in C#](https://gist.github.com/fileformat-words-gists/1c1eaf2878e5e25717561a3f3cbe43d6)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Gets images from the word document. Decodes image using SkiaSharp and encode to JPG.
* Resize image to 250(height) and 200(width).
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/1c1eaf2878e5e25717561a3f3cbe43d6)

## [Create Word Document Tables in C#](https://gist.github.com/fileformat-words-gists/c05f8f128080801fff348a41e38d0364)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates 5(rows) x 3(cols) tables with table styles defined by the Word document template.
* Appends each table to the body of the word document.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/c05f8f128080801fff348a41e38d0364)      

## [Read Word Document Tables in C#](https://gist.github.com/fileformat-words-gists/fb5d9fa3c0576b45140ee3be87405c79)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses tables and displays associated styles as defined by the Word document template.
* Traverses through each row and then traverses columns within the row.
* Traverses through paragrpahs within each cell and displays paragraph plain text
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/fb5d9fa3c0576b45140ee3be87405c79)

## [Modify Word Document Tables in C#](https://gist.github.com/fileformat-words-gists/20db884541540196dd00a9f313d9f77b)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Modifies tables by setting column widths to 2000
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/20db884541540196dd00a9f313d9f77b)

## Resources
* [Docs](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/index.html)
* [API Reference](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/api/index.html)
* [Articles](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/articles/index.html)

## System Requirements
* [Pre-Requisite](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/api/index.html#pre-requisite)
* Make sure to have below nuget packages installed:
  * [FileFormat.Words for .NET](https://www.nuget.org/packages/FileFormat.Words)
  * [SkiaSharp](https://www.nuget.org/packages/SkiaSharp)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
