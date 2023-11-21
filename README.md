# FileFormat.Words for .NET Gists
## [Create Word Document Paragraphs in C#](https://gist.github.com/fileformat-words-gists/0f5c7fa92216dec7c8b1b07f5a8060ea)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates paragraphs with heading styles defined by the Word document template.
* Adds normal paragraphs under each heading paragraph.
* Includes text runs with various fonts as per the template.
* Saves the newly created Word Document.        
## [Read Word Document Paragraphs in C#](https://gist.github.com/fileformat-words-gists/84eb759e58049ddc28c25943d2d3c121)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses paragraphs and displays associated styles as defined by the Word document template.
* Traverses through each run (text fragment) within each paragraph and displays fragment values.
## [Modify Word Document Paragraphs in C#](https://gist.github.com/fileformat-words-gists/53dbf77cd1168f06320f4b1a447bc4d1)
* Loads an existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Modifies paragraphs by prepending 'Modified Heading :' for styled paragraphs and 'Modified Run :' for each run within normal paragraphs, preserving the existing format.
* Saves the modified Word Document.
