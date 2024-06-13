# ⚠️ Important Notice

**This repository is no longer maintained.**

We have moved our code and updates to a new repository. Please visit the new repository for the latest version:

👉 **New Repository: [Openize.Words for .NET](https://github.com/openize-words-gists/Openize.Words-for-.NET)**

Thank you for your interest and support!

# C# Word Document API Examples

**FileFormat.Words.Examples** - C# code examples using [FileFormat.Words for .NET](https://github.com/fileformat-words/FileFormat.Words-for-.NET) - A versatile API for creating, loading, and modifying MS Word documents.

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
- [Multiple Word Documents Concurrent Updating in C#](#multiple-word-documents-concurrent-updating-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/efe7dfee71953a577e13988ed99e20f1)
- [Create Word Paragraph Alignment in C#](#create-word-paragraph-alignment-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/4cabcd2dd727ebb5dc1c27104b02b1bd)
- [Read Word Paragraph Alignment in C#](#read-word-paragraph-alignment-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/bd68c603de13475f65e2369969072b46)
- [Modify Word Paragraph Alignment in C#](#modify-word-paragraph-alignment-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/bc539ee85e5ad873ed27c5d1c7ff41b2)
- [Create Word Paragraph Indent in C#](#create-word-paragraph-indent-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/87676cf05f3fe03a30a8a08087ce3faf)
- [Read Word Paragraph Indent in C#](#read-word-paragraph-indent-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/951eb4906ef57d8d6b93b1f009700a54)
- [Modify Word Paragraph Indent in C#](#modify-word-paragraph-indent-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/95efc98b9e1b579c3201309697c4ef97)
- [Create Numbered Word Paragraphs in C#](#create-numbered-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/ea4a47288075c796ff5dd6bb97fccf1f)
- [Read Numbered Word Paragraphs in C#](#read-numbered-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/03003fe62a750d20fc3c3351c401fce1)
- [Modify Numbered Word Paragraphs in C#](#modify-numbered-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/54529e570031b98be281539a81713ef1)
- [Create Roman Alphabetic Word Paragraphs in C#](#create-roman-alphabetic-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/fbef77098e338fc29d2e5ba5108f0169)
- [Read Roman Alphabetic Word Paragraphs in C#](#read-roman-alphabetic-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/291df28a345ff6556c6c1f19f2b08a09)
- [Modify Roman Alphabetic Word Paragraphs in C#](#modify-roman-alphabetic-word-paragraphs-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/45a28fd403bfc13157e8f937641d80a1)
- [Create Multiple Multilevel List Paragraphs of Word Document in C#](#create-multiple-multilevel-list-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/28b1e2ba2d553059a2b85031208a2a2a)
- [Read Multiple Multilevel List Paragraphs of Word Document in C#](#read-multiple-multilevel-list-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/e0a00f9068eef510f44acdf064910f95)
- [Modify Multiple Multilevel List Paragraphs of Word Document in C#](#modify-multiple-multilevel-list-paragraphs-of-word-document-in-c) - Explore [gist](https://gist.github.com/fileformat-words-gists/6abca4875309fac7605518ac368de4c2)
- [Resources](#resources)
- [System Requirements](#system-requirements)
- [Quick Start](#quick-start)
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

## [Multiple Word Documents Concurrent Updating in C#](https://gist.github.com/fileformat-words-gists/efe7dfee71953a577e13988ed99e20f1)
* Loads 3 existing Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Appends paragraphs concurrently on three documents
* Saves the modified Word Documents.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/efe7dfee71953a577e13988ed99e20f1)

## [Create Word Paragraph Alignment in C#](https://gist.github.com/fileformat-words-gists/4cabcd2dd727ebb5dc1c27104b02b1bd)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates paragraphs with different alignments including left, center, right and justify.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/4cabcd2dd727ebb5dc1c27104b02b1bd)

## [Read Word Paragraph Alignment in C#](https://gist.github.com/fileformat-words-gists/bd68c603de13475f65e2369969072b46)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses paragraphs and displays its text along with alignment.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/bd68c603de13475f65e2369969072b46)

## [Modify Word Paragraph Alignment in C#](https://gist.github.com/fileformat-words-gists/bc539ee85e5ad873ed27c5d1c7ff41b2)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses through all paragraphs within the document.
* Modifies paragraphs by appending ' (alignment modified to justify)' with italic format and justify alignment.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/bc539ee85e5ad873ed27c5d1c7ff41b2)

## [Create Word Paragraph Indent in C#](https://gist.github.com/fileformat-words-gists/87676cf05f3fe03a30a8a08087ce3faf)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates paragraphs with different indents including left, right, firstline and hanging.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/87676cf05f3fe03a30a8a08087ce3faf)

## [Read Word Paragraph Indent in C#](https://gist.github.com/fileformat-words-gists/951eb4906ef57d8d6b93b1f009700a54)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses paragraphs and displays its text along with indentation.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/951eb4906ef57d8d6b93b1f009700a54)

## [Modify Word Paragraph Indent in C#](https://gist.github.com/fileformat-words-gists/95efc98b9e1b579c3201309697c4ef97)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses through all paragraphs within the document.
* Modifies paragraphs by appending the relevant indent message in italic format and modifies the indent (if found) to 0.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/95efc98b9e1b579c3201309697c4ef97)

## [Create Numbered Word Paragraphs in C#](https://gist.github.com/fileformat-words-gists/ea4a47288075c796ff5dd6bb97fccf1f)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates numbered paragraphs with nested levels.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/ea4a47288075c796ff5dd6bb97fccf1f)

## [Read Numbered Word Paragraphs in C#](https://gist.github.com/fileformat-words-gists/03003fe62a750d20fc3c3351c401fce1)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses paragraphs and displays its text, numbering and level.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/03003fe62a750d20fc3c3351c401fce1)

## [Modify Numbered Word Paragraphs in C#](https://gist.github.com/fileformat-words-gists/54529e570031b98be281539a81713ef1)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses through all paragraphs within the document.
* If numbered, modifies paragraphs by appending ' (numering removed)' with italic format and paragraph number is removed.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/54529e570031b98be281539a81713ef1)

## [Create Roman Alphabetic Word Paragraphs in C#](https://gist.github.com/fileformat-words-gists/fbef77098e338fc29d2e5ba5108f0169)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates roman and alphabetic paragraphs with nested levels.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/fbef77098e338fc29d2e5ba5108f0169)

## [Read Roman Alphabetic Word Paragraphs in C#](https://gist.github.com/fileformat-words-gists/291df28a345ff6556c6c1f19f2b08a09)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses paragraphs and displays its text, roman/alphabetic status and level.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/291df28a345ff6556c6c1f19f2b08a09)

## [Modify Roman Alphabetic Word Paragraphs in C#](https://gist.github.com/fileformat-words-gists/45a28fd403bfc13157e8f937641d80a1)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses through all paragraphs within the document.
* If alphabetic, modifies paragraphs by appending ' (alphabetic removed)' with italic format and paragraph alphabetic is removed.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/45a28fd403bfc13157e8f937641d80a1)

## [Create Multiple Multilevel List Paragraphs of Word Document in C#](https://gist.github.com/fileformat-words-gists/28b1e2ba2d553059a2b85031208a2a2a)
* Creates a new Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Generates two multilevel lists with different prefixes at different levels.
* Saves the newly created Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/28b1e2ba2d553059a2b85031208a2a2a)

## [Read Multiple Multilevel List Paragraphs of Word Document in C#](https://gist.github.com/fileformat-words-gists/e0a00f9068eef510f44acdf064910f95)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses paragraphs and displays its text, numbering id, numbering type and level.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/e0a00f9068eef510f44acdf064910f95)

## [Modify Multiple Multilevel List Paragraphs of Word Document in C#](https://gist.github.com/fileformat-words-gists/6abca4875309fac7605518ac368de4c2)
* Loads a Word Document with structured content using [FileFormat.Words](https://www.nuget.org/packages/FileFormat.Words)
* Traverses through all paragraphs within the document.
* If list paragraphs, modifies paragraphs by appending ' (numbering type changed to numeric)' with italic format and paragraph numbering type is changed to numeric.
* Saves the modified Word Document.
* [Click here to explore gist](https://gist.github.com/fileformat-words-gists/6abca4875309fac7605518ac368de4c2)

## Resources
* [Docs](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/index.html)
* [API Reference](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/api/index.html)
* [Articles](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/articles/index.html)

## System Requirements
* [Pre-Requisite](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/api/index.html#pre-requisite)
* Make sure to have below nuget packages installed:
  * [FileFormat.Words for .NET](https://www.nuget.org/packages/FileFormat.Words)
  * [SkiaSharp](https://www.nuget.org/packages/SkiaSharp)

## Quick Start
* [Create, Read and Modify Word Paragraphs](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.Examples.ParagraphExamples.html#FileFormat_Words_Examples_ParagraphExamples_examples).
* [Create, Read and Modify Word Images](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.Examples.ImageExamples.html#FileFormat_Words_Examples_ImageExamples_examples).
* [Create, Read and Modify Word Tables](https://fileformat-words-gists.github.io/FileFormat.Words-for-.NET/api/FileFormat.Words.Examples.TableExamples.html#FileFormat_Words_Examples_TableExamples_examples).

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
