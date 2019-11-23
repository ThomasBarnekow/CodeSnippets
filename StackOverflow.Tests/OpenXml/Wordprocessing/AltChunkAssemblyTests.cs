//
// AltChunkAssemblyTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace StackOverflow.Tests.OpenXml.Wordprocessing
{
    /// <summary>
    /// This class demonstrates how to use AlternativeFormatImportParts and AltChunk
    /// instances to merge multiple Word documents into one. Further, it shows how
    /// to transform the component Word documents by replacing the content control
    /// contents before inserting them into the target document.
    /// </summary>
    /// <remarks>
    /// See https://stackoverflow.com/questions/58995378/how-to-modify-content-in-filestream-while-using-open-xml
    /// </remarks>
    public class AltChunkAssemblyTests
    {
        // Sample template file names for unit testing purposes.
        private readonly string[] _templateFileNames =
        {
            "report-Part1.docx",
            "report-Part2.docx",
            "report-Part3.docx"
        };

        // Sample content maps for unit testing purposes.
        // Each Dictionary<string, string> represents data used to replace the
        // content of block-level w:sdt elements identified by w:tag values of
        // "firstTag" and "secondTag".
        private readonly List<Dictionary<string, string>> _contentMaps = new List<Dictionary<string, string>>
        {
            new Dictionary<string, string>
            {
                { "firstTag", "report-Part1: First value" },
                { "secondTag", "report-Part1: Second value" }
            },
            new Dictionary<string, string>
            {
                { "firstTag", "report-Part2: First value" },
                { "secondTag", "report-Part2: Second value" }
            },
            new Dictionary<string, string>
            {
                { "firstTag", "report-Part3: First value" },
                { "secondTag", "report-Part3: Second value" }
            }
        };

        [Fact]
        public void CanAssembleDocumentUsingAltChunks()
        {
            // Create some sample "templates" (technically documents) for unit
            // testing purposes.
            CreateSampleTemplates();

            // Create a an empty result document.
            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                "AltChunk.docx", WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            var body = new Body();
            mainPart.Document = new Document(body);

            // Add one w:altChunk element for each sample template, using the
            // sample content maps for mapping sample data to the content
            // controls contained in the templates.
            for (var index = 0; index < 3; index++)
            {
                if (index > 0) body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                body.AppendChild(CreateAltChunk(_templateFileNames[index], _contentMaps[index], wordDocument));
            }
        }

        private void CreateSampleTemplates()
        {
            // Create a sample template for each sample template file names.
            foreach (string templateFileName in _templateFileNames)
            {
                CreateSampleTemplate(templateFileName);
            }
        }

        private static void CreateSampleTemplate(string templateFileName)
        {
            // Create a new Word document with paragraphs marking the start and
            // end of the template (for testing purposes) and two block-level
            // structured document tags identified by w:tag elements with values
            // "firstTag" and "secondTag" and values that are going to be
            // replaced by the ContentControlWriter during document assembly.
            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                templateFileName, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document =
                new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text($"Start of template '{templateFileName}'"))),
                        new SdtBlock(
                            new SdtProperties(
                                new Tag { Val = "firstTag" }),
                            new SdtContentBlock(
                                new Paragraph(
                                    new Run(
                                        new Text("First template value"))))),
                        new SdtBlock(
                            new SdtProperties(
                                new Tag { Val = "secondTag" }),
                            new SdtContentBlock(
                                new Paragraph(
                                    new Run(
                                        new Text("Second template value"))))),
                        new Paragraph(
                            new Run(
                                new Text($"End of template '{templateFileName}'")))));
        }

        private static AltChunk CreateAltChunk(
            string templateFileName,
            Dictionary<string, string> contentMap,
            WordprocessingDocument wordDocument)
        {
            // Copy the template file contents to a MemoryStream to be able to
            // update the content controls without altering the template file.
            using FileStream fileStream = File.Open(templateFileName, FileMode.Open);
            using var memoryStream = new MemoryStream();
            fileStream.CopyTo(memoryStream);

            // Open the copy of the template on the MemoryStream, update the
            // content controls, save the updated template back to the
            // MemoryStream, and reset the position within the MemoryStream.
            using (WordprocessingDocument chunkDocument = WordprocessingDocument.Open(memoryStream, true))
            {
                var contentControlWriter = new ContentControlWriter(contentMap);
                contentControlWriter.WriteContentControls(chunkDocument);
            }

            memoryStream.Seek(0, SeekOrigin.Begin);

            // Create an AlternativeFormatImportPart from the MemoryStream.
            string altChunkId = "AltChunkId" + Guid.NewGuid();
            AlternativeFormatImportPart chunk = wordDocument.MainDocumentPart.AddAlternativeFormatImportPart(
                AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            chunk.FeedData(memoryStream);

            // Return the w:altChunk element to be added to the w:body element.
            return new AltChunk { Id = altChunkId };
        }
    }
}
