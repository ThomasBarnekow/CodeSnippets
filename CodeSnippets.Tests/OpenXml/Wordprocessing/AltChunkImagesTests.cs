//
// AltChunkImagesTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class AltChunkImagesTests
    {
        private static AltChunk CreateAltChunkFromWordDocument(
            string sourcePath,
            WordprocessingDocument destWordDocument)
        {
            string altChunkId = "AltChunkId-" + Guid.NewGuid();

            AlternativeFormatImportPart chunk = destWordDocument.MainDocumentPart.AddAlternativeFormatImportPart(
                AlternativeFormatImportPartType.WordprocessingML, altChunkId);

            using FileStream fs = File.Open(sourcePath, FileMode.Open);
            chunk.FeedData(fs);

            return new AltChunk { Id = altChunkId };
        }

        [Fact]
        public void CanAppendDocumentWithImage()
        {
            const string destinationDoc = "test-destination.docx";
            const string sourceDoc = "Resources/test-source.docx";
            const string appendedDoc = "Resources/test-append.docx";

            File.Copy(sourceDoc, destinationDoc, true);

            using WordprocessingDocument destWordDocument = WordprocessingDocument.Open(destinationDoc, true);
            Body body = destWordDocument.MainDocumentPart.Document.Body;

            Paragraph p2 = body.InsertAfter(
                new Paragraph(
                    new ParagraphProperties(
                        new PageBreakBefore()),
                    new Run(
                        new RunProperties(new Bold()),
                        new TabChar(),
                        new Text(appendedDoc)),
                    new Run(
                        new Break()),
                    new Run(
                        new RunProperties(new Italic()),
                        new TabChar(),
                        new Text("Uploaded by:")),
                    new Run(
                        new RunProperties(new Italic()),
                        new TabChar(),
                        new Text("Test User"))),
                body.Elements<Paragraph>().Last());

            p2.InsertAfterSelf(
                CreateAltChunkFromWordDocument(appendedDoc, destWordDocument));
        }
    }
}
