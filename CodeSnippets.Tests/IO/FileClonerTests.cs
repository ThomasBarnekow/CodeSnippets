//
// FileClonerTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Collections.Generic;
using System.IO;
using System.Linq;
using CodeSnippets.IO;
using CodeSnippets.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Tests.IO
{
    public class FileClonerTests
    {
        private const int ParagraphCount = 100;
        private const string Text = "The quick brown fox jumps over the lazy dog.";

        [Fact]
        public void ReadAllBytesToMemoryStream_TestDocument_SuccessfullyCloned()
        {
            // Arrange.
            const string path = "ReadAllBytesToMemoryStream_Source.docx";
            CreateTestDocument(path);

            // Act.
            MemoryStream destStream = FileCloner.ReadAllBytesToMemoryStream(path);

            // Assert.
            AssertDocumentHasExpectedContents(destStream);
        }

        [Fact]
        public void CopyFileStreamToMemoryStream_TestDocument_SuccessfullyCloned()
        {
            // Arrange.
            const string path = "CopyFileStreamToMemoryStream_Source.docx";
            CreateTestDocument(path);

            // Act.
            MemoryStream destStream = FileCloner.CopyFileStreamToMemoryStream(path);

            // Assert.
            AssertDocumentHasExpectedContents(destStream);
        }

        [Fact]
        public void CopyFileStreamToFileStream_TestDocument_SuccessfullyCloned()
        {
            // Arrange.
            const string sourcePath = "CopyFileStreamToFileStream_Source.docx";
            const string destPath = "CopyFileStreamToFileStream_Destination.docx";

            CreateTestDocument(sourcePath);
            File.Delete(destPath);

            // Act.
            FileStream destStream = FileCloner.CopyFileStreamToFileStream(sourcePath, destPath);

            // Assert.
            AssertDocumentHasExpectedContents(destStream);
        }

        [Fact]
        public void CopyFileAndOpenFileStream_TestDocument_SuccessfullyCloned()
        {
            // Arrange.
            const string sourcePath = "CopyFileAndOpenFileStream_Source.docx";
            const string destPath = "CopyFileAndOpenFileStream_Destination.docx";

            CreateTestDocument(sourcePath);
            File.Delete(destPath);

            // Act.
            FileStream destStream = FileCloner.CopyFileAndOpenFileStream(sourcePath, destPath);

            // Assert.
            AssertDocumentHasExpectedContents(destStream);
        }

        private static void CreateTestDocument(string path, int count = ParagraphCount, string text = Text)
        {
            using WordprocessingDocument unused = WordprocessingDocumentFactory.Create(
                path, Enumerable.Range(0, count).Select(i => text));
        }

        private static void AssertDocumentHasExpectedContents(Stream stream)
        {
            // Open the WordprocessingDocument on the FileStream and assert
            // that it has the expected contents.
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true);
            Body body = wordDocument.MainDocumentPart.Document.Body;
            List<Paragraph> paragraphs = body.Elements<Paragraph>().ToList();

            Assert.Equal(ParagraphCount, paragraphs.Count);
            Assert.All(paragraphs, p => Assert.Equal(Text, p.InnerText));
        }
    }
}
