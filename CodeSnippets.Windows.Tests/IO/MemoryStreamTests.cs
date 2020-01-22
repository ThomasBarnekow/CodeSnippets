//
// MemoryStreamTests.cs
//
// Copyright 2020 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Windows.Tests.IO
{
    public class MemoryStreamTests
    {
        private static byte[] CreateEmptyWordDocument()
        {
            using var stream = new MemoryStream();

            using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
                mainDocumentPart.Document = new Document(new Body(new Paragraph()));
            }

            return stream.ToArray();
        }

        private static void AddParagraphs(Stream stream)
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true);

            IEnumerable<Paragraph> paragraphs = Enumerable
                .Range(1, 10)
                .Select(i => new Paragraph(new Run(new Text($"This is paragraph #{i}."))));

            Body body = wordDocument.MainDocumentPart.Document.Body;
            body.Append(paragraphs);
        }

        [Fact]
        public void CannotIncreaseSizeOfWordprocessingDocumentWithNonResizableMemoryStream()
        {
            byte[] buffer = CreateEmptyWordDocument();

            // Create a non-resizable MemoryStream.
            using var stream = new MemoryStream(buffer);

            // Add paragraphs, increasing the size of the document.
            Assert.Throws<NotSupportedException>(() => AddParagraphs(stream));
        }

        [Fact]
        public void CanIncreaseSizeOfWordprocessingDocumentWithResizableMemoryStream()
        {
            byte[] buffer = CreateEmptyWordDocument();

            // Create a resizable MemoryStream and copy the buffer to it.
            using var stream = new MemoryStream();
            stream.Write(buffer, 0, buffer.Length);
            stream.Seek(0, SeekOrigin.Begin);

            // Add paragraphs, increasing the size of the document.
            AddParagraphs(stream);

            Assert.True(stream.Length > buffer.Length);
        }

        [Fact]
        public void Write_NonResizableMemoryStream_ThrowsNotSupportedException()
        {
            var buffer = new byte[3];
            using var stream = new MemoryStream(buffer);

            Assert.Throws<NotSupportedException>(() => stream.Write(new byte[] { 1, 2, 3, 4 }, 0, 4));
        }
    }
}
