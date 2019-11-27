//
// BookmarkTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using CodeSnippets.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class BookmarkTests
    {
        /// <summary>
        /// The w:name value of our bookmark.
        /// </summary>
        private const string BookmarkName = "_Bm001";

        /// <summary>
        /// The w:id value of our bookmark.
        /// </summary>
        private const int BookmarkId = 1;

        /// <summary>
        /// The test w:document with our bookmark, which encloses the two runs
        /// with inner texts "Second" and "Third".
        /// </summary>
        private static readonly XElement Document =
            new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                new XElement(W.body,
                    new XElement(W.p,
                        new XElement(W.r,
                            new XElement(W.t, "First"))),
                    new XElement(W.bookmarkStart,
                        new XAttribute(W.id, BookmarkId),
                        new XAttribute(W.name, BookmarkName)),
                    new XElement(W.p,
                        new XElement(W.r,
                            new XElement(W.t, "Second"))),
                    new XElement(W.p,
                        new XElement(W.r,
                            new XElement(W.t, "Third"))),
                    new XElement(W.bookmarkEnd,
                        new XAttribute(W.id, BookmarkId)),
                    new XElement(W.p,
                        new XElement(W.r,
                            new XElement(W.t, "Fourth")))
                )
            );

        /// <summary>
        /// Creates a <see cref="WordprocessingDocument"/> for on a <see cref="MemoryStream"/>
        /// testing purposes, using the given <paramref name="document"/> as the w:document
        /// root element of the main document part.
        /// </summary>
        /// <param name="document">The w:document root element.</param>
        /// <returns>The <see cref="MemoryStream"/> containing the <see cref="WordprocessingDocument"/>.</returns>
        private static MemoryStream CreateWordprocessingDocument(XElement document)
        {
            var stream = new MemoryStream();
            const WordprocessingDocumentType type = WordprocessingDocumentType.Document;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, type))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                part.PutXDocument(new XDocument(document));
            }

            return stream;
        }

        [Fact]
        public void GetRuns_WordprocessingDocumentWithBookmarks_CorrectRunsReturned()
        {
            // Arrange.
            // Create a new Word document on a Stream, using the test w:document
            // as the main document part.
            Stream stream = CreateWordprocessingDocument(Document);

            // Open the WordprocessingDocument on the Stream, using the Open XML SDK.
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true);

            // Get the w:document element from the main document part and find
            // our bookmark.
            XElement document = wordDocument.MainDocumentPart.GetXElement();
            Bookmark bookmark = Bookmark.Find(document, BookmarkName);

            // Act, getting the bookmarked runs.
            IEnumerable<XElement> runs = bookmark.GetRuns();

            // Assert.
            Assert.Equal(new[] {"Second", "Third"}, runs.Select(run => run.Value));
        }

        [Fact]
        public void GetText_WordprocessingDocumentWithBookmarks_CorrectRunsReturned()
        {
            // Arrange.
            // Create a new Word document on a Stream, using the test w:document
            // as the main document part.
            Stream stream = CreateWordprocessingDocument(Document);

            // Open the WordprocessingDocument on the Stream, using the Open XML SDK.
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true);

            // Get the w:document element from the main document part and find
            // our bookmark.
            XElement document = wordDocument.MainDocumentPart.GetXElement();
            Bookmark bookmark = Bookmark.Find(document, BookmarkName);

            // Act, getting the concatenated text contents of the bookmarked runs.
            string text = bookmark.GetValue();

            // Assert.
            Assert.Equal("SecondThird", text);
        }
    }
}
