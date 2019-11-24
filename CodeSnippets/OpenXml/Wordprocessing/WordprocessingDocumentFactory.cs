//
// DocumentFactory.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CodeSnippets.OpenXml.Wordprocessing
{
    /// <summary>
    /// Utility class that creates <see cref="WordprocessingDocument" /> instances.
    /// </summary>
    public static class WordprocessingDocumentFactory
    {
        /// <summary>
        /// Creates an empty <see cref="WordprocessingDocument" /> in the file system.
        /// </summary>
        /// <param name="path">The file system path of the document to be created.</param>
        /// <param name="document"></param>
        /// <returns>A new <see cref="WordprocessingDocument" />.</returns>
        public static WordprocessingDocument Create(
            string path,
            Document document = null)
        {
            WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                path, WordprocessingDocumentType.Document);

            MainDocumentPart part = wordDocument.AddMainDocumentPart();
            part.Document = document ?? new Document(new Body());

            return wordDocument;
        }

        /// <summary>
        /// Creates a <see cref="WordprocessingDocument" /> on the given <see cref="Stream" />.
        /// </summary>
        /// <param name="stream">The <see cref="Stream" />.</param>
        /// <param name="document">The <see cref="Document" /></param>
        /// <returns>A new <see cref="WordprocessingDocument" />.</returns>
        public static WordprocessingDocument Create(
            Stream stream,
            Document document = null)
        {
            WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                stream, WordprocessingDocumentType.Document);

            MainDocumentPart part = wordDocument.AddMainDocumentPart();
            part.Document = document ?? new Document(new Body());

            return wordDocument;
        }

        /// <summary>
        /// Creates a simple <see cref="WordprocessingDocument" /> for testing purposes.
        /// The document will have unformatted paragraphs with the given texts.
        /// </summary>
        /// <param name="path">The file system path of the document to be created.</param>
        /// <param name="texts">The collection of paragraph inner texts.</param>
        /// <returns>A new <see cref="WordprocessingDocument" />.</returns>
        public static WordprocessingDocument Create(string path, IEnumerable<string> texts)
        {
            return Create(
                path,
                new Document(
                    new Body(texts.Select(text =>
                        new Paragraph(
                            new Run(
                                new Text(text)))))));
        }

        /// <summary>
        /// Creates a simple <see cref="WordprocessingDocument" /> for testing purposes.
        /// The document will have unformatted paragraphs with the given texts.
        /// </summary>
        /// <param name="stream">The <see cref="Stream" /> on which to create the document.</param>
        /// <param name="texts">The collection of paragraph inner texts.</param>
        /// <returns>A new <see cref="WordprocessingDocument" />.</returns>
        public static WordprocessingDocument Create(Stream stream, IEnumerable<string> texts)
        {
            return Create(
                stream,
                new Document(
                    new Body(texts.Select(text =>
                        new Paragraph(
                            new Run(
                                new Text(text)))))));
        }
    }
}
