//
// XmlTransformationTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class XmlTransformationTests
    {
        private const string Xml =
            @"<p id=""_fab91699-6d85-4ce5-b0b5-a17197520a7f"">" +
            @"This document is amongst a series of International Standards dealing with the conversion of systems of writing produced by Technical Committee ISO/TC 46, " +
            @"<em>Information and documentation</em>" +
            @", WG 3 " +
            @"<em>Conversion of written languages</em>" +
            @"." +
            @"</p>";

        private const string OuterXml =
            @"<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" +
            @"<w:r><w:t xml:space=""preserve"">This document is amongst a series of International Standards dealing with the conversion of systems of writing produced by Technical Committee ISO/TC 46, </w:t></w:r>" +
            @"<w:r><w:rPr><w:i /></w:rPr><w:t>Information and documentation</w:t></w:r>" +
            @"<w:r><w:t xml:space=""preserve"">, WG 3 </w:t></w:r>" +
            @"<w:r><w:rPr><w:i /></w:rPr><w:t>Conversion of written languages</w:t></w:r>" +
            @"<w:r><w:t>.</w:t></w:r>" +
            @"</w:p>";

        private static OpenXmlElement TransformElementToOpenXml(XElement element)
        {
            return element.Name.LocalName switch
            {
                "p" => new Paragraph(element.Nodes().Select(TransformNodeToOpenXml)),
                "em" => new Run(new RunProperties(new Italic()), CreateText(element.Value)),
                "b" => new Run(new RunProperties(new Bold()), CreateText(element.Value)),
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        private static OpenXmlElement TransformNodeToOpenXml(XNode node)
        {
            return node switch
            {
                XElement element => TransformElementToOpenXml(element),
                XText text => new Run(CreateText(text.Value)),
                _ => throw new ArgumentOutOfRangeException()
            };
        }

        private static Text CreateText(string text)
        {
            return new Text(text)
            {
                Space = text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1]))
                    ? new EnumValue<SpaceProcessingModeValues>(SpaceProcessingModeValues.Preserve)
                    : null
            };
        }

        private const string MoviesXml =
            @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Movies>
  <Movie>
    <Name>Crash</Name>
    <Released>2005</Released>
  </Movie>
  <Movie>
    <Name>The Departed</Name>
    <Released>2006</Released>
  </Movie>
  <Movie>
    <Name>The Pursuit of Happiness</Name>
    <Released>2006</Released>
  </Movie>
  <Movie>
    <Name>The Bucket List</Name>
    <Released>2007</Released>
  </Movie>
</Movies>";

        private static Table TransformMovies(XElement movies)
        {
            var headerRow = new[]
            {
                new TableRow(movies
                    .Elements()
                    .First()
                    .Elements()
                    .Select(e => new TableCell(new Paragraph(new Run(new Text(e.Name.LocalName))))))
            };
            IEnumerable<OpenXmlElement> movieRows = movies.Elements().Select(TransformMovie);

            return new Table(headerRow.Concat(movieRows));
        }

        private static OpenXmlElement TransformMovie(XElement element)
        {
            return element.Name.LocalName switch
            {
                "Movie" => new TableRow(element.Elements().Select(TransformMovie)),
                _ => new TableCell(new Paragraph(new Run(new Text(element.Value))))
            };
        }

        private static void InsertTable(WordprocessingDocument wordprocessingDocument)
        {
            const string xmlFile = @"Resources\Movies.xml";
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

            // How to read the XML and create a table in word document filled with data from XML

            // First, read the Movies XML string from your XML file.
            string moviesXml = File.ReadAllText(xmlFile);

            // Second, create the table as previously shown in the unit test method.
            XElement movies = XElement.Parse(moviesXml);
            Table table = TransformMovies(movies);

            // Third, append the Table to your empty Body.
            mainPart.Document.Body.AppendChild(table);
        }

        private static void AssertTableIsCorrect(XElement movies, Table table)
        {
            static (string Name, string Released) GetRowTexts(XElement movie)
            {
                string name = movie.Element("Name")?.Value ?? throw new ArgumentException();
                string released = movie.Element("Released")?.Value ?? throw new ArgumentException();
                return (name, released);
            }

            static bool HasRowTexts(TableRow tr, (string Name, string Released) rowTexts)
            {
                (string name, string released) = rowTexts;

                return tr.Elements<TableCell>().ElementAt(0).Descendants<Text>().Any(t => t.Text == name) &&
                       tr.Elements<TableCell>().ElementAt(1).Descendants<Text>().Any(t => t.Text == released);
            }

            Assert.All(
                movies.Elements().Select(GetRowTexts),
                rowTexts => Assert.Contains(table.Elements<TableRow>(), tr => HasRowTexts(tr, rowTexts)));
        }

        private static XElement GetMovies()
        {
            const string path = @"Resources\Movies.xml";
            string moviesXml = File.ReadAllText(path);
            return XElement.Parse(moviesXml);
        }

        [Fact]
        public void CanCreateTableFromXml()
        {
            XElement movies = XElement.Parse(MoviesXml);
            Table table = TransformMovies(movies);

            Assert.NotNull(table);
        }

        [Fact]
        public void CanInsertTable()
        {
            // Arrange, creating a new WordprocessingDocument.
            using var stream = new MemoryStream();
            using var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);

            MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());

            // Act, inserting the table.
            InsertTable(wordDocument);

            // Assert, demonstrating the table was added.
            XElement movies = GetMovies();
            Table table = wordDocument.MainDocumentPart.Document.Body.Elements<Table>().Single();
            AssertTableIsCorrect(movies, table);
        }

        [Fact]
        public void CanTransformXmlToOpenXml()
        {
            // Arrange, creating an XElement based on the given XML.
            XElement xmlParagraph = XElement.Parse(Xml);

            // Act, transforming the XML into Open XML.
            var paragraph = (Paragraph) TransformElementToOpenXml(xmlParagraph);

            // Assert, demonstrating that we have indeed created an Open XML Paragraph instance.
            Assert.Equal(OuterXml, paragraph.OuterXml);
        }
    }
}
