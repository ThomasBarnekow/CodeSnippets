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
    public partial class XmlTransformationTests
    {
        /// <summary>
        /// Reads
        /// </summary>
        /// <returns></returns>
        private static XElement GetMovies()
        {
            const string path = @"Resources\Movies.xml";
            string moviesXml = File.ReadAllText(path);
            return XElement.Parse(moviesXml);
        }

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

        /// <summary>
        /// This is the non-working example from Enigma State.
        /// </summary>
        private static Table TransformMoviesTheWrongWay(XElement movies)
        {
            Table table = new Table();

            var headerRow = new[]
            {
                new TableRow(movies
                    .Elements()
                    .First()
                    .Elements()
                    .Select(e => new TableCell(new Paragraph(new Run(new Text(e.Name.LocalName))))))
            };
            var movieRows = movies.Elements().Select(TransformMovie);
            headerRow.Concat(movieRows);
            TableProperties tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder()
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 24
                    },
                    new BottomBorder()
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 24
                    },
                    new LeftBorder()
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 24
                    },
                    new RightBorder()
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 24
                    },
                    new InsideHorizontalBorder()
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 24
                    },
                    new InsideVerticalBorder()
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 24
                    }
                )
            );
            table.Append(headerRow);
            table.AppendChild(tblProp);
            //headerRow.Concat(movieRows);

            return table;
        }

        private static Table TransformMoviesWithBorders(XElement movies)
        {
            var table = new Table();

            var tblPr = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 24 },
                    new BottomBorder { Val = BorderValues.Single, Size = 24 },
                    new LeftBorder { Val = BorderValues.Single, Size = 24 },
                    new RightBorder { Val = BorderValues.Single, Size = 24 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 24 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 24 }));

            var headerRow = new TableRow(movies
                .Elements()
                .First()
                .Elements()
                .Select(e => new TableCell(new Paragraph(new Run(new Text(e.Name.LocalName))))));

            IEnumerable<OpenXmlElement> movieRows = movies.Elements().Select(TransformMovie);

            // Append child elements in the right order.
            table.AppendChild(tblPr);
            table.AppendChild(headerRow);
            table.Append(movieRows);

            return table;
        }

        private static OpenXmlElement TransformMovie(XElement element)
        {
            return element.Name.LocalName switch
            {
                "Movie" => new TableRow(element.Elements().Select(TransformMovie)),
                _ => new TableCell(new Paragraph(new Run(new Text(element.Value))))
            };
        }

        /// <summary>
        /// This answers the additional question of how to insert a <see cref="Table"/>
        /// into a <see cref="WordprocessingDocument"/>.
        /// </summary>
        /// <param name="wordprocessingDocument">The <see cref="WordprocessingDocument"/>.</param>
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

        [Fact]
        public void CanCreateTableFromXml()
        {
            XElement movies = GetMovies();
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
        public void CanInsertTableWithBorders()
        {
            using var stream = new MemoryStream();

            using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Act, inserting the table.
                XElement movies = GetMovies();
                Table table = TransformMoviesWithBorders(movies);
                mainPart.Document.Body.AppendChild(table);
            }

            File.WriteAllBytes("TableWithBorders.docx", stream.ToArray());
        }

        [Fact]
        public void CannotCreateTableTheWrongWay()
        {
            XElement movies = GetMovies();
            Table table = TransformMoviesTheWrongWay(movies);

            Assert.NotNull(table);
        }
    }
}
