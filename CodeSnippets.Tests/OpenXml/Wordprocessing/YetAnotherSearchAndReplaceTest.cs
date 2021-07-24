using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class YetAnotherSearchAndReplaceTest
    {
        [Fact]
        public void CanSearchAndReplaceStringInOpenXmlPartAlthoughThisIsNotTheWayToSearchAndReplaceText()
        {
            // Arrange.
            using var docxStream = new MemoryStream();
            using (var wordDocument = WordprocessingDocument.Create(docxStream, WordprocessingDocumentType.Document))
            {
                MainDocumentPart part = wordDocument.AddMainDocumentPart();
                var p1 = new Paragraph(
                    new Run(
                        new Text("Hello world!")));

                var p2 = new Paragraph(
                    new Run(
                        new Text("Hello ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(
                        new Text("world!")));

                part.Document = new Document(new Body(p1, p2));

                Assert.Equal("Hello world!", p1.InnerText);
                Assert.Equal("Hello world!", p2.InnerText);
            }

            // Act.
            SearchAndReplace(docxStream);

            // Assert.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(docxStream, false))
            {
                MainDocumentPart part = wordDocument.MainDocumentPart;
                Paragraph p1 = part.Document.Descendants<Paragraph>().First();
                Paragraph p2 = part.Document.Descendants<Paragraph>().Last();

                Assert.Equal("Hi Everyone!", p1.InnerText);
                Assert.Equal("Hello world!", p2.InnerText);
            }
        }

        private static void SearchAndReplace(MemoryStream docxStream)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(docxStream, true))
            {
                // If you wanted to read the part's contents as text, this is how you
                // would do it.
                string partText = ReadPartText(wordDocument.MainDocumentPart);
                wordDocument.MainDocumentPart.GetXDocument();

                // Note that this is not the way in which you should search and replace
                // text in Open XML documents. The text might be split across multiple
                // w:r elements, so you would not find the text in that case.
                var regex = new Regex("Hello world!");
                partText = regex.Replace(partText, "Hi Everyone!");

                // If you wanted to write changed text back to the part, this is how
                // you would do it.
                WritePartText(wordDocument.MainDocumentPart, partText);
            }

            docxStream.Seek(0, SeekOrigin.Begin);
        }

        private static string ReadPartText(OpenXmlPart part)
        {
            using Stream partStream = part.GetStream(FileMode.OpenOrCreate, FileAccess.Read);
            using var sr = new StreamReader(partStream);
            return sr.ReadToEnd();
        }

        private static void WritePartText(OpenXmlPart part, string text)
        {
            using Stream partStream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var sw = new StreamWriter(partStream);
            sw.Write(text);
        }
    }
}
