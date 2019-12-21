//
// OpenXmlReaderTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class OpenXmlReaderTests
    {
        private const string NamespaceUriW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly string XmlnsW = $"xmlns:w=\"{NamespaceUriW}\"";

        private static readonly string IndText =
            $@"<w:ind {XmlnsW} w:firstLine=""10"" w:left=""20"" w:right=""30""/>";

        private static readonly string DocumentText =
            $@"<w:document {XmlnsW}><w:body><w:p><w:pPr>{IndText}</w:pPr></w:p></w:body></w:document>";

        [Fact]
        public void ConvertingDocumentChangesIndProperties()
        {
            XElement element = XElement.Parse(DocumentText);

            var document = (Document) element.ToOpenXmlElement();
            Indentation ind = document.Descendants<Indentation>().First();

            Assert.Null(ind.Left);
            Assert.Null(ind.Right);

            Assert.Equal("10", ind.FirstLine);
            Assert.Equal("20", ind.Start);
            Assert.Equal("30", ind.End);
        }

        [Fact]
        public void ConvertingIndDoesNotChangeIndProperties()
        {
            XElement element = XElement.Parse(IndText);

            var ind = (OpenXmlUnknownElement) element.ToOpenXmlElement();

            Assert.Equal("10", ind.GetAttribute("firstLine", NamespaceUriW).Value);
            Assert.Equal("20", ind.GetAttribute("left", NamespaceUriW).Value);
            Assert.Equal("30", ind.GetAttribute("right", NamespaceUriW).Value);
        }
    }

    public static class XElementExtensions
    {
        public static OpenXmlElement ToOpenXmlElement(this XElement element)
        {
            // Write XElement to MemoryStream.
            using var stream = new MemoryStream();
            element.Save(stream);
            stream.Seek(0, SeekOrigin.Begin);

            // Read OpenXmlElement from MemoryStream.
            using OpenXmlReader reader = OpenXmlReader.Create(stream);
            reader.Read();
            return reader.LoadCurrentElement();
        }

        public static OpenXmlElement ToOpenXmlElement2(this XElement element)
        {
            using (var stream = new MemoryStream())
            {
                // Write XElement to MemoryStream.
                element.Save(stream);
                stream.Seek(0, SeekOrigin.Begin);

                // Read OpenXmlElement from MemoryStream.
                using OpenXmlReader reader = OpenXmlReader.Create(stream);
                {
                    reader.Read();
                    return reader.LoadCurrentElement();
                }
            }
        }
    }
}
