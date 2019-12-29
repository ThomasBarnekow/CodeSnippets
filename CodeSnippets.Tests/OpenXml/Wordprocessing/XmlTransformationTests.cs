//
// XmlTransformationTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
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

        [Fact]
        public void CanTransformXmlToOpenXml()
        {
            // Arrange, creating an XElement based on the given XML.
            var xmlParagraph = XElement.Parse(Xml);

            // Act, transforming the XML into Open XML.
            var paragraph = (Paragraph) TransformElementToOpenXml(xmlParagraph);

            // Assert, demonstrating that we have indeed created an Open XML Paragraph instance.
            Assert.Equal(OuterXml, paragraph.OuterXml);
        }

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
    }
}
