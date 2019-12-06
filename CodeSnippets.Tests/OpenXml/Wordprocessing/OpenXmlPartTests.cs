//
// OpenXmlPartTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class OpenXmlPartTests
    {
        [Fact]
        public void AddNewPart_NewMainDocumentPart_SuccessfullyAdded()
        {
            const string path = "Document.docx";
            const WordprocessingDocumentType type = WordprocessingDocumentType.Document;

            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(path, type);

            // Create minimum main document part.
            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            mainDocumentPart.Document = new Document(new Body(new Paragraph()));

            // Create empty style definitions part.
            var styleDefinitionsPart = mainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            styleDefinitionsPart.Styles = new Styles();

            // Create empty numbering definitions part.
            var numberingDefinitionsPart = mainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
            numberingDefinitionsPart.Numbering = new Numbering();
        }
    }
}
