using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class SmartTagTests
    {
        private const string Xml =
            @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
    <w:body>
        <w:p>
            <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""PersonName"">
                <w:r w:rsidRPr=""00BF444F"">
                    <w:rPr>
                        <w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/>
                        <w:b/>
                        <w:bCs/>
                        <w:sz w:val=""40""/>
                        <w:szCs w:val=""40""/>
                    </w:rPr>
                    <w:t>ST</w:t>
                </w:r>
            </w:smartTag>
            <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""PersonName"">
                <w:r w:rsidRPr=""00BF444F"">
                    <w:rPr>
                        <w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/>
                        <w:b/>
                        <w:bCs/>
                        <w:sz w:val=""40""/>
                        <w:szCs w:val=""40""/>
                    </w:rPr>
                    <w:t>AR</w:t>
                </w:r>
            </w:smartTag>
            <w:r w:rsidRPr=""00BF444F"">
                <w:rPr>
                    <w:rFonts w:ascii=""Arial"" w:hAnsi=""Arial"" w:cs=""Arial""/>
                    <w:b/>
                    <w:bCs/>
                    <w:sz w:val=""40""/>
                    <w:szCs w:val=""40""/>
                </w:rPr>
                <w:t xml:space=""preserve"">T</w:t>
            </w:r>
        </w:p>
    </w:body>
</w:document>";

        [Fact]
        public void CanStripSmartTags()
        {
            // Say you have a WordprocessingDocument stored on a stream (e.g., read from a file).
            using Stream stream = CreateTestWordprocessingDocument();

            // Open the WordprocessingDocument and inspect it using the strongly typed classes.
            // This shows that we find OpenXmlUnknownElement instances are found and only a
            // single Run instance is recognized.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false))
            {
                // Now, get the w:document as a strongly typed Document instance and demonstrate
                // that the document contains three Run instances.
                MainDocumentPart part = wordDocument.MainDocumentPart;
                Document document = part.Document;

                Assert.Single(document.Descendants<Run>());
                Assert.NotEmpty(document.Descendants<OpenXmlUnknownElement>());
            }

            // Now, open that WordprocessingDocument to make edits, using Linq to XML.
            // Do NOT use the strongly typed classes in this context.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
            {
                // Get the w:document as an XElement and demonstrate that this w:document contains
                // w:smartTag elements.
                MainDocumentPart part = wordDocument.MainDocumentPart;
                string xml = ReadString(part);
                XElement document = XElement.Parse(xml);

                Assert.NotEmpty(document.Descendants().Where(d => d.Name.LocalName == "smartTag"));

                // Transform the w:document, stripping all w:smartTag elements and demonstrate
                // that the transformed w:document no longer contains w:smartTag elements.
                var transformedDocument = (XElement) StripSmartTags(document);

                Assert.Empty(transformedDocument.Descendants().Where(d => d.Name.LocalName == "smartTag"));

                // Write the transformed document back to the part.
                WriteString(part, transformedDocument.ToString(SaveOptions.DisableFormatting));
            }

            // Open the WordprocessingDocument again and inspect it using the strongly typed classes.
            // This demonstrates that all Run instances are now recognized.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false))
            {
                // Now, get the w:document as a strongly typed Document instance and demonstrate
                // that the document contains three Run instances.
                MainDocumentPart part = wordDocument.MainDocumentPart;
                Document document = part.Document;

                Assert.Equal(3, document.Descendants<Run>().Count());
                Assert.Empty(document.Descendants<OpenXmlUnknownElement>());
            }
        }

        /// <summary>
        /// Recursive, pure functional transform that removes all w:smartTag elements.
        /// </summary>
        /// <param name="node">The <see cref="XNode" /> to be transformed.</param>
        /// <returns>The transformed <see cref="XNode" />.</returns>
        private static object StripSmartTags(XNode node)
        {
            if (!(node is XElement element))
            {
                return node;
            }

            if (element.Name.LocalName == "smartTag")
            {
                return element.Elements();
            }

            return new XElement(element.Name, element.Attributes(),
                element.Nodes().Select(StripSmartTags));
        }

        private static Stream CreateTestWordprocessingDocument()
        {
            var stream = new MemoryStream();

            using var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
            MainDocumentPart part = wordDocument.AddMainDocumentPart();
            WriteString(part, Xml);

            return stream;
        }

        #region Generic Open XML Utilities

        private static string ReadString(OpenXmlPart part)
        {
            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var streamReader = new StreamReader(stream);
            return streamReader.ReadToEnd();
        }

        private static void WriteString(OpenXmlPart part, string text)
        {
            using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var streamWriter = new StreamWriter(stream);
            streamWriter.Write(text);
        }

        #endregion
    }
}
