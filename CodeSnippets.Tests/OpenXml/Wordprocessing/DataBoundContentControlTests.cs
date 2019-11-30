//
// DataBoundContentControlTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class DataBoundContentControlTests
    {
        private const WordprocessingDocumentType Type = WordprocessingDocumentType.Document;

        private const string NsPrefix = "ex";
        private const string NsName = "http://example.com";
        private static readonly XNamespace Ns = NsName;

        private static string CreateCustomXmlPart(MainDocumentPart mainDocumentPart, XElement rootElement)
        {
            CustomXmlPart customXmlPart = mainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPart.PutXDocument(new XDocument(rootElement));
            return CreateCustomXmlPropertiesPart(customXmlPart);
        }

        private static string CreateCustomXmlPropertiesPart(CustomXmlPart customXmlPart)
        {
            XElement rootElement = customXmlPart.GetXElement();
            if (rootElement == null) throw new InvalidOperationException();

            string storeItemId = "{" + Guid.NewGuid().ToString().ToUpper() + "}";

            // Create a ds:dataStoreItem associated with the custom XML part's root element.
            var dataStoreItem = new DataStoreItem
            {
                ItemId = storeItemId,
                SchemaReferences = new SchemaReferences()
            };

            if (rootElement.Name.Namespace != XNamespace.None)
            {
                dataStoreItem.SchemaReferences.AppendChild(new SchemaReference {Uri = rootElement.Name.NamespaceName});
            }

            // Create the custom XML properties part.
            var propertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            propertiesPart.DataStoreItem = dataStoreItem;
            propertiesPart.DataStoreItem.Save();

            return storeItemId;
        }

        [Fact]
        public void CanDataBindBlockLevelSdtToCustomXmlWithNsPrefixIfNsPrefixInPrefixMapping()
        {
            // The following root element has an explicitly created attribute
            // xmlns:ex="http://example.com":
            //
            //   <ex:Root xmlns:ex="http://example.com">
            //     <ex:Node>VALUE1</ex:Node>
            //   </ex:Root>
            //
            var customXmlRootElement =
                new XElement(Ns + "Root",
                    new XAttribute(XNamespace.Xmlns + NsPrefix, NsName),
                    new XElement(Ns + "Node", "VALUE1"));

            using WordprocessingDocument wordDocument =
                WordprocessingDocument.Create("SdtBlock_NsPrefix_WithNsPrefixInMapping.docx", Type);

            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            string storeItemId = CreateCustomXmlPart(mainDocumentPart, customXmlRootElement);

            mainDocumentPart.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                    new XElement(W.body,
                        new XElement(W.sdt,
                            new XElement(W.sdtPr,
                                new XElement(W.dataBinding,
                                    // Note the w:prefixMapping attribute WITH a namespace
                                    // prefix and the corresponding w:xpath atttibute.
                                    new XAttribute(W.prefixMappings, $"xmlns:{NsPrefix}='{NsName}'"),
                                    new XAttribute(W.xpath, $"{NsPrefix}:Root[1]/{NsPrefix}:Node[1]"),
                                    new XAttribute(W.storeItemID, storeItemId))),
                            new XElement(W.sdtContent,
                                new XElement(W.p)))))));

            // Note that we just added an empty w:p to the w:sdtContent element.
            // However, if you open the Word document created by the above code
            // in Microsoft Word, you should see a single paragraph saying
            // "VALUE1".
        }

        [Fact]
        public void CanDataBindBlockLevelSdtToCustomXmlWithoutNsPrefixIfNsPrefixInPrefixMapping()
        {
            // The following root element has an implicitly created attribute
            // xmlns='http://example.com':
            //
            //   <Root xmlns="http://example.com">
            //     <Node>VALUE1</Node>
            //   </Root>
            //
            var customXmlRootElement =
                new XElement(Ns + "Root",
                    new XElement(Ns + "Node", "VALUE1"));

            using WordprocessingDocument wordDocument =
                WordprocessingDocument.Create("SdtBlock_DefaultNs_WithNsPrefixInMapping.docx", Type);

            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            string storeItemId = CreateCustomXmlPart(mainDocumentPart, customXmlRootElement);

            mainDocumentPart.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                    new XElement(W.body,
                        new XElement(W.sdt,
                            new XElement(W.sdtPr,
                                new XElement(W.dataBinding,
                                    // Note the w:prefixMapping attribute WITH a namespace
                                    // prefix and the corresponding w:xpath atttibute.
                                    new XAttribute(W.prefixMappings, $"xmlns:{NsPrefix}='{NsName}'"),
                                    new XAttribute(W.xpath, $"{NsPrefix}:Root[1]/{NsPrefix}:Node[1]"),
                                    new XAttribute(W.storeItemID, storeItemId))),
                            new XElement(W.sdtContent,
                                new XElement(W.p)))))));

            // Note that we just added an empty w:p to the w:sdtContent element.
            // However, if you open the Word document created by the above code
            // in Microsoft Word, you should see a single paragraph saying
            // "VALUE1".
        }

        [Fact]
        public void CannotDataBindBlockLevelSdtToCustomXmlWithDefaultNsIfNotNsPrefixInPrefixMapping()
        {
            // The following root element has an implicitly created attribute
            // xmlns='http://example.com':
            //
            //   <Root xmlns="http://example.com">
            //     <Node>VALUE1</Node>
            //   </Root>
            //
            var customXmlRootElement =
                new XElement(Ns + "Root",
                    new XElement(Ns + "Node", "VALUE1"));

            using WordprocessingDocument wordDocument =
                WordprocessingDocument.Create("SdtBlock_DefaultNs_WithoutNsPrefixInMapping.docx", Type);

            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            string storeItemId = CreateCustomXmlPart(mainDocumentPart, customXmlRootElement);

            mainDocumentPart.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                    new XElement(W.body,
                        new XElement(W.sdt,
                            new XElement(W.sdtPr,
                                new XElement(W.dataBinding,
                                    // Note the w:prefixMapping attribute WITHOUT a namespace
                                    // prefix and the corresponding w:xpath atttibute.
                                    new XAttribute(W.prefixMappings, $"xmlns='{NsName}'"),
                                    new XAttribute(W.xpath, "Root[1]/Node[1]"),
                                    new XAttribute(W.storeItemID, storeItemId))),
                            new XElement(W.sdtContent,
                                new XElement(W.p)))))));

            // Note that we just added an empty w:p to the w:sdtContent element.
            // If you open the Word document created by the above code in Microsoft
            // Microsoft Word, you will only see an EMPTY paragraph.
        }

        [Fact]
        public void CanUpdateCustomXmlAndMainDocumentPart()
        {
            // Define the initial and updated values of our custom XML element and
            // the data-bound w:sdt element.
            const string initialValue = "VALUE1";
            const string updatedValue = "value2";

            // Create the root element of the custom XML part with the initial value.
            var customXmlRoot =
                new XElement(Ns + "Root",
                    new XAttribute(XNamespace.Xmlns + NsPrefix, NsName),
                    new XElement(Ns + "Node", initialValue));

            // Create the w:sdtContent child element of our w:sdt with the initial value.
            var sdtContent =
                new XElement(W.sdtContent,
                    new XElement(W.p,
                        new XElement(W.r,
                            new XElement(W.t, initialValue))));

            // Create a WordprocessingDocument with the initial values.
            using var stream = new MemoryStream();
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, Type))
            {
                InitializeWordprocessingDocument(wordDocument, customXmlRoot, sdtContent);
            }

            // Assert the WordprocessingDocument has the expected, initial values.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
            {
                AssertValuesAreAsExpected(wordDocument, initialValue);
            }

            // Update the WordprocessingDocument, using the updated value.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
            {
                MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;

                // Change custom XML element, again using the simplifying assumption
                // that we only have a single custom XML part and a single ex:Node
                // element.
                CustomXmlPart customXmlPart = mainDocumentPart.CustomXmlParts.Single();
                XElement root = customXmlPart.GetXElement();
                XElement node = root.Elements(Ns + "Node").Single();
                node.Value = updatedValue;
                customXmlPart.PutXDocument();

                // Change the w:sdt contained in the MainDocumentPart.
                XElement document = mainDocumentPart.GetXElement();
                XElement sdt = FindSdtWithTag("Node", document);
                sdtContent = sdt.Elements(W.sdtContent).Single();
                sdtContent.ReplaceAll(
                    new XElement(W.p,
                        new XElement(W.r,
                            new XElement(W.t, updatedValue))));

                mainDocumentPart.PutXDocument();
            }

            // Assert the WordprocessingDocument has the expected, updated values.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
            {
                AssertValuesAreAsExpected(wordDocument, updatedValue);
            }
        }

        private static void InitializeWordprocessingDocument(
            WordprocessingDocument wordDocument,
            XElement customXmlRoot,
            XElement sdtContent)
        {
            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            string storeItemId = CreateCustomXmlPart(mainDocumentPart, customXmlRoot);

            mainDocumentPart.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                    new XElement(W.body,
                        new XElement(W.sdt,
                            new XElement(W.sdtPr,
                                new XElement(W.tag, new XAttribute(W.val, "Node")),
                                new XElement(W.dataBinding,
                                    new XAttribute(W.prefixMappings, $"xmlns:{NsPrefix}='{NsName}'"),
                                    new XAttribute(W.xpath, $"{NsPrefix}:Root[1]/{NsPrefix}:Node[1]"),
                                    new XAttribute(W.storeItemID, storeItemId))),
                            sdtContent)))));
        }

        private static void AssertValuesAreAsExpected(
            WordprocessingDocument wordDocument,
            string expectedValue)
        {
            // Retrieve inner text of w:sdt element.
            MainDocumentPart mainDocumentPart = wordDocument.MainDocumentPart;
            XElement sdt = FindSdtWithTag("Node", mainDocumentPart.GetXElement());
            string sdtInnerText = GetInnerText(sdt);

            // Retrieve inner text of custom XML element, making the simplifying
            // assumption that we only have a single custom XML part. In reality,
            // we would have to find the custom XML part to which our w:sdt elements
            // are bound among any number of custom XML parts. Further, in our
            // simplified example, we also assume there is a single ex:Node element.
            CustomXmlPart customXmlPart = mainDocumentPart.CustomXmlParts.Single();
            XElement root = customXmlPart.GetXElement();
            XElement node = root.Elements(Ns + "Node").Single();
            string nodeInnerText = node.Value;

            // Assert those inner text are indeed equal.
            Assert.Equal(expectedValue, sdtInnerText);
            Assert.Equal(expectedValue, nodeInnerText);
        }

        private static XElement FindSdtWithTag(string tagValue, XElement openXmlCompositeElement)
        {
            return openXmlCompositeElement
                .Descendants(W.sdt)
                .FirstOrDefault(e => e
                    .Elements(W.sdtPr)
                    .Elements(W.tag)
                    .Any(tag => (string) tag.Attribute(W.val) == tagValue));
        }

        private static string GetInnerText(XElement openXmlElement)
        {
            return openXmlElement
                .DescendantsAndSelf(W.r)
                .Select(UnicodeMapper.RunToString)
                .StringConcatenate();
        }
    }
}
