//
// DataBoundContentControlTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
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
    }
}
