//
// ContentControlWriter.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CodeSnippets.OpenXml.Wordprocessing
{
    /// <summary>
    /// Sample class that transforms block-level structured document tags (SDTs)
    /// by replacing their contents with texts specified in a content map. The
    /// latter maps the SDTs' w:tag values to the texts to be used for such
    /// SDTs.
    /// </summary>
    public class ContentControlWriter
    {
        private readonly IDictionary<string, string> _contentMap;

        /// <summary>
        /// Initializes a new ContentControlWriter instance.
        /// </summary>
        /// <param name="contentMap">The mapping of content control tags to content control texts.
        /// </param>
        public ContentControlWriter(IDictionary<string, string> contentMap)
        {
            _contentMap = contentMap;
        }

        /// <summary>
        /// Transforms the given WordprocessingDocument by setting the content
        /// of relevant block-level content controls.
        /// </summary>
        /// <param name="wordDocument">The WordprocessingDocument to be transformed.</param>
        public void WriteContentControls(WordprocessingDocument wordDocument)
        {
            MainDocumentPart part = wordDocument.MainDocumentPart;
            part.Document = (Document) TransformDocument(part.Document);
        }

        private object TransformDocument(OpenXmlElement element)
        {
            if (element is SdtBlock sdt)
            {
                string tagValue = GetTagValue(sdt);
                if (_contentMap.TryGetValue(tagValue, out string text))
                {
                    return TransformSdtBlock(sdt, text);
                }
            }

            return Transform(element, TransformDocument);
        }

        private static object TransformSdtBlock(OpenXmlElement element, string text)
        {
            return element is SdtContentBlock
                ? new SdtContentBlock(new Paragraph(new Run(new Text(text))))
                : Transform(element, e => TransformSdtBlock(e, text));
        }

        private static string GetTagValue(SdtElement sdt) => sdt
            .Descendants<Tag>()
            .Select(tag => tag.Val.Value)
            .FirstOrDefault();

        private static T Transform<T>(T element, Func<OpenXmlElement, object> transformation)
            where T : OpenXmlElement
        {
            var transformedElement = (T) element.CloneNode(false);
            transformedElement.Append(element.Elements().Select(e => (OpenXmlElement) transformation(e)));
            return transformedElement;
        }
    }
}
