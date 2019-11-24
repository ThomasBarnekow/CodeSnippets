//
// OpenXmlPowerToolsWrapper.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;

namespace CodeSnippets.OpenXmlWrapper
{
    public class OpenXmlPowerToolsWrapper
    {
        public static string GetMainDocumentPart(string path)
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);
            return wordDocument.MainDocumentPart.GetXElement().ToString();
        }

        public static string FinishReview(string path)
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(path, true);

            var settings = new SimplifyMarkupSettings
            {
                AcceptRevisions = true,
                RemoveComments = true
            };

            MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
            return wordDocument.MainDocumentPart.GetXElement().ToString();
        }
    }
}
