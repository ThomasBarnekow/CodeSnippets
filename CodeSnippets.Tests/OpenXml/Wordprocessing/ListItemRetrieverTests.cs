//
// RetrieveListItemTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;
using Xunit.Abstractions;
using W = DocumentFormat.OpenXml.Linq.W;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class ListItemRetrieverTests
    {
        private readonly ITestOutputHelper _output;

        public ListItemRetrieverTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void RetrieveListItem_DocumentWithNumberedLists_ListItemSuccessfullyRetrieved()
        {
            const string path = "Resources\\Numbered Lists.docx";
            using WordprocessingDocument wordDoc = WordprocessingDocument.Open(path, false);

            XElement document = OpenXmlPartRootXElementExtensions.GetXElement(wordDoc.MainDocumentPart!)!;

            foreach (XElement paragraph in document.Descendants(W.p))
            {
                string listItem = ListItemRetriever.RetrieveListItem(wordDoc, paragraph);
                string text = paragraph.Descendants(W.t).Select(t => t.Value).StringConcatenate();

                _output.WriteLine(string.IsNullOrEmpty(listItem) ? text : $"{listItem} {text}");
            }
        }
    }
}
