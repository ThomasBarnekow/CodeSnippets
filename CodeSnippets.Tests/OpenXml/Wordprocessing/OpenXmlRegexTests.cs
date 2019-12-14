//
// OpenXmlRegexTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlPowerTools;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class OpenXmlRegexTests
    {
        [Theory]
        [InlineData("Hello Firstname ", new[] { "Firstname" })]
        [InlineData("Hello Firstname ", new[] { "F", "irstname" })]
        [InlineData("Hello Firstname ", new[] { "F", "i", "r", "s", "t", "n", "a", "m", "e" })]
        public void InnerText_ParagraphWithSymbols_SymbolIgnored(string expectedInnerText, IEnumerable<string> runTexts)
        {
            using MemoryStream stream = CreateWordprocessingDocument(runTexts);
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false);

            Document document = wordDocument.MainDocumentPart.Document;
            Paragraph paragraph = document.Descendants<Paragraph>().Single();
            string innerText = paragraph.InnerText;

            Assert.Equal(expectedInnerText, innerText);
        }

        [Theory]
        [InlineData("1 Run", "Firstname", new[] { "Firstname" }, "Albert")]
        [InlineData("2 Runs", "Firstname", new[] { "F", "irstname" }, "Bernie")]
        [InlineData("9 Runs", "Firstname", new[] { "F", "i", "r", "s", "t", "n", "a", "m", "e" }, "Charly")]
        public void Replace_PlaceholderInOneOrMoreRuns_SuccessfullyReplaced(
            string example,
            string propName,
            IEnumerable<string> runTexts,
            string replacement)
        {
            // Create a test WordprocessingDocument on a MemoryStream.
            using MemoryStream stream = CreateWordprocessingDocument(runTexts);

            // Save the Word document before replacing the placeholder.
            // You can use this to inspect the input Word document.
            File.WriteAllBytes($"{example} before Replacing.docx", stream.ToArray());

            // Replace the placeholder identified by propName with the replacement text.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
            {
                // Read the root element, a w:document in this case.
                // Note that GetXElement() is a shortcut for GetXDocument().Root.
                // This caches the root element and we can later write it back
                // to the main document part, using the PutXDocument() method.
                XElement document = wordDocument.MainDocumentPart.GetXElement();

                // Specify the parameters of the OpenXmlRegex.Replace() method,
                // noting that the replacement is given as a parameter.
                IEnumerable<XElement> content = document.Descendants(W.p);
                var regex = new Regex(propName);

                // Perform the replacement, thereby modifying the root element.
                OpenXmlRegex.Replace(content, regex, replacement, null);

                // Write the changed root element back to the main document part.
                wordDocument.MainDocumentPart.PutXDocument();
            }

            // Assert that we have done it right.
            AssertReplacementWasSuccessful(stream, replacement);

            // Save the Word document after having replaced the placeholder.
            // You can use this to inspect the output Word document.
            File.WriteAllBytes($"{example} after Replacing.docx", stream.ToArray());
        }

        private static MemoryStream CreateWordprocessingDocument(IEnumerable<string> runTexts)
        {
            var stream = new MemoryStream();
            const WordprocessingDocumentType type = WordprocessingDocumentType.Document;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, type))
            {
                MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
                mainDocumentPart.PutXDocument(new XDocument(CreateDocument(runTexts)));
            }

            return stream;
        }

        private static XElement CreateDocument(IEnumerable<string> runTexts)
        {
            // Produce a w:document with a single w:p that contains:
            // (1) one italic run with some lead-in, i.e., "Hello " in this example;
            // (2) one or more bold runs for the placeholder, which might or might not be split;
            // (3) one run with just a space; and
            // (4) one run with a symbol (i.e., a Wingdings smiley face).
            return new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
                new XElement(W.body,
                    new XElement(W.p,
                        new XElement(W.r,
                            new XElement(W.rPr,
                                new XElement(W.i)),
                            new XElement(W.t,
                                new XAttribute(XNamespace.Xml + "space", "preserve"),
                                "Hello ")),
                        runTexts.Select(rt =>
                            new XElement(W.r,
                                new XElement(W.rPr,
                                    new XElement(W.b)),
                                new XElement(W.t, rt))),
                        new XElement(W.r,
                            new XElement(W.t,
                                new XAttribute(XNamespace.Xml + "space", "preserve"),
                                " ")),
                        new XElement(W.r,
                            new XElement(W.sym,
                                new XAttribute(W.font, "Wingdings"),
                                new XAttribute(W._char, "F04A"))))));
        }

        private static void AssertReplacementWasSuccessful(MemoryStream stream, string replacement)
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false);

            XElement document = wordDocument.MainDocumentPart.GetXElement();
            XElement paragraph = document.Descendants(W.p).Single();
            List<XElement> runs = paragraph.Elements(W.r).ToList();

            // We have the expected number of runs, i.e., the lead-in, the first name,
            // a space character, and the symbol.
            Assert.Equal(4, runs.Count);

            // We still have the lead-in "Hello " and it is still formatted in italics.
            Assert.True(runs[0].Value == "Hello " && runs[0].Elements(W.rPr).Elements(W.i).Any());

            // We have successfully replaced our "Firstname" placeholder and the
            // concrete first name is formatted in bold, exactly like the placeholder.
            Assert.True(runs[1].Value == replacement && runs[1].Elements(W.rPr).Elements(W.b).Any());

            // We still have the space between the first name and the symbol and it
            // is unformatted.
            Assert.True(runs[2].Value == " " && !runs[2].Elements(W.rPr).Any());

            // Finally, we still have our smiley face symbol run.
            Assert.True(IsSymbolRun(runs[3], "Wingdings", "F04A"));
        }

        private static bool IsSymbolRun(XElement run, string fontValue, string charValue)
        {
            XElement sym = run.Elements(W.sym).FirstOrDefault();
            if (sym == null) return false;

            return (string) sym.Attribute(W.font) == fontValue &&
                   (string) sym.Attribute(W._char) == charValue;
        }

        [Theory]
        [InlineData("1 Run", "Firstname", new[] { "Firstname" }, "Albert")]
        [InlineData("2 Runs", "Firstname", new[] { "F", "irstname" }, "Bernie")]
        [InlineData("9 Runs", "Firstname", new[] { "F", "i", "r", "s", "t", "n", "a", "m", "e" }, "Charly")]
        public void TagReplacer_FormattedDocument_StylingAndSymbolsMissing(
            string example,
            string propName,
            IEnumerable<string> runTexts,
            string replacement)
        {
            // Arrange.
            using MemoryStream stream = CreateWordprocessingDocument(runTexts);

            // Act.
            TagReplacer.ReplaceTags(stream, propName, replacement);

            // Assert.
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false);
            XElement document = wordDocument.MainDocumentPart.GetXElement();
            XElement paragraph = document.Descendants(W.p).Single();
            List<XElement> runs = paragraph.Elements(W.r).ToList();

            // All run formatting is gone.
            Assert.DoesNotContain(runs, r => r.Elements(W.rPr).Any());

            // All symbols are gone.
            Assert.DoesNotContain(runs, r => r.Elements(W.sym).Any());

            // Save document for inspection.
            File.WriteAllBytes($"TagReplacer {example} after Replacing.docx", stream.ToArray());
        }

        [Fact]
        public void TagReplacer_OtherDocument_AsExpected()
        {
            // Arrange.
            using MemoryStream stream = CreateOtherWordprocessingDocument();

            // Act.
            TagReplacer.ReplaceTags(stream, "Firstname", "Donald");

            // Save document for inspection.
            File.WriteAllBytes($"TagReplacer Other after Replacing.docx", stream.ToArray());
        }

        private static MemoryStream CreateOtherWordprocessingDocument()
        {
            var stream = new MemoryStream();
            const WordprocessingDocumentType type = WordprocessingDocumentType.Document;

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, type))
            {
                MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
                mainDocumentPart.Document = CreateOtherDocument();
            }

            return stream;
        }

        private static Document CreateOtherDocument()
        {
            return new Document(
                new Body(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId { Val = "Heading1" }),
                        new BookmarkStart { Id = "1", Name = "_Ref123456" },
                        new Run(
                            new Text("About Firstname")),
                        new BookmarkEnd { Id = "1" })));
        }

        /// <summary>
        /// This class is based on @petelid's answer to
        /// https://stackoverflow.com/questions/28697701/openxml-tag-search,
        /// </summary>
        [SuppressMessage("ReSharper", "ConvertToUsingDeclaration")]
        [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
        private static class TagReplacer
        {
            public static void ReplaceTags(MemoryStream stream, string pattern, string replacement)
            {
                Regex regex = new Regex(pattern, RegexOptions.Compiled);

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, true))
                {
                    //grab the header parts and replace tags there
                    foreach (HeaderPart headerPart in wordDocument.MainDocumentPart.HeaderParts)
                    {
                        ReplaceParagraphParts(headerPart.Header, regex, replacement);
                    }
                    //now do the document
                    ReplaceParagraphParts(wordDocument.MainDocumentPart.Document, regex, replacement);
                    //now replace the footer parts
                    foreach (FooterPart footerPart in wordDocument.MainDocumentPart.FooterParts)
                    {
                        ReplaceParagraphParts(footerPart.Footer, regex, replacement);
                    }
                }
            }

            private static void ReplaceParagraphParts(OpenXmlElement element, Regex regex, string replacement)
            {
                foreach (var paragraph in element.Descendants<Paragraph>())
                {
                    Match match = regex.Match(paragraph.InnerText);
                    if (match.Success)
                    {
                        //create a new run and set its value to the correct text
                        //this must be done before the child runs are removed otherwise
                        //paragraph.InnerText will be empty
                        Run newRun = new Run();
                        newRun.AppendChild(new Text(paragraph.InnerText.Replace(match.Value, replacement)));
                        //remove any child runs
                        paragraph.RemoveAllChildren<Run>();
                        //add the newly created run
                        paragraph.AppendChild(newRun);
                    }
                }
            }
        }
    }
}
