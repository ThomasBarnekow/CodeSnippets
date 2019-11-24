//
// OrdinalNumberFormattingTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class OrdinalNumberFormattingTests
    {
        private static readonly Regex OrdinalNumberSuffixRegex = new Regex("(?<=[0-9]+)(st|nd|rd|th)");

        [Theory]
        [InlineData("Take the 1st, 2nd, or 3rd element.", 3)]
        [InlineData("1st or 2nd", 2)]
        [InlineData("Sorry, this text does not contain any ordinal number.", 0)]
        public void FormatSuperscript_MultipleOccurrences_CorrectlyFormatted(string innerText, int count)
        {
            Paragraph paragraph = FormatSuperscript(innerText);
            Assert.Equal(count, paragraph.Descendants<VerticalTextAlignment>().Count());
        }

        /// <summary>
        /// Creates a new <see cref="Paragraph" /> with ordinal number suffixes
        /// (i.e., "st", "nd", "rd", and "4th") formatted as a superscript.
        /// </summary>
        /// <param name="innerText">The paragraph's inner text.</param>
        /// <returns>A new, formatted <see cref="Paragraph" />.</returns>
        public static Paragraph FormatSuperscript(string innerText)
        {
            var destParagraph = new Paragraph();
            var startIndex = 0;

            foreach (Match match in OrdinalNumberSuffixRegex.Matches(innerText))
            {
                if (match.Index > startIndex)
                {
                    string text = innerText[startIndex..match.Index];
                    destParagraph.AppendChild(new Run(CreateText(text)));
                }

                destParagraph.AppendChild(
                    new Run(
                        new RunProperties(
                            new VerticalTextAlignment
                            {
                                Val = VerticalPositionValues.Superscript
                            }),
                        CreateText(match.Value)));

                startIndex = match.Index + match.Length;
            }

            if (startIndex < innerText.Length)
            {
                string text = innerText.Substring(startIndex);
                destParagraph.AppendChild(new Run(CreateText(text)));
            }

            return destParagraph;
        }

        /// <summary>
        /// Creates a new <see cref="Text" /> instance with the correct xml:space
        /// attribute value.
        /// </summary>
        /// <param name="text">The text.</param>
        /// <returns>A new <see cref="Text" /> instance.</returns>
        public static Text CreateText(string text)
        {
            if (string.IsNullOrEmpty(text)) return new Text();

            if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[^1]))
                return new Text(text) {Space = SpaceProcessingModeValues.Preserve};

            return new Text(text);
        }

        [Fact]
        public void FormatSuperscript_DateOfBirth_CorrectlyFormatted()
        {
            // Say we have a body or other container with a number of paragraphs, one
            // of which is the paragraph that we want to format. In our case, we want
            // the paragraph the inner text of which starts with "Date of birth:"
            var body =
                new Body(
                    new Paragraph(new Run(new Text("Full name: Phung Anh Tu"))),
                    new Paragraph(new Run(new Text("Date of birth: October 15th, 2019"))),
                    new Paragraph(new Run(new Text("Gender: male"))));

            Paragraph sourceParagraph = body
                .Descendants<Paragraph>()
                .First(p => p.InnerText.StartsWith("Date of birth:"));

            // In a first step, we'll create a new, formatted paragraph.
            Paragraph destParagraph = FormatSuperscript(sourceParagraph.InnerText);

            // Next, we format the existing paragraph by replacing it with the new,
            // formatted one.
            body.ReplaceChild(destParagraph, sourceParagraph);

            // Finally, let's verify that we have a single "th" run that is:
            // - preceded by one run with inner text "Date of birth: October 15",
            // - followed by one run with inner text ", 2019", and
            // - formatted as a superscript.
            Assert.Single(body
                .Descendants<Run>()
                .Where(r => r.InnerText == "th" &&
                            r.PreviousSibling().InnerText == "Date of birth: October 15" &&
                            r.NextSibling().InnerText == ", 2019" &&
                            r.RunProperties.VerticalTextAlignment.Val == VerticalPositionValues.Superscript));
        }
    }
}
