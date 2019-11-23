//
// CommentTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace StackOverflow.Tests.OpenXml.Wordprocessing
{
    /// <summary>
    /// Demonstrates how to work with Word comments, using the Open XML SDK.
    /// </summary>
    public class CommentTests
    {
        /// <summary>
        /// Demonstrates how to create a Word comment, adding a WordprocessingCommentsPart
        /// before adding the comment.
        /// </summary>
        [Fact]
        public void CanCreateComment()
        {
            // Let's say we have a document called Comments.docx with a single
            // paragraph saying "Hello World!".
            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                "Comments.docx", WordprocessingDocumentType.Document);

            MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
            mainDocumentPart.Document =
                new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text("Hello World!")))));

            // We can add a WordprocessingCommentsPart to the MainDocumentPart
            // and make sure it has a w:comments root element.
            var commentsPart = mainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();

            // Now, say we want to comment on the whole paragraph. As a first
            // step, we create a comment and add it to the w:comments element.
            var comment = new Comment(new Paragraph(new Run(new Text("This is my comment."))))
            {
                Id = "1",
                Author = "Thomas Barnekow"
            };

            commentsPart.Comments.AppendChild(comment);

            // Then, we need to add w:commentRangeStart and w:commentRangeEnd
            // elements to the text on which we want to comment. In this example,
            // we are getting "some" w:p element and some first and last w:r
            // elements and insert the w:commentRangeStart before the first and
            // the w:commentRangeEnd after the last w:r.
            Paragraph p = mainDocumentPart.Document.Descendants<Paragraph>().First();
            Run firstRun = p.Elements<Run>().First();
            Run lastRun = p.Elements<Run>().Last();

            firstRun.InsertBeforeSelf(new CommentRangeStart { Id = "1" });
            CommentRangeEnd commentRangeEnd = lastRun.InsertAfterSelf(new CommentRangeEnd { Id = "1" });
            commentRangeEnd.InsertAfterSelf(new Run(new CommentReference { Id = "1" }));
        }
    }
}
