//
// SimpleFieldTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace StackOverflow.Tests.OpenXml.Wordprocessing
{
    public class SimpleFieldTests
    {
        /// <summary>
        /// Creates a Word document called "SimpleField.docx" that contains a
        /// w:fldSimple element with an incorrect field instruction "sdtContent".
        /// </summary>
        /// <remarks>
        /// When opening the document with Microsoft Word and updating the fields,
        /// Word inserts an error text "Error! Bookmark not defined."
        /// For security reasons, Word shows a dialog, asking the user whether
        /// fields should be updated. This dialog cannot be avoided.
        ///
        /// See https://stackoverflow.com/questions/58637790/openxml-table-of-contents-update-giving-error-bookmark-not-defined
        /// </remarks>
        [Fact]
        public void SimpleFieldWithIncorrectInstructionMakesWordInsertErrorText()
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                "SimpleField.docx", WordprocessingDocumentType.Document);

            MainDocumentPart part = wordDocument.AddMainDocumentPart();
            part.Document =
                new Document(
                    new Body(
                        new Paragraph(
                            new SimpleField
                            {
                                Instruction = "sdtContent",
                                Dirty = true
                            })));
        }

        /// <summary>
        /// Creates a Word document called "SimpleFieldRef.docx" that contains a
        /// w:fldSimple element with a reference to an existing bookmark.
        /// </summary>
        /// <remarks>
        /// When opening the document with Microsoft Word and updating the fields,
        /// Word does NOT insert any error text.
        /// For security reasons, Word shows a dialog, asking the user whether
        /// fields should be updated. This dialog cannot be avoided.
        ///
        /// See https://stackoverflow.com/questions/58637790/openxml-table-of-contents-update-giving-error-bookmark-not-defined
        /// </remarks>
        [Fact]
        public void SimpleFieldWithReferencingExistingBookmarkWorksFine()
        {
            using WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                "SimpleFieldRef.docx", WordprocessingDocumentType.Document);

            MainDocumentPart part = wordDocument.AddMainDocumentPart();
            part.Document =
                new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new Text("Hello World!"))),
                        new BookmarkStart { Id = "32767", Name = "_RefUpdate" },
                        new BookmarkEnd { Id = "32767" },
                        new Paragraph(
                            new SimpleField
                            {
                                Instruction = "REF _RefUpdate",
                                Dirty = true
                            }),
                        new SectionProperties()));
        }
    }
}
