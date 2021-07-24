using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class RepeatingSectionsTests
    {
        public void CanInsertRepeatingSections()
        {
            using WordprocessingDocument doc = WordprocessingDocument.Open(@"C:\in\test.docx", true);

            // Get the w:sdtContent element of the first block-level w:sdt element,
            // noting that "sdtContent" is called "mainDoc" in the question.
            SdtContentBlock sdtContent = doc.MainDocumentPart.Document.Body
                .Elements<SdtBlock>()
                .Select(sdt => sdt.SdtContentBlock)
                .First();

            // Get last element within SdtContentBlock. This seems to represent a "person".
            SdtBlock person = sdtContent.Elements<SdtBlock>().Last();

            // Create a clone and remove an existing w:id element from the clone's w:sdtPr
            // element, to ensure we don't repeat it. Note that the w:id element is optional
            // and Word will add one when it saves the document.
            var clone = (SdtBlock) person.CloneNode(true);
            SdtId id = clone.SdtProperties?.Elements<SdtId>().FirstOrDefault();
            id?.Remove();

            // Add the clone as the new last element.
            person.InsertAfterSelf(clone);
        }
    }
}
