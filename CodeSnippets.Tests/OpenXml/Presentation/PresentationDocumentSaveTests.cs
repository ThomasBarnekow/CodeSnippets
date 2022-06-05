using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Linq;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Presentation;

public class PresentationDocumentSaveTests
{
    [Fact]
    public void CanRemoveSectionListUsingStronglyTypedClasses()
    {
        using MemoryStream stream = ReadAllBytesToMemoryStream("Resources\\base.pptx");
        using (PresentationDocument doc = PresentationDocument.Open(stream, true))
        {
            SectionList sectionList = doc.PresentationPart!.Presentation.PresentationExtensionList!.Descendants<SectionList>().First();
            sectionList.Remove();
        }

        byte[] bytes = stream.ToArray();
        using var newStream = new MemoryStream(bytes);
        using PresentationDocument newDoc = PresentationDocument.Open(newStream, true);
        Assert.Empty(newDoc.PresentationPart!.Presentation.PresentationExtensionList!.Descendants<SectionList>());
    }

    [Fact]
    public void CanRemoveSectionListUsingLinqToXml()
    {
        var hasher = SHA1.Create();

        using MemoryStream stream = ReadAllBytesToMemoryStream("Resources\\base.pptx");
        string hashBefore = Convert.ToBase64String(hasher.ComputeHash(stream));

        using PresentationDocument doc = PresentationDocument.Open(stream, true);

        XElement presentation = doc.PresentationPart!.GetXElement()!;
        XElement sectionList = presentation.Elements(P.extLst).Descendants(P14.sectionLst).First();
        sectionList.Remove();

        doc.PresentationPart.SaveXElement();
        doc.PresentationPart.Presentation.Save();
        doc.Close();

        byte[] bytes = stream.ToArray();
        using var newStream = new MemoryStream(bytes);
        string hashAfter = Convert.ToBase64String(hasher.ComputeHash(newStream));

        using PresentationDocument newDoc = PresentationDocument.Open(newStream, true);

        Assert.Empty(newDoc.PresentationPart!.Presentation.PresentationExtensionList!.Descendants<SectionList>());
    }

    private static MemoryStream ReadAllBytesToMemoryStream(string path)
    {
        byte[] buffer = File.ReadAllBytes(path);
        var destStream = new MemoryStream(buffer.Length);
        destStream.Write(buffer, 0, buffer.Length);
        destStream.Seek(0, SeekOrigin.Begin);
        return destStream;
    }
}
