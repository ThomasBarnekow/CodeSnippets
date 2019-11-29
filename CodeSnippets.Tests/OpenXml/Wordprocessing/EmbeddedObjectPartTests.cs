//
// OleFileTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Diagnostics.CodeAnalysis;
using System.IO;
using CodeSnippets.Windows;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Wordprocessing
{
    public class EmbeddedObjectPartTests
    {
        [SuppressMessage("ReSharper", "ConvertToUsingDeclaration")]
        private static void ExtractFile(EmbeddedObjectPart part, string destinationFolderPath)
        {
            // Determine the file name and destination path of the binary,
            // structured storage file.
            string binaryFileName = Path.GetFileName(part.Uri.ToString());
            string binaryFilePath = Path.Combine(destinationFolderPath, binaryFileName);

            // Ensure the destination directory exists.
            Directory.CreateDirectory(destinationFolderPath);

            // Copy part contents to structured storage file.
            using (Stream partStream = part.GetStream())
            using (FileStream fileStream = File.Create(binaryFilePath))
            {
                partStream.CopyTo(fileStream);
            }

            // Extract the embedded file from the structured storage file.
            Ole10Native.ExtractFile(binaryFilePath, destinationFolderPath);

            // Remove the structured storage file.
            File.Delete(binaryFilePath);
        }

        [Fact]
        public void CanExtractEmbeddedZipFile()
        {
            const string documentPath = "Resources\\ZipContainer.docx";
            const string destinationFolderPath = "Output";
            string destinationFilePath = Path.Combine(destinationFolderPath, "ZipContents.zip");

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(documentPath, false);

            // Extract all embedded objects.
            foreach (EmbeddedObjectPart part in wordDocument.MainDocumentPart.EmbeddedObjectParts)
            {
                ExtractFile(part, destinationFolderPath);
            }

            Assert.True(File.Exists(destinationFilePath));
        }
    }
}
