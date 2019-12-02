//
// OpenXmlDigitalSignatureManagerTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using CodeSnippets.IO;
using CodeSnippets.Windows.IO.Packaging;
using CodeSnippets.Windows.OpenXml.Packaging;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace CodeSnippets.Windows.Tests.OpenXml.Packaging
{
    public class OpenXmlDigitalSignatureManagerTests
    {
        [Theory]
        [InlineData("Resources\\UnsignedDocument.docx", "OpenXml_CertPassed_SignedDocument.docx")]
        [InlineData("Resources\\UnsignedPresentation.pptx", "OpenXml_CertPassed_SignedPresentation.pptx")]
        [InlineData("Resources\\UnsignedWorkbook.xlsx", "OpenXml_CertPassed_SignedWorkbook.xlsx")]
        public void Sign_CertificatePassed_Success(string sourcePath, string destPath)
        {
            using MemoryStream stream = FileCloner.CopyFileStreamToMemoryStream(sourcePath);

            using (OpenXmlPackage openXmlPackage = Open(stream, true, sourcePath))
            {
                X509Certificate2 certificate = DigitalSignatureManager
                    .GetSigningCertificates()
                    .OfType<X509Certificate2>()
                    .First();

                PackageDigitalSignature signature = OpenXmlDigitalSignatureManager.Sign(openXmlPackage, certificate);

                VerifyResult verifyResult = signature.Verify();
                Assert.Equal(VerifyResult.Success, verifyResult);
            }

            File.WriteAllBytes(destPath, stream.ToArray());
        }

        [Fact]
        public void Sign_CertificateNotPassed_Success()
        {
            const string path = "Resources\\UnsignedDocument.docx";
            using MemoryStream stream = FileCloner.CopyFileStreamToMemoryStream(path);

            using (OpenXmlPackage openXmlPackage = Open(stream, true, path))
            {
                PackageDigitalSignature signature = OpenXmlDigitalSignatureManager.Sign(openXmlPackage);

                VerifyResult verifyResult = signature.Verify();
                Assert.Equal(VerifyResult.Success, verifyResult);
            }

            File.WriteAllBytes("OpenXml_CertNotPassed_SignedDocument.docx", stream.ToArray());
        }

        private static OpenXmlPackage Open(Stream stream, bool isEditable, string path)
        {
            path = path.ToLowerInvariant();

            if (path.EndsWith(".docx")) return WordprocessingDocument.Open(stream, isEditable);
            if (path.EndsWith(".pptx")) return PresentationDocument.Open(stream, isEditable);
            if (path.EndsWith(".xlsx")) return SpreadsheetDocument.Open(stream, isEditable);

            throw new NotSupportedException();
        }
    }
}
