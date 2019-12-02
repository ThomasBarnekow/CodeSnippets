//
// PackageSignerTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using CodeSnippets.IO;
using CodeSnippets.Windows.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace CodeSnippets.Windows.Tests.IO.Packaging
{
    public class DigitalSignatureManagerTests
    {
        [Fact]
        public void Sign_NoCertificatePassed_ExcelWorkbookSigned()
        {
            const string path = "Resources\\UnsignedWorkbook.xlsx";
            using MemoryStream stream = FileCloner.CopyFileStreamToMemoryStream(path);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, true))
            {
                DigitalSignatureManager.Sign(spreadsheetDocument.Package);
            }

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                VerifyResult verifyResult = DigitalSignatureManager.VerifySignature(spreadsheetDocument.Package);
                Assert.Equal(VerifyResult.Success, verifyResult);
            }

            File.WriteAllBytes("SignedWorkbook.xlsx", stream.ToArray());
        }

        [Fact]
        public void Sign_CertificatePassed_ExcelWorkbookSigned()
        {
            const string path = "Resources\\UnsignedWorkbook.xlsx";
            using MemoryStream stream = FileCloner.CopyFileStreamToMemoryStream(path);

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, true))
            {
                X509Certificate2 certificate = DigitalSignatureManager
                    .GetSigningCertificates()
                    .OfType<X509Certificate2>()
                    .First();

                DigitalSignatureManager.Sign(spreadsheetDocument.Package, certificate);
            }

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
            {
                VerifyResult verifyResult = DigitalSignatureManager.VerifySignature(spreadsheetDocument.Package);
                Assert.Equal(VerifyResult.Success, verifyResult);
            }

            File.WriteAllBytes("SignedWorkbook.xlsx", stream.ToArray());
        }
    }
}
