//
// PackageSignerTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.IO;
using System.IO.Packaging;
using System.Security.Cryptography.X509Certificates;
using CodeSnippets.IO;
using CodeSnippets.Windows.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace CodeSnippets.Windows.Tests.IO.Packaging
{
    public class DigitalSignatureManagerTests
    {
        private readonly ITestOutputHelper _output;

        public DigitalSignatureManagerTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void CanSignExcelWorkbook()
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
        public void CanAccessCertificateStore()
        {
            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadOnly);
            foreach (X509Certificate2 certificate in store.Certificates)
            {
                _output.WriteLine(certificate.Subject);
            }
        }
    }
}
