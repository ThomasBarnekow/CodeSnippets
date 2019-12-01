//
// DigitalSignatureTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using CodeSnippets.IO;
using CodeSnippets.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace CodeSnippets.Tests.IO.Packaging
{
    public class DigitalSignatureManagerTests
    {
        public DigitalSignatureManagerTests(ITestOutputHelper output)
        {
            _output = output;
        }

        private readonly ITestOutputHelper _output;

        [Fact]
        public void GetDigitalSignatureCertificates_MyStoreCurrentUser_CertificatesListed()
        {
            IEnumerable<X509Certificate2> certificates = DigitalSignatureManager.GetDigitalSignatureCertificates();

            foreach (X509Certificate2 certificate in certificates)
            {
                _output.WriteLine(certificate.ToString());
            }
        }

        [Fact]
        public void Sign_UnsignedWorkbook_SuccessfullySigned()
        {
            // Load our unsigned Excel workbook into a MemoryStream for processing.
            const string path = "Resources\\UnsignedWorkbook.xlsx";
            using MemoryStream stream = FileCloner.ReadAllBytesToMemoryStream(path);

            // Open the SpreadsheetDocument on the MemoryStream and sign it, using
            // the first digital signature certificate that we find.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, true))
            {
                X509Certificate2 certificate = DigitalSignatureManager.GetDigitalSignatureCertificates().First();
                DigitalSignatureManager.Sign(spreadsheetDocument, certificate);
            }

            // Save the signed Excel workbook to disk. When you open this workbook,
            // Excel should tell you that it has found a valid signature.
            File.WriteAllBytes("SignedWorkbook.xlsx", stream.ToArray());
        }
    }
}
