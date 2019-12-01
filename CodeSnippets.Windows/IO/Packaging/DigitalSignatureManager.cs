//
// PackageSigner.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;

namespace CodeSnippets.Windows.IO.Packaging
{
    public static class DigitalSignatureManager
    {
        public static void Sign(Package package)
        {
            var dsm = new PackageDigitalSignatureManager(package)
            {
                CertificateOption = CertificateEmbeddingOption.InSignaturePart
            };

            List<Uri> parts = package.GetParts()
                .Select(part => part.Uri)
                .Concat(new[]
                {
                    // Include the DigitalSignatureOriginPart and corresponding
                    // relationship part, since those will only be added when
                    // signing.
                    dsm.SignatureOrigin,
                    PackUriHelper.GetRelationshipPartUri(dsm.SignatureOrigin)
                })
                .ToList();

            dsm.Sign(parts);
        }

        public static VerifyResult VerifySignature(Package package)
        {
            var dsm = new PackageDigitalSignatureManager(package);
            return dsm.VerifySignatures(true);
        }
    }
}
