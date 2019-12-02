//
// OpenXmlPackageSignatureManager.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using CodeSnippets.Windows.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;

namespace CodeSnippets.Windows.OpenXml.Packaging
{
    public static class OpenXmlDigitalSignatureManager
    {
        private static readonly IReadOnlyCollection<string> ExcludedContentTypes = new List<string>
        {
            "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            "application/vnd.openxmlformats-package.core-properties+xml",
            "application/vnd.openxmlformats-package.relationships+xml"
        };

        private static readonly IReadOnlyCollection<string> ExcludedRelationshipTypes = new List<string>
        {
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        };

        public static PackageDigitalSignature Sign(OpenXmlPackage openXmlPackage)
        {
            if (openXmlPackage == null) throw new ArgumentNullException(nameof(openXmlPackage));

            X509Certificate2 certificate = DigitalSignatureManager.PromptForSigningCertificate(IntPtr.Zero);
            return certificate != null ? Sign(openXmlPackage, certificate) : null;
        }

        public static PackageDigitalSignature Sign(
            OpenXmlPackage openXmlPackage,
            X509Certificate2 certificate,
            string signatureId =  "idPackageSignature")
        {
            if (openXmlPackage == null) throw new ArgumentNullException(nameof(openXmlPackage));
            if (certificate == null) throw new ArgumentNullException(nameof(certificate));
            if (signatureId == null) throw new ArgumentNullException(nameof(signatureId));

            Package package = openXmlPackage.Package;
            var dsm = new PackageDigitalSignatureManager(package)
            {
                CertificateOption = CertificateEmbeddingOption.InSignaturePart
            };

            return dsm.Sign(GetParts(package), certificate, GetRelationshipSelectors(package), signatureId);
        }

        private static IEnumerable<Uri> GetParts(Package package)
        {
            return package.GetParts()
                .Where(part => !ExcludedContentTypes.Contains(part.ContentType))
                .Select(part => part.Uri)
                .ToList();
        }

        private static IEnumerable<PackageRelationshipSelector> GetRelationshipSelectors(Package package)
        {
            return new PackageRelationships(package)
                .Where(r => !ExcludedRelationshipTypes.Contains(r.RelationshipType))
                .Select(r => new PackageRelationshipSelector(r.SourceUri, PackageRelationshipSelectorType.Id, r.Id))
                .ToList();
        }

        public static VerifyResult VerifySignature(OpenXmlPackage openXmlPackage)
        {
            if (openXmlPackage == null) throw new ArgumentNullException(nameof(openXmlPackage));

            var dsm = new PackageDigitalSignatureManager(openXmlPackage.Package);
            return dsm.VerifySignatures(true);
        }
    }
}
