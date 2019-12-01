//
// DigitalSignatureManager.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;
using CodeSnippets.Xml;
using DocumentFormat.OpenXml.Packaging;

namespace CodeSnippets.IO.Packaging
{
    public class DigitalSignatureManager
    {
        #region Certificate Store

        /// <summary>
        /// Retrieves the valid digital signature certificates from the specified
        /// certificate store.
        /// </summary>
        /// <param name="storeName">The store name.</param>
        /// <param name="storeLocation">The store location.</param>
        /// <returns>The collection of valid digital signature <see cref="X509Certificate2"/> instances.</returns>
        public static IEnumerable<X509Certificate2> GetDigitalSignatureCertificates(
            StoreName storeName = StoreName.My,
            StoreLocation storeLocation = StoreLocation.CurrentUser)
        {
            using var store = new X509Store(storeName, storeLocation);
            store.Open(OpenFlags.ReadOnly);

            return store
                .Certificates
                .OfType<X509Certificate2>()
                .Where(IsValidDigitalSignatureCertificate)
                .ToList();
        }

        private static bool IsValidDigitalSignatureCertificate(X509Certificate2 certificate) =>
            certificate.HasPrivateKey &&
            certificate.NotAfter > DateTime.Now &&
            certificate.Subject != certificate.Issuer &&
            IsUsedForDigitalSignature(certificate);

        private static bool IsUsedForDigitalSignature(X509Certificate2 certificate) =>
            certificate
                .Extensions
                .OfType<X509KeyUsageExtension>()
                .Any(e => (e.KeyUsages & X509KeyUsageFlags.DigitalSignature) != 0);

        #endregion

        #region Digital Signature

        /// <summary>
        /// Signs the given <paramref name="openXmlPackage"/>, using the given
        /// <paramref name="certificate"/>.
        /// </summary>
        /// <param name="openXmlPackage">The <see cref="OpenXmlPackage"/>.</param>
        /// <param name="certificate">The <see cref="X509Certificate2"/>.</param>
        public static void Sign(OpenXmlPackage openXmlPackage, X509Certificate2 certificate)
        {
            if (openXmlPackage == null) throw new ArgumentNullException(nameof(openXmlPackage));
            if (certificate == null) throw new ArgumentNullException(nameof(certificate));

            RSA privateKey = certificate.GetRSAPrivateKey();
            using SHA256 hashAlgorithm = SHA256.Create();

            // Create KeyInfo.
            var keyInfo = new KeyInfo();
            keyInfo.AddClause(new KeyInfoX509Data(certificate));

            // Create a Signature XmlElement.
            var signedXml = new SignedXml { SigningKey = privateKey, KeyInfo = keyInfo };
            signedXml.Signature.Id = Constants.PackageSignatureId;
            signedXml.SignedInfo.SignatureMethod = Constants.SignatureMethod;
            signedXml.AddReference(CreatePackageObjectReference());
            signedXml.AddObject(CreatePackageObject(openXmlPackage.Package, hashAlgorithm));
            signedXml.ComputeSignature();
            XmlElement signature = signedXml.GetXml();

            // Get or create the DigitalSignatureOriginPart.
            DigitalSignatureOriginPart dsOriginPart =
                openXmlPackage.GetPartsOfType<DigitalSignatureOriginPart>().FirstOrDefault() ??
                openXmlPackage.AddNewPart<DigitalSignatureOriginPart>();

            var xmlSignaturePart = dsOriginPart.AddNewPart<XmlSignaturePart>();

            // Write the Signature XmlElement to the XmlSignaturePart.
            using Stream stream = xmlSignaturePart.GetStream(FileMode.Create, FileAccess.Write);
            using XmlWriter writer = XmlWriter.Create(stream);
            signature.WriteTo(writer);
        }

        private static Reference CreatePackageObjectReference()
        {
            // Create a Reference with an URI as the reference target.
            return new Reference("#" + Constants.PackageObjectId)
            {
                Type = Constants.TypeObject,
                DigestMethod = Constants.DigestMethod
            };
        }

        private static DataObject CreatePackageObject(Package package, SHA256 hashAlgorithm)
        {
            var xmlDocument = new XmlDocument();

            XmlNode documentElement = xmlDocument.AppendChild(xmlDocument.CreateElement("Root"));
            documentElement.AppendChild(CreateManifest(xmlDocument, package, hashAlgorithm));
            documentElement.AppendChild(CreateSignatureProperties(xmlDocument));

            return new DataObject
            {
                Id = Constants.PackageObjectId,
                Data = documentElement.ChildNodes
            };
        }

        private static XmlElement CreateManifest(
            XmlDocument xmlDocument,
            Package package,
            SHA256 hashAlgorithm)
        {
            XmlElement manifest = xmlDocument.CreateElement(DS.Manifest);

            // At the moment, we sign all parts parts, except for relationship parts.
            // Those require special treatment.
            IEnumerable<PackagePart> parts = package
                .GetParts()
                .Where(part => part.ContentType != "application/vnd.openxmlformats-package.relationships+xml");

            foreach (PackagePart part in parts)
            {
                AppendReference(xmlDocument, manifest, part, hashAlgorithm);
            }

            return manifest;
        }

        private static void AppendReference(
            XmlDocument xmlDocument,
            XmlElement manifest,
            PackagePart part,
            SHA256 hashAlgorithm)
        {
            Reference reference = CreateReference(part, hashAlgorithm);
            manifest.AppendChild(xmlDocument.ImportNode(reference.GetXml(), true));
        }

        private static Reference CreateReference(PackagePart part, SHA256 hashAlgorithm)
        {
            using Stream stream = part.GetStream();
            byte[] digestValue = hashAlgorithm.ComputeHash(stream);

            return new Reference
            {
                Uri = $"{part.Uri}?ContentType={part.ContentType}",
                DigestMethod = Constants.DigestMethod,
                DigestValue = digestValue
            };
        }

        private static XmlElement CreateSignatureProperties(XmlDocument xmlDocument)
        {
            const string xmlDateTimeFormat = "YYYY-MM-DDThh:mm:ssTZD";
            const string dateTimeFormat = "yyyy-MM-ddTHH:mm:ssZ";

            string xmlDateTimeValue = DateTimeToXmlFormattedTime(DateTime.UtcNow, dateTimeFormat);

            XmlElement signatureProperties = xmlDocument.CreateElement(DS.SignatureProperties);
            XmlElement signatureProperty = xmlDocument.CreateElement(DS.SignatureProperty);
            XmlElement signatureTime = xmlDocument.CreateElement(MDSSI.SignatureTime);
            XmlElement format = xmlDocument.CreateElement(MDSSI.Format, xmlDateTimeFormat);
            XmlElement value = xmlDocument.CreateElement(MDSSI.Value, xmlDateTimeValue);

            signatureProperties.AppendChild(signatureProperty);
            signatureProperty.AppendChild(signatureTime);
            signatureTime.AppendChild(format);
            signatureTime.AppendChild(value);

            signatureProperty.SetAttribute("Id", Constants.SignatureTimeId);
            signatureProperty.SetAttribute("Target", "#" + Constants.PackageSignatureId);

            return signatureProperties;
        }

        private static string DateTimeToXmlFormattedTime(DateTime dateTime, string format)
        {
            return dateTime.ToString(format, new DateTimeFormatInfo { FullDateTimePattern = format });
        }

        #endregion

        private static class Constants
        {
            public const string SignatureMethod = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
            public const string DigestMethod = "http://www.w3.org/2001/04/xmlenc#sha256";

            public const string TypeObject = "http://www.w3.org/2000/09/xmldsig#Object";

            public const string PackageSignatureId = "idPackageSignature";
            public const string PackageObjectId = "idPackageObject";
            public const string SignatureTimeId = "idSignatureTime";
        }
    }
}
