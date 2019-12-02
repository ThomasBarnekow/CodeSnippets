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
using System.Security.Cryptography.X509Certificates;

namespace CodeSnippets.Windows.IO.Packaging
{
    public static class DigitalSignatureManager
    {
        #region Certificate Store

        /// <summary>
        /// Prompts the user to select the desired signing certificate.
        /// </summary>
        /// <param name="hwndParent">The owner window's handle.</param>
        /// <returns>The selected <see cref="X509Certificate2"/> or null.</returns>
        public static X509Certificate2 PromptForSigningCertificate(IntPtr hwndParent)
        {
            X509Certificate2 certificate = null;

            X509Certificate2Collection certificates = GetSigningCertificates();
            if (certificates.Count > 0)
            {
                X509Certificate2Collection certificate2Collection = X509Certificate2UI.SelectFromCollection(
                    certificates, "Digital Signature", "Select a certificate", X509SelectionFlag.SingleSelection, hwndParent);

                if (certificate2Collection.Count > 0)
                {
                    certificate = certificate2Collection[0];
                }
            }

            return certificate;
        }

        /// <summary>
        /// Retrieves the valid digital signature certificates from the current
        /// user's certificate store.
        /// </summary>
        /// <returns>An <see cref="X509Certificate2Collection"/>.</returns>
        public static X509Certificate2Collection GetSigningCertificates()
        {
            using var x509Store = new X509Store(StoreLocation.CurrentUser);
            x509Store.Open(OpenFlags.OpenExistingOnly);

            X509Certificate2Collection certificates = x509Store
                .Certificates
                .Find(X509FindType.FindByTimeValid, DateTime.Now, true)
                .Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, false);

            for (int index = certificates.Count - 1; index >= 0; --index)
            {
                if (!certificates[index].HasPrivateKey)
                {
                    certificates.RemoveAt(index);
                }
            }

            return certificates;
        }

        #endregion

        #region PackageDigitalSignatureManager

        public static void Sign(Package package)
        {
            var dsm = new PackageDigitalSignatureManager(package)
            {
                CertificateOption = CertificateEmbeddingOption.InSignaturePart
            };

            List<Uri> parts = package.GetParts()
                .Select(part => part.Uri)
                .Where(uri => !PackUriHelper.IsRelationshipPartUri(uri))
                .ToList();

            dsm.Sign(parts);
        }

        public static void Sign(Package package, X509Certificate2 certificate)
        {
            var dsm = new PackageDigitalSignatureManager(package)
            {
                CertificateOption = CertificateEmbeddingOption.InSignaturePart
            };

            List<Uri> parts = package.GetParts()
                .Select(part => part.Uri)
                .Where(uri => !PackUriHelper.IsRelationshipPartUri(uri))
                .ToList();

            dsm.Sign(parts, certificate);
        }

        public static VerifyResult VerifySignature(Package package)
        {
            var dsm = new PackageDigitalSignatureManager(package);
            return dsm.VerifySignatures(true);
        }

        #endregion
    }
}
