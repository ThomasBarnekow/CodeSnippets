using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Xml;
using Disig.TimeStampClient;
using Org.BouncyCastle.Asn1.Tsp;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.X509;
using Xunit;
using Oid = Disig.TimeStampClient.Oid;
using TimeStampToken = Org.BouncyCastle.Tsp.TimeStampToken;

namespace CodeSnippets.Windows.Tests.Security
{
    public class TimeStampingTests
    {
        private const string DsNamespaceUri = "http://www.w3.org/2000/09/xmldsig#";
        private const string XadesNamespaceUri = "http://uri.etsi.org/01903/v1.3.2#";

        private readonly XmlNameTable _nameTable;
        private readonly XmlNamespaceManager _namespaceManager;

        public TimeStampingTests()
        {
            _nameTable = new NameTable();

            _namespaceManager = new XmlNamespaceManager(_nameTable);
            _namespaceManager.AddNamespace("ds", DsNamespaceUri);
            _namespaceManager.AddNamespace("xades", XadesNamespaceUri);
        }

        private XmlElement GetSignatureValue(XmlNode root)
        {
            return (XmlElement) root.SelectSingleNode(
                "descendant-or-self::ds:Signature/ds:SignatureValue",
                _namespaceManager);
        }

        private string GetCanonicalizationMethod(XmlNode root)
        {
            var element = (XmlElement) root.SelectSingleNode(
                "descendant-or-self::ds:Signature/ds:SignedInfo/ds:CanonicalizationMethod",
                _namespaceManager);

            return element?.GetAttribute("Algorithm");
        }

        private static byte[] Canonicalize(XmlNode node, string canonicalizationMethod)
        {
            var document = new XmlDocument();
            document.AppendChild(document.ImportNode(node, true));

            var transform = (Transform) CryptoConfig.CreateFromName(canonicalizationMethod);
            transform.LoadInput(document);

            using var stream = (MemoryStream) transform.GetOutput();
            return stream.ToArray();
        }

        private byte[] GetEncapsulatedTimeStamp(XmlNode root)
        {
            XmlNode textNode = root.SelectSingleNode(
                "descendant-or-self::xades:SignatureTimeStamp/xades:EncapsulatedTimeStamp/text()",
                _namespaceManager);

            return textNode != null ? Convert.FromBase64String(textNode.Value) : null;
        }

        [Fact]
        public void CanTimeStampSignature()
        {
            var document = new XmlDocument();
            document.Load("Resources\\Signature.xml");

            // Get data.
            XmlElement signatureValue = GetSignatureValue(document);
            string canonicalizationMethod = GetCanonicalizationMethod(document);
            byte[] data = Canonicalize(signatureValue, canonicalizationMethod);

            // Compute data digest.
            using var digestAlgorithm = SHA256.Create();
            byte[] dataDigest = digestAlgorithm.ComputeHash(data);

            // Compute Nonce
            var nonce = new byte[8];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(nonce);
            }

            // Request a time stamp token.
            var request = new Request(dataDigest, Oid.SHA256, nonce, certReq: true);
            byte[] derEncodedTimeStampToken = TimeStampClient
                .RequestTimeStampToken("http://timestamp.comodoca.com", request)
                .ToByteArray();

            // Create and validate CmsSignedData.
            var cmsSignedData = new CmsSignedData(derEncodedTimeStampToken);

            SignerInformation signer = cmsSignedData
                .GetSignerInfos()
                .GetSigners()
                .OfType<SignerInformation>()
                .Single();

            X509Certificate cert = cmsSignedData
                .GetCertificates("Collection")
                .GetMatches(signer.SignerID)
                .OfType<X509Certificate>()
                .Single();

            var timeStampToken = new TimeStampToken(cmsSignedData);
            timeStampToken.Validate(cert);

            TstInfo tstInfo = timeStampToken.TimeStampInfo.TstInfo;
            Assert.Equal(dataDigest, tstInfo.MessageImprint.GetHashedMessage());
            Assert.Equal(nonce, tstInfo.Nonce.Value.ToByteArray());
        }

        [Fact]
        public void CanValidateEncapsulatedTimeStamp()
        {
            var document = new XmlDocument(_nameTable);
            document.Load("Resources\\XadesEnvelopedSignature.xml");

            // Create and validate TimeStampToken.
            byte[] derEncodedTimeStampToken = GetEncapsulatedTimeStamp(document);
            var cmsSignedData = new CmsSignedData(derEncodedTimeStampToken);

            SignerInformation signer = cmsSignedData
                .GetSignerInfos()
                .GetSigners()
                .OfType<SignerInformation>()
                .Single();

            X509Certificate cert = cmsSignedData
                .GetCertificates("Collection")
                .GetMatches(signer.SignerID)
                .OfType<X509Certificate>()
                .Single();

            // Create and validate TimeStampToken.
            var timeStampToken = new TimeStampToken(cmsSignedData);
            timeStampToken.Validate(cert);

            // Get original digest value.
            TstInfo tstInfo = timeStampToken.TimeStampInfo.TstInfo;
            byte[] originalSignatureDigestValue = tstInfo.MessageImprint.GetHashedMessage();

            // Compute current digest value.
            XmlElement signatureValue = GetSignatureValue(document);
            string canonicalizationMethod = GetCanonicalizationMethod(document);
            byte[] data = Canonicalize(signatureValue, canonicalizationMethod);

            using var digestAlgorithm = SHA256.Create();
            byte[] currentSignatureDigestValue = digestAlgorithm.ComputeHash(data);

            // Check whether original and current digest values are equal.
            Assert.Equal(originalSignatureDigestValue, currentSignatureDigestValue);
        }
    }
}
