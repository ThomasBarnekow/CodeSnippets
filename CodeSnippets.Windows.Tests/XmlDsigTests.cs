using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using Xunit;

namespace CodeSnippets.Windows.Tests
{
    public class XmlDsigTests
    {
        // Note: The following might be useful when related exceptions are thrown.
        // AppContext.SetSwitch("Switch.System.Security.Cryptography.Xml.UseInsecureHashAlgorithms", true);

        private static readonly RSA Rsa = new RSACng();

        private static XmlElement CreateDetachedSignature(string path, RSA rsa)
        {
            // Create Reference.
            byte[] buffer = File.ReadAllBytes(path);
            using var stream = new MemoryStream(buffer);
            var reference = new Reference(stream)
            {
                Uri = Uri.EscapeUriString(path.Replace('\\', '/')),
                DigestMethod = SignedXml.XmlDsigSHA256Url
            };

            // Create KeyInfo.
            var keyInfo = new KeyInfo();
            keyInfo.AddClause(new RSAKeyValue(rsa));

            // Create SignedXml.
            var signedXml = new SignedXml
            {
                SigningKey = rsa,
                KeyInfo = keyInfo
            };

            signedXml.AddReference(reference);
            signedXml.ComputeSignature();

            return signedXml.GetXml();
        }

        [Fact]
        public void CanCheckDetachedSignature()
        {
            const string path = "Resources\\UnsignedDocument.docx";
            XmlElement signature = CreateDetachedSignature(path, Rsa);

            var signedXml = new SignedXml();
            signedXml.LoadXml(signature);

            // Remove all Reference elements because they are targeting a URI that
            // we will not be able to resolve.
            ArrayList uriReferenceArrayList = signedXml.SignedInfo.References;
            List<Reference> uriReferences = uriReferenceArrayList.OfType<Reference>().ToList();
            uriReferenceArrayList.Clear();

            // Add those Reference elements back, making sure they reference a Stream
            // rather than a URI.
            foreach (Reference uriReference in uriReferences)
            {
                // Get the XML representation of the URI Reference's state.
                XmlElement uriReferenceState = uriReference.GetXml();

                // Create a Stream from the referenced URI.
                string uri = uriReferenceState.GetAttribute("URI");
                byte[] buffer = File.ReadAllBytes(uri);
                var stream = new MemoryStream(buffer);

                // Create a Reference having a Stream as its target.
                var streamReference = new Reference(stream);
                streamReference.LoadXml(uriReferenceState);

                // Add the Stream reference back to the SignedInfo.
                signedXml.AddReference(streamReference);
            }

            // Check signature.
            Assert.True(signedXml.CheckSignature());
        }

        [Fact]
        public void CanCreateDetachedSignature()
        {
            // Create KeyInfo.
            XmlElement signature = CreateDetachedSignature("Resources\\UnsignedDocument.docx", Rsa);
            using var writer = new XmlTextWriter("DetachedSignature.xml", new UTF8Encoding(false));
            signature.WriteTo(writer);
        }
    }
}
