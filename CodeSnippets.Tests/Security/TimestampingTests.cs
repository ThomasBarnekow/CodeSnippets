using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Security.Cryptography.Pkcs;
using System.Threading.Tasks;
using Xunit;

namespace CodeSnippets.Tests.Security
{
    public class TimestampingTests
    {
        [Fact]
        public void CanUseTimeStampingAuthority()
        {

        }

        private static async Task<Rfc3161TimestampToken> GetTimestamp(SignedCms toSign, CmsSigner newSigner, Uri timeStampAuthorityUri)
        {
            if (timeStampAuthorityUri == null)
            {
                throw new ArgumentNullException(nameof(timeStampAuthorityUri));
            }

            // This example figures out which signer is new by it being "the only signer"
            if (toSign.SignerInfos.Count > 0)
            {
                throw new ArgumentException("We must have only one signer", nameof(toSign));
            }

            toSign.ComputeSignature(newSigner);
            SignerInfo newSignerInfo = toSign.SignerInfos[0];

            var nonce = new byte[8];
            using (var rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(nonce);
            }

            var request = Rfc3161TimestampRequest.CreateFromSignerInfo(
                newSignerInfo,
                HashAlgorithmName.SHA384,
                requestSignerCertificates: true,
                nonce: nonce);

            var client = new HttpClient();
            var content = new ReadOnlyMemoryContent(request.Encode());
            content.Headers.ContentType = new MediaTypeHeaderValue("application/timestamp-query");
            HttpResponseMessage httpResponse = await client.PostAsync(timeStampAuthorityUri, content).ConfigureAwait(false);
            if (!httpResponse.IsSuccessStatusCode)
            {
                throw new CryptographicException(
                    $"There was a error from the timestamp authority. It responded with {httpResponse.StatusCode} {(int)httpResponse.StatusCode}: {httpResponse.Content}");
            }

            if (httpResponse.Content.Headers.ContentType.MediaType != "application/timestamp-reply")
            {
                throw new CryptographicException("The reply from the time stamp server was in a invalid format.");
            }

            byte[] data = await httpResponse.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
            return request.ProcessResponse(data, out int _);
        }
    }
}
