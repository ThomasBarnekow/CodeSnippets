using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Threading;
using Xunit;

namespace CodeSnippets.Windows.Tests
{
    public class RandomTests
    {
        [Fact]
        public void TestRandomNumberGenerator()
        {
            const int n = 100;
            const int m = 100;

            var hexStringSet = new HashSet<string>();
            var hexStringSequences = new List<List<string>>();

            for (var i = 0; i < n; i++)
            {
                var rnd = new Random();

                var hexStringSequence = new List<string>();
                hexStringSequences.Add(hexStringSequence);

                for (var j = 0; j < m; j++)
                {
                    int randomNumber = rnd.Next(1, int.MaxValue);
                    var hexString = randomNumber.ToString("X8");

                    hexStringSet.Add(hexString);
                    hexStringSequence.Add(hexString);
                }
            }

            // Assert.True(hexStringSet.Count < n * m);
            Assert.Equal(m, hexStringSet.Count);

            for (var j = 1; j < n; j++)
            {
                Assert.Equal(hexStringSequences[0], hexStringSequences[j]);
            }
        }

        [Fact]
        public void TestCryptographicRandomNumberGenerator()
        {
            const int n = 100;
            const int m = 100;

            var hexStringSet = new HashSet<string>();
            var hexStringSequences = new List<List<string>>();

            for (var i = 0; i < n; i++)
            {
                using var generator = new RNGCryptoServiceProvider();

                var hexStringSequence = new List<string>();
                hexStringSequences.Add(hexStringSequence);

                for (var j = 0; j < m; j++)
                {
                    string hexString = CryptographicRandonNumberGenerator.CreateRandomLongHexNumber(generator, 0x7f);

                    hexStringSet.Add(hexString);
                    hexStringSequence.Add(hexString);
                }
            }

            Assert.Equal(n * m, hexStringSet.Count);

            for (var i = 0; i < n - 1; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    Assert.NotEqual(hexStringSequences[i], hexStringSequences[j]);
                }
            }
        }
    }

    public static class CryptographicRandonNumberGenerator
    {
        /// <summary>
        /// Creates an ST_LongHexNumber value, masking the most significant byte with
        /// the given <paramref name="msbMask" />.
        /// </summary>
        /// <param name="generator"></param>
        /// <param name="msbMask">The most significant byte mask.</param>
        public static string CreateRandomLongHexNumber(RNGCryptoServiceProvider generator, byte msbMask = 0xff)
        {
            // Create a four-byte random number, noting that the first byte (data[0])
            // will become the most significant byte in the string value created by
            // the ToHexString() method.
            var data = new byte[4];
            generator.GetNonZeroBytes(data);
            data[0] &= msbMask;

            return data.ToHexString();
        }

        /// <summary>
        /// Converts the given value into a hexadecimal string, with the first
        /// byte in the list being the most significant byte in the resulting
        /// string.
        /// </summary>
        /// <param name="source">The list of bytes to be converted.</param>
        /// <returns>A hexadecimal string.</returns>
        public static string ToHexString(this IReadOnlyList<byte> source)
        {
            var dest = new char[source.Count * 2];

            var i = 0;
            var j = 0;

            while (i < source.Count)
            {
                byte b = source[i++];
                dest[j++] = ToCharUpper(b >> 4);
                dest[j++] = ToCharUpper(b);
            }

            return new string(dest);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static char ToCharUpper(int value)
        {
            value &= 0xF;
            value += '0';

            if (value > '9')
            {
                value += ('A' - ('9' + 1));
            }

            return (char)value;
        }
    }
}
