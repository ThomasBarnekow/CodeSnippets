using System;
using System.Collections.Generic;
using Xunit;

namespace CodeSnippets.Tests
{
    public class RandomTests
    {
        [Fact]
        public void TestRandomNumberGenerator()
        {
            const int n = 10;
            const int m = 100;

            var hexStringSequences = new List<List<string>>();

            for (var i = 0; i < n; i++)
            {
                var rnd = new Random();
                var hexStringSequence = new List<string>();
                hexStringSequences.Add(hexStringSequence);

                for (var j = 0; j < m; j++)
                {
                    int randomNumber = rnd.Next(1, int.MaxValue);
                    hexStringSequence.Add(randomNumber.ToString("X8"));
                }
            }

            for (var i = 1; i < n; i++)
            {
                Assert.NotEqual(hexStringSequences[0], hexStringSequences[i]);
            }
        }
    }
}
