using System.IO;
using Xunit;

namespace CodeSnippets.Tests.OpenXml.Spreadsheet
{
    public class EmbeddedPackagePartTests
    {
        [Fact]
        public void DoNotNeedToSeek()
        {
            using FileStream stream = File.Open("Resources\\UnsignedWorkbook.xlsx", FileMode.Open, FileAccess.Read);

            Assert.Equal(0, stream.Position);
            Assert.NotEqual(0, stream.Length);

            stream.Seek(0, SeekOrigin.End);
            Assert.Equal(stream.Length, stream.Position);
        }
    }
}
