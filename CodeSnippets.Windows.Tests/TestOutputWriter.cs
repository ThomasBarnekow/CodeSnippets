//
// TestOutputWriter.cs
//
// Copyright 2022 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info
//

using System.Collections.Generic;
using System.IO;
using System.Text;
using JetBrains.Annotations;
using Xunit.Abstractions;

namespace CodeSnippets.Windows.Tests
{
    [PublicAPI]
    public class TestOutputWriter : TextWriter
    {
        private readonly ITestOutputHelper _output;

        private readonly List<char> _buffer = new List<char>();

        public TestOutputWriter(ITestOutputHelper output)
        {
            _output = output;
        }

        /// <inheritdoc />
        public override Encoding Encoding => Encoding.UTF8;

        /// <inheritdoc />
        public override void Flush()
        {
            WriteLine();
        }

        /// <inheritdoc />
        public override void Write(char value)
        {
            _buffer.Add(value);
        }

        /// <inheritdoc />
        public override void WriteLine()
        {
            _output.WriteLine(new string(_buffer.ToArray()));
            _buffer.Clear();
        }

        /// <inheritdoc />
        public override void WriteLine(string value)
        {
            _output.WriteLine(value);
        }
    }
}
