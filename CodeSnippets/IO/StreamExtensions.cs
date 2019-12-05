//
// StreamExtensions.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.IO;
using System.Threading.Tasks;

namespace CodeSnippets.IO
{
    public static class StreamExtensions
    {
        public static byte[] ToArray(this Stream stream)
        {
            switch (stream)
            {
                case null:
                {
                    throw new ArgumentNullException(nameof(stream));
                }
                case MemoryStream memoryStream:
                {
                    return memoryStream.ToArray();
                }
                default:
                {
                    if (stream.CanSeek) stream.Seek(0, SeekOrigin.Begin);
                    using var destination = new MemoryStream();
                    stream.CopyTo(destination);
                    return destination.ToArray();
                }
            }
        }

        public static async Task<byte[]> ToArrayAsync(this Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));

            if (stream.CanSeek) stream.Seek(0, SeekOrigin.Begin);
            using var destination = new MemoryStream();
            await stream.CopyToAsync(destination).ConfigureAwait(false);
            return destination.ToArray();
        }
    }
}
