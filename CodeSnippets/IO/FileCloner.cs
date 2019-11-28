//
// FileCloner.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.IO;

namespace CodeSnippets.IO
{
    /// <summary>
    /// This class demonstrates multiple ways to clone files stored in the file system.
    /// In all cases, the source file is stored in the file system. Where the return type
    /// is a <see cref="MemoryStream"/>, the destination file will be stored only on that
    /// <see cref="MemoryStream"/>. Where the return type is a <see cref="FileStream"/>,
    /// the destination file will be stored in the file system and opened on that
    /// <see cref="FileStream"/>.
    /// </summary>
    /// <remarks>
    /// The contents of the <see cref="MemoryStream"/> instances returned by the sample
    /// methods can be written to a file as follows:
    ///
    ///     var stream = ReadAllBytesToMemoryStream(sourcePath);
    ///     File.WriteAllBytes(destPath, stream.GetBuffer());
    ///
    /// You can use <see cref="MemoryStream.GetBuffer"/> in cases where the MemoryStream
    /// was created using <see cref="MemoryStream()"/> or <see cref="MemoryStream(int)"/>.
    /// In other cases, you can use the <see cref="MemoryStream.ToArray"/> method, which
    /// copies the internal buffer to a new byte array. Thus, GetBuffer() should be a tad
    /// faster.
    /// </remarks>
    public static class FileCloner
    {
        public static MemoryStream ReadAllBytesToMemoryStream(string path)
        {
            byte[] buffer = File.ReadAllBytes(path);
            var destStream = new MemoryStream(buffer.Length);
            destStream.Write(buffer, 0, buffer.Length);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static MemoryStream CopyFileStreamToMemoryStream(string path)
        {
            using FileStream sourceStream = File.OpenRead(path);
            var destStream = new MemoryStream((int) sourceStream.Length);
            sourceStream.CopyTo(destStream);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static FileStream CopyFileStreamToFileStream(string sourcePath, string destPath)
        {
            using FileStream sourceStream = File.OpenRead(sourcePath);
            FileStream destStream = File.Create(destPath);
            sourceStream.CopyTo(destStream);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static FileStream CopyFileAndOpenFileStream(string sourcePath, string destPath)
        {
            File.Copy(sourcePath, destPath, true);
            return new FileStream(destPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
    }
}
