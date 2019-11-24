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
