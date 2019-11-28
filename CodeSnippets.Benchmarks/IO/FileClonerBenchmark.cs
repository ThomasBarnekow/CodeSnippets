//
// ClonerBenchmark.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using CodeSnippets.IO;
using CodeSnippets.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CodeSnippets.Benchmarks.IO
{
    [SuppressMessage("ReSharper", "MemberCanBePrivate.Global")]
    [SuppressMessage("ReSharper", "UnassignedField.Global")]
    public class FileClonerBenchmark
    {
        #region Setup and Helpers

        private const string SourcePath = "Source.docx";
        private const string DestPath = "Destination.docx";

        [Params(1, 10, 100, 1000)]
        public static int ParagraphCount;

        [GlobalSetup]
        public void GlobalSetup()
        {
            CreateTestDocument(SourcePath);
            CreateTestDocument(DestPath);
        }

        private static void CreateTestDocument(string path)
        {
            const string sentence = "The quick brown fox jumps over the lazy dog.";
            string text = string.Join(" ", Enumerable.Range(0, 22).Select(i => sentence));
            IEnumerable<string> texts = Enumerable.Range(0, ParagraphCount).Select(i => text);
            using WordprocessingDocument unused = WordprocessingDocumentFactory.Create(path, texts);
        }

        private static void ChangeWordprocessingDocument(WordprocessingDocument wordDocument)
        {
            Body body = wordDocument.MainDocumentPart.Document.Body;
            Text text = body.Descendants<Text>().First();
            text.Text = DateTimeOffset.UtcNow.Ticks.ToString();
        }

        #endregion

        #region Benchmarks

        [Benchmark(Baseline = true)]
        public void DoWorkUsingReadAllBytesToMemoryStream()
        {
            using MemoryStream destStream = FileCloner.ReadAllBytesToMemoryStream(SourcePath);

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(destStream, true))
            {
                ChangeWordprocessingDocument(wordDocument);
            }

            File.WriteAllBytes(DestPath, destStream.GetBuffer());
        }

        [Benchmark]
        public void DoWorkUsingCopyFileStreamToMemoryStream()
        {
            using MemoryStream destStream = FileCloner.CopyFileStreamToMemoryStream(SourcePath);

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(destStream, true))
            {
                ChangeWordprocessingDocument(wordDocument);
            }

            File.WriteAllBytes(DestPath, destStream.GetBuffer());
        }

        [Benchmark]
        public void DoWorkUsingCopyFileStreamToFileStream()
        {
            using FileStream destStream = FileCloner.CopyFileStreamToFileStream(SourcePath, DestPath);
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(destStream, true);
            ChangeWordprocessingDocument(wordDocument);
        }

        [Benchmark]
        public void DoWorkUsingCopyFileAndOpenFileStream()
        {
            using FileStream destStream = FileCloner.CopyFileAndOpenFileStream(SourcePath, DestPath);
            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(destStream, true);
            ChangeWordprocessingDocument(wordDocument);
        }

        [Benchmark]
        public void DoWorkCloningOpenXmlPackage()
        {
            using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(SourcePath, false);
            using var wordDocument = (WordprocessingDocument) sourceWordDocument.Clone(DestPath, true);
            ChangeWordprocessingDocument(wordDocument);
        }

        #endregion
    }
}
