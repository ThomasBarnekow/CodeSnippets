//
// Program.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using BenchmarkDotNet.Running;
using CodeSnippets.Benchmarks.IO;

namespace CodeSnippets.Benchmarks
{
    public static class Program
    {
        public static void Main()
        {
            BenchmarkRunner.Run<FileClonerBenchmark>();
        }
    }
}
