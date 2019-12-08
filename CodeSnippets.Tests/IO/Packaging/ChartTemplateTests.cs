//
// ChartTemplateTests.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using CodeSnippets.IO;
using Xunit;

namespace CodeSnippets.Tests.IO.Packaging
{
    public class ChartTemplateTests
    {
        private static XElement LoadRootElement(PackagePart packagePart)
        {
            using Stream stream = packagePart.GetStream(FileMode.Open, FileAccess.Read);
            return XElement.Load(stream);
        }

        private static void SaveRootElement(PackagePart packagePart, XElement rootElement)
        {
            using Stream stream = packagePart.GetStream(FileMode.Create, FileAccess.Write);
            rootElement.Save(stream, SaveOptions.DisableFormatting);
        }

        [Fact]
        public void LoadRootElement_Chart_SuccessfullyLoaded()
        {
            using Package package = Package.Open("Resources\\ChartTemplate.crtx", FileMode.Open, FileAccess.Read);
            PackagePart packagePart = package.GetPart(new Uri("/chart/chart.xml", UriKind.Relative));

            XElement rootElement = LoadRootElement(packagePart);

            Assert.Equal(C.chartSpace, rootElement.Name);
            Assert.NotEmpty(rootElement.Elements(C.chart).Elements(C.title));
            Assert.NotEmpty(rootElement.Elements(C.chart).Elements(C.plotArea));
            Assert.NotEmpty(rootElement.Elements(C.chart).Elements(C.legend));
        }

        [Fact]
        public void SaveRootElement_Chart_SuccessfullySaved()
        {
            using MemoryStream stream = FileCloner.ReadAllBytesToMemoryStream("Resources\\ChartTemplate.crtx");

            using (Package package = Package.Open(stream, FileMode.Open, FileAccess.ReadWrite))
            {
                // Get the package part, its root element, and the c:lang element.
                // Note that the val attribute value is "en-US".
                PackagePart packagePart = package.GetPart(new Uri("/chart/chart.xml", UriKind.Relative));
                XElement rootElement = LoadRootElement(packagePart);
                XElement lang = rootElement.Elements(C.lang).First();

                Assert.Equal("en-US", (string) lang.Attribute(C.val));

                // Change the val attribute value to "de-DE" and save the root element.
                lang.SetAttributeValue(C.val, "de-DE");

                // Act, saving then root element.
                SaveRootElement(packagePart, rootElement);
            }

            // Save the modified chart template for manual inspection.
            File.WriteAllBytes("ModifiedChartTemplate.crtx", stream.ToArray());

            // Assert that we have indeed changed the part.
            using (Package package = Package.Open("ModifiedChartTemplate.crtx"))
            {
                PackagePart packagePart = package.GetPart(new Uri("/chart/chart.xml", UriKind.Relative));
                XElement rootElement = LoadRootElement(packagePart);
                XElement lang = rootElement.Elements(C.lang).First();

                Assert.Equal("de-DE", (string) lang.Attribute(C.val));
            }
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        [SuppressMessage("ReSharper", "UnusedMember.Local")]
        private static class A
        {
            public static readonly XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            public static readonly XName lumMod = a + "lumMod";
            public static readonly XName lumOff = a + "lumOff";
            public static readonly XName noFill = a + "noFill";
            public static readonly XName schemeClr = a + "schemeClr";
            public static readonly XName solidFill = a + "solidFill";
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        private static class C
        {
            public static readonly XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";

            public static readonly XName chart = c + "chart";
            public static readonly XName chartSpace = c + "chartSpace";
            public static readonly XName lang = c + "lang";
            public static readonly XName legend = c + "legend";
            public static readonly XName plotArea = c + "plotArea";
            public static readonly XName title = c + "title";

            public static readonly XName val = "val";
        }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        [SuppressMessage("ReSharper", "UnusedMember.Local")]
        private static class CS
        {
            public static readonly XNamespace cs = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

            public static readonly XName colorStyle = cs + "colorStyle";
            public static readonly XName variation = cs + "variation";
        }
    }
}
