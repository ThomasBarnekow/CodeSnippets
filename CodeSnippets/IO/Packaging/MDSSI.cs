//
// MDSSI.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Diagnostics.CodeAnalysis;
using System.Xml.Linq;

namespace CodeSnippets.IO.Packaging
{
    [SuppressMessage("ReSharper", "InconsistentNaming")]
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public class MDSSI
    {
        public static readonly XNamespace Namespace =
            "http://schemas.openxmlformats.org/package/2006/digital-signature";

        public static readonly XName SignatureTime = Namespace + "SignatureTime";
        public static readonly XName Format = Namespace + "Format";
        public static readonly XName Value = Namespace + "Value";
    }
}
