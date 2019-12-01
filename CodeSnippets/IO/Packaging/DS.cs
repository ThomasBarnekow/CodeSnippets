//
// DS.cs
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
    public static class DS
    {
        public static readonly XNamespace Namespace = "http://www.w3.org/2000/09/xmldsig#";

        public static readonly XName Signature = Namespace + "Signature";
        public static readonly XName SignedInfo = Namespace + "SignedInfo";
        public static readonly XName CanonicalizationMethod = Namespace + "CanonicalizationMethod";
        public static readonly XName SignatureMethod = Namespace + "SignatureMethod";
        public static readonly XName Reference = Namespace + "Reference";
        public static readonly XName DigestMethod = Namespace + "DigestMethod";
        public static readonly XName DigestValue = Namespace + "DigestValue";
        public static readonly XName SignatureValue = Namespace + "SignatureValue";

        public static readonly XName Object = Namespace + "Object";
        public static readonly XName Manifest = Namespace + "Manifest";

        public static readonly XName SignatureProperties = Namespace + "SignatureProperties";
        public static readonly XName SignatureProperty = Namespace + "SignatureProperty";

        public static readonly XName Algorithm = "Algorithm";
        public static readonly XName Id = "Id";
        public static readonly XName Target = "Target";
    }
}
