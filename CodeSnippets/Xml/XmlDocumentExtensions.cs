//
// XmlDocumentExtensions.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Xml;
using System.Xml.Linq;

namespace CodeSnippets.Xml
{
    public static class XmlDocumentExtensions
    {
        public static XmlElement CreateElement(this XmlDocument xmlDocument, XName name, string text = null)
        {
            XmlElement element = xmlDocument.CreateElement(name.LocalName, name.NamespaceName);

            if (text != null)
            {
                XmlText textNode = xmlDocument.CreateTextNode(text);
                element.AppendChild(textNode);
            }

            return element;
        }
    }
}
