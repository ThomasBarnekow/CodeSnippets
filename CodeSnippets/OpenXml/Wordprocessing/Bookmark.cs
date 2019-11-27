//
// Bookmark.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using OpenXmlPowerTools;

namespace CodeSnippets.OpenXml.Wordprocessing
{
    /// <summary>
    /// Represents a corresponding pair of w:bookmarkStart and w:bookmarkEnd elements.
    /// </summary>
    public class Bookmark
    {
        private Bookmark(XElement root, string bookmarkName)
        {
            Root = root;

            BookmarkStart = new XElement(W.bookmarkStart,
                new XAttribute(W.id, -1),
                new XAttribute(W.name, bookmarkName));

            BookmarkEnd = new XElement(W.bookmarkEnd,
                new XAttribute(W.id, -1));
        }

        private Bookmark(XElement root, XElement bookmarkStart, XElement bookmarkEnd)
        {
            Root = root;
            BookmarkStart = bookmarkStart;
            BookmarkEnd = bookmarkEnd;
        }

        /// <summary>
        /// The root element containing both <see cref="BookmarkStart"/> and
        /// <see cref="BookmarkEnd"/>.
        /// </summary>
        public XElement Root { get; }

        /// <summary>
        /// The w:bookmarkStart element.
        /// </summary>
        public XElement BookmarkStart { get; }

        /// <summary>
        /// The w:bookmarkEnd element.
        /// </summary>
        public XElement BookmarkEnd { get; }

        /// <summary>
        /// Finds a pair of w:bookmarkStart and w:bookmarkEnd elements in the given
        /// <paramref name="root"/> element, where the w:name attribute value of the
        /// w:bookmarkStart element is equal to <paramref name="bookmarkName"/>.
        /// </summary>
        /// <param name="root">The root <see cref="XElement"/>.</param>
        /// <param name="bookmarkName">The bookmark name.</param>
        /// <returns>A new <see cref="Bookmark"/> instance representing the bookmark.</returns>
        public static Bookmark Find(XElement root, string bookmarkName)
        {
            XElement bookmarkStart = root
                .Descendants(W.bookmarkStart)
                .FirstOrDefault(e => (string) e.Attribute(W.name) == bookmarkName);

            string id = bookmarkStart?.Attribute(W.id)?.Value;
            if (id == null) return new Bookmark(root, bookmarkName);

            XElement bookmarkEnd = root
                .Descendants(W.bookmarkEnd)
                .FirstOrDefault(e => (string) e.Attribute(W.id) == id);

            return bookmarkEnd != null
                ? new Bookmark(root, bookmarkStart, bookmarkEnd)
                : new Bookmark(root, bookmarkName);
        }

        /// <summary>
        /// Gets all w:r elements between the bookmark's w:bookmarkStart and
        /// w:bookmarkEnd elements.
        /// </summary>
        /// <returns>A collection of w:r elements.</returns>
        public IEnumerable<XElement> GetRuns()
        {
            return Root
                .Descendants()
                .SkipWhile(d => d != BookmarkStart)
                .Skip(1)
                .TakeWhile(d => d != BookmarkEnd)
                .Where(d => d.Name == W.r);
        }

        /// <summary>
        /// Gets the concatenated inner text of all runs between the bookmark's
        /// w:bookmarkStart and w:bookmarkEnd elements, ignoring paragraph marks
        /// and page breaks.
        /// </summary>
        /// <remarks>
        /// The output of this method can be compared to the output of the
        /// <see cref="XElement.Value"/> property.
        /// </remarks>
        /// <returns>The concatenated inner text.</returns>
        public string GetValue()
        {
            return GetRuns().Select(UnicodeMapper.RunToString).StringConcatenate();
        }
    }
}
