//
// PackageRelationships.cs
//
// Copyright 2019 Thomas Barnekow
//
// Developer: Thomas Barnekow
// Email: thomas<at/>barnekow<dot/>info

using System.Collections;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;

namespace CodeSnippets.Windows.IO.Packaging
{
    public class PackageRelationships : IEnumerable<PackageRelationship>
    {
        private readonly Package _package;
        private readonly ISet<string> _relationshipKeys;
        private readonly ISet<PackageRelationship> _packageRelationships;

        public PackageRelationships(Package package)
        {
            _package = package;
            _relationshipKeys = new HashSet<string>();
            _packageRelationships = new HashSet<PackageRelationship>();

            _package.GetRelationships().ToList().ForEach(Enumerate);
        }

        private void Enumerate(PackageRelationship r)
        {
            string key = r.SourceUri + ":" + r.TargetUri;
            if (_relationshipKeys.Contains(key)) return;

            _relationshipKeys.Add(key);
            _packageRelationships.Add(r);

            _package
                .GetPart(PackUriHelper.ResolvePartUri(r.SourceUri, r.TargetUri))
                .GetRelationships()
                .ToList()
                .ForEach(Enumerate);
        }

        public IEnumerator<PackageRelationship> GetEnumerator()
        {
            return _packageRelationships.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
