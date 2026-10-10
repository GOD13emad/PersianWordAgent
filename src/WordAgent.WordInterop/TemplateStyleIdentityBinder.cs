using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace WordAgent.WordInterop;

/// <summary>Rebases document-local template style IDs without changing the source identity or numbering graph.</summary>
internal static class TemplateStyleIdentityBinder
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly HashSet<XName> ReferenceElements = new()
    {
        W + "basedOn", W + "next", W + "link", W + "pStyle", W + "rStyle", W + "tblStyle",
        W + "styleLink", W + "numStyleLink"
    };

    internal sealed record BoundTemplate(XDocument Document, XDocument Styles, IReadOnlyDictionary<string, string> IdMap);

    internal static BoundTemplate Bind(XDocument templateDocument, XDocument templateStyles, XDocument sourceStyles)
    {
        ArgumentNullException.ThrowIfNull(templateDocument);
        ArgumentNullException.ThrowIfNull(templateStyles);
        ArgumentNullException.ThrowIfNull(sourceStyles);
        var source = Index(sourceStyles);
        var template = Index(templateStyles);
        var names = new Dictionary<string, List<XElement>>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in source.Values)
        {
            var key = Identity(item);
            if (key is null) continue;
            if (!names.TryGetValue(key, out var entries)) names[key] = entries = new List<XElement>();
            entries.Add(item);
        }
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var occupied = new HashSet<string>(source.Keys, StringComparer.OrdinalIgnoreCase);
        occupied.UnionWith(template.Keys);
        var destinations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var templateNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in template.OrderBy(p => p.Key, StringComparer.Ordinal))
        {
            var original = pair.Key;
            var key = Identity(pair.Value);
            if (key is not null && !templateNames.Add(key))
                throw new InvalidDataException("TEMPLATE_STYLE_NAME_AMBIGUOUS: " + key);
            string target;
            if (key is not null && names.TryGetValue(key, out var matches))
            {
                if (matches.Count != 1) throw new InvalidDataException("SOURCE_STYLE_NAME_AMBIGUOUS: " + key);
                target = Id(matches[0]);
            }
            else if (!source.TryGetValue(original, out var existing) ||
                (key is null && Identity(existing) is null && Type(existing) == Type(pair.Value)))
            {
                target = original;
            }
            else
            {
                // Deterministic collision-free identity. Never overwrite an unrelated source style.
                var digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Type(pair.Value) + "\n" + key + "\n" + original)))[..16];
                target = "WA_T_" + digest;
                var suffix = 0;
                while (occupied.Contains(target) || destinations.Contains(target))
                    target = "WA_T_" + digest + "_" + (++suffix).ToString(CultureInfo.InvariantCulture);
            }
            if (!destinations.Add(target)) throw new InvalidDataException("TEMPLATE_STYLE_DESTINATION_AMBIGUOUS: " + target);
            map.Add(original, target);
        }
        // Simultaneous graph rewrite on clones also handles ID swaps; no input XDocument is mutated.
        var document = new XDocument(templateDocument);
        var styles = new XDocument(templateStyles);
        foreach (var style in styles.Root!.Elements(W + "style"))
        {
            var original = Id(style); var target = map[original];
            style.SetAttributeValue(W + "styleId", target);
            if (source.TryGetValue(target, out var originalSource) && Identity(style) is not null &&
                string.Equals(Identity(style), Identity(originalSource), StringComparison.OrdinalIgnoreCase))
            {
                // Field instructions resolve style names, so retain the source spelling exactly.
                style.Element(W + "name")!.SetAttributeValue(W + "val", (string?)originalSource.Element(W + "name")?.Attribute(W + "val"));
            }
        }
        foreach (var root in new[] { document, styles })
        foreach (var reference in root.Descendants().Where(e => ReferenceElements.Contains(e.Name)))
        {
            var old = (string?)reference.Attribute(W + "val");
            if (old is not null && map.TryGetValue(old, out var updated)) reference.SetAttributeValue(W + "val", updated);
        }
        _ = Index(styles);
        return new BoundTemplate(document, styles, map);
    }

    private static Dictionary<string, XElement> Index(XDocument styles)
    {
        if (styles.Root?.Name != W + "styles") throw new InvalidDataException("INVALID_STYLE_ROOT");
        var result = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        foreach (var style in styles.Root.Elements(W + "style"))
        {
            var id = Id(style);
            if (string.IsNullOrWhiteSpace(id) || !result.TryAdd(id, style))
                throw new InvalidDataException("STYLE_ID_MISSING_OR_DUPLICATE: " + id);
        }
        return result;
    }
    private static string Id(XElement style) => (string?)style.Attribute(W + "styleId") ?? "";
    private static string Type(XElement style) => ((string?)style.Attribute(W + "type") ?? "").Trim().ToLowerInvariant();
    private static string? Identity(XElement style)
    {
        var name = ((string?)style.Element(W + "name")?.Attribute(W + "val"))?.Trim().Normalize(NormalizationForm.FormC);
        return string.IsNullOrWhiteSpace(name) ? null : Type(style) + "\n" + name;
    }
}
