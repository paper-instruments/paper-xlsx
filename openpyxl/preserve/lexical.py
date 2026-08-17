# paper-xlsx: lexical XML patching

"""Patch modeled XML changes without reserializing unchanged syntax.

The preserve writer compares a settled model serialization from load time
with the current serialization.  When both have the same element shape, this
module transfers only the changed attribute values and leaf text into the
original bytes.  Attribute order, quote style, omitted defaults, namespace
spelling, whitespace, and unknown children therefore remain untouched.
"""

from xml.etree.ElementTree import ParseError, fromstring

from .crosspart import apply_edits, scan_small
from .xmlscan import ScanRefusal, _ATTR_RE, _scan_name_end


def _local(name):
    if isinstance(name, bytes):
        name = name.decode("latin-1")
    if name.startswith("{"):
        return name.rsplit("}", 1)[-1]
    return name.split(":", 1)[-1]


def _namespace_context(node, default_ns=None, prefixes=None):
    """Return the namespace context in scope for ``node`` itself."""
    current_default = node.attrs.get("xmlns", default_ns)
    current_prefixes = dict(prefixes or {})
    for key, value in node.attrs.items():
        if key.startswith("xmlns:"):
            current_prefixes[key.split(":", 1)[1]] = value
    return current_default, current_prefixes


def _expanded_scan_name(node, default_ns, prefixes):
    raw = node.name.decode("latin-1") \
        if isinstance(node.name, bytes) else node.name
    if ":" in raw:
        prefix, local = raw.split(":", 1)
        if prefix not in prefixes:
            raise ValueError("undeclared XML namespace prefix")
        return prefixes[prefix], local
    return default_ns, raw


def _expanded_element_name(name):
    if name.startswith("{"):
        namespace, local = name[1:].split("}", 1)
        return namespace, local
    return None, name


def _nodes_by_path(root):
    out = {}

    def visit(node, path, default_ns, prefixes):
        out[path] = node
        counts = {}
        for child in node.children:
            child_default, child_prefixes = _namespace_context(
                child, default_ns, prefixes)
            tag = _expanded_scan_name(
                child, child_default, child_prefixes)
            index = counts.get(tag, 0)
            counts[tag] = index + 1
            visit(child, path + ((tag, index),),
                  child_default, child_prefixes)

    default_ns, prefixes = _namespace_context(root)
    root_tag = _expanded_scan_name(root, default_ns, prefixes)
    visit(root, ((root_tag, 0),), default_ns, prefixes)
    return out


def _elements_by_path(data):
    root = fromstring(data)
    out = {}

    def visit(element, path):
        out[path] = element
        counts = {}
        for child in list(element):
            tag = _expanded_element_name(child.tag)
            index = counts.get(tag, 0)
            counts[tag] = index + 1
            visit(child, path + ((tag, index),))

    visit(root, ((_expanded_element_name(root.tag), 0),))
    return out


def _model_attrs(node):
    return {key: value for key, value in node.attrs.items()
            if key != "xmlns" and not key.startswith("xmlns:")}


def _raw_attrs(data, node):
    end = node.end if node.self_closing else node.content_start
    head = data[node.start:end]
    name_start = 1
    name_end = name_start + _scan_name_end(head[name_start:])
    found = {}
    for match in _ATTR_RE.finditer(head, name_end):
        key = match.group(1).decode("latin-1")
        value_group = 3 if match.group(3) is not None else 4
        found[key] = (match, value_group, head)
    return found


def _match_attr(attrs, modeled_key):
    if modeled_key in attrs:
        return modeled_key
    local = _local(modeled_key)
    matches = [key for key in attrs if _local(key) == local]
    if len(matches) == 1:
        return matches[0]
    return None


def _escape_attr(value, quote):
    value = str(value).replace("&", "&amp;").replace("<", "&lt;")
    if quote == b'"':
        value = value.replace('"', "&quot;")
    else:
        value = value.replace("'", "&apos;")
    return value.encode("utf-8")


def _escape_text(value):
    return (value.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").encode("utf-8"))


def patch_xml(original, baseline, current, expected_root):
    """Return ``original`` with only model-observed lexical deltas applied.

    ``None`` means the edit changed element shape or could not be mapped
    unambiguously.  Callers may then use a guarded structural fallback, but
    must not silently normalize XML that carries unowned content.
    """
    if not original or not baseline or not current:
        return None
    try:
        original_nodes = _nodes_by_path(
            scan_small(original, expected_root, max_depth=64))
        baseline_nodes = _nodes_by_path(
            scan_small(baseline, expected_root, max_depth=64))
        current_nodes = _nodes_by_path(
            scan_small(current, expected_root, max_depth=64))
        baseline_elements = _elements_by_path(baseline)
        current_elements = _elements_by_path(current)
    except (ParseError, ScanRefusal, ValueError, TypeError):
        return None

    model_paths = set(baseline_nodes)
    if model_paths != set(current_nodes):
        return None
    if not model_paths.issubset(original_nodes):
        return None

    edits = []
    for path in sorted(model_paths, key=len):
        before = baseline_nodes[path]
        after = current_nodes[path]
        target = original_nodes[path]
        before_attrs = _model_attrs(before)
        after_attrs = _model_attrs(after)
        if before_attrs != after_attrs:
            raw_attrs = _raw_attrs(original, target)
            for key in sorted(set(before_attrs) | set(after_attrs)):
                if before_attrs.get(key) == after_attrs.get(key):
                    continue
                raw_key = _match_attr(raw_attrs, key)
                if key not in after_attrs:
                    if raw_key is None:
                        return None
                    match = raw_attrs[raw_key][0]
                    start = target.start + match.start(0)
                    while start > target.start and original[start - 1:start] \
                            in (b" ", b"\t", b"\r", b"\n"):
                        start -= 1
                    edits.append((start, target.start + match.end(0), b""))
                    continue
                if raw_key is not None:
                    match, group, _head = raw_attrs[raw_key]
                    value_start, value_end = match.span(group)
                    quote = match.group(2)[:1]
                    edits.append((target.start + value_start,
                                  target.start + value_end,
                                  _escape_attr(after_attrs[key], quote)))
                    continue
                if key in before_attrs:
                    return None
                end = (target.end - 2 if target.self_closing
                       else target.content_start - 1)
                payload = (b" " + key.encode("latin-1") + b'="'
                           + _escape_attr(after_attrs[key], b'"') + b'"')
                edits.append((end, end, payload))

        before_element = baseline_elements[path]
        after_element = current_elements[path]
        if len(list(before_element)) or len(list(after_element)):
            continue
        before_text = before_element.text or ""
        after_text = after_element.text or ""
        if before_text == after_text:
            continue
        if target.self_closing or target.children:
            return None
        edits.append((target.content_start, target.content_end,
                      _escape_text(after_text)))

    return apply_edits(original, edits)
