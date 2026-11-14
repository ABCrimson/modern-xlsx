//! Shared XML helpers for every OOXML part parser and writer: attribute
//! decoding, xsd:boolean and xsd:double lexical rules, entity resolution,
//! whitespace preservation, namespace bookkeeping for verbatim fragments, and
//! the namespace constants.

use core::hint::cold_path;
use std::borrow::Cow;
use std::cell::Cell;
use std::collections::{BTreeMap, btree_map};
use std::str::FromStr;

use quick_xml::events::attributes::{Attribute, Attributes};
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Prefix, PrefixDeclaration, QName};
use quick_xml::{Reader, XmlVersion};

use crate::{ModernXlsxError, Result};

/// The SpreadsheetML main namespace (ECMA-376 Part 1, §18), the default
/// namespace of every `xl/*.xml` part this crate writes.
pub(crate) const SPREADSHEET_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/// The officeDocument relationships namespace (ECMA-376 Part 1, §22.8), bound
/// to the `r:` prefix wherever a part carries `r:id` references.
pub(crate) const REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// The DrawingML charts namespace (ECMA-376 Part 1, §21.2), bound to the `c:`
/// prefix of a chart part and the `uri` of a chart's `a:graphicData`.
pub(crate) const DML_CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

/// The DrawingML main namespace (ECMA-376 Part 1, §20.1), bound to the `a:`
/// prefix of chart and drawing parts.
pub(crate) const DML_MAIN_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

/// The DrawingML spreadsheet drawing namespace (ECMA-376 Part 1, §20.5), bound
/// to the `xdr:` prefix of an `xl/drawings/drawing{n}.xml` part.
pub(crate) const XDR_NS: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

/// XML 1.0 white space, the `S` production (§2.3): space, tab, LF and CR.
/// Wider Unicode spaces (U+00A0, U+3000, ...) are not XML white space.
pub(crate) const XML_S: [char; 4] = [' ', '\t', '\n', '\r'];

/// 2^53: below this magnitude every integral f64 is an exact i64, so the
/// integer digits are the shortest round-trip form.
const EXACT_INT_LIMIT: f64 = 9_007_199_254_740_992.0;

/// The decoded value of attribute `a`, per XML 1.0 attribute-value
/// normalization (§3.3.3): the five predefined entities and character
/// references are resolved, and literal tab/LF/CR become spaces. Borrows when
/// the raw value needs no change.
///
/// # Errors
///
/// `XmlParse` naming the attribute when the value holds an undeclared entity
/// or a malformed character reference.
pub(crate) fn attr_str<'a>(a: &Attribute<'a>) -> Result<Cow<'a, str>> {
    a.normalized_value(XmlVersion::Implicit1_0).map_err(|e| {
        cold_path();
        ModernXlsxError::XmlParse(format!("attribute @{}: {e}", a.key.as_ref()))
    })
}

/// Parse an xsd:boolean lexical value (XML Schema Part 2, §3.2.2): `1` and
/// `true` are true, `0` and `false` are false, case-sensitive, after collapsing
/// the surrounding XML white space. Anything else is `None`, so each caller
/// applies its own schema default with `xsd_bool(v).unwrap_or(DEFAULT)`.
#[inline]
pub(crate) fn xsd_bool(s: &str) -> Option<bool> {
    match s.trim_matches(XML_S) {
        "1" | "true" => Some(true),
        "0" | "false" => Some(false),
        _ => None,
    }
}

/// The attributes of `e`. A malformed attribute list (a duplicate name, a
/// missing quote, ...) is an `XmlParse` error naming the element's qualified
/// name, `<{name}> attributes: {err}`, never a silently dropped attribute.
pub(crate) fn attrs<'a>(e: &'a BytesStart<'_>) -> impl Iterator<Item = Result<Attribute<'a>>> {
    e.attributes().map(move |a| {
        a.map_err(|err| {
            cold_path();
            ModernXlsxError::XmlParse(format!("<{}> attributes: {err}", e.name().as_ref()))
        })
    })
}

/// The `val` attribute (matched by local name) of a property element such as
/// `<sz val="11"/>`.
///
/// # Errors
///
/// `XmlParse` when the attribute list is malformed.
pub(crate) fn val_attr<'a>(e: &'a BytesStart<'_>) -> Result<Option<Attribute<'a>>> {
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.local_name().as_ref() == "val" {
            return Ok(Some(attr));
        }
    }
    Ok(None)
}

/// The decoded string `val` of a property element.
///
/// # Errors
///
/// `XmlParse` when the attribute list is malformed or `val` holds a bad entity.
pub(crate) fn val_string(e: &BytesStart<'_>) -> Result<Option<String>> {
    val_attr(e)?
        .map(|a| attr_str(&a).map(Cow::into_owned))
        .transpose()
}

/// Read a DrawingML `CT_Boolean` element such as `<c:delete/>`: its `val`
/// attribute is optional with schema default `true` (dml-chart.xsd), so an
/// absent or invalid `val` reads as `true`.
///
/// # Errors
///
/// `XmlParse` when the attribute list is malformed (for example a duplicate
/// attribute) or `val` holds a bad entity.
pub(crate) fn ct_boolean(e: &BytesStart<'_>) -> Result<bool> {
    let mut val = true;
    for attr in attrs(e) {
        let attr = attr?;
        if attr.key.as_ref() == "val" {
            val = xsd_bool(&attr_str(&attr)?).unwrap_or(true);
        }
    }
    Ok(val)
}

/// Parse attribute `a` of element `elem` as `T`, after collapsing the
/// surrounding XML white space that numeric schema types (`whiteSpace =
/// collapse`) allow. A malformed value is an error, never a silent default.
///
/// # Errors
///
/// `XmlParse("<elem> @<key>: invalid value '<raw>'")` when the decoded value
/// does not parse, or the `attr_str` error when it does not decode.
pub(crate) fn parse_attr<T: FromStr>(elem: &str, a: &Attribute<'_>) -> Result<T> {
    let raw = attr_str(a)?;
    raw.trim_matches(XML_S).parse().map_err(|_| {
        cold_path();
        ModernXlsxError::XmlParse(format!("{elem} @{}: invalid value '{raw}'", a.key.as_ref()))
    })
}

/// Whether text written into `<t>` needs `xml:space="preserve"`: it starts or
/// ends with XML white space (§2.3 `S`: space, tab, LF, CR), which a consumer
/// may otherwise strip.
#[inline]
pub(crate) fn needs_preserve(s: &str) -> bool {
    s.starts_with(XML_S) || s.ends_with(XML_S)
}

/// A `(key, value)` pair for `BytesStart::push_attribute` whose value has the
/// characters illegal in XML 1.0 removed (see `sanitize_xml_text`); quick-xml
/// then escapes it. Borrows when the value is already clean.
#[inline]
pub(crate) fn xml_attr<'a>(k: &'a str, v: &'a str) -> (&'a str, Cow<'a, str>) {
    (k, sanitize_xml_text(v))
}

/// Append the text a quick-xml `Event::GeneralRef` stands for to `buf`.
///
/// Character references (`&#N;`, `&#xN;`) resolve under XML 1.0 §4.1: a sign,
/// an uppercase `X`, code point 0, a surrogate or a value above U+10FFFF is
/// rejected. The five predefined entities (§4.6) resolve to their character.
/// Any other named entity is undeclared in OOXML parts (they carry no DTD) and
/// is dropped, keeping the surrounding text.
///
/// # Errors
///
/// `XmlParse` for a malformed character reference.
pub(crate) fn push_general_ref(buf: &mut String, r: &BytesRef<'_>) -> Result<()> {
    match r.resolve_char_ref() {
        Ok(Some(c)) => buf.push(c),
        Ok(None) => {
            if let Some(s) = quick_xml::escape::resolve_xml_entity(r) {
                buf.push_str(s);
            }
        }
        Err(e) => {
            cold_path();
            return Err(ModernXlsxError::XmlParse(format!(
                "character reference &{};: {e}",
                &**r
            )));
        }
    }
    Ok(())
}

/// The namespace declarations (`xmlns` and `xmlns:p`, XML Namespaces §3) on a
/// part's root start tag, in document order, as `(prefix, uri)`; the default
/// namespace has the empty prefix.
///
/// # Errors
///
/// `XmlParse` when the attribute list is malformed or a URI holds a bad entity.
pub(crate) fn root_namespaces(e: &BytesStart<'_>) -> Result<Vec<(String, String)>> {
    let mut out = Vec::new();
    for attr in attrs(e) {
        let attr = attr?;
        let prefix = match attr.key.as_namespace_binding() {
            Some(PrefixDeclaration::Default) => "",
            Some(PrefixDeclaration::Named(p)) => p,
            None => continue,
        };
        out.push((prefix.to_owned(), attr_str(&attr)?.into_owned()));
    }
    Ok(out)
}

/// The namespace declarations of a part's root start tag
/// ([`root_namespaces`]) with an index by prefix, built once per root of the
/// part so that each verbatim fragment of the part looks up only the prefixes
/// it uses ([`self_contained_fragment`]). The index holds, for the first
/// binding of each prefix, the declaration a fragment receives, rendered and
/// escaped once, so inserting a declaration copies that text and nothing
/// more.
#[derive(Debug, Default)]
pub(crate) struct RootNamespaces {
    /// `(prefix, uri)` in document order.
    list: Vec<(String, String)>,
    /// One entry per bound prefix, in the document order of its first
    /// binding: the index into `list` of that binding, and the end in
    /// `rendered` of its declaration, which starts where the previous entry's
    /// ends.
    entries: Vec<(usize, usize)>,
    /// The declaration of each entry, in entry order: ` xmlns:p="uri"` with
    /// the URI escaped by `quick_xml::escape::escape`, and ` xmlns="uri"` for
    /// the default namespace.
    rendered: String,
    /// The entries sorted by prefix, for a binary search. Every entry has a
    /// prefix of its own, so a search finds at most one.
    by_prefix: Vec<usize>,
    /// For each entry, the number of the last [`inherited_declarations`] call
    /// that inserted its declaration, so a fragment inserts each declaration
    /// once without a set of its own.
    inserted_in: Vec<Cell<u32>>,
    /// The number of the current [`inherited_declarations`] call.
    call: Cell<u32>,
}

impl RootNamespaces {
    /// Index `list` (in document order) by prefix and render the declaration
    /// of each prefix's first binding.
    pub(crate) fn new(list: Vec<(String, String)>) -> Self {
        // The first binding of each prefix: a stable sort keeps the bindings
        // of a prefix in document order, and the dedup keeps the first.
        let mut first: Vec<usize> = (0..list.len()).collect();
        first.sort_by(|&a, &b| list[a].0.cmp(&list[b].0));
        first.dedup_by(|later, kept| list[*later].0 == list[*kept].0);
        first.sort_unstable();
        let mut rendered = String::with_capacity(
            first
                .iter()
                .map(|&i| list[i].0.len() + list[i].1.len() + 10)
                .sum(),
        );
        let mut entries = Vec::with_capacity(first.len());
        for i in first {
            let (prefix, uri) = &list[i];
            if prefix.is_empty() {
                rendered.push_str(" xmlns=\"");
            } else {
                rendered.push_str(" xmlns:");
                rendered.push_str(prefix);
                rendered.push_str("=\"");
            }
            rendered.push_str(&quick_xml::escape::escape(uri));
            rendered.push('"');
            entries.push((i, rendered.len()));
        }
        let mut by_prefix: Vec<usize> = (0..entries.len()).collect();
        by_prefix.sort_unstable_by(|&a, &b| list[entries[a].0].0.cmp(&list[entries[b].0].0));
        Self {
            inserted_in: vec![Cell::new(0); entries.len()],
            list,
            entries,
            rendered,
            by_prefix,
            call: Cell::new(0),
        }
    }

    /// The entry of `prefix`: a binary search of the entries by prefix.
    fn entry(&self, prefix: &str) -> Option<usize> {
        self.by_prefix
            .binary_search_by(|&at| self.list[self.entries[at].0].0.as_str().cmp(prefix))
            .ok()
            .map(|i| self.by_prefix[i])
    }

    /// The rendered declaration of entry `at`.
    fn rendered_at(&self, at: usize) -> &str {
        let start = at.checked_sub(1).map_or(0, |b| self.entries[b].1);
        &self.rendered[start..self.entries[at].1]
    }

    /// The declarations in document order.
    #[cfg(test)]
    fn list(&self) -> &[(String, String)] {
        &self.list
    }

    /// The prefixes, other than the empty one, whose first binding is `uri`,
    /// in sorted order and each once. On a root start tag, whose attribute
    /// checks refuse a prefix declared twice, that is every non-empty prefix
    /// bound to `uri`.
    pub(crate) fn prefixes_bound_to<'s>(&'s self, uri: &'s str) -> impl Iterator<Item = &'s str> {
        self.by_prefix.iter().filter_map(move |&at| {
            let (prefix, bound) = &self.list[self.entries[at].0];
            (!prefix.is_empty() && bound == uri).then_some(prefix.as_str())
        })
    }

    /// The rendered declaration of the first binding of `prefix` (the empty
    /// prefix is the default namespace), if any.
    #[cfg(test)]
    fn declaration(&self, prefix: &str) -> Option<&str> {
        self.entry(prefix).map(|at| self.rendered_at(at))
    }

    /// Start an [`inherited_declarations`] call, after which each entry's
    /// [`Self::first_insert`] is true once.
    fn begin_call(&self) {
        let call = self.call.get().wrapping_add(1);
        if call == 0 {
            // After 2^32 - 1 calls, forget them all.
            self.inserted_in.iter().for_each(|c| c.set(0));
            self.call.set(1);
        } else {
            self.call.set(call);
        }
    }

    /// Whether entry `at` is inserted for the first time in this call.
    fn first_insert(&self, at: usize) -> bool {
        self.inserted_in[at].replace(self.call.get()) != self.call.get()
    }

    /// The bytes these declarations and their index keep: each pair and its
    /// strings, each entry with its place in the sorted order and its call
    /// number, and the rendered declarations.
    pub(crate) fn stored_bytes(&self) -> usize {
        self.list
            .iter()
            .map(|(p, u)| size_of::<(String, String)>() + p.capacity() + u.capacity())
            .sum::<usize>()
            + self.entries.capacity() * size_of::<(usize, usize)>()
            + self.by_prefix.capacity() * size_of::<usize>()
            + self.inserted_in.capacity() * size_of::<Cell<u32>>()
            + self.rendered.capacity()
    }
}

impl FromIterator<(String, String)> for RootNamespaces {
    fn from_iter<I: IntoIterator<Item = (String, String)>>(iter: I) -> Self {
        Self::new(iter.into_iter().collect())
    }
}

/// Make a verbatim XML fragment (an `<ext>` span kept for round-trip) carry
/// the namespace declarations it relies on, so it stays well-formed under the
/// Namespaces in XML rule that every prefix in use be declared in scope.
///
/// Each element or attribute prefix used without an in-scope declaration
/// inside the fragment gets `xmlns:p="uri"` from `inherited` (the
/// [`RootNamespaces`] of the source part), inserted into the fragment's first
/// start tag right after its name, in first-use order. The reserved `xml`
/// prefix is never declared. Unprefixed names are left to the host part's
/// default namespace, and a name that begins with a colon (the empty prefix)
/// names no declaration and is left as is. A prefix absent from `inherited`
/// cannot be resolved and is left as is; a prefix bound twice there takes its
/// first binding. Returns a copy of the input when nothing is missing or when
/// the fragment is not well-formed. See [`inherited_declarations`] for the
/// cost.
pub(crate) fn self_contained_fragment(fragment: &str, inherited: &RootNamespaces) -> String {
    match inherited_declarations(fragment, &[inherited]) {
        Some((insert_at, decls)) => with_declarations(fragment, insert_at, &decls),
        None => fragment.to_owned(),
    }
}

/// The declarations [`self_contained_fragment`] inserts into `fragment`, and
/// the byte offset they go at, with the bindings of `scopes` searched in
/// order (an inner element's declarations before the part root's, so the
/// first scope that binds a prefix gives its declaration). `None` when
/// nothing is inserted, as when the fragment is not well-formed.
///
/// A use of a prefix inside the fragment is looked up in the declarations of
/// its own element, sorted, and then in those of its open ancestors, counted
/// by prefix in an ordered map keyed by slices of the fragment (a B-tree,
/// whose cost does not depend on a hasher, so no choice of prefixes makes its
/// lookups collide); each use still undeclared is a binary search of each
/// scope's index until one binds it, whose entry gives the declaration text
/// its index rendered and whether this call has inserted it already. So a
/// fragment costs its own length plus searches, logarithmic in the
/// declarations the fragment and the scopes hold, per prefix use and
/// declaration, and nothing is allocated or formatted per inserted
/// declaration.
pub(crate) fn inherited_declarations<'a>(
    fragment: &'a str,
    scopes: &[&'a RootNamespaces],
) -> Option<(usize, Vec<&'a str>)> {
    for ns in scopes {
        ns.begin_call();
    }
    let mut decls = Vec::new();
    let insert_at = undeclared_prefixes(fragment, |p| {
        for ns in scopes {
            if let Some(at) = ns.entry(p) {
                if ns.first_insert(at) {
                    decls.push(ns.rendered_at(at));
                }
                return;
            }
        }
    })?;
    (!decls.is_empty()).then_some((insert_at, decls))
}

/// `fragment` with `decls` inserted at byte `insert_at`, in one allocation of
/// the exact length.
pub(crate) fn with_declarations(fragment: &str, insert_at: usize, decls: &[&str]) -> String {
    let added: usize = decls.iter().map(|d| d.len()).sum();
    let mut out = String::with_capacity(fragment.len() + added);
    out.push_str(&fragment[..insert_at]);
    for decl in decls {
        out.push_str(decl);
    }
    out.push_str(&fragment[insert_at..]);
    out
}

/// Scan `fragment`, call `on_use` with each use of a prefix outside the
/// scope of any declaration inside the fragment (other than the reserved
/// `xml` and the empty prefix of a name that begins with a colon, which
/// names no declaration), in document order, as a slice of the fragment, and
/// return the byte offset just after the first start tag's name. `None` when
/// the fragment has no element or is not well-formed.
fn undeclared_prefixes<'a>(fragment: &'a str, mut on_use: impl FnMut(&'a str)) -> Option<usize> {
    /// Take back the declarations of one element from `in_scope`, one search
    /// per declaration.
    fn release<'a>(in_scope: &mut BTreeMap<&'a str, usize>, declared: &[&'a str]) {
        for &p in declared {
            if let btree_map::Entry::Occupied(mut n) = in_scope.entry(p) {
                *n.get_mut() -= 1;
                if *n.get() == 0 {
                    n.remove();
                }
            }
        }
    }

    let mut reader = Reader::from_str(fragment);
    // The prefixes each open element declares, innermost last.
    let mut scopes: Vec<Vec<&str>> = Vec::new();
    // How many declarations of the first `counted` open elements bind each
    // prefix, so a use the element does not declare itself is one lookup.
    // The open elements are counted only when a use needs the map (a prefix
    // the using element does not declare itself), so an element with no such
    // use in its content never enters it.
    let mut in_scope: BTreeMap<&str, usize> = BTreeMap::new();
    let mut counted = 0;
    // The prefixes the current element uses, reused across elements.
    let mut used: Vec<&str> = Vec::new();
    let mut insert_at = None;
    loop {
        let (e, opens_scope) = match reader.read_event().ok()? {
            Event::Start(e) => (e, true),
            Event::Empty(e) => (e, false),
            Event::End(_) => {
                if let Some(declared) = scopes.pop()
                    && scopes.len() < counted
                {
                    counted = scopes.len();
                    release(&mut in_scope, &declared);
                }
                continue;
            }
            Event::Eof => break,
            _ => continue,
        };
        // The tag ends at the reader position, after `>` for a start tag and
        // `/>` for an empty one, so its name and attributes are the slice of
        // `fragment` before that, which the names below borrow.
        let end = usize::try_from(reader.buffer_position()).ok()?;
        let content_end = end.checked_sub(if opens_scope { 1 } else { 2 })?;
        let content_start = content_end.checked_sub(e.len())?;
        let tag = fragment.get(content_start..content_end)?;
        debug_assert_eq!(tag, &*e);
        let name_len = e.name().as_ref().len();
        let name = QName(tag.get(..name_len)?);
        insert_at.get_or_insert(content_start + name_len);
        let mut declared = Vec::new();
        used.clear();
        used.extend(name.prefix().map(Prefix::into_inner));
        for attr in Attributes::new(tag, name_len) {
            let attr = attr.ok()?;
            match attr.key.as_namespace_binding() {
                Some(PrefixDeclaration::Named(p)) => declared.push(p),
                Some(PrefixDeclaration::Default) => {}
                None => used.extend(attr.key.prefix().map(Prefix::into_inner)),
            }
        }
        // The element's own declarations, sorted for a binary search (the
        // attribute checks refuse a prefix declared twice on one element).
        declared.sort_unstable();
        for &p in &used {
            if p.is_empty() || p == "xml" || declared.binary_search(&p).is_ok() {
                continue;
            }
            for open in &scopes[counted..] {
                for &q in open {
                    *in_scope.entry(q).or_insert(0) += 1;
                }
            }
            counted = scopes.len();
            if in_scope.is_empty() || !in_scope.contains_key(p) {
                on_use(p);
            }
        }
        // Only an element with content passes its declarations on.
        if opens_scope {
            scopes.push(declared);
        }
    }
    insert_at
}

/// Reusable scratch space for `fmt_xsd_double`: a zmij float buffer and an
/// itoa integer buffer, both stack-allocated.
#[derive(Default)]
pub(crate) struct XsdBuf {
    float: zmij::Buffer,
    int: itoa::Buffer,
}

impl XsdBuf {
    pub(crate) fn new() -> Self {
        Self::default()
    }
}

/// Format `v` as an xsd:double lexical value (XML Schema Part 2, §3.2.5) that
/// parses back to the same bits. Integral values below 2^53 in magnitude are
/// written as plain integers (`15`, not `15.0`), negative zero as `-0`, and
/// everything else in zmij's shortest round-trip form (`8.43`, `1e+16`).
///
/// # Errors
///
/// `XmlWrite("<ctx>: non-finite xsd:double")` for NaN and the infinities,
/// which OOXML's numeric attributes cannot carry.
pub(crate) fn fmt_xsd_double<'b>(v: f64, buf: &'b mut XsdBuf, ctx: &str) -> Result<&'b str> {
    if !v.is_finite() {
        cold_path();
        return Err(ModernXlsxError::XmlWrite(format!(
            "{ctx}: non-finite xsd:double"
        )));
    }
    if v == 0.0 && v.is_sign_negative() {
        return Ok("-0");
    }
    if v.fract() == 0.0 && v.abs() < EXACT_INT_LIMIT {
        return Ok(buf.int.format(v as i64));
    }
    Ok(buf.float.format_finite(v))
}

/// Returns `true` if `c` is forbidden in an XML 1.0 document: the C0 control
/// characters other than tab/LF/CR, plus the permanently-unassigned U+FFFE and
/// U+FFFF noncharacters. These fall outside the XML 1.0 `Char` production
/// (§2.2), so a document that contains them is not well-formed and a conforming
/// parser rejects it.
#[inline]
pub(crate) fn is_illegal_xml_char(c: char) -> bool {
    let u = c as u32;
    (u < 0x20 && c != '\t' && c != '\n' && c != '\r') || u == 0xFFFE || u == 0xFFFF
}

/// Drop characters illegal in XML 1.0 from user-supplied text so the writer can
/// never emit a corrupt document. Returns the input borrowed (zero allocation)
/// when it is already clean, which is the overwhelmingly common case.
pub(crate) fn sanitize_xml_text(s: &str) -> Cow<'_, str> {
    if s.chars().any(is_illegal_xml_char) {
        Cow::Owned(s.chars().filter(|c| !is_illegal_xml_char(*c)).collect())
    } else {
        Cow::Borrowed(s)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::errors::ModernXlsxError;
    use proptest::prelude::*;
    use quick_xml::Reader;
    use quick_xml::events::attributes::Attribute;
    use quick_xml::events::{BytesStart, Event};
    use std::collections::{HashMap, HashSet};
    use std::fmt::Write as _;

    /// The first start/empty tag of `xml`, owned.
    fn first_tag(xml: &str) -> BytesStart<'static> {
        let mut r = Reader::from_str(xml);
        loop {
            match r.read_event().unwrap() {
                Event::Start(e) | Event::Empty(e) => return e.into_owned(),
                Event::Eof => panic!("no start tag in {xml}"),
                _ => {}
            }
        }
    }

    /// The attribute `key` of `e`.
    fn find_attr<'a>(e: &'a BytesStart<'_>, key: &str) -> Attribute<'a> {
        e.attributes()
            .map(|a| a.expect("well-formed attribute"))
            .find(|a| a.key.as_ref() == key)
            .unwrap()
    }

    /// Text of `<a>…</a>` built with `push_general_ref` for every reference.
    fn resolve_refs(xml: &str) -> Result<String> {
        let mut r = Reader::from_str(xml);
        let mut out = String::new();
        loop {
            match r.read_event().unwrap() {
                Event::Text(t) => out.push_str(t.as_ref()),
                Event::GeneralRef(e) => push_general_ref(&mut out, &e)?,
                Event::Eof => return Ok(out),
                _ => {}
            }
        }
    }

    #[test]
    fn xsd_bool_accepts_exact_lexical_forms() {
        for s in ["1", "true", " true "] {
            assert_eq!(xsd_bool(s), Some(true), "{s:?}");
        }
        for s in ["0", "false", "\t0\r\n"] {
            assert_eq!(xsd_bool(s), Some(false), "{s:?}");
        }
        for s in ["TRUE", "True", "yes", "", " ", "\u{a0}1"] {
            assert_eq!(xsd_bool(s), None, "{s:?}");
        }
    }

    #[test]
    fn ct_boolean_defaults_to_true() {
        assert!(ct_boolean(&first_tag("<c:delete/>")).unwrap());
        assert!(!ct_boolean(&first_tag(r#"<c:delete val="0"/>"#)).unwrap());
        assert!(!ct_boolean(&first_tag(r#"<c:delete val="false"/>"#)).unwrap());
        assert!(ct_boolean(&first_tag(r#"<c:delete val="1"/>"#)).unwrap());
        assert!(ct_boolean(&first_tag(r#"<c:delete val="bogus"/>"#)).unwrap());
    }

    #[test]
    fn ct_boolean_propagates_attribute_errors() {
        let e = ct_boolean(&first_tag(r#"<c:delete val="A &bad; B"/>"#)).unwrap_err();
        assert!(matches!(e, ModernXlsxError::XmlParse(_)), "{e:?}");
        let e = ct_boolean(&first_tag(r#"<c:delete val="1" val="0"/>"#)).unwrap_err();
        assert!(matches!(e, ModernXlsxError::XmlParse(_)), "{e:?}");
    }

    #[test]
    fn needs_preserve_uses_the_xml_s_set() {
        for s in ["x\r", "\rx", " a", "a\t", "\na", "a "] {
            assert!(needs_preserve(s), "{s:?}");
        }
        for s in ["a b", "", "a", "\u{a0}a", "a\u{3000}"] {
            assert!(!needs_preserve(s), "{s:?}");
        }
    }

    #[test]
    fn xml_attr_sanitizes_and_borrows_when_clean() {
        let (k, v) = xml_attr("display", "a\u{1}b");
        assert_eq!(k, "display");
        assert_eq!(v, "ab");
        assert!(matches!(
            xml_attr("display", "clean").1,
            Cow::Borrowed("clean")
        ));
        // quick-xml escapes the sanitized value on the way out.
        let mut e = BytesStart::new("hyperlink");
        e.push_attribute(xml_attr("display", "a\u{1}<b"));
        assert_eq!(e.attributes_raw(), r#" display="a&lt;b""#);
    }

    #[test]
    fn push_general_ref_resolves_valid_references() {
        assert_eq!(resolve_refs("<a>&#x41;</a>").unwrap(), "A");
        assert_eq!(resolve_refs("<a>&#65;</a>").unwrap(), "A");
        assert_eq!(resolve_refs("<a>&amp;</a>").unwrap(), "&");
        assert_eq!(
            resolve_refs("<a>x&lt;&gt;&quot;&apos;y</a>").unwrap(),
            "x<>\"'y"
        );
        assert_eq!(resolve_refs("<a>&#x1F600;</a>").unwrap(), "😀");
    }

    #[test]
    fn push_general_ref_drops_unknown_named_entities() {
        assert_eq!(resolve_refs("<a>&unknown;</a>").unwrap(), "");
        assert_eq!(resolve_refs("<a>a&nbsp;b</a>").unwrap(), "ab");
    }

    #[test]
    fn push_general_ref_rejects_malformed_char_refs() {
        for xml in [
            "<a>&#x+41;</a>",
            "<a>&#0;</a>",
            "<a>&#X41;</a>",
            "<a>&#xD800;</a>",
        ] {
            let e = resolve_refs(xml).unwrap_err();
            assert!(matches!(e, ModernXlsxError::XmlParse(_)), "{xml}: {e:?}");
            assert_eq!(e.code(), "XML_PARSE");
        }
    }

    #[test]
    fn attr_str_unescapes_once() {
        assert_eq!(
            attr_str(&find_attr(
                &first_tag(r#"<h display="A &amp; B"/>"#),
                "display"
            ))
            .unwrap(),
            "A & B"
        );
        assert_eq!(
            attr_str(&find_attr(
                &first_tag(r#"<h tooltip="x&#10;y"/>"#),
                "tooltip"
            ))
            .unwrap(),
            "x\ny"
        );
        assert!(matches!(
            attr_str(&find_attr(&first_tag(r#"<h display="plain"/>"#), "display")).unwrap(),
            Cow::Borrowed("plain")
        ));
    }

    #[test]
    fn attr_str_rejects_unknown_entities_naming_the_key() {
        let e = attr_str(&find_attr(&first_tag(r#"<s name="A &bad; B"/>"#), "name")).unwrap_err();
        assert!(matches!(e, ModernXlsxError::XmlParse(_)), "{e:?}");
        assert!(e.to_string().contains("name"), "{e}");
    }

    #[test]
    fn parse_attr_trims_and_reports_bad_values() {
        let e = first_tag(r#"<numFmt numFmtId=" 14 "/>"#);
        let a = find_attr(&e, "numFmtId");
        assert_eq!(parse_attr::<u32>("numFmt", &a).unwrap(), 14);
        let e = first_tag(r#"<numFmt numFmtId="&#9;7&#10;"/>"#);
        let a = find_attr(&e, "numFmtId");
        assert_eq!(parse_attr::<u32>("numFmt", &a).unwrap(), 7);
        let e = first_tag(r#"<numFmt numFmtId="x"/>"#);
        let a = find_attr(&e, "numFmtId");
        let err = parse_attr::<u32>("numFmt", &a).unwrap_err();
        assert!(matches!(err, ModernXlsxError::XmlParse(_)), "{err:?}");
        let msg = err.to_string();
        assert!(msg.contains("numFmt @numFmtId"), "{msg}");
        assert!(msg.contains("'x'"), "{msg}");
    }

    #[test]
    fn root_namespaces_collects_declarations() {
        let root = first_tag(
            r#"<worksheet xmlns="M" xmlns:r="R" xmlns:x14ac="X&amp;Y" mc:Ignorable="x14ac"/>"#,
        );
        assert_eq!(
            root_namespaces(&root).unwrap(),
            vec![
                (String::new(), "M".to_owned()),
                ("r".to_owned(), "R".to_owned()),
                ("x14ac".to_owned(), "X&Y".to_owned()),
            ]
        );
    }

    #[test]
    fn self_contained_fragment_inserts_missing_prefixes() {
        let inherited: RootNamespaces = [("x14".to_owned(), "U".to_owned())].into_iter().collect();
        assert_eq!(
            self_contained_fragment(r#"<ext uri="u"><x14:cf/></ext>"#, &inherited),
            r#"<ext xmlns:x14="U" uri="u"><x14:cf/></ext>"#
        );
        let declared = r#"<ext uri="u" xmlns:x14="U"><x14:cf/></ext>"#;
        assert_eq!(self_contained_fragment(declared, &inherited), declared);
    }

    #[test]
    fn self_contained_fragment_handles_attributes_scope_and_reserved_prefixes() {
        let inherited = [
            (String::new(), "M".to_owned()),
            ("r".to_owned(), "R".to_owned()),
            ("xm".to_owned(), "XM&\"".to_owned()),
        ]
        .into_iter()
        .collect::<RootNamespaces>();
        // Attribute prefix, element prefix, reserved xml prefix, escaped URI, and a
        // declaration on a sibling that does not cover the other sibling.
        let frag = r#"<ext uri="u"><a xmlns:xm="Z"><xm:f/></a><b r:id="rId1" xml:space="preserve"><xm:f/></b></ext>"#;
        assert_eq!(
            self_contained_fragment(frag, &inherited),
            r#"<ext xmlns:r="R" xmlns:xm="XM&amp;&quot;" uri="u"><a xmlns:xm="Z"><xm:f/></a><b r:id="rId1" xml:space="preserve"><xm:f/></b></ext>"#
        );
        // Unprefixed names are left to the host part's default namespace.
        assert_eq!(self_contained_fragment("<ext/>", &inherited), "<ext/>");
        // A prefix with no inherited binding cannot be fixed and is left alone.
        assert_eq!(
            self_contained_fragment("<ext><q:z/></ext>", &inherited),
            "<ext><q:z/></ext>"
        );
        // Leading whitespace before the first tag is kept.
        assert_eq!(
            self_contained_fragment("\n <r:x/>", &inherited),
            "\n <r:x xmlns:r=\"R\"/>"
        );
    }

    /// A name that begins with a colon (`<:a/>`, `:k="1"`) has the empty
    /// prefix for quick-xml. It does not use the default namespace, so no
    /// default declaration is inserted for it, whether or not the fragment
    /// declares a default namespace of its own: inserting one put a second
    /// `xmlns` on an `<ext>` that has one, which is not well-formed.
    #[test]
    fn colon_leading_names_insert_no_default_namespace() {
        let inherited: RootNamespaces = [
            (String::new(), "M".to_owned()),
            ("r".to_owned(), "R".to_owned()),
        ]
        .into_iter()
        .collect();
        for frag in [
            r#"<ext xmlns="urn:x"><:a :k="1"/></ext>"#,
            "<ext><:a/></ext>",
            r#"<:ext :k="1"/>"#,
        ] {
            assert_eq!(self_contained_fragment(frag, &inherited), frag);
            let mut uses = Vec::new();
            assert!(undeclared_prefixes(frag, |p| uses.push(p)).is_some());
            assert!(uses.is_empty(), "{frag}: {uses:?}");
        }
        // A prefixed name beside it is still declared.
        assert_eq!(
            self_contained_fragment(r#"<ext><:a r:id="1"/></ext>"#, &inherited),
            r#"<ext xmlns:r="R"><:a r:id="1"/></ext>"#
        );
    }

    /// Every declaration [`inherited_declarations`] returns lies inside the
    /// rendered index of one of its scopes: a declaration formatted, escaped
    /// or copied per insert would lie in a new allocation, so this fails in
    /// every profile and on every machine, whatever its speed.
    #[test]
    fn inherited_declarations_borrow_the_rendered_index() {
        let root: RootNamespaces = (0..200)
            .map(|i| (format!("p{i}"), format!("u&{i}")))
            .chain([(String::new(), "M".to_owned())])
            .collect();
        let ext_lst: RootNamespaces = (0..50)
            .map(|i| (format!("p{}", i * 3), format!("e\"{i}")))
            .collect();
        let mut frag = String::from("<ext");
        for i in 0..100 {
            write!(frag, r#" p{i}:a="1""#).unwrap();
        }
        frag.push_str("><e/>");
        for i in (0..200).rev() {
            write!(frag, "<p{i}:x/>").unwrap();
        }
        frag.push_str("</ext>");
        let within = |d: &str, ns: &RootNamespaces| {
            let index = ns.rendered.as_bytes().as_ptr_range();
            let decl = d.as_bytes().as_ptr_range();
            index.contains(&decl.start) && decl.end <= index.end
        };
        for scopes in [&[&root][..], &[&ext_lst, &root][..]] {
            let (_, decls) = inherited_declarations(&frag, scopes).expect("declarations inserted");
            assert_eq!(decls.len(), 200);
            for d in &decls {
                assert!(
                    scopes.iter().any(|ns| within(d, ns)),
                    "{d} is not a slice of a rendered index"
                );
            }
        }
    }

    #[test]
    fn self_contained_fragment_is_linear_in_prefixes_and_declarations() {
        const N: usize = 50_000;
        // The first binding of a prefix wins; p0 is bound twice.
        let mut inherited: Vec<(String, String)> =
            (0..N).map(|i| (format!("p{i}"), format!("u{i}"))).collect();
        inherited.push(("p0".to_owned(), "later".to_owned()));
        let inherited = RootNamespaces::new(inherited);
        // Used in reverse order, so each prefix sits at the far end of `inherited`.
        let mut frag = String::from("<ext>");
        for i in (0..N).rev() {
            write!(frag, "<p{i}:x/>").unwrap();
        }
        frag.push_str("</ext>");
        let start = std::time::Instant::now();
        let out = self_contained_fragment(&frag, &inherited);
        let elapsed = start.elapsed();
        assert!(
            out.starts_with(r#"<ext xmlns:p49999="u49999" xmlns:p49998="u49998""#),
            "{}",
            &out[..80]
        );
        assert_eq!(out.matches(" xmlns:").count(), N);
        assert!(out.contains(r#" xmlns:p0="u0">"#));
        assert!(
            elapsed < std::time::Duration::from_secs(5),
            "{N} prefixes took {elapsed:?}"
        );
    }

    /// A reference for the output of `self_contained_fragment`: it builds a
    /// map of every inherited declaration on each call and formats each
    /// inserted declaration.
    fn reference_fragment(fragment: &str, inherited: &[(String, String)]) -> String {
        let Some((insert_at, missing)) = reference_undeclared(fragment) else {
            return fragment.to_owned();
        };
        let mut by_prefix: HashMap<&str, &str> = HashMap::with_capacity(inherited.len());
        for (p, uri) in inherited {
            by_prefix.entry(p.as_str()).or_insert(uri.as_str());
        }
        let mut decls = String::new();
        for p in &missing {
            if let Some(&uri) = by_prefix.get(p.as_str()) {
                let _ = write!(decls, " xmlns:{p}=\"{}\"", quick_xml::escape::escape(uri));
            }
        }
        if decls.is_empty() {
            return fragment.to_owned();
        }
        let mut out = String::with_capacity(fragment.len() + decls.len());
        out.push_str(&fragment[..insert_at]);
        out.push_str(&decls);
        out.push_str(&fragment[insert_at..]);
        out
    }

    /// A reference for `undeclared_prefixes`: each use scans every
    /// declaration in scope.
    fn reference_undeclared(fragment: &str) -> Option<(usize, Vec<String>)> {
        let mut reader = Reader::from_str(fragment);
        let mut scopes: Vec<Vec<String>> = Vec::new();
        let mut missing: Vec<String> = Vec::new();
        let mut seen: HashSet<String> = HashSet::new();
        let mut insert_at = None;
        loop {
            let (e, opens_scope) = match reader.read_event().ok()? {
                Event::Start(e) => (e, true),
                Event::Empty(e) => (e, false),
                Event::End(_) => {
                    scopes.pop();
                    continue;
                }
                Event::Eof => break,
                _ => continue,
            };
            let name = e.name();
            if insert_at.is_none() {
                let end = usize::try_from(reader.buffer_position()).ok()?;
                insert_at = Some(fragment.get(..end)?.rfind('<')? + 1 + name.as_ref().len());
            }
            let mut declared = Vec::new();
            let mut used: Vec<&str> = name.prefix().map(Prefix::into_inner).into_iter().collect();
            for attr in e.attributes() {
                let attr = attr.ok()?;
                match attr.key.as_namespace_binding() {
                    Some(PrefixDeclaration::Named(p)) => declared.push(p.to_owned()),
                    Some(PrefixDeclaration::Default) => {}
                    None => used.extend(attr.key.prefix().map(Prefix::into_inner)),
                }
            }
            for p in used {
                let in_scope = declared
                    .iter()
                    .chain(scopes.iter().flatten())
                    .any(|d| d == p);
                if !p.is_empty() && p != "xml" && !in_scope && !seen.contains(p) {
                    seen.insert(p.to_owned());
                    missing.push(p.to_owned());
                }
            }
            if opens_scope {
                scopes.push(declared);
            }
        }
        Some((insert_at?, missing))
    }

    /// A xorshift generator of fragment pieces.
    struct Gen(u64);

    impl Gen {
        fn below(&mut self, n: usize) -> usize {
            self.0 ^= self.0 << 13;
            self.0 ^= self.0 >> 7;
            self.0 ^= self.0 << 17;
            (self.0 % n as u64) as usize
        }

        fn pick<'a>(&mut self, from: &[&'a str]) -> &'a str {
            from[self.below(from.len())]
        }

        /// One element with random prefixes, declarations and children.
        fn element(&mut self, out: &mut String, depth: usize) {
            // `None` is an unprefixed name; `Some("")` a colon-leading one.
            const PREFIXES: [Option<&str>; 8] = [
                None,
                None,
                Some("a"),
                Some("b"),
                Some("x14"),
                Some("xml"),
                Some("r"),
                Some(""),
            ];
            let name = match PREFIXES[self.below(PREFIXES.len())] {
                None => "e".to_owned(),
                Some(prefix) => format!("{prefix}:e"),
            };
            let _ = write!(out, "<{name}");
            for _ in 0..self.below(4) {
                match self.below(5) {
                    0 => {
                        let p = self.pick(&["a", "b", "x14", "r", "c"]);
                        let _ = write!(out, r#" xmlns:{p}="d{}""#, self.below(3));
                    }
                    1 => out.push_str(r#" xmlns="m""#),
                    2 => {
                        let p = self.pick(&["a", "b", "x14", "r", "xml", "c", ""]);
                        let _ = write!(out, r#" {p}:k{}="v""#, self.below(1_000));
                    }
                    _ => {
                        let _ = write!(out, r#" k{}="v&amp;""#, self.below(1_000));
                    }
                }
            }
            // White space before the tag's end, which the tag's slice keeps.
            let space = self.pick(&[
                "", "", "", " ", "
 ",
            ]);
            if depth >= 4 || self.below(3) == 0 {
                let _ = write!(out, "{space}/>");
                return;
            }
            let _ = write!(out, "{space}>");
            for _ in 0..self.below(4) {
                if self.below(4) == 0 {
                    out.push_str("t&lt;");
                } else {
                    self.element(out, depth + 1);
                }
            }
            let _ = write!(out, "</{name}>");
        }
    }

    /// Random `(prefix, uri)` bindings, the default namespace, a prefix
    /// bound twice and URIs that need escaping among them.
    fn bindings(g: &mut Gen, max: usize) -> Vec<(String, String)> {
        let uris = ["U", "X&Y", "q\"<", "", "urn:a", "a'b>c"];
        (0..g.below(max))
            .map(|_| {
                (
                    g.pick(&["", "a", "b", "x14", "r", "c", "xml"]).to_owned(),
                    g.pick(&uris).to_owned(),
                )
            })
            .collect()
    }

    /// 300,000 generated fragments (prefixed elements and attributes,
    /// colon-leading names, declarations that shadow at any depth, the
    /// default namespace, cut fragments) give the reference's output, with one
    /// scope of bindings (the part root's, some binding a prefix twice) and
    /// with two (an `<extLst>`'s before the root's, which the reference sees
    /// as one list, inner bindings first, since the first binding of a prefix
    /// wins).
    #[test]
    fn self_contained_fragment_matches_the_reference_on_generated_fragments() {
        let mut g = Gen(0x9E37_79B9_7F4A_7C15);
        let mut changed = 0;
        let mut twice = 0;
        let mut two_scopes = 0;
        for i in 0..300_000 {
            let mut frag = String::new();
            if g.below(4) == 0 {
                frag.push_str("\n ");
            }
            g.element(&mut frag, 0);
            if g.below(10) == 0 {
                // Not well-formed: cut anywhere on a character boundary.
                let mut cut = g.below(frag.len());
                while !frag.is_char_boundary(cut) {
                    cut -= 1;
                }
                frag.truncate(cut);
            }
            let mut uses = Vec::new();
            let mut seen = HashSet::new();
            let at = undeclared_prefixes(&frag, |p| {
                if seen.insert(p) {
                    uses.push(p.to_owned());
                }
            });
            assert_eq!(
                at.map(|at| (at, uses)),
                reference_undeclared(&frag),
                "case {i}: {frag}"
            );
            let inherited = bindings(&mut g, 8);
            let root = RootNamespaces::new(inherited.clone());
            twice += usize::from(root.entries.len() < inherited.len());
            let (got, expected) = if g.below(3) == 0 {
                two_scopes += 1;
                let inner = bindings(&mut g, 3);
                let all: Vec<(String, String)> = inner.iter().chain(&inherited).cloned().collect();
                let inner = RootNamespaces::new(inner);
                let got = match inherited_declarations(&frag, &[&inner, &root]) {
                    Some((at, decls)) => with_declarations(&frag, at, &decls),
                    None => frag.clone(),
                };
                (got, reference_fragment(&frag, &all))
            } else {
                (
                    self_contained_fragment(&frag, &root),
                    reference_fragment(&frag, &inherited),
                )
            };
            assert_eq!(got, expected, "case {i}: {frag} with {inherited:?}");
            changed += usize::from(got != frag);
        }
        // The generator reaches the inserting branch, a prefix bound twice
        // and two scopes, not only the copy.
        assert!(changed > 30_000, "{changed}");
        assert!(twice > 30_000, "{twice}");
        assert!(two_scopes > 30_000, "{two_scopes}");
    }

    #[test]
    fn root_namespaces_render_each_first_binding_once() {
        let ns: RootNamespaces = [
            (String::new(), "M&".to_owned()),
            ("b".to_owned(), "B\"<".to_owned()),
            ("a".to_owned(), "A".to_owned()),
            ("b".to_owned(), "later".to_owned()),
        ]
        .into_iter()
        .collect();
        assert_eq!(ns.declaration(""), Some(r#" xmlns="M&amp;""#));
        assert_eq!(ns.declaration("a"), Some(r#" xmlns:a="A""#));
        assert_eq!(ns.declaration("b"), Some(r#" xmlns:b="B&quot;&lt;""#));
        assert_eq!(ns.declaration("c"), None);
        assert_eq!(ns.declaration("ab"), None);
        assert_eq!(ns.list().len(), 4);
        assert!(ns.stored_bytes() >= ns.rendered.len() + 4 * size_of::<(String, String)>());
    }

    #[test]
    fn self_contained_fragment_is_linear_in_root_declarations_times_calls() {
        // 50,000 root declarations and 100,000 fragments, each using one of
        // them and declaring 1,000 prefixes of its own on one tag: the cost
        // is each fragment's length plus its prefixes, never a product.
        let inherited: RootNamespaces = (0..50_000)
            .map(|i| (format!("p{i}"), format!("u{i}")))
            .collect();
        let start = std::time::Instant::now();
        for i in 0..100_000 {
            let frag = format!("<ext><p{}:a/></ext>", i % 50_000);
            let out = self_contained_fragment(&frag, &inherited);
            assert!(out.starts_with("<ext xmlns:p"), "{out}");
        }
        let mut own = String::from("<ext");
        for i in 0..1_000 {
            let _ = write!(own, r#" xmlns:q{i}="v""#);
        }
        own.push('>');
        for i in 0..1_000 {
            let _ = write!(own, "<q{i}:a/>");
        }
        own.push_str("</ext>");
        for _ in 0..100 {
            assert_eq!(self_contained_fragment(&own, &inherited), own);
        }
        let elapsed = start.elapsed();
        assert!(
            elapsed < std::time::Duration::from_secs(5),
            "100,000 fragments took {elapsed:?}"
        );
    }

    #[test]
    fn fmt_xsd_double_known_values() {
        let mut b = XsdBuf::new();
        assert_eq!(fmt_xsd_double(15.0, &mut b, "t").unwrap(), "15");
        assert_eq!(fmt_xsd_double(-3.0, &mut b, "t").unwrap(), "-3");
        assert_eq!(fmt_xsd_double(0.0, &mut b, "t").unwrap(), "0");
        assert_eq!(fmt_xsd_double(-0.0, &mut b, "t").unwrap(), "-0");
        assert_eq!(fmt_xsd_double(8.43, &mut b, "t").unwrap(), "8.43");
        assert_eq!(
            fmt_xsd_double(1e15, &mut b, "t").unwrap(),
            "1000000000000000"
        );
        assert_eq!(fmt_xsd_double(1e16, &mut b, "t").unwrap(), "1e+16");
        for v in [1e16, 1e20, 9.3e18, 5e-324, f64::MAX, f64::MIN_POSITIVE] {
            let s = fmt_xsd_double(v, &mut b, "t").unwrap().to_owned();
            assert_eq!(
                s.parse::<f64>().unwrap().to_bits(),
                v.to_bits(),
                "{v} -> {s}"
            );
        }
        let s = fmt_xsd_double(1e20, &mut b, "t").unwrap();
        assert_ne!(s, "9223372036854775807");
    }

    #[test]
    fn fmt_xsd_double_rejects_non_finite() {
        let mut b = XsdBuf::new();
        for v in [f64::INFINITY, f64::NEG_INFINITY, f64::NAN] {
            let e = fmt_xsd_double(v, &mut b, "col width").unwrap_err();
            assert!(matches!(e, ModernXlsxError::XmlWrite(_)), "{v}: {e:?}");
            assert!(e.to_string().contains("col width"), "{e}");
        }
    }

    proptest! {
        /// Every finite f64 round-trips bit-exact (sign of zero included).
        #[test]
        fn fmt_xsd_double_roundtrips_every_finite(
            v in proptest::num::f64::ANY.prop_filter("finite", |v| v.is_finite())
        ) {
            let mut b = XsdBuf::new();
            let s = fmt_xsd_double(v, &mut b, "prop").unwrap();
            let back: f64 = s.parse().unwrap();
            prop_assert_eq!(back.to_bits(), v.to_bits(), "{} -> {}", v, s);
        }
    }

    #[test]
    fn sanitize_xml_text_borrows_when_clean() {
        // Clean text (including the legal control chars tab/LF/CR) is returned
        // borrowed — no allocation.
        let s = "hello\tworld\r\nok ✓ 日本語";
        assert!(matches!(sanitize_xml_text(s), Cow::Borrowed(_)));
        assert_eq!(sanitize_xml_text(s), s);
    }

    #[test]
    fn sanitize_xml_text_drops_illegal_chars() {
        // C0 control chars (other than tab/LF/CR) and the U+FFFE/U+FFFF
        // noncharacters are removed; legal whitespace is preserved.
        let dirty = "a\u{0001}b\u{0008}c\u{000B}d\u{001F}e\u{FFFE}f\u{FFFF}g\th\ni";
        assert_eq!(sanitize_xml_text(dirty), "abcdefg\th\ni");
    }

    #[test]
    fn is_illegal_xml_char_classification() {
        for c in [
            '\u{0000}', '\u{0001}', '\u{0008}', '\u{000B}', '\u{001F}', '\u{FFFE}', '\u{FFFF}',
        ] {
            assert!(is_illegal_xml_char(c), "{c:?} should be illegal");
        }
        for c in ['\t', '\n', '\r', ' ', 'a', '✓', '\u{FFFD}'] {
            assert!(!is_illegal_xml_char(c), "{c:?} should be legal");
        }
    }
}
