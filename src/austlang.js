// Optional AUSTLANG language identification (filename-based only), ported from
// corpus-tools-dyirbal's index.js. PDF *content* extraction is NOT ported.
// Uses the global fetch, so it works in both Node (>=18) and the browser.
// In the browser it may be blocked by CORS — callers handle that gracefully.

const GENERIC_TERMS = new Set([
  "north", "south", "east", "west", "northern", "southern", "eastern", "western",
  "northeast", "northwest", "southeast", "southwest",
  "man", "woman", "men", "women", "people", "tribe", "group", "all", "our", "country",
  "none", "n/a", "na", "unknown", "unspecified", "nil", "various", "other",
]);

const DATAPACK_SEARCH_URL = "https://lookups.ldaca.edu.au/data/_search";
const DATAPACK_API_KEY = "ApiKey bXJWcEVvY0JrZXVEdG93dy14c046YndJOVBLcGFUVk9zQW0xN282NERSQQ=="; // read-only

function pBasename(p) { const i = p.lastIndexOf("/"); return i >= 0 ? p.slice(i + 1) : p; }
function pStripExt(name) { const b = pBasename(name); const i = b.lastIndexOf("."); return i > 0 ? b.slice(0, i) : b; }
function containsWholeWord(haystack, needle) {
  const esc = needle.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp(`\\b${esc}\\b`, "i").test(haystack);
}
function isSpecificEnough(name) {
  if (name.length < 4) return false;
  return !name.toLowerCase().split(/\s+/).every((w) => GENERIC_TERMS.has(w));
}

async function searchDatapack({ query, fields, limit = 10 }) {
  const res = await fetch(DATAPACK_SEARCH_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: DATAPACK_API_KEY },
    body: JSON.stringify({ query: { multi_match: { query, fields } }, from: 0, size: limit, sort: [] }),
  });
  if (res.status !== 200) return [];
  const body = await res.json();
  return body.hits.hits.map((hit) => ({ ...hit._source }));
}

async function findAustlangMatches(queryText, haystack, includeAlt) {
  const hits = await searchDatapack({ query: queryText, fields: ["name", "alternateName"], limit: 20 });
  const matches = hits.filter((hit) => {
    if (hit.source !== "Austlang") return false;
    const names = [hit.name, ...(includeAlt ? hit.alternateName || [] : [])].filter(isSpecificEnough);
    return names.some((name) => containsWholeWord(haystack, name));
  });
  matches.forEach((hit) => {
    if (hit.geo) hit.geo["@id"] = `#${String(hit.geo["@id"]).replace(/^#?_*/, "")}`;
    hit["@id"] = `#AUSTLANG_${String(hit.languageCode).replace(/[^A-Za-z0-9._-]/g, "_")}`;
    hit["custom:austlangCode"] = hit.languageCode;
    delete hit.languageCode; delete hit.source;
    if (hit["iso639-3"]) { hit["custom:iso639-3"] = hit["iso639-3"]; delete hit["iso639-3"]; }
    if (hit.glottologCode) { hit["custom:glottologCode"] = hit.glottologCode; delete hit.glottologCode; }
  });
  return matches;
}

async function mapWithConcurrency(items, concurrency, fn) {
  const results = new Array(items.length);
  let next = 0;
  async function worker() { while (next < items.length) { const i = next++; results[i] = await fn(items[i], i); } }
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) || 1 }, worker));
  return results;
}

/**
 * Identify subject languages for each file by filename.
 * @returns array aligned to filesWithMeta: [{ matchedLanguages: [...] }]
 */
export async function identifyAllLanguages(filesWithMeta, includeAlt, log = () => {}) {
  log(`Starting AUSTLANG lookups for ${filesWithMeta.length} file(s) (by filename)…`, "muted");
  const cache = new Map();
  let failed = false;
  return mapWithConcurrency(filesWithMeta, 8, async (file) => {
    if (failed) return { matchedLanguages: [] };
    const base = pStripExt(file.fileName);
    if (cache.has(base)) return { matchedLanguages: cache.get(base) };
    let matches = [];
    try {
      matches = await findAustlangMatches(base, base, includeAlt);
    } catch (e) {
      if (!failed) { failed = true; log("AUSTLANG lookup failed (likely CORS/network) — continuing without subject languages.", "warn"); }
      matches = [];
    }
    cache.set(base, matches);
    if (matches.length) log(`  ${file.fileName} → ${matches.map((m) => m.name).join(", ")}`, "muted");
    return { matchedLanguages: matches };
  });
}
