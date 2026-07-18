// Crate assembly + output generation, ported from corpus-tools-dyirbal's
// index.js to use the ro-crate library directly (as the original does), plus
// xlsx (ro-crate-excel) and html (ro-crate-html-lite) generation.
//
// This module is ISOMORPHIC: it imports only browser-safe entry points
// (ro-crate, ro-crate-html-lite, and ro-crate-excel's lib/workbook.js — which
// avoids that package's Node-only shelljs/fs-extra modules), and returns bytes
// / strings rather than writing files. The caller (browser or Node) does I/O.

import { ROCrate } from "ro-crate";
import { renderSinglePage } from "ro-crate-html-lite";
import Workbook from "ro-crate-excel/lib/workbook.js";
import { CUSTOM_PROPERTIES } from "./defaults.js";
import { DEFAULT_LAYOUT } from "./default_layout.js";

/* Files that are generated output or local control files — never treated as
 * corpus data (mirrors GENERATED_FILENAMES in the original). */
export const GENERATED_FILENAMES = new Set([
  "ro-crate-metadata.json", "ro-crate-metadata.jsonld", "ro-crate-metadata.xlsx", "ro-crate-preview.html",
]);
export const CONTROL_FILENAMES = new Set(["config.json", "sample-data.json"]);

/* ---------- path + name helpers (relative paths use "/" separators) ---------- */
function pBasename(p) { const i = p.lastIndexOf("/"); return i >= 0 ? p.slice(i + 1) : p; }
function pStripExt(name) { const b = pBasename(name); const i = b.lastIndexOf("."); return i > 0 ? b.slice(0, i) : b; }
function pDirname(p) { const i = p.lastIndexOf("/"); return i >= 0 ? p.slice(0, i) : ""; }
function sanitizeUrl(rel) { return rel.replace(/ /g, "_"); }
function normalizeName(fileName) {
  return pStripExt(fileName).toLowerCase()
    .replace(/\b(copy|duplicate)\b/g, "")
    .replace(/\(\d+\)/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/* ---------- per-file metadata (folder grouping + duplicate detection) ---------- */
export function buildFileMetadata(files) {
  const filesWithMeta = files.map((file) => {
    const folders = pDirname(file.relativePath).split("/").filter((p) => p !== "" && p !== ".");
    const topLevelName = folders.length > 0 ? folders[0] : pStripExt(file.fileName);
    return {
      ...file,
      id: file.relativePath,
      isPartOfId: `#${sanitizeUrl(topLevelName)}`,
      isPartOfName: topLevelName,
    };
  });
  const groups = new Map();
  filesWithMeta.forEach((f) => {
    const key = normalizeName(f.fileName);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(f);
  });
  filesWithMeta.forEach((f) => {
    const group = groups.get(normalizeName(f.fileName));
    f.possibleDuplicates = group.filter((o) => o !== f).map((o) => o.id);
  });
  return filesWithMeta;
}

/* ---------- entity builders (using the ROCrate library) ---------- */
function addSampleData(crate, sampleData) {
  [...(sampleData.people || []), ...(sampleData.places || []), ...(sampleData.localities || [])]
    .forEach((entity) => crate.addEntity(entity));
}

function addFolderEntities(crate, filesWithMeta) {
  const folderGroups = new Map();
  filesWithMeta.forEach((file) => {
    if (!folderGroups.has(file.isPartOfId)) folderGroups.set(file.isPartOfId, { name: file.isPartOfName, fileIds: [] });
    folderGroups.get(file.isPartOfId).fileIds.push(file.id);
  });
  folderGroups.forEach((group, id) => {
    crate.addEntity({
      "@id": id,
      "@type": "RepositoryObject",
      conformsTo: { "@id": "https://w3id.org/ldac/profile#Object" },
      name: group.name,
      description: "",
      datePublished: "",
      ...(crate.rootDataset.license?.length
        ? { license: crate.rootDataset.license.map((license) => ({ "@id": license["@id"] })) }
        : {}),
      hasPart: group.fileIds.map((fileId) => ({ "@id": fileId })),
    });
  });
  crate.rootDataset["pcdm:hasMember"] = [...folderGroups.keys()].map((id) => ({ "@id": id }));
}

function addFileEntities(crate, filesWithMeta, langByIndex) {
  filesWithMeta.forEach((file, index) => {
    const matched = langByIndex ? langByIndex[index].matchedLanguages : [];
    crate.addEntity({
      "@id": file.id,
      "@type": "File",
      name: file.fileName,
      description: "",
      datePublished: "",
      "custom:participant": "",
      "custom:compiler": "",
      contentLocation: "",
      isPartOf: { "@id": file.isPartOfId },
      ...(file.possibleDuplicates.length
        ? { "custom:possibleDuplicate": file.possibleDuplicates.map((id) => ({ "@id": id })) }
        : {}),
      ...(matched.length
        ? { "ldac:subjectLanguage": matched.map((l) => ({ "@id": l["@id"] })) }
        : {}),
    });
  });
}

function addLanguageEntities(crate, langByIndex) {
  const identified = new Map();
  langByIndex.forEach((r) => r.matchedLanguages.forEach((l) => identified.set(l["@id"], l)));
  identified.forEach((language) => {
    crate.addEntity(language);
    if (language.geo) crate.addEntity(language.geo);
  });
  return identified.size;
}

/* Rewrite hash @ids of RepositoryObject entities to arcp form, and every
 * reference to them across the graph (mirrors rewriteHashIdsForExport). */
function rewriteHashIdsForExport(crate) {
  const datasetId = String(crate.rootDataset["@id"] || "").trim();
  if (!datasetId) return;
  const arcpBase = `${datasetId}/`;
  const idMap = new Map();
  crate.graph.forEach((entity) => {
    const types = Array.isArray(entity?.["@type"]) ? entity["@type"] : [entity?.["@type"]];
    const oldId = entity?.["@id"];
    if (types.includes("RepositoryObject") && typeof oldId === "string" && oldId.startsWith("#")) {
      idMap.set(oldId, `${arcpBase}${oldId.slice(1)}`);
    }
  });
  if (!idMap.size) return;
  const seen = new WeakSet();
  (function visit(value) {
    if (!value || typeof value !== "object" || seen.has(value)) return;
    seen.add(value);
    if (Array.isArray(value)) { value.forEach(visit); return; }
    if (typeof value["@id"] === "string" && idMap.has(value["@id"])) value["@id"] = idMap.get(value["@id"]);
    Object.keys(value).forEach((k) => visit(value[k]));
  })(crate.graph);
}

/* ---------- top-level: build the ROCrate ---------- */
export function buildCrate(filesWithMeta, config, sampleData, langByIndex, log = () => {}) {
  const crate = new ROCrate({ array: true, link: true });
  crate.addContext({ ldac: "https://w3id.org/ldac/terms#" });
  crate.addContext({ pcdm: "http://pcdm.org/models#" });
  crate.addContext({ custom: "arcp://name,custom/terms#" });
  crate.addContext({ AUSTLANG: "https://collection.aiatsis.gov.au/austlang/language/" });

  Object.assign(crate.rootDataset, config.rootDataset);
  if (typeof crate.rootDataset["@id"] === "string" && crate.rootDataset["@id"].trim()) {
    crate.descriptor.about = { "@id": crate.rootDataset["@id"] };
  }
  if (config.metadataLicence?.["@id"]) {
    crate.descriptor.license = { "@id": config.metadataLicence["@id"] };
    crate.addEntity(config.metadataLicence);
  }

  for (const p of CUSTOM_PROPERTIES) crate.addEntity(p);
  addSampleData(crate, sampleData);
  addFolderEntities(crate, filesWithMeta);
  addFileEntities(crate, filesWithMeta, langByIndex);
  if (langByIndex) {
    const n = addLanguageEntities(crate, langByIndex);
    log(`Identified ${n} unique language(s).`, n ? "ok" : "muted");
  }
  rewriteHashIdsForExport(crate);
  return crate;
}

/* ---------- output generators ---------- */
export function crateToJsonString(crate) {
  return JSON.stringify(crate.getJson(), null, 2);
}

// Returns bytes for ro-crate-metadata.xlsx (Uint8Array in browser, Buffer in Node).
export async function crateToXlsxBytes(crate) {
  const workbook = new Workbook({ crate });
  await workbook.crateToWorkbook();
  return workbook.workbook.xlsx.writeBuffer();
}

// Returns the ro-crate-preview.html string. By default it passes a bundled
// `layouts` object so ro-crate-html-lite does NOT fetch its default layout from
// GitHub at runtime (that fetch is fragile and CORS-blocked in the browser).
export async function crateToPreviewHtml(crate, layouts = { default: DEFAULT_LAYOUT }) {
  let html = await renderSinglePage({ crate, layouts });
  // ro-crate-html-lite urlencodes file links wholesale, turning "/" into "%2F"
  // and breaking relative navigation; only href values (not "#" anchors) are real links.
  html = html.replace(/href="([^"#][^"]*)"/g, (match, href) =>
    href.includes("%2F") ? `href="${href.replace(/%2F/g, "/")}"` : match
  );
  return html;
}
