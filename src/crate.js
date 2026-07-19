// Crate assembly + output generation, ported from corpus-tools-dyirbal's
// index.js to use the ro-crate library directly (as the original does), plus
// xlsx (ro-crate-excel) and html (ro-crate-html-lite) generation.
//
// This module is ISOMORPHIC: it imports only browser-safe entry points
// (ro-crate, ro-crate-html-lite, and ro-crate-excel's lib/workbook.js — which
// avoids that package's Node-only shelljs/fs-extra modules), and returns bytes
// / strings rather than writing files. The caller (browser or Node) does I/O.

import { ROCrate } from "ro-crate";
// lib/preview.js is the package's real module (index.js only re-exports
// renderSinglePage). renderTemplate + roCrateToJSON drive the styled preview
// (custom template + config + css) on the dyirbal-workshop branch.
import { renderSinglePage, renderTemplate, roCrateToJSON } from "ro-crate-html-lite/lib/preview.js";
import Workbook from "ro-crate-excel/lib/workbook.js";
import ExcelJS from "exceljs";
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

// Returns the ro-crate-preview.html string.
//
// Two modes:
//  - Plain (default): the precompiled single-page template via renderSinglePage.
//    A bundled `layouts` object is passed so the library does NOT fetch its
//    default layout from GitHub at runtime (fragile + CORS-blocked in browser).
//  - Styled: pass opts.template (a template string, e.g. the tabular template),
//    opts.config (a preview config object: propertyGroups, settings,
//    navigationByType, termMapping, footer…), and opts.css (a stylesheet string).
//    Rendered via roCrateToJSON + renderTemplate, exactly as the CLI does.
export async function crateToPreviewHtml(crate, opts = {}) {
  const { layouts = { default: DEFAULT_LAYOUT }, template = null, config = null, css = "" } = opts;
  let html;
  if (template) {
    const cfg = config || {};
    const layout = (Array.isArray(cfg.propertyGroups) && cfg.propertyGroups.length)
      ? cfg.propertyGroups : DEFAULT_LAYOUT;
    const data = await roCrateToJSON(crate, cfg, layout);
    data.cratePath = "";
    data.layout = layout;
    data.hasLayout = true;
    html = await renderTemplate({ data, template, config: { ...cfg, propertyGroups: layout }, css, layout });
  } else {
    html = await renderSinglePage({ crate, layouts });
  }
  // ro-crate-html-lite urlencodes file links wholesale, turning "/" into "%2F"
  // and breaking relative navigation; only href values (not "#" anchors) are real links.
  html = html.replace(/href="([^"#][^"]*)"/g, (match, href) =>
    href.includes("%2F") ? `href="${href.replace(/%2F/g, "/")}"` : match
  );
  return html;
}

/* ---------- spreadsheet merge (ported from corpus-tools-dyirbal/merge.js) ---------- */

// ExcelJS cell values can be plain scalars or rich objects (formula results,
// rich text, hyperlinks); normalise to a plain string.
function cellText(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (typeof v.text === "string") return v.text;
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("");
    if (v.result !== undefined) return String(v.result);
    if (v.hyperlink && v.text) return String(v.text);
    return "";
  }
  return String(v);
}

// Merge a spreadsheet's rows into matching crate entities (by an "@id" column),
// applying the config's column→property mappings. Typed mappings split on comma
// or slash and generate linked entities. Mutates `crate` in place; returns stats.
export async function mergeXlsxIntoCrate(crate, xlsxData, mergeConfig, log = () => {}) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(xlsxData);

  let sheet;
  if (wb.worksheets.length > 1) {
    sheet = mergeConfig.sheet ? wb.getWorksheet(mergeConfig.sheet) : wb.worksheets[0];
    if (!sheet) throw new Error(`Sheet "${mergeConfig.sheet}" not found in the workbook`);
  } else {
    sheet = wb.worksheets[0];
  }
  if (!sheet) throw new Error("The spreadsheet has no worksheets");

  const headers = [];
  const dataRows = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) headers.push(...row.values.slice(1).map(cellText));
    else if (row.values.length > 1) dataRows.push(row.values.slice(1));
  });

  const idCol = headers.indexOf("@id");
  if (idCol === -1) throw new Error('The spreadsheet needs an "@id" column');

  const entityById = new Map();
  crate.graph.forEach((e) => { if (e["@id"]) entityById.set(e["@id"], e); });

  const mappings = Array.isArray(mergeConfig.mapping) ? mergeConfig.mapping : [];
  let merged = 0, generated = 0;
  const missingCols = new Set();
  const matchedIds = new Set();   // entity @ids that a spreadsheet row matched
  const unmatchedRowIds = [];     // spreadsheet @ids with no matching entity

  for (const row of dataRows) {
    const entityId = cellText(row[idCol]).trim();
    if (!entityId) continue;
    const entity = entityById.get(entityId);
    if (!entity) { unmatchedRowIds.push(entityId); continue; }
    matchedIds.add(entityId);

    for (const mapping of mappings) {
      const col = headers.indexOf(mapping.source);
      if (col === -1) { missingCols.add(mapping.source); continue; }
      const value = cellText(row[col]).trim();
      if (!value) continue;

      if (mapping.type) {
        const values = value
          .split(/\s*[,/]\s*/).map((v) => v.trim()).filter(Boolean)
          .map((v) => v.replace(/[\[\]?()']/g, "").trim()).filter(Boolean);
        if (!values.length) continue;
        const refs = values.map((val) => {
          const id = `#${val.replace(/[^a-zA-Z0-9]/g, "_").toLowerCase()}`;
          if (!entityById.get(id)) {
            const ge = { "@id": id, "@type": mapping.type, name: val };
            crate.addEntity(ge);
            entityById.set(id, ge);
            generated++;
          }
          return { "@id": id };
        });
        entity[mapping.target] = refs.length === 1 ? refs[0] : refs;
        merged++;
      } else {
        entity[mapping.target] = value;
        merged++;
      }
    }
  }

  // File entities in the crate that no spreadsheet row matched (by exact @id) —
  // these get no merged metadata (e.g. no encodingFormat). Usually a path/name
  // mismatch between the folder and the spreadsheet's @id column.
  const isFile = (e) => { const t = e["@type"]; return Array.isArray(t) ? t.includes("File") : t === "File"; };
  const unmatchedFiles = crate.graph.filter((e) => isFile(e) && e["@id"] && !matchedIds.has(e["@id"])).map((e) => e["@id"]);

  const sample = (arr, n = 12) => arr.slice(0, n).map((s) => `\n    • ${s}`).join("") + (arr.length > n ? `\n    …and ${arr.length - n} more` : "");

  if (missingCols.size) log(`Merge: columns not in spreadsheet, skipped: ${[...missingCols].join(", ")}.`, "warn");
  if (unmatchedFiles.length)
    log(`Merge: ${unmatchedFiles.length} file(s) had NO matching spreadsheet row — no metadata merged (check the @id path):${sample(unmatchedFiles)}`, "warn");
  if (unmatchedRowIds.length)
    log(`Merge: ${unmatchedRowIds.length} spreadsheet row(s) matched no entity in the crate:${sample(unmatchedRowIds)}`, "warn");
  log(`Merged ${merged} value(s) from "${sheet.name}" into ${matchedIds.size} entity/ies; generated ${generated} new entity/ies.`, "ok");
  return { merged, generated, skipped: unmatchedRowIds.length, unmatchedFiles: unmatchedFiles.length, sheet: sheet.name };
}
