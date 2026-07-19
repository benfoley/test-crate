// resources2crate — browser UI + File System Access wiring.
// The crate assembly and output generation live in ./crate.js (library-based,
// isomorphic). This file only handles picking a folder, reading/writing files,
// and the stepped Build/Show UI.

import {
  buildFileMetadata, buildCrate, crateToJsonString, crateToXlsxBytes, crateToPreviewHtml,
  mergeXlsxIntoCrate, GENERATED_FILENAMES, CONTROL_FILENAMES,
} from "./crate.js";
// ./austlang.js (and its bundled AUSTLANG data pack) is loaded lazily via
// dynamic import() only when language lookups are enabled — see run() — so the
// ~730 kB data pack stays out of the main bundle.
import { DEFAULT_CONFIG, DEFAULT_SAMPLE_DATA } from "./defaults.js";
// Bundled defaults for the styled ("tabular") HTML preview. A folder may override
// these with its own preview-config.json / preview-style.css (see processFolder).
import PREVIEW_TEMPLATE from "./preview_template.html?raw";
import PREVIEW_CONFIG from "./preview_config.json";
import PREVIEW_STYLE from "./preview_style.css?raw";
// Default column→property mapping for the spreadsheet merge. A folder may
// override it with its own merge-config.json (see processFolder).
import MERGE_CONFIG from "./merge_config.json";

const JSON_FILE = "ro-crate-metadata.json";
const XLSX_FILE = "ro-crate-metadata.xlsx";
const HTML_FILE = "ro-crate-preview.html";

const OPTION_SCHEMA = [
  { key: "makeXlsx", label: "Generate ro-crate-metadata.xlsx", default: true },
  { key: "makeHtml", label: "Generate ro-crate-preview.html", default: true, children: [
    { key: "styledPreview", label: "Styled tabular preview (custom template + CSS)", default: true,
      hint: "Off = the library's plain preview.", children: [
      { key: "configFile", type: "file", label: "Config (JSON)", accept: ".json,application/json",
        hint: "Optional. Overrides the bundled config and any preview-config.json in the folder." },
      { key: "styleFile", type: "file", label: "Style (CSS)", accept: ".css,text/css",
        hint: "Optional. Overrides the bundled style and any preview-style.css in the folder." },
    ] },
  ] },
  { key: "enableLanguageLookups", label: "Identify subject languages (AUSTLANG, by filename)", default: false,
    hint: "Matches filenames against a bundled copy of the AUSTLANG data pack — fully offline, no network." },
  { key: "includeAlternateNames", label: "…also match AUSTLANG alternate names", default: false,
    hint: "Only applies with the option above. More matches, more false positives." },
  { key: "merge", label: "Merge metadata from a spreadsheet", default: false,
    hint: "Reads an .xlsx and merges its columns into matching entities (by their @id) before generating outputs.", children: [
    { key: "mergeFile", type: "file", label: "Spreadsheet (XLSX)", binary: true,
      accept: ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      hint: "Rows are matched to entities by the @id column." },
    { key: "mergeConfigFile", type: "file", label: "Mapping config (JSON)", accept: ".json,application/json",
      hint: "Optional. Column→property mappings and sheet name. Overrides the bundled default and any merge-config.json in the folder." },
  ] },
  { key: "overwrite", label: "Overwrite existing outputs", default: true },
];

/* ---------- DOM helpers ---------- */
const $ = (id) => document.getElementById(id);
const logEl = () => $("log");
function log(msg, cls = "info") {
  const span = document.createElement("span");
  span.className = "l-" + cls;
  span.textContent = msg + "\n";
  logEl().appendChild(span);
  logEl().scrollTop = logEl().scrollHeight;
}
function clearLog() { logEl().textContent = ""; }

/* ---------- view routing ---------- */
const VIEWS = ["view-folder", "view-mode", "view-build", "view-show"];
function showView(name) {
  for (const v of VIEWS) $(v).classList.toggle("hidden", v !== name);
  $("contextBar").classList.toggle("hidden", !dirHandle);
  $("menuBtn").classList.toggle("hidden", !(name === "view-build" || name === "view-show"));
}

/* ---------- options form ---------- */
// Uploaded config/style files (from the dropzones), keyed by option key:
// { configFile: { name, text }, styleFile: { name, text } }
const uploads = {};

function hintEl(text) { const h = document.createElement("div"); h.className = "hint"; h.textContent = text; return h; }

function buildForm() {
  const form = $("optionsForm");
  form.innerHTML = "";
  Object.keys(uploads).forEach((k) => delete uploads[k]);
  renderOptions(OPTION_SCHEMA, form);
}

function renderOptions(schema, parent) {
  for (const opt of schema) {
    if (opt.type === "file") { parent.appendChild(buildFileField(opt)); continue; }

    const wrap = document.createElement("div");
    wrap.className = "field";
    const row = document.createElement("div");
    row.className = "checkbox";
    const input = document.createElement("input");
    input.type = "checkbox"; input.id = "opt_" + opt.key; input.checked = !!opt.default;
    const label = document.createElement("label");
    label.htmlFor = input.id; label.textContent = opt.label;
    row.append(input, label);
    wrap.appendChild(row);
    if (opt.hint) wrap.appendChild(hintEl(opt.hint));

    if (opt.children) {
      const panel = document.createElement("div");
      panel.className = "subpanel"; panel.id = "panel_" + opt.key;
      renderOptions(opt.children, panel);
      wrap.appendChild(panel);
      const sync = () => panel.classList.toggle("hidden", !input.checked);
      input.addEventListener("change", sync);
      sync();
    }
    parent.appendChild(wrap);
  }
}

function buildFileField(opt) {
  const wrap = document.createElement("div");
  wrap.className = "field file-field";
  wrap.appendChild(Object.assign(document.createElement("div"), { className: "file-label", textContent: opt.label }));

  const drop = document.createElement("label");
  drop.className = "dropzone"; drop.htmlFor = "file_" + opt.key;
  const dz = Object.assign(document.createElement("span"), { className: "dz-text", textContent: "Drop a file or click to choose" });
  drop.appendChild(dz);

  const input = document.createElement("input");
  input.type = "file"; input.id = "file_" + opt.key; input.accept = opt.accept || ""; input.className = "hidden";

  const clear = document.createElement("button");
  clear.type = "button"; clear.className = "secondary dz-clear hidden"; clear.textContent = "Remove";

  // Store the File itself; its bytes/text are read at build time (supports
  // binary files like .xlsx as well as text config/style).
  const setFile = (file) => {
    uploads[opt.key] = { name: file.name, file };
    dz.textContent = file.name; drop.classList.add("has-file"); clear.classList.remove("hidden");
  };
  const clearFile = () => {
    delete uploads[opt.key];
    dz.textContent = "Drop a file or click to choose"; drop.classList.remove("has-file");
    clear.classList.add("hidden"); input.value = "";
  };

  input.addEventListener("change", () => { const f = input.files[0]; if (f) setFile(f); });
  drop.addEventListener("dragover", (e) => { e.preventDefault(); drop.classList.add("drag"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("drag"));
  drop.addEventListener("drop", (e) => {
    e.preventDefault(); drop.classList.remove("drag");
    const f = e.dataTransfer.files[0]; if (f) setFile(f);
  });
  clear.addEventListener("click", clearFile);

  wrap.append(drop, input, clear);
  if (opt.hint) wrap.appendChild(hintEl(opt.hint));
  return wrap;
}

function collectOptions(schema, o) {
  for (const opt of schema) {
    if (opt.type === "file") continue;
    const el = $("opt_" + opt.key);
    if (el) o[opt.key] = el.checked;
    if (opt.children) collectOptions(opt.children, o);
  }
}
function readOptions() {
  const o = {};
  collectOptions(OPTION_SCHEMA, o);
  o.configUpload = uploads.configFile || null;
  o.styleUpload = uploads.styleFile || null;
  o.mergeUpload = uploads.mergeFile || null;
  o.mergeConfigUpload = uploads.mergeConfigFile || null;
  return o;
}

/* ---------- File System Access ---------- */
let dirHandle = null;

async function verifyPermission(handle, readWrite) {
  const opts = { mode: readWrite ? "readwrite" : "read" };
  if ((await handle.queryPermission(opts)) === "granted") return true;
  if ((await handle.requestPermission(opts)) === "granted") return true;
  return false;
}
async function walkDirectory(handle, prefix = "") {
  const files = [];
  for await (const entry of handle.values()) {
    const nm = entry.name;
    if (nm.startsWith(".") || nm.startsWith("~$")) continue;
    if (GENERATED_FILENAMES.has(nm) || CONTROL_FILENAMES.has(nm)) continue;
    const rel = prefix ? prefix + "/" + nm : nm;
    if (entry.kind === "file") files.push({ fileName: nm, relativePath: rel });
    else if (entry.kind === "directory") files.push(...await walkDirectory(entry, rel));
  }
  return files;
}
async function writeFile(handle, filename, contents) {
  const fh = await handle.getFileHandle(filename, { create: true });
  const w = await fh.createWritable();
  await w.write(contents);
  await w.close();
}
async function fileExists(handle, filename) {
  try { await handle.getFileHandle(filename, { create: false }); return true; }
  catch { return false; }
}
async function readFileText(handle, filename) {
  try {
    const fh = await handle.getFileHandle(filename, { create: false });
    return await (await fh.getFile()).text();
  } catch (e) {
    if (e && e.name === "NotFoundError") return null;
    throw e;
  }
}
async function readJsonFromFolder(handle, filename) {
  const text = await readFileText(handle, filename);
  if (text === null) return null;
  try { return JSON.parse(text); }
  catch (e) { throw new Error(`${filename} in the folder is not valid JSON: ${e.message}`); }
}

/* ---------- Build ---------- */
async function processFolder(dirHandle, files, options) {
  const config = (await readJsonFromFolder(dirHandle, "config.json")) || DEFAULT_CONFIG;
  const sampleData = (await readJsonFromFolder(dirHandle, "sample-data.json")) || DEFAULT_SAMPLE_DATA;
  log(
    `Config: ${config === DEFAULT_CONFIG ? "built-in default" : "config.json from folder"} · ` +
    `Sample data: ${sampleData === DEFAULT_SAMPLE_DATA ? "built-in default" : "sample-data.json from folder"}.`,
    "muted"
  );

  files.sort((a, b) => a.relativePath.localeCompare(b.relativePath));
  const filesWithMeta = buildFileMetadata(files);
  log(`Scanned ${filesWithMeta.length} file(s).`, "info");

  let langByIndex = null;
  if (options.enableLanguageLookups) {
    const { identifyAllLanguages } = await import("./austlang.js");
    langByIndex = await identifyAllLanguages(filesWithMeta, options.includeAlternateNames, log);
  }

  const crate = buildCrate(filesWithMeta, config, sampleData, langByIndex, log);

  // Optional: merge metadata from an uploaded spreadsheet (before outputs).
  if (options.merge && options.mergeUpload) {
    // Mapping config precedence: uploaded file → folder file → bundled default.
    let mergeConfig = MERGE_CONFIG, mcSrc = "bundled default";
    if (options.mergeConfigUpload) {
      const mcText = await options.mergeConfigUpload.file.text();
      try { mergeConfig = JSON.parse(mcText); }
      catch (e) { throw new Error(`uploaded merge config "${options.mergeConfigUpload.name}" is not valid JSON: ${e.message}`); }
      mcSrc = `uploaded (${options.mergeConfigUpload.name})`;
    } else {
      const folderMc = await readJsonFromFolder(dirHandle, "merge-config.json");
      if (folderMc) { mergeConfig = folderMc; mcSrc = "merge-config.json from folder"; }
    }
    log(`Merging ${options.mergeUpload.name} · mapping ${mcSrc}.`, "muted");
    const bytes = await options.mergeUpload.file.arrayBuffer();
    await mergeXlsxIntoCrate(crate, bytes, mergeConfig, log);
  } else if (options.merge && !options.mergeUpload) {
    log("Merge is on but no spreadsheet was selected — skipping merge.", "warn");
  }

  const entities = crate.getJson()["@graph"].length;

  // ro-crate-metadata.json
  if (options.overwrite || !(await fileExists(dirHandle, JSON_FILE))) {
    await writeFile(dirHandle, JSON_FILE, crateToJsonString(crate));
    log(`Wrote ${JSON_FILE}.`, "ok");
  } else log(`${JSON_FILE} exists and overwrite is off — skipped.`, "warn");

  // ro-crate-metadata.xlsx
  if (options.makeXlsx) {
    if (options.overwrite || !(await fileExists(dirHandle, XLSX_FILE))) {
      const bytes = await crateToXlsxBytes(crate);
      await writeFile(dirHandle, XLSX_FILE, bytes);
      log(`Wrote ${XLSX_FILE}.`, "ok");
    } else log(`${XLSX_FILE} exists and overwrite is off — skipped.`, "warn");
  }

  // ro-crate-preview.html
  if (options.makeHtml) {
    if (options.overwrite || !(await fileExists(dirHandle, HTML_FILE))) {
      try {
        let html;
        if (options.styledPreview) {
          // Precedence for both config + style: uploaded file → folder file → bundled default.
          let cfg = PREVIEW_CONFIG, cfgSrc = "bundled default";
          if (options.configUpload) {
            const cfgText = await options.configUpload.file.text();
            try { cfg = JSON.parse(cfgText); }
            catch (e) { throw new Error(`uploaded config "${options.configUpload.name}" is not valid JSON: ${e.message}`); }
            cfgSrc = `uploaded (${options.configUpload.name})`;
          } else {
            const folderCfg = await readJsonFromFolder(dirHandle, "preview-config.json");
            if (folderCfg) { cfg = folderCfg; cfgSrc = "preview-config.json from folder"; }
          }
          let css = PREVIEW_STYLE, cssSrc = "bundled default";
          if (options.styleUpload) { css = await options.styleUpload.file.text(); cssSrc = `uploaded (${options.styleUpload.name})`; }
          else {
            const folderCss = await readFileText(dirHandle, "preview-style.css");
            if (folderCss !== null) { css = folderCss; cssSrc = "preview-style.css from folder"; }
          }
          log(`Preview: styled tabular · config ${cfgSrc} · style ${cssSrc}.`, "muted");
          html = await crateToPreviewHtml(crate, { template: PREVIEW_TEMPLATE, config: cfg, css });
        } else {
          log("Preview: plain (library default template).", "muted");
          html = await crateToPreviewHtml(crate);
        }
        await writeFile(dirHandle, HTML_FILE, html);
        log(`Wrote ${HTML_FILE}.`, "ok");
      } catch (e) {
        log(`HTML preview failed: ${e.message}`, "err");
      }
    } else log(`${HTML_FILE} exists and overwrite is off — skipped.`, "warn");
  }

  return { files: filesWithMeta.length, entities };
}

let buildHtml = null;  // ro-crate-preview.html captured after the last successful build

async function run() {
  if (!dirHandle) return;
  const runBtn = $("runBtn");
  runBtn.disabled = true; runBtn.textContent = "Building…";
  $("showHtmlBtn").classList.add("hidden"); buildHtml = null;
  const started = performance.now();
  $("statFiles").textContent = "—"; $("statEntities").textContent = "—"; $("statTime").textContent = "—";
  try {
    if (!(await verifyPermission(dirHandle, true))) { log("Permission to read/write the folder was denied.", "err"); return; }
    const options = readOptions();
    const files = await walkDirectory(dirHandle);
    const result = await processFolder(dirHandle, files, options);
    $("statFiles").textContent = result.files;
    $("statEntities").textContent = result.entities;
    const secs = ((performance.now() - started) / 1000).toFixed(2);
    $("statTime").textContent = secs + "s";
    log("Done in " + secs + "s.", "ok");
    // Capture the generated preview so the build-view button can open it in a
    // new tab synchronously (no await between the click and window.open).
    buildHtml = await readFileText(dirHandle, HTML_FILE);
    if (buildHtml !== null) $("showHtmlBtn").classList.remove("hidden");
  } catch (e) {
    log("Error: " + (e && e.message ? e.message : e), "err");
    console.error(e);
  } finally {
    runBtn.disabled = false; runBtn.textContent = "Build RO-Crate";
    $("saveLogBtn").disabled = false;
  }
}

/* ---------- actions ---------- */
async function pickFolder() {
  try {
    dirHandle = await window.showDirectoryPicker({ mode: "readwrite" });
  } catch (e) {
    if (e && e.name === "AbortError") return;
    console.error("Could not open folder:", e);
    return;
  }
  $("ctxFolder").textContent = dirHandle.name;
  await refreshModeCards();
  showView("view-mode");
}

// Offer "Show" only when the folder already has crate output to view:
// an ro-crate-metadata.json or an ro-crate-preview.html. A fresh folder with
// neither shows the Build card alone.
async function refreshModeCards() {
  let hasJson = false, hasHtml = false;
  if (dirHandle) {
    try {
      hasJson = await fileExists(dirHandle, JSON_FILE);
      hasHtml = await fileExists(dirHandle, HTML_FILE);
    } catch { /* treat as none → hide Show */ }
  }
  $("cardShow").classList.toggle("hidden", !(hasJson || hasHtml));
}
function openBuild() {
  clearLog();
  $("showHtmlBtn").classList.add("hidden");
  $("saveLogBtn").disabled = true;
  log("Set your options, then click Build RO-Crate.", "muted");
  showView("view-build");
}

// Download the current build log as a .log file.
function saveLog() {
  const text = $("log").textContent || "";
  if (!text.trim()) return;
  const name = `resources2crate-${dirHandle ? dirHandle.name : "build"}.log`;
  const url = URL.createObjectURL(new Blob([text], { type: "text/plain" }));
  const a = document.createElement("a");
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}
let showHtml = null, showJson = null, previewUrl = null;

async function openShow() {
  if (!dirHandle) return;
  try {
    if (!(await verifyPermission(dirHandle, false))) return;
    showHtml = await readFileText(dirHandle, HTML_FILE);
    showJson = await readFileText(dirHandle, JSON_FILE);
    if (showHtml === null && showJson === null) { $("modal").classList.remove("hidden"); return; }
    showView("view-show");
    renderShow(showHtml !== null ? "preview" : "json");
  } catch (e) {
    $("showFileName").textContent = "";
    $("showPreview").classList.add("hidden");
    const pane = $("showPane");
    pane.classList.remove("hidden");
    pane.textContent = "Error reading the RO-Crate: " + (e && e.message ? e.message : e);
    showView("view-show");
  }
}

function renderShow(mode) {
  const preview = $("showPreview"), pane = $("showPane");
  const tabP = $("showTabPreview"), tabJ = $("showTabJson");
  // Fall back to whichever file is present if the requested one is missing.
  if (mode === "preview" && showHtml === null) mode = "json";
  if (mode === "json" && showJson === null) mode = "preview";

  tabP.disabled = showHtml === null;
  tabJ.disabled = showJson === null;
  tabP.classList.toggle("active", mode === "preview");
  tabJ.classList.toggle("active", mode === "json");

  if (mode === "preview") {
    $("showFileName").textContent = HTML_FILE;
    pane.classList.add("hidden");
    pane.textContent = "";
    preview.classList.remove("hidden");
  } else {
    let pretty = showJson;
    try { pretty = JSON.stringify(JSON.parse(showJson), null, 2); } catch { /* raw */ }
    $("showFileName").textContent = JSON_FILE;
    preview.classList.add("hidden");
    pane.classList.remove("hidden");
    pane.textContent = pretty;
  }
}

// Open HTML as a real document in a new browser tab. The generated
// ro-crate-preview.html relies on in-page (:target) links to toggle tables,
// which don't work inside an embedded/srcdoc frame, so it needs its own URL.
// Must be called synchronously from a click handler (no awaits before it) so
// the browser doesn't treat window.open as an unsolicited popup.
function openHtmlInNewTab(html) {
  if (!html) return;
  if (previewUrl) URL.revokeObjectURL(previewUrl);
  previewUrl = URL.createObjectURL(new Blob([html], { type: "text/html" }));
  window.open(previewUrl, "_blank");
}
function openPreviewWindow() { openHtmlInNewTab(showHtml); }

/* ---------- boot ---------- */
function boot() {
  if (!("showDirectoryPicker" in window)) { $("unsupported").classList.remove("hidden"); return; }
  $("app").classList.remove("hidden");
  buildForm();
  showView("view-folder");

  $("pickBtn").addEventListener("click", pickFolder);
  $("changeFolderBtn").addEventListener("click", pickFolder);
  $("menuBtn").addEventListener("click", async () => { await refreshModeCards(); showView("view-mode"); });
  $("cardBuild").addEventListener("click", openBuild);
  $("cardShow").addEventListener("click", openShow);
  $("showTabPreview").addEventListener("click", () => renderShow("preview"));
  $("showTabJson").addEventListener("click", () => renderShow("json"));
  $("openPreviewBtn").addEventListener("click", openPreviewWindow);
  const key = (fn) => (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); fn(); } };
  $("cardBuild").addEventListener("keydown", key(openBuild));
  $("cardShow").addEventListener("keydown", key(openShow));
  $("runBtn").addEventListener("click", run);
  $("showHtmlBtn").addEventListener("click", () => openHtmlInNewTab(buildHtml));
  $("saveLogBtn").addEventListener("click", saveLog);
  $("rebuildBtn").addEventListener("click", openBuild);
  $("modalCancel").addEventListener("click", () => $("modal").classList.add("hidden"));
  $("modalBuild").addEventListener("click", () => { $("modal").classList.add("hidden"); openBuild(); });
}
boot();
