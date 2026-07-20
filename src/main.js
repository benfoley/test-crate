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
// Default column→property mapping for the spreadsheet merge. A folder may
// override it with its own merge-config.json (see processFolder).
import MERGE_CONFIG from "./merge_config.json";

const JSON_FILE = "ro-crate-metadata.json";
const XLSX_FILE = "ro-crate-metadata.xlsx";
const HTML_FILE = "ro-crate-preview.html";
const TEMPLATE_REPO_OWNER = "benfoley";
const TEMPLATE_REPO_NAME = "rocss-template-repo";
const TEMPLATE_REPO_REF = "main";

const OPTION_SCHEMA = [
  { key: "makeHtml", label: "Generate ro-crate-preview.html", default: true, children: [
    { key: "templateRepoPreset", label: "Template from rocss-template-repo", default: false,
      hint: "Pick a folder from the online template repo.", children: [
      { key: "templateRepoFolder", type: "select", label: "Template folder",
        placeholder: "Loading folders…", hint: "Select one folder from the template repo." },
    ] },
    { key: "templateSourceUrl", label: "Template folder URL", default: false,
      hint: "Use a public GitHub folder containing template/config/style files.", children: [
      { key: "templateSourceUrlValue", type: "url", label: "Template folder URL",
        placeholder: "https://github.com/<owner>/<repo>/tree/<branch>/<folder>",
        hint: "Downloads template/config/style from this folder." },
    ] },
    { key: "styledPreview", label: "Upload template files", default: false,
      hint: "Off = the library's plain preview.", children: [
      { key: "templateFile", type: "file", label: "Template (HTML)", accept: ".html,text/html",
        hint: "Optional. Uses your custom preview template; if omitted, the library default preview is used." },
      { key: "configFile", type: "file", label: "Config (JSON)", accept: ".json,application/json",
        hint: "Optional. Upload to override preview-config.json from the folder." },
      { key: "styleFile", type: "file", label: "Style (CSS)", accept: ".css,text/css",
        hint: "Optional. Upload to override preview-style.css from the folder." },
    ] },
  ] },
  { key: "enableLanguageLookups", label: "Identify subject languages (AUSTLANG, by filename)", default: false,
    hint: "Matches filenames against a bundled copy of the AUSTLANG data pack — fully offline, no network." },
  { key: "merge", label: "Merge metadata from a spreadsheet", default: false,
    hint: "Reads an .xlsx and merges its columns into matching entities (by their @id) before generating outputs.", children: [
    { key: "mergeFile", type: "file", label: "Spreadsheet (XLSX)", binary: true,
      accept: ".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      hint: "Rows are matched to entities by the @id column." },
    { key: "mergeConfigFile", type: "file", label: "Mapping config (JSON)", accept: ".json,application/json",
      hint: "Optional. Column→property mappings and sheet name. Overrides the bundled default and any merge-config.json in the folder." },
  ] },
];

// Shown in the Settings modal (accessed from the button next to Menu).
const SETTINGS_SCHEMA = [
  { key: "makeXlsx", label: "Generate ro-crate-metadata.xlsx", default: true },
  { key: "includeAlternateNames", label: "Match AUSTLANG alternate names", default: false,
    hint: "Only applies when “Identify subject languages” is on. More matches, more false positives." },
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
  $("settingsBtn").classList.toggle("hidden", name !== "view-build");
  $("rebuildBtn").classList.toggle("hidden", name !== "view-show");
}

/* ---------- options form ---------- */
// Uploaded config/style files (from the dropzones), keyed by option key:
// { templateFile: { name, file }, configFile: { name, file }, styleFile: { name, file } }
const uploads = {};

function hintEl(text) { const h = document.createElement("div"); h.className = "hint"; h.textContent = text; return h; }

function buildForm() {
  Object.keys(uploads).forEach((k) => delete uploads[k]);
  const form = $("optionsForm");
  form.innerHTML = "";
  renderOptions(OPTION_SCHEMA, form);
  setupTemplateSourceExclusivity();
  loadTemplateRepoFolderOptions();
  const settings = $("settingsForm");
  settings.innerHTML = "";
  renderOptions(SETTINGS_SCHEMA, settings);
}

function setupTemplateSourceExclusivity() {
  const uploadOpt = $("opt_styledPreview");
  const urlOpt = $("opt_templateSourceUrl");
  const repoOpt = $("opt_templateRepoPreset");
  if (!uploadOpt || !urlOpt || !repoOpt) return;

  const sync = (changed) => {
    const active = [];
    if (uploadOpt.checked) active.push("upload");
    if (urlOpt.checked) active.push("url");
    if (repoOpt.checked) active.push("repo");
    if (active.length <= 1) return;

    if (changed === "upload") { urlOpt.checked = false; repoOpt.checked = false; }
    else if (changed === "url") { uploadOpt.checked = false; repoOpt.checked = false; }
    else if (changed === "repo") { uploadOpt.checked = false; urlOpt.checked = false; }
    else { urlOpt.checked = false; repoOpt.checked = false; }

    uploadOpt.dispatchEvent(new Event("change"));
    urlOpt.dispatchEvent(new Event("change"));
    repoOpt.dispatchEvent(new Event("change"));
  };

  uploadOpt.addEventListener("change", () => sync("upload"));
  urlOpt.addEventListener("change", () => sync("url"));
  repoOpt.addEventListener("change", () => sync("repo"));
}

function renderOptions(schema, parent) {
  for (const opt of schema) {
    if (opt.type === "file") { parent.appendChild(buildFileField(opt)); continue; }
    if (opt.type === "url") { parent.appendChild(buildUrlField(opt)); continue; }
    if (opt.type === "select") { parent.appendChild(buildSelectField(opt)); continue; }

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

function buildUrlField(opt) {
  const wrap = document.createElement("div");
  wrap.className = "field";
  wrap.appendChild(Object.assign(document.createElement("div"), { className: "file-label", textContent: opt.label }));

  const input = document.createElement("input");
  input.type = "url";
  input.id = "opt_" + opt.key;
  input.placeholder = opt.placeholder || "https://";
  input.autocomplete = "off";
  input.style.width = "100%";
  input.style.padding = "9px 10px";
  input.style.borderRadius = "8px";
  input.style.border = "1px solid var(--border)";
  input.style.background = "var(--panel-2)";
  input.style.color = "var(--text)";
  input.style.fontFamily = "var(--mono)";
  input.style.fontSize = "12px";

  wrap.appendChild(input);
  if (opt.hint) wrap.appendChild(hintEl(opt.hint));
  return wrap;
}

function buildSelectField(opt) {
  const wrap = document.createElement("div");
  wrap.className = "field";
  wrap.appendChild(Object.assign(document.createElement("div"), { className: "file-label", textContent: opt.label }));

  const select = document.createElement("select");
  select.id = "opt_" + opt.key;
  select.style.width = "100%";
  select.style.padding = "9px 10px";
  select.style.borderRadius = "8px";
  select.style.border = "1px solid var(--border)";
  select.style.background = "var(--panel-2)";
  select.style.color = "var(--text)";
  select.style.fontFamily = "var(--mono)";
  select.style.fontSize = "12px";

  const ph = document.createElement("option");
  ph.value = "";
  ph.textContent = opt.placeholder || "Select…";
  select.appendChild(ph);

  wrap.appendChild(select);
  if (opt.hint) wrap.appendChild(hintEl(opt.hint));
  return wrap;
}

async function loadTemplateRepoFolderOptions() {
  const select = $("opt_templateRepoFolder");
  if (!select) return;
  select.disabled = true;
  try {
    const apiUrl = `https://api.github.com/repos/${TEMPLATE_REPO_OWNER}/${TEMPLATE_REPO_NAME}/contents?ref=${encodeURIComponent(TEMPLATE_REPO_REF)}`;
    const res = await fetch(apiUrl, { headers: { Accept: "application/vnd.github+json" } });
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    const entries = await res.json();
    if (!Array.isArray(entries)) throw new Error("Unexpected API response");
    const folders = entries
      .filter((e) => e && e.type === "dir" && typeof e.name === "string")
      .map((e) => e.name)
      .sort((a, b) => a.localeCompare(b));

    select.innerHTML = "";
    const ph = document.createElement("option");
    ph.value = "";
    ph.textContent = folders.length ? "Select a template folder…" : "No folders found";
    select.appendChild(ph);
    folders.forEach((name) => {
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      select.appendChild(opt);
    });
    select.disabled = folders.length === 0;
  } catch (e) {
    select.innerHTML = "";
    const ph = document.createElement("option");
    ph.value = "";
    ph.textContent = `Could not load folders (${e.message})`;
    select.appendChild(ph);
    select.disabled = true;
  }
}

function collectOptions(schema, o) {
  for (const opt of schema) {
    if (opt.type === "file") continue;
    if (opt.type === "url") {
      const el = $("opt_" + opt.key);
      if (el) o[opt.key] = (el.value || "").trim();
      continue;
    }
    if (opt.type === "select") {
      const el = $("opt_" + opt.key);
      if (el) o[opt.key] = el.value || "";
      continue;
    }
    const el = $("opt_" + opt.key);
    if (el) o[opt.key] = el.checked;
    if (opt.children) collectOptions(opt.children, o);
  }
}
function readOptions() {
  const o = {};
  collectOptions(OPTION_SCHEMA, o);
  collectOptions(SETTINGS_SCHEMA, o);
  o.templateUpload = uploads.templateFile || null;
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

function parseGitHubFolderUrl(rawUrl) {
  let url;
  try { url = new URL(rawUrl); }
  catch { throw new Error("Template folder URL is not a valid URL."); }
  if (url.hostname !== "github.com") {
    throw new Error("Template folder URL must be a github.com URL.");
  }
  const parts = url.pathname.split("/").filter(Boolean);
  if (parts.length < 5 || parts[2] !== "tree") {
    throw new Error("Use a GitHub folder URL like /owner/repo/tree/branch/path.");
  }
  const owner = parts[0];
  const repo = parts[1];
  const ref = decodeURIComponent(parts[3]);
  const folderPath = decodeURIComponent(parts.slice(4).join("/"));
  return { owner, repo, ref, folderPath };
}

function pickPreferredFile(files, ext, hints = []) {
  const byExt = files.filter((f) => f && f.type === "file" && typeof f.name === "string" && f.name.toLowerCase().endsWith(ext));
  if (!byExt.length) return null;
  for (const h of hints) {
    const found = byExt.find((f) => f.name.toLowerCase().includes(h));
    if (found) return found;
  }
  return byExt[0];
}

async function fetchTemplateBundleFromUrl(rawUrl) {
  const { owner, repo, ref, folderPath } = parseGitHubFolderUrl(rawUrl);
  const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${encodeURIComponent(folderPath).replace(/%2F/g, "/")}?ref=${encodeURIComponent(ref)}`;
  const res = await fetch(apiUrl, { headers: { Accept: "application/vnd.github+json" } });
  if (!res.ok) {
    throw new Error(`Could not read template folder (${res.status} ${res.statusText}).`);
  }
  const entries = await res.json();
  if (!Array.isArray(entries)) {
    throw new Error("Template URL does not point to a folder.");
  }

  const htmlFile = pickPreferredFile(entries, ".html", ["template", "preview"]);
  const jsonFile = pickPreferredFile(entries, ".json", ["config", "preview"]);
  const cssFile = pickPreferredFile(entries, ".css", ["style", "preview"]);

  const readText = async (entry) => {
    if (!entry || !entry.download_url) return null;
    const r = await fetch(entry.download_url);
    if (!r.ok) throw new Error(`Could not download ${entry.name} (${r.status} ${r.statusText}).`);
    return r.text();
  };

  const template = await readText(htmlFile);
  const configText = await readText(jsonFile);
  const css = (await readText(cssFile)) || "";
  let config = null;
  if (configText !== null) {
    try { config = JSON.parse(configText); }
    catch (e) { throw new Error(`Downloaded config file "${jsonFile.name}" is not valid JSON: ${e.message}`); }
  }

  return {
    template,
    config,
    css,
    sourceLabel: `${owner}/${repo}@${ref}/${folderPath}`,
    files: {
      template: htmlFile ? htmlFile.name : null,
      config: jsonFile ? jsonFile.name : null,
      style: cssFile ? cssFile.name : null,
    },
  };
}

function buildGitHubTreeUrl(owner, repo, ref, folderPath) {
  const safePath = String(folderPath || "").split("/").map((p) => encodeURIComponent(p)).join("/");
  return `https://github.com/${owner}/${repo}/tree/${encodeURIComponent(ref)}/${safePath}`;
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
        const urlSelected = !!options.templateSourceUrl;
        const repoSelected = !!options.templateRepoPreset;
        const sourceUrl = (options.templateSourceUrlValue || "").trim();
        if (options.styledPreview || urlSelected || repoSelected) {
          // Precedence for template/config/style: uploaded file → URL folder → local folder.
          let template = null, templateSrc = "none";
          let cfg = null, cfgSrc = "none";
          let css = "", cssSrc = "none";

          if (repoSelected) {
            const selectedFolder = (options.templateRepoFolder || "").trim();
            if (!selectedFolder) throw new Error("Template repo source is selected but no template folder was chosen.");
            const repoUrl = buildGitHubTreeUrl(TEMPLATE_REPO_OWNER, TEMPLATE_REPO_NAME, TEMPLATE_REPO_REF, selectedFolder);
            const remote = await fetchTemplateBundleFromUrl(repoUrl);
            template = remote.template;
            cfg = remote.config;
            css = remote.css;
            templateSrc = remote.files.template ? `repo (${selectedFolder}/${remote.files.template})` : `repo (${selectedFolder}; no html found)`;
            cfgSrc = remote.files.config ? `repo (${selectedFolder}/${remote.files.config})` : "none";
            cssSrc = remote.files.style ? `repo (${selectedFolder}/${remote.files.style})` : "none";
          }

          if (urlSelected) {
            if (!sourceUrl) throw new Error("Template folder URL is selected but no URL was provided.");
            const remote = await fetchTemplateBundleFromUrl(sourceUrl);
            template = remote.template;
            cfg = remote.config;
            css = remote.css;
            templateSrc = remote.files.template ? `URL (${remote.sourceLabel}/${remote.files.template})` : `URL (${remote.sourceLabel}; no html found)`;
            cfgSrc = remote.files.config ? `URL (${remote.sourceLabel}/${remote.files.config})` : "none";
            cssSrc = remote.files.style ? `URL (${remote.sourceLabel}/${remote.files.style})` : "none";
          }

          if (options.styledPreview && options.templateUpload) {
            template = await options.templateUpload.file.text();
            templateSrc = `uploaded (${options.templateUpload.name})`;
          } else {
            const folderTemplate = (!urlSelected && !repoSelected) ? await readFileText(dirHandle, "preview-template.html") : null;
            if (folderTemplate !== null) { template = folderTemplate; templateSrc = "preview-template.html from folder"; }
          }

          if (options.styledPreview && options.configUpload) {
            const cfgText = await options.configUpload.file.text();
            try { cfg = JSON.parse(cfgText); }
            catch (e) { throw new Error(`uploaded config "${options.configUpload.name}" is not valid JSON: ${e.message}`); }
            cfgSrc = `uploaded (${options.configUpload.name})`;
          } else if (!urlSelected && !repoSelected) {
            const folderCfg = await readJsonFromFolder(dirHandle, "preview-config.json");
            if (folderCfg) { cfg = folderCfg; cfgSrc = "preview-config.json from folder"; }
          }

          if (options.styledPreview && options.styleUpload) { css = await options.styleUpload.file.text(); cssSrc = `uploaded (${options.styleUpload.name})`; }
          else if (!urlSelected && !repoSelected) {
            const folderCss = await readFileText(dirHandle, "preview-style.css");
            if (folderCss !== null) { css = folderCss; cssSrc = "preview-style.css from folder"; }
          }
          if (template) {
            log(`Preview: styled tabular · template ${templateSrc} · config ${cfgSrc} · style ${cssSrc}.`, "muted");
            html = await crateToPreviewHtml(crate, { template, config: cfg, css });
          } else {
            log("Preview: plain (library default template; no custom template file provided).", "muted");
            html = await crateToPreviewHtml(crate);
          }
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
let previewFileUrls = [];

function revokePreviewUrls() {
  if (previewUrl) {
    URL.revokeObjectURL(previewUrl);
    previewUrl = null;
  }
  previewFileUrls.forEach((u) => URL.revokeObjectURL(u));
  previewFileUrls = [];
}

function isAbsoluteLikeUrl(value) {
  return /^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith("//") || value.startsWith("#");
}

function splitUrlParts(value) {
  const hashIdx = value.indexOf("#");
  const queryIdx = value.indexOf("?");
  const cut = [hashIdx, queryIdx].filter((n) => n >= 0).reduce((a, b) => Math.min(a, b), value.length);
  return {
    base: value.slice(0, cut),
    suffix: value.slice(cut),
  };
}

function normalizeRelativePath(value) {
  let v = (value || "").trim();
  if (!v) return "";
  try { v = decodeURIComponent(v); } catch { /* keep as-is */ }
  v = v.replace(/\\/g, "/").replace(/^\.\//, "").replace(/^\/+/, "");
  const out = [];
  for (const part of v.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") { out.pop(); continue; }
    out.push(part);
  }
  return out.join("/");
}

async function buildFileUrlMap(handle) {
  const map = new Map();
  const created = [];
  async function walk(h, prefix = "") {
    for await (const entry of h.values()) {
      if (entry.kind === "directory") {
        const next = prefix ? `${prefix}/${entry.name}` : entry.name;
        await walk(entry, next);
        continue;
      }
      if (entry.kind !== "file") continue;
      const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
      const file = await entry.getFile();
      const url = URL.createObjectURL(file);
      created.push(url);
      map.set(rel, url);
      // Accept both encoded and decoded lookup forms.
      map.set(encodeURI(rel), url);
    }
  }
  await walk(handle, "");
  return { map, created };
}

async function materializePreviewHtml(html, handle) {
  const { map, created } = await buildFileUrlMap(handle);
  const doc = new DOMParser().parseFromString(html, "text/html");
  for (const el of doc.querySelectorAll("[src],[href]")) {
    for (const attr of ["src", "href"]) {
      const raw = el.getAttribute(attr);
      if (!raw || isAbsoluteLikeUrl(raw)) continue;
      const { base, suffix } = splitUrlParts(raw);
      const key = normalizeRelativePath(base);
      if (!key) continue;
      const mapped = map.get(key) || map.get(encodeURI(key));
      if (mapped) el.setAttribute(attr, mapped + suffix);
    }
  }
  return { html: `<!doctype html>\n${doc.documentElement.outerHTML}`, created };
}

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
async function openHtmlInNewTab(html) {
  if (!html) return;
  const popup = window.open("about:blank", "_blank");
  if (!popup) return;
  popup.document.title = "Loading preview...";
  popup.document.body.textContent = "Loading preview...";
  try {
    revokePreviewUrls();
    let toOpen = html;
    if (dirHandle) {
      const materialized = await materializePreviewHtml(html, dirHandle);
      toOpen = materialized.html;
      previewFileUrls = materialized.created;
    }
    previewUrl = URL.createObjectURL(new Blob([toOpen], { type: "text/html" }));
    popup.location.replace(previewUrl);
  } catch (e) {
    popup.document.body.textContent = "Failed to open preview: " + (e && e.message ? e.message : e);
    console.error(e);
  }
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
  $("settingsBtn").addEventListener("click", () => $("settingsModal").classList.remove("hidden"));
  $("settingsClose").addEventListener("click", () => $("settingsModal").classList.add("hidden"));
  $("settingsModal").addEventListener("click", (e) => { if (e.target === $("settingsModal")) $("settingsModal").classList.add("hidden"); });
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
