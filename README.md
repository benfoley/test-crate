# resources2crate

A zero-install, single-file web app that lets a user **pick a local folder, configure
some options, and run a command over that folder in place** — reading its files and
writing results back — entirely in the browser. No Node, no installer, no code signing.

It's built on the **File System Access API**, so it works in Chromium browsers
(Chrome / Edge) and needs a *secure context* (`https://` or `http://localhost`).

---

## Running it

The one catch with the File System Access API: it will **not** run from a
`file://` URL (opening `index.html` by double-clicking won't work). Serve it over
`localhost` or host it over HTTPS.

**Locally (quickest):**

```bash
cd resources2crate

# Python (already on macOS):
python3 -m http.server 8000

# …or Node:
npx serve .
```

Then open <http://localhost:8000> in **Chrome or Edge**.

**For real use / handing to others:** upload `index.html` to any static HTTPS host
(GitHub Pages, Netlify, Cloudflare Pages, S3+CloudFront, an internal server…). Because
it's one static file with no backend, hosting is trivial and there's nothing to sign.
The user's files never leave their machine — all reading and writing happens locally in
the browser.

> Browser support: Chrome/Edge (and Chromium derivatives). Safari and Firefox do **not**
> support directory access, so they'll see the "can't run" notice.

---

## Where your logic goes

Everything is in `index.html`. You only need to touch two spots:

1. **`OPTION_SCHEMA`** (top of the `<script>`) — declaratively defines the config form.
   Add/remove entries and the UI plus the `options` object update automatically.
   Supported field types: `text`, `number`, `checkbox`, `select`.

2. **`processFolder(dirHandle, files, options, ctx)`** — the single function marked
   `>>> YOUR LOGIC GOES HERE <<<`. Replace the placeholder body with your real command.
   It receives:

   | arg         | what it is                                                            |
   |-------------|-----------------------------------------------------------------------|
   | `dirHandle` | the chosen directory handle (read **and** write)                      |
   | `files`     | array of `{ path, name, size, lastModified, type }` from the walk     |
   | `options`   | the config object from the form (keys match `OPTION_SCHEMA`)          |
   | `ctx`       | helpers: `log(msg, level)`, `writeFile(dir, name, text)`, `fileExists(dir, name)` |

   Return an object; the `written` count you report is shown in the Results panel.

The placeholder just scans the folder and writes a `manifest.json` back into it — enough
to prove the read + write round-trip works end to end.

### Handy building blocks already wired up

- `walkDirectory(handle, options)` — recursive folder walk, respects `maxDepth` and
  `includeHidden`.
- `writeFile(handle, filename, contents)` — create/overwrite a file in the folder.
- `fileExists(handle, filename)` — top-level existence check.
- `verifyPermission(handle, readWrite)` — re-prompts for access if needed.
- `log(msg, level)` — append to the log pane; levels: `info`, `muted`, `ok`, `warn`, `err`.

To read a file's contents inside `processFolder`, get its handle and call `.getFile()`:

```js
const fh = await dirHandle.getFileHandle(someName);
const text = await (await fh.getFile()).text();
```

For nested writes (into subfolders), get the subdirectory handle first with
`dirHandle.getDirectoryHandle(name, { create: true })`.

---

## Current options (placeholder)

| key            | type     | purpose                                             |
|----------------|----------|-----------------------------------------------------|
| `crateName`    | text     | name written into the manifest                      |
| `outputFile`   | text     | filename created/overwritten in the folder          |
| `maxDepth`     | number   | how many levels deep to scan (0 = top level only)   |
| `format`       | select   | `json` (pretty) or `jsonl` (one line per file)      |
| `includeHidden`| checkbox | include dot-files                                   |
| `overwrite`    | checkbox | overwrite the output file if it already exists      |

Swap these for whatever your command actually needs.
# test-crate
