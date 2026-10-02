# Office MCP

A local-first Model Context Protocol server for Microsoft Excel, Word, and PowerPoint files and applications.

This is an independent open-source project and is not affiliated with or endorsed by Microsoft.

Office MCP gives an orchestrating AI deterministic handles for Microsoft Office. It can work directly with `.xlsx`, `.docx`, and `.pptx` packages on any supported platform, or use the real installed Excel, Word, and PowerPoint applications on Windows for full-fidelity editing, calculation, charts, PivotTables, PDF export, and slide rendering. It runs locally over stdio and requires no API key.

The server does not decide what a report, model, or presentation should say. The model plans the work and calls these tools to inspect inputs, construct the documents, render the result, visually review it, and iterate.

## Quick start (Windows PowerShell)

You need [Node.js 20+](https://nodejs.org/), Git, and an MCP client such as Codex/ChatGPT desktop or Claude Code. Desktop Microsoft Office is **optional** for basic `.xlsx`, `.docx`, and `.pptx` file creation; it is required for Windows-native features such as animations and PowerPoint rendering. This repository is not yet published on npm.

1. Download and build the server:

   ```powershell
   git clone https://github.com/gauravvvvvvvvvv/office-mcp.git
   Set-Location office-mcp
   npm ci
   npm run build
   ```

2. In the **same PowerShell window**, choose which files the AI may access and connect **one** client. This example allows your Documents folder; change `$allowed` to a narrower folder if you prefer:

   ```powershell
   $entry = (Resolve-Path .\dist\index.js).Path
   $allowed = (Resolve-Path ([Environment]::GetFolderPath('MyDocuments'))).Path
   ```

   For **Codex CLI or ChatGPT desktop**, run:

   ```powershell
   codex mcp add office --env "OFFICE_MCP_ROOTS=$allowed" -- node $entry
   codex mcp list
   ```

   For **Claude Code**, run instead:

   ```powershell
   claude mcp add --scope user --env "OFFICE_MCP_ROOTS=$allowed" --transport stdio office -- node $entry
   claude mcp get office
   ```

3. Restart an already-open desktop client or start a new chat. Ask: “Call `office_capabilities` and tell me which Office features are available.” Then try: “Use `powerpoint_create_designed_presentation` to create a two-slide presentation in my Documents folder, inspect it, and report the saved path.” The server runs when the client launches it; do **not** start `node dist/index.js` in a separate terminal.

If you use both clients, run both registration commands. Existing `office` entries should be checked before re-adding them; `codex mcp list` and `claude mcp get office` show what is already configured. The checkout must remain at this path, because the clients launch its compiled `dist/index.js`. See [other platforms and clients](#other-platforms-and-clients) or [troubleshooting](#installation-troubleshooting) below.

## Why this server

- Hybrid operation: portable Office-file creation and inspection without desktop Office, plus native Windows Office operations when installed.
- Prompt-specific quality contracts, file comparisons, render evidence, and a finalization gate that refuses stale or incomplete review.
- PowerPoint generation that stays in the background by default; it does not repeatedly steal focus from the user's work.
- Bounded filesystem roots, explicit overwrite behavior, and no arbitrary COM execution tool.

These are design choices, not proof that this project is more capable or produces better-looking documents than every alternative. Its 34-tool surface is intentionally bounded; a human-quality result still depends on the orchestrating model, source material, and visual review.

## What works

| Application | Operations |
| --- | --- |
| Excel | Create and inspect workbooks; read/write/format ranges; descriptive analysis and read-only filtered/grouped queries; formulas and native recalculation; tables, charts, sorting, filters, named ranges, conditional formatting, validation, and PivotTables |
| Word | Create and inspect documents; sections and page layout; headers, footers, page numbers, TOC fields, source-backed citation text, headings, lists, tables, images, links, footnotes, replacement, native comments and tracked changes, and PDF export |
| PowerPoint | Create, inspect, and audit decks; reuse existing PPTX masters/layouts or duplicate styled slides; text, shapes, images, media, tables, charts, notes; native animation timeline editing, transitions, PDF export, and PNG rendering |
| Common | Discover files in directories, list active Office files, inspect and compare supported files, export PDF, run prompt-specific quality checks, prepare render-backed review evidence, and finalize only reviewed drafts |

The server exposes 34 MCP tools:

- `office_capabilities`
- `office_inspect_file`
- `office_search_files`
- `office_convert_to_pdf`
- `office_export_pdf`
- `office_quality_check`
- `office_compare_files`
- `office_prepare_review`
- `office_finalize`
- `office_get_review_image`
- `excel_create_workbook`
- `excel_inspect_workbook`
- `excel_read_range`
- `excel_analyze_dataset`
- `excel_query_dataset`
- `excel_write_range`
- `excel_set_formulas`
- `word_create_document`
- `word_inspect_document`
- `word_append_paragraph`
- `word_replace_text`
- `powerpoint_create_presentation`
- `powerpoint_create_designed_presentation`
- `powerpoint_inspect_presentation`
- `powerpoint_audit_presentation`
- `powerpoint_inspect_template`
- `powerpoint_create_from_template`
- `powerpoint_replace_text`
- `office_native_status`
- `office_list_open_files`
- `excel_native_batch`
- `word_native_batch`
- `powerpoint_native_batch`
- `powerpoint_render_presentation`

## Requirements

- Node.js 20 or newer
- Microsoft Office on Windows for native automation, active-document control, and PowerPoint rendering
- LibreOffice only for portable PDF conversion when native Office is unavailable

Microsoft Office does not need to be installed for the portable file operations.

### Platform compatibility

| Platform | Portable `.xlsx` / `.docx` / `.pptx` tools | Native Office automation and animations | PDF / visual review |
| --- | --- | --- | --- |
| Windows | Supported with Node.js | Requires installed desktop Microsoft Office | PowerPoint slide PNGs require desktop PowerPoint; PDF export can use Office or LibreOffice |
| macOS and Linux | Designed to work with Node.js; verify in your environment | Not supported by the Windows COM bridge | PDF conversion needs LibreOffice; PowerPoint slide PNG review is unavailable |

The portable path edits document files, whether or not they are open in an Office application. It does not control an open Office window. Only Windows native automation can operate on the active document, add PowerPoint animations, use Excel's calculation engine, or render slides through PowerPoint. This is a local stdio server, not a hosted service: its client must be able to launch the Node.js process and access the allowed files.

## Verify a development checkout

These commands are for contributors and maintainers, **not** required to install the server:

```powershell
npm run check
npm run smoke
```

`npm run check` type-checks, runs the portable integration tests, and builds `dist/`. `npm run smoke` launches the compiled server through an MCP client and creates a test deck. On Windows with desktop Microsoft Office installed, `npm run test:native` additionally tests native Excel, Word, and PowerPoint behavior. Do not run that native test on a machine without Office.

## One-prompt workflow

For a request such as “analyze the financial workbook in this folder and create a board presentation,” an orchestrator can:

1. Call `office_search_files` to discover workbooks and templates.
2. Call `office_inspect_file` or `excel_native_batch` to understand data, formulas, and workbook structure.
3. Call `excel_analyze_dataset` for deterministic data-quality statistics and `excel_query_dataset` for focused filters, group-by summaries, and sorting; then reason about the business question and plan the story.
4. Call `powerpoint_create_designed_presentation` for a new narrative deck, `powerpoint_create_presentation` for a custom layout, or `powerpoint_inspect_template` followed by `powerpoint_create_from_template` to preserve a company deck's layouts.
5. If motion is requested, call `powerpoint_native_batch` to add object animations and slide transitions, then save and reopen to inspect the timeline.
6. Define a quality contract from the actual prompt, including required text, structure, numerical assertions, and—when editing—an optional baseline source file. Call `office_quality_check`; repair blocking issues.
7. Call `office_prepare_review` to render every slide or export a Word/Excel PDF, and inspect the result. For a long deck, request slides beyond the inline preview with `office_get_review_image`.
8. Visually critique alignment, overflow, hierarchy, and chart readability; apply corrections and rerun the check and render. PNG renders cannot demonstrate motion.
9. Call `office_finalize` with honest pass/fail notes for every slide, sheet, or document and every prompt criterion. It verifies the draft, contract, baseline, render evidence, and review coverage again, then copies to a separate final file without overwriting.

This is the intended separation: the AI supplies judgment; Office MCP supplies reliable handles.

### Quality gate and honest limits

`office_quality_check` validates ZIP package relationships, checks explicit prompt assertions, and returns the draft's SHA-256 hash and review units. It can check Excel sheets, headers, exact cells, and formula errors; Word headings, tables, and image descriptions; and PowerPoint slide counts, minimum font sizes, speaker notes, and structural audit flags. A `baseline.path` compares an edited file to its source and blocks removed sheets, slides, or Word paragraphs by default. Contract thresholds can explicitly permit intended removal or bound changed Excel cells. `office_compare_files` returns the detailed semantic difference independently of finalization.

`office_prepare_review` creates hashed slide PNGs or a PDF and a manifest tied to the exact draft, contract, and baseline. `office_finalize` refuses failed machine checks, a stale or incomplete render, a changed draft, unreviewed units, failed criteria, or unaccepted warnings. The server cannot determine whether a claim is true, whether a chart tells the right story, whether a design is beautiful, or whether a reviewer genuinely inspected the output. The orchestrator must use the visual evidence and source data honestly; don't mark a criterion passed if it was not checked. Native rendering and PDF export need installed Office or LibreOffice, and PowerPoint slide PNG review specifically needs desktop PowerPoint on Windows.

### Quiet background generation

Portable PPTX creation edits files directly and never opens PowerPoint. Native file-based PowerPoint operations now create or open presentations with hidden windows by default, including rendering, animation edits, and PDF export. They do not repeatedly bring a deck to the foreground.

PowerPoint automation uses a shared application instance on Windows. To avoid altering or closing a presentation the user is actively working in, the server refuses file-based native work while PowerPoint is already running. The agent should continue with portable creation and defer native rendering or animation work until PowerPoint is closed. `target.active: true` explicitly opts into the running session, and `target.visible: true` requests a visible window; neither should be used without the user's consent. Native rendering does not expose an active-session override.

### Improving presentation quality

`powerpoint_create_designed_presentation` provides five widescreen compositions (`cover`, `imageText`, `statement`, `comparison`, `panorama`) and two visual styles (`editorial`, `cinematic`). `panorama` pairs a wide scene with an editable title and supporting copy; `comparison` accepts an optional `takeaway` to state the conclusion. It uses readable typography and crops supplied images to their frames. Supply relevant, distinct image assets when they help the story; the tool does not generate images or invent content. The lower-level PowerPoint tool remains available for charts, tables, and exact placement. Its image elements support `fit: "cover"`, `"contain"`, or `"stretch"`.

The optional companion Codex skill lives at `skills/office-presentation-quality/SKILL.md`. Install that folder into your Codex skills directory to guide the agent through planning, render review, and revision. This skill improves the workflow; it cannot guarantee visual taste or replace a good reference deck. To use the updated server and skill in an existing Codex session, restart that session after rebuilding `dist/`.

`powerpoint_audit_presentation` works directly on any PPTX without opening Office. It reports off-slide objects, small explicitly sized text, unusually dense slides, pictures without meaningful alt text, low contrast where both colors are explicit solids, overlapping text boxes, empty slides, and grouped objects it cannot measure reliably. These are review flags, not proof of poor design. The creation tools accept image alt text (`altText` for custom images, `imageAltText` for designed layouts); without it, they use a reviewable placeholder rather than leaking a local image path into the deck.

### Reusing a PowerPoint template

`powerpoint_inspect_template` opens an existing `.pptx` invisibly through desktop PowerPoint and reports its designs, custom layouts, placeholder indexes, and slide shape names. `powerpoint_create_from_template` can then add slides from those layouts or duplicate existing slides, fill placeholders or named text shapes, add text boxes, and save a separate `.pptx`. It keeps the source deck untouched and removes its original slides from the output unless `keepTemplateSlides` is true. Coordinates for new native text boxes are points. This requires Windows desktop PowerPoint and is subject to the quiet-background rule above; it does not import `.potx` files or offer arbitrary master editing yet.

### Word structure and Excel analysis

`word_create_document` accepts either the original `blocks` array or a `sections` array with separate page size/orientation, margins, headers, footers, and page numbers. Blocks can include an updateable `toc`, a `cited_paragraph` referring to user-supplied `sources`, a `bibliography`, images with required alt text or an explicit decorative flag, hyperlinks, page breaks, and footnotes. Word fills the TOC entries when it updates the field. Citations are formatted document text, not entries in Word's built-in citation manager. Native Word batches can apply styles, insert sections, generate or refresh TOCs, and add comments to an existing document.

`excel_analyze_dataset` profiles a sheet or selected range for missing and duplicate rows, column types, descriptive numeric statistics, IQR outliers, top text values, dates, formula/error counts, and strongest Pearson correlations. It analyzes at most 50,000 rows and 100 columns per call. It does not clean the data, infer causation, run statistical tests, or calculate uncached formulas; use native Excel recalculation first when formula results matter.

`excel_query_dataset` reads up to 50,000 rows and 100 columns, applies typed filters, optionally groups and aggregates (`count`, `countDistinct`, `sum`, `average`, `min`, `max`), sorts, and returns at most 1,000 rows. It never writes back. Native Excel batches can define names, add conditional formatting, and set data validation.

### Animations and transitions

On Windows with installed desktop PowerPoint, `powerpoint_native_batch` can apply object entrance (`appear`, `fade`, `fly`, `wipe`, `zoom`), emphasis (`spin`, `growShrink`), and exit (`fade`, `fly`, `wipe`, `zoom`) effects. Triggers are `onClick`, `withPrevious`, and `afterPrevious`; duration and delay use seconds. Call `list_shapes` to get a shape's name or 1-based index before targeting it. `rename_shape`, `group_shapes`, `ungroup_shape`, `set_z_order`, and `update_shape` provide the object controls needed for cinematic sequences. `list_animations`, `update_animation`, `move_animation`, `delete_animation`, and `clear_animations` support timeline inspection and revision, including repeat and reverse settings. `add_media` embeds a local audio or video asset. `set_transition` supports `none`, `cut`, `fade`, left/right push and wipe, zoom-in, and true PowerPoint Morph by object, word, or character. For deterministic Morph matching, give paired shapes the same name beginning with `!!`, duplicate the slide, then move or resize those shapes on the duplicate. `export_video` creates an MP4 or WMV and waits for PowerPoint to finish encoding it. Include a `save` operation and reopen the result to verify persistence.

This is an expanding, validated Office capability set, not a wrapper around every Office command. Arbitrary COM execution is intentionally not exposed. Animation playback itself is not captured by the static PNG renderer.

## Other platforms and clients

The [quick start](#quick-start-windows-powershell) is the shortest route for Windows. The same built server works with any local stdio MCP client that can launch Node.js. Client configuration is separate: adding Office MCP to Codex does not automatically add it to Claude Code.

### Codex and ChatGPT desktop

The `codex mcp add` command in the quick start writes Codex's user configuration, which the ChatGPT desktop app, Codex CLI, and Codex IDE extension share. In ChatGPT desktop, restart after adding the server and type `/mcp` to inspect it. You can also add it through **Settings → MCP servers → Add server → STDIO**. ChatGPT web does **not** read local Codex configuration; this repository does not provide a hosted MCP endpoint. See [OpenAI's MCP guide](https://learn.chatgpt.com/docs/extend/mcp?surface=app).

### Claude Code

The quick start uses Claude Code's `user` scope, so the server is available in all your local projects. Run `claude mcp list` or use `/mcp` inside Claude Code to check its connection. See [Claude Code's MCP setup guide](https://code.claude.com/docs/en/mcp).

### macOS or Linux

Portable `.xlsx`, `.docx`, and `.pptx` tools are designed for Node.js on these platforms, but this project's native Office bridge and PowerPoint PNG rendering are Windows-only. After cloning and running `npm ci` and `npm run build`, run the following from the repository root in a Bash-compatible shell; choose **one** client command:

```bash
entry="$(pwd -P)/dist/index.js"
allowed="$(pwd -P)" # Change to a directory containing the documents you want to use.
codex mcp add office --env "OFFICE_MCP_ROOTS=$allowed" -- node "$entry"
# Or: claude mcp add --scope user --env "OFFICE_MCP_ROOTS=$allowed" --transport stdio office -- node "$entry"
```

The default example grants access only to the repository. Set `allowed` to an absolute documents directory if your files live elsewhere. Restart the client, check its MCP connection, and call `office_capabilities`.

### Other local MCP clients

Configure a stdio server with `node` as the command, the absolute path to `dist/index.js` as its argument, and `OFFICE_MCP_ROOTS` as an environment variable. For example, adapt this JSON to your client's format (the paths shown are placeholders):

```json
{
  "command": "node",
  "args": ["C:\\absolute\\path\\to\\office-mcp\\dist\\index.js"],
  "env": { "OFFICE_MCP_ROOTS": "C:\\allowed\\documents" }
}
```

### Installation troubleshooting

| Symptom | Check |
| --- | --- |
| `node` or `npm` is not recognized | Install Node.js 20+ and open a new terminal; run `node --version` and `npm --version`. |
| `codex` or `claude` is not recognized | Install the corresponding client CLI, or use that client's graphical MCP settings. |
| Client says the server cannot start | Run `npm run build`; check that `dist/index.js` exists and that the configured path still points to this checkout. Do not expect a standalone `node dist/index.js` process to print a success message; it waits for MCP input. |
| `office` already exists | Inspect the existing entry with `codex mcp list` or `claude mcp get office`; don't create a duplicate. If you intend to replace it, remove the old entry in that same client and rerun the setup command. |
| A file is rejected as outside allowed roots | Set `OFFICE_MCP_ROOTS` to an absolute parent folder containing that file, then restart the client. Multiple Windows roots use `;`; macOS/Linux roots use `:`. |
| Native animations, Excel recalculation, or PowerPoint rendering are unavailable | These require Windows and installed desktop Microsoft Office. Start by calling `office_capabilities` and `office_native_status`. |
| ChatGPT web cannot see the server | Web chats do not load local stdio MCP configuration; use ChatGPT desktop/Codex or build a separately hosted integration. |

## Path security

By default, the server may only access files below its working directory. Set `OFFICE_MCP_ROOTS` to one or more allowed roots.

On Windows, separate multiple roots with a semicolon:

```text
C:\Users\your-name\Documents;D:\Company Files
```

On macOS and Linux, separate roots with a colon. The server resolves every input and output path before use and rejects path traversal outside these roots.

## PDF conversion

`office_convert_to_pdf` uses LibreOffice in headless mode. The server checks standard installation locations and the system `PATH`. For a custom installation, set:

```text
OFFICE_MCP_SOFFICE=C:\path\to\LibreOffice\program\soffice.exe
```

Call `office_capabilities` to see whether the converter was detected.

## Development

```powershell
npm run dev
```

The process waits silently for MCP messages on standard input. Do not write logs to standard output because stdout carries the MCP protocol. Diagnostics belong on stderr.

## Manual releases

This repository has no GitHub Actions workflows or Dependabot configuration. Nothing runs automatically on pushes or pull requests, and pushing a tag does not create a release or publish to npm. It is a local stdio server, not a hosted Office service.

Before making a public release manually, update `package.json` and the lockfile; run `npm run check` and `npm run smoke`; and, on Windows with desktop Office, run `npm run test:native`. Verify a fresh packaged install and audit the bundled dependencies. Then decide whether to create a GitHub Release and attach the tested package. See [CONTRIBUTING.md](CONTRIBUTING.md) and [SECURITY.md](SECURITY.md). No npm publication is currently configured.

Project layout:

```text
src/
  index.ts          MCP tool registration and schemas
  office.ts         Cross-format inspection and PDF conversion
  excel.ts          XLSX adapter
  word.ts           DOCX adapter
  powerpoint.ts     PPTX adapter
  paths.ts          Allowed-root enforcement
  errors.ts         Structured MCP results and errors
  native.ts         Secure Node-to-native-Office bridge
native/
  office-bridge.ps1 Excel, Word, and PowerPoint COM automation
tests/
  office.test.ts    End-to-end file tests
```

## Important limitations

- Native active-document control currently requires Windows and installed desktop Microsoft Office. A future Office.js bridge is required for the same live control on Mac, web, and iPad.
- ExcelJS does not support every Excel feature. Re-saving complex workbooks may not preserve unsupported objects such as embedded charts or VBA projects. Prefer a separate `outputPath` when editing valuable files.
- Word and PowerPoint replacement can match text split across formatting runs. When a replacement crosses multiple runs, it consolidates the replacement into the first run, so mixed character formatting inside that phrase is not retained.
- PowerPoint creation is declarative. Positions and sizes use inches, matching PptxGenJS conventions.
- The designed layouts are starting points. Template reuse is available through native PowerPoint but does not import `.potx`, edit a master arbitrarily, or automatically judge aesthetics. Inspect renders before delivery.
- The PPTX audit cannot determine whether text actually fits, whether contrast inherited from a theme or image background is sufficient, or whether the story and aesthetics are strong. Visually inspect rendered slides.
- Object animation and slide-transition editing currently requires desktop PowerPoint on Windows. Static slide images verify visual layout, not how effects play during a slideshow.
- Background native PowerPoint work can be deferred when PowerPoint is already running. Hidden presentations avoid pop-up windows, but this is not a fully isolated headless PowerPoint service.
- Formula values are written but not calculated by Node.js. The workbook is marked for recalculation when opened in Excel.
- The server does not offer every command in Microsoft Office, nor does it act as an autonomous design model. Quality still depends on the orchestrating model, source assets, templates, and render review.

## Technology

- [MCP TypeScript SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- ExcelJS for `.xlsx`
- docx for document creation
- PptxGenJS for presentation creation
- JSZip for structure-preserving OOXML inspection and targeted edits

## License

MIT
