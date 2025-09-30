# VBA-Cleaner
Macro code cleaner for Excel workbooks - Reduce file-bloat and improve portability &amp; reliability
This is a modern, free alternative to legacy “CodeCleaner” utilities. Works on 32/64-bit Office, no VBIDE reference required (late binding), handles mixed-locale CodeNames (`Sheet1`/`Ark1`), avoids `VERSION/Attribute` clutter in document modules, and can avoids issues with OneDrive/SharePoint file paths.

## When/why should you "clean" a Visual Basic for Applications (VBA) enabled Microsoft Excel workbook?

VBA projects can accumulate 'cruft' (accumulated unnecessary leftovers, stale settings, old artifacts) that can bloat the file size, and odd behavior (such as compilation errors) when other users open your file. Here are some examples:

* **Stale compiled p-code**: the compiled stream inside modules can desync from source across Office versions or after heavy edits. VBA stores **source code** *and* compiled **p-code** inside the project. That p-code is **not guaranteed portable** across environments. Symptoms of incompatible p-code: odd compile/runtime errors, “Can’t find project or library,” or just flaky behavior that disappears after **Debug → Compile VBAProject**. You’ll need a recompile when, for example:
  * **Office/VBA runtime version changes** (e.g., opening a workbook compiled under Office X in Office Y).
  * **Different host/type library versions** (e.g., Excel/Word versions; or changing references—ADO 2.8 vs 6.1, different GUIDs/versions).
  * **32-bit vs 64-bit Office** (especially if `PtrSafe`/Declare signatures differ or API calls change).
  * **Locale/encoding & build differences** (rare, but it happens).
  * **Conditional compilation args** or project options change.
* **Fragmented module streams**: lots of edits can bloat text streams; forms may carry large, outdated `.frx` payloads. Each user-designed form has a **binary companion file** with the same base name and the extension **`.frx`**. That `.frx` holds the **binary resources** for the form: embedded images, control property blobs (fonts, colors, serialized states), icon handles, etc. In the workbook itself, those resources live inside the file; the `.frx` appears only when you **export** a form.
* **Workbook container cruft** (separate from VBA): orphaned names/styles, bloated UsedRange, old PivotCaches/QueryTables, OLE caches, etc. None of this is *inherently* bad; it’s the **excess, duplication, or orphaning** that hurts. A DeepClean (copy sheets to a new workbook) plus hygiene passes (reset UsedRange, prune styles/names, remove unused connections) typically trims most of it. Examples:
  * **Defined Names**: hidden or orphaned names; names with `#REF!`; thousands of stray names imported from other files.
  * **Styles**: style explosion (hundreds/thousands of “Normal 2”, “Normal 3”, …).
  * **UsedRange bloat**: Excel thinks a sheet spans to row 1,048,576 because a cell was once touched; inflates file size.
  * **Conditional Formatting**: duplicated/overlapping rules everywhere.
  * **PivotCaches / QueryTables / Connections**: old caches or connections not used anymore.
  * **Power Query / Data Model**: retained query results or model artifacts after structure changes.
  * **External links & OLE/ActiveX leftovers**: broken links, ghost controls, stale OLE caches.
  * **Shapes/Images**: hidden/off-sheet shapes, large images no longer needed.
  * **Printer setup & page breaks**: per-sheet printer settings bloating XML.
  * **Custom XML parts / Doc properties**: add-ins and tools leave behind custom parts you no longer need.
  * **Chart caches / Slicer caches**: redundant cache data after edits.
  * **Shared strings (xlsx)**: very large `sharedStrings.xml` due to past content.
<br>

---

## What’s included in this repositry?

Two tools are included in the same VBA module:

1. **"VBA_Cleaner"**: cleans an open VBA project **rebuilds modules from clean source** so p-code regenerates. 
2. **"VBA_DeepClean"**: goes further than VBA_Cleaner; it rebuilds the **workbook container** too, and saves the file without p-code, improving file portability.

## When to use which?

1. Run **VBA Cleaner**:
  * occasionally during development to keep p-code/source aligned and reduce mysterious compile issues.
  * when you want a fast rebuild of the VBA project to remove stale p-code/stream fragmentation without losing any global VBA-project settings or touching the workbook container.
2. Run **DeepClean**:
  * on bloated files, to minimize file size when the workbook feels “heavy”, or behaves strangely and wish to improve reliability by refreshing both the VBA project and the workbook container.
  * on inherited files, or before distributing to others, to clear incompatible p-code, and to facilitate a clean code-compilation.
<br>

---

## 1) Basic **VBA_Cleaner** (p-code cleaner)

* **Scope:** Current/open VBProject only.
* **What it does:**
  * Exports all components to temp.
  * Removes **non-document** components (Std modules, Classes, Forms) and re-imports them fresh (forms bring their `.frx`).
  * For **document modules** (`ThisWorkbook`, sheet/chartsheet modules), it **clears all lines** and **re-inserts only the code body** from the exported `.cls` (strips `VERSION/BEGIN…End` and `Attribute` lines so they don’t appear in the code window).
  * It doesn't meddle with global VBA-project settings such as Tools→References, project password, name, or compile settings.

| **Pros** | **Cons** |
| -------- | -------- |
| <ul><li>Fast and in-place (i.e. acts on an open workbook).</li><li>Retains all VBA-project settings such as project password and references.</li><li>Keeps workbook intact (names, styles, etc).</li></ul> | <ul><li>Doesn’t fix workbook/container bloat.</li><li>If forms are extremely large, you’ll still carry their `.frx` content (though exported/imported cleanly).</li></ul> |

### How to use VBA_Cleaner:

1. Open the workbook / VBProject you want to clean (unlock VBProject if it is password-protected).
2. Run macro `VBA_Cleaner`, e.g. via ALT+F8.
3. Pick the project (or press Enter for the active one).
4. It runs in-place on the open workbook, and shows a success message.

### What exactly happens under the hood of VBA_Cleaner?

* Exports every component from the workbook to a temp. folder, but keeps the workbook itself (ant all its global settings) as an empty shell.
* Removes **non-document** components, then imports them back.
* For **document** components, does not delete them (ensuring that important workbook settings are not lost), but the code modeule is deleted, and replaced with clean code.
<br>

---

## 2) **VBA_DeepClean** (Whole workbook rebuild)

* **Scope:** An **open workbook** (selected from a dialog picker).
* **What it does:**
  * Exports all components from source.
  * **Creates the destination by copying the first sheet to a new workbook** (prevents CodeName renumbering), then copies remaining sheets and chartsheets.
  * Imports **non-document** components.
  * For **document modules**, injects **only** the code body (header/attribute lines stripped) into the matching destination modules, matched by **tab name** → destination CodeName (robust across locales and renumbering).
  * Saves to a **local path** (auto-maps SharePoint/OneDrive URLs to local mirrors or falls back to `%TEMP%`) and **closes** the new file so your users can manually open it and trigger `Workbook_Open`/`Auto_Open`.
  * You get a **fresh workbook container**. That incidentally clears many forms of workbook bloat and odd metadata.
 
| Pros | Cons |
| ---- | ---- | 
| <ul><li>Cleans all both **VBA streams**, UserForms’ `.frx` binaries, and **workbook container**, reducing file size, and improving reliability.</li><li>Saves the file without p-code, improving portability</li><li>Robust matching of sheet/chartsheet code in mixed international locales (`Ark1`/`Sheet1`) and after renumbering.</li><li>Avoids `VERSION/Attribute` junk in code windows.</li><li>Uses **late binding** (no VBIDE reference needed)</ul></ul> | <ul><li>**Project password is cleared** (destination VBProject has no password; set it again if needed).</li><li>**Project name, conditional compilation args, and some project-level settings** (Break on Unhandled Errors, etc.) revert to defaults in the destination.</li><li>Tools→References **are not cloned**; they remain whatever Excel assigns by default for the new file. Re-set custom references if you had them.</li><li>Breakpoints, watches, code pane positions aren’t preserved (VBE limitations).</li><li>Copies sheets “as is”: if your workbook had excessive styles, hidden names, etc., many are mitigated by the new container, but sheet-local cruft (e.g., wildly expanded UsedRange) may persist unless you reset it separately.</li></ul> | 

### How to use VBA_DeepClean:

1. Open the source workbook (unlock VBProject if it is password-protected)
2. Run macro `VBA_DeepClean`, e.g. via ALT+F8.
3. Pick the workbook to deep-clean, in the dialog window that appears.
4. Select file name to save, in the SaveAs file browser that appears. By default the file name is appended with '_DeepClean.xlsm' 
5. The tool exports, rebuilds into a **new** workbook, **saves and closes** it without p-code.
6. If necessary re-set any **project password**, **project name**, and **references** you need.

### Exactly what happens udner to hood of VBA_DeepClean?

* Exports everything!
* Creates a destination by `firstSheet.Copy` to a **new workbook** (no placeholder sheet), then copies remaining sheets/charts back.
* Imports non-document components back.
* For `ThisWorkbook` and each tab (worksheet or chart) strips and re-imports all source code.
* Finally saves to a **local path** (OneDrive/SharePoint URL → local mirror), the closes file without compiling p-code.
<br>

---

## Known limitations

* Doesn’t “fix” broken **Tools→References** automatically (both tools). DeepClean’s new file may need custom references re-set.
* Doesn’t clone **VBE UI state** (breakpoints/watches).
* DeepClean resets **VBProject password** and **project-level settings** to defaults. If you rely on **VBProject passwords** or **specific project names** for external tooling—re-apply them after DeepClean or stick to the basic Cleaner.
* Saving directly to a SharePoint **URL** is avoided on purpose; we save to local disk for reliability.
* Do **not** use if you’re auditing a workbook under legal/forensic constraints where changing file hashes is undesirable.
* Do **not** use DeepClean if you expect references to auto-migrate. Set them explicitly.

---

## Requirements & Safety

* The code requires that you first set Excel **Options → Trust Center → Trust Center Settings… → Macro Settings →** check **“Trust access to the VBA project object model.”**. The macros check this automatically, and inform you when necessary.
* Unlock any password-protected project before running.
* Document module code is inserted with `AddFromString` after stripping header/attribute lines to avoid “`VERSION 1.0 CLASS …`” showing in the editor.
* This repositry is provided without warranty of any kind. To be sure, take backups.

---

## Credits & license

* Inspired by Rob Bovey’s canonical **VBA Code Cleaner** tool (RIP that 32-bit limitation).
* This project is original work, designed for 32/64-bit Office with late binding.
* Lisence: MIT (simple & permissive)

if you want, i can also generate a **short CHANGELOG**, a **CONTRIBUTING.md** (with guidance on PRs and common test cases), and a sample **demo workbook** you can include for users to test the tools on.
