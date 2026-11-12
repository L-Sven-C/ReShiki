# macOS Word integration test report

Run date: 2026-11-12 (Asia/Singapore). **Phase 1 acceptance is incomplete.**
A real Word task pane loaded, inserted a native drawing and acknowledged native
saves. The saved DOCX contains one displayed image and the matching native data.
The document was closed and reopened; a new editing session restored identical
native bytes. A subsequent changed save and the remaining failure/isolation cases must still pass before this is described as reliable Mac Word support.

## Build and environment

| Item | Observed value |
| --- | --- |
| Upstream baseline | `d254c8d8cd686d02893ed1283f528ac5d98b9177` |
| Native fix / tested app source | `3f6b43d008bbb5f8cee0aaec2fab27b39c050b80` |
| Fork branch | `L-Sven-C/ReShiki`, `feature/macos-word-integration` |
| Host | Apple Silicon, macOS 27.2, build 26B123 |
| Word | Microsoft Word for Mac 16.100.4, Mac App Store installation |
| Rust | rustup 1.99.1, rustfmt and Clippy; repository toolchain pin |
| Node | 26.12.1 |
| Xcode | 26.6, build 17F113 |
| Python | 3.14.6 |
| Development app | `target/debug/ReShiki.app`, built and ad hoc signed by upstream packager |
| Manifest | Generated Word TaskPaneApp, WordApi 1.4, localhost HTTPS port 43127 |
| Test files | Local disposable DOCX, upstream ethanol fixture; no research documents |

The host run used network access for Microsoft's Office.js runtime. Cold-start
offline use and other Word/macOS versions are not validated. Word's launch URL
reports `16.01` as protocol metadata; that is not the installed Word version.

## Automated checks

Full local logs are under `~/Library/Logs/ReShiki-word-dev/`; they are not
committed because build logs contain machine paths. Commands run at repository
root unless stated otherwise. Prefix Rust commands with
`PATH="$HOME/.cargo/bin:$PATH"` when Homebrew Rust comes first.

| Check | Actual result | Evidence log |
| --- | --- | --- |
| Initial `(cd integrations/office && node --test)` | Failed: missing pinned `@xmldom/xmldom` dependency in five test files | `node-baseline.log` |
| Same Node suite after installing pinned root dev tools | **158 passed, 0 failed** | `node-baseline-ready.log` |
| Initial Rust Office tests with Homebrew Rust 1.96 | Compile failed on upstream `isolate_lowest_one`; toolchain too old | `rust-office-baseline.log` |
| `cargo test --locked -p reshiki --lib office_addin` with 1.99.1, baseline | **5 passed** | `rust-office-baseline-1.99.log` |
| New regressions before fix | **2 failed, 1 passed**, reproducing draft loss | `recovery-red.log` |
| Same Office tests after fix | **8 passed, 0 failed**, 388 unrelated tests filtered out | `recovery-green.log` |
| `cargo fmt --all --check` | Passed | `fmt-final.log` |
| `cargo check --workspace --all-targets --locked` | Passed | `cargo-check.log` |
| `cargo clippy --workspace --all-targets --locked -- -D warnings` | Passed workspace gate | `clippy.log` |
| `node_modules/.bin/oxlint --deny-warnings website integrations/office packages/editor` | Passed | `lint-baseline.log` |
| `python3 scripts/build_macos_app.py` | Built runnable native preview worker and app bundle; signing completed | `build-app.log` |
| `node scripts/check_docs_tree.mjs` | Passed, 902 referenced assets | `docs-tree.log` |
| `npx --yes office-addin-manifest validate integrations/office/manifests/word.xml` | Microsoft manifest validator: valid | `manifest-validation.log` |

Upstream dependency warnings remain (vendored iced deprecations and future
incompatibility of `block` 0.1.6). The debug linker warned about a large unwind
section. These are not suppressed or described as a warning-free build.
The full Rust chemistry/integration suite, Bun formatter, website build and
website checks were not run; the Office and native checks above do not imply
those passed. No JavaScript runtime logic was modified in this patch.

The three new native regression tests exercise missing, malformed, oversized,
unsupported-version and invalid-identity session metadata; symlink metadata;
and failed atomic draft replacement. They establish that the latest bytes are
retained before failed authorization, no new update request is published, and a
failed disk write cannot claim a saved recovery draft. They are real filesystem
tests, **not Office host tests**.

## Real-host procedure and observed results

All rows below use the OS, Word build and app commit in the environment table.
Evidence is retained locally under
`~/Library/Application Support/ReShiki-Office-Dev/evidence/`.
The working file is `ReShiki-Word-acceptance-20261112.docx`.
Neither DOCX files, certificate keys, session capabilities nor recovery drawings
are included in the Git commit.

| Case and reproducible steps | Actual result / evidence | Status |
| --- | --- | --- |
| Install trusted localhost certificate, run setup and companion; inspect listener and fetch taskpane without TLS bypass | User completed Keychain authorization; certificate verification and HTTPS fetch succeeded; listener restricted to `127.0.0.1:43127` | Pass |
| Copy generated manifest into Word `wef`, restart, Home → Add-ins → Developer Add-ins → ReShiki | ReShiki appeared. Automated entry clicks had no effect; user clicked it and actual pane displayed **Connected to Word**. This was an automation obstacle, not established add-in incompatibility | Pass, manual activation |
| With no editable ReShiki clipboard payload, choose Paste editable drawing | Explicit “Copy an editable drawing in ReShiki…” error; no guessed chemistry was inserted. Message currently includes a confusing “LibreOffice integration failed” worker prefix | Pass for refusal; wording issue remains |
| Attempt insertion with existing ReShiki object still selected | Explicit cursor-outside-drawing error; no extra image created. Moving the cursor outside allowed the second insertion | Pass |
| Insert `ethanol.rsk` using Insert a drawing file | Native file selected in actual Word file picker; tagged rich-text control and inline structure appeared. Initial persisted native hash equals fixture preview hash `7c5087f6e96abaedc71261ce0e174d089ddaa78648a500c36897ad6695bb899c` | Pass |
| Select inserted object, Edit selected drawing | Actual native `drawing.rsk` session opened. First loaded view was not captured before user edits, so exact initial native fidelity is supported by the stored record, not a first-frame screenshot | Partial |
| Modify in native editor and Save; verify original object updates | User modified test drawing. Native screenshot showed six-atom carboxylic-acid drawing and “Drawing updated in Office”; Word pane independently confirmed update. Matching request/ACK and saved native hash `f115e90c365874df9f62b867f401ab971b5898d33336d0a2d5a1ffaf0b6b332c` were verified | Pass for observed saves |
| Save DOCX, inspect package for one image and matching current native record | Actual Word Save completed. `01-after-native-update.docx` and `01-after-native-update.json`: **1 image, 1 logical object, 3 immutable versions**, current tag references latest record, native SHA-256 matches revision. Same control ID `-1687056187` retained | Pass |
| Close DOCX, reopen it, select same drawing, Edit, change and Save again | Save/close/reopen and Edit completed. Pane reconnected automatically; new session draft matches the previous native bytes exactly (`02-reopen-native.json`). Changed save in that session is still pending; multiple same-name native processes complicate UI automation | Partial; changed Save **Pending** |
| Insert second, distinct molecule; record both payload/image hashes; edit first; verify second unchanged | Second molecule (R-lactic-acid fixture) inserted and saved. `03-two-objects-before-save.json` / `.docx` show two images and distinct logical IDs; per-object native and PNG hashes recorded. First-object changed save still pending | Partial; isolation assertion **Pending** |
| Start editing first molecule; change Word selection to second before native Save; verify captured first target only changes | First edit remains open; second image selected in actual Word (image focus and Picture Format tab observed). Awaiting changed native Save to compare hashes | Partial; update assertion **Pending** |
| Copy/paste one object in same document, edit copy, compare logical IDs and original data | Not completed | **Pending** |
| Resize and move inline drawing, edit with same extent and changed extent; compare dimensions and paragraph placement | Not completed; current algorithm preserves scale, not necessarily fixed frame dimensions | **Pending** |
| Undo and redo a native update in Word; reopen editor for each restored state | Not completed | **Pending** |
| Enable AutoSave on an explicitly approved test location, update and reopen after save completes | Not completed; current test DOCX is local with AutoSave off. No test document uploaded | **Pending** |
| Protect/delete captured target or force Word update failure; Save in native app; verify old image/data and retained new draft, no false success | Not completed in real Word; automated rejection/rollback and native receipt tests pass | **Pending** |
| Close pane during edit; change and Save; verify no ACK, dirty editor and recoverable draft; explicitly reimport draft | Not completed | **Pending** |
| Exit companion/native/Word abnormally; restart; inspect drafts and document before explicit recovery | Not completed | **Pending** |
| Simulate delayed document synchronization and verify no ACK is mistaken for durable disk/cloud save | Not completed | **Pending** |
| Insert complete synthesis route and complex stereochemical drawing; repeat update/reopen checks | Not completed; single small drawing does not validate route or macrocycle layout fidelity | **Pending** |

Temporary Word runtime logging was enabled using Microsoft’s documented
`CEFRuntimeLoggingFile` preference while diagnosing startup, then the preference
was removed. No matching runtime log file was produced.

The first round trip included user interaction between automated observations.
The report distinguishes observed state, persisted evidence and missing steps;
it does not claim unattended end-to-end automation or complete acceptance.

## Remaining priorities

1. Finish save/close/reopen/re-edit and the two-object/selection-change tests
   before changing the UI or PNG format. Capture real `getOoxml` output only from
   disposable test drawings if a live-host failure needs a regression fixture.
2. Verify native-tab closure releases the reservation: an observed native window
   remained open with an Untitled tab after the drawing tab closed. Process-exit
   cleanup alone may require the pane's explicit End edit session in this case.
3. Run real failure/rollback/recovery, duplicate, undo/redo and resize cases.
4. Correct the reused LibreOffice error prefix for Office worker failures, and
   investigate clearer native tab titles (upstream issue #1119).
5. Only after the acceptance gate, consider ribbon actions and local Office.js
   packaging, then vector display/print fidelity. No Windows-style OLE or
   double-click support is assumed for Mac Word.
