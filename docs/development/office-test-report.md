# macOS Word integration test report

Run date: 2026-11-12 (Asia/Singapore). **Phase 1 acceptance is incomplete.**
Real Word tests now cover repeat editing, independent copied objects, undo/redo,
inline resizing, changed drawing extent, movement between paragraphs, rejected
updates and explicit recovery. The independent Word ribbon and its native
editing path were also tested; details and upgrade limitations are recorded below.
A native-tab lifecycle defect was reproduced and fixed: closing a drawing tab
while the app stayed open previously left its object reserved. The patched app
released the reservation and reopened the same object with identical native data.
AutoSave, synchronization delays, abnormal exits, floating placement and
partial host-mutation rollback remain unvalidated; this is not complete acceptance.

## Build and environment

| Item | Observed value |
| --- | --- |
| Upstream baseline | `d254c8d8cd686d02893ed1283f528ac5d98b9177` |
| Initial native app source, cases 01–20 | `3f6b43d008bbb5f8cee0aaec2fab27b39c050b80` |
| Initial companion source | `642055a55f82fbc1d4681cea81301838449029cf` (unchanged upstream runtime) |
| Tab-close fix / rebuilt app and companion, cases 21–25 | `008548d706b3e9995b1d9790ebc8db73bce8e6a8` |
| Ribbon add-in / companion, cases 26–28 | `174665b7` (native executable remains `008548d7`) |
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
those passed. The tab-close correction modifies companion session lifecycle logic.
The original manifest and Word mutation adapter are unchanged.

The continuation run added **6 Node tests and 4 Rust tests**. The initial five
Node cases failed before the lifecycle correction (`tab-close-red.log`), then
the full suite passed. These checks are separate from the real-host cases below.

| Continuation check | Actual result | Evidence log |
| --- | --- | --- |
| `(cd integrations/office && node --test)` | **164 passed, 0 failed** | `tab-close-node.log` |
| `cargo test --lib office_addin` | **12 passed, 0 failed**, 388 other root-library tests filtered; other workspace library tests also filtered | `tab-close-rust.log` |
| `cargo fmt --all --check`, `git diff --check` | Passed | terminal result |
| `cargo clippy --workspace --all-targets --locked -- -D warnings` | Passed, including native binding call site | `tab-close-clippy.log` |
| `node_modules/.bin/oxlint --deny-warnings website integrations/office packages/editor` | Passed | `tab-close-lint.log` |
| `node scripts/check_docs_tree.mjs` | Passed, 902 referenced assets | `tab-close-docs.log` |
| `python3 scripts/build_macos_app.py` | Built and signed patched app; exercised in real Word cases 21–25 | `tab-close-build.log` |

New coverage checks closure with no process exit, reopen before the next poll,
unread/pending save barriers, late old-process exit, rejected-save draft retention,
invalid/malformed/oversized/symlink close notices and idempotent native detachment.
Only the companion closes authorization; no close notice can fabricate a save ACK.

The three new native regression tests exercise missing, malformed, oversized,
unsupported-version and invalid-identity session metadata; symlink metadata;
and failed atomic draft replacement. They establish that the latest bytes are
retained before failed authorization, no new update request is published, and a
failed disk write cannot claim a saved recovery draft. They are real filesystem
tests, **not Office host tests**.

## Real-host procedure and observed results

All rows below use the OS and Word build in the environment table. Cases 01–20
use the initial app; cases 21–25 use the rebuilt fix.
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
| Insert `ethanol.rsk` using Insert a drawing file | Native file selected in actual Word file picker; tagged rich-text control and inline structure appeared. Initial persisted native hash equals the fixture native hash `7c5087f6e96abaedc71261ce0e174d089ddaa78648a500c36897ad6695bb899c` | Pass |
| Select inserted object, Edit selected drawing | Actual native `drawing.rsk` session opened. First loaded view was not captured before user edits, so exact initial native fidelity is supported by the stored record, not a first-frame screenshot | Partial |
| Modify in native editor and Save; verify original object updates | User modified test drawing. Native screenshot showed six-atom carboxylic-acid drawing and “Drawing updated in Office”; Word pane independently confirmed update. Matching request/ACK and saved native hash `f115e90c365874df9f62b867f401ab971b5898d33336d0a2d5a1ffaf0b6b332c` were verified | Pass for observed saves |
| Save DOCX, inspect package for one image and matching current native record | Actual Word Save completed. `01-after-native-update.docx` and `01-after-native-update.json`: **1 image, 1 logical object, 3 immutable versions**, current tag references latest record, native SHA-256 matches revision. Same control ID `-1687056187` retained | Pass |
| Close DOCX, reopen it, select same drawing, Edit, change and Save again | Save/close/reopen and Edit completed; pane reconnected and draft bytes matched (`02-reopen-native.json`). User requested restarting the edit. The restarted session loaded the same bytes, rotated the drawing 180°, saved with a matching ACK and persisted revision `5c013c75d807c4e3d7b50b55ffe514c36a1bbd41e6de5c07e256971dde16e070`. `04-restarted-before-save` and `05-after-restarted-save` JSON/DOCX snapshots plus `05-isolation-result.json` record the result | Pass for observed repeat edit |
| Insert second, distinct molecule; record both payload/image hashes; edit first; verify second unchanged | R-lactic-acid fixture inserted; a third drawing appeared during user interaction between observations. The 04/05 snapshots show three pictures before and after, unchanged first control/logical IDs, changed first native and PNG, and identical native/preview/displayed-media hashes and dimensions for both other objects | Pass for observed isolation |
| Start editing first molecule; change Word selection to second before native Save; verify captured first target only changes | Second image selected before native Save; Word still reported that image focused after Save. Exact request/session/revision ACK matched the first object’s new draft and persisted data; `05-isolation-result.json` confirms the other objects unchanged | Pass |
| Copy/paste one object in same document, edit copy, compare logical IDs and original data | Full-control copy forks an independent ID on Edit; editing copy and later original leaves the other unchanged. Pane editable Copy/Paste also creates a new ID. Detailed 07–12 and 19–20 evidence below | **Pass** for same-document cases |
| Resize and move inline drawing, edit with same extent and changed extent; compare dimensions and paragraph placement | Doubled inline height, right-aligned its paragraph while edit was open, then updated. Dimensions differ by less than 0.05 pt; paragraph identity and right alignment retained (13–14). Changed intrinsic extent retains scale (23); whole-control cut/paste retains identity and updates the new paragraph (24–25). Floating placement is unsupported | **Pass** for tested inline cases |
| Undo and redo a native update in Word; reopen editor for each restored state | One Undo restores the complete pre-save object snapshot, Redo restores the post-save snapshot; each state opens the matching native bytes (09–12) | **Pass** for one update cycle |
| Enable AutoSave on an explicitly approved test location, update and reopen after save completes | Not completed; current test DOCX is local with AutoSave off. No test document uploaded | **Pending** |
| Protect/delete captured target or force Word update failure; Save in native app; verify old image/data and retained new draft, no false success | Actual Word 90° picture rotation triggers preflight rejection: saved document XML and object data unchanged, latest draft retained, no matching ACK. Undo rotation then Save succeeds (15–17). Protected/deleted target and partial-mutation rollback remain untested in host | **Pass** for preflight failure/retry; other cases **Pending** |
| Close pane during edit; change and Save; verify no matching ACK, dirty editor and recoverable draft; explicitly reimport draft | Closed the actual pane, rotated again and saved. Native title stayed dirty and displayed “Draft saved at …” in error status. Request revision `97cd837d39a3c02e9ca8462f89b9cf8f3b332765e1aed34f43128d942ddb6347` matches the retained draft; only the previous request’s ACK exists. Saving Word and comparing 05/06 snapshots confirms all three objects unchanged. Native preview worker successfully parsed/rendered the retained draft. `06-failure-result.json`, `06-after-pane-closed-save.docx/.json` and private `06-recoverable-draft.rsk` retain evidence. Native Save As plus actual Word file-picker import and subsequent Edit restored identical bytes (18) | **Pass** including UI recovery |
| Exit companion/native/Word abnormally; restart; inspect drafts and document before explicit recovery | Not completed | **Pending** |
| Simulate delayed document synchronization and verify no ACK is mistaken for durable disk/cloud save | Not completed | **Pending** |
| Insert complete synthesis route and complex stereochemical drawing; repeat update/reopen checks | Not completed; single small drawing does not validate route or macrocycle layout fidelity | **Pending** |

Temporary Word runtime logging was enabled using Microsoft’s documented
`CEFRuntimeLoggingFile` preference while diagnosing startup, then the preference
was removed. No matching runtime log file was produced.

The first round trip included user interaction between automated observations.
A third drawing appeared during that interaction; its creation is not used as
evidence of automated route insertion. The later controlled save retained all
three objects. The first picture width changed from 877203 to 876935 EMU
(about 0.021 pt) after rotation; its height stayed 514985 EMU. This is evidence of
near size preservation in that case, not exact layout invariance.
The report distinguishes observed state, persisted evidence and missing steps;
it does not claim unattended end-to-end automation or complete acceptance.

## Continuation evidence: duplicate, undo, layout and recovery

Snapshots listed here include both `.docx` and `.json` unless otherwise noted.
The JSON associates content-control ID, logical object ID, native and preview
hashes, displayed-media hash and inline dimensions. Each snapshot follows an
actual Word Save. Original Word controls and embedded native data were inspected
from saved packages; the document was not modified by an XML editing script.

| Steps in actual desktop apps | Result and local evidence | Status |
| --- | --- | --- |
| Select only the second picture → Cmd+C → move outside → Cmd+V → Edit pasted picture | `07-picture-only-copy`: one ordinary picture added, no editable control. Pane explicitly refused to edit a plain image. No chemistry was guessed | Pass for refusal; picture-only paste does not preserve editability |
| Select second control's blue title → copy/paste outside → Edit duplicate → rotate in ReShiki → Save | `08-full-control-copy-before-edit`, `09-copy-forked-before-save`, `10-copy-after-update`: copy control `1850909119` gets new object ID `d1ff044d-f0c1-43ac-906d-cc2849e8b62b`; its native revision changes from `27d7e060…` to `f5c0d8fa…`. First three objects unchanged; picture count stays 5 | Pass |
| Undo that update once in Word → Save → Edit; quit without changes → Redo → Save → Edit | `11-undo-once` matches the complete 09 active-object snapshot; `12-redo-once` matches 10. Reopened native drafts match the corresponding revisions. `12-copy-undo-redo-result.json` records comparisons | Pass |
| While copy edit is open, set Word picture height to 3.1 cm with aspect lock, then right-align its paragraph; native rotate/save | `13-resized-right-before-update` → `14-resized-right-after-update`: extent 1499863 × 1116000 → 1499235 × 1115695 EMU (about 0.0495 × 0.0240 pt difference). `14-layout-paragraph-result.json`: same paragraph ID `3E81A817`, same right alignment. Other objects unchanged | Pass for this inline placement; extension below |
| Rotate Word copy 90° → native change/save → wait for failure → save Word; undo Word rotation → native Save again | `15-word-rotation-before-rejection`, `16-word-rotation-rejected`, `16-rejection-result.json`: all objects and `word/document.xml` unchanged by rejected update; draft matches failed request, only an older ACK exists. Pane reports unsupported transformation. `17-retry-after-word-reset` records successful later update. This tests rejection before mutation, not rollback after a partial write | Pass |
| In previously failed pane-close edit, use native Save As; insert that file using Word file picker; select inserted object → Edit | `06-recovered-via-save-as.rsk`, `18-recovered-via-ui`, `18-recovery-result.json`: saved file, embedded native and reopened draft all hash to `97cd837d39a3c02e9ca8462f89b9cf8f3b332765e1aed34f43128d942ddb6347`. Existing objects unchanged. Recovery is an explicit new insertion | Pass |
| Select original second object → pane Copy editable drawing → outside cursor → Paste editable drawing; then edit and save original | `19-editable-clipboard-copy`, `20-original-edited-after-copies`, `20-reverse-isolation-result.json`: new clipboard copy has independent ID `28bcbbde-dfb6-4e1c-b35a-cb2c08222f9d`. Editing original changes neither copy nor other objects; picture count remains 7 | Pass |
| Before fix: save native original → Cmd+W leaves Untitled window → Word Edit same object | Pane rejects with “This drawing already has an open edit session”. Explicit End edit session was needed. Reproduced on initial native source; no data loss | **Failed, fixed below** |
| Rebuild/restart patched companion/app; edit original → rotate/save → Word Save → native Cmd+W, retain Untitled process → Word Edit same object | On `008548d7`, pane automatically reports Edit finished, then opens new session. `21-tab-close-fixed-save`, `21-tab-close-fixed-result.json`: same object/native bytes loaded, exact prior save ACK, closed notice and authorization; old and new native processes coexist. Other objects unchanged, 7 pictures | **Pass after fix** |
| In reopened native session, Save As independent file; rotate/save that file while app stays open | Pane automatically finishes old binding. `22-after-save-as-detachment`, `22-save-as-result.json`: standalone file changes; Word and session draft stay unchanged; no new session request/ACK. `21-22-host-observations.txt` records actual UI steps | **Pass after fix** |
| Change native structure rotation to 90° (not Word picture rotation) → Save → Word Save | On `008548d7`, `23-changed-native-extent`, `23-24-layout-result.json`: intrinsic extent 2082 × 1550 → 1691 × 1686; displayed extent 1499235 × 1115695 → 1217678 × 1213588 EMU. Both dimensions match the existing scale formula within 0.00003 pt. Same paragraph and right alignment, other objects unchanged. Frame dimensions intentionally follow native extent | Pass for scale retention; not fixed-frame layout |
| Select whole control via blue title → Cmd+X → outside cursor/new paragraph at end → Cmd+V → Save → Edit → native rotate 180° and Save → Word Save | `24-full-control-moved`, `24-moved-reopen-result.json`, `25-moved-object-after-update`, `25-move-update-result.json`: control `1850909119` and logical ID retained; paragraph changes from `3E81A817` to `36961CD0`. Reopen loads identical native bytes. Update retains the new paragraph/alignment and matches exact ACK, other objects unchanged and picture count stays 7. Width/height rounding after update stays below 0.031 pt | Pass for this whole-control cut/move/update |

The pane was reloaded through **Add-in Options → Reload** after the companion
restart; it reconnected without closing Word. That is a clean restart with no
active save, not evidence of interrupted-save or crash recovery. Test drawings,
DOCX snapshots and full session data remain local and are not pushed to GitHub.

## Independent ReShiki ribbon: 2026-11-12

At the user's explicit request, added the ribbon before the remaining phase-1
acceptance cases. Code under test: `174665b7`; native app `008548d7`.
Host: the same Apple Silicon macOS 27.2 (26B123), Word 16.100.4 listed above.
Installed Word manifest version `1.1.0.0`, SHA-256
`f1221abbaedbfebf837fa3b0252a21f496385e6bab775a52961986fcb7144296`.
Both the private generated manifest and Word sideload copy match this hash.
The user manually activated the updated Developer Add-ins entry; subsequent
ribbon interactions below were automated through the actual Word UI.

| Check | Result | Local evidence |
| --- | --- | --- |
| Full `cd integrations/office && node --test` | **172 passed, 0 failed** | `ribbon-node.log` |
| `cargo test --locked -p reshiki --lib office_addin` | **12 passed**, 388 filtered | `ribbon-rust.log` |
| `oxlint --deny-warnings website integrations/office packages/editor` | Passed | `ribbon-lint.log` |
| Microsoft `office-addin-manifest validate integrations/office/manifests/word.xml` | Valid; schema and add-in commands accepted | `ribbon-manifest.log` |
| `node scripts/check_docs_tree.mjs`; `git diff --check` | Passed; 902 referenced assets, no whitespace errors | `ribbon-docs.log` |

The eight new Node tests cover command readiness/visibility/completion, failure
handling, shared resources and custom port generation, checked-in manifests,
one client/update loop, same-object duplicate refusal, plain-image refusal and
busy-command non-queuing. HTTP tests also check serving `ribbon.js` and rejecting
Office launch metadata on that resource. Mocks are not the following host tests.
No native source changed in this ribbon patch; this is not a new full-workspace
Rust run or a full Bun build/format pass.

| Real-host steps | Actual observation / evidence | Result |
| --- | --- | --- |
| Activate add-in → click ReShiki tab → Open pane → close pane → Open pane | Three Chinese buttons visible; **打开窗格** restores Connected to Word | Pass |
| Cursor outside drawings → **插入结构** → choose local ethanol `.rsk` → Word Save | `26-ribbon-insert`: 8 pictures, 7 editable controls; all six prior objects unchanged. `26-ribbon-insert-result.json` | Pass |
| Select inserted image → close pane → **编辑选中结构** | Pane opens, native editor launches. Recovery draft hash equals embedded inserted native hash; no guessed structure | Pass |
| Native rotate 180°; Word selects first object; hide pane → Open pane → Insert button | Same original recovery session remains displayed; Insert only reveals controls and does not replace runtime or add a picture | Pass |
| Explicitly click Edit on the other object, then save the original native window → Word Save | Separate second session is allowed; original session updates only its captured target. `27-ribbon-update-result.json`: count remains 8; other six objects unchanged; same object/control ID; native, preview and displayed media changed. ACK accepted with exact native revision | Pass for isolation |
| Compare saved frame before/after | 561240 × 267480 → 560705 × 266615 EMU (about 0.0421 × 0.0681 pt difference). Dimensions are not byte-exact | Pass with recorded sub-0.07 pt difference; no exact-frame promise |
| End sessions → select plain picture → ribbon Edit | Explicit “A plain image does not contain editable drawing data”; no new session shown | Pass |
| Word Save → close DOCX → reopen → dismiss stale pane → select updated object → ribbon Edit | Ribbon persists, native draft hash matches `27-ribbon-update` exactly. `28-ribbon-reopen-result.json` | Pass with upgrade workaround below |
| Click ribbon Edit again on the same object | “This drawing already has an open edit session”; same single session remains | Pass |

Raw packages/results `26-*`–`28-*` and `ribbon-host-observations.txt` are in the
private evidence directory used above. UI screenshots were inspected in the
interactive run; no screenshots or user documents are committed. Host test
sessions were explicitly ended, recovery drafts retained, and the DOCX saved.

**Upgrade issue remains:** the older acceptance DOCX still contains a developer
`we:reference version="1.0.0.0"` after installing manifest `1.1.0.0`. Word reopens
an unavailable-add-in pane for that reference. Closing it and using the new
**ReShiki → 编辑选中结构** successfully launches the new runtime without the
Developer Add-ins menu. Native payloads survive unchanged. Automatic migration
of this host-owned reference and fresh-document activation need separate tests;
this upgrade experience is not considered seamless.

SharedRuntime changes the meaning of pane closure: hiding/reopening the new pane
preserved the live session in this run. Earlier case 06 tested runtime loss with
the old non-shared manifest. It must not be presented as a failure test of the new
pane's close button. Saving while hidden, killing the new runtime, and recovering
an in-flight save after Word/companion crash remain Pending. The tested update
was saved after the pane was visible again. Keep it visible to read confirmation.

## Observed image-quality limitation

A direct package inspection found that this Word save **downsampled** the
displayed PNG media: first object 1153 × 677 → 211 × 124 pixels, second object
984 × 733 → 180 × 134. The full-size original previews remain in the custom XML
records beside the native data. This is not just a compression-byte change;
decoded RGBA dimensions differ (`03-displayed-image-check.json` and
`03-decoded-image-check.json`, inspected with Pillow in an isolated uv runtime).
The observed chemistry remained visually recognizable, but pixel equality and
publication print fidelity are not established. Do not claim that an ACK or
native payload hash verifies the saved Word image quality. Test Word image
compression settings and printed output before phase 3; no global Word quality
settings were changed during this run.

## Remaining priorities

1. Verify partial Word mutation/rollback and protected/deleted targets in the
   real host; retain before/after package evidence and recoverable drafts.
2. Extend layout stress tests to multi-page schemes and table cells. Floating
   pictures, cropping and Word picture rotation are currently rejected; inline
   scale retention and whole-control cut/move do not imply arbitrary positioning.
3. Exercise abnormal native/companion/Word exit and delayed document sync.
   AutoSave/coauthoring requires an explicitly approved test location; no local
   drawing was uploaded merely to enable it. Cold-start offline remains Pending.
4. Test complete synthesis routes and complex stereochemical structures across
   save/reopen and undo histories; investigate PNG downsampling/print quality.
5. Correct the reused LibreOffice worker error prefix and improve ambiguous
   native tab titles (upstream #1119). Migrate stale developer manifest references
   and test fresh-document ribbon activation; SVG remains a later phase;
   no Windows OLE-style double-click capability is assumed.
