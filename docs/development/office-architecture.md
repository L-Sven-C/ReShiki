# Office architecture audit (macOS Word fork)

Audit baseline: `d254c8d8cd686d02893ed1283f528ac5d98b9177`, 2026-11-12.
The fork keeps the existing integration; source implementation and mock coverage
are not evidence of a successful Word desktop round trip. See
[the test report](office-test-report.md) for observed host results.

## Existing data flow

1. `integrations/office/manifests/word.xml` declares a `TaskPaneApp`,
   `ReadWriteDocument`, WordApi 1.4 and `https://localhost:43127/taskpane.html`.
   `setup.js:33` generates the same manifest for a configured port.
2. `taskpane.js:196` waits for Office, checks the runtime requirement set, obtains
   a runtime token and creates a per-pane client capability. `host-adapters/index.js`
   rejects unsupported hosts; the stated Word Mac floor is 16.64, with the runtime
   `isSetSupported` check authoritative.
3. `word.js:155` inserts one PNG inside a rich-text content control. Its tag is
   `reshiki:<recordId>`. The document's custom XML stores an immutable version
   record under `urn:reshiki:office-object:1` (`common.js:137`). Each record has a
   logical `objectId`, version `recordId`, host owner and envelope. The envelope
   contains native bytes, PNG, extent, format version and native SHA-256 revision
   (`protocol.js:40`). It carries the complete native drawing, not inferred atoms.
4. `word.js:134` finds a uniquely selected tagged control or the selection's
   parent control. `readSelected` forks a same-document copied control into an
   independent logical object when its owner differs (`word.js:214`). Ordinary
   images have no matching native record and are explicitly refused by the pane.
5. `companion/sessions.js:119` validates the captured target and native drawing,
   creates a private UUID recovery directory, persists `drawing.rsk` and
   `session.json`, and spawns only the configured executable with fixed arguments
   `--office-addin-edit --open <draft>`. The selection is never consulted again
   for that edit. A second live edit of the same logical object is refused.
6. `src/session.rs:171` recognizes the native edit mode. The app save dispatch in
   `src/app/files.rs:864` calls `src/office_addin.rs`. Save writes a draft and a
   unique request, and waits for a matching session/request/revision acknowledgement.
7. The pane polls; `word.js:250` verifies the captured control ID, revision,
   exclusive container, protection and supported picture transformations. It
   stages the new XML before replacing the PNG in the original control. It reads
   back native data and geometry. `edit-flow.js` performs another read before ACK.
   The companion writes `ack.json` only for the exact request and target.
8. A known failed write is rolled back and verified where possible. An uncertain
   partial write blocks the session. A missing acknowledgement leaves the editor
   dirty. End-session and expiry retain drafts rather than deleting them.
9. Native tab detachment writes a versioned, session-specific `closed.json`
   notice. Companion polling or same-object reopening observes it and retires
   the session only after unread/pending saves have been handled. Process exit
   uses the same barrier; closing a tab does not require quitting the app.

## Status and gaps

| Capability | Source evidence | Audit status |
| --- | --- | --- |
| Native molecule / complete drawing insertion | `taskpane.js:61`, `word.js:155`, native preview worker | Implemented; current UX is native Copy → pane Paste or `.rsk` file selection, not a native one-click Word command |
| Selected object → native editor | `word.js:134`, `sessions.js:119` | Implemented; live-host acceptance recorded separately |
| Selection-independent update | `word.js:242`, `word.js:250`, `common.js:72` | Implemented; controlled Mac Word save with another object selected updated only the captured control (test report, 04/05 evidence) |
| PNG + native version readback and exact ACK | `edit-flow.js:19`, `sessions.js:221`, `src/office_addin.rs` | Implemented; ACK means document model update, not disk save or cloud sync |
| Rollback and recoverable drafts | `word.js:282`, `sessions.js:276`, `src/office_addin.rs` | Implemented; metadata-loss draft retention fixed in `3f6b43d0`, with three new native regressions |
| Same-document duplicate independence | `word.js:214`, immutable records in `common.js` | Full-control copy and pane editable copy verified; editing original and copy independently retained other objects. Picture-only paste loses editability |
| DOCX persistence | document-scoped custom XML + tagged control | Close/reopen/re-edit and one real undo/redo/re-edit cycle verified; AutoSave remains Pending |
| Layout retention | `word.js:122`, `word-ooxml.js:192`, `common.js:354` | Inline resize/right alignment, changed intrinsic extent and whole-control cut/move/update verified. IDs and destination paragraph retained. Current algorithm preserves scale, not fixed frame dimensions |
| Crop/rotation/effects/floating pictures | `word-ooxml.js`, `word.js:56` | Unsupported transformations rejected; no silent flattening |
| Pane closure / companion restart | `sessions.js:97`, README Recovery | Pane-close failure retained a valid draft and unchanged Word objects; native Save As, Word UI reimport and Edit verified. Clean companion restart plus pane Reload verified; interrupted-save restart recovery remains Pending |
| Native tab closure / Save As | `src/session.rs:141`, `src/office_addin.rs:77`, `sessions.js:313` | Fixed in `008548d7`; real host releases the reservation while the old process remains running, and reopening loads the saved native bytes |
| Ribbon insertion/edit/pane entry points | Word manifest, `setup.js`, `ribbon.js`, `taskpane.js` | Word-only CustomTab; shared runtime reuses existing file/clipboard insertion, selected-object edit and polling. See host report for acceptance |
| Double-click and SVG preview | PNG-only protocol | Not implemented |

Records are retained for copied objects and undo; the document limit is 256
records / 64 MiB XML (`common.js:9`). Repeated saves consume versions, so the limit
can be reached even with few visible structures. No automatic history pruning
is safe without proving that undo and duplicate references no longer need it.

## Fork correction

`src/office_addin.rs` (`save_with_timeout`) atomically persists the latest drawing before reading
`session.json`. Missing, corrupt or symlinked metadata still rejects authorization
and never emits a new request, but no longer prevents recovery of the latest
edit. Failed disk writes return errors without claiming recovery.
`src/office_addin/recovery_tests.rs` adds three regressions; the original receipt
and stale-ACK checks remain in place.

The continuation reproduced a reservation leak: native Cmd+W replaces the last
drawing with an Untitled tab while the process remains alive. The existing
`Binding::finish` lifecycle hook now calls `office_addin::finish` for standalone
Microsoft 365 bindings. The native writer validates bounded, non-symlink session
metadata and atomically emits only the close notice; it never changes companion
authorization or a save receipt. The companion validates notice size, version
and session identity, then shares the process-exit save barrier. Same-object
start checks the notice before rejecting a live reservation. An invalid notice
does not block other objects, and late old-process exit cannot retire a new edit.
Drafts are never removed by this path. Notice-write failure logs an explicit
End edit session fallback. Six Node regressions, four native filesystem tests
and actual tab-close/Save As host retests cover the change.

## Word ribbon entry points

Manifest version 1.1.0.0 retains the add-in ID and introduces `CustomTab` under
`PrimaryCommandSurface`. Insert, Edit, Open Pane and Preferences are `ExecuteFunction` commands.
Each first shows the shared pane; Open Pane returns to the drawing view and
Preferences switches views without navigating or unloading the runtime. All reference one task-pane page with a long-lived
SharedRuntime 1.1, and no `TaskpaneId`. Excel and PowerPoint manifests are unchanged.
`ribbon.js` registers the four command handlers before asynchronous initialization
finishes, waits for readiness and pane visibility, and always completes the
Office command event. `taskpane.js` routes Edit through the same operation gate
and exact-target workflow as the pane button. A busy operation is reported and
not queued against a later document selection. Insert exposes existing controls
without mutating the document until the user chooses a file or editable paste.

This uses Microsoft's [shared runtime configuration](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime)
and [pane visibility API](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/show-hide-add-in).
Distinct ShowTaskpane URLs can replace pane contents, so command routing never
navigates to a per-action URL or creates another pane client. The runtime may
remain alive while the pane is hidden; document closure or unloading still ends
it. No auto-start-on-document-open behavior is enabled. Companion authentication,
loopback binding, exact save ACKs and recovery storage remain unchanged; the
static resource allowlist adds `/ribbon.js`, `/preferences.js` and the fixed
16/32/80 PNG command icon names.

`preferences.js` persists only a validated language, compact-spacing boolean and
hint-visibility boolean in local webview storage. User-facing session messages
are re-rendered without altering target IDs, envelopes, receipts or polling.
Untranslated diagnostic details remain intact. Settings view changes never stop
the heartbeat or create a second client.

Office.js manifest labels are installation resources; no unsupported runtime
label API is assumed. `setup.js` can generate English (default) or Simplified
Chinese using `manifest(HOSTS[0], port, language)`. The optional trusted local
`wordManifest` config path enables `/api/ribbon-language`, behind existing
Origin/token/client checks and a Word-host check. Requests contain only a language.
`companion/ribbon-language.js` serializes updates, checks for an exact known
manifest, rejects symlinks/custom content and atomically replaces that file with
a generated manifest of the same ID/version/runtime/permissions. No arbitrary
XML or path comes from the pane or document. Restart/re-activation may be needed;
this service does not close Word or change global Office language settings.

## Security boundary

`companion/server.js:171` binds only `127.0.0.1` with TLS 1.2 or later. The handler
requires an exact localhost Host, Origin, JSON content type, runtime token and
per-pane capability. Routes and static resources are allowlisted; no arbitrary
filesystem or command endpoint exists. Worker arguments are fixed and use
`shell: false`. Native bytes are bounded at 16 MiB, PNG at 8 MiB / 32 megapixels,
worker runtime at 15 seconds, and inactive pane leases at 30 seconds.

Certificates, keys, config, recovery drafts and test Word documents stay outside
the Git repository. Native drawing bytes go to the loopback service and the
current Office file. Office itself may sync a document stored in a cloud folder;
tests must use disposable local documents.

Office.js still comes from Microsoft's CDN. This is a local data workflow, not
a verified cold-start offline product. Microsoft's [Office.js loading guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/referencing-the-javascript-api-for-office-library-from-its-cdn)
permits local references for internal/development use, but a complete local
runtime packaging and compatibility test is separate work.

The real Word save in this run downsampled displayed PNG media while retaining
the original preview in custom XML. See the test report for pixel dimensions;
native editability and publication-quality rendering require separate checks.

## Official API boundaries and upstream work

- [Word event reference](https://learn.microsoft.com/en-us/office/dev/add-ins/word/word-add-ins-events)
  documents content-control and selection events, not a custom picture
  double-click activation contract. Do not claim Windows OLE parity. VBA's
  [WindowBeforeDoubleClick](https://learn.microsoft.com/en-us/office/vba/api/word.application.windowbeforedoubleclick)
  is a different extension surface and would require separate Mac testing,
  deployment and macro-security decisions. It is not used here.
- [Word content controls](https://learn.microsoft.com/en-us/javascript/api/word/word.contentcontrol)
  and [inline pictures](https://learn.microsoft.com/en-us/javascript/api/word/word.inlinepicture)
  define the APIs used for the selected-object workflow.
- [#1118](https://github.com/Ameyanagi/ReShiki/issues/1118) leaves Microsoft 365
  add-in, Word, save failures and macOS Office regression checks open.
- [#1119](https://github.com/Ameyanagi/ReShiki/issues/1119) tracks ambiguous native
  tab titles for multiple Office drawings.
- [#462](https://github.com/Ameyanagi/ReShiki/issues/462) tracks editable picture
  payloads and clipboard round trips; that is distinct from add-in acceptance.
  Its past probes cannot establish this fork's Word behavior.

## Licensing and upstream maintenance

The fork retains all 3,242 baseline commits (a non-shallow, blob-filtered clone),
`origin` points to `L-Sven-C/ReShiki`, and `upstream` points to `Ameyanagi/ReShiki`
with its push URL disabled locally. Work is on `feature/macos-word-integration`.
Fetch upstream and merge/rebase deliberately; do not push to upstream or merge
this work into main without authorization.

Original code uses **MIT OR Apache-2.0**. Preserve `LICENSE`, `LICENSE-MIT`,
`LICENSE-APACHE`, `NOTICE`, file notices, `licenses/` and `vendor/`. `NOTICE` also
identifies separately licensed fonts, chemistry data and adapted algorithms:
RDKit BSD-3-Clause, centres BSD-2-Clause, OPSIN/openclatura MIT, Python-2.0
adaptations, LLVM exceptions, Liberation OFL-1.1, DejaVu/Bitstream Vera, ChEBI
CC BY 4.0 and separately licensed NMR data. Dependency pins retain their own
license ledgers. `scripts/license_notices.py` and the existing app packager
assemble distribution notices; a future Release must use them. This is a source
and attribution audit, not a completed binary-distribution license audit.
