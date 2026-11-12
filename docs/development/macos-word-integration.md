# macOS Word integration development

This fork prioritizes repeatable native editing of drawings embedded in Word.
It retains upstream's PNG + native `.rsk` + custom XML design. See
[architecture and limitations](office-architecture.md) and the
[observed test results](office-test-report.md) before using it for research work.

## Checkout and toolchain

```sh
git clone --filter=blob:none https://github.com/L-Sven-C/ReShiki.git
cd ReShiki
git remote add upstream https://github.com/Ameyanagi/ReShiki.git
git remote set-url --push upstream DISABLED
git switch feature/macos-word-integration
```

This is not a shallow clone: commit history is retained, and old file blobs can
be fetched on demand. Do not rename the project, overwrite upstream history,
merge into main or publish a Release as part of this phase.

- Apple Silicon is the primary test architecture. The app bundle currently
  declares macOS 14.0 as its minimum (`scripts/build_release.py`). Only the OS
  and Word build actually recorded in the test report have been inspected here.
- Microsoft 365 Word desktop, WordApi 1.4. The upstream version floor is Word
  Mac 16.64; the runtime API check, activation and real behavior must be verified.
- Rust **1.99.1**, with rustfmt and Clippy, as pinned in `rust-toolchain.toml`.
  The older prose in `docs/development.md` is not the authoritative pin.
- Node.js 22+, Apple Clang/Xcode command line tools, Python 3.12+ for packaging.
- Bun 1.4.4 for the repository's complete JavaScript development workflow.
  The companion has no npm runtime dependencies; tests require the pinned
  development dependency `@xmldom/xmldom`.

```sh
rustup toolchain install 1.99.1 --profile minimal --component rustfmt --component clippy
bun install --frozen-lockfile
(cd integrations/office && node --test)
cargo fmt --all --check
cargo test --locked -p reshiki --lib office_addin
cargo check --workspace --all-targets --locked
cargo clippy --workspace --all-targets --locked -- -D warnings
python3 scripts/build_macos_app.py
```

If Homebrew's Rust precedes rustup, use `PATH="$HOME/.cargo/bin:$PATH"` for these
commands. For Office-only testing without Bun, `npm install --ignore-scripts
--no-save --package-lock=false --workspaces=false` in the repository root installs
its pinned development tools without changing the lockfile; it does not establish
that the website build or Bun formatting checks pass.

The development app is `target/debug/ReShiki.app`. Packaging preserves upstream
license notices. The full Rust workspace tests additionally need the native
InChI and naming helpers described in `docs/development.md`; a filtered Office
test pass must not be reported as a full workspace test pass.

## Local HTTPS and Word installation

Use Microsoft's development-certificate tooling or an organization-approved
localhost certificate. Trust installation is a machine change: the user must
complete any macOS authorization prompt themselves. Never put certificate keys
or recovery files in this repository.

```sh
npx --yes office-addin-dev-certs@3.0.1 install --days 30
node integrations/office/setup.js \
  --executable "$PWD/target/debug/ReShiki.app/Contents/MacOS/reshiki" \
  --cert "$HOME/.office-addin-dev-certs/localhost.crt" \
  --key "$HOME/.office-addin-dev-certs/localhost.key" \
  --directory "$HOME/Library/Application Support/ReShiki-Office-Dev"
node integrations/office/companion/server.js \
  --config "$HOME/Library/Application Support/ReShiki-Office-Dev/config.json"
```

Keep that terminal running. The server must listen only on `127.0.0.1:43127`.
Check `https://localhost:43127/taskpane.html` without bypassing any TLS warning.
`setup.js` validates certificate/key matching, validity, localhost SAN and the
executable, and never overwrites conflicting configuration. Renew expired
development certificates using the same reviewed tooling; do not disable TLS.

In another terminal:

```sh
mkdir -p "$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
cp -n "$HOME/Library/Application Support/ReShiki-Office-Dev/manifests/word.xml" \
  "$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef/reshiki-word-dev.xml"
```

Close/restart Word after saving any open documents. Open a disposable **local**
DOCX, then choose **Home → Add-ins → Developer Add-ins → ReShiki**. This was
the working entry on Word 16.100.4; the older Insert → My Add-ins dialog did not
list this sideload. Automated accessibility clicks may not activate that entry;
a manual click opened the pane during testing. Follow Microsoft's
[Mac sideloading instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac).
Only the Word manifest is needed. A generated manifest does not prove that Word
has loaded it successfully.

## Editing workflow

1. Draw in the native app, Copy, then choose **Paste editable drawing** in the
   Word pane. Alternatively use **Insert a drawing file** with a native `.rsk`.
   A complete route remains one editable drawing, including its native objects.
2. Select the inserted drawing and click **Edit selected drawing**. Keep the
   pane and companion open. The target is captured before the native editor opens.
3. Modify and Save in ReShiki. Wait for confirmed update in the pane and native
   editor. Save the Word document separately; a native ACK is not a DOCX disk-save
   or cloud-sync receipt.
4. Close the native editor after the confirmed save, or explicitly end its
   session in the pane before opening the same object again.
5. Save, close and reopen the DOCX, reopen the pane, select the object and edit
   again. Check the test report for actual verification of this sequence.

Plain PNG/SVG images without native records are not editable through this route.
The software does not reconstruct chemistry from pixels. Keep pictures inline
and free of crop, rotation or picture effects for this phase. Current sizing
preserves the user's scale relative to the drawing's extent, so adding content
can enlarge the frame; fixed-frame layout is not yet promised.

## Recovery

Each edit displays its recovery path. `drawing.rsk` is the native draft;
`session.json`, `request.json` and `ack.json` identify authorization and saves.
After a failure, do not discard the native window or delete recovery files.
Use Save As to keep an independent `.rsk`. If the pane closes or the companion
restarts, open the retained draft explicitly, then import it as a new object in
Word; no automatic reconnection to a possibly different document occurs.

If Word is locked or an update is safely rolled back, resolve the error and
Save again in ReShiki. An uncertain partial update requires explicit recovery.
Compare the original and recovered object before replacing anything manually.
Do not infer success merely because a PNG or XML part was written.

## Next phases

The first gate is the real Word test matrix. Only after that gate passes should
work proceed to ribbon commands, fewer clicks, window activation, or SVG.
No custom double-click/OLE behavior is assumed. The documented Office.js event
surface does not provide that activation contract; VBA/macros are a separate
potential investigation, not part of the current implementation.

SVG research must test printed bond widths, text, stereochemistry and Word
save/reopen behavior while retaining independent native data. It must not remove
the working PNG path. Offline Office.js packaging and repeated-edit history
capacity are also separate, explicit follow-up items.
