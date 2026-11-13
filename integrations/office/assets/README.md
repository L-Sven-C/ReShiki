# Office command icons

`insert.svg`, `edit.svg`, `pane.svg` and `preferences.svg` are original source
artwork for this fork, under the repository's MIT OR Apache-2.0 license.
They distinguish file insertion, chemical drawing editing, pane display and
settings. Zotero's command grouping informed the layout; no Zotero artwork or
code is copied. The existing `icon-*` assets remain the add-in identity icons.

Office manifest buttons consume the checked-in 16/32/80 PNG exports. To regenerate
one with librsvg:

```sh
rsvg-convert -w 32 -h 32 preferences.svg -o preferences-32.png
```

Repeat for each source and size. Word chooses the actual ribbon layout and
button sizes; this manifest does not promise RibbonX-style small/large controls.
