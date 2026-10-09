# WinVectorPptAddon — drop PDFs into PowerPoint as vector graphics

On a Mac, dragging a PDF onto a PowerPoint slide inserts it as a crisp vector
image. On Windows you get an embedded-object icon instead. This add-in fixes
that: drag a PDF from Explorer **straight onto a slide** and it is inserted as
a **vector graphic**, centered where you dropped it.

Nothing to click, no extra window: install once, then just drag and drop.

## Install

Pick one. All three are per-user (registry under `HKCU` only), need **no
admin rights**, and install nothing but this folder. Restart PowerPoint
afterwards.

**One line** — open PowerShell and run:

```powershell
irm https://raw.githubusercontent.com/KLG/WinVectorPptAddon/main/get.ps1 | iex
```

This downloads the latest release (add-in + bundled PDF converter) into
`%LOCALAPPDATA%\WinVectorPptAddon` and registers it. Run it again to update.

**Zip** — download `WinVectorPptAddon-<version>.zip` from the
[latest release](https://github.com/KLG/WinVectorPptAddon/releases/latest),
extract it anywhere you want to keep it, and double-click **`Install.bat`**.

**Git** — `git clone https://github.com/KLG/WinVectorPptAddon` and
double-click `Install.bat`. The repository does not contain the 70 MB Poppler
binaries; the installer downloads them on first run.

To remove: double-click `Uninstall.bat` (or run `uninstall.ps1`) and restart
PowerPoint. Deleting the folder afterwards leaves nothing behind.

Requirements: Windows 10/11 and PowerPoint from Microsoft 365 or Office
2016+ (32- or 64-bit). Tested with Microsoft 365 PowerPoint 16.0 (32-bit) on
Windows 11.

## What happens on a drop

- The graphic is placed **centered on the drop point**, at the PDF page's
  natural size (capped at 95 % of the slide so it always fits), with the aspect
  ratio locked.
- Multi-page PDFs: page 1 goes where you dropped it; each further page gets its
  own new blank slide right after.
- Several PDFs dropped together are handled the same way, in order.
- Conversion runs in the background, so PowerPoint stays responsive; the
  graphic appears about a second after the drop (dense plots take longer, see below).
- Anything that is not a PDF-only drop (images, text, Office files, …) is passed
  to PowerPoint untouched, so nothing else changes.

Text is converted to outlines, so the result looks exactly like the PDF
regardless of installed fonts (like the Mac's PDF import, it is not editable
text). Right-click the graphic → *Convert to Shape* to edit the paths in
PowerPoint.

## How it works

PowerPoint registers an OLE drop target on each document window. The add-in
(`src\PdfDropAddin.cs`, a plain COM add-in in C#, no dependencies) runs inside
PowerPoint and wraps that drop target: PDF drops are handled by the add-in,
everything else is forwarded to PowerPoint's original handler.

PowerPoint cannot read PDF, but it imports SVG as a real vector graphic. Each
page is converted to SVG with **pdftocairo** from
[Poppler](https://poppler.freedesktop.org/) (bundled in `tools\poppler\`:
only the files pdftocairo needs, about 70 MB, see `NOTICE.txt` there), then
inserted with PowerPoint's own `Shapes.AddPicture`;
`DocumentWindow.PointsToScreenPixels` maps the drop position onto the slide.

If `tools\poppler\` is missing, `install.ps1` downloads the latest
[poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)
release and recreates it. Failing that, the add-in also looks for `pdftocairo`
on `PATH` or in a MiKTeX install, then for Inkscape, and as a last resort
inserts pages as 300 DPI bitmaps so a drop always does something useful.

### Dense plots (flow cytometry, large scatter plots)

PowerPoint imports SVG on its user-interface thread and needs roughly 0.25 ms
per drawn element, so a plot with 50 000 dots would freeze it for 10 to 15
seconds. Before inserting, the add-in therefore rewrites heavy pages in the
background (`src\SvgOptimizer.cs`); pages PowerPoint can import in under half
a second are left exactly as converted.

1. **Merge.** Neighbouring shapes with identical opaque paint are joined into
   one path (up to 500 per path). The result is pixel-identical and still
   100 % vector.
2. **Dense layers.** If the page would still take more than about 3 seconds
   (typically semi-transparent dots, which cannot be merged without changing
   how overlaps darken), the big runs of dots are drawn into a transparent
   600 DPI image that sits at the same place in the stacking order. Axes, text,
   gates, legends and everything else stay vector.

Time from drop to graphic on an 8-core laptop, 500 × 400 pt test plots:

| Plot | Without | With | Result |
| --- | --- | --- | --- |
| 50 000 opaque dots, 4 colours in sequence | 15.7 s | 5.4 s | all vector (merged) |
| 50 000 dots at 30 % opacity | 15.2 s | 4.5 s | vector axes + dot layer |
| 20 000 outlined dots | 10.6 s | 3.3 s | vector axes + dot layer |
| 200 000 dots at 30 % opacity | 30.3 s | 12.1 s | vector axes + dot layer |

Without the rewrite PowerPoint is frozen for nearly all of that time; with it,
for the last 1 to 4 seconds only.

### Settings

Optional values under `HKCU\Software\PdfDrop`:

| Value | Meaning |
| --- | --- |
| `Mode` | `vector` (default) or `raster` to force bitmaps |
| `PdfToCairo` | full path to `pdftocairo.exe` |
| `Inkscape` | full path to `inkscape.com` |
| `DenseLayers` | `auto` (default), `vector` (merge only, never draw a layer as image; dense plots import slowly) or `off` (insert pdftocairo's SVG unchanged) |
| `DenseLayerDpi` | resolution of dense layers, default `600` |

A short log is kept at `%LOCALAPPDATA%\PdfDrop\log.txt` (it also records which
converter was used, how long each conversion took and what was optimized).

### Files

| File | Purpose |
| --- | --- |
| `Install.bat` / `Uninstall.bat` | Double-click installers (run the `.ps1` scripts below) |
| `install.ps1` / `uninstall.ps1` | Register / unregister the add-in for the current user; fetch Poppler if missing |
| `get.ps1` | The one-line web installer |
| `make-release.ps1` | Builds `dist\WinVectorPptAddon-<version>.zip` (folder + Poppler) to attach to a GitHub release |
| `bin\PdfDropAddin.dll` | Prebuilt add-in (what gets registered) |
| `src\*.cs`, `build.ps1` | Source (`PdfDropAddin.cs`: the add-in; `SvgOptimizer.cs`, `SvgRaster.cs`: dense-plot handling) and build script (building needs the Windows SDK; not required to install) |
| `tools\poppler\` | Bundled Poppler `pdftocairo` (PDF → SVG) |
| `DropPdfToPowerPoint.ps1`, `Drop PDF to PowerPoint.bat` | Older stand-alone fallback (bitmap only) that needs no installation |

## Troubleshooting

- **Nothing happens on drop** → In PowerPoint, *File → Options → Add-ins →
  Manage: COM Add-ins → Go…* and check that "PDF drop" is listed and ticked.
  If PowerPoint disabled it after a crash, re-run `Install.bat` (it clears the
  disabled-items list) and restart PowerPoint.
- **Graphic is a bitmap, not vector** → `tools\poppler\bin\pdftocairo.exe` is
  missing or could not run; the log says which converter was detected. Re-run
  `Install.bat` to download Poppler again, or set `PdfToCairo` in the registry
  as above.
- **Still inserts an icon** → the drop also contained non-PDF files; the add-in
  only takes over drops that consist solely of `.pdf` files.
- **Password-protected PDFs** are not supported (an error message is shown).
- **"Running scripts is disabled"** → use `Install.bat`, or run
  `powershell -ExecutionPolicy Bypass -File install.ps1`.
