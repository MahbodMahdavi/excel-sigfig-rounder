# Excel Sig-Fig Rounder

A one-click Excel add-in that rounds the numeric cells in your current selection
to **3 significant figures**, overwriting the cells in place — no VBA knowledge,
no XLSTART folder juggling, no re-running a macro by hand.

Replaces an older XLSTART-based macro with a proper, distributable `.xlam`
add-in that loads automatically every time Excel starts.

## What's in this repo

| File | Purpose |
|---|---|
| `RoundToSigFigs.bas` | The VBA module — rounds every numeric cell in the current selection to 3 significant figures. |
| `customUI.xml` | Ribbon XML that adds a **Data Tools** tab with a **Round** button. |
| `round_icon.png` | Custom icon for the Ribbon button (blue circle, embossed "R"). |
| `RoundToSigFigs.xlam` | The finished, ready-to-install Excel add-in (macro + Ribbon button + icon, all packaged together). |
| `Install_Guide_RoundToSigFigs.md` | Full build-from-scratch instructions, for rebuilding or modifying the add-in. |

## Quick install (most people only need this)

1. Download `RoundToSigFigs.xlam` from this repo.
2. In Excel: **File → Options → Add-ins**.
3. Next to "Manage," make sure **Excel Add-ins** is selected → click **Go…**
4. Click **Browse…** → select `RoundToSigFigs.xlam` → OK.
5. Make sure its checkbox is checked → OK.
6. Restart Excel.

You'll now see a **Data Tools** tab with a **Round** button, every time Excel opens.
No admin rights required for this step.

## How to use it

1. Select a range of unrounded numeric cells.
2. Click **Round** on the **Data Tools** tab.
3. Selected cells are rounded to 3 significant figures, in place.

Example:

| Before | After |
|---|---|
| 4582 | 4580 |
| 0.004582 | 0.00458 |
| 123.456 | 123 |

## Rebuilding the add-in from source

See `Install_Guide_RoundToSigFigs.md` for full step-by-step instructions, including:
- Importing `RoundToSigFigs.bas` into a blank workbook
- Saving as `.xlam`
- Adding the Ribbon button with [OfficeRibbonXEditor](https://github.com/fernandreu/office-ribbonx-editor)
- Installing the result

## Notes

- The Ribbon button calls `RoundToSigFigsModule.RoundToSigFigs`, which accepts
  an optional `IRibbonControl` parameter (required by the Ribbon callback
  signature) but can also be run manually from Developer → Macros.
- Rounding logic: for a given number, decimal places are calculated as
  `3 - 1 - floor(log10(abs(value)))`, then applied via `Round()`. This mirrors
  standard "3 significant figures" rounding, including for values below 1 and
  negative numbers.
