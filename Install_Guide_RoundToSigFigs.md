# Packaging & Installing the "Round to 3 Sig Figs" Add-in

This guide turns the `RoundToSigFigs.bas` module into a proper Excel add-in (`.xlam`)
that loads automatically every time Excel starts — no XLSTART folder, no per-file
macro security prompts, and a permanent Ribbon button.

You only need to do **Part 1** once, to build the add-in file. Everyone else only
needs **Part 2**.

---

## Part 1 — Build the .xlam (do this once)

### Step 1: Enable the Developer tab
1. File → Options → Customize Ribbon.
2. On the right-hand list, check **Developer**, then click OK.

### Step 2: Create a new blank workbook
1. Open a new blank Excel workbook.
2. Go to the **Developer** tab → click **Visual Basic** (or press Alt+F11).

### Step 3: Import the macro module
1. In the VBA editor, right-click on the workbook's project in the left pane (e.g. `VBAProject (Book1)`).
2. Select **Import File…**
3. Browse to `RoundToSigFigs.bas` and import it.
4. You should now see a module called `RoundToSigFigsModule` in the project tree.

### Step 4: Test it
1. Close the VBA editor.
2. Type a few unrounded numbers into some cells (e.g. `4582`, `0.004582`, `123.456`).
3. Select those cells.
4. Run the macro: Developer tab → Macros → select `RoundToSigFigs` → Run.
5. Confirm the cells update in place to 3 significant figures.

### Step 5: Save as an Add-in
1. File → Save As.
2. In "Save as type," choose **Excel Add-in (*.xlam)**.
3. Name it something clear, e.g. `RoundToSigFigs.xlam`.
4. Save it somewhere you'll remember (you'll distribute this file to your team).

At this point the macro logic is packaged — but it's still only accessible via
Developer → Macros. Part of your original ask was a **one-click button**, so:

### Step 6: Add a Ribbon button (recommended)
Excel doesn't let you edit Ribbon XML directly inside the file — you need a small
free utility called **Custom UI Editor for Microsoft Office** (a standard, widely
used Microsoft-ecosystem tool, not a third-party macro risk).

1. Download and install "Custom UI Editor for Office" (search for it — it's a free
   tool commonly used for exactly this purpose).
2. Open your `RoundToSigFigs.xlam` file in Custom UI Editor.
3. Go to Insert → **Office 2010+ Custom UI Part**.
4. Paste in the contents of `customUI.xml` (provided alongside this guide).
5. Save and close Custom UI Editor.

This adds a **"Data Tools" tab** to the Ribbon with a **"Round to 3 Sig Figs"**
button, wired to the macro automatically.

> **Simpler alternative (no extra tool):** skip Step 6 and instead have each user
> add the macro to their **Quick Access Toolbar** once the add-in is installed
> (Excel Options → Quick Access Toolbar → choose "Macros" from the dropdown →
> add `RoundToSigFigsModule.RoundToSigFigs`). Slightly less elegant than a Ribbon
> tab, but zero extra software needed.

### Step 7 (optional but recommended): Digitally sign the file
If your organization enforces macro security, get the `.xlam` digitally signed
(Tools → Digital Signature in the VBA editor, using a certificate from your IT
department) so users don't see security warnings when it loads.

### Step 8: Distribute
Place the finished `.xlam` on a shared drive, SharePoint site, or send it directly
to your team — this is the file each person installs in Part 2.

---

## Part 2 — Installing the add-in (each user does this once)

1. Save the `RoundToSigFigs.xlam` file somewhere permanent on your computer
   (e.g. `Documents\Excel Add-ins\`). Don't leave it in Downloads, since moving it
   later will break the link.
2. In Excel: File → Options → Add-ins.
3. At the bottom, next to "Manage," make sure **Excel Add-ins** is selected, then
   click **Go…**
4. Click **Browse…**, select your saved `RoundToSigFigs.xlam` file, and click OK.
5. Make sure the checkbox next to it is checked, then click OK.
6. Restart Excel.

From now on, every time Excel opens:
- If you used the Ribbon button method: you'll see a **Data Tools** tab with a
  **Round to 3 Sig Figs** button.
- If you used the QAT method: the button will be pinned to your Quick Access
  Toolbar (top-left, always visible).

**To use it:** highlight the range of unrounded numbers → click the button.
The selected cells are rounded to 3 significant figures in place — same
one-click behavior as the original XLSTART macro.

---

## Notes

- Because this loads through Excel's official Add-ins system rather than a raw
  file drop into XLSTART, it survives Excel updates and reinstalls more reliably.
- If you improve the macro later, you only need to update the one shared `.xlam`
  file — users don't need to do anything unless the file path changes.
- If your organization uses centrally managed software deployment (e.g. SCCM,
  Intune, or a Group Policy startup script), IT can push the `.xlam` to each
  user's Add-ins folder automatically, skipping Part 2 entirely for future hires.
