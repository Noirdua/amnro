# amnro

A digital-root arithmetic board for Excel. Each digit lives in its own square, and
anything you select is broken down into its number, sum, digital root, opposite
(nines-complement) pairs, Gauss pairings and letter values.

- **Workbook:** `amnro.xlsm`
- **Sheets:** `Menu`, `Add`, `Subtract`, `Multiply`, `Divide` (plus a hidden `_IntelliSense_` sheet)
- **Platform:** Windows desktop Excel (64-bit or 32-bit)

---

## Setup

1. **Open `amnro.xlsm` in desktop Excel.** It is macro-enabled (`.xlsm`).
2. **Enable macros.** On the yellow bar click **Enable Content**, or set
   *File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings ->
   Enable VBA macros*.
3. **Trust the folder (recommended).** *Trust Center -> Trusted Locations -> Add new
   location* and point it at the folder holding `amnro.xlsm`. Tick *Subfolders* if you
   keep add-ins below it. A trusted location stops the macro and add-in prompts each time.
4. **Formula autocomplete (optional, recommended).** The workbook fills in argument help
   for its own formulas (`AddSteps`, `TimesSteps`). That help is drawn by the free
   Excel-DNA IntelliSense add-in:
   - The 64-bit file `ExcelDna.IntelliSense64.xll` is included. Use the 32-bit build
     instead (from https://github.com/Excel-DNA/IntelliSense/releases) if your Excel is 32-bit.
   - Put the matching `.xll` in this folder. On open the workbook registers it, so it
     loads automatically; you do not need to add it by hand each session.
   - The hints themselves live in the hidden `_IntelliSense_` sheet and in
     `amnro.intellisense.xml`, which sits next to the workbook.
5. **Reopen the workbook once** after adding the add-in so the hints are registered.

If your Excel bitness and the `.xll` bitness do not match, the add-in can appear ticked
in the Add-ins dialog yet never actually load. Match the file to the Excel build.

---

## Getting started

The four operation sheets (`Add`, `Subtract`, `Multiply`, `Divide`) each hold a 30 x 30
grid of digit squares.

- **Type a digit** into any square to fill it.
- **`###`** fills the row to the right with digit squares.
- A single digit stays in its cell. A multi-digit number is drawn as one outlined block.
- A decimal point or a leading `+`/`-` takes its own cell.
- **Delete** or **Backspace** clears the selected squares.

### Facts

Select any set of digit squares and the facts are computed for the selection:

- **number** - the selected digits read by row, then column
- **sum** - the running total, with reduction steps when the sum exceeds 9
- **root** - the digital root
- **oppos** - the nines-complement pair of each digit (`0=9`, `1=8`, `2=7`, ...)
- **gauss** - outside-in pairing (`1234` -> `1+4=5, 2+3=5`)
- **letters** - the matching English / Greek / Hebrew letters

The facts show on the **status bar** by default. Turn the floating **Facts overlay** on
or off with **Ctrl+Shift+F** or the **Facts** item on the right-click menu.

---

## Actions

Everything is on the right-click **Arithmetic** menu (right-click any cell). The same
commands appear as buttons on the **Actions** window, opened from the popup's *Actions*
button when the Facts overlay is visible.

| Action | What it does |
| --- | --- |
| **Facts  Ctrl+Shift+F** | Show or hide the floating facts overlay |
| **Invert +/-** | Flip the sign of the selected numbers |
| **Opposite** | Apply the nines-complement |
| **Reverse** | Reverse the digits of the selected signed number(s) in place |
| **Rotate** | Transpose the run around the selected cell (horizontal <-> vertical) |
| **Add x** | Repeated addition (opens the Steps window) |
| **Times** | Step-by-step multiplication (opens the Steps window) |
| **Clear** | Empty the selected squares |
| **Drop digit** | Remove every selected cell equal to a chosen digit, keeping the rest |
| **Copy > Number / Sum / Gauss / Letters / All** | Copy the matching fact |
| **Colors** | Toggle the per-digit colour scheme on the whole workbook |
| **Root tint** | Colour the selection by its digital root |
| **Letter tint** | Show each cell's own letter and colour it by that digit |
| **Letter level** | Cycle the letter tint level (1, 2, 3) |

### Steps window (Add x / Times)

- **Add x** asks for **Number**, **Add** and **Times**.
- **Times** asks for **Number** and **Times**.
- **Times** runs from 1 to 99.
- The results are written starting in the cell you selected, each step as its own digits.
- **Pad** pads single digits (`2` -> `02`). **Gap** leaves one blank cell between steps.
- **Across / Down** chooses the direction. **Across** is the default; **Down** writes one number
  per row.
- **Rows / Digits** applies to Down: **Rows** keeps a number's digits side by side, one number
  per row; **Digits** puts each digit on its own line.
- A live preview and the equivalent formula are shown as you type.

### Drop digit

Select many cells, choose **Drop digit**, then click a digit `0`-`9`. Every selected cell
equal to that digit is removed from the selection; the rest stays selected so the sum and
root update to show the difference. Reading order stays row-by-row.

### Root tint and Letter tint

- **Root tint** colours the whole selection with the colour of the selection's digital
  root. Moving the selection restores the previous cells.
- **Letter tint** shows **each cell's own letter**, chosen by its value and the current
  level, and colours each cell by its own digit. Turning it off puts the digits back.

**Letter level** picks how the letter is derived. The alphabet is numbered `A=1 .. Z=26`,
and each `9` behaves like a `0`, so the level selects which letter of the root's group is
used. For a cell value `v`, the letter index is `v + 9 x (level - 1)`:

| Value | Level 1 | Level 2 | Level 3 |
| --- | --- | --- | --- |
| 1 | A | J | S |
| 2 | B | K | T |
| 8 | H | Q | Z |
| 9 | I | R | (past Z, none) |

So two cells `1` and `2` read `AB` at level 1, `JK` at level 2 and `ST` at level 3.
Letter index `19` is `S` at level 3, which is the "`100`" case: `9 + 9 + 1`.

---

## Formulas

Two worksheet formulas spill the same sequences the Steps window writes:

```
=AddSteps(Start, Add, Times, [Gap], [Pad], [Down], [Split])
=TimesSteps(Number, Times, [Gap], [Pad], [Down], [Split])
```

| Argument | Meaning |
| --- | --- |
| `Start` / `Number` | the first value |
| `Add` | the amount added each step (Add x only) |
| `Times` | how many totals. Counts the first number as one of them |
| `Gap` | `TRUE` leaves a blank cell between values |
| `Pad` | `TRUE` returns padded text such as `01`. Omit for plain numbers |
| `Down` | `TRUE` spills down a column instead of across a row |
| `Split` | `TRUE` puts each digit, and a decimal point, in its own cell |

`Gap`, `Pad`, `Down` and `Split` all default to `FALSE`, so `=AddSteps(1,9,10)` spills
plain numbers.

```
=AddSteps(1,9,10)                         ->  1  10  19  28  37  46  55  64  73  82
=AddSteps(3,9,11)                         ->  3  12  21  ...  93
=TimesSteps(8,9)                          ->  8  16  24  ...  72
=AddSteps(2,9,10,FALSE,TRUE)              ->  02 11 20 29 38 47 56 65 74 83
=TimesSteps(8,9,FALSE,FALSE,FALSE,TRUE)   ->  8 1 6 2 4 3 2 4 0 ... (each digit)
```

A formula is a live formula, so a single-digit result can be padded with a cell number
format; type the formula and press **Enter**.

---

## Digit colours

`Colors` toggles a colour scheme where each digit gets a background and text colour:

| Digit | Background | Text |
| --- | --- | --- |
| 0 | white | black |
| 9 | black | white |
| 1 | red | violet |
| 8 | violet | red |
| 2 | purple | orange |
| 7 | orange | purple |
| 3 | blue | yellow |
| 6 | yellow | blue |
| 4 | green | grey |
| 5 | grey | green |

Opposite digits share a pair: 0-9, 1-8, 2-7, 3-6, 4-5.

---

## Notes

- Formula results are cached by Excel. Values written by a formula recalculate the next
  time the workbook is opened in Excel.
- The workbook uses the right-click menu and the **Ctrl+Shift+F** hotkey rather than a
  ribbon, because a ribbon needs the workbook's custom UI which this tooling cannot write.

## Files

| File | Purpose |
| --- | --- |
| `amnro.xlsm` | the workbook, with all the VBA inside it |
| `ExcelDna.IntelliSense64.xll` | 64-bit formula autocomplete add-in (optional) |
| `amnro.intellisense.xml` | formula argument help read by the add-in |
| `AGENTS.md` | instructions for AI agents editing this workbook |
| `kilo.json` | agent tool configuration (not committed) |
