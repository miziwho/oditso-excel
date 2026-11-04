# Excel Work Tools v1.1

*[한국어 README](README.md)*

**39 everyday Excel jobs, one key each.** Unpivot columns, multi-VLOOKUP, compare two
lists — the things Power Query or a long manual routine used to be needed for.
Excel 2016+ on Windows.

## Install (3 steps)

1. Close every Excel window.
2. Open `Install.xlsm`, press **[Enable Content]**, then **[Install]**.
3. Quit Excel completely, start it again, and press `Ctrl+Shift+M` to check.

> Get `Install.xlsm` and `PERSONAL.XLSB` from the [latest release](../../releases/latest).

**If the installer is blocked by a security policy,** put `PERSONAL.XLSB` into
`%APPDATA%\Microsoft\Excel\XLSTART` yourself — it behaves exactly the same.
Step-by-step instructions are in [`src/en/INSTALL.txt`](src/en/INSTALL.txt) and in the guide.

## If you only learn five

| Key | What it does |
|---|---|
| `Ctrl+Shift+M` | **Work Tools menu** — forget the rest, this one is enough |
| `Ctrl+Shift+X` | Comma format |
| `Ctrl+Shift+A` | Filter on / off |
| `Ctrl+Shift+W` | AutoFit rows and columns |
| `Ctrl+Shift+K` | Stick the shortcut list onto the sheet |

There is no need to memorise 39. `Ctrl+Shift+M` opens a menu holding every one of them,
each printed with its shortcut.

## All 39 shortcuts

29 on `Ctrl+Shift`, 10 on `Ctrl+Alt`.

### Edit

| Key | What it does | Undo |
|---|---|---|
| `Ctrl+Shift+O` | Open New Window (same file in two windows) | — |
| `Ctrl+Shift+Q` | Copy to New Workbook (current sheet into a new workbook) | — |
| `Ctrl+Shift+X` | Comma Format (set the negative and zero look in `Ctrl+Shift+M`) | `Ctrl+Z` |
| `Ctrl+Shift+Y` | Text to Columns - as General (text numbers become real numbers) | `Ctrl+Z` |
| `Ctrl+Shift+I` | Text to Columns - as Text | `Ctrl+Z` |
| `Ctrl+Shift+U` | Table Borders (bold header + vertical dividers) | `Ctrl+Z` |
| `Ctrl+Shift+W` | AutoFit rows and columns | `Ctrl+Z` |
| `Ctrl+Shift+S` | Freeze / unfreeze header row (header detected for you) | press again |
| `Ctrl+Shift+H` | Tidy All Sheets (every sheet to A1, zoom 100%) | — |
| `Ctrl+Alt+Z` | Center Align / Default (toggle) | `Ctrl+Z` |
| `Ctrl+Alt+S` | Center Across Selection (centered without merging) | `Ctrl+Z` |
| `Ctrl+Alt+X` | Format Brush (copy and paste with one key) | `Ctrl+Z` |
| `Ctrl+Alt+F` | Find by Value (matches the result, selects every hit) | — |

### Fill & Select

| Key | What it does | Undo |
|---|---|---|
| `Ctrl+Shift+1` | Fill Formula Down (to the end of the table) | `Ctrl+Z` |
| `Ctrl+Shift+2` | Insert Index Column (No, `=ROW()-header row`) | `Ctrl+Z` |
| `Ctrl+Shift+3` | Select to End of Table (ignores blanks) | — |
| `Ctrl+Shift+4` | Fill Blanks From Above (tidy up unmerged tables) | `Ctrl+Z` |
| `Ctrl+Alt+A` | Flash Fill (like `Ctrl+E`, but to the end of the table with no range) | `Ctrl+Z` |
| `Ctrl+Alt+Q` | Auto Fill Series (`1,2,3` / `Mon,Tue,Wed` to the end of the table) | `Ctrl+Z` |
| `Ctrl+Alt+E` | Remove Rows Above Header (tidy up before exporting) | `Ctrl+Z` |

### Data

| Key | What it does | Undo |
|---|---|---|
| `Ctrl+Alt+B` | Column Headers to Row Data (`Jan Feb Mar` columns into one) | — |
| `Ctrl+Shift+L` | Multi VLOOKUP (group matching values and spread them out) | `Ctrl+Z` |
| `Ctrl+Shift+V` | Duplicates by Key Column (highlight · number · extract) | `Ctrl+Z` |
| `Ctrl+Alt+D` | Compare Two Columns (between lists · within each row) | `Ctrl+Z` |
| `Ctrl+Shift+D` | Highlight Duplicate Rows (Ctrl-click cells to compare values instead) | `Ctrl+Z` |
| `Ctrl+Shift+C` | Clear Cell Fill (removes fill color in the selection) | `Ctrl+Z` |
| `Ctrl+Shift+J` | Create PivotTable (new sheet, range detected) | — |
| `Ctrl+Shift+T` | Classic PivotTable Layout (removes subtotals) | **cannot be undone** |
| `Ctrl+Shift+G` | SUBTOTAL Sum (in the cell above the header, press again to undo) | `Ctrl+Z` |
| `Ctrl+Alt+G` | SUBTOTAL Count (filled cells, follows the filter) | `Ctrl+Z` |

### Filter

| Key | What it does | Undo |
|---|---|---|
| `Ctrl+Shift+A` | Filter on / off | press again |
| `Ctrl+Shift+Z` | Filter by selected cells (press again to clear those columns) | press again |
| `Ctrl+Shift+F` | Run Advanced Filter (press again to clear) | press again |
| `Ctrl+Shift+R` | Set Up Criteria Columns (headers copied for you) | `Ctrl+Z` |

### Export & other

| Key | What it does | Undo |
|---|---|---|
| `Ctrl+Shift+E` | Export Sheet as Text (for txt to dat conversion) | — |
| `Ctrl+Shift+N` | Copy File Path (full path to the clipboard) | — |
| `Ctrl+Shift+P` | Set Password (open password, default `1234`) | — |
| `Ctrl+Shift+K` | Paste / remove this list | `Ctrl+Z` |
| `Ctrl+Shift+M` | Work Tools menu — when in doubt, press this one | — |

On top of these, **closing a saved file with no open password asks whether to set one.**

## Undo

| | Count | |
|---|---|---|
| `Ctrl+Z` takes it back | 23 | |
| press the same key again | 4 | filter toggles |
| nothing to undo | 11 | selecting cells · opening a window · copying a path |
| **cannot be undone** | **1** | **`Ctrl+Shift+T` Classic PivotTable Layout** |

- Only **one** thing changes data without a way back: `Ctrl+Shift+T`. Save before running it.
- `Ctrl+Z` has to come **straight afterwards**. Touch another cell in between and Excel
  throws the registration away.
- Anything that creates a sheet (`Ctrl+Alt+B`, `Ctrl+Shift+J`, and the new-sheet extracts of
  `Ctrl+Shift+L` / `Ctrl+Shift+V`) is undone by right-clicking the tab and deleting it.

## Things to know

- If you already use a `PERSONAL.XLSB`, installing replaces it.
  **The v1.1 installer backs the old one up first** —
  `%APPDATA%\Microsoft\Excel\PERSONAL_backup_<date>.XLSB`
- It takes over these built-in Excel shortcuts: `Ctrl+Shift+1` `3` `4` (number formats),
  `Ctrl+Shift+A` `F` `P`.
- `Ctrl+Shift+B` is deliberately left empty — on some machines another background program
  takes that key first.
- On many keyboard layouts the **right Alt is AltGr**, which the layout may use for accented
  characters. Use the **left Alt** for the `Ctrl+Alt` shortcuts.
- Settings (comma format look · ask on close · default password) live inside `PERSONAL.XLSB`.
  Nothing is written to the registry, so moving that one file takes the settings with it.

## Uninstall

In `Install.xlsm`, press `Alt+F8` and run `PERSONAL_Remove`.
If you installed it by hand, delete `PERSONAL.XLSB` from the `XLSTART` folder.

## Requirements

- Excel 2016 or newer, Windows, 32/64-bit
- No external references, no DLLs

## User guide

- [`docs/Excel_Work_Tools_v1.1_Guide.html`](docs/Excel_Work_Tools_v1.1_Guide.html)

Install, all 39 features, screenshots and a FAQ on one page. It opens in English and the
**한국어** button in the sidebar switches the whole page to Korean.

Ten of its screenshots still show Korean Windows and Korean Excel dialogs — those are
Windows' and Excel's own screens, which appear in your language on your machine, so the
steps still match.

## Repository layout

```
oditso-excel/
├─ src/        VBA source (CP949 encoded — do not convert to UTF-8)
│   ├─ mod업무도구.bas           Korean module
│   ├─ ThisWorkbook_코드.txt     Korean ThisWorkbook code
│   ├─ modInstall.bas            Korean installer logic
│   └─ en/                       English edition
│       ├─ mod_WorkTools.bas        the tools
│       ├─ ThisWorkbook_code.txt    paste into ThisWorkbook
│       ├─ modInstall.bas           installer logic
│       └─ INSTALL.txt              install guide for users
├─ docs/       HTML user guides (Korean, English)
└─ dist/       built files
```

## Licence

[CC BY-NC-ND 4.0](LICENSE) — attribution: **oditso (oditso.com)**

## Links

- https://oditso.com
- If it saved you time, ☕ [buy me a coffee](https://www.oditso.com/donate)
