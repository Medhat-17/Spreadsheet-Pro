# Spreadsheet Pro (Java Swing)

Excel-style spreadsheet in pure Java (JDK 8+, no external libraries).

## Compile and run

Linux / macOS:
    chmod +x build.sh && ./build.sh        # compiles and starts the app

Windows:
    build.bat

Manual:
    mkdir out
    javac -encoding UTF-8 -d out *.java
    java -cp out SpreadsheetApp

## Architecture

| File                        | Layer                                                       |
|-----------------------------|-------------------------------------------------------------|
| SpreadsheetApp.java         | UI: ribbon, formula bar, tabs, status bar, CSV import/export |
| SpreadsheetModel.java       | Data model (AbstractTableModel): raw / formula / style maps  |
| FormulaEngine.java          | Expression-tree formula parser, 64 functions, error codes    |
| Icons.java                  | Vector icons drawn with Java2D (no image files)              |
| DrawShape.java / DrawingLayer.java | Pen, highlighter, shapes and eraser painted over the sheet |
| FunctionCatalog.java        | Names, syntax and examples of every function                 |
| InsertFunctionDialog.java   | fx function picker (search + categories)                     |
| SpreadsheetCellRenderer.java| Cell rendering (bold, italic, alignment, colours, borders)   |
| CellStyle.java              | Per-cell formatting record                                   |
| InsertDialog.java           | Insert date / random data / sales table / chart placeholder  |
| MacroIDEDialog.java         | VBA editor + console                                         |
| MacroEngine.java            | VBA-subset interpreter                                       |
| WorkbookIO.java             | File layer: save / load .spx workbooks, optional encryption  |
| SheetProtection.java        | Sheet protection state + salted PBKDF2 password hash         |
| SheetImage.java             | A picture on a sheet (PNG bytes + position + size)           |
| ImageLayer.java             | Draws pictures over the grid; select / drag / resize         |

## Formulas
Start with `=`. Operators `+ - * / ^ %` and parentheses; references `A1`; ranges `A1:A10`.
64 functions. Click **fx** (or Insert > Function, Shift+F3) to search and insert them; Help > Formula Reference lists all.

- Math & Statistics: SUM AVERAGE MIN MAX COUNT COUNTA COUNTBLANK PRODUCT MEDIAN STDEV VAR LARGE SMALL ABS SQRT ROUND ROUNDUP ROUNDDOWN INT MOD POWER EXP LN LOG LOG10 SIGN PI RAND RANDBETWEEN
- Logical: IF AND OR NOT IFERROR ISNUMBER ISTEXT ISBLANK ISERROR COUNTIF SUMIF AVERAGEIF
- Text: CONCATENATE CONCAT LEN UPPER LOWER PROPER TRIM LEFT RIGHT MID SUBSTITUTE FIND REPT VALUE (and the & operator)
- Lookup: VLOOKUP MATCH INDEX
- Date: TODAY NOW YEAR MONTH DAY DAYS

Also comparison operators (= <> < > <= >=), text in "quotes", and TRUE/FALSE values. IF only evaluates the branch it picks, so `=IF(A1=0,0,10/A1)` is safe. Extra error: `#N/A` (lookup not found).

## Zoom and column width

- Zoom 25%-300%: View menu, status-bar slider / - / + buttons, Ctrl+mouse wheel, Ctrl +/-/0. Each sheet remembers its own zoom.
- Resize a column by dragging the right edge of its header; double-click the edge to fit the content. View > Reset Column Widths restores the default.
- Zoom, column widths and pictures are saved in the .spx file (older files still open).
Errors: #DIV/0!, #SYNTAX!, #NAME?, #ERROR! (bad reference, text used as number, circular reference).

## Macro example
    Sub Demo()
        For i = 1 To 10
            Cells(i, 1).Value = i * i
            If i Mod 2 = 0 Then Cells(i, 1).Interior.Color = RGB(198, 229, 209)
        Next i
        Range("A11").Formula = "=SUM(A1:A10)"
        Range("A11").Font.Bold = True
    End Sub
Supported: Dim, For/Next (Step, Exit For), If/ElseIf/Else/End If (and single-line If),
Cells(r,c) / Range("A1:B3") with .Value .Formula .Font.Bold/Italic/Size/Color .Interior.Color
.HorizontalAlignment, ClearContents, Clear, MsgBox, Debug.Print, and Int Abs Sqr Rnd Round
Len UCase LCase Str Val Mid RGB.

## Shortcuts
Ctrl+B bold, Ctrl+I italic, Ctrl+L/E/R align, Delete clear, F2 / double-click edit,
Ctrl+N new sheet. Double-click a sheet tab to rename it; the x closes it.

## Files (File menu)
- New Workbook (Ctrl+N), Open... (Ctrl+O), Save (Ctrl+S), Save As... (Ctrl+Shift+S).
- Workbooks are saved as `.spx` and keep all sheets, values, formulas, formatting, pictures
  and sheet protection. The title bar shows `*` for unsaved changes and you are asked
  before anything is discarded.
- Set Open Password...: encrypts the saved file (AES-256-GCM, key from PBKDF2). The password is
  asked when the file is opened. A forgotten password cannot be recovered.
- Import CSV / Export CSV still work as before.

## Pictures
- Insert > Picture from File... (or the Picture button): PNG, JPG, GIF, BMP.
- Insert > Picture from Clipboard: pastes an image copied from another program.
- Click a picture to select it, drag to move, drag the small square in its corner to resize
  (hold Shift to stretch freely), press Delete to remove it.

## Sheet protection (Review menu or the Protect Sheet button)
- Protect with a password, or leave it blank to protect without one.
- A protected sheet cannot be edited, formatted, cleared, auto-summed, filled by a macro or
  have its pictures changed or deleted, and cannot be deleted. A padlock shows on its tab.
- Unprotect Sheet asks for the password. Passwords are never stored, only a salted hash.
- Sheet protection stops changes inside the app (like Excel). To keep a file's contents
  private use Set Open Password, which encrypts it.

## Ribbon, icons and interactivity

- The ribbon has tabs: **Home**, **Insert**, **Draw** and **View**. Every button has an icon and a tooltip with its shortcut.
- Bold / Italic / alignment buttons light up to match the selected cell.
- The cell, column header and row header under the mouse get a faint highlight.
- Right-click a cell for a menu (insert function, clear, bold/italic, alignment, fill colour, fit column, insert picture).
- Menus have icons too. The window and task-bar icon is drawn in code.

## Drawing and painting (Draw tab)

Tools: Pen, Highlighter, Line, Rectangle, Ellipse, Eraser (the arrow, "Select", goes back to normal cell selection; switching to another ribbon tab also turns drawing off).

- Pick one of 8 quick colours or any colour with **More**; the slider sets the brush thickness; **Fill shapes** fills rectangles and ellipses.
- Hold **Shift** for lines at 45 degree steps and perfect squares / circles.
- The eraser removes the whole stroke or shape you touch.
- **Undo** (Ctrl+Z), **Redo** (Ctrl+Y) and **Clear All** (also undoable).
- Drawings follow the zoom level, are saved in the .spx file (format v3; older files still open) and are locked while the sheet is protected.
