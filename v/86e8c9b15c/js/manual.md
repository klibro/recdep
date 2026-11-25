# RecDep — manual

This is how RecDep is used: what each feature does and how to use it,
with references as a user writes them, keys, menus and tabs.

## 1. Model

A sheet is a grid of cells, and any cell can hold a grid of its own, so
grids nest to any depth. A cell's address is the path to it from the top
grid: `B5` is a cell of the top grid, `B5>A1` is cell A1 of the grid
inside B5. Each cell holds one Clojure form: a literal is itself, anything
else is evaluated, and a result that is a function can be called by other
cells (`B19 = (fn [x] (* x x))`, `C19 = (B19 12)`). A cell can also have a
number format (§10) and a renderer, the function that draws its value
(§4). How to write a reference is in §2.

## 2. Reference syntax

```
ref      := endpoint ( ':' endpoint )?          range if two endpoints
          | ( endpoint '>' )? $?COL ':' $?COL   whole columns, as Excel's B:B
endpoint := step ( '>' step )*                  last step must be a cell
step     := '$'                                 root grid (first step only)
          | '..'                                parent cell
          | cell
          | NAME                                a named cell (see Names)
cell     := $?COL $?ROW move*                   B5, $B$5, B$5, B5u1
          | $?COL move+                         $Au0  (col A, this row)
          | move+                               u1, u1l2, ub
move     := [udlr] DIGITS                       up/down/left/right by n
          | 'ub' | 'lb'                         to row 1 / column A  (b = beginning)
          | 'de' | 're'                         to last row / last column (e = end)
COL      := [A-Z]+        ROW := [1-9][0-9]*
NAME     := [A-Za-z_][A-Za-z0-9_?!*+'-]*         and not spelling a cell step
```

Semantics of a `cell` step: start at the referencing cell's position, set the
column/row if `COL`/`ROW` is given, then apply the moves left to right.
So `$Au0` = column A of this row, `ubl1` = row 1 of the column to the left.
A `$` before the column or row fixes it against copying (see "Copying
formulas"); in a formula it doesn't change what the reference means. In
a validation or conditional-format rule it does, as in Excel: `B$2` is
row 2 of each covered cell's column (§10, §16).

Examples:

| ref           | meaning (from cell C5 in a 10×30 grid) |
|---------------|------------------------------------------|
| `u2`          | C3 |
| `l1`          | B5 |
| `u1l2`        | A4 |
| `Au0`, `$Au0` | A5 |
| `r1:re`       | D5:J5 — the rest of the row |
| `ub:u1`       | C1:C4 — everything above |
| `$>A1`        | A1 of the root grid, from any depth |
| `..>l1`       | (inside a nested grid) the cell left of the parent |
| `..>..>l1`    | left of the grandparent |
| `B5>A1`       | A1 of the grid held in B5 |
| `B:B`, `$B:$D`| the whole of column B, of B to D: rows 1 to 30 |
| `model>B:B`   | the whole of column B of the grid named model |

A whole column (`B:B`, as Excel writes it) is read from the grid's
first row to its last, so it takes in rows the grid grows by. It stays
spelt that way: copied across, it moves a column unless fixed (`$B:$B`);
columns inserted or deleted before it move it; rows never change it;
F4 fixes and frees it. A range between two cells named in capitals
alone (`QTY:TOTAL`) reads as columns, as in Excel: write such a range
by address. Excel's whole rows (`2:5`) can't be written in a formula,
which reads `2:5` as numbers; one imported from Excel comes in as
`$A2:$A5re`.

### Names

A cell can be given a name (`subtotal`, `tax-rate`), and the name can be
used wherever a cell step can: `subtotal`, `subtotal:tax-rate`,
`..>subtotal`, `B5>total`, `$>total`.

* **Scope: the grid.** A name belongs to the grid its cell is in, and is
  unique within that grid. The same name can be reused in another grid. A
  bare name resolves only in the referencing cell's own grid; a sub-grid
  reaches its parent's names with `..>name`.
* **A bare symbol is a name ref only if that name is in scope.** Otherwise
  every symbol (`map`, `x`) would be a candidate. Paths that start with
  `$`, `..` or a coordinate are always refs, and a missing name in them is
  a `#REF`.
* **Names can't hide functions.** Refs are bound as function parameters, so
  a cell named `count` would shadow `count` in every formula in the grid.
  Names of functions the grid can call (Clojure's core, or a library the
  grid sees), special forms, ref-shaped strings (`B5`, `u1`, `re`) and duplicates are refused with a message.
  The threading macros (`->`, `some->`, `map->Foo`, …) never parse as refs.
* **Renaming rewrites what used the name.** Give a named cell another
  name (the Inspector's Name field, the name box, a parameter's rename)
  and every reference that meant it says the new name: formulas (`rate`,
  `..>rate`, `B5>rate`, `inner>rate`, a range's ends), renderers, format
  and validation rules, scenario inputs, the grid's test cases, and the
  key binding it when it's a grid's input or parameter (`{:rate C5}`).
  A `rate` that means another grid's cell, or a header, is left alone.
  It's one undo step, and on a big sheet it's prepared without freezing
  the page (§3).
  - Keywords in a formula aren't references: `(call loan {:rate 0.1})`
    keeps `:rate`, and reports the missing input.
  - If a formula has a local of the new name that would take the
    reference (`(let [rate2 1] (* rate 2))`, renaming `rate` to `rate2`),
    the rename is refused, naming the cell.
  - Taking a name away (clearing it) rewrites nothing: formulas that
    used it become `#NAME` errors.
* **Another grid's names.** A path's first name is looked up in the
  grid the formula is in: from inside a sheet, another sheet's cells
  are `$>lists>cities` (or `..>lists>…`). Written without it, the
  error offers the path that reaches it (*lists is a cell of $, the top
  grid: write $>lists>cities>city*).
* **Clearing names in bulk.** The cell menu's *Clear names* takes the
  names off every named cell in the selection, in one undo step; *Clear
  ▸ Clear all* takes them with the contents and formatting. Delete keeps
  them: it empties the cells, so a form's named input, a search box with
  a drop-down, is blank and still an input. The status
  bar says how many went and how many formulas outside the selection used
  them (they now show `#NAME`). An input bound from outside keeps its name,
  as it keeps its formula: the binding names it (§6).

### Ranges

Always 2-D: a vector of rows, each a vector of values (`nil` for empty cells),
even for a single row or column. Excel's functions take them as they are
(`(SUM A1:C3)`, `(AVERAGE B:B)`); the standard library's `cells`, `nums`
and `column` flatten them for Clojure's own.

* If an endpoint is before the start on either axis (e.g. `ub:u1` evaluated on
  row 1), the range is empty: `[]`. This makes "everything above me" work on
  the first row.
* Otherwise an endpoint outside the grid is a `#REF` error.
* Both endpoints must land in the same grid.
* **A second endpoint that is a single step continues where the first
  endpoint's path led**, as Excel's `Sheet2!A1:C1` applies the sheet to both
  ends: `H15>A1:C1` is A1:C1 inside H15, `..>A1:C3` is in the parent grid,
  `..>l1:r1` runs from left to right of the parent cell. Within one grid
  (`u1:d1`, `r1:re`) this is the ordinary meaning. A second endpoint with
  its own path (`H15>A1:H15>C1`) is resolved on its own. Copying shifts an
  inherited endpoint only along with its grid (`H15>A1:C1` → `I16>A1:C1`),
  and point-and-click writes the short form.

## 3. Evaluation

* Errors propagate: if any referenced value is an error, the cell is a
  `#UP` error naming the upstream cell, where the error started when
  that's further up, and the error itself: `A2 has an error, from A1:
  #ERR there's no item 3: there are no items`. The Inspector's *Starts at*
  row links to the cell where it started, and every cell or range an
  error's message names is a link there (`A2`, `A1`; `base (C1)` as one
  link): click your way back up the chain to where it went wrong. The
  same goes for *Step through*'s summary. (Excel's error values, `#N/A`,
  `#DIV/0!` and the rest, are a different thing: ordinary values that the
  operators and Excel's functions pass along, §17.)
* **Catching another cell's error**. A formula that only *tests* a
  cell, as the argument of `IFERROR` / `iferror` (the first),
  `ISERROR`, `ISERR`, `ERROR.TYPE`, `error?` or `error-of`, runs even
  when that cell holds an error, and the test sees it:
  * `(IFERROR A1 0)` is 0 while A1 shows `#ERR`; `(ISERROR A1)` is TRUE;
  * `(error-of A1)` says what it is: `{:code "#ERR" :msg "there's no item
    3: there are no items" :at "A1"}`, `:at` where it started when A1
    only passed it on; `ERROR.TYPE` gives Excel's number where there is
    one (`#REF` 4, `#NAME` 5, `#SPILL` 9), else `#N/A`;
  * a range is tested cell by cell by `IFERROR`, `iferror`, `ISERROR` and
    `ISERR`, so `(SUM (iferror A1:A9 0))` adds the range with its bad
    cells as 0, where `(SUM A1:A9)` is `#UP`; `error?`, `error-of` and
    `ERROR.TYPE` take one value, so a range there isn't caught;
  * a grid's cell, `(iferror E2 0)`, catches an error in the grid;
  * the formula's own bad reference is caught as well, as Excel's IFERROR
    catches `#REF!`: `(IFERROR Z99 0)` is 0, and so is a misspelt path
    such as `(IFERROR G1>ratee 0)`. The status bar still counts the cells
    where errors start, so a typo hidden this way isn't lost.

  Used any other way, even once in the same formula (`(+ A1 (IFERROR A1
  0))`, IFERROR's second argument, `IFNA`), the cell's error stops the
  formula with `#UP`. A cell that ran out of time is caught
  like any other. When the formula then fails for a reason of its own,
  `(if (error? A1) (throw (ex-info "A1 is broken" {})) 1)`, that's what
  it shows, not A1's error.
* **`IFERROR` catches a failure in its own argument**, as well as an
  error value: `(IFERROR (nth xs 5) 0)`, `(IFERROR (call E2 {:idx 5}
  :pick) "none")` give the fallback instead of `#ERR`. Not a formula
  running out of time (`#TIME` stays: the time is spent). (A cell it
  reads is caught too: see *Catching another cell's error*.) `ISERROR` is a
  function, so its argument fails before it runs: `(ISERROR (nth [] 3))`
  is `#ERR`; use `IFERROR` (or Clojure's `try`) around a part that may
  fail.
* **Excel's error values say why.** Hovering a cell showing `#N/A`,
  `#DIV/0!` … says why it was made, where RecDep knows, else what
  the code means, and whether it was made there or read from a cell
  holding it; so does the Inspector, under the value, and Steps' summary:
  * `#VALUE!` from an operator or a function reading a number: "the text
    "abc" isn't a number"; from a condition: "the text "maybe" isn't
    TRUE or FALSE";
  * `#DIV/0!`: "12 divided by 0"; `#NUM!`: "the result isn't a finite
    number";
  * the lookups: `#N/A` "the text "cherry" isn't in the first column (4
    rows)" (VLOOKUP, HLOOKUP, MATCH, XMATCH, LOOKUP, XLOOKUP; "nothing at
    or below …" for an approximate match), `#REF!` "column 7, but there
    are 2 columns" (VLOOKUP, HLOOKUP, INDEX);
  * a formula that failed for a reason of its own while one of its inputs
    held an error value shows that value, as Excel does, and adds "the
    formula failed: there's no item 3 …".

  The value is still the same `#N/A`: IFNA, ISNA and `=` see no
  difference. A reason names values, not cells (the functions get
  values); Step through shows the part that made it.
* **Error codes.** A cell that can't give a value shows a code, with the
  reason on hover and in the Inspector. When the reason is in one part of
  the formula, the formula bar underlines that part in red, and the hover
  and the Inspector's *At* row name it (`at (nth xs n)`; a misspelt name
  is the name itself). When it happened inside a library function, they
  say which, and where: `in third-item, line 2 of the library of picks (E2)` (the
  Inspector's *In* row): the library function the line belongs to, even
  when the error happened in a core function it called (`into`, `vec`,
  a `for`) or in an anonymous `fn` inside it. Some errors have no part: one in a core function
  run lazily after the formula returned, `(map first [1 2])`, or one in
  the whole formula. The codes:
  * `#PARSE`: the formula can't be read (a bracket not closed, `+ 1 2`)
    or is malformed: `(let [x] x)`, `(if)`, `(fn)`, a library macro
    refusing what it was given.
  * `#NAME`: a name that isn't there: a typo, a function from another
    grid's library, `Math/sqrt`.
  * `#REF`: a reference off the grid, or a `REF!` left where a reference
    was deleted or moved off the grid by a copy (§5).
  * `#ERR`: the formula failed while running (`(nth [] 3)`, a `throw`).
  * `#CYCLE`, `#TIME` (§8), `#SPILL` (§13), `#UNIT` (§25), and `#UP`:
    an error passed on from a cell it reads, saying where it started.
* **Messages in plain words.** JavaScript's and ClojureScript's own error
  texts are rewritten to say what was wrong, with which value, and what
  to write: `(A1 + A2)` with 5 in A1 says "A1 holds the number 5, which
  isn't a function: in Clojure the function comes first, so write (+ A1
  A2)"; `(nth A1:A2 5)` says "there's no row 5: the range A1:A2 has 2
  rows, numbered 0 to 1 (a range is a vector of rows)"; `(count 5)`
  says "count needs a collection …, got the number 5"; `(+ 1 2` says
  "the ( at column 1 is never closed: add a )". The original text is in
  the Inspector, folded under *Details*. A message RecDep doesn't know
  is shown as it came, and your own `(throw (ex-info "…" {}))` text is
  never changed.
* **Argument counts**. A call with an argument count its function
  doesn't take is refused when the function is Excel's, RecDep's (`call`,
  `what-if` …) or a library's: `(tax 100 0.2)`, where `tax` takes `[x]`,
  is `#ERR` "tax takes (tax x), not 2 arguments", the call underlined,
  rather than quietly ignoring 0.2; `(VLOOKUP "a" A1:B1)` says what
  VLOOKUP takes. Excel's optional arguments count as optional: any number
  from the shortest form to the longest. Clojure's own functions are only
  marked (§26): `(inc 1 2)` shows 2. Not checked: functions passed as
  values (`(map f xs)`, `apply`), a cell holding a function, locals, and
  calls inside macros other than a few of Clojure's (`when`, `cond`,
  `and`, `or`, `do`, `->`, `->>`) and Excel's (`IF`, `IFERROR`, `IFS` …):
  not inside `cond->`, `doto` or `lazy-seq`, say. Excel's functions say
  how many they take ("SORT takes 1 to 4 arguments, not 5").
* **Formulas written without brackets.** A cell holds one form, so `+ 1
  2`, `A1 + A2` or `SUM A1:A9` (Excel's way, without the `=`) is `#PARSE`
  and says the call to write: `(+ 1 2)`, `(+ A1 A2)`, `(SUM A1:A9)`; as
  `SUM(A1:A9)` says `(SUM A1:A9)`.
* **Did you mean.** A name that isn't there is checked against the cell
  names in scope and everything the formula can call (its libraries, the
  standard library, Excel's functions, Clojure's), ignoring case: the
  hover and the Inspector add "did you mean total?" (`totl`), "did you
  mean SUM?" (`sum`, `SUMM`), "did you mean map?" (`mpa`). Only close
  names are offered (one letter off for short names, two for longer);
  a cell named afterwards is offered too. When the name is defined in a
  library that failed to load, the message says so instead: "tax is
  defined in the top grid's library, which failed to load: line 2: …
  Fix that first". The same goes for `call`'s cell names (`:rat`, did you
  mean `:rate`?), a renderer's names, a format rule's (`poss?`, did you
  mean pos? on the Formats panel), and Excel formulas typed with `=`
  (`=SUMM(…)`, did you mean SUM?). The hint is on the cell itself, not
  repeated in the `#UP` of the cells that read it.
* **Operators are Excel's.** RecDep's own `+ - * /` and `< > <= >=` read
  values as Excel's operators do: an empty cell is 0 (or `""` next to
  text), text holding a number is that number, other text is `#VALUE!`,
  an error value passes through, dividing by zero is `#DIV/0!`, a result
  that isn't finite is `#NUM!`, and a range works element by element
  (`(* A1:A3 2)` spills). Two numbers take a fast path, so on numbers they
  are Clojure's (and as fast in a loop), and Clojure's arities stay: `(- x)` negates, `(/ x)` is
  1/x, `(+)` is 0, `(< a b c)` chains. Excel's `=` is `==` (text ignores
  case, an empty cell equals 0, `""` and FALSE), its `<>` is `<>`, its `&`
  is `&` (joins as Excel writes values). `=` stays Clojure's equality,
  because vectors, maps and keywords need structural comparison and
  `(== A1:A9 "x")` has to broadcast. `^` is the reader's metadata
  character, so Excel's `^` is `POWER`; its postfix `%` is `(/ x 100)`.
  The standard library runs on these operators too.
* **Clojure's other number functions are Clojure's**:
  `inc`, `dec`, `max`, `min`, `abs` don't read values as Excel does, so
  `(inc "5")` is `"51"` (JavaScript joins the text) and `(max "x" 3)` is
  3. For a cell that may hold text or be empty, use `(+ A1 1)` or
  Excel's `MAX`, `MIN`, `ABS`, which read it as the operators do.

* **Calculating while you work.** An edit that changes many values (one
  that tens of thousands of cells depend on, the first unit typed into a
  sheet, F9, a tick) doesn't freeze the page: the cells are calculated in
  short slices, and you can scroll, select and keep editing meanwhile.
  - A cell not yet calculated shows its **old value, dimmed**; it
    brightens once calculated. The status bar says **Calculating… 12,000
    cells left**, and counts down a few times a second.
  - The cell you typed into shows its new value at once.
  - The selection's count, sum and average, and the Inspector's value,
    are dimmed too while any of their cells is: they're from the old
    values.
  - **Edit, undo and redo as usual.** An edit made while calculating
    joins in: what it affects is calculated as well. Undo takes you back
    to the sheet as it was before the edit, still calculating if it was.
  - What needs values works on calculated ones: copying, Paste values,
    Make table a sub-grid, sort, filter and group views, goal seek, test
    cases, scenarios, validation of what you type, Save grid. They
    calculate what they read first, at once.
  - A conditional format follows the value shown: a dimmed cell keeps its
    old format until calculated. A cell that ticks (§28) waits for its
    next tick while the sheet is calculating.
  - On a small sheet it's all over before the page redraws, and nothing
    is dimmed.
* **Automatic or Manual.** The status bar's **Auto ▾** chooses how the
  sheet calculates, saved with the sheet. *Automatic* calculates every
  change, as above. *Manual* is for a heavy sheet, when you have several
  things to change and only want the answer once they're all in:
  - a change works out only the cells changed: what you type shows as
    typed, and a formula you type shows its value (with any waiting
    cells it reads worked out first). Everything that reads them waits,
    showing its old value, dimmed, and the status bar says **Manual ▾
    4,672 cells to calculate · F9**;
  - **F9**, or a click on that count, calculates everything waiting,
    with the page answering meanwhile (*Calculating…*); choosing
    *Automatic* again does too. F9 also asks the sheet's queries and web
    requests again, as it always does, and what reads them is calculated
    when their answers come in;
  - cells that update on their own (a tick, a schedule, a web answer
    arriving) wait for F9 as well;
  - operations that need values (copying, sorting, goal seek…) still
    calculate what they read first;
  - opening a sheet saved in Manual calculates it, as any sheet is, and
    Manual applies to your changes from then on.
* **Loading and preparing.** Opening a sheet, changing a library (or My
  library, or whether cycles iterate), inserting, deleting or moving
  rows, columns and cells, and renaming a cell that many formulas use
  don't freeze the page either. Opening and a library change compile
  every formula again; inserting, deleting and moving update only the
  formulas they touch (their references rewritten, nothing compiled
  again unless it must be). This is done first, in short slices; meanwhile
  the sheet you had stays on the page, and the status bar says
  **Loading… 40%** (opening) or **Preparing… 40%**.
  - **What you do meanwhile waits for it.** Edits, undo and redo, and
    Save are made in order once the new sheet is ready, each an undo step
    as usual; the status bar says how many are waiting ("· 2 waiting").
    So a value you type while a sheet opens lands in the opened sheet,
    as its first undo step (opening isn't one). The same goes for what you
    do while the file is being read, before *Loading…* starts.
  - An edit waits with the address you gave it: typed into E50 while a
    row is inserted above it, it goes into E50 of the new sheet, which is
    the row below the cell you typed over. A fill or a paste made
    meanwhile writes what it took from the sheet showing.
  - Copying, Paste values, sorting and other operations that read values
    finish the preparing at once first (the page waits for them), since
    they'd otherwise read the sheet as it was before.
  - Then the sheet is calculated as above: after opening or a library
    change every cell shows its **value from before, dimmed**, until it's
    calculated again. After inserting, deleting or moving, cells keep
    their values, moved with them, and only what the change affects is
    calculated again (dimmed until then): a total over a range the new
    row falls into, a formula that now reads other cells, one that asks
    where it is (`here`, `down`), one whose reference was deleted.
    Histories (§28) move with their cells too.
    **The cells on screen are calculated first**, then the rest; scroll
    elsewhere meanwhile and the cells you scrolled to come next.
  - On a small sheet it's all over at once: nothing says *Preparing…*.
* **Opening a sheet shows its values at once.** Save keeps the values
  the sheet had calculated in the file, so Open shows them, dimmed, as
  soon as the file is read: the sheet you opened is on the page while
  *Loading… 40%* compiles its formulas, instead of the one before it. At
  300,000 cells they show after about 4 s (30,000 cells: half a second).
  - The values shown are what the file says, not results: each cell
    brightens once it's calculated again (the cells on screen first). Until
    then nothing uses them: formulas, the selection's sum, copying, goal
    seek, the status bar's counts (they show **…**) and Save all wait for
    calculated values.
  - What you do while it loads waits, as above, and lands on the opened
    sheet.
  - Drawn plainly at first: a cell's renderer (a bar, a chart) appears
    once the formulas are compiled.
  - A cell shows **blank, dimmed** if its value wasn't kept: it was still
    calculating when the file was saved, or its value can't be written
    down (a function, an endless sequence, a JavaScript object, an Excel
    error value). A grid made by a formula (§22) appears once its formula
    is calculated.
  - Saving straight after opening, before anything is calculated, keeps
    no values (they'd be the file's word, not calculated): the file still
    opens, its cells blank until calculated. A file saved without values,
    or by **Save grid**, leaves the sheet you had on screen until the new
    one is ready.

## 4. Renderers

**Units in a renderer's options** are as in a formula: a plain number is
in base units, a quantity is written as text, `(data-bar "50 km")`, and a
bound read from cells (`(MAX B2:B9)`) is already in base units, whatever
the cells show. Controls work the same way (`(slider "0 km" "20 km" "0.5
km")`), though without a step they step by one of the cell's unit (a
km, not a metre), and `number-input` shows and takes the number in the
cell's unit; a validation rule's bounds read quantities too, `(between
"0 km" "10 km")`.
What a control writes keeps the cell's unit: `set-by-ref` and
`update-by-ref` take numbers as values (base units) and write them in the
cell's declared unit, or with the unit its value was typed in (`7 km`),
so `(update-cell! ctx 'l1 #(+ % 1))` on a km cell adds a metre rather than
rereading 5001 as kilometres. `duration` takes a cell with a time unit
(s, min, h) as that time, a plain number as seconds; `file-size` takes
bytes, which is what a data unit (B, kB, MB) holds.

The predefined set lives in the **standard library**, shared by every
sheet (§14); a grid's own library can add more. Cell formulas can use
library functions too. Renderers and renderer makers are marked
`^:renderer` (`(defn ^:renderer plain …)`), which is how the Renderer
field's completion tells them from other functions (§12).

**A renderer's options can read cells**: `(heat 0 (MAX B2:B20))` colours
each cell against the column's largest value, a heat map over cells that
each hold one number. As in a validation rule (§16), the references are
parameters: the source compiles to a maker, `(fn [B2:B20] (heat 0 (MAX
B2:B20)))`, and each cell drawn with it binds them from where it is, so in
a format rule `(data-bar r1)` reads each cell's right neighbour. A
reference is a free symbol spelled like one (`B5`, `l1`, `B2:B20`,
`$>B6>C9`, a cell's name) that names nothing in the renderer's namespace;
quoted refs (`'l1`, for `set-cell!`) aren't references. A referenced cell
with an error draws `#RENDER` saying which; a bare word that's no cell's
name reads "Could not resolve symbol", as a misspelt function would.

**Lists and maps as values.** A cell holding a flat list of values shows
them as a list, each in the cell's number format: a list of dates with a
date format reads *18 Jan 2027, 2 Apr 2027*, not `[46405 46479]`. A map
is drawn as a table of its keys and values (`key-values`), nested maps
and tables inside it as tables; give it the renderer `plain` to see it
as Clojure. Hover the list for it as Clojure. Numbers in a map all take
the cell's number format: a map of dates reads as dates with a date
format on its cell.

### The shelf: what there is, where you'd look for it

What a selection can be **in its own cells** is in the popover of each
field; what it can be **as a new cell** is in the side panel's Shelf
tab:

* **Renderers**: the Renderer field's gallery, refiners and all
  (below; the cell menu's *Show as…* opens it on a selection).
* **Number formats**: the format field's gallery (§10; *Number
  format…*).
* **Functions** of a range (a total, an average, a library's
  `^:summary` …), **views** of a table (`^:view`), **charts** of the
  selection, and **templates** (grids put on the shelf, their inputs
  wired to the selection, showing the display area their author
  marked): the Shelf tab, each placed in the grid.
* Both galleries and the tab find things by meaning (*heat map*, *dial*,
  *euro*, *pivot*, *mean*), from each piece's own metadata (`:accepts`,
  `:suggest`, `:aka`, `^:view`, `^:summary`).
* My library and the shelf go to disk, and to other browsers, as one
  file (the library file, below).

* **Units**: the unit field's popover (§25), units found by search and
  drawn with the cell's value, your own defined, changed and removed
  there.

### The gallery: choosing a renderer by how it looks

A renderer's whole point is how it looks, yet the Reference and the
Renderer field's completion only name them. The gallery shows them: the
active cell drawn by every renderer that suits its value, each with
options worked out from the cells around it. It's the first part of the
shelf (above): renderers, surfaced visually.

**Using it**

* **Focus the Renderer field** and the gallery opens under the formula
  bar, on the right (the grid's selection stays in sight): *Default*
  first, then the grid's own and My library's renderers, then the
  standard ones by section (*Numbers and text*, *Indicators*,
  *Collections*, *Charts*, *Interactive* …). Each tile is the cell's own
  value drawn as the grid would draw it, with its number format and unit,
  at the column's width, the source written under it.
* **A selection of several cells** is drawn as a strip: each tile shows
  its first rows and columns (up to 6 × 3), each cell with its own value,
  so `heat` over a column reads as a heat map and `data-bar` as bars of
  different lengths. It's judged by what **most** of its cells hold, not
  by the active cell: a column of numbers selected from its heading, or
  by its letter, is numbers (the heading is drawn too, as it would be).
  The gallery says which cell it judged by when that isn't the active
  one. *Indicators* come first, since several cells are most often
  selected to compare them.
* **Hover a tile** and the selection's cells are drawn with it in the
  grid, until the pointer leaves. A formula's result that spills is drawn
  whole, as a renderer of its own would keep it in its cell.
* **Click a tile** to write its source into the field, which keeps the
  focus, with its first option selected for typing over (`(MAX B2:B5)` in
  `(data-bar (MAX B2:B5))`). Enter applies it as typing would, and so
  does a click on the tile the field already holds (the one just
  clicked, or the cell's own); Esc leaves the cell as it was. Choosing
  *Default* empties the field. A renderer that can be refined opens its
  refine panel in the gallery's place (below).
* **Typing filters** the tiles by the renderer's name while only a name is
  written (`dat`, `(st`); once the call has options, every tile shows
  again, the one written marked. What the field held when it was focused
  shows every tile, its own marked.
* **Only what suits the value** is offered:
  * by the **value's shape**: a number gets number formats, indicators
    and controls; text gets `markdown`, `badge`, `truncate` …; a list,
    a range or a map gets `bars`, `sparkline`, `pie`, `bullets` … and
    Vega-Lite charts (below); a table of row maps, or a range under a
    row of headings, gets `table` and charts of its columns; a grid cell
    gets `grid-link`, `grid-view`, `key-values`. A number in a date
    format gets `date-input` and `date-time-input`, in a time format `time-input`.
    A `#SPILL` cell is judged by the result it couldn't lay out, since a
    renderer keeps that in its cell and so clears the `#SPILL` (§13):
    `grid-view` comes first, and a line under the gallery's heading says why.
    Text (or a number) also gets a `button` that adds 1 to the cell to
    its right and a `counter-button`, and a `dropdown` and `radio` whose
    options are the values of its kind around it (the column's, or the
    selection's: a heading left out), else the value and two more to
    edit, else the cell's validation's;
  * by the **selection's shape**: comparisons (`data-bar`, `heat`,
    `traffic-light`) need something to compare with, a selection of
    several cells or a cell in a column of numbers; showcases of one
    value (`kpi`, `gauge`) aren't offered for several cells;
  * **controls** (`slider`, `stepper`, `number-input`, `toggle`,
    `dropdown` …), which write the cell's own value, only when the cell
    holds a value typed in (in a selection, every cell of the kind
    judged by): on a formula they'd replace it;
  * a renderer that draws the value just as `plain` does, or fails, is
    left out, and a few only show where they change something
    (`compact` from 1000). Nothing that would only do a number format's
    job is offered: those are the format bar's.
* **Options from the cells around.** Indicators read the numbers of the
  selection, or of the run of numbers in the active cell's column up and
  down from it, as references, so the renderer follows them:
  `(data-bar (MAX B2:B9))`, `(heat (MIN B2:B9) (MAX B2:B9))`, `(gauge 0
  (MAX B2:B9))`, `(progress)` for fractions. Alone, a number gets round
  bounds (`(gauge 0 50)`). `(kpi "Revenue")` is captioned by the cell's
  name or the text to its left; `slider` takes the cell's validation
  bounds when it has some (§16), `dropdown` and `radio` its options.
* **Vega-Lite charts** (`chart`) are written for the value's columns: a
  table's, or a range's under its headings, named as `chart` keys them
  (`Month` → `:month`). The first column that isn't numbers goes across
  and the first that is goes up, as columns (`(chart :bar :month
  :sales)`), a line, and a pie when there are at most 8 categories and
  no negative parts; a second text column of 2–8 values colours the
  columns (`{:color :region}`); two number columns and no text make a
  scatter (`(chart :point :w :h)`). Categories not already in order keep
  the rows' (`{:x {:sort nil}}`, and `{:color {:sort nil}}` for a pie's
  legend), since Vega-Lite would sort them: months stay Jan, Feb, Mar. A
  map {label number} gets columns and a pie (`(chart :bar)`, `(chart
  :arc)`), a list of numbers a line and columns by position (`(chart
  :line :index :value)`). Chart tiles are taller (190 px), so their axes
  show. In the grid a chart is 180 px tall unless its row has a height
  or it spans cells (§4), which the preview shows as it is.
* **Big values** are drawn from their first 60 items (the gallery says
  so); the grid draws them all.

### Ways in, and finding by meaning

The galleries open from their fields: the Renderer field's, the format
field's. Two ways in that don't depend on clicking a field, and a
search that finds things by what they do:

* **The cell menu**: *Show as… (renderers)* opens the
  renderer gallery on the selection, browsing (not refining what the
  field holds: *All renderers* already done); *Number format…* opens the
  format gallery. Each focuses its field, which opens its popover, so
  the popover is the same as from the field, refiners and all.
* **Each gallery has a search box**, its field being for code (a
  renderer call, a format code): the renderer gallery's finds renderers
  by name, the words they're known by, section and docstring (*heat map* or *colour scale* finds `heat`,
  *dial* `gauge`, *rating* `stars`, *arrow* `trend`); the format
  gallery's finds formats by caption, section and code (*euro*,
  *percent*, *date long*, *thousands*). Enter or Esc in a search box goes
  back to the field, where they use or cancel as ever.
* **Typing in the Renderer field** is code, and completes names: the
  gallery narrows to the renderers whose name begins with what's typed
  (`dat`), as completion does; it doesn't search.

### Refining a renderer: variations on a theme

Choosing a renderer is the start: `data-bar` against what, showing the
number how; a chart of which columns, summed or averaged, in what
order. Once the Renderer field holds a call of a renderer that can be
refined, the gallery gives way to its **refine panel**: rows of
variations on that call, each drawn with the selection's own values, and
controls for what is dialled in or named rather than picked.

**Using it**

* It opens when a tile is chosen, when a refinable call is typed, or
  when the field already holds one (a cell with a chart opens straight
  into it). *← All renderers* goes back to the gallery.
* At the top, the field's call drawn on the selection, *use this*: a
  click (or Enter) applies it. Under it, a row per thing to change:
  * **variations** (the current one marked): `data-bar`'s *Full at* (the
    column's `(MAX …)`, round numbers), *Empty at*, *Number* (`plain`,
    `compact`); `heat`'s *Colours*; a chart's *Kind*
    (columns, line, area, points, pie, donut) and *Order*; a button's
    *When clicked* (add 1, take 1, set to 0, set to today, switch on and
    off). **A variation is drawn only where it looks different**: those
    that would draw alike (a button's actions, `plain` and `compact`
    on 40, a bar full at 1 or at the column's maximum of 40) are offered
    by name instead, as small labelled options under the tiles, and so is
    the whole *Wrap in* row, whose tooltip doesn't change the look;
  * **controls**: a number (a step, with suggestions in its placeholder;
    with a slider when the values around give a
    range), text (`kpi`'s caption, a chart's title, a drop-down's
    options as a list, `Low, Medium, High`), a select (a chart's columns,
    summary, colour; the cell a button changes), a tick (dates across, a
    zero baseline, points on a line).
  * **Wrap in**, for every renderer that draws text: none, or value on
    hover (`tooltip`). Refining a wrapped call refines the renderer
    inside, keeping the wrapper. Negatives in red, centred, bold, zero
    as blank aren't wrappers: they're the cell's number format and style
    (§4), set in the format bar, and every renderer follows them.
* Each change is written into the field (the call stays readable and
  typeable) and the grid draws the selection with it, live, before it's
  used: Enter (in the field or a control), a click on the top tile, or
  focus going anywhere else uses it, as one undoable change; Esc puts
  the cell back.

**Writing refinements.** A renderer says how it's refined in its
metadata, beside its definition (the standard library's do, in the attr
map after the docstring), so a library's own renderers take part:

* **Parameter specs**: `{:refine [{:param :hi :label "Full at" :kind
  :bound :end :hi} {:param :lo :kind :bound :end :lo :default 0} {:param
  :inner :label "Number" :kind :renderer :default 'plain}]}`. A spec
  names a parameter of the renderer's arglists (`:params [:from :to]`
  for several set together), with a `:kind`:
  * `:choice` — `:options [[value caption] …]`, `nil` for leaving it out;
  * `:bound` — a range's end (`:end :hi` or `:lo`): the selection's `MAX`
    or `MIN` as references, and round numbers around the value;
  * `:renderer` — another renderer drawing the value (`plain`,
    `compact`; decimals, percent and currency are the cell's format);
  * `:number` and `:int` (`:min :max :step`; with neither bound given,
    the range of the values around, and a slider beside the box;
    `:placeholder` suggests values, `"1 (or 0.5, 0.1, 0.01)"`, emptying
    the box leaves the parameter out), `:text`, `:toggle`, and `:list` (a
    vector, typed as `a, b, c`).
  Where the variations would look alike, prefer a control: a slider's
  step is a number box, not five sliders that differ only in their
  caption.
  `:fill` gives a value to write when a parameter is needed but not set:
  setting `slider`'s *From* alone writes `(slider 2 100)`, its *To* filled
  with 100.
  A changed parameter is written back by the shortest arglist holding
  every parameter given, the rest filled in by their `:default`s, and a
  parameter equal to its default is left out: `(data-bar 0 100 plain)`
  is written `(data-bar 100)`.
* **A refiner of its own**, for a renderer the specs can't describe:
  `{:refine my-refine}`, a function of the call (as a form) and `info`
  (`:value`, the selection's numbers as `:range {:range "B2:B9" :lo
  :hi}`, `:label`, `:format`, `:domain`, `:choices`), giving rows.
  A row is `{:label … :choices [{:caption … :form …}]}` (drawn as
  tiles where they look different; `:named? true` offers them all by
  name) or a control with `:set`, a function of its new value giving
  the new call: `{:select {:value v :options [[v caption] …]}}`, `{:text
  {:value s}}`, `{:number {:value n :min :max :step}}`, `{:toggle
  {:value b}}`. `chart` (`chart-refine`) and `button` (`button-refine`)
  are refined so, in the standard library, as worked examples.
* **A refinement UI of its own**: a refiner's row can be `{:label …
  :ui (fn [call set-call!] hiccup)}`, drawn as it is (in an error
  boundary), `set-call!` writing a new call into the field: buttons, a
  palette, a little editor, whatever the renderer needs. A refiner runs
  within a time limit (2 s).

### Charts and views dropped on a selection

A renderer draws one cell's value, so a column of numbers can be given
`heat` or `data-bar`, cell by cell, but not one chart of all of them; nor
can a table be filtered, grouped or cross-tabulated where it is. Each is
a new cell: its formula reads the selection, its renderer draws the
result, and it spans a block of empty cells for room. Neither is a
renderer of the selection's cells, so the renderer gallery doesn't offer
them: the side panel's **Shelf** tab does (with templates, below), and
placing one is the drop.

**Using it**

* Select several cells and open the **Shelf** tab (or right-click →
  *Functions, views, charts and templates…*), which follows the selection: *As one
  chart* offers the charts of the selection, drawn from its values:
  * a table under a row of headings (`Month`, `Sales`): its columns, as
    the gallery charts such a value (§4): columns, a line, a pie (a few
    positive parts), columns coloured by a second text column, a scatter
    of two number columns;
  * a column of names beside a column of numbers (`apple 1`, `pear 2`),
    with headings or without: columns, a line and, for a few positive
    parts, a pie with a legend naming them, in the rows' order; and
    labelled bars (`bar-list`), reading the rows under the headings;
  * two columns of numbers (x beside a y worked out from it), with
    headings or without: the second against the first, as a line (first
    when x rises row by row) and as points, the axes in whole steps when
    x is whole numbers;
  * a first column (numbers, or names) beside several columns of numbers,
    with headings or without: one chart with a series for each, first
    (*lines of sales, cost*): against numbers as lines and points,
    against names as lines and columns side by side. The legend names
    the series by their headings, in the columns' order; without
    headings they are y1, y2 …. Up to eight series;
  * one column (or row) of numbers: a line and columns by position, with
    axes, titled by its heading when it has one (the heading isn't
    charted), and the standard small `sparkline` and `bars`. No pie: one
    of numbers by position names nothing;
  * a column of dates (§15) beside columns of numbers: the dates across,
    first, the numbers up (*lines of qty, price* over the dates), as a
    line, points, stacked areas and columns; with one column of numbers
    and a text column beside them, a line a colour each (each region
    over time). Then by the text column, as above. A date is never one
    of the series. A column of dates alone has nothing to chart over
    time: the tab says to select it with the numbers beside it.
  A whole column selected by its letter is read to its last filled row.
  Nothing is offered for one cell, text alone, or a first column mixing
  numbers and names; for several cells the tab says why, and what would
  give a chart (the cell holding text among the numbers, a row of
  headings, one column of numbers).
* `chart` takes several fields up as a vector, `(chart :line :x [:sales
  :cost])`: a series each.
* A chart's colours and legend are options: `{:colors ["#08519c" …]}`
  colours the slices or series in the legend's order (a chart without a
  legend takes the first as its colour, over the cell's text colour);
  `{:legend :bottom}` (`:top`, `:left`, `:right`) places the legend,
  `false` hides it, and a map is Vega-Lite's legend settings (`{:orient
  :bottom :title "Fruit"}`); `{:names {:y1 "Square"}}` renames series.
  The refiner sets them: *Palette* (the chart colours, blues, greens,
  warm), *Legend* (right, bottom, top, left, none), *Legend
  title*, and for series *Series colours* (a colour picker each) and
  *Series names* (a box each, blank for the column's own); a chart with
  no legend has *Colour* instead, a picker and *the theme's* to go back.
* A chart's refiner (§4) keeps to what makes sense: pie and donut are
  among its kinds only over named parts, none below zero (not numbers by
  position, not several series),
  switching between a pie and columns keeps the rows' order, and a pie,
  coloured by what's across, has no *Colour by*, nor do series, coloured
  one by one. *Up* offers one column, or each of the others as a series.
* Above the charts, *As a view of its table* offers the views (§20) of
  a table: of a selection under a row of headings, read with `as-table`
  (`(filter-view (as-table $A$1:$C$9))`); of a selection without
  headings, every row data, its columns named by their letters in the
  formula (`(filter-view (as-table $C$2:$E$9 [:C :D :E]))`: type names
  over the letters, `[:month :region :qty]`, to read them by name,
  `v>qty`, since `v>C` means a cell named C; until then read it by
  position, `v>B2:B9`; a column inserted to the left moves the range
  and keeps them); one
  column is enough; or of a single cell holding a table (a grid with
  headers, rows as maps), read by its name (`(filter-view orders)`).
  When the first row may be headings or data (text over text, or text
  with a blank among it) both are offered, each tile saying which (*·
  first row as headings*, *· no headings*), and a selection holding
  something that no view works on says so. A long table's tiles are
  worked out from its first 1,000 rows, a good indication of what you'll
  get, and the tab says so above them (*These previews use only the
  first 1,000 of 30,000 rows …*): a placed view reads every row.
  Offered: every view in scope, marked
  `^:view` (§22): the standard *filter and sort*, *group and summarise*,
  *cross-table*, *one row at a time*, *each column summed up*, then those of the grid's libraries and My library (the
  tour's tables section offers its *order chart* and *top orders*). A
  view that needs columns the table lacks (its `:needs`, §22), that
  fails on it, or that gives no grid isn't offered; under the views,
  *n views aren't offered for it: why* lists those one column short and
  those that failed on it, each with its reason (*gantt needs 2 date
  columns (start, finish); it has 1 (Begin)*), and when none is offered,
  the tab says so with the same reasons. Each tile is the view drawn as its grid will be,
  from its layout: a panel's buttons, drop-downs and chart, a view's
  header controls (still, as in every tile). A view is drawn in place,
  `(grid-view)`, its controls its own, and spans its own size: a row per
  row of its grid, a row with a height of its own as many as it takes (a
  panel's 150 px chart takes seven), and one more (a header, a name over
  a drop-down, is two lines tall), at most 12 × 6, a bigger one
  scrolling inside.
* Choosing one **places** it: the block it will
  span (a chart's 8 rows and columns enough for about 360 px; a view's
  own size) is outlined in the grid, first at the suggested place, to
  the right of the selection a column apart (else touching it, else
  below it, else the first free block, else, in a grid with no room
  left, below everything in it, the grid growing to fit), then under the
  pointer. Its place and its size are both yours:
  * a click puts it there at that size; **a drag marks out the block**
    corner to corner, any size, in any direction;
  * the arrow keys move the outline, **Shift+arrows resize it** (down
    and right grow it, up and left shrink it); Enter places it;
  * Esc doesn't place it. A block that isn't empty (a value, a grid, a
    span over it) is outlined in red and refused; one running past the
    grid's edge grows the grid when it's placed, as the views of §20 do.
  A banner under the formula bar says what to do and the size it will be
  (*Place the chart, 10 × 4*).
* The new cell holds the range, fixed (`$A$1:$B$5`, so it can be moved
  and copied and still read the same cells), its renderer the chart, and
  its span the block: one undoable step. It's selected, a chart opened
  in refine mode, to adjust at once (a view is adjusted by its own
  controls).
* `as-table` turns a range under a row of headings into a table, a row
  map per row keyed by its headings as a grid with headers keys its rows
  (`"Unit price"` is `:unit-price`), rows holding
  nothing left out: the views and table functions (§20) take a range so.
  Given keys, one per column, it reads a range without headings, every
  row data: `(as-table C2:E9 [:C :D :E])`. Too few or too many keys, or
  one twice, is an error saying so. A range's numbers can't say they are
  dates, so the columns holding dates are an option, last:
  `(as-table A1:E9 {:dates [:date]})`. The Shelf writes it for a column
  whose kind is dates or whose cells show dates; the views then show
  them as dates and don't add them up. (A grid with headers knows its
  columns' kinds: nothing to write.)

### Functions dropped on a range

Σ writes a SUM under a column; the other common questions of a column
(its average, how many, the biggest, the middle) are the same move with
another function, and a library's own (a spread, a weighted mean) are
too. The Shelf tab offers them, each with its answer for the selection.

**Using it**

* Select a column of values (from its heading, or the whole column by
  its letter), a block of columns, or a row (a label first, maybe), and
  open the **Shelf** tab: *As a function of it, below* (or *to its
  right*, for a row) offers SUM, AVERAGE, COUNT, MIN, MAX, MEDIAN and
  STDEV.S for the columns holding numbers, COUNTA (*count of entries*)
  for every column, then the library functions marked `^:summary`. Each
  tile shows its results, one per cell it writes, in the number format
  they'll have.
* **A block goes either way.** For several rows and columns the section
  has a switch, *↓ below each column* (first) and *→ right of each
  row*: by row, each data row gets a result in the column right of the
  block, the heading row left out, and a first column of labels (text
  in every row, values beside it) left out of what each reads. A single
  column goes below, a single row to the right. The switch is kept for
  the next block. (Σ sums by column, as Excel's does.) The search finds them by what they do (*total*, *mean*,
  *biggest*, *how many*).
* A click **places** it, as views and charts are placed, but its cells
  keep their shape: the outline is exactly the cells it writes (under
  each column of numbers, say, not under a text column between them),
  suggested under the columns in the row below the data (right of a
  row), else, when a cell there is taken, the first free place; a click
  or a drag puts it where the pointer is let go, arrows move it, and
  Shift+arrows and dragging don't resize it. Enter places it; the cells
  written are selected; one undo takes them away.
* Each cell is the function of its column's data, headings left out:
  `(SUM B2:B9)`, its references written from where it goes in the
  formula bar's style (as Σ writes them; placed elsewhere, it reads the
  same cells). A function that keeps its column's kind (SUM, AVERAGE,
  MIN, MAX, MEDIAN, STDEV.S) gives its cell the column's number format
  when it wouldn't have it already: a total of prices is a price.
* **Dates** (a column whose kind is dates, or whose cells show dates,
  §15): MIN and MAX (the earliest, the latest), MEDIAN and AVERAGE, each
  a date in the column's format (`yyyy-mm-dd` when only its kind says
  dates), COUNT and COUNTA. Never SUM or STDEV.S: a sum of dates is no
  date. On a block, SUM goes under the other columns only. By row, a row
  mixing dates and numbers is only counted. Your own `^:summary`
  functions are offered on numbers, not dates. For the span in days,
  write `(- (MAX C2:C9) (MIN C2:C9))`. Σ leaves dates out too, and on
  dates alone says so.
* **Your own:** a function marked `^:summary` in a grid's library or My
  library is offered too, called as a formula calls it, `(spread
  B2:B9)`: given the range as rows of values (`[[10] [20] [30]]`, as
  any formula is; the standard library's `cells` flattens them), its
  `:caption` naming its tile and `:aka` words found by the search.

  ```clojure
  (defn ^:summary spread "The largest less the smallest." {:caption "spread"}
    [xs] (let [ns (filter number? (cells xs))] (- (apply max ns) (apply min ns))))
  ```

### The Shelf tab: the selection as a new cell

What a selection can be *in its own cells* is a renderer (the renderer
gallery) or a number format (the format gallery), each in the popover of
its field, refiners and all. What it can be *as a new cell* (a view of
its table, a chart of it, a template wired to it) isn't a renderer of
those cells, and has its own place: the side panel's **Shelf** tab. It
duplicates neither popover.

* It's in two parts. **For this selection** holds only what fits the
  selection: the card of the template you've picked, if any (it stays
  there as the selection changes), *Repeat … as one formula*, *Templates
  it wires* (those whose inputs the selection labels), *As a function of
  it, below*
  (or *to its right*; see *Functions dropped on a range*, above), *As a
  view of its table*, *As a formula*, *As steps* and *As one chart*, each
  tile drawn from the selection's values, once the selection rests:
  while you drag out a range it waits, then fills in a section at a time
  (*Working out what it can be…* until it has). A click places it
  (above). **The shelf** lists every template on it, one line each (its
  name, what it does, its inputs; point at a line and its face shows,
  at its own values), to pick from whatever is selected;
  when templates are placed other than as a copy, it says how (a card
  changes it). Placing one per row, a template is offered for the
  selection when a heading names one of its inputs, and the tab says so
  when none does.
* **A few tiles, the rest by name.** The first six views and the first
  four templates are drawn as tiles; the others are named under them
  (*1 more: top orders*), and a click draws them too. Views past the
  first six aren't even tried on the selection until that click, so a
  library of many views costs nothing more.
* **What you've used comes first.** What you place from the tab is
  remembered, in this browser, for the kind of selection it was placed
  on (two date columns and one of text; a single number). On a selection
  of the same kind, those come first in each section, the most used
  first.
* The search box finds by caption, the words a view is known by,
  docstring and formula (*pivot*, *pie*), and a template by its name,
  what it does, the words it's known by, and its inputs and outputs
  (*loan*, *repayments*, *rate*).
* A part of the tab that can't be worked out for the selection says so,
  and why, at the top.

### Templates on the shelf, dropped with their inputs wired

A grid that works on its own (§6: inputs, no references out) is worth
reusing: the loan calculator, a little dashboard. Put on the shelf, it's
kept in the browser and offered in every sheet; dropped on a selection,
a copy goes in a new cell with its inputs bound to the cells the
selection labels for them.

**Using it**

* **Put grid … on the shelf…** (cell menu, or the Shelf tab's button)
  on a cell holding a grid, or
  on any cell inside one (a right-click inside a grid drawn in place, or
  zoomed into it, selects a cell of it): asks for a name (the grid's own,
  by default), and keeps a copy, as Save grid writes it, in the browser,
  as My library is kept; the Shelf tab opens on it. References reaching
  outside the grid, which a copy breaks, are listed first, to go on or
  not (*Make input* turns each into an input, §6). A name already on the
  shelf is replaced, asked first.
* **Its display area comes with it.** A template is usually a
  presentation (inputs, results, a chart) and scratch cells (interim
  results). Its author marks the presentation as the grid's display area
  (below, §6): every copy dropped shows only that, the scratch cells
  working out of sight. It's the template's to say, not the dropper's.
* The **Shelf** tab's *Templates it wires* offers the templates whose
  inputs the selection labels (see *Wiring*, below), the most wired
  first; *The shelf* lists them all. A tile is the template's face, at
  its own values: its display area (else all of it, at most 12 × 12, its
  columns sized to fit the tile), each cell by its own renderer and
  number format (a calendar shows as the calendar), and under it which
  input the selection wires to what (*amount ← $B$1*, *years (its own)*,
  *(2 of 3)*). Its own numbers say what sort of thing it is; your
  values come in once you pick it.
* **Picking a template** (a click on its tile, or its line in *The
  shelf*) opens its card at the top of *For this selection*: its tile again, now worked out with your
  values (the drop it will make), how to place it (*Place as a copy*,
  *a linked instance*, *inline*, *one per row*), a field for each input
  to wire by hand (kept for this selection only: select other cells and
  it's worked out afresh), what would refuse it, what it is (inputs, outputs,
  the size it drops as, when it was put there), and managing it:
  *Describe…*, *Rename…*, 🗑 to take it off the shelf (asked first), and
  ✕ to close the card.
  A click on the card's tile places it.
* **What it does.** A template says it itself: its grid's *About* (the
  Inspector, §6) and the words it's known by (*loan, mortgage,
  repayment*) go onto the shelf with it, and with every copy of it.
  Putting a grid that says nothing on the shelf asks, in a few words
  (you can leave it blank), and writes the answer onto the grid; the
  card's *Describe…* changes the template's own. The search finds it by
  them, and its line in *The shelf* shows them.
* **Wiring**, by name: a cell of the selection holding a value is
  labelled by the text to its left, or by the heading at the top of its
  column, and an input whose name is that label (as a header is keyed:
  *Amount* is `amount`) is bound to it; failing that, a named cell of the
  grid with the input's name; failing that, the one label holding all
  the input name's words (*Interest rate* for `rate`; two such labels,
  and neither is guessed). Failing those, when you've selected a range
  and exactly one input left takes data (its own value is a list or a
  table; or it's the template's only input, with no value of its own),
  that input reads the range, under its heading if it has one: a
  column-stats template's `rng` reads `$B$2:$B$9`. Two such inputs, and
  neither is guessed.
* **Data is matched to the template's own.** An input whose own value is
  data is the template's example of what it reads, and a template reads
  its data by position or by heading, so a range or table wired to it
  must match that example. Its own value a list, or rows of values: the
  same number of columns, each holding the same kind (numbers, a date
  counting as one, or text); a column-stats `rng` isn't given a column of
  names, nor two columns. Its own value a table of rows by heading
  (`[{:date … :amount …}]`): a column headed the same for each, of the
  same kind, others no matter; the selection is then given with its
  headings, `(as-table $A$1:$C$9)`. What doesn't match isn't wired, and
  the card says why (*rng takes whole numbers in Name, which holds
  text*); what does, says which column is which (*rng reads Score for
  column 1*). What you wire on the card is taken as you give it. A value of the wrong kind isn't wired: a
  table for an input whose own value is a number, a number for one whose
  own value is a table, or text for a number. An input left unwired
  keeps the template's own value, and is flagged (below).
* **Placing** is as for views and charts (a click, a drag to size,
  arrows, Enter, Esc; the grid grows if it's full), the block the size
  of its display area (else its whole grid) and a header row, at most
  12 × 6. The new cell holds a copy of the template (as Insert does: its
  library, its tests, its display area), its inputs bound by the grid
  cell's formula (`{:amount $B$1 :years $B$2}`, §6), drawn in place,
  `(grid-view)`, and named after the template when the name is free:
  one undoable step. The Inspector opens on it, its *Inputs* showing
  each binding, and the unwired ones' fields to bind; a message says
  what was wired.
* **Inputs left unwired are flagged.** A template's own values are
  its author's: realistic data to build it with, and what makes its
  tile on the shelf worth looking at, but seldom what the place it's
  dropped means. So a placed copy or instance with inputs nothing wired
  gets a formula check (§26): a green corner, counted in the status
  bar's `◤ n to check`, and in the Inspector each such input marked ◤
  in *Inputs*. Give each one (a binding, an override, or a copy's own
  value typed in it) and the mark goes; if the template's value is
  right, *Ignore: keep the template's* under *Check*. The template
  itself, and its face on the shelf, keep their values.
* **A copy or a linked instance.** The Shelf tab's *Place as a copy*
  and *Place as a linked instance* (on a picked template's card) say how a
  template is placed. A copy is a grid of its own, to change on its own.
  An instance (§21) follows its template, by name: `(instance colstats
  {:data $B$2:$B$9})`. If a template of that name is in scope already
  and the same, it's used; else the sheet takes a copy of the shelf's,
  held in its top grid's margin (⚙ Parameters shows it as `▦ template`,
  *follows the shelf's*), which every instance uses, and which follows
  the shelf's changes (§35). The wired inputs are the instance's
  overrides, written in its formula, where any input can be given as
  any formula: a range below it (`d1:de`), to the right of it (`r1:re`),
  anything. Copy the instance and paste it along a row, and each copy
  reads its own column.
* **Inline**. *Place as inline* lays the template's cells into
  your grid, as if you had written its formulas there: no grid around
  them and no link to the template. Each input is replaced by what it's
  wired to, `(BYCOL grid AVERAGE)` with `grid` wired to `$A$1` becomes
  `(BYCOL $A$1 AVERAGE)`, or by the template's own value when nothing
  wires it. References between its cells are aimed at where they land
  (`$B$2` stays fixed, `u1` relative), and one out of it (`..>B6`) reads
  the same cell from your grid. Its names are prefixed with the
  template's (`total` is `loan-total`; `loan-2-total` placed again), its
  format rules come as rules on the block it fills, and its library's
  functions are added to your grid's library. A function your grid
  already has the same (a copy of a standard helper, say) isn't added
  again; one it has differently keeps the template's version under the
  template's prefix, `cal-option-number`, with every use of it in the
  template renamed too, and the card says so before you place it. It
  fills its whole own size, working cells too: when the tile shows less
  (a display area), the card warns how many working cells come with it.
  What stops it is said on the card, by its tile, with what to do: an
  input to wire, or a grid inside the template (inline doesn't lay those
  out). One undo takes it all back.
* **One per row.** *Place as one per row* drops a linked instance for
  each row of a table, or of a range under headings (select it, or a
  cell in it): an input named like a heading reads that column, row by
  row (`Amount` gives `amount`; a heading holding all of an input's
  words, `Interest rate` for `rate`, does too). It's one formula, `(MAP
  $A$2:$A$4 $B$2:$B$4 (fn [amount years] (instance loan {:amount amount
  :years years})))`, spilling a card to a cell, beside the rows. In a
  grid with headers that has a name, the columns are written by their
  headings, `$>loans>amount`, so a row added gets a card too. An input
  no heading names is the same for every card: the template's own
  value, or what you wire on the card. A single column without a
  heading goes to the one input left when you've wired the others. Type
  into a card's input and the value goes to its row's cell. When a
  selection can't be read this way (no heading names an input, or
  nothing is under the headings), the card says so and what to change.
* **Wiring on the card.** Each input has a field on the template's card,
  `grid ← …`, showing what labels and names wired it to, or the
  template's own value as its placeholder. Type a reference (or any
  formula), or click a cell, a grid or a range in the sheet while the
  field has focus; Enter wires it, and the tile is worked out again with
  it. A reference written with names (`$>SmallGrid`, `data>B2`) is kept
  as you wrote it; a cell or range by address is written fixed (`$A$1`),
  or by its name when it has one. It works for all three ways of placing. Inline won't place a
  template whose unwired input has sample data as its own value (a grid,
  or rows of values): the card says to wire it first.

### The library file: My library, the shelf and saved steps on disk

My library, the shelf of templates and the Data tab's saved steps (§34)
live in the browser, so clearing its data, or moving to another browser
or machine, would lose them. The library file holds all three, so one
file is all there is to keep.

**Using it**

* The toolbar's **Library ▾** says what's kept (*My library (12
  definitions), the shelf (3 templates) and the Data tab's saved steps
  (4 sets)*) and offers **Save library
  file…** (the save dialog, `my-library.recdeplib` suggested; then
  remembered, so the next save writes back to it, with *Save library
  file as…* for another) and **Open library file…**.
* Opening one shows what it holds against what's here: *3 definitions: 1
  new, 1 differ from yours (spread), 1 the same; 2 templates: …; 4 saved
  step sets: …*. Then
  **Merge** adds what's new (definitions at the end of My library, with
  the comment lines above them; templates and saved steps by name) and settles what
  differs as chosen once for all of it, *keep yours* (the default) or
  *take the file's* (a definition replaced where it stands); **Replace**
  makes My library, the shelf and the saved steps the file's (asked first). A sheet,
  Clojure source (My library's own *Load file…*, in the Library tab,
  reads that into its editor), or anything else is said and changes
  nothing.
* My library is in the sheet's state, so Undo takes a merge or replace
  of it back; the shelf and the saved steps, outside any sheet, stay as
  they are.

* **Its own extension, `.recdeplib`**, though it's EDN as a sheet is: the
  sheets' Open and Insert dialogs list `.edn` (and Excel, CSV) files, not
  library files, and *Open library file…* lists library files, not
  sheets or Clojure source, so neither is picked for the other by
  mistake. One given to
  Open or Insert anyway (renamed) is said: *That's a library file … open
  it with Library ▾*; a sheet given to *Open library file…* likewise.
  My library's own *Load file…* (Library tab) reads Clojure source only.

### Markdown and links to cells

`markdown` draws text as Markdown, wrapped to the cell: headings,
emphasis, inline code, lists (nested by indenting), quotes, fenced code,
rules, GFM tables, web links and images. `note` is the same, set closer,
for a column of commentary (the tour's notes). Two departures from
CommonMark suit text typed into cells: a line break is a line break (as in
GitHub comments), and there's no raw HTML (the output is hiccup; URLs
become links only if they're a safe kind, as for `link`).

`[[ref]]` and `[[ref|words]]` are links to cells (quotes round the words
are dropped), for documentation inside the sheet. The reference is read
as a formula in the note's cell reads it: `B6>C9` is C9 in the grid in this
grid's B6, `$>B6>C9` starts at the top grid, `..>A1` goes up one, names
and ranges work. A click calls ctx's `:go-to!`, which selects the target,
changing the view to the grid holding it when it isn't on the page; a link
that goes nowhere says why in the message bar. The tour's notes link its
sections this way (`[[$>loan>B2|the loan template]]`), by name and with
words rather than section numbers.

### Colours, brackets and the predefined renderers

The text and fill colours (and the `color-input` renderer, through the
host function `:color-picker` in its ctx) open the app's own colour picker,
not the browser's: a fixed palette of greys and ten hues in
four shades, a row of the colours applied recently (kept in the browser,
shared by all three), a custom colour from the browser's wheel or typed
(`#rrggbb`, a CSS name), None for the format bar, and a now → new preview.
Nothing is written until **Apply** (Enter, or a double-click on a swatch);
**Cancel**, Escape or a click outside leave the sheet and its undo history
as they were.

**A renderer draws in the cell's text colour**. The text colour
(the format bar's A, a column's or range's format rule, or a conditional
rule) colours what a renderer draws, not only its text: a gauge's arc, a
progress or data bar, `bars`, a `sparkline`, a `bar-list`, a `histogram`
or `band`, `stars`, the deepest colour of `heat` and `heatmap`,
and the marks of `chart` and `vega-lite`. Without one they draw in the
theme's blue. So one cell's gauge can be green and its neighbour's red,
and a rule colours by the value:

    {:area "B2:B9" :when "#(> % 90)" :style {:color "red"}}

turns each gauge in B2:B9 red past 90, blue below. What stays: negative
bars are red (`--neg`); colours with a meaning (a trend's arrow, a
traffic light, a check, a test's result) and colours chosen per value
(a `badge`'s, a `pie`'s slices, `(heat lo hi from to)`) keep theirs; a
progress bar's percentage is in the theme's text colour, since it sits
on the bar. A fill colour alone doesn't change a renderer's colour.
The colour is drawn as chosen, even where the cell's text gives way to
the theme's (an orange or yellow text colour in the light theme, §4,
*Light and dark*): a bar or an arc shows where text wouldn't. Text
colour is the one setting for both, so a cell's number and its gauge
share a colour; a `link` keeps the link colour.

Your own renderers do the same with `(ink)`, the colour as CSS:
`[:rect {:style {:fill (ink)}}]`, or `(ink "teal")` to draw in teal
when the cell has no text colour. Text needs nothing: it takes the
cell's colour anyway.

**Brackets are optional** where there are no options: `grid-view` is
`(grid-view)`, `(key-values)` is `key-values`, `(trend compact)` is
`(trend (compact))`. Which one a name is comes from its var: a
`^:renderer` that can be called with no arguments is a maker, so written
bare it's called; one that can't is a renderer, so `(name)` is the
renderer itself. Only the compiled form changes, never the source, and
only the renderer field's top-level form and the options of calls to
renderer makers — not a `fn`'s body. A maker that needs options, named
bare (`heat`), draws `#RENDER` saying how to call it. So does `(heat)`:
with no zero-argument arity it's taken as the renderer `heat`, since
two-argument makers can't be told from `(fn [v ctx])` by their arity.

Predefined (each documented in the Library tab's Reference; renderers
written bare, makers called with options, though either spelling works):

| kind | renderers |
|------|-----------|
| numbers and text | `plain` (in the cell's number format and unit), `(compact)`, `file-size`, `duration`, `(truncate n)`, `code`, `html` (value *is* hiccup), `markdown` (`(markdown "template")` for a record or a table's rows, §56), `note` (markdown set close, for notes), `link`, `(image)`, `check`, `swatch`, `(badge)` |
| combining | `(render-if pred then else)`, `(tooltip text)` |
| indicators | `(progress)`, `(data-bar hi)`, `(heat lo hi)`, `(trend)`, `(traffic-light lo hi)`, `(stars)`, `(kpi label)`, `(gauge lo hi)` |
| collections | `table` (a 2-D value as a read-only grid; a table of row maps with a header row, §19), `(table-of inner)`, `size` (a range as "3×2 range", kept in its cell), `heatmap`, `bullets`, `numbered`, `key-values` (a map as keys and values; a grid's values as its named cells, labelled by the text to their left and in their own formats, or laid out as in the grid when it has no names) |
| charts | `bars`, `sparkline`, `bar-list`, `pie`, `donut` (SVG; `(pie colors)` for your own colours) |
| controls | `(button f)`, `(counter-button ref step)`, `(slider lo hi)`, `(stepper)`, `toggle`, `text-input`, `number-input`, `date-input`, `color-input`, `key-input` (a web key's field, §56), `(dropdown opts)`, `(radio opts)`, `(checklist opts)` (ticks: the cell's value a vector of those ticked) (bare `dropdown`, `radio`, `checklist`, `(slider)`: from the cell's validation, §16) |
| grids | `grid-link`, `(grid-view)`, `(grid-view rows cols)`, `instance-view`, `(instance-view :clean)` (§21) |

`default-render` picks `grid-link` for grid cells, `instance-view` for a
linked instance (§21), `table` for vectors of vectors and tables of row
maps, `λ` for functions, and `plain` otherwise. (A cell whose own result
is 2-D spills instead, see §13, unless it has a renderer of its own: give it
`table` to keep the result in the cell.) Renderers take any value without
failing: what they don't draw specially, they show `plain`.

### Spans: a cell drawn over its neighbours

A chart, or a grid shown in a cell, often needs more room than one cell.
A **span** gives it a block of cells: the cell is drawn over the
rectangle of `rows × cols` cells from it, sized by those cells' widths
and heights.

* **Set it** by selecting the block and right-click → *Span*, which gives
  the top-left cell `:span [rows cols]`; *Remove span* (or the
  Inspector's Remove) takes it away. A formula's grid can give its cells
  spans (`as-grid`'s and `panel`'s `:spans {[r c] [rows cols]}`, §22),
  and Excel's merged cells come in as spans (§17).
* **Resize it** by its handle, the small square at the span's
  bottom-right corner (shown when its cell is active or the span is
  hovered): drag it to the cell the span should reach. While dragging,
  the span turns see-through and the block it would cover is outlined —
  red when a cell in it isn't empty or another span's block overlaps it,
  and such a size isn't taken (the span stays as it was, with a
  message). Dragging back onto the cell itself removes the span; Esc
  cancels. One drag is one undo step.
* **Move it** by the square at its top-left corner: drag it to the cell
  the span should start at. The cell moves there with everything it has
  (formula, formats, name, span, the grid in it), as Excel's
  drag-and-drop does, and references keep naming the same cells:
  formulas that read it, or reach into its grid (`B5>A1`), follow it;
  its own relative references (`u1`, `..>l1` inside its grid) are aimed
  back at their cells; a range that includes it keeps its cells (as in
  Excel). The block at the new place must be empty (but for the span's
  own old block) and clear of other spans; the outline shows it, red
  when it isn't. Only a span set on the cell has handles: one a
  formula's grid gives is changed in the formula.
* **A span is a way of drawing**, like a renderer: the cell's value, its
  formula and the cells under it are unchanged. The covered cells are
  still cells — addresses, references, the name box and the arrow keys
  reach them. With one of them active, the span fades so it shows, and
  clicks go through it. Clicking the span selects its cell.
* **The covered cells must be empty** (no formula, grid, value, spilled
  value or span of their own). If one isn't, the span is *blocked*: the
  cell is drawn in its own cell only, hatched, and its tooltip and the
  Inspector name the blocking cell — as a spill is blocked (§13). Nothing
  is hidden.
* **Sizes** come from the cells: widen the columns or set the rows'
  heights to make the span bigger in pixels; the handle changes how many
  cells it covers. A span is clipped to its grid. Rows
  and columns inserted inside it don't change its counts.
* **What fills it**: `svg` charts (`bars`, `sparkline`, `pie`, `donut`)
  fill the block; a `(grid-view)` stretches its rows and columns to it,
  as it does to a cell with a set height (§6). A grid drawn in a span is
  live like any inline grid.
* Copying a cell copies its span (it goes with the formats: *Paste
  formats* carries it, *Clear formats* removes it).

### Charts with Vega-Lite

Beside the small SVG renderers (`bars`, `sparkline`, `pie` …), charts with
axes, legends and tooltips are drawn by [Vega-Lite](https://vega.github.io/vega-lite/).

```clojure
(chart :bar :region :qty)                         ; qty summed by region
(chart :line :day :qty {:x {:type :temporal} :color :region})
(chart :arc)                                      ; a pie of a map {label n}
(vega-lite {:mark {:type :line :point true}       ; any Vega-Lite spec
            :encoding {:x {:field :month :type :ordinal :sort nil}
                       :y {:field :sales :type :quantitative}}})
```

* **Specs are Clojure maps** in Vega-Lite's JSON shape; keywords become
  strings. `chart` covers Excel's kinds (§42): a kind, the field across
  and the field up, typed from the data (numbers are quantities, the
  rest categories); bars and arcs sum (`:aggregate nil` for none); its
  own options, each a refiner row (§42); anything else in its options
  map is merged into the spec.
* **The cell's value is the data** (unless the spec has `:data`): a table's
  row maps as they are; a map `{label n}` — a grid's values, a panel's
  totals — as rows `{:key label :value n}` (what `chart` reads without
  fields); a range with a header row of text keyed by its headers, as a
  headed grid's rows are; a list of numbers as `{:index i :value x}`.
  `vl-rows` gives these rows, `vega-spec` the spec with them.
* **Dates are numbers** (§15). Wherever Vega-Lite reads dates, date
  numbers are turned into dates: a field or a `:datum` typed `:temporal`
  or given a time unit; a bar's end (`:x2`, `:y2`), which takes its type
  from `:x` or `:y`; a field or datum without a type on an axis of dates
  (another layer's); a tooltip's entries typed `:temporal`, and those
  without a type naming a field read as dates elsewhere. Bounds and tick
  values written on a date axis (`:scale {:domain [d1 d2]}`, `:axis
  {:values …}`) are dates too. All are drawn in UTC, axes and tooltips,
  so no time zone shifts a day.

  ```clojure
  (vega-lite {:layer [{:mark :bar :encoding {:y {:field :task}
                                             :x {:field :start :type :temporal}
                                             :x2 {:field :finish}
                                             :tooltip [{:field :task} {:field :start} {:field :finish}]}}
                      {:mark {:type :rule :color "red"}
                       :encoding {:x {:datum 46430 :type :temporal}}}]})
  ```
* **Size**: a chart is as wide as its cell and as tall as its row's set
  height (180px otherwise); a span (above) or a panel's `:spans` and
  `:heights` (§22) give it room. It's redrawn when the cell is resized.
* **Colours** come from the app's (see Light and dark, below): text, axes
  and grid lines in the theme's colours, categories in `chart-colors`.
* **Only the sheet's data.** A spec's `data.url` (or any other URL) isn't
  fetched: a sheet's charts draw from its cells, the same everywhere and
  offline. The cell says so instead.
* **Limits.** Dates are read in every view: layers, concatenations, a
  facet and its spec, a repeat's spec (but not the fields a repeat
  names, `{:repeat :column}`). Not in a spec's own `:data`, nor
  in what a `:transform` computes or filters on (write those as
  milliseconds, or as text "2026-11-07"). Layers whose scale is
  resolved `:independent` don't lend each other dates. A field typed `:quantitative`
  that is read as dates elsewhere shows milliseconds. Charts don't
  write back.

### Light and dark

The toolbar's ◐ picks **Auto** (as the system is set), **Light** or
**Dark**; the choice is kept in the browser. Every colour of the app is a
CSS variable with a light and a dark value, which follows the choice, as
do the browser's own parts (scrollbars, drop-downs, checkboxes, date pickers).
Renderers use the variables, so they follow too; charts redraw.

## 5. Working with ranges

* **Selection** is the active cell plus the selection's far corner, within
  one grid. As in Excel, the active cell is where the selection started:
  drag, Shift-click and Shift-arrows move only the far corner. Row/column
  headers and Ctrl-A select whole rows/columns/grids.
* **Fill.** Ctrl-Enter in the formula bar puts the formula, as written for
  the active cell, into every selected cell, shifted for each cell as a copy
  would be (Excel's Ctrl-Enter). Ctrl-D / Ctrl-R copy the first row / column
  across the selection, also shifting. A single cell (or a single row for
  Ctrl-D, column for Ctrl-R) holding something goes onward instead: into
  the cells below / to the right, up to the grid's edge or the first cell
  that isn't empty (a spilled value counts), whichever is sooner, so
  nothing is overwritten. Type a value, Fill down from the cell menu, and
  the column fills. An empty one takes the row above / column left of it,
  as in Excel (a spilled value there isn't a cell to copy).
* **Number format, renderer and style** edits apply to the whole selection.
  Delete (and Backspace) clears the cells' contents, as Excel's does:
  values, formulas and grids go; names, renderers (a drop-down, a
  toggle), formats and comments stay. On cells holding nothing it says
  so, and what stayed. The cell menu has *Clear contents* (Delete's) and,
  under *Clear ▸*, *Clear all* (contents, formatting, comments and names
  at once, in one undo step: formulas elsewhere that used the names show
  `#NAME`; bound inputs keep theirs, §2 Names), *Clear names* (only the
  names), *Clear style* (keeping the number format and renderer) and
  *Clear formats* (all three).
* **Point mode.** While writing a formula, after `(`, `[`, `{`, a space or a
  comma (and in a formula starting with `=`, after that `=` and after
  `+`, `-`, `*`, `/`, `^` or `&`), clicking or dragging over cells inserts a ref instead
  of selecting; so do the arrow keys, as in Excel: `=` then ↓ ↓ gives
  `=A3` (each arrow moves the pointed cell, replacing the reference it
  put in), Shift-arrows grow a range (`=A3:A4`), and typing `+` then
  an arrow points again from the cell you're writing. A click right
  after a reference you pointed at replaces it too, so you can keep
  clicking until it's the one you want. A click on a cell picks that
  cell whatever it shows: a date picker, a slider, a toggle or a button
  in it doesn't act, and a grid's link (`▦ 4×2 grid ›`) picks the grid.
  Double-click a grid's cell to go into it and pick one of its cells
  instead. In a grid drawn in its cell, a click on one of its cells
  picks that cell (`box>A1`), and a click on its corner ▦ picks the
  grid (`box`).
  The formula bar's `refs:` toggle cycles between `B2:B5` (or the cell's
  name), `$B$2:$B$5` (kept by copies) and relative `u4:u1`; plain A1 is the
  default because it reads as what was clicked. The Inspector's *Inputs*
  fields work the same way (a grid cell's bindings, an instance's
  overrides, §6 and §21): with the cursor in one where a reference could
  start (an empty field too), click a cell or drag out a range, and its
  reference goes in, written from the grid cell or instance; the
  Inspector stays where it is, and Enter gives it.
  Pressing a column's letter puts in the whole column: in a table, by
  its heading (`(SUM ..>resources>days)`), which reads down to the
  table's last row however many rows it has; elsewhere by its letter,
  as Excel writes it (`B:B`). A table drawn in a cell
  (`(grid-view)`, a facade's link) has its letters too.
  To pick a cell in another grid, just go there while you write: the
  Outline, the breadcrumb, a sheet tab, a double-click on a grid's link, or Alt-↓ / Alt-↑
  show that grid and leave the formula open; click the cell (or point
  with the arrows) and its reference goes in by names (`model>rate`).
  Enter enters the formula and takes you back (§11, *Picking while you
  write a formula*).
* **AutoSum (Σ)** writes `(SUM …)` under each selected column holding
  numbers, its heading left out, in the row below its last filled one
  (or right of a selected row, its label left out), the grid growing
  if need be; on a single cell it writes `(SUM ub:u1)`. It goes where
  the Shelf tab's *sum* goes: a cell in the way, or a grid with headers (where
  a row below the data would join its table), is said, not overwritten;
  the Shelf tab then places the sum elsewhere.
* **Copy/paste**: see below.

### Column widths, row heights, overflow

* **Columns have fixed widths**, as in Excel: content never changes them.
  Each grid stores its own column widths and row heights (saved, undoable,
  kept on resize, copied with the grid).
  Unset columns use a default: 120px in the main view, 72px for a grid drawn
  inside a cell. Drag the right edge of a column header or the bottom edge of
  a row header to resize (one drag is one undo step); double-click the edge
  to go back to the default. Rows fit their content unless given a height.
* **Back to the defaults.** Right-click whole columns (their headers) for
  *Reset width of column B* (or of B–D), whole rows for *Reset height*, and
  any cell for *Reset all column widths and row heights in* its grid: the
  grid cell selected, else the grid the selection is in. A grid shrunk
  keeps the sizes of the columns and rows it lost, for when it grows again;
  the reset clears those too (the item says so when there are any). Each
  is offered only when there's something set to clear, and is one undo step.
* **Text overflow**: a cell showing plain text (a string with the default
  renderer, not centred or right-aligned) runs on over the empty cells to
  its right and stops at the first non-empty one; numbers and other content
  are clipped with an ellipsis. A cell counts as empty if it has no formula,
  no grid and no value; formatting doesn't block overflow. Clicks on the
  covered cells still reach them.
* **A grid drawn whole in a cell** widens its column to fit and makes its
  row taller (§40); a grid in a window, `(grid-view rows cols)`, fills its
  cell and scrolls (§6).

### Frozen rows and columns

A grid can keep its first rows and columns in view while the rest
scrolls, as Excel's Freeze Panes: the names down the left of a table of
weeks, its heading row along the top.

* **Right-click a cell → *Freeze and protect ▸ Freeze rows above 2 and
  columns left of B***
  (the item names the active cell): the rows above it and the columns
  to its left stay put. Select a cell in column A for rows only, in row
  1 for columns only. *Unfreeze rows and columns* undoes it. Each is one
  undo step, and the grid keeps it: saved, copied with it.
* **A table's heading row is frozen** to start with (a grid with *Row 1
  is headers*, §19): scroll down a long table and its headings stay.
  Unfreeze to let it scroll.
* Up to 9 columns can be frozen. It works in the grid you're in (the
  main view), big grids too. A grid drawn in a cell scrolls with the
  sheet, so it's frozen only when you zoom into it.

### Copying formulas

Copying a formula — paste, Ctrl-Enter fill, Ctrl-D / Ctrl-R — rewrites its
references as Excel does:

* **Coordinates move with the formula unless fixed with `$`.** `(* B2 C2)`
  pasted one row down becomes `(* B3 C3)`; `$D$6` never moves; `$D6` keeps
  the column, `D$6` the row; `$Au0` keeps column A.
* **Relative refs (`u1`, `r1:re`, `ub:u1`) and names never change.** They
  already mean the same thing wherever the formula sits.
* **Only the formula's own grid is shifted.** In a path, only the first
  step moves, and only if it's a coordinate: `B5>A1` → `B6>A1` one row down.
  `$>A1`, `..>B2` and the parts inside other grids are left alone, since
  the formula's displacement in its own grid doesn't mean anything there.
* **A grid copied: what it reads around it moves too.** Copying a grid
  cell is copying its formulas along with it, so a reference inside it to
  the grid the copy is made in moves as the cell's own would: in a grid
  copied from B7 to B9, `..>B8` becomes `..>B10` (and `..>..>B8` from a grid
  inside it). `$` fixes it as ever; references within the copied grid,
  moves (`..>u1`) and paths further out stay as written.
* **Off the top or left edge** a reference becomes `REF!` (Excel writes
  `#REF!`, which isn't a Clojure symbol). It shows as a `#REF` error.

Paste modes, as in Excel's Paste Special (right-click a cell):

| mode | writes | keeps from the destination |
|------|--------|----------------------------|
| Paste (Ctrl-V) | formula (shifted), number format, renderer, style | — |
| Paste values (Ctrl-Shift-V) | the copied values, as literals | number format, renderer, style |
| Paste formulas | formula (shifted) | number format, renderer, style |
| Paste formats | number format, renderer, style | formula |
| Paste link | a link to each copied cell (`$>model>rate`), in its number format (§54) | renderer, style |

* The clipboard is a snapshot taken at copy time, marked by a dashed
  outline until Esc. The system clipboard gets tab-separated values, each
  as the cell shows it (a date as `2026-11-07`, money as `$1,234.50`), and
  text from other apps pastes as numbers, dates (with a date format) or
  strings.
* Values that can't be written back as Clojure (functions, errors) are
  skipped with a message.
* One copied cell pasted onto a selection fills it; a block goes at the
  selection's top-left.
* The menu's paste items use our own clipboard; if there isn't one they ask
  the browser for the system clipboard, which may prompt for permission.

### Copying to other programs: nested grids as one table

A copy also puts a table on the system clipboard, which Excel, Word,
PowerPoint and Google Sheets paste in preference to the text: every grid
drawn whole in the selected cells (§6), at any depth, laid out as cells
of the one table. A cell holding a 3-column grid becomes three columns,
and a label beside it, or a plain cell above it, is one cell merged
across them. So a sheet that presents its results in grids within grids
pastes into a slide or a workbook as one table.

* **What's copied is what's drawn.** A grid showing only a block of
  itself (its display area) gives that block; an instance (§21) and a
  2-D result kept in its cell (`table`) give their cells; a
  grid drawn as a link gives its link's text, a chart nothing. Each
  cell's text is as it shows (a date as `2026-11-07`), its own bold,
  italics and colours kept (not the theme's), numbers aligned right.
* **On a mesh (§41), the lines you see.** Every column and row edge at
  every depth falls on the mesh, so the table has a column wherever any
  grid in the selection has a column edge, and a row wherever any has a
  row edge: a grid with one wide column above a grid with two narrow ones
  pastes as one cell merged across two. A grid's headings and the room
  round it make no columns or rows of their own.
* **Off the mesh, by nesting.** Edges don't meet there, so each grid's
  column is as many columns as the most any grid inside it has, and
  likewise rows: a 3 × 2 grid above a 4 × 4 one pastes as 7 rows × 4
  columns, the narrower grid's rows padded.
* RecDep's own paste still takes formulas, formats and grids, as above:
  the table is only for other programs.
* **Limits.** Selections over 20,000 cells copy as text alone. Cells of a
  big grid's rows not drawn (scrolled far away) are copied as their text,
  without the grids inside them. A 2-D result's columns are as wide as
  their text, so on a mesh they don't line up with the grids around and
  make narrow columns of their own. Formulas can't fuse nested grids;
  `VSTACK` and `HSTACK` stack flat ones.

### Cut and paste: moving cells

**Cut (Ctrl-X) then Paste (Ctrl-V) moves the cells**, as in Excel. Cut
only marks them (the dashed outline); nothing changes until you paste.
Then, as one undo step:

* The cells go to the selection's top-left with everything they have:
  formula, value, formats, renderer, style, name, a grid with its
  contents.
* **References keep naming the same cells.** Formulas elsewhere that
  referred to the moved cells now refer to where they went (`(* A3 10)`
  becomes `(* E7 10)`, `$A$3` becomes `$E$7`, a relative `l2` is aimed
  anew); the moved formulas still read the cells they read before; a
  range follows when all its cells moved (`(SUM A1:A2)` → `(SUM E5:E6)`)
  and keeps its cells when only some did. Renderers, rules, scenario
  inputs and test cases follow too.
* **Pasting over cells that hold something asks first**: they're
  replaced, and formulas referring to them show `#REF` (`REF!` in the
  formula), as in Excel.
* The old and new places may overlap (moving a block down one row).
* Esc cancels the cut, and so does **any change to the sheet** before
  you paste (an edit, an undo, another sheet opened), as in Excel: the
  outline goes, and a paste then says so. Cut again.
* A formula's spill (§13) moves with the formula: cut its whole range
  or none of it.

What isn't a move:

* **Into another grid**, the cut cells are copied there and emptied where
  they were, in one undo step, and the status bar says their references
  weren't changed: references follow cells moved within a grid only.
* **In another tab** (another sheet), a cut pastes as a copy, and says
  so: the cut cells stay in the sheet they came from.
* **Paste values, Paste formulas and Paste formats** don't take a cut:
  copy the cells (Ctrl-C) for those.
* A grid whose cells come from its formula (§22) can't have its cells
  moved.

One cell can also be moved with a span's move handle (§4).

### Inserting and deleting rows and columns

(The cell menu folds related items into submenus, opened by hovering
them: *Paste special ▸* (values, formulas, formats), *Clear ▸* (contents,
names, style, formats; *Clear all* stays at the top), *Insert ▸* and
*Delete ▸* (rows and columns; with whole rows or columns selected
they're listed as they are), *Freeze and protect ▸*, and on a table's
heading *Column holds ▸*.)

Right-click a row or column header (or a cell) for **Insert N rows
above / below**, **Insert N columns left / right**, **Delete N rows**
and **Delete N columns**, N being the selection's height or width. A
header menu offers only its own axis when whole rows or columns are
selected. The selection stays put, so after inserting above or left it
covers the new, empty lines, and after deleting it covers the lines that
moved up or left into its place (as in Excel). Each is one undo step.

* **Cells move with everything they have**: formula, renderer, style,
  name, and any grid they hold with all its contents. Column widths, row
  heights and the areas of format rules (§10) move with their columns and
  rows. The grid grows by N (Excel's sheet has a fixed size and drops the
  last rows instead), and cells hidden beyond a shrunk grid's edge move
  too. Spills (§13) are laid out again from their moved anchors, so an
  insertion inside a spill's area gets filled by it.
* **Absolute references keep naming the same cells**, anywhere in the
  sheet: a coordinate that lands in the grid at or after the insertion
  moves by N — `B5`, and also `$B$5`, since `$` fixes a reference against
  copying, not against its cells moving. Every step of a path is checked,
  not just the first: `E15>B2` (into a grid), `..>B5` (from a grid inside),
  `$>B5`, `total>A1`, and a range's single-step second end, which lives in
  the first end's grid (`E15>A1:C3`). A range's ends move independently,
  so inserting inside `B2:B9` grows it to `B2:B10`; inserting at its first
  row moves it; inserting after its last row leaves it alone.
* **Relative references follow their cells too**, as in Excel: a counted
  move (`u1`, `d4`, `r2`) is measured again from where the formula is now
  to where its cell is now. `(SUM d4:de)` in A1 becomes `(SUM d6:de)`
  when 2 rows are inserted at row 3, and a `u1` with a row inserted
  between it and its cell becomes `u2`; inserted before both, or after
  both, nothing changes. Only moves along the edited axis are measured
  (`d2l1` keeps its `l1` on a row insertion); several along it become one
  when they change (`d1d1` → `d3`). Edge moves (`ub`, `de`, `lb`, `re`)
  mean the grid's edge wherever it is and don't change, so `(SUM ub:u1)`
  still sums everything above. In a mixed reference each part moves:
  `B5u1` becomes `B6u2` when a row is inserted at row 5, `$Au0` becomes
  `$Bu0` on a column insertion before A. A cell's renderer follows like
  its formula. Format and validation rules (§10, §16) keep their relative
  moves as written: `(data-bar r1)` over column D means each cell's right
  neighbour, wherever the columns are.
* **Names don't change**: they mean the same cell wherever it is.

* **Not rewritten**: references that aren't symbols in a formula — the
  quoted refs and strings renderers pass to `set-cell!` (`'l1`, `"B5"`),
  keys into a grid's value (`(:B2 E15)`) and `call`'s cell keys
  (`{:A1 5}`). Use relative refs or names there, which insertion doesn't
  disturb. (Excel's `INDIRECT` has the same problem.) A formula that
  already holds a `REF!` (from a copy or a deletion) still has its other
  references rewritten, so `(+ REF! B6)` reads `(+ REF! B5)` after a row
  above B6 is deleted.

**Deleting** is insertion run backwards, plus a rule for what was in the
deleted lines:

* **The deleted cells go with everything they have**, grids inside them
  included; so do their row heights or column widths. Cells after them
  (hidden ones beyond a shrunk grid's edge too) move back by N, and the
  grid shrinks by N. A grid keeps at least one row and one column, so the
  menu doesn't offer deleting all of them.
* **Absolute references keep naming the same cells**, as for insertion: a
  coordinate after the deleted lines moves back by N, `$` or not, at any
  step of a path.
* **A reference to a deleted cell becomes `REF!`**, as for a copy that
  moves one off the edge, and the formula shows `#REF`. So does a path
  that goes through one (`E15>B2` once E15's row is gone: the grid it led
  into went with it), and a range whose cells were all deleted.
* **A range loses the deleted lines** (as in Excel): deleting inside
  `B2:B9` shrinks it to `B2:B7`; deleting its first rows starts it at the
  first row kept, deleting its last rows ends it at the last row kept.
  `(SUM B2:B9)` thus still sums what's left of its column.
* **Relative references follow their cells**, as for insertion:
  `(SUM d4:de)` in A1 becomes `(SUM d2:de)` when rows 2–3 are deleted. One
  whose cell was deleted becomes `REF!` like an absolute one (a `u1`
  whose row went, as in Excel), and a range by moves loses the deleted
  lines like any range (`d2:d5` in A1 is `d1:d3` once rows 2–3 are
  deleted).
  Edge moves (`ub`, `de`) don't change, nor do rules' relative moves.
* **Names are left as written**: a
  name stays as written when its cell is deleted, and the formula reports
  that no cell has that name, so naming another cell repairs it; Excel
  turns such names into `#REF!` instead.
* **Format rules** lose the deleted lines like ranges do, and a rule whose
  whole area was deleted goes (§10). Spills are laid out again: deleting
  the anchor's row removes the spill, deleting a row under it just moves
  the rest up and fills it again.
* Undo brings everything back, `REF!`s included, since it restores the
  previous sheet.

### Undo

* One user action is one step: a 300-cell fill, a paste, AutoSum, a
  library change, a resize.
* **Opening a workbook starts the history afresh**: Open, New, the
  guided tour, an example or a link that opens a sheet. Undo goes back as
  far as the moment it was opened, then says *Nothing to undo since this
  workbook was opened*; it doesn't go back to the workbook you had
  before, so save that one first if you want to keep it (the confirmation
  before replacing it says so).
* Changes an interactive renderer makes to the same place less than a
  second apart merge, so a slider drag or typing into a text-input cell is
  one step. (Rapid clicks on the same
  button merge too.)
* A button that writes several cells (`(set-cell! ctx 'A3 …)` and
  `(set-cell! ctx 'B1 …)` in one handler) makes one step: one Undo takes
  all of it back, as for a form typed into the REPL.
* Ctrl-Z / Ctrl-Y (or Ctrl-Shift-Z) work anywhere except inside the
  formula/library editors and input cells, which keep their own text undo.
  After you click a button in a cell, its cell is selected and the grid
  has the keyboard again: Ctrl-Z, the arrows and typing work at once.
* At most 200 steps; history lives for the session and isn't saved.
* Big steps are let go to save memory: a step that rewrites a large block
  (a Refresh of a 200,000-row grid, a big paste, opening a file) keeps a
  copy of everything it replaced, hundreds of megabytes. Undo always keeps
  your last step, however big, and lets an older one go once about
  1,000,000 cells' worth of changes have been made after it (typing into
  a few cells after a Refresh doesn't; a second Refresh does). When undo
  runs out because of that, the status bar says so.

### Errors where you're looking

A refused change or any other error is a message in the status bar,
which is far from where you're working: easy to miss, so what you did
looks as if it did nothing. So every error and refusal also comes up in
a box where you acted, with a brief shake:

* next to the control you pressed (a button, a field, a link), below it
  or, near the bottom of the window, above;
* else by the selected cell, for a key or a menu on the grid: Ctrl-Z
  with nothing to undo, typing into a protected cell, Insert on a
  protected grid, a paste refused;
* else, for an answer that comes later (a file read, a page fetched),
  above the status bar's right end.

The box never takes the keyboard: carry on typing. It goes at your next
key or click, or after a few seconds. The message stays in the status
bar (marked ⚠) as a reminder of what was refused, until you click it
away, another takes its place, or you go on to change the sheet some
other way (an edit, Undo, another sheet opened): then it's about
something past, and goes. Notes saying what was done (*Added column M
to tasks*, *Replaced in 3 cells*) are beside it in the status bar,
plain, and go the same way: at the next note, or the next change after
it. Moving about the sheet, reading, opening menus keep both.

### Errors an edit made

A change can break a cell far from the one you typed into: a task ID
typed twice in a list stops the plan worked out from it, on another
sheet. So when an edit makes new errors, the page says so once the
sheet has recalculated:

* The status bar's **⚠ errors** count flashes.
* A box above the cell you edited names where the new error starts
  (`plan (B21)`), with a link that goes there, and its message. When
  one error breaks others further on, it names the one nearest your
  edit, the cause, then where it ends up: the first cells showing it,
  each a link, and how many more; and how many other places an error
  starts. It doesn't take the keyboard: carry on
  typing. It goes after a few seconds (not while the pointer is on it),
  with ✕, at your next edit, or once the error's gone (undo). An error in the cell you typed into
  itself only flashes the count: you can see it already.
* 🔔 in the box turns on a short sound with it (off until you do; the
  setting is kept in this browser). It plays once as you turn it on, so
  you know what to listen for; if you don't hear it then, check the
  browser tab isn't muted and your speakers are on. Like the box, it's
  for errors outside the cell you typed into. If the browser holds a
  sound back, the box says so; turning 🔔 off and on again lets it play.

Only errors the edit has just made count, not those already there:
inserting a row above an error moves it without alerting, and undo,
opening a sheet and cells that update on a timer never alert.

## 6. Grids within grids

A cell holds either a formula or a grid. The grid's cells are ordinary cells
at path addresses (`B5>A1`), in the same dependency graph, with their own
formulas, renderers, styles and names.

* **Value.** A grid cell's value is the map of its non-empty inner values,
  keyed by cell keyword, in row order: `{:A1 1 :B1 10 :A2 8}`. Each is what
  the cell shows: a spill's formula cell its result's top-left, as a range
  over it reads it (the rest is in the cells it spilled into), so nothing
  is counted twice. Nested grids
  give nested maps. Use it as data: `(:A2 B5)`, `(SUM B5)` (Excel's
  functions and the library's `cells` flatten maps). Where an Excel
  function wants a range by shape — `(VLOOKUP "b" B5 2)`, `(INDEX B5 3 2)`,
  `(ROWS B5)`, `(* B5 10)` — a map keyed by cell names is read as the
  range of its cells by position, from A1 to its last row and column
  holding a value (trailing empty rows and columns aren't counted); a
  headed grid, as its data rows without the header row (§19). A grid that
  shows only a block of itself (its display area, below) has that block
  as its value: its working cells, out of sight, are out of it too, and
  read as a range it starts at the block's top-left, not A1, so
  `(VSTACK B4 B5)` stacks what B4 and B5 show. Reach a working cell by
  its path (`B5>A1`) or `(call B5 :A1)`. Named inner cells aren't duplicated in the map;
  reach them with `B5>total`. An error in any inner cell makes the grid
  cell's value an error, as with ranges. A grid whose first row is
  headers has a table as its value instead, a vector of row maps (§19).

* **References.** The innermost grid is assumed: `A1` in B5's grid is B5>A1.
  `..>` goes up (`..>l1`, `..>..>A1`), `B5>A1` goes down, `$>` starts from
  the root. `..` alone is the cell holding the grid around (a tests grid's
  tests call the grid they test with it, §27). In the top grid's own tests
  grid it's the top grid; anywhere else that would mean the top
  grid, it's an error. A name can step
  into a grid kept in a margin (`loan>_tests>A1`). Names are per grid. Clicking a cell in another grid while writing
  a formula writes the shortest path (`..>B2`, `B5>A1`, `..>C2>B1`).
* **Showing a grid is a rendering.** A grid cell's renderer
  gets `:grid {:rows :cols}`, `:zoom!`, and `:show-grid`, which returns the
  live grid as hiccup: `(show)` for all of it (its display area, when it
  has one), `(show rows cols)` for a window of that size, each taking
  a last `{:look :clean}` or `{:look :grid}` to say how it's drawn. Two
  library renderers use them:
  * `grid-link` — the default: a button that zooms into the grid, the same
    as the In button;
  * `(grid-view)` — the whole grid drawn in the cell;
  * `(grid-view rows cols)` — that many rows and columns in view, the rest
    reachable by scrolling (headers stay put);
  * `(grid-view :clean)`, `(grid-view :grid)`, `(grid-view rows cols :clean)`
    — the same, saying how it's drawn (below, *Drawn clean*).

  **A grid's display area.** A grid often holds a presentation (inputs,
  results, a chart) and scratch cells (interim results) its author
  would rather not show where it's drawn. Which block is which is the
  grid's own business, so the grid says it, not the renderer of the
  cell it sits in: inside the grid (zoomed in, or in it drawn in place)
  select the block and right-click → *Show only this in …'s cell*; *Show
  all of … in its cell* (or the Inspector's *Show all*) undoes it. Then
  wherever the grid is drawn in a cell, `(grid-view)` draws only that
  block, with no scrollbars to the rest, the grid's value is that block
  alone (above), and `(grid-view rows cols)` a
  window within it; an instance (§21) of it shows the block too. Going
  to one of its cells outside the block (a status-bar count, Find, the
  Outline, a link) zooms into the grid, where every cell shows; the arrow
  keys stop at the block's edge, as they do at a grid's. A block selected
  up to the grid's right-hand or bottom edge **sticks to that edge**
  (`A3:re`, row 3 to the right-hand edge; `A1:Cde`, columns A–C to the
  bottom), so a result that grows the grid (§38) shows all of it; the
  Inspector says *and to the right as the grid grows*.

  **Drawn clean.** How a grid is drawn in its cell is its renderer's
  choice, like any other drawing: `(grid-view :clean)` draws it clean,
  with no row and column headings, no frame round it (only the lines
  between its cells: the outer grid's own lines frame the cell holding
  it), no frames round spilled results, and no marks for names, formula
  checks or bound inputs (a failed validation and a units mistake still
  show: they're about the values). Values and interactive renderers
  stay. `(grid-view :grid)` draws it as a grid, with all of those. Plain
  `(grid-view)` goes by what shows: marking a display area says the
  block is a finished presentation, so a part is drawn clean and all of
  a grid as a grid. The look belongs to the cell, so a copy of the cell
  draws the same way. Choose it as any renderer is chosen: in the
  Renderer field, `grid-view`'s refine panel has a *Drawn* row (*as it
  shows*, *clean*, *as a grid*), keeping a window's rows and columns.
  The part drawn as a grid keeps its headings, spill frames and marks
  only for its own cells: a block from A4 is headed 4, the rest of the
  grid still out of sight.
  The Inspector on the grid's cell says what shows: its *Shows* line is
  the part (*A4:B4*) or *All of it*, the current one marked. Showing all
  of it keeps the part, so its button stays and brings it back as it
  was (sticking to the edges or not); a new part chosen inside the grid
  replaces it. With no part chosen the line says how to choose one.
  Switching keeps how it's drawn. Without headings there's no ▦ corner: hovering the grid
  shows a ▦ handle in its top-left corner (over the first cell) that
  selects the grid's cell, double-clicked zooms in, and Alt-↑ works as
  ever. Zoomed in, everything shows.

  On a cell without a grid, either `grid-view` draws a 2-D value or a
  table of row maps as `table` does: a read-only table as big as the
  value. So `(filter-rows orders :region "north")` with `(grid-view)` is
  a box of the matching rows that grows and shrinks with them, with no
  sub-grid to size (a renderer keeps the result from spilling).
  Both **fill the cell** (except on a mesh, §41, where a grid is drawn at
  its columns' width so its lines meet the sheet's): columns stretch (in
  proportion to their widths) to the column's width, so widening the column widens the grid's columns
  while keeping the same ones in view. Rows stretch to fill the cell when
  you've set the row's height; an auto-height row is as tall as its
  tallest content, so grids don't stretch it (two filling grids in one row
  would otherwise hold each other up). If the cell is smaller than needed,
  `(grid-view rows cols)` narrows its columns (to half at most) to keep
  `cols` in view. The whole-grid view of a grid is never smaller than
  the grid: its column and row make room for it (§40), and a set row
  height stretches it only up to that height. Sizing is measured in
  the browser, and refits when the cell is resized or contents change.
  A cell shorter than its grid (a window in a set row height, a span)
  scrolls it up and down, and its columns then fit the width beside that scrollbar.
  Any other renderer works too, e.g. `bars` over the grid's values. When the
  contents have an error, the renderer gets a nil value and `:error`, and
  `grid-view` still draws the grid so the faulty inner cell can be fixed.
* **Inline grids are live.** Click an inner cell to select and edit it;
  arrows move within its grid, scrolling the window to keep the active cell
  in view. The ▦ corner of an inline grid (or Alt-Up)
  selects the cell holding it.
* **Zoom.** Click a grid cell's link, double-click it, or press Enter or
  Alt-Down, and that grid becomes the main view, with a breadcrumb
  (`$ › E15 › C1`) and Out / Alt-Up to go back. The rows/cols inputs size
  the viewed grid.
* **The grid itself, zoomed in.** The ▦ in the top-left corner of the
  grid you're zoomed into (or the breadcrumb's last step) selects the
  grid as a whole, staying where you are: the Inspector shows the grid's
  own settings (its name, About, size, what shows in its cell) and the
  formula bar its cell's formula; the corner and the breadcrumb's step
  are lit while it's selected.
  Double-click the corner to zoom out. An arrow key, Tab, Enter or Esc
  goes back into the grid, at its first cell; typing or Delete would
  replace the grid you're in out of sight, so they're refused, saying
  how (zoom out and select its cell there), and so are cut and paste;
  copy copies it. The top grid has no cell of its own, so
  its corner selects nothing.
* **One trail: the breadcrumb.** The breadcrumb says which grid you're
  in; the name box says where in it: `A1`, or `hike>A2` in a grid drawn
  inside it, not the whole path again. A path typed there is read from
  the grid you're in first (`hike>A2`), then from the top grid
  (`basics>A1`); `$>` starts at the top always. Beside its size, the
  breadcrumb says what's protected in the grid you're in, with the lock
  that turns protection off (§47).
* **A name, and its address beside it.** Wherever RecDep names a cell,
  a grid or a range for you to read (the Inspector, messages and
  refusals, tooltips, the breadcrumb, sheet tabs, the Outline, Find,
  the Map, Profile, the Tests tab, an error's box, an error a cell
  shows), it writes the names on its path and then its address in
  brackets: `order-card (B5)`, `settings>calendar (B6>B9)`,
  `tasks>A1 (B20>A1)`, `resources>rate (B19>D2:D8)` for a table's whole
  column by its heading. So the address the sheet shows (the row and
  column headers, the name box) is always in what's said, and the name
  too. A place with no name anywhere on its path is its address alone
  (`B6>B9`), with nothing in brackets. A step into a parameter's grid
  is the parameter's name, `⚙loan`, in the brackets too (parameters have no address on the sheet). Inside a grid, a list of that
  grid's own cells writes them from it: `calendar (B9)` in Settings.
  An error that names a cell says it again when that cell is renamed.
* **Paths you type.** The name box reads a path by names or by
  addresses, as a formula may write it (`basics>A1`, `B6>A1`,
  `basics>A1:B3`), and shows the active cell from the grid you're in
  (`hike>A2`). What you type, and formulas, keep the form you wrote:
  RecDep never rewrites a name into an address there. The URL hash and
  files keep addresses.
* **Every one a link.** In the Inspector (*Depends on*, *Used by*,
  *Starts at*, *References*, the heading), the Tests tab and an error's
  box, each cell named is a link to it. *References* says what each
  reference in the formula resolves to: `settings>calendar →
  settings>calendar (B6>B9)`.
* **Lists as you'd type them.** A cell holding a list shows it in the
  Inspector's *Value* as you'd type it, `Ana, Ben, Dee`, not
  `["Ana" "Ben" "Dee"]`; so do the steps' previews. A list that typing
  wouldn't read back (an item with a comma in it) is shown as written in
  a formula.
* **Editing.** The Inspector's *Make grid* puts a grid in the active cell,
  replacing its formula; the Inspector resizes or removes it. A grid cell
  can't take a formula. Delete removes a grid with everything inside it
  (undoable). Shrinking a grid (the Inspector's Resize, or the rows and
  columns boxes for the grid in view) removes what it leaves past the new
  edge, so nothing is kept out of sight, to be carried by copies, files
  and the shelf unseen. When any of it holds content (a formula, a value,
  a grid), you're asked first, with where it is; formatting alone (a
  style, a number format) just goes. Undo brings it all back.
* **Copying.** Copy, paste and fills carry a grid with all its contents,
  including nested grids and inner names (a copied grid
  is a new scope). Formulas inside keep their positions in the new grid, and
  `..>l1` reads the new neighbour; a reference to a coordinate in the grid
  around, `..>B8`, moves with the copy as the cell's own would (§5). Pasting content onto
  a grid cell replaces the grid; paste values writes the grid's map.
* **Sizes.** A grid's size is set by hand, except for a grid that a
  formula makes, `(as-grid …)`, which is as big as the formula's result
  and follows it (§22).

A grid can also be called as a function, with some of its cells set to
other values: see §7.

### Saving grids and inserting files

* **Save** saves the whole sheet, with the values it has calculated, so
  that Open shows them at once (§3, *Opening a sheet shows its values at
  once*; the file's `:values`, §9), wherever you are zoomed in.
  **Save grid**, in the bar above the grid just before *Export CSV…*,
  saves just the grid in view as a sheet of its own, without values
  (at `$`, the whole sheet, it's greyed out: that's Save's):
  the grid becomes the file's top grid, with its
  nested grids, names, sizes and libraries: its own library, with the
  libraries of the grids around it put first, becomes the file's top
  library (§14), so its formulas keep working.
* **The file's name.** Save opens the system's save dialog (Chrome,
  Edge: the File System Access API), the name and folder yours, the
  sheet's or the grid's name suggested (`sheet.recdep.edn`, the grid
  cell's name, or its address, `E15_C1.recdep.edn`). The file is
  remembered for the visit: the next Save writes back to it with no
  dialog (its tooltip says where), **Save as…** (beside Save once there's
  a file) picks another, and a sheet opened with **Open** (its open
  dialog) is saved back to the file it came from (an Excel workbook
  isn't: Save writes a sheet). The sheet and each grid saved on its own
  are remembered apart;
  New, an example and an Open forget the sheet's. If writing back fails (the
  file moved, permission refused) the dialog asks where. Without the
  dialog (Firefox, Safari) Save asks for the name, `.recdep.edn` added
  if it's left off, and downloads the file; every Save asks again.
* **While a file is read or saved** the status bar says so, with how far
  it got: *Reading sheet.recdep.edn… 40%*, *Saving… 60%*. The page stays
  usable meanwhile (the work is done off the page's thread, §9). What's
  saved is the sheet as it was when you pressed Save; pressing Save again
  before it's done says *Still saving*. While a file is being read, the
  sheet you see is still yours to edit; the opened one replaces it when
  it's ready, and Undo brings back the old one, edits and all. Another
  Open, New or an example meanwhile wins over the file being read.
* **Broken links.** References inside the grid that reach outside it
  (`..>B2`, a `call` to an outside grid) or start at the root (`$>A1`, which
  would mean the file's own top grid) won't survive. Save lists them and
  asks before saving. `..>` that stays inside the saved
  grid is fine.
* **Insert** reads a file into the selected cell as a sub-grid,
  replacing what's there (asking first; undoable):
  * a saved sheet (`.edn`): its top grid, with everything in it. Its top
    library becomes the new grid's library (§14), less the copies of what
    the grids around it already define the same. References starting at
    `$>` now point at this sheet's top grid, and are reported;
  * CSV / TSV: RFC 4180 — quoted fields may hold the
    separator, line breaks and doubled quotes; the separator (comma, tab or
    semicolon) is detected from the first line. Numbers become numbers,
    `true`/`false` booleans, dates and times dates (with a date format,
    as typed), everything else text; blank fields stay empty;
  * an Excel workbook (`.xlsx`): its contents page, with every sheet a
    grid inside it (§17).

* Big files: see §9 for what sizes work well and what's limited.

### Inputs bound from outside, and parameters

`..>` and `$>` let a grid read its surroundings, which is quick but ties
the grid to where it is: saved on its own, inserted elsewhere, or used
as a template (§21), those references break or read the wrong cells. A
grid can instead take *inputs*, given values where it is used, so
nothing inside it mentions what's around it.

**Using it**

* **Inputs are parameters.** A grid's inputs are its parameters: named
  values, or tables, in its margin, above its cells (⚙ Parameters, in
  the bar over the grid, shows them). The value a parameter holds is its
  *default* (a table's rows, its *sample*), so the grid works on its
  own. Its named cells inside are its own, its data and workings, never
  inputs; the named ones are its *outputs*. A cell inside can show a
  parameter by holding just its name (`amount`): an instance's card then
  shows the parameter there, to type into.
* **What a parameter takes.** Beside each parameter in ⚙ Parameters,
  *takes anything* (or *takes any table*) is a choice: a value parameter
  takes numbers, whole numbers, dates, text, TRUE or FALSE, lists, or
  numbers at least or between bounds you type, or a table of a shape in
  scope, a built-in one (a calendar) or one of your own (a route): its
  own value then becomes a sample of the shape (its headings and an
  example row, to fill in), so the grid still works on its own. A
  table parameter takes any table or a shape. A whole grid given where a
  table is taken (a block, `working-calendar`) is refused, naming the
  tables in it that fit (`working-calendar>calendar`). Then anything
  else it's given, bound from outside, by an
  instance's formula, or as its own default or sample, is refused: the
  parameter is an error saying why (*shapes>trip>⚙pts (B25>B8>⚙pts) takes a route: row 2's x is
  "abc": a route's x holds numbers*), shown beside it in the strip, and
  nothing in the grid computes with it. Strict: a table needs the
  shape's needed columns, and every value in a column the shape gives a
  kind must be of it (a blank is fine). Lookups, IDs and dates on or
  after another column aren't checked here; the table's own cells flag
  those. A parameter taking dates is given a date format if it has none.
* **What fits, offered.** While nothing gives a parameter that says
  what it takes one, or what does is refused, the strip lists beside it
  what in the sheet fits: the tables of its shape (*fits: hike*), or the
  named values of its kind, nearest first. A click binds it.
* **Values as they're formatted.** A parameter's value shows in the
  strip as its number format says, as a cell would: a date as a date
  (*4 Jan 2027*), a currency as money, whether it's its own or given.
* **Make parameter.** Right-click a named value or table inside a grid
  → *Make … a parameter*: it moves to the margin under its name (a table
  with its sample), and the formulas reading it by name go on doing so.
  Giving anything else to a grid from outside is refused, and the
  message points to it.
* **Private names.** A name starting with `_` (`_first-day`,
  `_days-in-week`) is private to its grid: its formulas read it as any
  name, but it's no part of the grid's interface, neither input nor
  output. Names make formulas readable; without this, every helper
  named for that reason would be an output (a template's tile, a recorded
  test case) or, holding a value typed in, an input to wire. One
  character, chosen with the name, is all the packaging there is: no
  switch to set. Left out: the ⚙ Parameters strip's wiring, a template's inputs
  and outputs on the shelf, what *Record a case* records, the tables a
  block's input is offered (a table so named, or inside a grid so named:
  the Project plan's Baselines block keeps its comparison in `_compare`,
  its own view, so a loading view isn't offered its tables), and completion
  from outside the grid (`cal>` offers its outputs only, and not the
  private cells' addresses). A reference to one from outside still
  works, but is flagged by the formula checks (§26): it reads something
  the grid doesn't give, free to change. A binding may still name one
  (`{:_rate 0.5}`), which makes it an input after all.
* **Bindings** are the formula of the grid cell itself: a map from the
  grid's inputs to formulas written in the grid cell's place.

  ```clojure
  ;; the tour's dashboard, C3 of the top grid; inside, A1 is named shopping, B1 tax-rate …
  {:shopping basics>grand-total :tax-rate basics>tax-rate :loan loan>B2>payment}
  ```

  Inside, formulas just say `shopping`. The bindings read as if written
  in the grid cell: relative references are relative to it (`{:amount
  Au0}` filled down a column binds each copy to its own row), names are
  the grid cell's grid's, and so is the library. A key is a name in the
  grid, a cell's or a parameter's, holding something: names are a grid's
  interface, so an address (`:B2`), which would point elsewhere once rows
  are inserted, is refused (`call` still takes one, for a quick look). A
  key the grid doesn't have makes the grid cell a `#REF` error, as `call`
  reports one; the inner cells keep working. A grid cell takes no other
  formula. Bind an input to the grid cell itself and it's a `#CYCLE`.
* **A bound input** has its binding's value, and keeps its own as the
  default: remove the binding (or the whole map) and it's back. It never
  spills: bound to a range or a table, that's its value. A grid bound to
  an input can be called (`{:loan loan>B2}`, then `(call loan {…}
  :payment)` inside), which passes a template in rather than reaching
  out for it.
* **Parameters** are named values in a grid's *margin*, outside its
  cells, for inputs that shouldn't sit in the grid's data. No range,
  `ub`/`de`/`re`, header column, spill or grid value reaches them, and
  inserting or deleting rows and columns doesn't move them. They're used
  by name, inside (`rate`) and from outside (`B2>rate`), and bound,
  called, pinned by instances and test cases, saved and copied as named
  cells are. A parameter holds a value, not a formula (bind it from
  outside instead), and can't be the end of a range. Names are shared
  with the grid's cells.
* **Copying** the grid cell carries its bindings, shifted as any formula
  (`{:total A1}` pasted one row down and one right is `{:total B2}`);
  inserting and deleting rows rewrites them, a deleted target becoming
  `REF!`; resizing keeps them.
* **Save grid** leaves the grid cell's bindings behind: the file's grid
  computes with its defaults, and they aren't reported as broken links.
  A grid inside it bound from outside the saved grid is.
* **Templates passed in.** A template (§7) used from inside a grid is
  best given to it as an input, `{:loan loan>B2}`, with `(call loan …)`,
  (the tour's Spilling, What-if and Linked instances sections take its
  loan template so),
  `(instance loan …)`, `(what-if loan …)` inside: the template's place
  is written once, in the grid cell's formula, and *Make input* on
  `(call $>loan>B2 …)` does exactly that. Save grid then copies the
  template into the file, as that input's default (a parameter can hold
  a grid), with its library and those around it, and an instance's
  overrides written in as values, so the grid's calls work on their own.
  Bound again after Insert, the binding wins over the copy. References
  out of the template are listed with the grid's (*Reaches out*, and
  Save grid's warning), since the copy would break them. The strip shows
  such a parameter as `▦ B12>B2` (given; a click goes to the template) or
  `▦ template` (its own, a template held there, §35; a click opens it, as an instance's ⧉ link and the
  Inspector's name link do). Inside, it's an ordinary grid, to read and
  edit; the breadcrumb and the Inspector name it by its parameter
  (`$ › ⚙ square`, `⚙square>A1`), and the Outline, error counts and
  checks see into it.
* **Test cases** (§27) run on the defaults of the grid's own inputs, not
  their bound values, so a grid's tests pass the same way before and
  after Save grid. `call` evaluates the grid where it is, so its bindings
  apply. *Record a case* counts bound inputs as inputs.
* **Make input** turns a reference out of a grid into an input of that
  grid, bound from outside: the refactoring from a quick `..>` to a
  grid that works on its own. On a cell using one, right-click → *Make
  input from ..>A1…*, or *Make input* beside it in the Inspector's
  *References* (or its *Check*, or a grid cell's *Reaches out*).
  * If the reference is the cell's whole formula (`..>basics>total`),
    the cell becomes the input: it keeps its name (or is asked for one),
    and the value it read becomes its default. Not when it spills (`..>A1:A9`
    alone): an input doesn't, so the cell reads a new parameter instead.
  * Otherwise a parameter is made, named from the reference
    (`..>basics>tax-rate` suggests `tax-rate`, numbered if taken) or
    asked for, holding the value read as its default.
  * Either way the grid cell's formula binds the input to the reference
    as written from there (`..>basics>tax-rate` becomes
    `basics>tax-rate`; written afresh when that wouldn't read the same
    cells, as for `..>A1:..>C3`), and every reference inside the grid to the same
    cells is rewritten to the input: `tax-rate`, or `..>tax-rate` from a
    grid deeper inside. One undoable step.
  * A reference in a binding (a grid inside, bound from its grid) makes
    an input one grid further out: repeat to bring a value through
    several levels. A grid read this way (a template, `$>loan>B2`) is
    passed in: see *Templates passed in*, above.
* **Inputs that take a concept.** A block's input can say what it
  takes: say it of the input's cell (This table is ▸ on the Shelf, for
  a table input). Placing the block (Insert, or the shelf as a copy)
  then binds it by what the sheet's tables are: one said to be that
  concept is bound at once (or, with several, the one named as the
  input is), and Insert's message says so. Otherwise its chip in the
  block's ⚙ Parameters lists what could be it: a table said to be it; one
  whose columns fit but which says nothing (*cal is capacity*: binding
  it says so); one that fits once a column is renamed (*booked,
  renaming Taken to Takes*: bound through the renames, which the
  binding shows). A table the block's own steps can't read (a join on a
  column it hasn't got) is shown with why; when nothing fits, the
  nearest is, with what it lacks.
  * **A finer table, summed.** When the input's own sample has fewer
    columns than the table (effort by task, from a log by resource, day
    and task), a measure it has is combined over the rest as its column
    says it combines (summed, averaged, the smallest, the largest), and
    the binding shows the group step that does it (*Hours summed by
    Task, over Resource, Day*). Only tables of facts are summed so: a
    calendar's rows are rules, never added up.
  * **Made from what's here.** When nothing is it, but something can be
    made from what the sheet has, the Inputs row offers that too: a plan
    by the Schedule step from the tasks and a calendar; bookings from a
    scheduled plan; a load (each resource's capacity and days booked,
    week by week) from resources, a calendar and bookings. At most two
    such steps in a row. Each offer says how, in words (*make: the
    Schedule step of jobs, cal*), and what it guesses (*start left at
    (TODAY)*: no cell named start; name one, holding a date, and it's
    used). Clicking it places what it makes beside the block, as an
    ordinary named table said to be what it is (*plan-table*), and binds
    the input to it; one Undo takes it all back. Insert's message says
    when an input can be made so. If what it would make has an error,
    nothing is placed, and the message says why.
  * **The best first.** The offers are ranked, the first three shown: a
    table that is it before one that has to be made; fewer guesses
    first; a table as it is before one renamed, and one renamed before
    one summed (renaming loses nothing; summing loses detail); then one
    named as the input is.
  * **Instances too.** Placed as an instance, a block's inputs are given
    in its formula, `(instance counter {:avail cal})`: placing binds the
    certain ones there, and the Inspector offers the rest for any input
    the instance doesn't give.
  * **When a binding stops fitting.** If what an input is bound to gets
    an error, or isn't a table of what it takes any more, its Inputs row
    says why and offers again.
* **An example.** Examples ▾ → *Project plan* opens on its facade
  (§54): the plan as its team sees it, the dates to set, where it
  stands, the Gantt and the task picked in it, with resources, tasks
  and load as its sheets. *Advanced*, beside Home under the grid, shows the rest: a
  project plan built from blocks (§46, *The project plan*): a working calendar, a
  timesheet, a scheduler, baselines, tracking, a Gantt, two loading
  views and a next project, each a self-contained grid saying what it
  takes and gives. The sheet holds the resources, the tasks and the
  dates, and binds each block's inputs to them or to another block's
  outputs: open a block and its ⚙ Parameters show each input, what it
  takes and what it's bound to; its body holds only its own data and
  workings. Tracking's time spent is the timesheet's effort,
  summed by task in its binding; the two loading views take the same
  load, so either can replace the other. Its tasks ask for people in
  three ways: by name (Resource), by role (Role: any one of the role's
  people, from the sheet's Roles table, where Ana is a designer and a
  developer), or as a follow-on (Follows on: whoever got that task).
  The Scheduler block takes the roles as pools and the timesheet's
  effort, from which it works out who is on each task begun; its plan's
  Resource is who was assigned, so Tracking and the Gantt read people.
  The Timesheet block takes the tasks too: its Task is picked from them
  with the ▾ and shows each task's name, though it holds the ID.
  *Assignments to check*, in its summary, counts the tasks whose
  who-check says something.
* **About.** A grid says what it's for: the Inspector's *About* row (a
  grid's cell selected) holds a few words, and the words it's known by
  beside them. They're its own, not a note on it: saved alone, put on
  the shelf, copied or inserted, it carries them, and the Shelf's search,
  Insert's message, the grid's hover and a formula's hover over a
  reference to it show them.
* **Self-contained: a block.** Tick it for a grid cell in the
  Inspector, and the grid is a **block**, with a boundary: a promise
  that it can be saved, copied or reused without breaking, kept as you
  go rather than found out at Save grid.
  * **It reads only its inputs.** A reference out of it, at any depth
    inside it (`..>A1`, or one starting at the root, `$>…`), is refused
    where it's written: the cell shows `#REF`, *reaches out of B2, a
    self-contained block, by ..>A1*, and *Make input* (its right-click
    menu, or the Inspector's *Reaches out*) binds the value from outside
    instead. The grid's own bindings are written outside it, so they
    aren't judged against it. A rule of a grid inside it is held to the
    same: a lookup's column out of it (*one of ..>..>tasks>id*) doesn't
    apply, and its rule says *reaches out of B27, a self-contained
    block…* in the Rules tab and on the column's heading.
  * **An input's columns are what's bound to it.** A table input's
    column, `tasks>id` (from a grid inside the block, `..>tasks>id`),
    reads the table bound to the input, not its sample: in a formula,
    and as the column a lookup picks from, so a block's Timesheet can
    take the tasks and pick its Task from them. Unbound, it's the
    sample's column again.
  * **Its code writes only inside it.** A button or control in it
    writing a cell out of it (`(set-cell! ctx "..>A1" 0)`, or one of its
    bound inputs, from another cell) is refused when pressed, saying so.
  * **What you type through it reaches outside only where you allow
    it.** Typing into a bound input inside a block (or a control on one,
    a date picker on its own cell) would write the cell it's bound to,
    outside, and so would typing into a view inside it showing that
    input's table (a plan's Task, read from the tasks bound in): that's
    refused until you tick *views may change it* beside the input's
    binding in the grid's *⚙ Parameters*, off by default. It's said as
    soon as you start (a key typed, F2, a double-click, a click in the
    formula bar, the ▾, the input's field in *⚙ Parameters*), naming the
    block and the cell it would write, so nothing typed is lost; until
    then the view's cells aren't marked as writing back. Ticked, it
    writes there, the cell's own rules applying.
  * A note's links out of it (`[[$>basics|…]]`) are flagged, since they
    break when the grid is saved alone. A grid cell's Inspector
  lists what *Reaches out* of its grid, ticked or not (what Save grid
  would warn about, writes and links included), and what it *Borrows*:
  the functions it uses from the libraries of the grids around it,
  which Save grid carries into the file (just those, and what they use,
  §14), so they aren't broken, but are a dependency all the same. The
  tour's dashboard is ticked.
* **Controls** (`slider`, `number-input`, …) on an input bound to a
  single cell holding a value (`{:rate C5}`) write that cell, following
  further bindings. Bound to anything else, a formula or a cell holding
  one, they're refused (they would only change the default, out of
  sight), as an instance's controls only change typed-in inputs (§21).
  Typing into such an input is refused the same way, as soon as you
  start: change the binding in the grid cell's formula, or the input's
  default in the Inspector.

**In the app**

* **The parameter strip.** *⚙ Parameters* in the breadcrumb shows or
  hides the parameters of the grid you're working in, in a strip above
  it (shown by itself when the grid has some): the grid in view, or,
  while the active cell is inside a grid drawn in place that has
  parameters, that one (*Parameters of J2*), as the Outline lists them,
  so a template a grid keeps (one Insert kept, say) is seen, and removed,
  where it is. Each is `name = value`: edit the value
  and Enter; × removes it (undoably); double-click the name to rename it
  (formulas using it, and the bindings giving it a value, follow the new
  name, as for cells, §2). *+ name =
  value* adds one. While writing a formula, clicking a parameter's name
  inserts it (`B2>rate` from outside the grid). A value only: a formula
  typed there is refused, with a pointer to binding it instead. The
  formula checks (§26) look at parameters too: one nothing reads, or one
  whose binding reaches out of a self-contained grid, is marked in the
  strip (with *Make input* for the latter), counted with the rest, and
  the status bar's ◤ link opens the strip to it.
* **Bound inputs** have a corner mark (bottom right) and say what binds
  them in their tooltip. The formula bar shows the value they get, not
  their default; typing a value writes it where the binding reads, when
  that's one cell holding a value (and says so), and is refused
  otherwise, pointing to the binding. Delete, paste and fill over a
  selection leave bound inputs' sources alone, and say so; their defaults are set
  in the Inspector. A bound parameter's value in the strip works the same way,
  drawn in the accent colour.
* **⚙ Parameters, where inputs are set.** Open the grid: each parameter
  in its strip shows its default (a table: *▦ table · 21 rows*, to open
  and change its sample), what it takes when it says (*a load*), and,
  beside ←, what the grid's cell gives it: a field holding the binding
  (a formula written in the grid cell's place; Enter changes that entry
  of the cell's map, one undoable step; blank unbinds it). When nothing
  gives a parameter that takes a concept, or what gives it stopped
  fitting, the offers are there too; for a block, *views may change it*
  is the edit grant. The Inspector, on the grid's cell, says it in a
  line (*1 parameter, 1 given (load ← load)*) with a button opening the
  strip. The map is edited in place, as text: the other entries,
  comments and layout stay as written, a new entry goes on a line of its
  own in a map written a line each, and an entry taken out goes with its
  line and its comment (the views' and panels' options and instances'
  overrides are edited the same way). The grid cell's formula, in
  the formula bar, is the same map, to edit whole.
  *Names here* lists parameters as such; clicking one shows the strip.

## 7. Grids as functions

A grid can be used as a reusable calculation — a spreadsheet LAMBDA made
from a nested grid. Build the calculation once, in a grid, then call it
from any cell with some of its cells set to other values; the grid itself
isn't changed.

```clojure
;; E32 holds a template: principal, rate, years -> total (all named)
(call $E$32 {:rate 0.07} :total)            ; one result: the cell named total
(call $E$32 {:years 20 :rate 0.06} :total)
(call $E$32 {:rate 0.07})                   ; the whole value map {:A1 .. :B4 ..}
(mapv #(call $E$32 {:rate (/ % 100)} :total) (range 1 6))   ; a what-if row
(call $E$32 :total)                         ; one cell, as the grid is
```

`what-if` and `goal-seek` (§18) evaluate a template the same way, many
times; `instance` (§21) keeps a template with inputs of its own, drawn and
linked.

* `(call grid overrides output)`: `grid` is any reference to the cell
  holding the grid (`E32`, `$E$32`, `..>Loan`, a name); `overrides` is a map
  from cells to values; `output` picks the cell whose value is returned.
  Without `output`, the grid's whole value map comes back; `(call grid
  :total)`, with a keyword in place of the overrides, reads one cell.
* Keys in `overrides` and `output` are cell addresses (`:B2`) or names in
  that grid (`:rate`). Name the input and output cells: calls then read
  like function calls.
* **Partial application comes free:** leave an input out and the template's
  own value is used. Override a cell that holds a formula and the formula is
  replaced, for that call only.
* Edit the template, or anything it reads, and every call re-runs. The
  Inspector's "Used by" on the grid cell lists its callers.
* **A call with an output works out only what that output needs**, and
  works that out the same way each time: called many times (in a
  simulation's runs, a what-if table, goal seek), it costs about the
  template's own work for that output, however big the sheet around it.
  So when you need one result, name it: `(call scheduler {:tasks t}
  :plan)` doesn't work out the Scheduler's bookings.
* Use `$` (`$E$32`) when the calling formula will be copied or filled, so
  every copy calls the same template.
* `call` works in cell formulas, not in renderers.
* When the template fails for these inputs, the error names the inner
  cell where it started, by its name and address: `call: E2>pick (E2>A3): No item 5 in
  vector of length 2`, even when the output is a cell that only reads
  `pick`. `what-if`, `goal-seek` and `instance` say the same. A time
  limit reached inside the template is the calling cell's `#TIME`.

## 8. Time limits

Each cell formula and each renderer call has a time limit, so an endless
loop or sequence doesn't freeze the page: going over it stops the
evaluation with an error.

* A formula that runs too long shows **`#TIME`** (a renderer, **`#RENDER`**);
  hover it, or see the Inspector, for the message. Cells that use it show
  `#UP`. Fix the formula and everything recovers.
* **Every cell's time** is 1 second, unless you set another: the status
  bar's **Auto ▾** menu, *Time for one cell: 1 s…*, asks for any number
  of seconds. It's kept as `(def time-limit-ms 3000)` in the top grid's
  library, so you can write it there too. Raise it for formulas that are
  slow but finite, or lower it to catch mistakes sooner.
* **One cell's own time.** A heavy cell (a simulation of many runs) may
  need longer while every other cell should still stop after a second.
  On a cell stopped with `#TIME`, its menu and the Inspector's *Time* row
  offer **Allow this cell more time…**, which asks for any number of
  seconds (twice what it has now, suggested). It's kept on the cell,
  saved with it, and the Inspector says *may run 3 s, its own*; *Change…*
  there changes it, and a blank gives it every cell's time again.
* Typical causes: `(range)`, `(iterate …)`, `(repeat x)` or `(cycle …)`
  consumed whole (`count`, `reduce`, `last`, or just returned as a value);
  `loop`/`recur` or `while` that never ends; a function that never stops
  recursing (often a stack overflow, `#ERR`, rather than `#TIME`).
* A `call` runs inside its caller's limit, and a timeout inside it is a
  timeout for the caller.
* What it **doesn't** do: the page still waits for up to the limit on a
  slow-but-finite formula, and several slow cells add up, since the limit
  is per cell. A single long core operation that doesn't step through a
  sequence (sorting a huge vector, say) can't be stopped part-way. Nor can
  one call of an Excel function (§17), which is built in rather than
  written in Clojure; each finishes in time proportional to its input,
  but a `LAMBDA` passed to `MAP` or `REDUCE` is user code and is checked.

## 9. Large grids

Grids can be big: a grid can have up to a million rows and a thousand
columns, and grids with hundreds of thousands of filled cells stay quick to
scroll, jump around in and edit. Getting the data in, and the browser's
memory, are what limit size.

* **Insert a big CSV** into a cell (Insert, §6) and open it (Alt-Down or its
  link). Measured in Chrome with 10 columns of numbers:

  | rows | filled cells | insert | open (draw) | edit a cell | jump to the end |
  |---:|---:|---:|---:|---:|---:|
  | 2,000 | 20,000 | 0.6 s | 0.12 s | 65 ms | 0.10 s |
  | 20,000 | 200,000 | 5.7 s | 0.08 s | 66 ms | 0.13 s |
  | 50,000 | 500,000 | 18.7 s | 0.13 s | 82 ms | 0.13 s |
* **Memory is the practical limit**: about 2 KB per filled cell (500,000
  cells ≈ 1 GB). Inserting a file with more than 200,000 filled cells asks
  first, with an estimate of the time and memory. A million cells (100,000
  rows × 10) takes over a minute to insert and about 3.5 GB — beyond a
  comfortable browser tab.
* **Empty space is free**: a million-row grid with little in it costs
  little.
* **Opening and saving a big sheet** don't hold up the page: the file is
  read, and the saved text finished, off the page's thread, with
  *Reading… n%* or *Saving… n%* in the status bar. For 300,000 cells,
  reading takes about 2 s and saving about 1 s. Building the sheet after
  reading it is done in slices too (*Loading… n%*, §3), then its values
  are calculated in slices, the cells on screen first. Meanwhile the
  values it was saved with show, dimmed (§3).
* **The file keeps the calculated values**, after the cells:
  `:values {"A1" 1 …}` for a value on its own, `:tagged-values {"B2"
  {:value 5000 :unit "km"} …}` for one with more (an error, a unit, a
  spill). They make the file about half as big again for a sheet of
  numbers and formulas (+55–58%), Save about 1.5 times as long and the
  reading about 1.8 times; they are only ever shown: a file without them
  opens all the same, and RecDep reads none of them as a cell's value.
* Sheets with more than 20,000 cells aren't **autosaved** in the browser
  (writing them after every change would stall the page, and browser
  storage holds about 5 MB): the status bar says so — use Save.
* Selections over 50,000 cells show their size in the status bar but no
  count / sum / average (it would be recomputed after every change).
* Formulas over big ranges (`(SUM A1>A2:A50001)`) cost one entry, and the
  Inspector shows the range as one link.
* **Just after opening a big sheet** (or changing a library, which
  rebuilds it, or inserting rows near its top, which moves most of its
  cells), the status bar shows **…** for a moment
  while it counts the sheet's errors, units warnings and cells to check;
  the Outline says *Counting the sheet's cells…*, and the green check
  corners and cells drawn over their neighbours (spans) appear a moment
  later. The page stays usable meanwhile: the counting is done in short
  slices between frames. After an edit of a few cells the counts are
  there at once; one that changes tens of thousands of values can show
  … for a moment too. The check corners and spans aren't held up by it.
* **Edits that change many values** are calculated in slices, the page
  answering meanwhile: cells not yet calculated are dimmed and the status
  bar counts them down (*Calculating while you work*, §3). At 300,000
  cells, an edit that 30,000 cells depend on holds the page for about
  60 ms, and the first unit typed into a sheet for none
  of its own; the values arrive over about 1 s and 10 s. Opening a sheet,
  changing a library and renaming a much-used cell compile every formula
  they touch again in slices (*Loading and preparing*, §3), then
  calculate in slices, the old values showing dimmed meanwhile.
  Inserting or deleting rows only moves what's below: a row inserted near
  the top of a 300,000-cell sheet takes about 9 s of slices, one near the
  bottom a moment.

## 10. Formats for columns, rows and ranges

How numbers look is a cell's **number format**, as in Excel. A number
format, renderer or style can belong to a whole column ("all of D is
currency"), a row, a range or a grid, and apply to every cell there —
including cells filled in later — without touching the cells. A rule can be conditional
(`{:when "neg?" :style {:color "red"}}`), which makes "negative numbers in
red" data rather than a hand-written renderer. A cell's own setting wins.
Validation rules (§16) sit in the same list.

* **Number formats.** The format bar's format field gives the selection a
  number format: an Excel format code (`#,##0.00;[Red]-#,##0.00`, `d mmm
  yyyy`, `[$€]#,##0.00`), typed, or chosen and varied in its gallery
  (below). It's
  stored as the cell's `:format` (or a rule's, below), shown by `plain`
  and so by `default-render`; renderers get it as `(:format ctx)`. A cell
  without one shows numbers as Excel's General does: at most 11
  characters, so `(+ 0.1 0.2)` shows 0.3 and a long number is rounded or
  goes to scientific notation, and booleans as TRUE / FALSE. That applies
  to every sheet. A cell a spill reaches, with no format of its own or by
  a rule, shows its value in the anchor's format, as Excel's dynamic
  arrays do. Copy, paste (all, and paste formats) and *Clear formats* (cell
  menu) carry or clear it with the renderer and style; Copy puts a cell's shown text on the
  clipboard (a date as `2026-11-07`), and Find searches numbers as they
  show. Entering something into a cell without a format can give it one,
  as in Excel: see §15.

* **Headers.** Select a whole column or row with its header (or the grid
  with Ctrl-A) and set a number format, renderer, bold, a colour,
  alignment: it becomes a rule for that column, row or grid. In a grid
  made by a formula with a header row (`as-grid`, a view), a column's
  rule names it by its header (`{:column "start" …}`, below): its columns
  move when the formula's rows change, and the rule moves with its
  column. That rule leaves the header row as it is, and the status bar
  says so the first time. As in Excel,
  the cells' own settings of what you just changed give way to it (their
  own formats or renderers, or the style properties you set); the rest of
  their own formatting stays.
  *Clear formats* (cell menu) on a header selection removes its plain
  rule as well; *Clear style* removes only that rule's style.
* **The Rules tab** lists the rules of the active cell's grid, one a
  row, formatting first (the higher one wins, *Which wins* below), then
  kinds, validations and protection (all apply): its type (*Format*, *Conditional*, *Validation*, *Kind*,
  *Protect*), where it applies (*column K (Done), from row 2*) and what
  it does (*as 0% · bold*, *when above 100 · text red*). Click a row:
  its cells are outlined on the grid, and its editor opens under it,
  with the fields its type has: where it applies (a column by its
  heading, the selection, or cells typed as `B2:B9`), a number format,
  a renderer, bold, italic, alignment, colours and a unit; for a
  conditional rule *When the value is* negative, above, below,
  between, equal to, text containing, or a condition written as a
  formula; for a validation *Must be* a number, a whole number,
  between, one of a column, one of values listed, text no longer than,
  or a formula, its message, and whether a failing value is refused;
  for a kind what the cells hold, with its bounds, its column or
  values, what a blank reads as and a message; for protection,
  locked or open. Nothing changes until **Apply**; what can't be applied
  (cells that aren't cells, a condition that doesn't work, a default
  outside the bounds, a protected cell) is said under its field.
  *+ Format*, *+ Conditional*, *+ Validation*, *+ Kind* and *+ Protect*
  open the editor for a new rule on the selection. Drag a row to move
  it (a line shows where it will go); ✕ removes. The Renderer field
  opens the renderer gallery beside it, its tiles drawn with the rule's
  first cells: hover one to see the rule's cells drawn with it, click to
  choose (and refine), Enter or a second click to use it. A rule written by hand (in a file) with settings the editor
  has no field for keeps them, and its row says *and more, written by
  hand*. Where the Inspector says a cell's format or kind comes from
  (*from K2:Kde*), click it to open that rule here.
* **The Inspector** shows the number format, renderer and style in effect,
  and which rules (or the spill's anchor) they came from; the formula bar's renderer field shows an inherited
  renderer as its placeholder.

```clojure
;; the grid's :formats
[{:area "D:D" :format "$#,##0.00" :style {:text-align "right"}}
 {:area "1:1" :style {:font-weight "bold" :background-color "#eef"}}
 {:area "D:D" :when "neg?" :style {:color "#c0392b"}}
 {:when "#{:done}" :style {:text-decoration "line-through"}}   ; whole grid
 {:area "B2:C9" :when "#(> % 100)" :format "#,##0"}
 {:area "E:E" :render "(data-bar 100)"}]
```

* `:area` — columns `"D:D"` / `"B:D"`, rows `"1:1"` / `"2:5"`, a range
  `"A1:C5"` or a cell `"B2"`, in the grid's own coordinates. Leave it out for
  the whole grid. (`"D"` and `"3"` are accepted and stored as `"D:D"`,
  `"3:3"`.) A range can run to the grid's edge with the moves references
  use (§2): `"B2:Bde"` is B2 to the bottom of column B (a column below its
  header), `"B2:re"` row 2 from B to the right-hand edge, `"B2:dere"` B2 to
  the bottom-right corner. It keeps reaching the edge as rows or columns are
  added. A column or row left out is the first corner's (`"B2:de"` is stored
  as `"B2:Bde"`).
* `:column` — instead of `:area`, in a grid with headers: the column with
  that header, below it (`{:column "start" :format "d mmm"}`; several,
  `{:column ["pv" "ev" "ac"] …}`). The header is matched as references
  match it (`"Unit price"`, `"unit price"` and `:unit-price` are the
  same). The rule follows the column: a column added before it, or a
  table worked out by a formula whose columns come out in another order
  (a function's rows gaining a column), leaves a rule by letter on the
  wrong column, but not this one. Retyping the header takes its rules
  along; retyped to a header another rule names, the column comes under
  that rule too: where both set the same thing (a unit), the higher one
  in the list wins. A rule naming a header the grid hasn't got formats nothing, and
  the Rules tab says so, listing the headers there are. It's the way to
  format a grid made by `as-grid`. When rules compete it counts as a
  range (the cells below its header), so it wins over a rule for the
  whole column or a row.
* `:when` — optional: a function of the value (`neg?`, `string?`,
  `#(> % 100)`, a set `#{:done}`, a library function). Tested only on cells
  with a value: not on empty cells or errors. One that throws, times out or
  isn't a function doesn't apply. It can read cells, as a validation
  rule does (§16): each reference is read from the cell being tested, so
  `{:area "B2:B9" :when "#(> % l1)"}` compares each cell with the one on
  its left (Excel's `=B2>A2`), and `#(> % limit)` with the cell named
  `limit`. Change the cell it reads and the cells it covers restyle. A
  reference by address (`#(> % D1)`) is read from every cell alike, and
  follows rows and columns inserted or deleted, as a formula's does.
  **Excel's mixed references** work as in Excel: the rule is written for
  its area's first cell and read as if copied to each, so `B$2` is row 2
  of each cell's column (`{:area "B3:H20" :when "#(> % B$2)"}`, Excel's
  `=B3>B$2`: each week's cell against that week's limit in row 2) and
  `$B3` column B of each cell's row (written for row 3). `$B$2` is B2
  everywhere. A plain `B2`, with no `$`, is B2 from every cell, unlike
  Excel's: for the cell itself write `%`, for its neighbours `l1`, `u1`. When
  a cell it reads has an error, the condition doesn't apply, and the
  Rules tab says why.
  **True or false, whatever the cell holds**: a condition can also be a
  formula that is true or false, as Excel's are: `(> limit 3)` applies
  to every cell it covers while `limit` is over 3. A name inside a grid
  isn't seen from the grids within it: there, write the path from the
  top (`$>order-card>order-head>B5`).
* `:format` — a number format code, as a cell's `:format`.
* `:render` — renderer source, as in a cell's renderer field.
* `:style` — a CSS map, as a cell's style.

**Which wins.** The cell's own number format and renderer, and its own
style property by property, win over every rule. Among rules, the higher
one in the Rules tab's list wins, as in Excel's conditional formatting:
where two rules set the same thing on the same cells, the upper one's
applies, whatever their areas or conditions. Drag a rule up to make it
win. A rule the format bar makes, and one added in the Rules tab, goes on
top; one changed keeps its place. Styles merge property by property
across all the rules that apply; the number format and the renderer each
come from the highest rule that has one. Kinds, validations and
protection aren't ranked: every one that covers a cell applies (a cell
left open stays open, whatever locks it).

**Scope.** Rules belong to one grid, like names: column D's rule in the top
grid doesn't reach the grid held in D5. Rule renderers don't apply to cells
that hold a grid — they keep their own renderer or `grid-link` — though
rule styles do. So a renderer set on whole columns, rows or the grid is
also given to the grid cells selected as their own: pick `grid-view` for
a grid held in a grid one row high (where a cell is its whole column)
and it's drawn, as the gallery previewed; clearing the column's renderer
clears theirs.

### The format gallery: building a code by how it looks

Renderers don't do a format's job (§4), so red negatives, hidden
zeros and text around a number are the format's, written in Excel's
codes, which few people know. The format field does for them what the
renderer gallery does for renderers (§4): it shows the choices drawn
with the cell's own value, and builds the code.

* **The field** holds the selection's code (empty is General). Focused,
  it opens a popover at the window's right, under the bars, so the
  selection it previews (usually under the format bar, on the left)
  stays in sight; typing a code previews it live. A change that removes
  the row whose control had the focus puts the focus back in the field,
  so Enter and Esc keep working.
* **The gallery**: formats by kind — General, Numbers, Currency, Percent
  (with scientific and fractions), Text, Dates, Times and durations —
  each tile the value drawn in that code (a date or a duration where
  the value isn't one), the kind that fits the value first, but every
  kind there (a bare 46000 can still be made a date). Hover a tile and
  the grid draws the selection in it; click it and its code goes into
  the field (a second click, or Enter, uses it).
* **Refining**: once the field holds a code the builder can read, the
  gallery gives way to rows of variations on it, each a whole new code
  drawn on a value that shows the difference: *Kind* (General, number,
  currency, percent, scientific, fraction), *Decimals*, *Thousands*,
  *Currency* ($ € £ ¥ CHF; not on a cell whose unit is money, which is
  its currency), *Scale* (thousands K, millions M), *Fraction*,
  *Negatives* (minus, red, brackets, red brackets — drawn on a negative),
  *Zero* (0, blank, – — drawn on 0), *Text before* and *Text after*
  (typed; quoted and escaped for you); a date's *Date* and *Time*
  styles. The panel's own tile draws the selection's first cells as the
  grid will, each by its own renderer: a data bar's label in the new
  format. *All formats* goes back to the gallery.
* **Using it**: the grid draws the draft (a date format drops the unit)
  until Enter, a click on the panel's tile, or focus going elsewhere uses
  it, as one undoable change (so a column
  header's selection sets the column's rule); Esc puts it back with no
  undo step. Only a date format changes more than how the cell looks
  (formulas on it see a date, §25), and that follows when it's used.

## 11. Finding your way around

Grids inside grids make a sheet deep: a reference can lead three grids
down, and a grid behind a link is out of sight. These tools work across
every grid at once.

* **Sheet tabs** (§36) under the grid: `$` and each grid of `$` drawn as
  a link, to go between them, and + to make another.

* **Outline** (side panel tab): every grid as a tree from the top grid `$`,
  with its name, size, how many cells hold something, errors (`⚠ 2`),
  cells failing validation (`✗ 1`, §16), cells to check (`◤ 3`, §26),
  cells of cycles that didn't settle (`⟳ 2`, §30) and cells with a
  comment (`◥ 2`, §33); folded nodes count the
  grids inside them too. Under each grid are its named cells, then the
  grids inside it. A named grid is listed once, by its name with its cell
  beside it (`basics B6`), and its contents are under it. Click a grid to
  show it in the main area, ⌖ to select the cell holding it, ⚠ to go to
  the next error, ✗ to the next invalid cell, ◥ to the next comment, a
  name to go to that cell.
  The viewed grid is highlighted, and the grids around the active cell
  shaded.
  * **The shapes a table is.** Beside a grid or name holding a table,
    a tag says which shapes it is (*a calendar*, *a route*), judged as a
    parameter taking one judges it: the shape's needed columns, every
    value of its column's kind. So a table that's the sort of thing you
    want tells you what to say a parameter takes (⚙ Parameters, beside
    the parameter: *takes a calendar*).
  * **Picking while you write a formula.** A click in the Outline
    always goes where you click, also while you're writing a formula (in
    a cell or the formula bar, once it starts with `(` or `=`, code in
    the REPL, or a field that takes a formula, like an instance's
    inputs): then it only *shows* that grid or name, and the formula
    stays open, so you see what's there before you pick it. Click the
    cell you want (or point at it with the arrows) and its reference
    goes in, by names (`model>rate`), so a formula reaching across
    grids, or a facade's links (§54), are picked rather than typed. The
    same goes for every other way of getting about: the breadcrumb, the
    sheet tabs, a double-click on a grid's link in a cell, and Alt-↓ into the grid the
    pointed cell holds and Alt-↑ out. (A box that takes typing, as
    Find's or the name box, takes the cursor: going there enters the
    formula.) The breadcrumb says *writing the formula of D1* while
    you're away from it, the Outline says so at its top, and the cell
    the arrows would point on from is outlined. Enter enters the formula
    and brings you back to its cell; Esc cancels it and brings you back
    too; clicking a cell where no reference can go selects that cell,
    entering the formula and staying there; so does pressing a row or
    column head when no reference can go at the cursor. Opening the
    Outline tab and unfolding its grids leave the cursor in the formula.
    The REPL picks the same way once something is typed in it.
* **Errors in the status bar**: `⚠ 2 errors` beside `✗ n invalid`
  counts the errors in the whole sheet, wherever you are, so stepping
  through them goes from one grid to the next. On one of them it reads
  `⚠ 1 of 2 errors`.
* **Everywhere | here**. Inside a grid, each of the status bar's
  counts comes in two halves: `⚠ 2 errors | 1 here`. The left counts the
  whole sheet; the right counts the grid in view and the grids inside
  it (`0 here` when there are none). Each half steps through its own:
  a click on the left goes to the next anywhere, on the right to the
  next in this grid, staying in it (Shift-click: the previous). On one
  of them, both say where you are (`⚠ 1 of 2 errors | 1 of 1 here`).
  While you step through "here", it stays the grid you started in,
  even when a step takes the view into a grid inside it or out
  to a grid's cell; moving anywhere yourself ends that. Stepping the
  everywhere half takes "here" with you: it is then the grid the step
  landed in.
  This holds for every count: errors, invalid cells, units warnings,
  cells to check, unsettled cycles, comments and failing tests. In the
  top grid, here is everywhere: both halves give the same number.
* **No address on the status row:** the active cell's address is in the
  name box, top left. The row keeps the selection's size, count, sum and
  average.
  * One broken cell turns every cell using it into an error too, so only
    the cells where an error **starts** are counted and gone to: `#PARSE`, `#REF`,
    `#NAME`, `#ERR`, `#CYCLE`, `#TIME`, `#SPILL`, `#UNIT` (§25), and Excel's error values
    (`#N/A`, `#DIV/0!` …, §17) made by the cell's own formula. A cell
    that shows an error it read — `#UP`, or an Excel error value from a
    cell it reads that holds an error — is **passed on**: fixing the
    source clears it. The count's tooltip says how many cells pass errors on.
  * When a grid has errors only passed on from outside it, its count
    steps through those instead, so a ⚠ is never a dead link.
  * The counts (status bar and Outline) are links: a click goes to the
    next error after the active cell, wrapping round; Shift-click the
    previous; zooming in or out as needed, as for invalid cells (§16).
* **Searching the outline** (its search box; Ctrl-G puts the cursor
  there). Typing narrows the tree to what matches, with the grids leading
  to each match kept and everything shown open:
  * names in any grid, fuzzily — a few letters in order: `prn` →
    `principal` (in `E32`); also by path, `E32>principal`;
  * grids, by address (`e15c1` → `E15>C1`) or by name; a named grid is
    found once, as the grid;
  * an address, path or range typed out — `B5>A1`, `..>C2`, `A1:B9` —
    read as a formula in the active cell would read it, and also from the
    top grid (as the name box reads paths), offered above the tree.

  The matched letters are marked. Enter goes to the best match (a typed
  address first): a grid opens in the main area; a cell or range is
  selected, zooming as needed. ↑ ↓ choose another, a click goes to any of
  them, Esc clears the search and a second Esc goes back to the grid.
* **Find and replace** (Find tab; Ctrl-F, Ctrl-H for replace):
  * search **formulas** as written, the **comments** written on cells
    (§33), or the **values** cells show — numbers
    as printed, errors by their code, so `#REF` finds every broken cell
    (and Excel's error values too: `#N/A`, `#DIV/0!`, §17); every error
    also reads with `#ERROR` after its code (`#REF #ERROR`), so `#ERROR`
    finds them all, passed-on ones included;
  * everywhere, or in the viewed grid and the grids inside it; match case;
    regular expressions (JavaScript; the replacement can use `$1`);
  * results in address order, depth first (`B5`, `B5>A1`, `B6`), with the
    matches marked. Enter goes to the next match after the active cell
    (Shift-Enter the previous), wrapping round; a click goes there,
    zooming out or in as needed;
  * **Replace** replaces in the active cell if it's a match and goes to the
    next one (Excel's Replace); **All** replaces every match as one undo
    step. Formulas and comments can be replaced, values can't: they're
    computed (a literal is its own formula, so `"apple"` is found and
    replaced in formulas).
    Replacing is plain text: it doesn't know about references, so
    replacing `B5` also changes `B50` unless a regex says otherwise
    (`\bB5\b`).
* **Trace arrows** ("⇠ Precedents" / "Dependents ⇢" in the breadcrumb
  bar): arrows from the cells the active cell uses (blue; a range boxed)
  and to the cells that use it (orange), following the selection.
  * They cross grid boundaries: an arrow reaches into a grid drawn in a
    cell, down to the inner cell.
  * A target inside a grid that isn't drawn (behind a link, or scrolled
    out of a `(grid-view rows cols)` window) gets a **dashed** arrow to
    the nearest drawn cell holding it, labelled with its path
    (`..>scheduler>tasks`). When that cell shows its grid as a
    `▦ grid ›` link, the arrow ends on the link and the label sits under it.
  * A target not on the page at all — outside the viewed grid when zoomed
    in, or far off in a big grid — is listed in a box beside the active
    cell (`from ..>B2:B5`). Every label is a link.
  * A label is the path a formula in the active cell would write, each
    grid on it and the cell by name where they have one
    (`..>scheduler>tasks`, not `..>B28>A0`). The `..>` steps stay: a
    name isn't seen from inside the grids within its own.
  * Ctrl-[ / Ctrl-] go to the first precedent / dependent, in whichever
    grid it is (as in Excel).
  * The grid cells holding a cell aren't drawn as its dependents, nor
    listed in the Inspector's *Used by*: their value contains it by
    structure, not by a reference. A grid cell's
    dependents include cells outside it that use something inside it.
  * A cell's dependents include the cells using a grid around it as a
    whole: a block given `{:tasks tasks}` reads every task, so a task's
    Work has the block's `tasks` parameter as a dependent, though no
    formula names that cell. The Inspector's *Used by* lists them *as
    part of tasks*.
  * What a cell's **renderer** reads counts too: a chart drawing
    `{:data prices}`, or a data bar's `(MAX B2:B7)`, has those as
    precedents, and they have it as a dependent. The Inspector says so:
    *Depends on* "its renderer reads prices", *Used by* "and the
    renderers of B3 draw from it". A format rule's renderer over a long
    column lists its first 2,000 cells.
  * Arrows run in the cells' margins, the strip inside each edge where
    no text is, so they never cover a value: down a column by its left
    edge, along a row under its top, and otherwise down one cell's column
    and along the other's row. Each starts at a dot in the margin of the
    cell it comes from and ends with its head on the edge of the cell it
    points at. Two arrows along the same cells take opposite edges
    (left and right, top and bottom), so each can be told apart.
  * A grid a formula makes (a steps table, `(steps actuals [...])`, or
    any `(as-grid ...)`) has the precedents its formula has: the
    `actuals` grid, and any cell a step reads. So does every cell in it:
    they hold what that formula made. A grid you fill in yourself
    doesn't trace what it holds (its own cells); its precedents are
    what its formula gives its parameters from outside: a block
    holding `{:resources resources}` has `resources` as a precedent
    (and a parameter so given has it too), and `resources` has the
    parameter as a dependent. The Inspector's *Depends on* says so,
    and selecting the block outlines `resources` (dashed), as
    selecting a formula outlines what it reads.
    Zoomed into the block with the grid itself selected, they're
    listed under its ▦ corner.

### Back and Forward

The browser's Back and Forward (Alt-← / Alt-→, a mouse's side buttons)
retrace jumps: zooming in and out (the ▦ link, Alt-↓ / Alt-↑, the
breadcrumb, the Outline) and every jump to a cell —
`[[links]]` in notes, the name box, the Outline and Find, stepping through
errors and invalid cells, Ctrl-[ / Ctrl-]. Each is a history entry
holding the view and the selection. Moving the selection by
clicks and arrow keys isn't a jump: it rewrites the current entry,
so Back returns to the cell you jumped
*from*, and Forward to where you left off after jumping. The URL's hash is
the active cell (`#B6>C9`): a URL with one opens the sheet there, and a
hash typed into the address bar goes there.

**An expanded side panel steps aside.** With the side panel at the
window's width (⇤ Expand), the grid is out of sight, so going to a cell
puts the panel back beside it: a link to a cell in any tab (the Tests
tab's test and ↗, the Inspector's, Profile's, the Map's *Go to the
cell*), a status-bar count, the Outline, and a jump from the name box.
Opening a grid in the Map, or looking through a table in the Data tab,
keeps it expanded: those show their result in the panel.

### Help: a guide, and a search

The side panel's **Help** tab has a search box at its top, then *How
do I…*, then Help's reference.

* **How do I…** lists the common tasks: type a table, make a range a
  table, sort and filter, add a computed column, total a column,
  summarise, make a chart, say what a column holds, look up a value,
  schedule a plan, fix an error. Click one to unfold a few lines on
  how, with two links: *In Help*, to the part of Help's reference that
  says more, and *In the manual*, to this manual's section.
* **The reference** follows, a heading per topic, with a line of links
  to each.

* **Search.** Type into the box at the top (F1 from anywhere puts the
  cursor there) and Help shows, as you type, what matches: the *How do
  I…* entries, Help's paragraphs, and this manual's headings with their
  first lines. Each word typed must be there; a word finds those it
  starts (*sort* finds sorting), and plurals find the singular
  (*charts*). `§20`, or `20`, finds the manual's section 20. The best
  matches come first, the *How do I…* entries before the rest.
* ↓ and ↑ choose a result, Enter opens it (the first while none is
  chosen), Esc empties the box and shows the guide again; a click opens
  one too.
* A Help result scrolls the reference to its paragraph, flashes it, and
  marks the words you typed in it. A manual result shows that section
  of the manual in the tab, with its subsections; *‹ Results* goes back.
  A `§` number in Help's text, or in a section shown, opens that
  section.
* The manual is loaded the first time it's wanted (the box focused, a
  *How do I…* entry unfolded, a manual link clicked). Until it's there, Help's own paragraphs are
  searched, and the results say so; if it can't be loaded, they say why,
  with *Try again*.

## 12. Help while writing formulas

The formula bar helps the way Excel's does: coloured references with
matching outlines on the grid, completion, F4 for `$`, and values on hover.

* **Coloured references.** Each distinct reference in the formula gets a
  colour (by its text, in order of first appearance; eight colours, then
  they repeat). While the formula bar has focus, the cells each one points
  at are outlined in its colour, one rectangle per range, in any grid on
  screen (inline grids too). The committed formula's dashed precedent
  outlines step aside meanwhile. Clicked references (point mode) are
  outlined as soon as they're inserted.
* **Completion** pops up while typing a symbol:
  * library functions, the host builtins and Excel's functions (with
    arglists, and the docstring alongside), then `clojure.core`;
  * cell names in scope (with address and value), ranked first;
  * grid cells of this grid as `E15>`, and `..>` / `$>`;
  * after a path (`B5>`, `..>`, `$>`, `B5>C1>`): the names and non-empty
    cells of the grid it leads to, with their values; grid cells continue
    with `>`;
  * after the `:` of a range, the part after the colon.

  **Tab** takes the highlighted option, or the first. Nothing is
  highlighted until you arrow to it, so **Enter** still commits the formula
  (typing `tax-rate` Enter commits, it doesn't stop at the popup); Escape
  closes the popup, a second Escape cancels the edit. Nothing pops up
  inside strings and comments. The renderer and library editors complete
  functions only, since they can't refer to cells. The Renderer field
  narrows further: its first symbol (`plain`, `(data-bar`) completes only
  renderers and renderer makers, the vars marked `^:renderer` (every one
  in the standard library is; a library's own are when written
  `(defn ^:renderer f …)`). Inside arguments (`(render-if neg? …)`,
  `(pie greens)`) it offers every function, renderers first.
* **F4** cycles the reference at (or just before) the cursor through
  `B5` → `$B$5` → `B$5` → `$B5` → relative `u2l1` → `B5`, as in Excel, plus
  the relative step. A relative reference is worked out to the cell it
  reaches, so `re` becomes the last column. In a range both ends change
  together; in a path only the coordinate the reference starts from (`B5`
  in `B5>A1`, `D5` in `..>D5`); names and `$>` paths are left alone.
* **Brackets by depth.** In the formula bar and the other code editors
  each bracket pair is coloured by how deeply it is nested (six colours,
  then they repeat; one set for light and one for dark), so the pairs of
  a long one-line formula can be told apart. A bracket without a partner
  (unclosed, a stray closer, `(]`) shows red. While the editor has focus,
  the bracket at the cursor and its partner are also boxed. Brackets
  inside strings, comments and character literals are left alone, and so
  is text in text mode.
* **Hover** a reference for where it points and its value (a range as rows
  of values, at most 50 × 20; an error as its code and message). Hover a
  function for its arglists and docstring, in the formula, renderer and
  library editors.
* **A long formula is read in the formula bar.** The Inspector's
  *Formula* row shows a formula that fits on one line as it is. A longer
  one (several lines, or one line wider than the panel) is a link, *10
  lines: edit in the formula bar ↑*: click it and the bar opens on it,
  laid out and coloured. In a protected cell the bar can't be typed
  into, so the link unfolds the formula in the Inspector instead, as
  written, with the way to edit it (open the lock 🔒 above the grid).

### What typing enters: text unless it says otherwise

Typing into a cell enters **text**, as Excel does, unless what's typed
says it's something else:

* **A value stays a value:** a number (`42`, `-5`, `1.5e3`), a date or a
  time (`2026-11-07`, `14:30`), a quantity with a unit (`5 km`), TRUE or
  FALSE in any case, and Clojure's literals (`:done`, `[1 2 3]`, `{:a
  1}`, `nil`, `"text in quotes"`).
* **A percentage is its number**, as in Excel: `50%` is 0.5, `12.5%`
  0.125. A cell that shows no number format gets one that shows it as
  typed (`0%`, `0.0%`); a column already formatted as a percentage
  shows it so. Pasted into a column of numbers, or in a column you make
  hold numbers, `50%` is read the same way. In a column of text or IDs
  it stays text.
* **`(` starts a formula**, as always: `(* B2 C2)`, `(SUM ub:u1)`.
* **`=` starts a formula too**, as in Excel: `=SUM(B2:B5)` is
  translated, and `=` before a reference or a name is that reference:
  `=B5`, `=tax-rate`, `=settings>status`. (Written bare, `tax-rate` is
  the text *tax-rate*.)
* `[`, `{` and `#` start a formula too: a vector or map with references
  in it, `#(* % 2)`.
* **A path is a reference**, with or without `=`: `$>loading`,
  `..>rate`, `plan>finish`, when it leads to a cell or range (it can
  only be a reference). A bare `rate` or `B5` is text: write `=rate`.
* **Anything else is text:** `Ben`, `User interviews`, `2, 3`, `12-2d`.
* **A leading `'` enters the rest as text, whatever it is**, as Excel's
  does: `'(draft)`, `'=not a formula`, `'00123`.
* In a table's column of text (§19, *Column kinds*), a number typed is
  text too: `007` stays `007`.

The formula bar's **ƒ / T** mark shows which it will be while you type:
**T** for text, **ƒ** for a formula or a value. Click it to enter what
you typed the other way. Completion, coloured references and clicking
cells to insert references start once it's a formula (after `(` or
`=`), so `Ana` Enter enters *Ana*, not a function that starts with it.
A cell already holding a formula is edited as a formula, and one holding
text as text: only what's typed afresh is read this way.

### Text cells: writing text without quotes

A cell's source is one Clojure form, so text is a string literal, with
quotes and escapes, which is clumsy for prose (notes, Markdown). The
formula bar therefore edits a cell whose source is nothing but one string
literal **as text**: the bar shows the text
itself, marked **T** where it shows ƒ for formulas, with no Clojure
highlighting or formula help; Enter commits, as in a formula,
Shift-Enter or Alt-Enter (Excel's) starts a new line, Ctrl-Enter commits
and stays (Excel's Ctrl-Enter), Esc cancels, and leaving the bar commits
as usual. The commit writes the text as a string literal (empty
text empties the cell), so the stored form doesn't change: files, copies,
formulas reading the cell and the Inspector see the same string literal.
Ctrl-; puts the date in bare, and clicking cells doesn't insert
references (point mode is for formulas).

The ƒ/T button switches the cell being edited, converting the draft:
text → formula gives its string literal, formula → text gives the
literal's text, or a formula's source as text. The choice lasts until
the edit ends; otherwise the mode follows the source. Typing over a
cell, or into an empty one, enters what's typed by the rule above:
text, unless it's a value or starts as a formula does.

Until it's focused, each formula-bar editor shows only its first line,
unwrapped, at a fixed height, with ⋯ when there's more (more lines, or a
line too long): selecting a cell with a long note or formula doesn't push
the grid down. Focusing it (Enter, F2, typing, a click in
it) opens it to its content.

### Typing into cells: Enter and Edit modes

What you type shows **in the cell** you're typing into, as well as in
the formula bar: the cell draws the draft over its content, with a
caret where the cursor is, running on over its neighbours when it's
long, so your eyes stay where the value is going. The completion popup
and the other formula help are still in the bar.

How the keys behave depends on how the edit started, as in Excel:

* **Enter mode**: you typed over the cell, or opened it with F2, Enter
  or a double-click. The arrows finish the entry and move that way, so
  a row fills as fast as a column does: `1` → `2` → `3` →. In a formula,
  where a reference can go (after `(`, `=`, a space, an operator in an
  `=` formula), they point at cells instead and put their reference in
  (§5, *Point mode*).
* **Edit mode**: you clicked into the formula bar (also while you're
  already writing), or switched ƒ/T. The arrows move the cursor in what
  you're writing.
* **F2** switches between the two while you write. The status bar shows
  which you're in (**Enter** or **Edit**, at its right end) for as long as
  the edit lasts.
* **Tab** / **Shift-Tab** finish the entry and move right / left, in
  either mode (when the completion popup is open, Tab still takes its
  option). **Enter** after a run of Tabs goes down to the column the run
  started from, so a table fills row by row: `1` Tab `2` Tab `3` Enter
  lands under the `1`. Tab also moves this way when you aren't typing.
* Halfway through a formula, while a bracket or a string is still open
  (`(SUM `, `"Tot`), an arrow moves the cursor even in Enter mode, rather than committing
  something that can't be read. When the completion popup is open, ↑
  and ↓ move through its options.
* Enter commits and moves down, Esc cancels and puts the cell back, and
  a refused entry (failing validation, §16) stays open to be fixed,
  whichever key committed it.

### At a grid's edge

Typing never lands on a cell that already holds something without your
asking. At a grid's last column or last row, the keys go on, and the
grid grows to make room, as Excel's sheet goes on:

* **Tab** at the last column adds a column and moves into it, whether
  or not you typed. In a table you're zoomed into, Tab goes to the faint
  empty column drawn after the last (§19), and typing there adds it.
* **Enter** after typing in the last row adds a row and moves down
  into it (after a run of Tabs, to the column the run started from), and
  so does **→** or **↓** in Enter mode. A table already keeps an empty
  row last, so Enter there just moves into it.
* The value and the new row or column are **one undo step**: Undo takes
  both back.
* **Paste** a block that runs past the edge, and the grid grows to take
  all of it, in one undo step (a table keeps an empty row after it).
  Nothing is cut off: where the grid can't grow, nothing is pasted, and
  the message says the block's size, the grid's, and why.
* This works in the top grid too (the sheet), in any grid in a cell,
  zoomed into or drawn in its cell, headed or not. A grid grows at its
  right and bottom only.

**Where a grid can't grow**, the cell stays selected and the message,
by the cell, says why and what to do:

* a protected grid (§47), while the lock is closed: open the lock 🔒 to
  add rows or columns;
* a grid a formula makes (§22), a view or an instance: it's as big as
  its result;
* a grid grown to fit its results (§38): past its own size its cells
  hold results only, so make it bigger with Resize;
* a grid drawn in its cell showing only its display area (§6): zoom in
  to work past it;
* Shift-Tab at column A, or ← and ↑ in Enter mode at the first column
  or row;
* Enter (or ↓ in Enter mode) from a table's totals row (§19): records go above
  it.

Then, while that cell holds something, typing over it is refused too,
saying so, so a quick `x` Tab `y` can't replace `x` with `y`. Press F2
to change it, Delete to clear it, or move to another cell.

The arrow keys without an edit stop at a grid's edge and add nothing.

### When an entry is refused

Some entries are refused, by the cell, saying why and what to type: a
date that could be either (`3/1/2026` could be 3 January or March 1), a
value its column's check refuses, an Excel formula that can't be
translated. The edit stays open with what you typed, so you can fix it.

* **Keys typed after it are kept out of it.** Typing on quickly, as
  into a row (`3/1/2026`, Tab, `10`, Tab …), the characters meant for
  the next cells aren't added to the refused entry: the message, by the
  cell, says what was kept out (`"10"`), to type again once it's fixed.
  Tab and Enter try the entry again, and are refused again.
* **To fix it**, press Backspace, Delete, F2 or an arrow key in Edit
  mode, or click in the formula bar: from then on, typing goes into it
  as usual. **Esc** leaves the cell as it was.
* **A fix offered.** An Excel formula left open, `=SUM(A1:A3`, is
  offered closed, as Excel does: the message's button *Close it:
  =SUM(A1:A3)* enters it, and so does **Enter** while the entry is still
  as you typed it. It's offered only when the closed formula reads;
  `=SUM(A1:A3,` isn't finished by closing it.

## 13. Spilling: results that fill a range

A formula whose result is a table spills it into the sheet: each element
lands in its own cell, below and to the right of the formula, as Excel's
dynamic arrays do. The table stays live — edit an input and the spilled
cells follow — and the cells it fills read like any others.

```clojure
;; in B38: a 5×3 table, filling B38:D42
(vec (for [n (range 1 6)] [n (* n n) (* n n n)]))
(SUM C38:C42)          ; elsewhere: 55, read from the spilled cells
(for [r rows] [(:name r) (:qty r)])   ; any sequence of rows spills
```

* **What spills:** a non-empty vector or sequence whose elements are all
  vectors or sequences (rows), at least one of them non-empty. Ragged rows
  leave the missing cells empty. A sequence of maps (a table, §19) spills
  too, under a header row of its keys. A flat vector (`[1 2 3]`), a map
  and a number don't spill. Neither does a range written on its own with a
  renderer (`B2:B5` shown as `bars`): **a cell with a renderer of its own
  keeps its result in place**, so `table` is the way to show a table inside
  one cell. A renderer from a column, row or range format (§10) doesn't
  count: those format the cells a result lands in. To lay a result out in
  a grid of its own instead, as big as the result, wrap it in `(as-grid
  …)`, or draw the cell with `(grid-view)`, which lays out a table or
  rows given it the same way (§22).
* **The formula's cell** (the anchor) holds the whole table as its value
  — `(count B38)` is 5 — and shows its top-left element. A **range** over
  the anchor reads what the cells show, so `(SUM B38:B42)` is 15.
* **The spilled cells** hold the elements: `C39` is 4. They have no formula
  of their own; selecting one shows the anchor's formula in the formula bar,
  greyed, and the Inspector says where it came from. A thin frame marks the
  area.
* **`#SPILL`** on the anchor: the table needs cells that hold a formula or a
  grid (formats and renderers don't count), cells another table already
  fills, or more room than the top grid has (a grid in a cell grows to fit,
  §38). Nothing is spilled partially. The message names the cell in the
  way; clear it, or grow the grid, and the table spills. Typing into a spilled cell is allowed and blocks the spill
  (as in Excel), so the value typed is never silently overwritten.
  Or **draw it inside its cell**: a cell with a renderer of its own keeps
  its result there, needing no room. `(grid-view)` lays it out in a grid
  of the cell's own (§22), whose columns you can widen and format; `table`
  draws it as a read-only table, its columns of dates (as the function or
  step that made it says: `bookings`' From and To, a schedule's start and
  finish) as dates, `2027-01-04`, or in the cell's own format when it has
  one. The message says so, and a `#SPILL` cell is offered the way
  out where you'd look: the **Shelf tab** opens on a section *As this
  cell's own renderer, out of its #SPILL*, whose tile, the result drawn
  by `(grid-view)`, gives the cell that renderer in a click; the
  **renderer gallery** (the Renderer field, or right-click → *Show as…*)
  draws the held-up result by every renderer that suits it, `grid-view`
  first, and says why. Formulas reading the cell still get its error
  until it has one.
* **`#CYCLE`** on the anchor: the table would spill into a cell the formula
  itself depends on.
* **Materialise** (right-click a spill): *Convert spill to values* replaces
  the formula with the values it laid out, one literal per cell; *Make
  table a sub-grid* turns the cell into a grid of literal cells holding the
  table (also for a table kept in its cell by `table`), ready to edit. Both
  are one undoable step. Paste values over a spill does the same as
  convert, cell by cell, since copying reads what cells show.
* Spills work inside grids too, and in a `call`: the called grid's value
  map has the spilled cells (and the formula's cell as it shows, its
  top-left; `(call B2 {…} :A1)` gives the whole result), a pinned input can
  make its tables grow, and a pinned 2-D value spills like a computed one.

## 14. Libraries: standard, My library, and per grid

The functions formulas and renderers use come from libraries. The
**standard library** — formula helpers, dates, every predefined renderer —
is part of the app, the same for every sheet and grid, and isn't saved with
sheets. **My library** is the user's own toolkit: what they find generally
useful, seen by every sheet they open, kept in the browser rather than in
sheets. **Any grid can also have a library of its own**, saved with it: a
pricing grid defines `discount` without cluttering the sheet, and the grid
can be copied, saved or inserted elsewhere with its functions.

A function that proves generally useful moves from a grid's library to My
library, instead of being folded into the app or copied into every sheet;
a standard function to adapt is copied, from its source, into either.

* **Library tab.** It shows the library of the active cell's grid; the
  "Library" menu switches to a grid around it (`$` is the top grid) or,
  last, to My library. A library is shown as documentation, like the
  standard library: its definitions by category, each with its arguments
  and docstring, under the same filter, rather than as one long text, so
  a large toolkit stays easy to find one's way in.
  * *Edit* beside a definition opens just that definition (with the
    comment lines above it) in a small editor in place; Apply
    (Ctrl/Cmd-Enter) puts it back where it was, Cancel (Esc) leaves it,
    Delete removes it.
  * *+ Add function* writes a new one at the end; *+ Add function here*,
    under a category, writes one at the end of that category.
  * *Edit all* opens the whole library as text, for categories, order or
    several definitions at once. An edit there that isn't applied is kept
    while you look at other grids (Revert drops it), and Edit / Add wait
    for it to be applied or reverted.
  * Remove deletes a grid's library. Every change is undoable.
* **Categories.** `;;; # Title` lines divide a library into categories, as
  they divide the standard library into sections; the `;;;` lines right
  after a heading introduce it. Definitions before the first heading are
  the library's own section. A definition without a docstring is
  documented by the `;;` comment lines directly above it, which the
  filter searches too.
* **My library** is between the standard library and the grids': it can
  redefine a standard name (`SUM`, `default-render`, `time-limit-ms`) for
  every sheet, and any grid's library can redefine its names in turn. It's
  kept in the browser (localStorage, a key of its own), so every sheet,
  New, the examples and Open included, gets it; Download saves it as a `.clj`
  file, and Load file puts a file's text in its editor to Apply. To keep
  it on disk, or take it to another browser, with the shelf, use the
  library file (§4, *The library file*). Edits are
  undoable like any other change, in the sheet's undo history, which has
  a consequence: undoing past an edit of My library takes My library back
  too, for every sheet. Another tab's change to it is picked up.
* **Files carry what they use of it.** Save (and Save grid) writes into the
  file, as `:my-library`, the definitions of My library the sheet uses:
  those whose names appear anywhere in its cells, renderers, format rules
  or libraries, `default-render`, `time-limit-ms` and `iterate-cycles`
  (the engine looks them up itself), and those they use in turn. Opening or inserting the file
  where My library lacks some of them puts those in the top (or inserted)
  grid's library, under a comment saying where they came from, and says
  so; names My library has keep its definitions. A sheet saved here works
  anywhere, without copying the toolkit into every sheet. The autosave in
  the browser doesn't carry it (it's the same browser).
* **Scope, as with names.** A grid's library is visible to the formulas,
  renderers and format rules of that grid and of every grid inside it, not
  outside it or in sibling grids. Inside, the nearest definition wins: a
  grid can redefine a standard function (`SUM`, `default-render`,
  `time-limit-ms` in the top grid) or an outer grid's, for itself and the
  grids inside it. The standard library's own functions keep using the
  standard definitions (Clojure's lexical scope): redefining `plain` doesn't
  change what `(data-bar 25)` falls back to. `default-render`,
  `time-limit-ms` and `iterate-cycles` (§30) are the exceptions, because
  the engine looks them up in the grid's namespace.
* **One set of functions.** The standard library doesn't duplicate
  Excel's: there's no `sum`, `avg`, `minimum`, `maximum`, `round-to`, `transpose`,
  lowercase maths builtins (`sqrt`, `abs`, `round`, `floor`, `log`, `sin`
  …), `today`, `now`, `days-between`, `hours-between` or `add-days`;
  `SUM`, `AVERAGE`, `MIN`, `MAX`, `ROUND`, `SQRT`, `TODAY` … are
  the functions (§17). What's left is what Excel has no word for: the
  helpers `cells`, `nums`, `column`, `clamp`, `add-time`, the fills, the
  validation predicates, the table helpers (`sort-rows`, `group-rows` …,
  §20), the renderers.
* **Reference.** Under the editor: everything the grid's formulas can use
  besides `clojure.core`, nearest library first, then My library, then
  Excel's functions by
  section (*Operators*, *Logic*, *Maths*, *Statistics*, *Lookup and
  arrays*, *Text*, *Dates and times*, *Financial*, *Information*, *Number
  formats*), then the standard library's other sections (*Dates and times*
  is one section for both, PI sits in *Maths*), each name with its
  arguments and docstring. Names redefined
  nearer are marked. A filter searches names and descriptions (and, with
  Show code on, the code). Hovering a
  name in any editor shows the same text, and completion offers it.
* **Show code.** The Reference is documentation; ticking *Show code* shows
  each definition's source under its docstring (Excel's functions and the
  host functions are ClojureScript, and say so), with *Adapt in …* (the
  library shown: a copy opens in a new definition's editor to be changed
  and applied; keeping the name replaces the original for that library's
  scope, another name keeps both. From the standard library it's the
  definition alone: the helpers it calls stay the standard ones, named in
  the editor's note; from another library its private helpers come too)
  and,
  for a grid's definitions, *Move to My library* (one undoable step; its
  comment lines go with it).
* **Copies of standard code.** A library shouldn't hold copies of the
  standard library's code: a copy freezes it, and the standard one may
  change. Save grid, a template put on the shelf, and Insert leave out
  copies the same as the standard library's. A copy of a standard
  helper that differs (an old copy) is flagged: *◤ 1 old copy* in the status bar opens the Library
  tab, which names it, with *Drop the old copies* to use the current
  ones (give one you changed on purpose a name of its own instead). A
  template on the shelf carrying one says so on its card, with *Use the
  current one*.
* **Every standard function and renderer has a docstring** saying what it
  does and how to write it (`plain` bare; `(data-bar 25)` with options).
* **Copying and files.** A grid's library is part of the grid: copy/paste,
  fills and inserting rows carry it; Save grid writes its library, with
  what it uses of the libraries of the grids around it put first (as a
  file carries what it uses of My library: the definitions it names, those
  the engine looks up itself, `default-render`, `custom-units` …, and
  those they use; a library running code, not just defining, goes whole),
  as the file's top library; Insert makes a file's top library the new
  grid's library. The grid cell's Inspector lists what it *Borrows* (§6).
* **Inserting where they're defined already.** Put back where it came
  from (or anywhere the grids around it, or My library, define the same),
  a saved grid's copies would shadow those definitions, and a later fix
  to them wouldn't reach it. So Insert leaves out each definition that
  means the same there: the one the grid would inherit has the same text
  (comments and spacing aside), and so does everything it uses, as each
  side finds it (a nearer grid redefining a helper keeps the copy that
  uses the helper as saved), and private helpers only those used go with
  them. A `;; ----` heading stays while something is left under it. What
  differs stays, as an override; a library running code is kept whole,
  and a library around it that failed to load (or runs code) stands in
  for nothing. Insert's message names what was left out.
* **Errors in libraries.** A library that fails to load says on the
  Library tab which line: `line 2: No item 1 in vector of length 0` (the
  definitions before it work). A cell whose formula fails inside a library
  function names the function and its line (§3, *Error codes*), and the
  formula bar underlines the cell's own call into it. A failing call
  that isn't a function's own, such as `(.toUpperCase v)`, gives the
  line the function starts on. A format rule failing in a library
  function says so on the Formats panel the same way.
* **Checks in libraries.** The Library tab lists, under *Check*, calls in
  the library's own code with argument counts their functions don't take
  ("line 3: tax takes (tax x), not 2 arguments"). The library still
  loads: they're for you to look at.
* **A broken library in the status bar.** While a library fails to load,
  the status bar shows `⚠ library` (`⚠ 2 libraries`): its tooltip says
  which and why, and a click opens the Library tab on it. A formula using
  a name it would have defined says so (§3, *Did you mean*).

## 15. Dates, times and fills

Two things every sheet needs: entering a
date or time, and a column or row of values — numbers going down, months
across — that runs to the edge of the grid without anyone counting rows.

* **Dates are numbers, as in Excel.** A date is a serial number — days
  since 30 December 1899, so 2026-11-07 is 46333 — and the time of day is
  the fraction (noon is .5, 14:30 on its own 0.604…). What makes a number
  look like a date is its cell's number format (§10): `yyyy-mm-dd`,
  `d mmm yyyy`, `h:mm AM/PM`. So `(+ A1 7)` is a week later, `(- B1 A1)`
  the days between, `(* 24 (- B1 A1))` the hours, and Excel's date
  functions (`YEAR`, `EDATE`, `NETWORKDAYS`, `TEXT` …) take the same
  numbers as RecDep's. The Inspector's *Value* shows a date as a date (in
  the cell's format, else as `2027-04-01`), and the number under it, in
  grey (*the number 46478*), for when you want it.
* **A cell holds dates** when its kind says so (*Holds ▸ dates*, §19),
  or, with no kind, when its number format shows a date (`d mmm yyyy`,
  not `h:mm`: hours add up). That is what the Shelf, the charts, the
  views and the Inspector go by: a column of dates is never summed or
  charted as a series of numbers (§ *Functions dropped on a range*, §20).
* **Typing dates and times.** Write them as you would, without quotes:
  `2026-11-07`, `2026-1-5`, `14:30`, `9:05:30`, `2:30 pm`,
  `2026-11-07 14:30` (or `T14:30`). The cell holds the serial number and,
  unless it already has a date format, gets one as written: `yyyy-mm-dd`,
  `yyyy-mm-dd hh:mm`, `hh:mm`, `hh:mm:ss`, `h:mm AM/PM`. The source stays
  as typed. A shape that isn't a real date or time (`2026-02-30`, `25:00`,
  `13:00 pm`) is a `#PARSE` error that says why. Inside a formula a date
  is written as text, `(- A2 "2026-01-01")`: the operators and the date
  functions read date text as the date it names. Pasting date text from
  another app, and a picker writing one, enter it as if typed.
* **Dates typed as the column shows them.** Typing also reads a date
  the way its cell (or its column) shows dates: in a column formatted
  `d mmm yyyy`, `10 Mar 2027` is 10 March 2027; in one formatted
  `dd/mm/yyyy`, `3/4/2027` is 3 April. Anywhere, a month's name reads:
  `10 Mar 2027`, `10-Mar-2027`, `10 March 2027`, `Mar 10 2027`,
  `March 10, 2027`, `Wed 10 Mar 2027`, `10th March 2027`. Numbers alone,
  `13/3/2027` or `3/13/2027`, read when only one order is a date; when
  both are (`3/4/2027`), and the column's format doesn't say which comes
  first, it's refused, by the cell, saying both readings and what to
  type: give the column a format with the day first (`dd/mm/yyyy`) or
  the month first, or type `2027-04-03` or `3 Apr 2027`. The year needs
  four digits (two when the column's format has `yy`); a date without a
  year (`10 Mar`) stays text. The cell holds the date as `2027-03-10`,
  shown in the column's format, or, in a cell without one, in the shape
  typed (`d mmm yyyy`, `mmmm d, yyyy`). A column of text keeps what's
  typed as text. Written as a formula (the formula bar's ƒ), such a
  shape is an error that says it's a date and how to write one in a
  formula, `(DATE 2027 3 10)`. Text already in a column of dates (typed
  with a leading `'`, or from before) is read again as a date when you
  edit it and press Enter, and date text pasted into a column of dates
  (from Excel, say) is read the same way; elsewhere a paste reads
  `2027-03-10` as a date and leaves other shapes as text.
* **A calendar's From and To take dates, not text.** A holiday typed as
  text (`'10 Mar 2027`, or text a formula writes) is flagged in its date
  column and refused by the calendar too, naming the rule: type it again
  as a date. A date given to a function as an argument,
  `(week-start "2027-01-20")`, still reads.
* **Formats from the entry.** As in Excel, a cell without a number format
  gets one from what's entered into it: a formula
  whose outermost function makes a date (`TODAY`, `DATE`, `DATEVALUE`,
  `EDATE`, `EOMONTH`, `WORKDAY`) gets a date format, `NOW` a date and time,
  `TIME` and `TIMEVALUE` a time; `add-time`, `series`, `down` and `across`
  from a date text get that date's; any other formula takes the format of
  the first cell it refers to (a range's first cell), so `(+ A1 7)` on a
  date is a date and `(* B2 2)` on dollars is dollars. It's set once, at
  entry, and shows in the format bar, where it can be changed; a cell that
  has a format, of its own or by a rule, keeps it — except that a typed
  date or time replaces a format that isn't a date format, as Excel does.
* **Ctrl-;** enters today's date and **Ctrl-Shift-;** the time now (as in
  Excel): into the active cell from the grid (one undo step), or at the
  cursor in an editor — bare when it's the whole formula, in quotes
  inside a larger one.
* **Showing.** A date format shows them; without one a date is just its
  number, and every renderer writes it in that format (`(kpi "Due")`
  on a cell with `dddd, mmmm d, yyyy` shows the weekday).
  `date-input`, `time-input` and `date-time-input` are pickers; they read
  the serial number and write what they pick as if typed. A day picked
  from the calendar is written at once; a date typed into the box is
  written when you press Enter or leave the box (Escape puts it back),
  so a date half typed is never written.
* **Weeks.** `(week-start d)` is the Monday of date d's week (`(week-start
  d :sunday)`, weeks from Sunday; any day of the week by name): what a
  column of days is grouped by to count weeks, as in a table step's
  worked-out column `#(week-start (:day %))`. Given a working calendar
  (§51), `(week-start d cal)` starts weeks the day after its weekly days
  off: Sunday for a Sunday-to-Thursday week.
* **Arithmetic.** Plain `+` and `-` for days; `add-time` moves a date or
  time by a number of `:seconds :minutes :hours :days :weeks :months
  :quarters :years`, as a serial number: `(add-time "2026-01-31" 1
  :months)` is 2026-02-28 (the day is clamped, the time of day kept), and
  `(add-time "23:30" 45 :minutes)` is 00:15 the next day (an `hh:mm` format
  shows 00:15, as in Excel). `TODAY` and `NOW` read the clock when the
  cell is computed — on load, when it or its inputs change, or on F9 —
  not continuously; `(tick 1000)` is a clock that runs (§28). `between`
  takes dates as bounds (`(between "2026-01-01" "2026-12-31")`), for
  validation (§16).
* **Filling.** `(down …)` lays values out from its cell to the bottom of
  the grid, `(across …)` to the right-hand edge; the result spills (§13).
  With no argument they count 1, 2, 3 …; with a start they carry on from
  it as Excel's fill handle does:

  | formula | fills |
  |---|---|
  | `(down)`, `(down 0)`, `(down 10 10)` | 1 2 3 …, 0 1 2 …, 10 20 30 … |
  | `(across "Jan")`, `(across "Jan" 3)` | Jan Feb Mar …, Jan Apr Jul Oct … |
  | `(across "Mon")`, `(down "friday")` | weekdays, in the case written |
  | `(across "Q1")`, `(down "Week 1")`, `(down "Lot 007")` | Q1–Q4 round again; text ending in a number counts on, keeping zeros |
  | `(down "2026-01-05" 7)`, `(down "2026-01-31" :months)`, `(down A1 1 :months)` | dates by days, or by a unit (a number with a unit is a date) |
  | `(down "09:00" 30 :minutes)` | times by hours unless a unit is given |
  | `(down (range 28))`, `(across ["Name" "Qty"])` | a list, as it is |

  Other values repeat. A list stops where it ends; anything else, or an
  endless sequence such as `(range)`, stops at the grid's edge. Grow the
  grid (resize, or insert rows) and the fill grows too. Something in the
  way makes it `#SPILL` like any spill. Like every spill, Convert spill to
  values turns it into literal cells.
* **The pieces.** `(series start step unit)` is the endless sequence
  behind `down` and `across`, for formulas of their own:
  `(take 12 (series "Jan"))`. `(here)` tells a formula where it is —
  `{:row 3 :col 2 :rows 30 :cols 10 :address "B3"}` (counted from 1) —
  and `rows-to-end` / `cols-to-end` count the cells left to the edge,
  its own included.
* **Not the Fill buttons.** Fill down / Fill right (Ctrl-D / Ctrl-R) copy the
  active cell over a selection, shifting references (§5). `down` and
  `across` are formulas: one cell holds the whole fill and it stays live.
  That's why they aren't called `fill-down`.

## 16. Validation and input constraints

A cost should stay under budget, a name short enough to print, a finish
after its start. A validation rule says so, beside the formats (§10):
cells whose values fail it are flagged, typed input that fails it can be
refused, and the status bar, the Outline, the Rules tab and Find count
the cells that fail and go to them.

A validation rule only checks. What cells *hold* (whole numbers, a date,
one of the sizes listed in D2:D9, one or more of the tasks' IDs) is a
**kind** (§19, *Holds ▸* on the cell menu): a kind checks its cells the
same way, and also says how a value goes in, the ▾ that picks one, and
what a bare drop-down or slider offers. Use a kind for what a cell
holds, a rule for a test of a value that may still be entered.

* **Rules.** A validation rule goes in the grid's `:formats`, with an area
  like any rule; the Rules tab's "+ Validation" adds one for the
  selection to start from.

  ```clojure
  [{:area "B2:B20" :valid "pos-int?" :msg "whole units"}
   {:area "C2:C20" :valid "(between 0 1000)" :msg "over budget" :refuse true}
   {:area "E:E" :valid "(all-of int? (between 1 10))"}
   {:area "F:F" :valid "#(> % l1)" :msg "must end after it starts"}
   {:area "A2:A50" :valid "#(<= (count %) 24)" :msg "a long name"
    :style {:background-color "#fff1c2"}}]
  ```

  * `:valid` — source for a predicate on the value: any function of it
    (`pos-int?`, `string?`, `#(< 0 % 100)`, a set `#{:todo :done}`, a
    library function), or one the standard library makes: `(one-of …)`,
    `(between lo hi)`, `(all-of p q …)`. References in it are read from
    each cell checked, as a formula in that cell would read them:
    `(one-of D2:D9)` is D2:D9 for every cell, `l1` the cell's left
    neighbour, a name the grid's named cell. A reference with a `$` is
    read as Excel reads it, the rule written for its area's first cell:
    over C3:C20, `#(<= % $B3)` checks each cell against column B of its
    own row, and `#(<= % C$2)` against row 2 of its column.
  * `:msg` — what a failing cell says. Without it, the predicate's own
    message (`one-of`, `between` and the standard predicates have one:
    "must be a whole number above 0"), else "must satisfy `<source>`".
  * `:refuse true` — refuse typed input that fails. Without it, a failing
    value is entered and flagged.
  * `:style` — a CSS map for the cells that fail this rule, instead of the
    magenta edge and tint: amber for a warning, say, red for a hard limit.
* **Flagged.** A cell whose value fails a rule gets a magenta edge, a
  light magenta tint and a magenta corner (top right), or, if the rule it fails has a
  `:style`, that style and the corner. The rule's style wins over every
  other style of the cell, its own included: being invalid is the thing
  to see. Hovering the cell, or the Inspector's Validation row, says why.
  The value stays what it is: formulas that use it still see it (it isn't
  an error), so a total over a column with one bad entry still adds up.
  Formula results are checked as well as typed values.
* **Going to them.** Each count is a link that goes to the next failing
  cell after the active one, in address order, wrapping round
  (Shift-click: the previous), zooming into the grid holding it as
  needed:
  * the **status bar**: "✗ 3 invalid", for the whole sheet ("✗ 1 of 3
    invalid" on one of them; the tooltip counts those in view);
  * the **Outline**: "✗ 2" beside a grid, as "⚠" counts errors (a folded
    grid counts the grids inside it too);
  * the **Rules tab**: "2 invalid" beside each validation rule, for the
    cells failing that rule.
* **Refused.** With `:refuse true`, typing a value that fails (in the
  formula bar, or Ctrl-Enter into a selection) changes nothing: a message
  says why, and after Enter the formula bar keeps the text to fix (leaving
  it by clicking elsewhere drops the text). A control writing
  into the cell (dropdown, stepper, counter-button…) is refused the same
  way, and so are pasting, Fill down and right, moving cut cells there,
  and Find's Replace: one refused value refuses the whole paste, fill or
  replace. Inserting a file or a template and changes upstream (a
  value taken off the list a cell looks up) aren't refused: the cells
  they make invalid are flagged.
* **What's checked.** Cells with a value: empty cells pass (as Excel's
  "ignore blank"), errors are already errors, and cells holding grids
  aren't checked. Every validation rule covering a cell must hold; the
  first that fails gives the message. A predicate that throws on the value
  fails it (`pos?` on text). A rule whose source has an error, or whose
  references reach a cell with an error, doesn't apply until that's fixed
  — the Rules tab shows its error.
* **Controls take what the cells hold, not what a rule checks.** A rule
  draws no ▾ and gives a control nothing: `(one-of D2:D9)` checks a
  value, and that is all. For a drop-down of D2:D9's values, give the
  cells the kind *one of D2:D9* (*Holds ▸ → one of a column…*, §19).
  Renderers get `:domain` in ctx, what the cell's kind allows:
  `{:options :several :min :max :step}`, a lookup's values or the ones
  listed, the kind's bounds, a step of 1 for whole numbers. In the
  standard library, a bare `dropdown`, `radio` or `checklist` (or
  `(dropdown)`, `(radio)`) offers the options; `number-input` sets its
  min, max and step; `(slider)` takes its bounds, and `stepper` without
  bounds keeps to them.
* **Find.** Searching values, a failing cell reads as its value followed
  by `#INVALID` (`12 #INVALID`), so Find `#INVALID` lists them all, as
  `#REF` lists broken cells — including in grids too big for the counts
  above.

## 17. Excel workbooks and Excel's functions

An `.xlsx` workbook can be opened as a sheet, or inserted into a cell,
with its values, formulas, styles, number formats, column widths,
merged cells (as spans, §4) and names. Formulas are translated to Clojure that calls functions with
Excel's own names and behaviour, so the imported sheet computes what
Excel computed, and keeps doing so as it's edited. Each formula's value is
then checked against the value Excel saved, and every difference is
listed. There's no export to Excel: the aim is a faithful import.

* **Open** (`.xlsx`, `.xlsm`) replaces the sheet with the workbook;
  **Insert** puts it into the selected cell as a sub-grid. Both are one
  undo step. The status bar says what came in: how many formulas were
  translated, how many were kept as Excel's values, and whether the
  computed values agree with Excel's.
* **Sheets are sub-grids.** The top grid is a contents page: one row per
  sheet with its name and the sheet itself, a grid cell named after it
  (`Sales Data` → `Sales_Data`; a name that isn't allowed, such as `Q1`,
  which is a cell reference, gets `_` appended). Click the link to open a
  sheet; Out (Alt-Up) comes back. Below are the workbook's defined names
  that aren't a single cell, then `import_notes`, a grid listing what
  didn't come across.
* **Formulas keep Excel's names.** `=SUM(B2:B9)*(1+TaxRate)` becomes
  `(* (SUM B2:B9) (+ 1 ..>Prices>TaxRate))`; `=IF(A2>0,"pos","neg")`
  becomes `(IF (> A2 0) "pos" "neg")`; `=A1=B1` becomes `(== A1 B1)`,
  `=A1^2` `(POWER A1 2)`, `=A1%` `(/ A1 100)`. References keep their spelling,
  `$` included, since RecDep's `A1` and `$A$1` mean what Excel's do.
  Another sheet's cells are reached through the contents page:
  `Sheet2!B3` → `..>Sheet2>B3`, which keeps working when the workbook is
  inserted into a cell.
* **Typing an Excel formula.** In any cell, a formula typed with a
  leading `=` is translated the same way when it's committed: `=A1+1`
  is stored as `(+ A1 1)`, and the status bar shows the translation.
  Function names and references can be typed in lower case. Names of
  cells in the grid, or the grid around it, work as Excel names do
  (`=total*2`), and `Sheet2!A1` reaches a sheet of an imported workbook
  from a neighbouring sheet. What can't be translated is refused with the
  reason, and nothing is written. (A bare `=` is Clojure's `=`, and
  anything that doesn't start with `=` is Clojure.)
* **Tables inside `=`.** A path works inside an Excel formula as it does
  anywhere: `=SUM(B19>D2:D4)`, `=SUMIFS(plan>cost,plan>phase,"Build")`,
  `=SUM(..>resources>rate)`. A `>` with no spaces around it is a path
  when what's left of it is a table (a grid, a cell holding a table such
  as a steps cell, `..` or `$`); otherwise it's Excel's comparison, so
  `=IF(A2>B2,"yes","no")` and `=IF(finish>deadline,…)` over two dates
  compare as in Excel. Spaces always compare: `=sales > 3`. A step that
  isn't in the table is refused rather than compared (*sales has no
  column qtty; its columns are … To compare, put spaces*); so is a name
  holding `-` when both it and the part before the `-` are there (put
  spaces around a `-` that subtracts).
* **Excel's table references** become RecDep's: `Sales[Qty]` is
  `sales>qty` (the column's heading, *Unit price* as `unit-price`);
  inside a table, `[Qty]` is its own column and `[@Qty]` (or
  `Sales[@Qty]`, `Sales[[#This Row],[Qty]]`) this row's cell in it,
  written `$Bu0` (column B, this row), so it stays this row's when the
  formula is filled, copied or moved. `[#Headers]`, `[#Totals]`,
  `[#All]`, a span of columns, and `[@Qty]` from outside the table are
  refused, each saying what to write instead.
* **A function where Excel passes one** is that function:
  `=GROUPBY(A2:A6,D2:D6,SUM)`, `=BYROW(B2:D9,MAX)`. A cell's name comes
  first (a cell named *rate* is that cell, not `RATE`).
* **A name that isn't there** says what to write when it's a table's
  column: *Qty is a column of this table: [@Qty] for this row's*, or
  *sales has a column qty: sales>qty*.
* **In a worked-out column's formula** (the form of *+ Step*), an Excel
  formula works over the row's columns: `=IF(finish>deadline,"late","on
  time")`, `=work*(1-done)`, `=[@Work]*2`. A name that's a column is the
  row's; a cell's or a function's name is that; any other name is read
  from the row, as in a formula written the RecDep way. The translation
  shows in the status bar; reopened, the form shows it.
* **A few names differ from Excel's**, where Excel's name spells a cell
  reference here (`ATAN2` is column ATAN, row 2): `ATAN2` is `ATAN-2`,
  `DAYS360` `DAYS-360`, and `LOG10(x)` is written `(LOG x 10)`. The translator makes these changes,
  typed formulas included.
* **The functions are RecDep's**, for any formula, not only imported
  ones: `(SUM B2:B9)`, `(VLOOKUP A2 Prices>A2:C50 3 false)`,
  `(NETWORKDAYS B2 C2)`, `(TEXT B5 "d mmm yyyy")`. The Library tab's
  Reference lists them first, by section, with their arguments;
  completion and hover show them too. Names are Excel's, upper case, so
  they don't collide with Clojure's (`COUNT`, `count`). `ROUND`'s digits
  may be left out (`(ROUND 2.5)` is 3), a RecDep extension.
* **Looking up many keys at once.** As in Excel 365, `XLOOKUP` and
  `XMATCH` take a range or an array of lookup values and give one result
  for each, in the same shape: `(XLOOKUP (TEXTSPLIT "b,c" ",") ids
  finishes)` is the finishes of b and c, side by side. A key that isn't
  there is `#N/A` in its own place (or the if-not-found value), with the
  others found. A column of keys looked up in a return range several
  columns wide gives a row for each key (Excel keeps only the first
  column); keys in rows and columns both, each giving several values,
  are `#VALUE!` saying to look up a single column of them. A list (a
  list column's value, `["Dee" "Chen"]`) gives a column; `SUMPRODUCT`
  pairs a list with a row or a column of as many values, so `(SUMPRODUCT
  booked (XLOOKUP resource resources>resource resources>rate 0))` is
  each person's days times their rate, summed.
* **What doesn't translate keeps Excel's value**, written as a literal
  with the original formula in a comment:
  `"apple" ;; Excel: =INDIRECT("A1") — not translated: INDIRECT builds
  references while it runs …`. It's listed in `import_notes`, and Find
  (formulas, `not translated`) finds them all.
* **Checked against Excel.** After the import, each translated formula's
  value is compared with the one Excel saved in the file (numbers to a
  relative 10⁻⁹). Differences go into `import_notes` ("computes 12,
  Excel saved 13"), so a translation that misbehaves is visible straight
  away rather than silently wrong.

### Not supported

Some of what a workbook can hold has no counterpart here, or can't be
had with fixed dependencies. These arrive as values, or not at all, and
are noted in `import_notes`:

* **Macros and code**: VBA (`.xlsm` opens, its macros don't come or run),
  add-ins and XLL / automation functions, Office Scripts, Python in
  Excel, `CALL`, `REGISTER.ID`.
* **Dynamic references**: `INDIRECT` and `OFFSET` build references while
  they run, and RecDep's dependencies are fixed when a formula is
  compiled (§2, §3); `CELL`, `INFO`, `SHEET`, `SHEETS`, `AREAS`,
  `FORMULATEXT`, `ISFORMULA` ask about the workbook itself. Reference
  intersections (`A1:B5 B2:C3`), unions (`(A1,B2)`) and ranges built from
  functions (`A1:INDEX(…)`) too.
* **Outside data**: links to other workbooks (`[1]Sheet1!A1`), data
  connections, Power Query and the data model, `WEBSERVICE`, `RTD`,
  `STOCKHISTORY`, cube functions, `GETPIVOTDATA`.
* **Objects and presentation**: charts, chart sheets, pictures, shapes,
  comments and notes, sparklines, conditional formatting, data
  validation, hidden rows and columns (shown), freeze panes, AutoFilter, outlines and
  grouping, page setup, protection. Pivot tables arrive as their last
  values.
* **Calculation settings**: manual calculation (F9 recomputes only the
  volatile cells, §28), precision as displayed; what-if data tables
  (`TABLE()`) keep their values (RecDep's own are `what-if`, §18).
  Iterative calculation is kept: the top grid's library gets
  `iterate-cycles`, with the workbook's rounds and change (§30).
* **The 1904 date system** (old Mac workbooks): dates show four years and
  a day early; noted.
* **`.xls`** (the binary format before 2007) and `.xlsb`: save as
  `.xlsx` first.

Known differences in what's translated:

* A single-cell argument counts as a value given directly, not as a
  reference: `SUM(A1)` with text `"3"` in A1 is 3 here, 0 in Excel
  (Excel skips text in references). Ranges behave as in Excel.
* `TODAY`, `NOW`, `RAND` and `RANDBETWEEN` are computed when the formula
  is, not on every recalculation (RecDep recomputes what changed); F9
  computes them again (§28).
* Excel's error values count as errors in the status bar and the outline's
  ⚠ (§11), as Excel's own error checking flags them: one made in the cell
  is where an error starts, one read from an error cell is passed on.
* A formula that is just a reference to an empty cell shows empty, where
  Excel shows 0.
* `ROW(B7)` and `COLUMN(B7)` are worked out when translating, so a copy
  elsewhere still gives 7. (`ROW()` without a reference is `(:row
  (here))` and follows its cell.)
* `SUBTOTAL` and `AGGREGATE` can't skip hidden rows (none are hidden
  here) or nested subtotals.
* Sheets are as big as their used area, so a whole-column reference
  (`A:A`) covers the rows the sheet has, and `ROWS(A:A)` isn't 1,048,576.
  Growing the sheet's grid (§6) grows what `A:A` covers.
* General shows up to 11 characters whatever the column's width.
* Copying an imported formula follows RecDep's rule (§5): only
  coordinates in the formula's own grid move, so `..>Sheet2>B3` stays.

## 18. What-if tables and goal seek

Two questions every model gets asked: what would the result be for each of
these inputs, and what input gives the result I want? Excel answers them
with data tables and Goal Seek, both dialogs. Here they're functions on
the "evaluate with overrides" mechanism of `call` (§7): a template grid is
evaluated with an input pinned, many times, and nothing in it changes.

```clojure
;; Loan holds a template: amount, rate, years -> payment (all named)
(what-if Loan :rate A3:A9 :payment)            ; a payment per rate: a column
(what-if Loan :rate [[0.03 0.05 0.07]] :payment)  ; a row of rates: a row
(what-if Loan :years [5 10] [:payment :years])    ; a column per output
(what-if Loan :rate A12:A14 :years B11:D11 :payment)  ; two inputs: a table
(goal-seek Loan :payment 500 :amount)          ; the loan a 500 payment buys
(goal-seek Loan :payment 300 :rate {:min 0 :max 0.2})
```

* **`(what-if grid input values output)`** is Excel's data table: the
  value of `output` for each of `values` given to `input`, laid out as the
  values are (a column for a column or a list, a row for a row). It spills
  (§13), so it sits beside the inputs it was given, as Excel's does. A
  vector of outputs gives a column each; a blank input gives a blank row.
  **With two inputs**, `(what-if grid in-1 values-1 in-2 values-2 output)`
  is a table: a row per value of the first, a column per value of the
  second.
* **`(goal-seek grid output target input)`** is Goal Seek: the value of
  `input` that makes `output` equal `target`. It starts from the input's
  own value; options `{:guess :min :max :tolerance}` give a start, bounds
  and how close is close enough (10⁻⁹ of the target by default). No
  value within 100 tries is an error that says how close it came.
* Inputs and outputs are cell names or addresses in the grid, as for
  `call`; the result follows the template and the values, like any
  formula. Both work on a linked instance (§21), starting from its own
  values.
* **Goal seek… on the cell menu** of any formula, in any grid, does the
  same for the sheet itself, as Excel's dialog does: it asks for the
  target and the cell to change (offering the nearest number the formula
  depends on, typed into its cell rather than computed), finds the value,
  and, if you keep it, types it into that cell (rounded to 12 significant
  digits; one undo step).
* **Several inputs, constraints, the largest or the smallest:** the
  solver, `optimize` and *Solver…* (§39).

## 19. Headed grids as tables

A grid whose first row is headers is a table, as Excel's tables and
Clojure's sequences of maps are: its value is a vector of row maps, each
header names its column, and a sequence of maps spills back into cells
with its header row. That joins grids, `call` and spills to ordinary
Clojure data.

* **Headers.** Tick *Row 1 is headers* for a grid cell in the Inspector,
  or right-click a cell of a sub-grid → *Use row 1 as headers* (and
  *Row 1 isn't headers*). Cells you typed anywhere become a table with
  *Make this a table* (below). The header row is drawn bold, underlined.
* **The value** of a headed grid is a vector of row maps, one per
  record (the rows below the headers down to the last holding
  anything, *A table's rows* below), keyed by the headers:
  `[{:item "Apples" :qty 30 :price 0.5} …]`. A header becomes a keyword
  in lower case, with runs of other characters as one `-` (`Unit price`
  → `:unit-price`, `Total ($)` → `:total`); a duplicate is numbered
  (`:qty-2`), and a column without a header keeps its letter (`:C`).
  Clojure works on it as on any data: `(map :qty orders)`, `(filter #(>
  (:qty %) 20) orders)`, `(sort-by :qty orders)`. `call` on it gives the
  table too. (`(:A2 orders)` doesn't read a cell: the value isn't a map
  of cells.)
* **Excel's functions** read a table as the range of its data rows, the
  header row left out, as Excel reads a table's name: `(VLOOKUP "b"
  orders 2)`, `(INDEX orders 1 1)` (the first data row), `(ROWS orders)`.
  Columns go by position (by key order, which is the headers'), so a
  repeated header doesn't move a column. A blank row among the records
  is one of them, as in Excel's tables: INDEX and ROWS count it, and
  its columns hold it, so the table and its columns always have the same
  rows.
* **Columns by header.** `orders>qty` is the qty column below the header
  row, to the table's last record, as a range: Excel's structured reference
  `Orders[Qty]` (an import turns those into plain ranges, §17).
  `B5>qty` works too, and from inside the grid `..>orders>qty`. A cell
  named `qty` in the grid comes first. Retyping a header renames its
  column: the formulas naming it follow (*A heading retyped*, below);
  clearing it, or turning headers off, makes them `#REF`. Inserting
  rows grows it.
  Completion after `orders>` offers the headers.
* **A column of a table a cell holds.** A cell whose value is a table
  (a formula's result kept by a button, `(steps …)` drawn in its cell)
  has columns too: `baseline>kept>finish` is the finish column of the
  table `kept` holds, as a list, `(MAX baseline>kept>finish)` its last
  date. On a cell holding a map it's a key: `settings>prefs>level`, where
  `prefs` holds `{:level true …}`.
  It follows the cell. A column the table hasn't is `#REF` naming those
  it has; so is an empty cell, or one holding a number or text.
* **Tables spill.** A formula returning a sequence of maps spills a header
  row of the keys and a row per map (§13); a missing key leaves its cell
  empty. The columns come in the order the rows were built: a row read
  from a grid, or made with `array-map`, keeps its order. A Clojure map of
  more than eight keys keeps none, so a row of the table given a ninth
  key (`(assoc row :finish …)`) would come out scrambled: its table's
  columns are kept in their order, and the new keys follow. Where the
  order matters, give it: `(as-grid rows {:columns [:id :task …]})`, or
  build the rows with `array-map`; and format such a grid by header
  (§10's `:column`), which no order can move. `(table-columns rows)` says
  the order a table's columns are laid out in. The formula's cell
  keeps the maps as its value and shows the first header. *Make table a
  sub-grid* on such a spill makes a grid with headers. `table` (and the
  default renderer) draws a table of maps with a header row.
* **Several spills, with headings, into one sub-grid.** Headings typed
  over columns that `down` fills are a table too, but each column is its
  own spill. A selection that touches a spill and reaches past the active
  cell's own one (the headings and the columns, or cells across two
  columns) offers *Make A1:B30 a sub-grid*: the selection, grown to take
  in whole every spill it touches, becomes a sub-grid in its top-left
  cell holding what its cells show, as literals with their number formats.
  The other
  cells are emptied (formats stay, as *Clear contents* leaves them); a first row of
  headings (text outside any spill, or a row-maps spill's header row)
  becomes the headers. One undo puts it all back. A selection holding a grid or a span is
  refused. Within one spill it's still *Make table a sub-grid*.

### Column kinds

Each column of a table can say what it holds: **text**, **numbers**,
**whole numbers**, **dates**, **TRUE or FALSE** or **days of the week**
(`Mon-Fri`, `Sat, Sun`: what a calendar's Weekdays holds). A value of another
kind is flagged as invalid where it is (the magenta corner, its reason
on hover: *must be a number*) and counted in the status bar's ✗, as any
failed validation is (§16), so text typed into a column of rates
doesn't go unnoticed.

* **Set when the headings are**: ticking *Row 1 is headers* gives each
  column whose values are all of one kind that kind. A column of mixed
  values gets none.
* **Shown on the heading**, small, at the right: *abc*, *1.5*, *123*,
  *date*, *T/F*, *ID*, *a, b* (lists), *Mon–Fri* (days of the week),
  *one of*, *list of* (one or more of), and *·1* after it for each value
  once. Hover it for what it means.
* **Changed on the heading**: right-click it → *Column holds: numbers ▸*
  and choose *text*, *numbers*, …, or *any value*; the current one is
  ticked. One undo step.
* A kind is a rule of its own over the column below its heading, `{:area
  "D2:Dde" :kind :number}`, on the Rules tab with the grid's other
  rules (it reads *holds numbers*): copied, saved and moved with the
  grid. It checks its cells as a validation rule does (§16), and it is
  also what the ▾ and a bare control offer; a validation rule only
  checks.
* The Inspector's *Holds* row says what the selected cell holds, where
  that comes from and how a value goes in: *one or more of tasks' ID
  (F2:Fde) · ticked with the ▾*.
* Typing a number into a column of text enters text (an ID, a code:
  `007`).

### Any cells: Holds ▸

A kind isn't only a table column's. Select any cells, a single cell of
an input form, a block, and right-click → *Holds ▸*: the same choices as
*Column holds*, for those cells. The form's item cell holds *one of a
column…* (`items>item`), and its drop-down offers the items; a status
cell holds *one of values listed…* (`Low, Medium, High`). One kind per
cell: on cells inside another kind's area, *Holds ▸* says which area
holds what and offers to select it, to change it there (or make it *any
value* first). On a cell of a table column with a kind, it is that
column's menu.

### What a whole table holds: shapes

Some tables have a known shape: a set of columns, each holding a kind,
some needed. RecDep knows five: *a calendar* (Resource, From, To,
Weekdays, and Capacity, Takes or Uses), *a list of dates*, *a
schedule's rows* (ID, Task, Work, Resource, After, Not before,
Priority, Started, Finished, Share, Done, Due), *a Gantt's rows*
(Name, Start, Finish, ID, After) and *a shape spec* (Column, Holds,
Needed, Also known as, Means, Example: what defines a shape of your
own). Every column of a shape says what it
holds: Takes is *a booking: the share of the day it takes from the
capacity*, Uses *a booking of a pool: how much of the capacity it
uses*. A workbook can add its own (*Shapes of your own*, below).

Nothing is declared: a table whose columns hold what a calendar's do is
a calendar, made however.

**This table holds**, on the Shelf. Select a table (or its cell) and the
Shelf tab's *This table holds* has a button for each shape the table is
or nearly is: ticked when it is one already, *renaming Start to From*
when a column is named nearly right, *a part of one* when it lacks a
column it needs (a calendar's layer gets its Capacity from a step).
Choosing one gives every column the shape has, by its heading, what that
column holds, in one undo step: the kinds, their bounds (To on or after
From, Capacity at least 0), a lookup of the sheet's column of names or
of the table's own IDs, and a date format on dates. The shapes the table
isn't are folded away under *… shapes it isn't: why*, each saying what
it needs and the columns the table has.

A table a formula makes (a steps formula's, such as a calendar's layers
combined, or what's left of it after its bookings) says the same, with
nothing to click: *left is made by its formula, and is a calendar, and
nearly a Gantt's rows*. Its columns are its formula's, so a step changes
what they hold.

**An empty grid** is laid out instead: a button per way the shape can
start. A shape with an either-or column has one per choice: *Lay out as
a calendar of Capacity* (Resource, From, To, Weekdays, Capacity: working
time), *of Takes* (bookings) and *of Uses* (a pool's bookings), each
headed and given an example row.

**Start a table.** At the top of the Shelf, *Start a table* places a new
table beside the selection, laid out the same way, for every shape and
every way it starts: a calendar of Capacity, of Takes or of Uses, a list
of dates, a schedule's rows, a Gantt's rows, a shape spec, any shape of your own, and
the working calendar in layers (§51).

**A table's other columns.** A layout is where a table starts, not all
it can be. Type a heading in a table on its way to a shape (it is one,
nearly is, or has two of its columns), in the faint empty column after
the last or over a heading, and completion offers the shape's columns it
hasn't got, each with its shape and what it holds (*Takes: numbers at
least 0: a booking …*); Tab takes one. Entered, a heading that names one
of the shape's columns gives its column what it holds, and says so
(*Takes holds numbers at least 0 …, as a calendar's does*). The shape's
columns typed before it, while the table wasn't on its way to the shape
yet, get theirs too: type x, then y, and both hold numbers as a route's
do (*y holds numbers …; x too*). So a
calendar gets Takes and Uses beside its Capacity, a schedule its Due.
A shape is its columns' kinds as well as their names: a table laid out
as one, or headed into one, holds what the shape's columns hold, and a
value of another kind typed in is flagged.

**Functions that take a shape.** A function says which of its
arguments take a table of a shape: the date functions' calendar
(`capacity`, `working-days`, `work-end`, `week-start`, `month-view`,
`calendar-for` …) and the Excel functions' holidays (WORKDAY,
NETWORKDAYS and their .INTL forms) take a calendar; `gantt-view` takes
a Gantt's rows. Then:

* Hovering the function in a formula shows, after its description,
  what each such argument takes: *calendar takes a calendar: a table
  with Capacity, Takes or Uses; also Resource, From, To, Weekdays (or a
  list of dates)*, and a line per column, its kind and what it means.
* Typing that argument, completion offers the tables in the sheet that
  are of the shape first (*cal — a calendar*), then those nearly of it,
  with what they lack.
* Given a table that isn't one, the call is refused before it runs, in
  the shape's words: *capacity: its calendar isn't a calendar. A
  calendar table needs a Capacity column … Its columns are Who, Hours.*
  (A shape whose columns a function picks by its options, a Gantt's
  rows, is left to the function: `gantt-view` says which option names a
  column the table hasn't got.)
* Select a cell holding a table of a shape, and the Shelf's *As an
  argument* offers the functions taking it, as a call: `(capacity
  (TODAY) cal)`, `(month-view (TODAY) cal)`, `(WORKDAY (TODAY) 10
  cal)`. A click places it beside the selection; change its other
  arguments in the formula bar.

**Shapes of your own.** `(shape spec)` makes a shape from a shape
spec, a table of column specs, a row per column:

| Column | Holds | Needed | Also known as | Means | Example |
|---|---|---|---|---|---|
| Person | one or more of people>name | yes | Who, Name | who worked | |
| Date | dates | yes | Day | the day worked | 2027-03-01 |
| Hours | numbers at least 0 | yes | Time | hours worked that day | 7.5 |
| Note | text | | Why | | |

* **Holds** is written as *Column holds ▸* says it: text, numbers,
  whole numbers, dates, TRUE/FALSE, IDs, lists, days of the week; with
  *at least 0*, *at most 1*, *between 0 and 1*, or *on or after* another
  column; or *one of* / *one or more of* a column (`people>name`, or
  *the Name column*; naming its table, `customers>id` or *customers'
  id*, so a database's tables, each with an id, are told apart); or
  *one of* values listed (*one of open, shipped, invoiced or
  cancelled*, their case kept). Blank: anything. Anything else is
  refused, saying what it may be.
* **Needed** is *yes*, blank (optional), or a group's name: the columns
  sharing a group need any one of them, as a calendar's Capacity, Takes
  and Uses do; a table of the shape starts once per member of a group.
* A table laid out as the shape gets its needed columns and those with
  an Example, headed as Column writes them (*x* stays *x*).

A shape spec is a shape itself: select an empty cell and *Start a
table ▸ shape spec* lays one out, its headings and an example row to
fill in. It is an ordinary table of text, nothing in it a formula. Select
it and the Shelf's *As an argument* offers `(shape spec)`, and places it
in a cell of its own beside the spec, named: *route-spec* gives *route*,
a spec named otherwise *my-shape* (numbered when taken); rename the cell
to rename the shape. Typed by hand, it goes in a cell of its own too,
never on the spec grid's own cell (a grid's cell only takes a map giving
its parameters values). Given a
table that isn't a shape spec, `shape` is refused before it runs: *its
spec isn't a shape spec. A shape spec needs a Column column …*.

Name the cell holding it (*timesheet*) and it is a shape in scope, for
the cells of its grid and the grids inside it, as the built-in ones are:
*This table holds* and *Start a table* offer it, headings complete its
columns, and a library's function says an argument takes it by its name:

```clojure
(defn hours-of {:takes {:sheet "timesheet"}} [sheet who]
  (reduce + (keep #(when (= who (:person %)) (:hours %)) sheet)))
```

Its hover then shows what a timesheet is, completion at `sheet` offers
the sheet's timesheets, and `(hours-of notes "Ana")` on a table that
isn't one is refused before it runs: *hours-of: its sheet isn't a
timesheet. A timesheet needs Person and Date and Hours … Its columns
are …*. A library function can take a built-in shape the same way:
`{:takes {:cal :calendar}}`. The tour's *Shapes of your own* section
works one through, a route of x, y and z, from its spec to a function
taking it; its Working calendars section has a timesheet.

**A table's shape.** In a sheet with a database (§53), `(table-shape
"customers")` is the shape a table of it is, read from the table
itself: a column per column, holding what its type and its rules say
(a foreign key *one of customers' id*, a `check (… in (…))` *one of*
its values, `check (x >= 0)` *at least 0*), its NOT NULL columns
without a default needed (but a key the database numbers), its key
said. Named, it is a shape in scope as a spec's is, and it follows the
table: there is one place to say what a customer is. What its columns
mean is said beside it, `(table-shape "customers" customer-notes)`: a
table of Column and any of Means, Also known as and Example, a row per
column (a shape spec without Holds and Needed, which are the table's
to say; giving them is refused). The other way, a
table a shape of yours is copied into the database (§53) takes the
shape's rules as the table's.

### What a column is for: This column is ▸

After the kind, a typed table's heading says what its column is for,
its **role**, worked out from what it holds:

| Tag | Role | Worked out from |
|---|---|---|
| key | the table's key: each row's own | the first column of IDs |
| name | the name of each row's key | the first text column after a key that numbers its rows |
| → resources | a reference: what each row is about | *one of* / *one or more of* another table's column |
| link | a link to rows of this same table | *one or more of* its own IDs (After) |
| day, week… | time, by the day (or week, month, year) | dates; two dates, one on or after the other, are an interval |
| Σ, avg, w.avg… | a measure, and how it combines | numbers (a rate, money per day, is averaged) |
| (none) | an attribute: carried along, grouped by | the rest |

Its hover says it in words: *Done is a measure, averaged, weighted by
Work when rows are combined.*

Where that's wrong, say otherwise: right-click the heading → **This
column is ▸**: the key, the name, time by the day, week, month or year
(or, for numbers, *days from a date…*, which asks for the date's cell), a
measure summed, averaged, averaged weighted by another measure of the
table, the smallest, the largest, the last, or not combined, or an
attribute. What was said is marked `*` on the heading; *As its kind
says* takes it back. A role is kept with the column's heading: retype
the heading and it goes along, and so does a weighting by it. While the
heading is protected, the change is refused, saying so.

**This table's rows are ▸**, on the same menu: things (one a row, each
with its key), occurrences (each about what its columns name), or rules
(a condition and a value, the last match winning: a calendar). It's
worked out too: a calendar's columns make rules, a key makes things,
the rest occurrences.

**What a role opens.** A totals row (Totals row, on a heading's menu)
totals each measure as its role combines it: Work summed, Done weighted
by Work (a 1-day task counting less than a 20-day one), a rate averaged.
Its cell's *Total ▸* offers *As its role combines it* beside Excel's
choices. And the Shelf, for a table you select, says what one more step
would open under **One more step**: *Done holds 0 to 1: if it's a part
of each row's Work, This column is ▸ a measure, averaged, weighted by
Work says so, and a totals row then weighs it.*

### What a table is: This table is

A table can be used as something: a plan is what a Gantt draws and a
cost roll-up totals; capacity is what a schedule fits work into. That's
its **concept**. Two tables can hold the same kind of columns and mean
different things (capacity and demand are both days of work, by
resource, by day), so a concept is said, not worked out, and checked
against the columns it needs.

Select a table (or its cell) and the Shelf's **This table is** says what
it is:

* **said**: *resources is resources (said): who or what does the work*,
  with *Take it back*;
* **given by what made it**: *plan is a plan (given by its Schedule
  step)*. Steps after it that keep its rows (a filter, a sort, a
  worked-out column, columns chosen, an append, a join) keep it; a step
  making new rows (a group, a pivot, a series) doesn't. A table worked
  out from another keeps what that one was said to be: say the
  calendar is capacity, and what's left of it after the plan's bookings
  is capacity too. `(bookings plan :task)` gives bookings;
* **said but not fitting**: ⚠ *tasks is said to be capacity, but it
  isn't a calendar: …*, with how to take it back.

Then a button for each concept it fits and isn't: **It is resources**,
**It is capacity** … A click says so on the table's cell (refused,
saying so, while the cell is protected). Blocks that take that concept
find it.

RecDep's own concepts, for planning: resources, tasks, capacity,
bookings, a plan, demand, effort and a load. Your own: give `shape`
options, `(shape spec {:refines :calendar :means "the support rota" :rows
:rules})`, or `(shape nil {:refines :calendar :means "…"})` for a
calendar with nothing more; name its cell, and This table is offers it.
Extra columns never matter: a plan with an approver and a budget is
still a plan, and which plan is which is the cell it's in.

**Inserting a grid** that defines a shape or concept by a name this
sheet defines differently says so: inside the inserted grid its own is
meant, outside it this sheet's.

### Bounds: at least, at most, between; on or after

Numbers, whole numbers and dates can have bounds, on the same menu:

* **at least…**, **at most…**, **between…**: it asks for the bound (two
  for *between*, `0 to 1`, both included; dates as `2027-01-04` or
  `4 Jan 2027`), and asks again, saying why, when what's typed isn't one.
  The heading then says *numbers at least 0*, *numbers 0 to 1*; a value
  outside is flagged in its cell, *must be a number, at least 0*, and the
  number controls (a slider, a stepper) keep to it. A column holding a
  capacity, a share of a day or a price. (A validation rule's
  `(between 0 1000)` checks, but bounds no control: it is a test of the
  value, not what the cells hold.)
* **dates on or after** another column of the same row: the menu offers
  the table's other columns of dates by their headings (*dates on or
  after From*), and *dates on or after a column…* asks for any heading.
  A To before its row's From is flagged, *is before its row's From (12
  Mar 2027)…*; a blank From isn't compared. The column is found by its
  heading, so moving or inserting columns doesn't change what it
  compares with, and retyping the heading (*From* to *Start*) keeps the
  rule on it.
* **numbers at least, or at most, another column** of the same row: the
  menu offers the table's nearest columns of numbers by their headings
  (*numbers at most Work*, *numbers at least Work*); the Rules tab's
  *At least its* and *At most its* take any. A three-point estimate's
  Low above its row's Work is flagged, *is above its row's Work (6):
  correct this one, or Work*, and with *Refuse a failing value* ticked
  it's refused where it's typed. A blank Work isn't compared. It goes
  with a bound of its own (*numbers at least 0, at most Work*), and the
  column is found by its heading, as for dates.
* Choosing the kind itself (*numbers*, *dates*) takes the bounds off.

### IDs, lists and lookups

Four more kinds, on the same menu:

* **IDs**: text, each once. A typed `3` is the text `"3"`, so a column
  of IDs matches text IDs in other tables (a join, a lookup). An ID used
  by another row too is flagged on both: *this value is used by another
  row too*. **IDs within** another column (*IDs within Project*, offered
  for the table's columns of text nearby): each once per project, so
  two projects can both number their tasks 1, 2, 3 (§49, *Several
  projects in one plan*).
* **lists**: typed `Dee, Chen`, the cell holds the list `["Dee"
  "Chen"]`. Commas or semicolons separate the items; spaces round them
  and empty items are dropped. It's shown as `Dee, Chen` and edited as
  that text (F2, or the formula bar). A list is a value formulas read as
  a list: `(some #{"Dee"} C3)`; a `steps` formula's *Unnest* step
  gives a row per item, so a list column can then be joined or counted.
  An item can't hold a comma or a semicolon.
* **one of** a column: the value is one of another column's values (an
  order's product, an invoice's customer). Or **one of values listed…**:
  `Low, Medium, High`, typed once.
* **one or more of** a column: a list whose items are each one of that
  column's values (a task's people; the tasks it waits for).

For the last two:

* **Choosing the column**: the menu offers the sheet's columns of IDs by
  name (*one or more of ..>resources>resource*, *one or more of A2:Ade*
  for the table's own ID column), nearest first. *one of a column…* asks
  for any other, written as a formula in the table would write it; if
  what you type isn't one column of a table, it asks again, saying why,
  with your text kept.
* **Picking**: select a cell and a ▾ shows at its right. It opens the
  column's values with a field to filter them: a click picks one; for a
  list, ticks add or remove items, the others kept as they are. Esc or
  *Done* closes it. What it writes is typed into the cell, so protection,
  undo and a refused value work as for typing.
* **Each value with its description**: when the column looked up is in
  a table with a column of text, the picker shows each value with that
  row's text beside it, from the nearest text column to its right: a
  task's ID with the task's name, a person with their role. Typing in
  the filter finds either, so `audit` finds *3 Content audit*; what's
  written is the value alone (`3`).
* **IDs shown with their names**: a cell holding IDs of a table with a
  column of text shows each with that row's text, the ID muted, as the
  picker lists them: a Timesheet's Task reads *12 Content migration*,
  an After *9 Page templates, 11 Content writing*, an ID with a lag by
  the task it names (*12-2d Content migration*). The cell still holds
  the IDs: formulas, steps and saving see `12`, Find searches the IDs,
  and renaming the task changes every cell showing it. Only IDs that
  are codes, with a digit in them (12, T1, Web:3), are shown so: a name
  (Ana) reads already. An ID on two rows of different names (one per
  project) is shown alone, since which is meant can't be told. A cell
  with a renderer of its own draws as that says. A table a formula
  makes of one (sorted, grouped, a column renamed, given to a block)
  shows its IDs with their names too, and its headings say what each
  column holds, as the table it came from says: *one of · tasks' ID*.
  A column the formula works out afresh holds whatever it makes.
* **A value that isn't there is refused** when it's typed, pasted,
  filled or written by a control, naming it, what the cells hold and
  what to do: *"Zed" isn't one of resources (B19)'s Resource (not entered):
  add it to resources (B19)'s Resource first, or pick one with the ▾*. A task can't
  name a resource before it's on Resources. One already there that
  stops being there (taken off Resources) is flagged instead, as is a
  column given a lookup with values already in it. The Rules tab's kind
  editor has the switch (*Refuse a value that isn't there when it's
  typed*); a lookup written by hand without `:refuse true` only flags.
* **An item with a lag**: an item names a value when it is that value,
  or that value followed by `+` or `-` and a number, with letters after
  it if you like: `3+2d`, `12-1d`, `T-1 + 2 days` all name an ID (`T-1`
  is matched whole first). The item keeps its lag for the sheet to read;
  anything else after it (`3x`) is flagged.
* **Items are text**, so `2, 3` matches text IDs; a single *one of*
  value typed as a number is text when the column it names holds text or
  IDs.
* **The column it names follows**: renaming its table, inserting columns
  before it, or retyping its heading keeps the lookup working. If the
  column goes (deleted), the heading shows ⚠ and says why on hover: its
  cells aren't checked until you choose the column again.

**Each value once**, at the end of the menu, goes with any kind: a
column of dates with one row per day, say. IDs have it already.

**A blank reads as…**, on the same menu (a column of numbers, whole
numbers, dates, text, TRUE or FALSE, or one of values listed), says what
a blank cell of the column reads as in the table's rows: Tasks' Done
reading 0 where nothing's typed, so the steps and views reading Tasks
get 0 there, not a blank to guard against. It asks for the value, typed
as the column's kind is; one of another kind, or outside the column's
bounds, is refused in the question, saying why; left empty, a blank
reads as blank again. The cell itself stays blank: in a row that holds
something it shows the value faintly, in grey italics, and its hover
and the Inspector's Value row say what it reads as. The heading says
it too: *numbers 0 to 1, a blank read as 0*. A reference to the cell
or to the column (`tasks>done`) reads the cell, blank, as Excel would;
an empty row isn't a row, so it gets nothing. An ID column has no
default: each row's is its own; nor have lists and lookups of another
column. Changing the kind keeps the default while it fits.

**Setting a kind converts what was typed** in the column, in the same
undo step, and says how many: numbers become text for text and IDs,
text becomes lists for lists (`Dee, Chen` → `["Dee" "Chen"]`), a list
becomes its text for the other kinds, and text holding a number becomes
the number for numbers. Formulas are left alone. Pasting into a column
of a kind converts what's pasted the same way.

### Tables that grow as you type

Type into the last row of a table and it grows by a row, so there's
always an empty row under the last one to type the next record into, as
Excel's tables extend. No need to resize the grid.

* The new row is part of the same undo step as what you typed.
* **Calculated columns fill in**: in a row that was empty, a column whose
  two rows above hold the same formula (as copied down) gets it too, so a
  new record's cost or total is worked out as soon as you start it.
* Format and validation rules that run to the bottom edge (`B2:Bde`, the
  column kinds) cover the new row already.
* A paste that runs past the last row grows the table to take it, and
  keeps an empty row after it (*At a grid's edge*, §12); filling
  doesn't grow it.
* With a totals row, the table grows above it: the totals stay last
  and take the new record.

**Columns grow too.** Zoomed into a table (a grid with headers that you
typed, not one a formula makes), an empty column is drawn after its
last one, its letter faint. Move right from the last column (or Tab, or
type its address into the name box: `tasks>M1`) and type there, and the
table has one more column, in the same undo step:

* typed in the heading row, what you type names the new column;
* typed lower down, the value goes in and the column is headed
  *Column1* (*Column2* if that's taken, …), as Excel names one: type over
  the heading to name it.

Rules that run to the right edge (`re`) cover the new column. In a
protected grid, typing there is refused, as Insert is, and the message
says so by the cell; open the lock 🔒 above the grid to add a column. A
table drawn in its cell, or more than 50 columns wide, doesn't have the
empty column: there Tab from the last column adds one, headed *Column1*
(§12, *At a grid's edge*).

### A table's rows

A table's **records** are its rows below the headings down to the last
one holding anything. The empty row typing keeps under them, where the
next record goes, isn't one: not in the table's value, not in
`orders>qty` or `Orders[Qty]`, not counted by `ROWS`, not a blank group
in `GROUPBY`. Type into it and it is. A blank row among the records is
one, as in Excel's tables: delete the row to take it out.

**A totals row.** Right-click a heading → *Totals row*: a row is added
at the foot, drawn bold under a line, *Total* in its first column and a
sum of the last column of numbers, as Excel's *Total Row* starts.

* *Total ▸* on a heading, or on a totals cell, picks that column's:
  *None*, *Sum*, *Average*, *Count* (what isn't blank), *Count numbers*,
  *Min*, *Max*. It writes the formula, `(SUM ..>orders>qty)`, which you
  can read and change; type any other formula there and it's kept.
* The totals row isn't a record either: not in the table's value, its
  columns or `ROWS`, so a total reading the table (`(ROWS ..>orders)`)
  works. Kinds aren't checked there: a count in a column of text is
  fine.
* Records go above it: typing into the empty row, a paste, Insert below
  the totals, a bigger size, all add rows above the totals. Enter (or ↓
  in Enter mode) from the totals row stays, saying to type the record in the row
  above it. Untick *Totals row* to take the row away (Undo brings it
  back); deleting the row does too.

**Fill down** (Ctrl-D) from one cell of a table fills down to its last
record, not to the bottom of the grid; on the last record, or below it,
there's nothing to fill, and it says so. Select the rows to fill a
range as before.

**The Shelf** places a chart, view or formula of a table's cells in the
grid around the table, beside it (to its right, else below), reading the
cells through the table: `orders>A1:D7`. Zoomed into the table, the page
shows the grid around while you place it. A function's cells under the
columns, or a template, would be records in the table: they're refused
there; select cells outside the table, or use *Totals row* for totals
under the columns.

**Make this a table.** Select cells you typed, with their headings in
the first row, and right-click → *Make A1:E6 a table* (on one cell,
*Make this a table* takes the block of cells around it, as Excel's
Ctrl-T does). In one undo step:

* the cells become a grid in the first one's place, row 1 its headings,
  each column given the kind its values have, with an empty row under the
  records for the next;
* it's named *table1* (*table2* … when that's taken): rename it in the
  name box, and every reference follows;
* formulas reading the cells read them in the table: `(SUM D2:D6)`
  becomes `(SUM table1>D2:D6)`, a cell named *rate* in it `table1>rate`,
  and formulas inside it reading cells outside reach them with `..>`.
  A formula whose range has cells both in and out of the selection is
  named, and nothing changes: select all the cells it reads, or none.

### A heading retyped

Retyping a table's heading renames its column, as renaming a column of
an Excel table does: what names the column follows, in the same undo
step. Retype Tasks' *Work* as *Effort* and:

* references by heading, `tasks>work` and `..>tasks>work`, become
  `tasks>effort`;
* the kinds and checks naming it (*one or more of tasks' ID*, *dates on
  or after Started*) keep naming it;
* **table steps over the table** follow: the fields naming the column
  (*Filter* on Work, a *Sort*, a *Columns* step keeping it, a
  Schedule step's Work, a *Join* bringing it in from this table), and a
  worked-out column's formula where it reads the row:
  `(fn [{:keys [work done]}] (* work (- 1 done)))` becomes
  `(fn [{:keys [effort done]}] (* effort (- 1 done)))`, `(:work %)`
  becomes `(:effort %)`. When the formula already uses the new name for
  something else, the old one is kept and given the column:
  `{work :effort :keys [done]}`;
* views of the table naming it follow: `gantt-view`, `filter-view`,
  `group-view` and `pivot-view` options, a `snapshot-button`'s
  `:keep`, `sort-rows`, `filter-rows` and `group-rows`; `picked`
  gets the pair `[:effort :work]`, the table's name for the column and
  the one the rows picked still use (they come from a chart of a table
  worked out by steps, which keeps its names).

**A table made by steps keeps its columns' names.** The plan reads
Tasks' Work and shows it as *work*; after the retype its steps read
*effort* and end by naming it *work* again (`:rename {:effort :work}` in
its last Columns step, or a Columns step of its own saying why), so
what reads the plan, `plan>work` in the summary, a chart, a calendar
reading a *Resource* column, is left as it was. The same happens
before an *Append*, which matches columns by name, and the table itself
appended (a calendar's working week) is given under the names it had,
so the match the append recorded (§43) still holds. Typing a heading
back to the name an append matched (after clearing it and typing it
wrong) mends that append instead.

**A name the steps make** is left to them: retype *Work* as
*Remaining*, and the plan, which works out a column *remaining* of its
own, isn't rewritten: its steps start by naming Tasks' column *work*
again, and the rest reads as before.

Not followed: a heading cleared (what named it shows `#REF`, and a
worked-out column reading it is refused, naming the columns there are),
two headings swapped in one paste, and Clojure of your own reading the
rows outside a step, `(map :work tasks)`: retype those yourself. A
heading retyped as one the table already has is numbered (`work-2`),
and followed as that.

## 20. Sort, filter and group

Sorting in place moves cells, which fights with references and spills.
Here the data stays where it is and a sorted, filtered or grouped *view*
of it is a spilling formula, so the view stays live and references to the
data keep their targets.

* **For ranges, Excel's functions:** `SORT`, `SORTBY`, `FILTER`, `UNIQUE`
  (§17), `GROUPBY` and `PIVOTBY`:

  ```clojure
  (GROUPBY A2:A50 C2:C50 SUM)           ; a row per region, its total, and a Total row
  (GROUPBY A1:A50 C1:D50 AVERAGE 3)     ; headers in row 1, shown; a column per value column
  (PIVOTBY A2:A50 B2:B50 C2:C50 SUM)    ; regions down, months across, with totals
  ```

  `GROUPBY row-fields values function [field-headers total-depth
  sort-order filter-array]` takes Excel's arguments: field-headers 0 none
  (the default), 1 the first row holds headers, 2 none but make some, 3
  headers, shown; total-depth 1 a Total row (default), 0 none, -1 at the
  top; sort-order n by the n-th result column, -n descending; filter-array
  TRUE for the rows to use. Keys compare as Excel's do (text ignoring
  case). `function` is any function of a column (`SUM`, `COUNTA`, a
  `LAMBDA` or `fn`). `PIVOTBY row-fields col-fields values function
  [field-headers row-total-depth row-sort-order col-total-depth
  col-sort-order filter-array]` likewise, one column each. With headers
  (§19) the columns are `orders>region`, `orders>qty`.
* **For tables (row maps, §19), the standard library's *Tables*
  section:** `(sort-rows orders :qty :desc)` (Excel's order: numbers, text
  ignoring case, TRUE / FALSE; empty values last either way; several keys
  as a vector), `(filter-rows orders :region "north")` (a value, a set, a
  function, or a function of the whole row), `(group-rows orders :region
  {:qty SUM :price AVERAGE})` (a row map per group; `[column f]` for
  another column; without aggregates, `:count`), and `(pivot-rows orders
  :region :month :qty SUM)` (a cross-table with a header row). `sort-rows`
  and `filter-rows` also take rows of a range, by column number.
  Clojure's `sort-by`, `filter` and `group-by` work too; the results
  spill with a header row, or, wrapped in `(as-grid …)`, fill a grid of
  their own as big as they are (§22).
* **Views as formulas, on the Shelf**. Select a block of data,
  or a whole column (its header) and right-click → *Functions, views,
  charts and templates…*, which selects the block of data around the
  column (the filled columns next to it, from its first filled row to its
  last) and opens the Shelf tab. Its *As a formula* section offers
  Excel's own, by the column picked there (the one clicked, else the
  active cell's): sorted A→Z and Z→A, filtered to the rows holding the
  column's first value (to edit in the formula), and grouped by it:

  ```clojure
  (VSTACK A1:D1 (SORT A2:D20 3 -1))                    ; sort, header row kept on top
  (VSTACK A1:D1 (FILTER A2:D20 (== C2:C20 "north") ""))   ; filter: edit the condition
  (GROUPBY C1:C20 (CHOOSECOLS A1:D20 2 4) SUM 3)        ; group: sums of the number columns
  ```

  Dates (§15) aren't summed by the group. A spilled result shows the
  values without the range's formats (as in Excel), so placing a sort
  or a filter gives its columns of dates their format where they spill.

  The first row counts as headers in a grid with headers, or when it's
  all text with other values below. Each tile shows its result; a click
  places the formula where you choose, as big as its result spills (the
  grid grows if it must). The view is live, and the data never moves,
  so references to it stay put. Beside them on the Shelf are the
  interactive views of the same data (§22) and its steps (§43).

## 21. Linked grid instances

Between `call` (a result; the template unchanged) and copying (an
independent copy) sits the *instance*: a cell that stores only its own
inputs, is drawn as the template is, and follows every later change to
the template's formulas, renderers and library — prototype inheritance
for grids. Edit the loan template once and every loan on the sheet
follows.

* **Making one.** Copy a cell holding a grid, right-click another cell →
  *Paste as linked instance*; or write the formula: `(instance $B$2)`, or
  with inputs of its own, `(instance $B$2 {:amount 35000 :years 10})`.
  Keys are cell names or addresses (`:B1`) in the template, as for
  `call` — name the inputs: address keys aren't rewritten when rows are
  inserted in the template (§5). The template's address is `$`-fixed, so
  copies of the instance keep it.
* **Drawn as the template.** An instance is drawn as the template's grid
  (`instance-view`, the default renderer for one), each cell with the
  instance's value through the template cell's renderer, number format and
  style; conditional rules test the instance's values. A line over it,
  *⧉ colstats · 1 own value*, names the template; its link opens the
  template itself, zoomed into, where it's defined (Back returns). The
  instance's own values are marked with a bar.
* **One per item.** A list of instances spills as a column of them,
  one card to a cell: `(map #(instance card {:x %}) [1 2 3])`, or `(for
  [[x] A2:A9] (instance card {:x x}))` for a card per row of a range.
  Each is drawn whole, its row given the room. The Shelf's *Repeat* on
  an instance writes this for you (§44). Type into a card's input and
  the value goes where the card's came from: the cell of the range the
  formula goes over (A15, for the second card of `(MAP A14:A16 …)`), or
  the item of the list written in the formula (`[1 2 3]` becomes
  `[1 7 3]`). An input the same for every card (`:years 3`) or worked
  out in the formula can't be typed into one card: you're told which
  formula to change.
* **Drawn clean.** `(instance-view :clean)` as an instance's renderer
  draws only its cells, their formats and the lines between them: no
  line over it, no frame, no bars marking its own values. Use it for an
  instance copied along a row, or one that's part of a finished page.
  Its inputs still take typing. The renderer gallery (*Show as…* on the instance's cell)
  offers it, and its refine panel's *Drawn* row switches between *as it
  shows*, *clean* and *with its ⧉ line*, `(instance-view :head)`. Cards spilled by one formula (*One per
  item*) are drawn clean as any spill is drawn: by a format rule over
  where they spill, `{:area "B10:Bde" :render "(instance-view :clean)"}`
  (to the grid's bottom edge, for a list of any length; §10). A renderer
  given to the formula's own cell would keep the whole list in it
  instead of a card to a cell.
  A template that presents itself (a display area, §6) has its
  instances drawn without the line by default, as its own cell is drawn
  clean: a column analysis repeated over ten columns shows only its
  results; `(instance-view :head)` puts the line back. Whichever way
  it's drawn, the Inspector's *Instance of* row names the template, with
  the same link and its count of own values.
* **Its inputs, in the Inspector.** Under *Instance of*, *Inputs* lists
  every input of the template, its parameters too (which the drawing
  doesn't show), each with what this instance gives it: blank shows the
  template's own value (`the template's: [1 2 3]`). Type a formula, as
  written in the instance's place (`A3:A7`, `d1:de`, `r1:re`,
  `(* 2 B3)`), or click a cell or drag out a range on the sheet while
  the field has the cursor (point mode, §5), and Enter: the overrides map is rewritten, `(instance
  $>colstats {:data d1:de})`, comments and the other entries kept;
  empty it to follow the template again. An instance whose overrides
  are computed (`(instance loan m)`) is changed in its formula.
* **By name.** A template held in a grid's margin is used by name from
  every grid inside it, `(instance colstats …)`, as a function is (§35).
* **From the shelf.** A template on the shelf can be placed as a linked
  instance (§4, *Templates on the shelf*): the sheet gets one copy of
  it, which its instances, and the shelf's changes, follow. Placed, the Inspector opens on it at
  its *Inputs*, and the message names those left at the template's own,
  which are flagged until given (§4, §26).
* **Changing its inputs.** The template's inputs are its parameters
  (§6): a cell of the template holding just a parameter's name
  (`amount`) is a field in the instance; everything else is drawn, not
  edited (a label row such as `cal`'s Mon … Sun, even one its formulas
  read, isn't an input). An override naming anything but a parameter is
  refused, saying so. Type a value and
  Enter, and the instance's overrides map is rewritten (`{:amount 15000
  :years 4}`), one undo step; empty the input (or type the template's
  value) to follow the template again. Controls in the template
  (`number-input`, `slider`, `dropdown` …) write the overrides the same
  way. An instance whose overrides are computed (`{:rate C5}`) is changed
  in its formula.
* **Reading it.** Its value is the template's value with the overrides
  (its display area alone, when it has one, §6: a column analysis gives
  its results, not its working cells):
  `(call inst :payment)` reads one cell (`(call grid :output)` reads a
  cell of any grid as it is), `(call inst {:years 3} :payment)`,
  `what-if` and `goal-seek` start from its overrides, and `(instance inst
  {:years 3})` is an instance of the template with both sets of
  overrides.
* **Detach instance** (cell menu) replaces it with an ordinary copy of
  the template, its own values typed into the copy's cells.

### Its own data

A template can hold tables of its own: a rota's working week, a
calendar's days off. Those are its **data**, and each instance keeps its
own rows of them, the template's rows being where it starts.

An instance draws each of its template's typed tables under the
template's headings, every cell an input, with a blank row last. Type
into it: the first row you type copies the template's rows into the
instance, a row past the last adds one, a row you empty goes, and each
is one undo step. What the template works out from the table (a total,
a calendar) follows the instance's rows; another instance of the same
template keeps its own, or the template's. The line over it says *its
own week*.

Change the template's table and each instance reads its own rows as the
template now has them, by heading: a column the template adds shows
blank, and is said (*none of its rows fill it yet: fill it in*); a
column it drops is kept but not read, and said; so is a whole table the
template no longer has. Nothing is changed for you.

A formula of the template that reads the whole table (`week`) sees the
instance's rows; one that reads its cells or a column of them
(`week>capacity`) reads the template's own cells. Write the first kind
in a template meant to have instances.

## 22. Grids made by formulas

A result laid out in a grid of its own: the formula's cell holds both the
formula and a grid as big as the result, which grows and shrinks with it.
Its cells are cells — select them, refer to them, format them, draw the
grid with `(grid-view)` — so a filtered or grouped table needs no
sub-grid sized by hand, no room below it to spill into, and keeps its
formatting as it changes size. Its cells hold values only; a grid of
formulas grows to fit its results instead (§38).

```clojure
;; in B6, named north: the North rows of orders, in a grid of their own
(as-grid (filter-rows orders :region "north"))
(SUM north>qty)                          ; elsewhere: its qty column
(as-grid (group-rows orders :region {:qty SUM}))
(as-grid (PIVOTBY orders>region orders>month orders>qty SUM))
(as-grid (for [i (range 5)] [i (* i i)]))   ; any rows
```

* **What it lays out:** a table of row maps, with a header row of its
  keys ("Row 1 is headers" ticked, so `north>qty` names a column); rows
  of values as they are; a sequence of single values as a column; a
  single value in a 1×1 grid. When nothing matches, a table keeps its
  header row.
* **What stays as it changes size:** everything that isn't a value —
  format rules (a `$` format on `E:E` covers new rows), renderers,
  styles, column widths, row 1 as headers, the cell's name and renderer
  (`(grid-view)` draws the grid in the cell). Another `as-grid` formula
  in the cell keeps them too.
* **What can't be changed:** its values — typing into it is refused,
  with a pointer to the formula, as you start (the formula bar shows 🔒
  for ƒ on its cells, and the first key gets the reason), not after
  you've typed — its size (the Inspector shows it, with no Resize), and
  its rows and columns (no inserting or deleting). Change the formula or
  its data instead. The same goes for the part of a grid there only
  while its results need the room (§38). **Except values read from a
  table**: a cell showing a value of a grid with headers, unchanged, can
  be typed over, and the typing goes into the table's cell (§48).
* **Convert to a plain grid** (cell menu, or the Inspector) writes the
  values into the grid's cells as literals and drops the formula: the
  grid, its formats and its cells stay, and are then edited as any grid.
* A formula without `as-grid` makes the grid go (the result spills, as
  §13), unless the cell is drawn with `(grid-view)` (below); clearing the cell takes the grid away; an error keeps the grid,
  emptied, and a reference into it (`plan>finish`, `plan>F2:F9`, `plan>F3`)
  is that error, as the cell is: a summary reading a table in error
  shows `#UP` from it, never a number worked out from blanks (a
  latest finish of 0, *0 tasks late*). `IFERROR` and `ISERROR` testing
  the reference catch it.
* Copying the cell copies the formula with the grid's formats; the copy
  makes its own grid.
* **No `as-grid` needed under `(grid-view)`.** A formula whose result is
  a table, or rows of values, drawn with `(grid-view)` is laid out in a
  grid of its own just the same, whatever made it (`bookings`, `filter`,
  a function of yours): `(bookings plan :task)` in a cell named booked,
  drawn with `(grid-view)`, is a grid, its columns `booked>from`, its
  width and format rules kept as it changes size, its formula live. Its
  columns start as wide as what they hold (a column of numbers or dates
  at least the usual width); drag them to change it. Take `(grid-view)`
  away and the result spills again (or stays in the cell under another
  renderer: `table` draws it as a read-only table). A grid of cells of
  its own (one you made, or a block's input) drawn with `(grid-view)`
  stays as it is.

### Views with their own controls

Filtering, sorting and grouping are interactive: you try a value, look,
try another. A *view* is a function whose result is a grid made by its
formula with controls in its header row, so that's done by clicking
rather than by editing a formula.

```clojure
(filter-view orders)        ; a drop-down and △ ▽ on every column's header
(filter-view orders {:region "North" :sort [:qty :desc]})   ; after two clicks
(filter-view orders {:search "pl" :page [2 25]})            ; searched, page 2 of 25 rows
(group-view orders {:by :region :qty :sum :price :average})
(group-view orders {:by [:region :item] :folded [["South"]]})   ; two levels, South folded
(pivot-view orders {:rows :region :cols :month :value :qty :f :sum :totals true})
(record-view orders {:at 3})  ; one row at a time, as a form
(profile-view orders)         ; a row per column: type, counts, least, greatest, histogram
(SUM north>qty)             ; the filter view named north: the rows it shows
```

* **`filter-view`**, Excel's filter: each column's header holds its
  name, △ ▽ to sort by it (the lit one, clicked again, stops sorting), and
  its values to tick, folded to a line saying which are (*(all)* when
  none is): click it to open them. Tick one or several to keep only their
  rows, untick them all to stop filtering by it; it stays open while you
  tick, and a click on the line folds it again. The values are those
  left by the other columns' filters, as in Excel. Values match as
  `filter-rows` matches them (text ignoring case). In the formula, one
  value is written as itself, several as a vector: `(filter-view orders
  {:region ["East" "North"]})`.
  * **Search**: the first column's header has a search box. Type and
    press Enter for the rows with that text in any column (ignoring
    case; a number matches as it's written, so `12` finds 112); × clears
    it. The columns' lists then offer what the search leaves. A date
    is a number in a table's rows, so the search doesn't find it as
    it's shown: `2026-11` finds no date (and `46` finds every date
    of 2026, as its serial number); filter the date column instead.
  * **Rows and pages**: the last column's header counts the rows (*12 of
    340 rows* when some are filtered out) and holds a drop-down of page
    sizes. Choose *25 a page* (or 50, 100, 500) to see the rows a page
    at a time, ◀ ▶ going from page to page (*26–50 of 340*); *all rows*
    shows every row again. On a page the view's value and its cells are
    the page's rows. A long table is also quicker to show a page at a
    time: only those rows are laid out.
  * An arrow with nowhere to go (the first page's ◀) is drawn faint, and
    its tooltip says why. A page that no longer exists (the rows
    shrank) shows the last one, saying so (*there's no page 9*).
  * A column named `sort`, `search` or `page` can be sorted by but not
    filtered here, as those names are the view's own options; its
    header says so. Filter it with steps.
* **`group-view`**: the first header picks the column to group by (a row
  per value, sorted); each other column's picks its summary: Sum,
  Average, Count, Min, Max, or — for none (the column stays, empty, so
  its drop-down can bring it back). Numbers are summed until chosen;
  dates (§15) aren't: a column of dates is — until chosen, and its
  drop-down offers Min, Max, Count and —, the Min and Max shown as dates.
  A Sum or Average of dates written into the formula (`{:start :sum}`)
  is an error saying what to choose.
  * **Several levels**: *then by…* under the grouping column adds
    another level, a column of its own: each region, say, with its row
    of summaries, and under it a row for each item in that region. A
    level's drop-down changes it, and its *(none)* takes it away (with
    the levels after it).
  * **Folding**: ▸ and ▾ beside each group. The groups of the last level
    unfold to their rows, the table's own rows under the group's
    summary; a group of an earlier level folds away the groups under it.
    Groups start with every level shown and no rows unfolded. Changing
    the levels starts the folding afresh.
  * A column named `by`, `folded` or `unfolded` can't be summarised
    here, as those names are the view's own options.
  * The view's value is the groups of the last level, a row each,
    whatever is folded; its cells are what it shows. So with rows
    unfolded or several levels, `(SUM gv>qty)` adds the summaries and
    the rows under them together: read the value, `(SUM (map :qty gv))`,
    to add the groups alone.
* **`pivot-view`**, a cross-table: a row per value of one column, a
  column per value of another (*(none)* for a single column), and in
  each cell a summary (Sum, Average, Count, Min, Max) of a third, with a
  Total row and column. Its column headings are data (the months), so its
  controls can't sit on them: they're a row of their own above the
  table — rows, columns, the summary and what it summarises, totals.
  Until chosen, the first text column goes down, the next across, and
  the first number column is summed (never a column of dates, §15);
  with no number column, it's counted (a column of names alone: how
  many rows have each), and so is a column of dates, whose summary
  offers Count, Min and Max (shown as dates). Dates down or across show
  as dates. Its value is the cross-table, header
  row first, without the controls (in the grid, the table starts at row
  2), so `(SUM pv)` and `pv>B3` work. Totals are the summary over all the
  rows in that row or column, not a sum of sums, so an average's total
  is right.
* **`record-view`**, one row at a time, as Excel's Data Form shows a
  table: a row for each column, its name beside the value, with ◀ ▶
  above to step through the rows (*3 of 40*) and a search box: type and
  press Enter, and ◀ ▶ step through the rows with that text in any
  column (*1 of 4 found*). Its value is the row shown, a map, so
  `(:qty rv)` reads its qty. A row number past the end (the table
  shrank) shows the last row, saying so; a search that finds nothing
  says so, and the fields are empty. **It edits**: a value read from a
  table's cell is typed over in place and written into that cell (§48),
  so it's a form for the table.
* **`profile-view`**, what's in each column of a table, as the Data
  tab's column statistics: a row per column with its type (whole
  numbers, numbers, text, true/false, mixed or empty), how many values
  it holds, how many are blank, how many are different (text ignoring
  case), its least and greatest (of its numbers, else of its text in
  sort order), and a small histogram: for a column of numbers, how many
  fall in each of up to 12 ranges; for any other, its commonest values,
  most first. Hover a bar for its range or value and its count. Its
  value is the profile, a row per column (`pv>blanks`). A column of
  dates (its kind, §19, or as-table's `:dates`) is *dates*, its least
  and greatest shown as dates.
* **`gantt-view`**, a Gantt chart of a table with a start and a finish
  date per row: a bar per row from its start to the end of its finish
  day, the names down the side, the dates across. Two drop-downs above
  it colour the bars by a column (resource, phase, status) and order them
  (as listed, or by start). More options are typed into the formula:

  ```clojure
  (gantt-view plan {:color :resource
                    :baseline [:baseline-start :baseline-finish]  ; the baseline's bars, thin and grey, beneath
                    :done :done                                   ; a column of 0–1: the part done, darker
                    :milestone {:days 0}                          ; milestones, as diamonds: here the rows of 0 days
                    :marker :deadline                             ; a tick on each row at its date, red when late
                    :line settings>status})                       ; a date as a red line: a status date, (TODAY)
  ```

  Hover a bar for its task, dates, baseline and how much is done.

  **Two versions compared.** `:baseline` can be any two date columns,
  so one version's bars can be drawn over another's: the Project plan's
  Baselines block draws the second baseline picked over the first.
  `:labels {:bars "Approved" :baseline "Original"}` names the two: the
  hover gives each one's dates under its name, and a line under the
  chart says which bars are which. A row with only the baseline's dates
  (a task gone since) is drawn as its grey bar alone.

  **Milestones.** `:milestone` names a column whose true values mark
  the milestones: true, a number other than 0, or *yes*, *y*, *x*
  (0, a blank and *no* don't). Or it gives a condition, as a filter
  does: `{:days 0}` marks the rows whose Days is 0 (`{:days [0 1]}`, 0
  or 1); a blank Days isn't 0.

  **Times of day.** A start or finish with a time draws to that time: a
  booking from 12:00 to 18:00 is six hours long. A finish that is a
  whole day, with a whole-day start, is the last day worked, so its bar
  runs to the end of that day; an all-day row among timed
  ones stays a day. A row from midnight to the next midnight reads as
  two whole days: give the day itself as its finish instead. With times,
  the dates across label hours as you zoom in, and the tooltip says
  them (*1 Mar 2027 12:00 – 18:00*).

  **Lanes.** `:lane :room` puts every row of the same room on one line,
  labelled with the room, each booking's name on its bar:

  ```clojure
  (gantt-view bookings {:lane :room :name :who :key :id :color :late})
  ```

  Bookings of one lane that overlap are stacked on extra lines under it,
  the time they overlap shaded red, and a line under the chart says
  which: *Overlapping in Blue: Smith and Jones, 1 Mar 2027 14:00 –
  16:00*. One ending as another starts doesn't overlap. A click still
  picks one booking, by its `:key` (or its name).

  **Markers.** `:marker :deadline` draws a tick on each row at the date
  in its Deadline column, `:marker [:deadline :promised]` one for each:
  red when the bar ends after it (late), grey when it's met. A deadline
  that is a whole day is met by a finish on that day. Hover a tick for
  *Build: deadline 12 Feb 2027, late by 3 days*.

  **Colours.** The colour drop-down offers every column. Text colours a
  value each; a column of true and false draws false grey and true in
  the alert red (light and dark themes each their own), so a *late*
  column reads at a glance; numbers draw from pale (low) to dark (high).
  Set colours of your own for values: `{:color :status :colors {"late"
  "#c0392b" "on time" "#9aa0a6"}}`; the other values keep theirs, and
  `{:color :late :colors {true "#d35400"}}` sets true's.

  **Dependencies.** Name the column of each row's ID with `:key` and the
  column of what each row waits for with `:depends-on`, and a line runs
  from each predecessor's finish down to the row and across to its
  start, ending in a small arrowhead:

  ```clojure
  (gantt-view tasks {:key :id :depends-on :after})
  ```

  The Shelf offers it written so for a table with an ID and an After
  column (a Gantt's rows: start and finish dates, a name, an ID, what
  each waits for, found by what the columns hold and their headings).

  What a row waits for is text of IDs separated by commas or semicolons
  (`2, 3`), a single ID, or a list of them. IDs are compared as text, so
  an ID typed as the number 3 matches `"3"` in a list. A line drawn red
  runs back: the row starts before what it waits for has finished. An
  ID that names no row (a typo, a row deleted) isn't skipped silently:
  a line under the chart says which row waits for what. A lag is read
  as the Schedule step reads it: `12-2d` and `3+2d` wait for 12 and 3
  (an ID that itself looks like one, `T-1d`, is that ID). A row allowed
  to start early by a lead (`12-2d`) isn't drawn red. A start-to-start
  link (`3SS`) runs from 3's start to the row's start; a finish-to-finish
  one (`3FF`) from 3's finish to the row's finish, its arrowhead pointing
  back into the bar's end. Hover a line for its type and lag.

  A table made by a Schedule step (§49) knows which column holds its
  links, so its chart draws them without being asked:
  `(gantt-view plan)` has the arrows, through any later steps that keep
  the ID and After columns. `:depends-on` names another column, and
  `{:depends-on false}` leaves the arrows out.

  **A column the table hasn't got is an error.** Every option naming a
  column (`:color`, `:start`, `:baseline`'s, `:done`, `:lane`,
  `:marker` …) is checked against the table: `{:color :phase}` on a
  table without a phase column shows *gantt-view: no column :phase to
  colour by; the table's columns are …* rather than drawing every bar
  one colour. So is an option it doesn't take, or one of the wrong kind:
  `{:lanes :room}` shows *gantt-view: there's no option :lanes (did you
  mean :lane?); its options are …*, and `{:line "2027-02-12"}` says
  `:line` is a date: `(DATE 2027 2 12)`, a cell holding one, or
  `(TODAY)`. Every view does the same (below).

  **Zooming and scrolling.** The third control above the chart says the
  dates shown and moves them: ◀ ▶ step back and forward by half the
  window, + narrows it (zoom in), − widens it, *all* shows the whole plan
  again. Each writes `:window [from to]` into the options, one undo step,
  so the view is kept with the sheet. Bars running past the window's
  edges are cut there. For a quick look that isn't kept, drag across the
  chart to move along it, and hold Shift with the wheel to zoom.

  **Picking a task.** Click a bar to pick its row: it's drawn strong and
  the others faint, and `:pick` is written into the options (the row's
  `:key`, such as `{:key :id}`; for a plan made by a Schedule step, its
  ID, or project and ID, `"Intranet:3"`, when the step's key is two
  columns; else its name). Click beside the bars to
  pick none. The view's value is what it picks: the row picked alone (a
  table of one), or no rows while none is, saying *Nothing picked:
  click a bar to pick its row*. So a cell beside it can show the
  details: `(record-view (picked tasks :id chart))` is the picked task's
  row of the Tasks sheet, to edit there (§48), and *Nothing picked …*
  while none is (§45, *A view's picked rows*). The Project plan's Gantt
  block does that under its chart. The
  columns are `:start`, `:finish` and, for the names, `:task`, `:name`
  or the first column of text, unless the options say others (`{:start
  :begin :finish :end :name :step}`); the Shelf offers it for any table
  with two date columns and writes the ones it found. Two rows of the
  same name are shown as `Review` and `Review (2)`. Only rows with both
  dates are charted (with `:baseline`, or both of the baseline's). Example: the tour's Tables section,
  and the Project plan example's Gantt block (§46).
* **The choices are kept in the formula**: its options map, its last
  argument, which the controls rewrite. Each choice is one undo step, is
  saved and copied with the cell, can be read in the formula bar, and
  can be typed. Nothing else holds state; the source never changes. A
  view written inside `if`, `when`, `let` and the like, `(if (seq
  orders) (record-view orders) "No orders yet.")`, has its own options
  rewritten, `(record-view orders {:search "Ana"})`; with two views in
  such a formula the controls can't tell which, and say so.
* **A column the table hasn't got is an error**, in every view: a
  filter or a sort of `filter-view`, a level or a summary of
  `group-view`, `pivot-view`'s rows, columns and values, `gantt-view`'s
  options, a chart's across, up, colour or size. The cell says which,
  and the columns there are: *filter-view: no column :regon to filter
  by; the table's columns are :item, :region, :qty*, rather than a
  filter keeping no rows or a chart drawn in one colour. A chart says it
  in its place. A view of
  your own does the same with one call: `(check-columns! "my-view" rows
  [["to sum" (:value opts)]])`.
* **An option a view doesn't take is an error**, as is one of the wrong
  kind, naming the options there are: *pivot-view: there's no option
  :row (did you mean :rows?)*, *record-view: :at is a number, not the
  text "2"*, and a summary `group-view` doesn't know (`{:qty :total}`)
  rather than none drawn. A view of your own does the same with
  `(check-options! "my-view" opts {:by :column :top :number})`: each
  option and its kind (`:column`, `:columns`, `:date`, `:number`,
  `:bool`, `:text`, `:map`, `:any`, a set of the values allowed such as
  `#{:asc :desc}`, or a vector of kinds such as `[:date :date]`).
* **The value** is the rows shown (a table), so `(SUM north>qty)` and
  `(count north)` read what the view shows; the dependency graph does
  the rest. Widths and formats are the grid's own, as for any generated
  grid; a renderer on the cell (`(grid-view)`) draws it in place.

### Panels: small UIs of your own

The views are built from a small kit any library function can use to
make its own UI, with buttons, drop-downs and the rest generated by the
program. Write the function in a grid's library (§14), call it from a
cell, and its controls rewrite the call's options map:

```clojure
(defn top-orders
  ([rows] (top-orders rows {}))
  ([rows {:keys [region n] :or {n 3} :as opts}]
   (panel opts
     (concat
       [(into ["Region"] (for [r (cons nil (sort (distinct (map :region rows))))]
                           (press (or r "All") [:region] r)))   ; a button per region
        ["Show" (press "−" [:n] #(max 1 (dec (or % 3)))) n
                (press "+" [:n] #(inc (or % 3)))]]
       (for [o (take n (sort-by :qty > (cond->> rows region
                                         (filter #(= region (:region %))))))]
         [(:item o) (:qty o)])))))
;; in a cell: (top-orders orders) → after two clicks (top-orders orders {:region "North" :n 4})
```

* **`(panel opts rows)`** lays `rows` out in a grid of their own
  (`as-grid`), row by row (rows may differ in length). Each cell is a
  value or a **gadget**:
  * `(input path control)` — any interactive renderer (`(dropdown …)`,
    `(slider 0 10)`, `toggle`, `text-input`, `(radio …)`) editing option
    `path`; its cell's value is the option's.
  * `(press label path v)` — a button setting option `path` to `v`, or to
    `(v current)` when `v` is a function (`(press "+" [:n] inc)`); path `[]`
    is the whole map (`(press "Reset" [] {})`). `(press label f)` runs `(f
    ctx)` instead, for anything else (`set-cell!` into the sheet).
  * `(drawn v renderer)` — value `v` drawn by any renderer: a chart of
    data the panel computed, `(drawn totals bar-list)`. Its cell's value
    is `v`.
  * `(beside …)` and `(stacked …)` — several items in one cell, side by
    side or one above another; the cell's value is the first item's.
  * `(with-style x style)` draws `x` with a style; `(with-value v x)` gives
    the cell value `v` (what references, and a header's column name, read).
* A third map: `:value` (the cell's value; else the cells' values, as
  rows), `:headers true` (row 1 names the columns), `:widths {0 120}` and
  `:heights {1 150}` (column widths and row heights until the grid has
  its own), `:spans {[1 0] [1 3]}` (cells drawn over a block, §4: a
  chart across the panel's width), and `:dates #{[2 1] 3}` (cells
  holding dates, by `[row column]` from 0, or a column's number for all
  of it below row 1: shown as dates where the grid gives them no number
  format, in the grid and in the Shelf's tile). `(date-column? rows k)`
  says whether a table's column holds dates (its kinds), and
  `(measure-column? rows k)` whether it holds numbers to add up: numbers,
  and not dates. The standard views use both.
* **Charts in panels** are a `drawn` cell spanning a row, under controls
  choosing what it shows (the demo's `order-chart`):

  ```clojure
  (panel o [[(stacked "by" (input [:by] (dropdown {:region "region" :month "month"})))
             (stacked "chart" (input [:kind] (dropdown {:bars "bars" :pie "pie"})))]
            [(drawn totals (if (= kind :pie) pie bar-list))]]
         {:value totals :spans {[1 0] [1 2]} :heights {1 150}})
  ```
* **The state is the options map** in the formula, nothing else: each
  click is one undo step, saved and copied with the cell, readable in the
  formula bar, and typeable. The function is pure — options in, cells
  out — so it's tested and reasoned about like any formula, and other
  cells can read its value.
* `filter-view` is a panel whose header cells are `(stacked (beside name
  sort-up sort-down) (input [k] (dropdown values)))`, and whose `:value`
  is the rows shown; `group-view` likewise.

**On the shelf.** A panel that takes a table as its first argument is a
view, and marked `^:view` it's offered by the Shelf tab for a table
(§4: dropped on a selection, in a new cell), next to the
standard views, which are marked so too:

```clojure
(defn ^:view gantt
  "Bars from start to finish …"
  {:caption "gantt"
   :needs {:start  {:kind :date}
           :finish {:kind :date}
           :name   {:kind :text :optional true}
           :link   {:kind :text :optional true :like "after|depends"}}}
  ([rows] (gantt rows {}))
  ([rows {:keys [start finish name link] :as opts}] (panel opts …)))
```

`:caption` names its tile (else its name). `:needs` says what columns
it takes, by what they hold, not by name: each entry is a slot, which
the Shelf gives a column of its own. So a view isn't even tried on a
table it can't use, which matters for one that takes a while to draw.

* `:kind` is what the column holds: `:date` (numbers shown as dates, by
  their number format), `:number` (any numbers but dates), `:integer`
  (whole numbers), `:text`, `:bool`, or `:any` (anything but an empty
  column); a set, `#{:text :number}`, takes either.
* `:optional true`: the view works without it. An optional slot is given
  a column only when one is left of its kind; with a `:like`, only one
  whose heading matches it.
* `:like` is a regular expression (or text, read as one, ignoring case)
  for the column's heading. For a slot that isn't optional it only says
  which column to prefer; otherwise the columns go in their order (the
  first date column is the start, the second the finish).

The columns found are written into the formula as the view's options,
so you can see and change them: `(gantt (as-table $A$1:$E$9) {:start
:begin :finish :end :name :task})`. They're ordinary options: called by
hand, the view takes any table and whatever options you give it. A view
without `:needs` is tried on any table, and offered when it gives a grid
without failing; the standard views are all of that kind. What a view
needs, and how its columns are chosen, isn't checked beyond that: that
a task starts before it finishes is the view's own business, said by
its failing. Its
tile is drawn from what it lays out, controls and charts included, and
its new cell is sized by it (`:heights` count). The tour's tables
section marks its `top-orders` and `order-chart`, each with `:needs`: a
column of text and one of numbers, preferring the ones headed *region*
and *qty*, so the Shelf offers them as `(order-chart orders {:by :region
:of :qty})`, and they work on any table with such columns.

## 23. Profiling and "why did this recompute"

Which cells are slow, and what a change set off: each formula's time when
it last ran, each renderer's when it last drew, and for the last change,
which cells re-ran, in what order, and why. Useful against per-cell time
limits (§8) and in big sheets (§9).

* The **Profile** tab. **Last change** sums it up — the cells it started
  from, how many cells re-ran, in how many passes, and how long the whole
  recompute took — then lists the cells that re-ran, in the order they
  ran, each with its time and why it ran:
  * *refers to B3, which changed* — a single reference; *(spilled from
    B1)* when B3 is part of a spill, which B1's formula filled;
  * *its range A1:A9 includes A4, which changed* — the cell in the range
    that changed first;
  * *refers to the grid at C5, which changed* — a grid value (§6), which
    re-ran because *a cell inside changed*;
  * *pass 2*: a spill grew (or a generated grid changed shape, §22), so
    the cells reading the new part ran in a further pass; *recompiled*:
    a formula referring into a generated grid that changed shape;
  * *a cell in its spill area changed* — typing into (or clearing) a cell
    a spill covers re-runs the spill's formula; *its spill was laid out
    again* — what blocked it went;
  * *changed* — the cells the edit started from; after loading a sheet,
    inserting or deleting rows or columns, or changing a library,
    *everything was recomputed*.
  The first 100 are listed, more on request. Each address is a link.
* **Slowest cells**: the 15 slowest formulas, by the time each took when it
  last ran (a cell that didn't re-run keeps its time). Only times of
  0.15 ms or more are kept: most cells take a few microseconds, and a cell
  faster than that says it *ran in under 0.15 ms* (a browser's clock
  counts in steps of 0.1 ms, so a fast cell can read 0.1 ms; the floor
  sits above that step) (the list says how many
  more took longer, the rest ran in under 0.15 ms). **Slowest
  renderers**: the 10 slowest, by how long each took to draw the value its
  cell holds now; renderers run as cells are drawn, so only cells that
  have been in view have a time, and **Refresh** shows the latest.
* **Overlay on the cells**: *Time* tints each formula cell by how long it
  took to run and to draw (a log scale up to the slowest cell, so one slow
  cell stands out; a cell under 0.15 ms stays pale), with a small tag
  giving the time where it's 0.1 ms or more; *Last change* numbers the cells the last change re-ran, #1, #2 …,
  and marks the cells it started from ✎. Hover a cell for the details.
* The **Inspector**'s *Profile* row says the same for the active cell:
  its time (or *ran in under 0.15 ms*), its renderer's, and whether the last change re-ran it, at
  which step and why. After a change that recomputed everything (a load,
  rows inserted, a library changed) it says just that.
* A `call`, `what-if` or `goal-seek` counts towards the cell that makes
  it: the grid's cells evaluated inside aren't listed.

## 24. The dependency map

How a sheet is put together, as a picture: the grid in view as a graph of
what reads what, inputs on the left and results on the right, with the
grids inside it as nodes to open. Trace arrows (§11) show one cell's
neighbours; the map shows a whole grid, and with grids as nodes, how the
grids of a sheet depend on each other.

* The **Map** tab. It draws the grid in view, so zooming into a grid (or
  out) changes the map; the breadcrumb above it opens the grids around it.
  ⇤ Expand gives it the window's width, and the map fits itself again
  whenever the panel changes size.
* **Nodes** are coloured by what they hold: values, formulas, **grids**
  (▦, double-click to open the grid's map, which is zooming into it),
  **another grid** (↗ `$` or `$>C3`: the cells outside this grid that it
  reads or that read it, one node per grid, double-click to open that
  grid's map) and **empty cells** a formula reads (dashed). A named cell
  shows its name (`grand-total D8`). A cell where an error starts has a red
  border, one showing an error passed on to it a dashed red one (§11).
* **Edges** run from what is read to what reads it; dashed when read
  through a range (`(SUM A1:A9)`), thicker for many references between
  the same two nodes. A cell inside a grid that is read from outside shows
  as its grid's node reading or being read: zoomed out, the edges between
  section grids are the references between sections.
* **The longest chain** of references (gold, apart from the errors' red;
  a node with an error stays red, the chain showing in its arrows;
  *Longest chain* turns it off), listed above the map with its steps and
  the time of its cells:
  the chain a change at its start has to run through, and where a sheet
  is deepest. A cycle counts once.
* **Time** tints each node by how long its cells took when they last ran
  (the Profile's times, §23, on a log scale up to the slowest node, so one
  slow node stands out), a grid node by everything inside it; *Kind*
  colours by what they hold. Cells under 0.15 ms have no time kept: a
  node says how many of its cells ran in under that ("1.2 ms, and 499
  cells under 0.15 ms each"), and the longest chain's time is then *at
  least* the sum shown.
* **Click** a node to light up everything it reads and everything that
  reads it, directly or not, and dim the rest; the box under the map gives
  its formula, its time, how many it reads and is read by, and a link to
  the cell (or to the grid's map). **Double-click** a cell's node to go to
  the cell. Clicking the background clears it. The node holding the active
  cell has a dashed ring that follows the selection.
* The wheel zooms around the pointer, dragging pans; +, − and Fit are in
  the corner. Fit shows the whole map; fitting by itself (opening the tab,
  another grid, resizing the panel) stops at a scale the labels can be read
  at, with the inputs' end in view, and the rest a drag away.
* **What's left out.** Text that nothing reads (headings, labels, notes)
  is never drawn; values and formulas that neither read nor are read are
  counted above the map and drawn with *Unconnected*. Grids and cells with
  errors are always drawn. On the demo's top grid that leaves the section
  grids and the cells reading them; its notes and headings are counted.
* **Big grids.** A grid with more than 300 filled cells is drawn by
  **column runs**: cells one under another in a column, of one kind (text
  apart from other values), as one node (`B2:B2501`), so a column of
  copied formulas reading the column beside it is two nodes and one edge
  (*Column runs* does this for a smaller grid too). At most 600 nodes are
  drawn, the least connected left out (and counted). A sheet with more than
  20,000 filled cells has its map drawn when the tab opens and on
  **Redraw**, not after every change, since drawing reads every cell: the
  map says when the sheet has changed since.

## 25. Units

A number can say what it measures: 5 km, 12.5 m/s, $40, 20 °C. Adding a
metre to a dollar is then flagged, multiplying and dividing make new
units (metres over seconds is m/s), and a kilometre plus 500 metres is
1.5 km.

* **Giving a value a unit.** Type it with the value: `5 km`, `12.5 m/s`,
  `$40`, `-3 kg`, `20 °C`. That unit belongs to the value, not to the
  cell: clearing the cell, or typing `7` or `3 h` over it, leaves no km
  behind.
* **Giving a cell a unit.** The format bar's **unit** field (beside
  Format) sets the unit of the selection; on a column, row or grid header
  it becomes a rule. The Rules tab takes `:unit` in a rule:
  `{:area "D:D" :unit "km"}`. Emptying the field, or *Clear formats*
  (cell menu), removes it.
  Typed, °F is `degF` and Δ°C `delta-degC`.
* **The unit popover.** Focusing or clicking the unit field opens a
  gallery of units, as the format field opens the format gallery:
  * a **search** (*fahrenheit*, *mile*, *speed*, *swiss*, by unit, its
    name and its kind; a unit's own text begun, too, so `$`, `€`, `°C`
    find theirs, as `CHF` does);
  * what's **typed in the search, read as a unit**, first, *As typed*:
    `$/m` is a tile, *$/m · money (USD) per length* (it says what a unit
    measures: its kind, else its parts), to give the
    cells; so a compound unit is written, not built. A search with a `/`,
    `*`, `·` or `^` is a unit being written, so it doesn't look for words
    (`$/m` doesn't list every unit with an m in it); one that isn't a
    unit says which part isn't (*“$/mx” isn't a unit here: mx isn't
    one*);
  * **your own units** first (below), then the units of the kind the
    active cell shows (a temperature's °C, °F, K; a change's Δ°C, Δ°F,
    K, never °F; a speed's km/h, m/s, mph …, and for a unit of parts
    such as $/h each part swapped: $/min, $/d), then every kind's usual
    units with their names (each group one kind, whose units convert to each other: current,
    voltage and resistance apart, and money labelled as currencies that
    don't; the cell's own kind isn't listed twice), and *No unit*;
  * each **tile** the active cell's value in that unit, when it's a
    number of that kind (*mi · mile · 3.11 mi*); a click gives the
    selection the unit and closes it. Typing in the field still works
    (Enter applies); Esc, or a click elsewhere, closes it.
* **What the number is.** A value is held in the **base unit** of its kind
  of quantity: metres, kilograms, kelvin, radians, amperes, moles, bytes,
  each currency as itself, and **days** for time. So `5 km` holds 5000, and
  `90 min` holds 0.0625. Days rather than seconds, because dates are serial
  days (§15): 10:30 plus `10 min` is 10:40. A plain number typed into a
  cell that has a unit is in that unit: 5 in a km column is 5 km (it holds
  5000), much as Excel reads 5 typed into a percent cell as 5%. Changing a
  cell's unit therefore *relabels* plain numbers (5 km becomes 5 mi), but
  *converts* numbers typed with a unit (a cell typed `5 km` shows 3.11 mi).
  A number typed with a unit of another kind than its cell's (`5 kg` in a
  km cell) shows in its own unit, and is flagged.
* **Where a unit shows.** The format bar's unit field shows the active
  cell's own unit (or its column's), with a formula's worked-out unit as
  its placeholder. The Inspector's Unit row says where the unit comes
  from (its own, a rule, typed with the number, worked out), and its
  Value row that the number is in base units. The status bar's sum and
  average are in the selection's unit when its numbers share a kind.
* **Showing.** A cell shows its value in its unit. The number format
  applies to the number in that unit, and the unit follows it: `1,234.50
  km`. A currency symbol goes in front: `$40`, `-€3.50`. The unit shown is
  the cell's own, else a rule's, else the unit worked out from its formula
  (below); a cell a spill reaches shows the anchor's. A cell whose number
  format shows a date or time doesn't add a unit, because the format already
  says what the number is.
* **Showing it, or not.** A column of CHF needn't say CHF in every cell,
  and a compound unit can be long. Whether it's beside the numbers is
  how they're shown, so it's with the number formats: the format
  popover's **Show … beside the numbers** switch (at its top, for a
  number with a unit; its tiles then drawn without it) turns it off for
  the selection (a whole column, row or grid as a
  rule, `{:area "B:B" :unit-shown false}`; a cell's own setting wins over
  its column's, so one cell can show it again): the numbers are still
  written in the unit (5 km in a km column is 5, not 5000), without it
  beside them, and the value keeps it (arithmetic, conversions, warnings,
  the Inspector, the unit field, `unit`). Only the unit the cell is meant
  to show goes: a value shown in another (7 EUR typed into a CHF column,
  flagged) keeps its unit, or a bare 7 would read as CHF. Where the unit
  is then said is the sheet's to lay out: a header cell can say it once
  with `unit`.
* **`(unit B2:B9)`** gives the unit the cells show, as text, `"CHF"`
  (`(str "Price (" (unit B2:B9) ")")` for a header): their own, a rule's,
  typed with the number or worked out from a formula; `""` when none has
  one; empty cells and text (a heading in the range) don't count. Cells
  showing **different units are an error, `#UNIT`**, naming each and
  where (*B2:B9 mixes units: CHF (B2:B7); EUR (B8)*), and counted and
  gone to as errors are: a header saying one unit for mixed cells would
  mislead, the error this exists to prevent. That includes units that
  convert (km beside m: as shown, with the units switched off, 5 and 300
  would sit side by side), and the message says so and how to fix it:
  *They're all lengths: give the cells one unit (the unit field) to show
  them all in it*: a column in m shows 5 km as 5000 m. A number without a
  unit beside numbers with one is mixed too. It takes a reference, not a
  value: a unit changed on the cells, or by a rule, recomputes it, as it
  recomputes what reads those cells. `(unit 5)` says it takes a
  reference. The Show switch reaches renderers as `(:unit-shown ctx)`,
  so every renderer that writes the cell's unit leaves it off.
* **Formulas compute with plain numbers**, so arithmetic is right whatever
  the units are: 1 km + 500 m is 5000 + 500 in metres, which shows as 1.5 km.
  The result's unit is worked out from the formula:
  * `+` and `-` keep the first operand's unit.
  * `*` and `/` combine units (km/h, m², $/h). Units of the same kind merge
    into the first one: km × m is km², km/h × s is km.
  * A result with no unit left is a plain number: 1 km / 500 m is 2.
  * Giving a formula cell a unit of the same kind shows its result in that
    unit (a speed worked out in m/s shown in km/h). A unit of another kind
    is a warning.
* **A sheet with no units doesn't work them out.** Until a sheet has a
  unit somewhere (a cell's or a rule's unit, a number typed with one, a
  formula giving one, such as `(with-unit A1 "km")` or `"5 km"`) or a
  date or time format, formulas don't work their units out, so they run
  faster (an edit reaching many cells takes 20–40% less). The first one
  makes the sheet work out every formula's unit, once, which takes a
  moment on a big sheet (under a second at 30,000 filled cells, about
  nine at 300,000); removing the last one does the same. Until then the
  Inspector's Unit row says *none* for every formula, where it would
  otherwise say *unknown* for one whose unit can't be worked out.
  Once a sheet has units, a formula works its unit out again only when
  it changes or something it reads changes kind (another unit, text,
  empty, an error, a unit or date format given or taken away): typing a
  new number of the same unit costs about what it would on a sheet with
  none. What an edit reaches through a spilled result or a grid made
  from a table still works its units out each time.
* **Units in formulas** are written as text, as dates are:
  `(+ A1 "10 min")`, `(> A1 "5 km")`, `(* B2 "$12/h")`. The operators and
  Excel's functions read quantity text as its value in base units, as they
  read `"$12"`, `"50%"` and `"2026-01-01"`. A plain number in a formula has
  no unit, so `(+ A1 1)` on a length adds one metre, and that's flagged.
  `(in-unit A1 "mi")` gives the plain number of miles in A1, and
  `(with-unit B1 "km")` takes a plain number B1 as kilometres.
* **Functions pass units through.** Every Excel function and every
  function of the standard library says how: `SUM` and `MAX` keep the
  unit of what they're given, `ROUND` and `INDEX` the first argument's,
  `COUNT` and `MATCH` give a plain number, `VAR.S` of lengths is in km².
  Dates: `WORKDAY`, `EDATE`, `EOMONTH` and `week-start` take a date and
  give a date, with a plain number of days or months beside it; `YEAR`,
  `WEEKDAY`, `NETWORKDAYS` and `DATEDIF` give plain numbers; `DATE` and
  `TODAY` make a date. Text functions (`LEFT`, `TEXT`, `str`) give text.
  Clojure's common ones (`max`, `inc`, `first`, `take`, `count`) do too.
  Giving a date function a length, say, is flagged.
* **Your own functions.** A function in a library has its unit worked
  out from its body, with its parameters standing for what it's given:
  `(defn twice [x] (* 2 x))` gives km when given km, and a mistake inside
  it (adding kg to x) is flagged on the cell calling it, naming the
  function. To say it yourself, give it `:units`, which wins:
  `(defn total {:units :same} [& xs] …)`; `:same`, `:plain`, `:text`,
  `{:from 0}` (the first argument's unit) and `:date` are the usual
  ones. A renderer or a view (`^:renderer`, `^:view`) gives what isn't
  a number, without saying. A body that can't be followed (one taking a
  function, as `map` does, a loop, a function calling itself) gives an
  unknown unit:
  nothing is shown or flagged, and the Inspector's Unit row names the
  function. A formula giving a table or a map says *not a number*
  there.
* **Completing units.** After a number in quantity text — in a string,
  `"10 delta`, or as a cell's whole formula, `20 fahr` — the formula bar
  offers units, grouped by kind as the unit dropdown's are, so Δ°C and °F
  needn't be typed: what's typed matches the start of a unit, of its
  other spellings (`delta-degC`, `degF`, `ohm`) or of a word of its name
  (fahrenheit, change, hour), its own spelling first. Tab or Enter on one
  takes it. After a `/` or `*` (`"12 $/hou`) only units of one part are
  offered, for that part. Nothing pops up for other strings, or while a
  number is being typed.
* **Points and differences.** An absolute temperature (°C, °F) and a date
  are *points* on a scale. A temperature change (`Δ°C`, `Δ°F`, or K) and a
  duration (`10 min`, `3 d`) are *differences*.
  * A point plus a difference is a point: 20 °C + 10 Δ°C is 30 °C (raised by
    10 degrees), and 10:30 + 10 min is 10:40.
  * A point minus a point is a difference: 30 °C − 20 °C is 10 Δ°C, and a
    date minus a date is in days.
  * Adding two points, or multiplying or dividing a point, is flagged.

  `"10 °C"` always means the temperature 10 °C. A rise of ten degrees is
  `"10 Δ°C"`, typed `delta-degC` (and `delta-degF`; `degC` and `degF` are
  °C and °F). A cell whose number format shows a date (with or without a
  time) is a point in time. A time-of-day format (`hh:mm`) counts as a
  duration, since Excel sheets use it for hours worked.
* **Warnings, not errors.** The value is still a correct number, so it
  stays: dependents see it, and nothing turns into an error. The cell gets
  an amber corner (bottom left: a name has the top left, a failed
  validation the top right), hovering says why, the Inspector's Unit row
  lists the warnings, and the status bar's "⚠ n units" goes to each
  flagged cell (Shift-click: back). The warnings are:
  * adding or comparing different kinds: km and $, km and kg;
  * adding a plain number to a quantity (except 0, or an empty cell);
  * adding two points, or multiplying a point;
  * a plain number expected but a quantity given (`EXP`, `LN`, `CONVERT`);
  * a date expected but another quantity given (`WORKDAY` on a length),
    or a quantity other than a time given as a number of days;
  * a range given to `SUM`, `AVERAGE`, `MAX` … that mixes kinds;
  * a formula giving m/s in a cell whose unit is kg;
  * a formula giving a plain number in a cell whose unit isn't scale 1
    (it's taken as base units: 6 shows as 0.006 km in a km cell);
  * a number typed with a unit of another kind than its cell's;
  * a value with a unit in a cell whose number format shows money
    (`$#,##0`), when the unit isn't that money: a rate, `56.25 USD/h`,
    typed into a column of day rates shown as dollars is flagged *a rate
    where an amount is shown*, and so is a formula giving such a rate in
    a cell shown as dollars (an hour is a 24th of a day, so 56.25 USD/h
    is 1,350 a day: a working hour is a unit of your own, *New unit…*);
  * a number typed with a unit of another kind than the values typed
    near it in its column (two or more, all of one kind): `56.25 USD/h`
    among `$600`, `$700` is *a rate among amounts*.
* **Currencies don't convert.** Each currency is a kind of its own, so
  $ + € is flagged, because no rate is fixed. A rate is a value in $/€:
  A1 in € times B1 in $/€ is in $.
* **The units.** Combine them with `/`, `*`, `·` or a space, with powers
  as `^2`, `²` or `m2`: `m/s^2`, `kg·m/s²`, `N m`, `$/h`, `1/s`.

  | kind | units |
  |---|---|
  | length | m km cm mm µm nm, in ft yd mi nmi |
  | mass | g kg mg µg t, lb oz |
  | time | s ms µs ns min h d wk yr |
  | temperature | K, °C °F (degC degF), Δ°C Δ°F (delta-degC delta-degF) |
  | area, volume | ha acre, L mL cL dL (or l ml …) gal |
  | speed | mph kn (and any length / time: km/h, m/s) |
  | force, energy, power | N kN lbf, J kJ MJ cal kcal Wh kWh MWh, W mW kW MW GW hp |
  | pressure | Pa hPa kPa MPa bar mbar atm psi |
  | electrical, frequency | A mA V mV kV Ω kΩ (ohm), Hz kHz MHz GHz |
  | amount, data | mol, B kB MB GB TB KiB MiB GiB bit |
  | angle | rad deg ° |
  | currency | $ € £ ¥, and USD EUR GBP JPY CHF CAD AUD NZD CNY INR SEK NOK DKK |

  An angle is a kind of its own (radians at base), so `(SIN A1)` on
  `30 deg` is 0.5. There's no "C" (coulomb, or Excel's Celsius in
  `CONVERT`), so `20 C` isn't quietly one or the other.
* **Units of your own** are defined in the unit popover: **+ New unit…**
  asks for a name, what it is (*equals* a number in units already known,
  `201.168 m`, or *a kind of its own*) and the SI prefixes it takes (`k`,
  `kM`); each of your units is listed with what it is, to **Edit** or
  remove (**✕**; refused while another unit is defined in it: crew in
  person). A mistake is said in the form, checked as the engine would: a name already a unit, a definition in unknown
  units, a circle. Each change rewrites the `custom-units` of the library
  whose map is in effect for the active cell (the nearest grid's that
  defines one, else My library), a unit a line, by name (comments inside
  the map aren't kept; one that isn't a map written out is left to edit
  by hand); when none defines any yet, the form asks: this sheet (the top
  grid's library) or My library. One undoable step, as any library edit.
  Underneath, it's a setting in a library: the top grid's for the whole sheet (saved with it), My
  library for every sheet, or a grid's for that grid and the grids inside
  it, the nearest definition winning, whole, as it does for functions:

  ```clojure
  (def custom-units {"furlong"   "201.168 m"
                     "fortnight" "14 d"
                     "person"    :kind
                     "crew"      "4 person"
                     "pace"      {:as "0.75 m" :prefixes "k"}})
  ```

  A unit is a number in units already known (they can use each other,
  in any order), or `:kind`: a kind of quantity of its own, which
  converts to nothing else, so `$/person` works and adding people to
  dollars is flagged. `:prefixes` lets it take SI prefixes (`2 kpace`).
  Refused, as the library's error (prefixed "units:"): a name that is
  already a unit (`m`), that isn't one word or ends in a digit (`2x`), a
  definition in unknown units or going round in a circle, a temperature on
  a scale (`"20 °C"`), and a plain number (`"12"`: make `dozen` a `:kind`
  of its own, or leave it a number). The good definitions still apply. `5
  apples`, typed with no such unit, says apples isn't a unit here.
* **Preferred units.** A worked-out result shows in the first unit of
  its kind in the formula (240 km + 1000 m is km, as most unit-aware tools
  do). A cell's own unit (the unit field) shows it in another. For a grid,
  the grids inside it, or (in the top grid's library) the whole sheet,
  another setting, defined where custom-units are:

  ```clojure
  (def preferred-units ["m" "s" "°F"])
  ```

  A result shows in the preferred unit of its kind; else each unit in it
  of a base kind is replaced by the preferred one (km/h becomes m/s with
  `"m"` and `"s"`; `"km/h"` itself in the list wins over that). A
  temperature shows in a preferred scale, and a change of temperature in
  its Δ (Δ°F), never with the scale's offset; K serves both. It applies
  when shown, so it recomputes nothing, and the Inspector says "worked
  out as km · shown in m (preferred-units)". Numbers typed with a unit
  show as typed, and a cell's own unit or a rule's always wins.
* **Copying and pasting** carry units: a typed `240 km` is its source,
  and Paste values writes a number in the unit it shows (`240 km`, `$40`),
  never the base number it holds. The last copy is also kept in the
  browser's storage, so another tab (another sheet) pastes it as cells —
  a grid with its formats, rules and library, formulas, units — and not
  only as text, when the system clipboard still holds what that copy put
  there. Text from other apps reads `240 km` and `$40` as quantities, as
  typed.
* **Excel's `CONVERT`** works on plain numbers, as in Excel:
  `(CONVERT 1 "mi" "km")` is 1.609…. It knows these units and Excel's own
  names for them (`"hr"`, `"mn"`, `"sec"`, `"lbm"`, `"C"`, `"F"`, …).

## 26. Formula checks

A formula can compute without error and still be wrong: a copy that
reads the row above, a number typed over a formula, a rate written into
one formula of many. Excel marks such cells with a green triangle and
leaves them alone. RecDep does the same: the checks only point, they
never change a value or stop a formula.

* **The marks.** A cell worth a second look gets a green corner, top left
  (a name's blue corner is drawn over it, smaller). Hover it for why; the
  Inspector's **Check** row says the same, one line per check.
* **What is checked.**
  * **A call with the wrong number of arguments** to one of Clojure's own
    functions: `(inc 1 2)` still shows 2 (the extra argument is ignored),
    but is marked "inc takes (inc x), not 2 arguments". The same mistake
    with Excel's, RecDep's or a library's function is an error instead
    (§3, *Argument counts*).
  * **Unlike its neighbours.** The cells on both sides of it, along a
    column or a row, hold copies of one formula, and it holds another
    formula, or a number (typed over one?). Copies are recognised by
    their shape: the formula with each relative reference written as an
    offset from its cell, so `(* B3 C3)` in D3 and `(* B4 C4)` in D4 are
    the same, and `(* B5 C4)` in D5 isn't. Only a cell with copies on
    *both* sides is compared, so a total under a column, or the first
    row, is the end of a run and not an odd one out. A cell that belongs
    to a run going the other way isn't flagged either: a formula matching
    a neighbour across it (a column consistent with itself, between two
    columns that happen to share a shape), or a number beside another
    number lying between the same copies (a column of inputs between two
    copied columns).
  * **A number stored as text.** A text cell holding `"1,250"`, `"42"`,
    `"-3.5"` or `"1e6"`: SUM leaves it out and comparisons treat it as
    text. Not `"007"` (a code, kept as text on purpose), `"+1"` or `"12%"`
    (a button's label, a format code), nor a headed grid's headers, which
    are names.
  * **A number written into a formula** that also reads cells:
    `(* D9 1.2)`. A rate or a limit kept in a formula is easy to miss and
    to leave stale in some copies; a cell of its own shows it and changes
    it in one place. Not counted: small numbers (−10 to 10, 12, 24, 60,
    100, 365, 1000: indexes, halves, months, percentages), numbers in
    maps, vectors and sets (options, styles, a lookup table, hiccup),
    the arguments of `date` and `time`, and the inputs given to `call`,
    `instance`, `what-if` and `goal-seek`, which are what those are for.
    Of a run of copies only the first (no copy above it or to its left)
    is marked, not every copy.
  * **An input nothing reads.** A named cell holding a constant (a number,
    a quantity, true or false) that no formula reads: not by a reference,
    a range around it, or a reference to a grid holding it, nor a
    validation rule (`(between min-qty max-qty)`, §16). Unnamed
    constants aren't inputs: most are data, read or not.
* **Big grids.** A grid with more than 20,000 cells is checked for its
  names only (inputs nothing reads, names hiding functions), not cell by
  cell; the count's tooltip says so when one is. Grids inside it are
  checked as usual.
  * **A name hiding a function.** A cell named like a function its grid's
    formulas could call (a library's, the standard library's, core's): in
    those formulas the name means the cell. Naming a cell so is refused
    (§2), but a library function defined after the name, or a file, can
    still make one.
  * **Out of a self-contained grid.** A reference out of a block is
    refused where it's written (§6), not flagged. Flagged here is what
    formulas don't say as references: a button or control writing out of
    it, where the address is written in (`(set-cell! ctx "..>A1" 0)`:
    refused when pressed, flagged before), and a note's links out of it
    (`[[$>basics]]`), which a saved grid breaks. (Parameters get this
    check, *an input nothing reads* and *a name hiding a function*, not
    the cell-by-cell ones: they have no neighbours.)
  * **A reference to another grid's private name.** A formula reading
    `cal>_first-day` from outside `cal`: a name starting with `_` is
    private to its grid (§6), no part of what it gives, so free to change
    or go. From inside the grid, or a grid inside it, it's an ordinary
    name. The fix is to read an output instead (or make the name public).
  * **A template hiding another of its name** (§35). A template a grid
    keeps in its margin, of the same name as one around it: its grid
    and the grids inside use it, and changes to the one around don't
    reach them. The mark is on the parameter's chip in the strip, and the
    ◤ link opens the template, whose bar says the same. *Use $'s* (beside
    it) drops it; *Keep this one* keeps it on purpose, unflagged. A copy
    kept on purpose (*Keep a copy here*, or a version Insert kept) isn't
    flagged.
  * **A test outside a tests grid** (§27). A cell whose formula calls an
    assertion, anywhere but in a grid's tests grid. The fix offered is
    *Move it there*: into its grid's tests grid (made, if the grid has
    none), its references written to mean the same from there (`total`
    becomes `..>total`). It's refused, saying why, when a formula reads
    the test (move it by hand), or when a reference wouldn't mean the
    same (one relative to where the test is, such as `u1`).
  * **A table shaped in code that a table step does** (§43). A formula
    grouping a table it reads, `(group-by :phase tasks)` or
    `(frequencies (map :phase tasks))`, choosing its columns, `(map
    #(select-keys % [:id :finish]) plan)`, or sorting it by a column,
    `(sort-by :finish plan)`: the check names the step that does it,
    `(steps tasks [{:op :group :by [:phase] :aggs [{:fn :count}]}])`,
    which the Inspector shows as a chip to read and change with the Data
    tab's forms. Filters and sums over a table are plain enough, and
    aren't marked; nor is code in a library.
  * **A placed template's inputs left at its own values** (§4). A
    template dropped from the shelf, as a copy or a linked instance,
    whose inputs nothing in the selection wired: they still hold the
    values its author built it with. The Inspector's *Inputs* marks each
    one ◤; giving it (a binding, an override, a copy's own value
    changed) clears it, and *Ignore: keep the template's* keeps them. A
    cell drawn over a block (a grid drawn in place, an instance, a
    chart) shows the corner over what it draws.
* **Counts.** The status bar's `◤ n to check` counts the marked cells in
  the whole sheet (`◤ 1 of 6` on one of them), and the Outline gives each
  grid a `◤ n` (folded, the grids inside it too). Both are links that
  step through the cells as the error counts do (§11): a click goes to
  the next after the active cell, Shift-click the previous, zooming as
  needed. Their tooltip says what the warnings are ("2 formulas differ
  from the ones beside them, 1 number is stored as text").
* **Ignoring.** The Inspector's **Ignore** beside a check silences it for
  that cell (an undoable change, saved with the sheet); the cell then
  lists it as ignored, with **Show again**. A few formulas of the tour
  keep numbers in them on purpose and are ignored that way. The tour's
  Formula checks section shows each check.

## 27. Tests and assertions

A sheet of formulas has no tests: change a template and nothing says
that a total it gave last week is now wrong. RecDep's tests are cells
that check the sheet, among them test cases calling a grid used as a function;
both re-run with every change, like any formula, and a Tests tab lists
them.

* **A test is a cell.** `(assert= B4 1000)` passes while B4 is 1000 and
  shows ✓; when it isn't, ✗ with what it got and what it expected ("got
  990, expected 1000"). A message says what's tested: `(assert= B4 1000
  "no interest")` shows "no interest: got 990, expected 1000".
  * `(assert= actual expected msg?)` — equal, as Clojure's `=`, except
    that numbers equal but for rounding (`(+ 0.1 0.2)` and `0.3`, to a
    billionth) are equal, also inside vectors and maps; a date or a
    quantity can be written as text, as arithmetic reads it (`(assert=
    A1 "2026-11-07")`, `(assert= B4 "5 km")`; a plain number as text,
    `"42"`, stays text); an Excel error value equals the same error
    (`(assert= (/ 1 0) (ERROR "#DIV/0!"))`).
  * `(assert-near actual expected tolerance msg?)` — within a tolerance.
  * `(assert-true x msg?)` — any condition; false, nil and an Excel error
    value fail.
  * A vector of results is one test: `(mapv (fn [[r p]] (assert-near
    (call Loan {:rate r} :payment) p 0.01)) cases)` shows ✓ 3, or how
    many fail and the first one's why. A result that spills is a test
    per cell.
* **Testing a grid used as a function** (§7): call it. `(assert= (call
  $E$32 {:rate 0} :total) 1000)`, copied down a table of inputs and
  expected outputs, is a table of tests.
* **Test cases**. `(assert-case .. {:rate 0} {:total 1000}
  "no interest")`, in a grid's tests grid, calls the grid with those
  inputs as `call` does, and passes when each expected value is what it
  gives (as `assert=` compares). Inputs left out keep the grid's own
  values; inputs bound from outside take their defaults, so a case tests
  the grid itself. Failing, it says which values differ: *no interest:
  total: got 1050, expected 1000*. It's a test like any other: listed,
  counted, stepped to, carried with the grid.
  * In the Tests tab, with the grid's cell (or a cell inside it)
    selected, **+ Record a case** writes one into the grid's tests grid
    (making it if need be), in the first free row of column A. It takes
    the grid's inputs as they are — its parameters (with none, every
    number or true/false typed in) — and what
    its formulas give with them (the named ones; with none named, all).
    Private names (`_first-day`, §6) are neither: no part of what the
    grid gives.
  * Renaming a cell rewrites the cases' keys. Address keys (`:B2`) don't
    move with inserted rows, as `call`'s don't: name the cells.
  * The top grid can have cases too, in its own tests grid.
  * **The cases are a table**, as the Scenarios tab's (§32): a
    column per case (each an `assert-case` cell, rewritten as you edit;
    ↗ goes to it), then a row per input and a row per expected value.
    Widen the panel with Expand when cases are many.
  * **Values are written as typed into the cell:** `0.05`, `5 km`,
    `2026-11-07`, `"north"` (text in quotes), `:done`, `[1 2]`. A value
    shows as its cell shows it (`$20,000.00`) when that reads back the
    same. A bare word, a formula, or a value the cell couldn't hold is
    refused, saying why.
  * **Empty means:** for an input, the grid's own value; for an
    expected value, not checked. A case must check something, so
    emptying its last expected value is refused.
  * **The keys are the Scenarios tab's:** Enter (or ↓) keeps a value and
    moves down the case's column, Shift-Enter (or ↑) up, Tab across;
    Escape puts it back. A refused value keeps its text, to fix.
  * **Rows:** *+ Input: rate* (*+ Expected: total*) adds the active
    cell, when it's a cell of the grid; or type a name or address beside
    it (Enter). A cell that has a row already is refused. ✕ on a row
    takes the key out of every case. A row marked ⚠ names a cell the
    grid no longer has.
  * **A case's column** has its name (click to change it), ✓ or ✗, ↗,
    ⧉ (a copy, in a new cell, to change an input and pin down another
    behaviour) and ✕ (clears its cell). A failing case offers **Accept these results**,
    which makes what the grid gives now its expected values. A value
    that differs is tinted, with what the grid gave (`got $371.23`).
* **A grid's tests grid.** A grid's tests belong in its **tests grid**,
  a grid kept in its margin (as templates are, §35), named `_tests`:
  * **It travels with the grid:** copies and paste, Save grid and Insert
    carry it, and a copy's tests test the copy. A grid saved alone is its
    file's top grid, and its tests still run there.
  * **The top grid's tests grid**. No cell holds the top grid, so
    a test there that uses `..` depends on no cell: it runs again after
    every change outside its tests grid, once the rest has settled. One
    that only calls it (`call`, `assert-case`, what-if… with `..` as the
    grid) is cheap: about 2 ms an edit in a sheet of 30,000 cells. One
    that reads its whole value (`(count ..)`, `(:total ..)`) builds the
    top grid's value map each time: about 130 ms an edit there. The
    value leaves out the top grid's cells in error, rather than being an
    error itself, so a call still works.
  * **Calls of the grid leave it out:** `call`, what-if, goal seek,
    simulations and instances don't run its tests, and its tests
    calling the grid don't run each other.
  * **In it, `..` is the grid it tests:** `(assert= (call .. {:rate 0}
    :total) 1000)` calls it with other inputs, and `(assert= ..>total
    1102.5)` reads a cell of it as it is.
  * **Making one:** the Tests tab's **+ Tests grid** (for the grid the
    active cell is in), or **+ tests grid** in the grid's ⚙ Parameters
    strip, makes one and opens it. Once made, the tab's *Open _tests*
    and the strip's **✓ tests grid** chip open it, and × removes it,
    asked first.
  * **Tests anywhere else are flagged** (§26), with *Move it there*.
  * **By path** a test is `loan>B2>_tests>A1`, and a note can link to it
    (`[[$>loan>B2>_tests>A1|its tests]]`); typed into the name box, the
    path goes there.
  * `_tests` is a private name (§6): no part of the grid's interface,
    not recorded by a case, not offered to formulas outside.
* **Results are values, not errors.** A failing test doesn't stop
  anything: formulas reading it go on computing, and a template holding
  a failing test still works for every call, what-if and instance. Read
  one with `(passed? B4)`; `test-outcome` gives `{:pass? :n :failed
  :msg}` for renderers of your own. A test whose formula has an error
  (`#UP` from a broken cell it checks, a `#NAME`) is listed as an error,
  not dropped.
* **The Tests tab** first counts the tests in the whole sheet (failing,
  erroring, passing) and lists the failing ones with why. Then, scoped
  to the grid the active cell is in (or holds), as the Inspector is: its
  tests grid (*Open*, or *+ Tests grid*), the tests in it (failing ones
  listed, passing ones folded; a tests grid's tests among them), and its
  cases. Each test is named by its path, as a formula may write it
  (`loan>B2>_tests>A4`, `tests>B5`), a link to its cell (a case's, to its
  grid). In a tests grid, the tab shows the grid it tests. The **status bar** counts failures (`✗ 2 tests
  fail`; `✗ 1 of 2` on one of them), a link stepping through them as the
  error counts do (§11); when every test passes it reads `✓ 9 tests`.
  A failing test case has no cell of its own, so stepping to it selects
  its grid (the top grid's: zooms out to it) and shows the Tests tab,
  where its cases are; the status bar says which fails and why (*loan>B2's
  test case “as it is” (payment: got 377.42, expected 1) fails*).
  **Find** reads a result as `#PASS` or `#FAIL` and why, so `#FAIL`
  finds the failing cells.
* Paste values turns a result into the formula that makes it, with the
  values it had (`(assert= 990 1000)`): still a test, frozen.
* Formula checks (§26) don't flag a test's expected values as numbers to
  put in a cell of their own: in a test that's where they belong.
* The tour's Tests section shows each; one test there fails on
  purpose, and the loan template (Grids as functions) carries two cases.

## 28. Cells that update over time

RecDep computes a cell when its formula or its inputs change, so `NOW`,
`TODAY` and `RAND` keep the value they had when their formula was last
computed. F9 computes them again, `(tick 1000)` makes a cell run on a
timer, and `(history B5 50)` keeps a cell's last values: with interactive
renderers and sparklines, live dashboards.

* **F9 recalculates the volatile cells**: those whose formula uses
  `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `tick`, or
  Clojure's `rand`, `rand-int`, `rand-nth`, `shuffle`, `random-uuid`,
  and everything that depends on them, as Excel's F9 does. **↻ F9** in
  the status bar does the same (it shows while the sheet has such a
  cell; its title says how many), for keyboards where F9 is the
  system's. Nothing else is computed: RecDep recomputes what changed, so
  there's nothing else stale.
* **`(tick ms)`** is the date and time now, to the millisecond, and the
  cell runs again every `ms` milliseconds (at least 100), with everything
  that reads it. Typed in, it gets the format `hh:mm:ss`, so `(tick
  1000)` is a clock. Cells reading the clock follow it: `(- clock start)`
  counts up, a formula with `RAND` reading it is a new sample every
  second, a chart of them moves.
  * `(when running (tick 1000))` ticks only while the cell `running` is
    true (a `toggle`, say): a tick that isn't called doesn't run. It
    gets `hh:mm:ss` too. With
    several ticks in a formula the shortest wins.
  * **The status bar** shows `⏸ 2 live` while cells tick; a click pauses
    every timer in the sheet (`▶ 2 live, paused`), another resumes them.
    Timers also stop while the page is hidden (another tab), and catch up
    when it shows.
* **`(history x n)`** is the last `n` values of `x`, oldest first: `(history
  B5 50)` adds B5's value each time B5 changes — a tick, F9, an edit —
  so `sparkline` or a chart of the cell shows how B5 moved, and `(AVERAGE
  (history B5 10))` is a moving average. It passes on B5's unit.
  * A sample is taken when a cell the formula refers to is computed
    again, even to the same value (a sensor reading 20 twice is two
    readings); typing the same value again into B5 isn't a change, so
    takes none. Changing the formula around it (`n`, or `(vec (history
    B5 50))`) keeps the values and adds none; a formula given another
    input as well takes a sample only when an input changes. A formula
    referring to no single cell, `(history (RAND) 10)`, takes one every
    time it runs.
  * An error upstream doesn't lose the values: the cell shows the error
    and carries them, and the next value goes on the end.
  * Of a range, each value kept is the range, a 2-D vector, so the
    result spills as a table of them; keep what's wanted instead,
    `(history (SUM A1:A3) 20)`. (With only ranges to read, a sample is
    taken each time the formula runs: when a cell in them changes.)
  * In a cell of its own, the values show as a vector: draw them with
    `sparkline`, `bars` or a chart.
* **Not changes.** Ticks and F9 change values only: they aren't undo
  steps, don't autosave, and don't replace the Profile tab's last change
  (§23). Undoing an edit brings back the values as they were then,
  histories included; the next tick goes on from there.
* `tick` and `history` must be written in the cell's own formula (not in
  a library function a cell calls, nor a renderer), and formula checks
  (§26) don't flag their intervals and lengths as constants to name.
* The tour's Cells that update over time has a clock, a pretend sensor with its history
  as a sparkline and a moving average, and a die for F9; its clock is off
  until switched on, so the tour doesn't tick for everyone.

## 29. Uncertain inputs (Monte Carlo)

A model's inputs are often estimates: about 20, give or take 2. A cell can
say so, `(normal 20 2)`, and `(simulate profit 2000)` works `profit`
out 2000 times, each time with a fresh draw from every estimate it
depends on: the spread of results, drawn as a histogram, instead of one
number. What-if (§18) and grids as functions (§7) re-evaluate part of the
sheet with other inputs; this does the same with drawn ones.

```clojure
(triangular 800 1000 1500)            ; units sold: shows 1100, its mean
(normal 20 2)                         ; price
(* (chance 0.1) 3000)                 ; a 10% risk of a 3000 cost: shows 300
(- (* units (- price unit-cost)) overheads risk)  ; profit: a formula
(simulate profit 2000)                ; 2000 runs, drawn as a histogram
(simulate (* units price))            ; an expression, 1000 runs
(simulate profit 2000 {:seed 7})      ; other draws
(chance-of neg? runs)                 ; the chance of a loss, 0–1
(:p95 (stats runs))                   ; 95% of runs come out below this
```

* **Distributions** are functions in a cell's formula: `(normal mean
  sd)`, `(lognormal mean sd)` (of the value itself, not its log),
  `(uniform lo hi)`, `(triangular lo most-likely hi)`, `(pert lo
  most-likely hi)` (the smooth three-point estimate), `(chance p)` (1 or
  0) and `(discrete values)` / `(discrete values weights)` (values can be
  a range). Parameters can be cells, and a formula can combine several,
  `(* B2 (normal 1 0.1))`.
* **On the sheet each is its mean**, so the model reads as a deterministic
  one: `(triangular 800 1000 1500)` shows 1100, `(chance 0.1)` 0.1, so
  `(* (chance 0.1) 3000)` shows the expected cost. They aren't volatile:
  F9 leaves them.
* **`(simulate x runs)`** is a vector of `runs` values of `x` (1000 by
  default, at most 100,000): in each run every uncertain cell `x` depends
  on, however far upstream and in whatever grid, is drawn afresh, the
  cells between are worked out again, then `x` is. `x` is a cell, a
  name, or any expression of them. It's drawn as a histogram, and it
  never spills, even when each run gives a vector (`(simulate [revenue
  cost])` is shown as its first runs: `(map first …)` takes one out).
* **A fixed seed.** The draws come from a seed (1, or `{:seed 7}`), so a
  simulation changes when the model does, not on every edit, and F9
  leaves it. For new draws each F9, `{:seed (RAND)}`. Two simulations
  with the same seed see the same draws in each run, so `(simulate
  revenue)` and `(simulate profit)` line up, run by run. Adding another
  uncertain input doesn't change the draws of the others.
* **RAND** (and `RANDBETWEEN`, `RANDARRAY`, Clojure's `rand`, `rand-int`,
  `rand-nth`, `shuffle`) draw from the seed in a run too: a model already
  built on `RAND` simulates as it is.
* **Reading the runs.** `(stats runs)` is `{:n :mean :sd :min :p5 :p25
  :p50 :p75 :p95 :max}`; `(chance-of pred runs)` the share of runs for
  which `pred` holds; Excel's `AVERAGE`, `MEDIAN`, `STDEV.S`,
  `PERCENTILE.INC` take them as they are. A percentile of whole days
  (finish dates) is interpolated between two days, as in Excel, and a
  date shows its day: for *80% sure by*, round it up,
  `(ROUNDUP (PERCENTILE.INC runs 0.8) 0)`, or it shows the day before.
  `chance-of`'s cell gets the
  format `0%`; a simulation gets its first reference's format, as any
  formula does, so a simulation of dollars labels its chart in dollars.
* **Drawing them.** The default renderer draws a simulation as a
  `histogram`: 24 bins, the outer 5% at each end paler, a dashed line at
  the mean, the least and greatest run written under its ends (in the
  cell's number format, so a simulation of dates shows dates), the
  figures on hover. It fills the cell's height: give the row a height
  for a taller chart. `band` is a percentile band: a line
  from the 5th to the 95th percentile, a box from the 25th to the 75th, a
  tick at the median, a dot at the mean, the 5th and 95th written at the
  ends. Both draw any list or range of numbers.
* Distribution parameters, `simulate`'s runs and `chance-of`'s question
  aren't flagged as constants to name (§26): estimates are typed where
  they're used. The tour's Uncertain inputs section is a small profit
  model; the Project plan example's *Schedule risk* (§46) runs a whole
  plan through its scheduler.
* **A table made by steps.** A simulation can read a column of a table
  a steps formula makes, `(simulate (MAX plan>finish) 300)`: each run
  works the steps out again, with its draws, without drawing the table
  on the sheet, so the steps stay a cell of their own, with their list
  in the Inspector.

## 30. Iterative calculation

A cycle is usually a mistake, and shows as `#CYCLE` (§3). Some are meant:
interest on the average balance (the interest depends on the closing
balance, which includes the interest), a fee that is a share of a total
that includes it, a small equilibrium model where price and demand
depend on each other. Excel works these out with iterative calculation,
repeating the cycle until it settles. A grid can ask RecDep to do the same.

```clojure
;; in a library: the grid's own, an outer grid's, or My library
(def iterate-cycles true)                                  ; Excel's defaults
(def iterate-cycles {:max-iterations 500 :max-change 1e-6})
(def iterate-cycles false)                                 ; not in this grid

opening   1000
rate      0.1
interest  (* rate (/ (+ opening closing) 2))               ; reads closing…
closing   (+ opening interest)                             ; …which reads interest
```

* **Opt-in per grid.** `iterate-cycles` (false in the standard library) is a library setting, like
  `time-limit-ms` and the units settings (§25). It applies to the grid
  whose library defines it and to the grids inside that grid, and the
  nearest definition wins. My library's applies to every sheet. Without
  it a cycle stays a `#CYCLE`.
* **From the cell.** On a `#CYCLE`, the cell menu's *Iterate cycles in
  this grid* and the Inspector's Cycle row write the definition into the
  grid's library (undoable), creating the library if the grid has none.
  The Inspector's *Stop iterating* takes it out again, or writes `(def
  iterate-cycles false)` where the grid would otherwise inherit `true`.
  The definition is only written where it differs from what the grid
  inherits.
* **Every grid the cycle passes through must iterate.** If a cycle runs
  through a grid that doesn't, all its cells are `#CYCLE`, and the
  message ends "not iterated: the grid D1 doesn't iterate cycles". A
  cycle across grids takes the most rounds and the smallest change any
  of them sets.
* **Errors.** A cell of the cycle that fails during the rounds (`#ERR`,
  `#TIME`, an upstream error) keeps its error, and the cells reading it
  are `#UP` naming it. Run again, it would only read the `#UP` it
  caused, so a cycle can't recover from a failure within one recompute.
  The next recompute starts it afresh.
* **Rounds.** Each round works out every cell of the cycle once, in
  address order. A cell reads the values the cells before it left in
  this round, and the rest the values from the round before (Gauss–Seidel,
  as Excel does). It stops after a round in which no cell changed by more
  than `:max-change` (0.001), or after `:max-iterations` rounds (100, at
  most 32767), whichever comes first. Numbers change by their
  difference; any other value counts as changed unless it's equal to the
  one before.
* **Starting values.** The first round starts from the values the cells
  already had, so an edit upstream of a settled cycle usually settles in
  a few rounds. After loading, a rebuild or a library edit, they start
  empty. A cell whose value was an error (the `#CYCLE` from before
  iterating was allowed, Excel's `#NUM!` from a run that diverged) starts
  empty too, so a cycle can recover once the cause is fixed.
* **Not settling isn't an error.** As in Excel, a cycle that hits the cap
  keeps the values its last round left. `A1: (+ A1 1)` counts up by
  `:max-iterations` on every recompute, as Excel's does by its maximum
  iterations. It is flagged the way errors, invalid values, units
  warnings and formula checks are:
  * its cells get a teal corner (bottom left, where units warnings go),
    and a tooltip;
  * the status bar's `⟳ n unsettled` counts them in the whole sheet
    (`⟳ 1 of 2` on one of them); a click goes to the next, Shift-click
    the previous, zooming as needed;
  * the Outline gives each grid its count, `⟳ 2`, a link that steps
    through them the same way;
  * Find in values: an unsettled cell's value reads with `#UNSETTLED`
    after it, so searching `#UNSETTLED` lists them all;
  * the Inspector says how many rounds ran and by how much the last one
    changed.
* **Each cell says how it went**: its value carries `:iterated {:rounds n
  :settled? b :change x :with [addr …]}`, which the Inspector shows as
  "worked out by iterating B2, B3: settled after 6 rounds".
* **One time limit.** All the rounds share one cell's time limit (§8):
  beyond it every cell of the cycle is `#TIME`, "the cycle didn't settle
  within the time limit".
* **Wherever cells are worked out again**, the cycle is iterated again:
  `what-if`, `goal-seek` (§18), `call` (§7) and `simulate` (§29). Each
  probe or run starts from the template's values. A `history` (§28) in a
  cycle gains one sample per recompute, not one per round.
* **Excel import.** A workbook saved with iterative calculation on
  (`calcPr iterate`) gets `iterate-cycles` in the top grid's library,
  with its `iterateCount` and `iterateDelta` when they aren't the
  defaults, so its cycles compute as they did in Excel. They agree
  within the tolerance, not exactly: both stop near the answer, at
  different places. The import's check accepts an iterated cell within
  ten times its `:max-change`.

What stays a `#CYCLE` even when the grid iterates: a cycle through a
grid's value (a formula reading the grid it sits in) or through a
spill's target, and a cycle one of whose results would spill. Only
formulas iterate. A formula that spilled or made a grid before it was
edited into a cycle is a plain formula again: its old spill or grid is
taken down first. For these, the message ends "not iterated: …" with
the reason. Some cycles are found elsewhere, not among the cells, and
keep their own message: a formula reading the grid it makes, a spill
landing on its own inputs, an instance inside its own template.
Iterating doesn't apply to them, and the cell menu doesn't offer it.

## 31. Stepping through a formula

Excel's Evaluate Formula, for Clojure. A formula that gives the wrong
number, `#ERR`, `#TIME` or `#DIV/0!` says only what went wrong,
not where in the formula: its error message, and its inputs in the
Inspector. The Inspector can run it again with the value of every
part recorded (each sub-expression, each `let` binding, each reference)
and draw them as a tree, the part that failed marked.

* Select a formula and click **Step through** on the Inspector's Trace
  row. Each part of the formula is a line, `(* B1 2) → 6`, its parts
  beneath it; a `let` binding is labelled with its name, `avg = (/ sum
  n) → 12.33`; a `for` binding with what it runs over, `x in [1 2 3]`.
  References are links to their cells. Parts that didn't run (the branch
  an `if` or `IF` didn't take) say *not evaluated*.
* **The part that failed** is marked ✗ with the error, and the parts
  around it are opened and shaded, so the way to it shows: `(nth items
  5)` inside a `let` inside a `+`. For `#TIME` it is the part that was
  running when time ran out: `(count (range))`, not `(range)`. The
  summary above the tree says it in one line: *Failed in (apply max …),
  on #DIV/0! from (/ %1 %2)*.
* **Error values** (`#DIV/0!`, `#N/A`, §17) are values, so a formula can
  pass one along without failing: the summary says where it was made
  (*#DIV/0! comes from (/ B2 0)*), following the parts that gave it
  from the outside in. One caught by `IFERROR` on the way isn't it.
* **A function's body** (`#(/ %1 %2)` given to `map`, a `loop`) runs
  many times: its parts sit under the call that runs them and say how
  many times they ran, the last value, and the first error value they
  gave and on which run: `(/ %1 %2) → 3.5 ×3, the last · #DIV/0! on run 2`.
* **An input with an error** (`#UP`) means the formula never ran: the
  reference to that cell is marked; click it, and Trace (which stays open
  from cell to cell) steps through that one.
* **The formula bar** marks the part under the mouse in the tree, and,
  while Trace is open, underlines the part that failed.
* A cell showing `#TIME` isn't run again by just selecting it: Trace
  says it would take up to the time limit and runs it on **Run it**.
* Values are as the formula sees them: numbers in base units (§25),
  distributions at their means (§29). A formula using `NOW`, `RAND` or
  `tick` (§28) gives new values each run, so the steps may differ from
  the cell; Trace says so. So does a cell in a cycle worked out by
  iterating (§30): stepping it is one more round from the values the
  cycle settled on. A bound input (§6) steps through its binding, read
  where it's written.
* **Stepping into a library function**. A call of a function from
  a grid's library, My library or the standard library has **↘ step
  into**: its body runs again on the arguments of that call (the last
  one, if it ran several times: the one that failed, if it did), part by
  part, as the formula did. Above the tree, **↑ back to the formula**
  (or to the function it was stepped into from) goes up a level, and the
  trail says where you are, `in the formula › third-item (the library of
  E2, line 2)`; inside, its own library calls can be stepped into in
  turn. Clojure's and Excel's functions aren't (their
  code isn't RecDep's to show), nor a function passed as a value
  (`(map third-item xs)` is a call of `map`), nor one made with `def` and
  `fn`. The formula bar marks only the cell's own formula.
* **Stepping through a renderer**. **Step through** under the
  Inspector's Renderer draws the cell's value with its renderer again,
  part by part, the tree opening right under that row (**Hide steps**
  closes it; the formula's own Trace and the renderer's are open one at
  a time): `(fn [v ctx] [:b (nth v 5)])` on a number marks
  `(nth v 5)`. A renderer that refers to cells is made from their values
  first (a bad one says so, as on the page); its library calls can be
  stepped into. It's drawn with the `ctx` the cell is (format, unit,
  style, domain, links), without a number format being previewed. A
  named renderer such as `(data-bar 100)` steps into how it's made; the
  drawing function it returns runs outside the run.

## 32. Scenarios

A model gets asked the same question with several sets of inputs: base,
pessimistic, the Q3 plan. A scenario is such a set, named and kept with
the sheet: show one with a click, or compare outputs for every one side
by side. A scenario changes inputs only,
never formulas.

* **The Scenarios tab** is a table: a column per scenario, a row per
  input, every scenario giving every input a value. A model has more
  inputs than scenarios, so the table grows down; when the scenarios
  are too many for the panel it scrolls sideways, the inputs' paths
  staying at the left, and each scenario's name, ✕ and Show sit
  together at the top of its column.
  * *+ Input: the active cell* adds a row: every scenario gives it the
    value it has now. A path typed beside it (and Enter) adds the input
    there, for a parameter, which isn't a cell one selects. *+ Scenario from the values now* adds a column holding
    the inputs' values now, named `scenario 3` (click the name to change
    it). ✕ on an input's row stops setting it, in every scenario; ✕ by
    a scenario's name removes the scenario. An error from any of these
    shows next to the button or field (§5, *Errors where you're
    looking*).
  * A value is written as it would be typed into the cell: `0.05`, `5
    km`, `2026-11-07`, `"north"`; a plain number in a cell with a unit is
    in that unit (§25). A formula is refused, and so is a value the cell
    couldn't hold (`2026-02-30`).
  * Typing values in, as in a sheet: Enter (or ↓) keeps the value and
    moves down the scenario's column to its next input, its text
    selected to type over; Shift-Enter (or ↑) moves up, to the name at
    the top; Tab moves across to the next scenario; Escape puts the
    value back. Enter on a refused value stays in the field, keeping the
    text to fix; leaving it by a click puts the kept value back.
  * **Show** types a scenario's values into its inputs, one undo step.
    It's checked first: a broken or bound input, or a value a validation
    rule (§16) refuses, stops it with nothing written. **● showing**
    marks the scenario whose values the inputs all hold; change one by
    hand and none is.
* **Inputs are paths**, read from the top grid as a formula there would
  read them: `growth`, `plan>growth`, `B30>b>price`, `B30>A3`. A path
  leads to a cell holding a value typed in, or a parameter, in any grid;
  each step is a name or an address. *+ Input* writes names where the
  cells have them, so `plan>growth` rather than `B30>A1`: an address
  works, and moves, but says less in the table and in
  `compare-scenarios`' rows.
  * Why not only named inputs, a grid's interface (§6)? Scenarios
    belong to the top grid and describe the whole model, so setting any
    typed-in value is like typing it, not like using a grid from
    outside. And the reason bindings take names only doesn't apply: a
    binding's key (`:B2`) is a keyword, which nothing rewrites when rows
    are inserted, while a path is a reference, moved like any other.
  * An input bound from outside (§6) is refused, naming its binding:
    writing it would only change its default, out of sight. Set what the
    binding reads instead, which is further out, so also reachable from
    the top grid.
  * Two paths to the same cell (`plan>growth` and `B27>growth`) are
    refused: each cell is set once.
  * Inserting and deleting rows and columns, in any grid along a path,
    moves it; an input whose cell is deleted (or the grid on its way) is
    removed from every scenario. A path that no longer leads to an input
    (a cell whose name was taken away, an emptied cell, or one now
    holding a formula) is marked ⚠ in the table, with why. Renaming the
    cell rewrites the path (§2).
* **`(compare-scenarios x)`** works `x` out for every scenario without
  changing the sheet: a row per scenario, its name, then `x`'s value, or
  `x`'s values when it's a list. `(compare-scenarios [profit npv])`
  spills a table of each scenario's profit and npv. It follows the model
  and the scenarios, like any formula. A broken input makes it an error
  that says which.
* **Files.** Scenarios are saved with the sheet. Save grid leaves them
  behind (they're about the sheet the grid is in), and so does Insert:
  an inserted sheet's scenarios are dropped, and Insert says so.
* The tour's Scenarios section has three, over a small profit model.

## 33. Comments

A comment is a note written on a cell for whoever reads the sheet: where
a number came from, what it leaves out, what's still to check. It
changes nothing the sheet computes. The tour's Comments section has
three.

* **Writing one.** Select the cell and type in the Inspector's
  **Comment** field, just under its name. Enter keeps it (Shift-Enter
  starts a new line), and so does clicking elsewhere; Esc drops the
  change. Right-click → *Add comment* (or *Edit comment*), or Shift-F2
  as in Excel, puts the cursor there. *Remove* beside the field, or
  emptying it, takes the comment away. Each is one undo step.
* **Seeing them.** A cell with a comment has a small gold triangle in
  its top right corner, as in Excel; hover the cell to read it. On a
  cell that also fails its validation, the gold sits inside the larger
  magenta corner. Any cell can have one, an empty cell too, and so can
  a cell holding a grid: a note about the whole grid.
* **Going through them.** As with errors and failing tests, the status
  bar counts the comments in the sheet, `◥ 3 comments`: click it to go
  to the next one after the active cell, Shift-click for the previous
  (it says `◥ 2 of 3 comments` while you're on one, and zooms into the
  grids they're in). The Outline (§11) shows `◥ n` beside each grid with
  comments, counting the grids inside it while it's folded, with the
  same links.
* **Finding them.** Find (§11, Ctrl-F) has **in comments** beside *in
  formulas* and *in values*: it searches every comment (in the viewed
  grid and the grids inside it, if you ask), and Replace and *All*
  rewrite them. A comment replaced by nothing is removed.
* **What happens to them.** A comment goes with its cell: inserting and
  deleting rows and columns, cut and paste, and Paste (Ctrl-V) of a copy
  carry it; Paste values, formulas or formats don't, and a fill (Ctrl-D,
  Ctrl-R) doesn't copy it either: a comment is about its own cell.
  Delete and *Clear contents* keep it; *Clear ▸ Clear all* removes it
  with the cell's contents, formats and name. Comments are saved with the sheet, and with a
  grid saved on its own.

Not supported: replies or threads, who wrote it and when, formatted text, and
reading the comments of an Excel file.

## 34. Data from database files

The **Data** tab reads a SQLite or DuckDB database file, or a CSV file,
shows its tables, shapes a table with steps (filter, clean values,
change a column's type, choose and rename columns, join another table,
sort, group and summarise, remove duplicates, work out a column, or a
query of your own), saves the result as a CSV file or puts the result into the sheet as a grid that
remembers where it came from and can be refreshed. Nothing is installed:
the database is read in the browser, and the file is never changed.
The tour's Data from database files section shows a grid promoted
this way.

* **Opening a file.** Data tab → *Open file…*: a SQLite file
  (`.db`, `.sqlite`, `.sqlite3` …), a DuckDB file (`.duckdb`, `.ddb`), a
  CSV file (`.csv`, `.tsv`, `.txt`) or an Excel workbook (`.xlsx`,
  `.xlsm`), one or several. A workbook's sheets are its tables (*Excel ·
  3 sheets*), each one's cells from its first used row and column: the
  values Excel saved (a formula's result as last calculated), dates as
  `2025-03-12` (with a time when they have one), errors as `#N/A`; row 1
  is the names when it looks like them. A big sheet takes a while
  (100,000 rows × 5: about 6 s). Each file is listed with its tables and views, their
  row counts, and (▸) their columns, with types, keys (⚷) and foreign
  keys (→). A table or view that can't be read (a view over a table since
  dropped) says why; the others still open.
  * In Chrome and Edge the browser remembers the file: after a reload it
    is read again by itself, or, when the browser wants to ask again,
    after one click on *Reconnect*. Elsewhere (Firefox, Safari), pick it
    again after a reload: the tab says so.
  * **A file on the web**: *Open a URL…* and its address
    (`https://…/sales.csv`). It's fetched whole and read as a picked file
    is. It is remembered in this browser by its address, and read again
    after a reload. *Refresh* asks its server whether it changed, and
    fetches it again only when it has. The file's server must let other
    pages read it (CORS). GitHub's raw files do, as does S3 with a CORS
    rule; when it doesn't, the tab says so: download the file and *Open
    file…* instead.
  * A SQLite database another program has open may keep its latest
    changes in a `-wal` file beside it. The tab says *WAL mode, no -wal
    file given*: *Add -wal file…* (under *All files* in the dialog), or
    pick both together, or those changes are missing.
  * *Refresh* reads the file again as it is now; *changed on disk* says
    when it has changed (noticed when you come back to the window).
    *Remove* takes it off the list (the file isn't touched); a grid
    promoted from it still finds it (*Refreshing it*, below).
  * Anything that isn't a SQLite or DuckDB database, or a CSV file, is
    refused.
* **DuckDB files.** The first one takes a moment longer: DuckDB itself
  (36 MB) is loaded then, not before. Then they work as SQLite files do, steps and all, with these
  differences:
  * *DuckDB · 5 tables* says which kind a file is. A table in a schema
    other than `main` is listed as `archive.orders`.
  * The file isn't copied into memory: DuckDB reads the parts it needs.
  * DuckDB keeps its latest changes in `x.duckdb.wal`, beside the file,
    until it checkpoints (or closes the file); its header doesn't say
    whether there are any, so *Add .wal file…* is always offered, and a
    file that shows no tables says to add it.
  * A file written by a DuckDB newer than the one reading it here
    (DuckDB 1.5.4, which reads storage versions 64 to 68) is refused
    before anything loads, saying how to save a copy that can be read:
    in your DuckDB, `ATTACH 'copy.duckdb' AS c (STORAGE_VERSION 'v1.2.0');
    COPY FROM DATABASE yours TO c;`, then open the copy.
  * Values are shown as DuckDB has them: dates as `2025-03-12`, a
    TIMESTAMP WITH TIME ZONE as the UTC time, lists, structs, intervals
    and times as DuckDB writes them as text, booleans as true and false
    (a cell's TRUE or FALSE when promoted), a whole number past 2^53
    exactly, as text, marked *big*.
  * A column has one type, so its details have no values of another kind,
    but for a column of dates kept as text.
  * A SQL step is DuckDB's SQL, and can name the file's other tables
    (`… from prev p join customers c on c.name = p.customer`).
  * A step that runs too long is stopped at DuckDB's first chance after
    30 s: for most queries that's then, but some (every row of one table
    with every row of another) can't stop until much later.
  * Nothing is fetched from the internet: no DuckDB extension is
    installed or loaded, and a query can't read a URL or another file.
* **CSV files** are read by DuckDB (loaded with the first one): a
  million rows in about a second, kept in memory while the page is open.
  A CSV is one table, named after the file (`orders.csv`: *orders*).
  * Every line is a row and every field is kept as written, as text: a
    stray line at the end, a row with fields missing (blank), a value of
    the wrong kind. Nothing is refused or dropped, so the details show
    what's wrong, and filters take it out: *Remove rows where qty is one
    of n/a*, *Remove rows where customer is blank*.
  * Its details judge each value as Insert reads a CSV's fields: whole
    numbers, numbers, dates, text. A column of numbers gets its spread,
    histogram and outliers, and its rogue values (`n/a`, `$12.50`)
    listed; a column of dates, the ones that aren't (`31/02/2025`).
  * Filters and sorts take the values as they read: `qty > 40` compares
    numbers and lets no text through; sorting qty puts 9 before 10 (as
    numbers), and its texts after them. A worked-out column needs the
    text made a number: `try_cast(qty as double) * try_cast(price as
    double)` (`qty * price` is refused: they're text).
  * Promoted, each field becomes a cell as Insert would make it: numbers,
    dates, true/false, text.
  * Row 1 is taken as the column names when it looks like them, and the
    separator (comma, tab, semicolon…) is guessed; under the file's name,
    *Row 1: column names / data* and *Separator* say otherwise when the
    guess is wrong. A file with notes or a title above its data (`Sales
    export, Acme Ltd`): *Skip* that many lines, and the line after them is
    read as the names (unskipped, DuckDB takes the notes for the names and
    the names for a row). A first line with fewer fields than the rows
    (such notes, unskipped) isn't taken for names at all: the columns are
    column0, column1… and the notes are rows, so you see them. The choice is kept with the file in this browser, and
    with a grid promoted from it, so Refresh reads it the same way. A file
    that isn't UTF-8 is read as Latin-1, and the tab says so.
* **Looking through a table.** Click a table's name: its rows show in a
  grid below (or beside, with *⇤ Expand*, which gives the tab the
  window's width). It scrolls through any number of rows: millions are
  read a block at a time as you go, never all at once. Row numbers stay
  at the left; drag a header's edge to widen a column. Numbers are shown
  in full, exactly as the file holds them; a whole number too large for a
  spreadsheet is shown exactly, as text, marked *big*; binary data shows
  its size; very long text is cut, with its length in the tooltip.
* **What the grid shows.** The drop-down at its top left: *Data* (the
  rows alone: no column is analysed), *Details + data* (what's in each
  column, above the rows) or *Correlations* (how the columns of numbers
  move together, instead of the rows). It's remembered in this browser.
* **What's in each column.** *Details + data* adds, under every
  header, what's in the column, a row each so columns line up:
  * a line with values of another kind than most of the column's (⚠),
    outliers (◆), blanks (◌) and the column's kind;
  * Kinds (whole, decimal, date — a text shaped like `2025-03-12` —,
    other text, empty text, null, binary), Blank, Distinct, Range;
  * for numbers: Spread (quartiles and median), Stats (the mean, the
    median and the standard deviation, of every row; for text, the
    shortest, longest and mean length), a Histogram with the mean (a
    solid line) and the median (dashed) on it and the
    outliers counted apart at either end, Outliers (beyond 1.5 × the
    middle half's spread from the quartiles), tinted in the grid too;
  * for dates: a Histogram of a bar a day, a month, a year or a decade,
    whichever its range takes (hover a bar for its count; click it to go
    to its rows);
  * Rogue: the values of another kind, the commonest first (`n/a` in a
    column of numbers, `31/02/2025` among dates);
  * Values: the commonest values, when there are few.

  Click Histogram, Outliers, Rogue or Values on its name to fold it to
  one line, and again to unfold it.

  **Going to them.** The counts and values in the details are links
  (dotted underline): a kind (*decimal 99*), the blanks, *all* the rogue
  values or one of them (*n/a*), the outliers below or above, one of the
  most extreme, one of the commonest values, and the line's ⚠, ◆ and ◌.
  Click one and the grid goes to the first row holding it, the cell
  ringed; a bar under the grid says which it is (*qty · decimal values: 1
  of 99 — row 50*), with *‹ previous* and *next ›* (after the last, the
  first). *Remove these rows* adds a step removing them, every other row
  kept (blanks too); *Keep only these* one keeping only them. What's
  found is what was counted: the step is the same test the count is.
  So removing the rogue values of a column of numbers is *Remove: qty is
  text, date, empty or binary*, a filter by kind (below), and removing
  its outliers *Remove: price is outside 4.975 and 25.015*, Tukey's
  fences at the time, kept as numbers in the step (Refresh doesn't work
  them out again). On a DuckDB or CSV file, filters, sums and sorts read
  a text as a number as Insert does: `inf` and `nan` are texts. A column too narrow for its details
  shows ⋯: click it (or the column's line) and they spill over the
  columns beside it. Counting takes a moment on a big table, and is done
  only while the details are shown; over 200,000 rows the spread and the
  histogram come from a sample (≈), everything else counts every row.
* **How the columns relate.** *Correlations* shows, in place of
  the rows, a matrix: a row per column (its name at the left), under the
  same headers, so a column's menu is still at hand. Each cell is how
  closely its two columns move together, from 1 (green: they rise
  together) through 0 (black: unrelated) to −1 (red: one falls as the
  other rises). It's of the rows after the steps: add a filter and it's
  worked out again for the rows left.
  * Every column with numbers takes part, as Excel's CORREL takes any
    two ranges, whatever kind the details give it: a column of numbers
    and many `n/a`s is a column with missing data, not text. A row's
    blank or other value (`n/a`) leaves that row out of its column's
    pairs only. (In a CSV, numbers are read as Insert reads them.) A
    column whose numbers are half its non-blank values or fewer is marked
    ⚠, on its name, in its cells' tooltips and in the strongest pairs:
    its correlations come from the few rows that have numbers. A column
    without numbers keeps its row and column, blank.
  * A pair is blank, too, when fewer than 3 rows have numbers in both, or
    one of them doesn't vary over those rows (hover it to see which).
  * Under the matrix, the measure: *Linear (Pearson r)*, how well a
    straight line fits (Excel's CORREL), or *Rank (Spearman ρ)*, the
    linear correlation of the values' ranks: 1 whenever one only ever
    rises as the other does, however curved, and an outlier counts for
    no more than any other row.
  * Beside it, the strongest pairs (|r| ≥ 0.3), each a link to its cell:
    ≈ marks a pair saying nearly the same thing (|r| ≥ 0.98: one may be
    worked out from the other, so you may not need both), ~ one whose
    linear and rank correlations differ by 0.2 or more (a curved
    relation, or a few outliers carrying it).
  * Hover a cell for both measures, in words, the rows they're from and
    the p-value: how often chance alone gives a correlation this strong
    between unrelated columns (it matters on a few rows; on thousands
    almost anything is "significant").
  * The line under the matrix says how many columns are numbers and how
    many rows it's from. Over 200,000 rows (or fewer, with very many
    columns of numbers) it's an even sample of about 100,000.
  * It's worked out in the background, once per table and steps: going
    back to *Data* and again to *Correlations* shows it at once.
* **Seeing a pair.** Click a cell of the matrix, or a pair under
  it among the strongest, and the two columns are plotted under the
  matrix (click the cell again, or ×, to close it): a scatter of one
  against the other, with the least-squares line (*ranks* draws their
  ranks instead, what rank correlation measures); for numbers and
  categories, a box plot per category; for two columns of categories, a
  table of how many rows have each pair of values; a column with itself,
  its histogram. Hover a point, box or cell for its values. It's of the
  rows after the steps, an even sample of at most 5,000 of those with
  both values. A column that the measure shown doesn't take (text, under
  *Linear*) says so instead, and that *Any relation* plots it.
* **Any relation.** The third measure, *Any relation (0–1)*,
  asks whether two columns are related at all, whatever the shape, and
  takes text columns too. Each cell is 0 (black: unrelated) to 1 (green),
  with no sign, by the measure the pair's kinds take:
  * two columns of numbers: **distance correlation**, which sees what
    linear and rank miss, a U (a cost against a size, high at both
    ends), a wave, a ring. Where it's clearly above 0 but rank is near
    0, the pair is marked ∿: related, but neither rising nor falling.
    It's quadratic in the rows, so a pair uses an even sample of at most
    1,500 of them (fewer with many columns of numbers); the line under
    the matrix says how many;
  * numbers and categories: the **correlation ratio η**, how much of the
    numbers' spread the categories explain (prices by region);
  * two columns of categories: **Cramér's V**, how far their categories
    go together (region and warehouse).
  * Each is corrected for chance, so unrelated columns read about 0 even
    over a few rows; the plain forms read well above 0 there.
  * A column is **categories** (▤ by its name) when it has no numbers, or
    few (half its values or fewer, its other values taking two or more
    different values: `open`, `closed` and a stray `0`), and its values
    repeat: at most 50 different ones, and at most half as many as its
    values. A column of numbers among `n/a`s stays numbers. Names, ids
    and free text (more different values) stay blank, saying why.
  * It's worked out when you pick it (and remembered, like the others):
    a few seconds on a wide table, in the background.
* **Correlations in the sheet.** The same analysis works on the
  sheet's own ranges, the same way:
  * `(correlations A1:E200)` is the matrix of the range's columns: a row
    of their names, then a row per column headed by its name. It spills;
    give the spilled cells the renderer `correlation-heat` (a range
    format over them) for the Data tab's colours, or give the formula's
    cell `correlation-table` to keep the matrix there as one table. The
    range's first row is its headings when it holds text and no numbers
    (`{:headers false}` or `true` says otherwise); a table of row maps
    works too, `{:columns [:qty :price]}` picking its columns in order.
    Every column with numbers takes part, as for CORREL; one without is
    blank (`{:numbers-only true}` leaves it out); `{:method :spearman}`
    gives the rank measure.
  * `(correlated-pairs A1:E200)` is the pairs as a table, strongest
    first: `a`, `b`, `linear`, `rank`, `rows`, `p` and a `note` (*nearly
    the same*, or *curved, or outliers*, and *few numbers in* a column
    whose numbers are a minority). `{:at-least 0.5}` keeps the
    pairs with |r| ≥ 0.5.
  * `{:method :any}` gives any relation in both: text columns
    take part as categories, and `correlated-pairs` gives `a`, `b`,
    `any`, `measure` (*distance*, *η* or *Cramér's V*), `rank` (beside
    distance correlation), `rows` and a `note` (*neither rising nor
    falling* for a U or a wave).
  * The tour's Data from database files section has one (B7), over
    the promoted grid beside it.
* **Shaping it with steps.** Above the grid, the table and its steps are
  chips: `orders ▸ qty > 40 ▸ Columns: 4 ▸ Sort: price ↓`. Click a chip to
  see the data after that step; ✎ (or a double-click) edits it, × removes
  it. *+ Filter*, *+ Columns* and *+ Step ▾* (the other kinds) add a step
  after the one shown, as does a column's menu: click a column's name in
  the grid (*region ▾*) to sort by it, filter by its values, remove
  duplicates by it, group by it, rename or hide it.
  * **Filter**: *Keep* (or *Remove*) rows where a column passes a test
    (= isn't < ≤ > ≥ between contains, starts with, ends with, is blank,
    isn't blank) against a value, or *is one of* the values you tick in a
    list of the column's values with their counts, the commonest first
    (search it; *tick all shown*; *(blank)* for blank rows; the 500
    commonest are listed, and with more than that the search looks
    through every value: *1 of 1,501 values hold “special”*). Values
    ticked but not listed stay at the top, to untick. A value that reads as a number is compared with
    numbers only (`qty > 40` doesn't let `n/a` through); *isn't* keeps
    blank rows, and *Remove* never removes a row only because it's blank;
    contains, starts with and ends with ignore case, take `%` and `_`
    literally, and read a number as it's written (*qty contains 2* keeps
    12 and 25). *is outside* two numbers: numbers below the one or above
    the other (texts and blanks don't pass, so *Remove* keeps them). *is of
    kind*: tick kinds, as the details count them (whole, decimal, date,
    text, empty, null, binary): *Remove rows where qty is of kind decimal*
    takes out the decimals of a mixed column, *Keep … null* shows the rows
    without a qty. *+ condition* adds another: rows are kept (or removed)
    when *all* of them hold, or *any* (*region = North or qty > 40*). A
    condition on *(any column)* or *(every column)* tests them all:
    *Remove rows where (any column) is blank* keeps only complete rows.
  * **Clean values** (a column's menu, *Clean its values…*): trim spaces
    at either end, UPPER or lower case (text only: numbers are left
    alone), replace a whole value (`n/a` by nothing: a blank; ` north ` by
    `North`), replace text within values (every `-`), or fill blanks.
    Before you apply it, it says how many values change, with the
    commonest before and after (*974 values change: “ north ” → “north”
    × 974*). *Fill down* (also *Fill it down* in a column's menu): a
    blank takes the value above it, for a report that names a group on
    its first row only. *Fill blanks with the mean* (or *median*): of the
    column's numbers (text among them left out). *Collapse spaces*: trim,
    and each run of spaces, tabs or line breaks between words one space.
    *Title Case*: each word capitalised (`ANN SMITH` → `Ann Smith`).
    *Replace by a pattern*: every match of a regular expression (`[0-9]+`
    digits, `\s+` spaces, `(…)` a group) replaced, `$1` putting a group
    back (`(\w+) (\w+)` with `$2, $1`: `Ann Smith` → `Smith, Ann`); with
    nothing, matches are removed. *Extract by a pattern*: the first match
    (or the first group's): `Qty: 12 units` with `[0-9]+` gives `12`; no
    match, blank. *Start from* offers a number, an email address, a date,
    the text in brackets. *ignore case* for either.
  * **Change type** (*Change its type…*): make a column numbers, whole
    numbers, dates (read in a layout you pick: `dd/mm/yyyy`, `mm/dd/yyyy`,
    `yyyy-mm-dd`…) or text, as Insert reads a field (`$12.50` and `1,234`
    aren't numbers; `31/02/2025` isn't a date, nor is `12/03/25`: a year
    has four digits, so none is guessed). Before you apply it, it
    counts how many convert, how many are blank (they stay blank) and how
    many don't, with the commonest of those (*1,994 don't (“n/a” ×
    1,994)*); those become blank, or, if you choose, their rows go.
    Promoted, a column made dates is dates, and one made text stays text
    (even in a CSV).
    *read as written* reads numbers as people write them: a currency sign
    or code (`$1,234.50`, `1,500 USD`), thousands in threes, `12%` (0.12),
    brackets for a negative (`(12.00)`); *as written, decimal comma* for
    `1.234,50 €`. Thousands must be in threes, so `12,5` with a decimal
    point is counted as not converting, not read as 125. Dates can also
    be read with a time (`2025-03-12 14:30`, `12/03/2025 14:30`: a date and
    time), with month names (`12 Mar 2025`, `March 12, 2025`), or as
    Excel's day numbers (`45728`).
  * **Split a column** (*Split it…*): cut a column's text at a separator
    into the columns you name (`Smith, John` at `,` → *last*, *first*);
    the last takes the rest (`Lee, Kim, Jr` → *Lee* | *Kim, Jr*). Parts are
    trimmed; one missing or empty is blank. The original goes, unless
    kept.
  * **Merge columns**: join the columns ticked, in the order ticked, with
    a separator between, into a new column where the first of them was
    (blanks skipped: *Brown*, not *Brown ␣*). The originals go, unless
    kept.
  * **Append rows**: another table's rows after these: the same file's
    (a SQLite file), or any open DuckDB, CSV, Excel or grid source's (a
    CSV file or a DuckDB file): monthly files of one layout, or rows
    exported, fixed and opened again. Columns are matched *by name* (its
    other columns are added after these, blank for these rows; the editor
    says which) or *by position*. These rows come first, then its, in
    order. Refresh reads the appended file again too. To append a CSV to
    a SQLite table, promote the table, then *Open in Data* on its grid,
    and append to that.
  * **Columns**: tick the columns to keep (*all*, *none*), put them in
    order (drag a column by its grip ⠿ or its name to where it goes, in
    one move; or ↑ ↓) and give any a new name; or type the names in a line,
    *Names, in order*: `id, customer, qty` (commas, or tabs: a row pasted
    from a sheet), for a file without them (column0, column1…). A
    column's menu, *Rename it…*, opens it with that column's new name to
    type.
  * **Sort**: by one column or more (*+ then by*), smallest or largest
    first.
  * **Group and summarise**: tick the columns to group by, and work out
    for each group a count of rows, a count of values (not blank),
    distinct values, the sum, average, smallest or largest of a
    column, or a **list** of its values; the four between count numbers
    only, so `n/a` among them is left out rather than read as 0. A list
    is each group's different values, blanks left out, in the order they
    come: grouping a timesheet by task, a list of Resource is who logged
    time to each task (`["Ana" "Ben"]`). Run in a database (§34), the
    values come joined by commas instead (`Ana,Ben`). Each gets a column (*count*, *average of
    price*), or a name you give it. Nothing ticked to group by: one row
    for the whole table.
  * **Remove duplicates**: keep the first row of each, as they come (so
    after a sort, the first in that order), by the columns ticked, or of
    each whole row when none is. Or *show only the duplicates*: every
    copy of each repeated one, the copies together (also *Show its
    duplicates* in a column's menu), to look at them before removing them.
    *ignore case and spaces* treats ` ann  smith` and `Ann Smith` as the
    same. Near matches (`Jon` and `John`) aren't found.
  * **Pivot**: a column's values become columns (*month*: `jan`, `feb`,
    `mar`), each filled from another column (*amount*), the rows grouped
    by the other columns. Fill with its value (the first, when a group has
    several), a count of its values, the sum, average, smallest or largest.
    Which values make columns: *every value the data has when the steps
    run* (the first choice), so the steps work on whatever data comes,
    next month's file making next month's columns, in the order each
    value first comes; or *only the values ticked*, for a one-off look:
    the values there are now are listed and ticked, untick any you don't
    want, and one that turns up later makes no column until it's ticked.
    (*Pivot: its values as columns…* in a column's menu.)
  * **Unpivot**: the other way: tick columns (*jan*, *feb*, *mar*), and
    each row becomes one per column, its name in one new column (*month*)
    and its value in another (*amount*), the other columns repeated.
    Blank values' rows are left out unless kept.
  * **Remove empty rows** (in *+ Step ▾*): rows blank in every column go
    (a filter: every column is blank, removed).
  * **Join**: another table of the same file (or, for a DuckDB, CSV,
    Excel or grid source, of any other one open: *customers.csv* to look
    names up from, a grid as a lookup table), its rows matched to these
    where a column here equals one there (more pairs: *+ and another
    pair*). It suggests the pair from a declared foreign key, or a
    column named after the table (`customer_id` for *customers*' `id`).
    *Keeping every row here* gives blanks where none matches; *only the
    rows that match* drops the rest. Tick the columns to bring in (one
    whose name is here already comes in as `customers.name`). Before you
    apply it, it says how the rows match: *4,982 of 5,000 rows match*;
    *some match several rows, so rows repeat: 5,000 become 26,569*; or
    that none match, and to try *compare as text* when one column holds
    numbers and the other the same numbers as text (`10` and `"10"`).
    The rows keep their order; Refresh reads the other file again. Anything
    more (a join on a condition) can be a SQL step: `select … from prev p join customers c
    on c.name = p.customer`.
  * **Skip or keep rows**: skip the first *n* rows, then keep the next
    *m* (blank: all the rest). Skip 2, keep 100: rows 3 to 102.
  * **⇄ Rows left out** (in *+ Step ▾*): the rows the steps before
    it leave out, instead of those they keep: every filter, skip, removed
    duplicate, value that didn't convert (when its rows are dropped) and
    row a matching-only join drops, with the same columns and the same
    cleaning. So after removing the rogue values and the outliers, it
    shows just those rows: export them (*Export CSV…*) to fix them
    elsewhere, or promote them. Remove the step to go back. Steps after it
    work on those rows; a second ⇄ inverts everything before it, so two in
    a row give back the rows kept. Not
    after *Group and summarise* or a SQL step, whose rows aren't the
    table's (the step says so).
  * **Column =**: a new column worked out from the others, in SQL:
    `qty * price`, `round(price * 1.2, 2)`, `upper(region)`.
  * **SQL**: any query over `prev`, the table as the step before leaves
    it (`select region, count(*) n from prev group by region`).
    Ctrl-Enter applies. Only a query runs: nothing can be written.

  A step that fails says which step and why, with a caret under the
  place in its SQL or its expression; the grid keeps showing the last
  steps that worked. A step that runs for more than 30 s is stopped.
  *▸ SQL* shows the query the steps make, with its values, and copies it.
  *Edit all as SQL* there turns the steps up to the one shown into a
  single SQL step, its values written in, to go on with by hand; it asks
  first. A table's steps last while the page is open.
  **↶ ↷** undo and redo changes to the steps: one added, edited or
  removed, *Edit all as SQL*, saved steps applied (Ctrl-Z and Ctrl-Y too,
  while you're working in the table; elsewhere they're the sheet's).
  ↶ ↷ appear once there's a change to undo. *+ Step ▾ › Save these
  steps…* keeps the steps up to the one shown by name, in this browser
  and the library file (not in the workbook file), and *Apply “name”* in
  the same menu adds them to any table later: next month's file, named
  otherwise, gets this month's cleaning in one click (after the step
  shown; ↶ takes them off). The menu lists the five used (or saved)
  last, × beside one forgets it, and **Saved steps…** below them opens
  all of them:
  * **a search** through their names, the files they came from and
    their steps (*unpivot jan*: every word, any case);
  * **how each fits the table shown**: *✓ fits this table*, or *✗
    missing: name, score*, the columns its steps read that this table
    hasn't (those steps would fail). A column a step makes (a split's, a
    pivot's, a rename) isn't counted; a SQL step's text and a worked-out
    column's expression aren't read, and it says so. Ordered by fit
    (those that fit first), by last used, or by name;
  * ▸ a name shows its steps; **Apply**, **Rename** (Enter; not onto a
    name taken) and **Forget**, and ticked ones forgotten together.
    Esc or Close closes it.

  They go in the library file (*Library ▾*, §4) with My library and the
  shelf: saved with them, and opened merged by name (the same steps
  are the same, whenever saved) or replacing, so they move to another
  browser or to a colleague in one file. Steps can hold some of your
  data (a filter's ticked values, a pivot's columns, file and table
  names), so look before passing it on; it's why a template on the shelf
  leaves a grid's steps behind. *Replace* with a library file that
  has no saved steps drops yours: *Merge* keeps them.
* **Saving it as a CSV file.** *Export CSV…* (beside *Promote*) saves the
  data as shown, every row (not only those on screen), column names
  first: the save dialog in Chrome and Edge, else a file name and a
  download. Numbers are written in full, dates as `2025-03-12`, blanks as
  empty fields; a field with a comma, a quote or a line break is quoted.
  The file starts with a byte-order mark, so Excel reads its accents
  right (RecDep's Insert and the Data tab skip it). A sheet's grid can be
  saved the same way: *Export CSV…* on the grid bar saves the grid shown,
  the values its cells show (not their formulas), numbers in full
  (1234.5, not $1,234.50), dates as they show, without the empty rows and
  columns at its end.
* **Promoting it into the sheet.** Select a cell in the sheet, then
  *Promote ▸ B2*: the data as shown (after the step shown) becomes a grid
  in that cell, row 1 its headers (*Row 1 is headers*, so a formula reads
  a column by name: `orders>qty`), the values typed in: numbers as
  numbers, text as text, and in a column declared DATE, dates. It's one
  undo step, and it asks first when the cell holds something, or the
  result is more than 200,000 cells. It's an ordinary grid from then on:
  edit it, sort, chart it, refer to it.
* **Refreshing it.** A promoted grid keeps where its data came from: the
  file, the table and the steps. Select its cell, or any cell inside it:
  the Inspector's **Data** row says where it came from (*orders ·
  sales.db · 2 steps · 4 rows, promoted 13 Nov, 19:14*; hover it for the
  steps). Inside it, the grid bar says so too (*Data: orders · sales.db
  ⟳ Refresh · Edit in Data*). *Refresh* runs
  the steps again on the file as it is now and replaces the data, as one
  undo step. Whatever you added beside the data stays (a column of
  formulas, say), and so do the grid's formats and widths. Before
  replacing anything it asks when it would disturb something: values you
  typed over the data, cells beside it that won't line up with its rows
  once their number changes, or a column that's gone (formulas reading it
  by header would show #REF).
  * Its file needn't be open in the Data tab: Refresh opens it (asking
    for permission, from the same click, if the browser wants to), and
    so does a file removed from the Data tab, since its handle is kept
    (it's found by its name: the message says which file it read).
    A web page never learns a file's path, so what it keeps is the
    browser's handle to it (Chrome, Edge), or, elsewhere, the file
    picked this visit.
  * If this browser hasn't got it (another computer, or Firefox after a
    reload), Refresh opens the open dialog for it.
* **Changing its steps.** *Edit in Data* opens the table with the grid's
  steps in the Data tab; change them, then *Update B2* writes the data
  and the steps back to the grid. *Update B2* is offered for every grid
  promoted from the table shown (up to three named), so steps changed
  in the Data tab can go back to a grid without *Edit in Data*; it
  replaces that grid's steps with these (undoable). And the other way:
  when the Data tab has other steps for the grid's table than the grid's
  own (you reordered its columns there, say), the grid bar and the Data
  row offer *Use the Data tab's steps* beside *Refresh* (hover it for
  them). *Refresh* itself always keeps the grid's own steps.
* **A grid as a source.** *Open in Data* on the grid bar puts the grid
  shown into the Data tab as a source (*grid B2*), its table (*B2*) shown:
  the values its cells show, row 1 its column names when the grid has
  headers, dates as `2025-03-12`. Filter it, clean it, find its rogue
  values, export it or promote the result into another grid, as with a
  file. *Refresh* (in the Data tab, or of a grid promoted from it) reads
  the grid's values again, as they are then; a grid promoted from it
  finds it again by its address after a reload. So data can go round:
  cleaned in the Data tab, worked on in the sheet, back to the Data tab.
  The result can't go back over the grid it's read from (or a cell of
  it, or a grid holding it, or a grid a step appends or joins): writing
  there would change what's read, and every read after, with no steps
  too, would be the result. *Promote ▸* and *Update* are greyed out
  there and say so; select a cell outside it. (The whole sheet as a
  source is the exception: its results go beside its data.)
* **Promote live.** For a grid of the sheet, *Promote live ▸ E2* beside
  *Promote* writes the steps into the cell as a formula instead of the
  values they give now: `(steps orders [{:op :window :fn :running-sum
  :column "qty"}])`, drawn as a table. It follows the grid as it
  changes, with no Refresh, and its steps are chips in the Inspector to
  change there (§43). Steps only the Data tab has, a SQL step, *Rows
  left out*, or a join or append of another of its tables, can't go
  live: you're told which, and Promote gives their values. The grid
  needs headers (*Use row 1 as headers*), since the Data tab names a
  headerless grid's columns column1, column2 ….
* **What goes with the grid.** Its steps are saved with the sheet, with
  the grid saved on its own (Save grid), and with copies. A template put
  on the shelf leaves them behind (they can hold some of your data).

Not supported: joining a SQLite file's table with another file's (DuckDB, CSV,
Excel and grid sources join one another; a SQL step can't), Parquet files and
DuckDB 2.0's own file format, a query's values taken from cells, a Stop
button for a long query, and SQLite files over about 1 GB.

## 35. Templates as definitions

A template is a grid you use as a function: `(call colstats {:data A3:A7}
:A1)`, an instance of it (§21), a what-if. Templates are treated as
functions are: kept somewhere, used by name from anywhere below, carried
along when a grid is saved, and kept once where they're the same.

* **Where a template lives.** A grid held in a grid's margin (⚙
  Parameters, the strip shows it as `▦ template`) is a template of that
  grid. Every formula in that grid, and in every grid inside it, uses it
  by its name, bare: `(instance colstats …)` two grids down means the
  `colstats` in `$`'s margin. The nearest one of a name wins, as a
  grid's library definition wins over its surroundings'. Only templates
  are seen from below: a value (a parameter holding a number) stays its
  grid's own, so what a grid reads from outside stays explicit, through
  its inputs (§6). A function's name wins: in a grid whose library (or
  one around it) defines `colstats`, `colstats` is the function.
* **The shelf is the outermost place.** A formula naming a template on
  the shelf (`colstats`, or `loan-calc` for *Loan calc*) works in any
  sheet: the sheet takes a copy of it, in its top grid's margin, marked
  *follows the shelf's*. Put a changed version on the shelf, and every
  open sheet following it moves on to it (another tab's too). A name
  that's a function's (`count`) is never taken for a shelf template.
* **Borrowing doesn't break a grid.** A grid using a template from
  around it isn't reaching out: Self-contained (§6) doesn't flag it,
  Save grid doesn't warn of it, and Make input isn't offered. The grid
  cell's Inspector lists it under *Borrows*, with where it comes from
  and its tag (`colstats · 3f9a` — the same content, the same tag). A
  template that uses another (`(call helper …)` inside it) is the same
  as another version only when what it uses is too: the shelf's version
  carries `helper`, so changing `helper` makes `colstats` differ from
  the shelf's, and putting it on the shelf again makes it the same. What
  a borrowed template itself reaches out to (`..>A1` inside it) is in
  Save grid's warning, as the copy in the file would read other cells.
* **Save grid carries what the grid borrows**, a copy of each template
  in the file's top grid's margin, so the file works on its own.
  **Insert** leaves a carried copy out where its name means the same
  template already (around the cell inserted into, or on the shelf), so
  the grid uses that one and follows its changes; one that differs
  stays, as the version the grid was made with, and the message says
  which were left out. Opening a sheet whose copies differ from your
  shelf's says so; they stay as they are.
* **Pin or follow.** Following the latest version and keeping the one a
  grid was built with are both right, at different times:
  * *Keep a copy here* (Inspector, *Borrows*) copies a borrowed template
    into the grid's margin: the grid keeps that version, whatever
    happens around it. *Use $'s* (beside a grid's own copy) removes it,
    and the grid follows the one around again. Each own copy says
    whether it's *the same as* or *differs from* the one around.
  * In `$`, a copy of the shelf's says *follows the shelf's*; *Keep this
    version* stops it following (the sheet's own from then on), and
    *Use the shelf's* makes a differing one the shelf's again.
  * *Move up to …* (beside a grid's own copy) appears when several grids
    inside one grid each keep the same copy: one copy goes in that grid's
    margin instead, and they all use it. It isn't offered when a formula
    in there would then mean a different version than it does now.
* **Renaming a template on the shelf** keeps the sheets following it:
  their copies are linked to it under its new name (open ones at once,
  others when they're opened), and their formulas keep the name they use.
* **Put on the shelf** says *on the shelf already, the same* when it is,
  rather than asking to replace it.
* **×** on a `▦ template` in the strip removes it, its grid too, and
  says what its uses mean now (the one of its name around it, the
  shelf's, or nothing: `#NAME`). One the sheet would take straight back
  from the shelf (the shelf has it, and cells use the name that nothing
  else would mean) stays, and the message names those cells: remove
  those uses first, or take it off the shelf.
* Renaming a template (its name field) renames the references to it in
  the grids inside, as for any named cell.

### Starting one, publishing it, and what an edit reaches

* **Start in the margin: *+ template*** (⚙ Parameters, beside the field
  for adding one) asks for a name and opens an empty grid, a template of
  the grid in view. Develop it there; instances you place show its
  results on the sheet meanwhile, and their ⧉ link brings you back.
* **Or make a grid one later: *Make template…*** (right-click a grid, or
  a cell in it). The grid moves into the margin of the grid it's in,
  under a name asked for (`colstats` for `colstats-dev`), and an instance
  of it takes its place: the cell keeps its name and size, and its
  bindings become the instance's inputs, so what you see, and what
  `(call colstats-dev …)` reads, stay the same. Formulas reading into its
  cells (`colstats-dev>A2`) can't read into an instance: they're listed
  first, as are references out of it by a move from its cell, to go on
  or not.
* **Publishing moves it.** *Put on the shelf* (the menu, or the button
  beside a template in the strip and *Borrows*) on a template: in `$`,
  it becomes the sheet's copy, following the shelf; in a grid inside, it
  leaves that grid, and its name there means the sheet's copy — one
  copy, not two (one undo step, with the shelf's change it makes here).
  If you were in the template, the view moves to the one its name now
  means. It stays where a different version around would take its
  place, or where a grid's binding gives it (a template passed in, §6),
  and the message says so. Develop the sheet's copy from then
  on, and publish again when ready; *Keep a copy here* gives a grid its
  own version on purpose.
* **A cell named like a template hides it.** In a grid, its own cell
  names come first: a grid (or any cell) named `cal` in B30 is what
  `cal` means for B30's formulas, not the template `cal`. It's flagged
  by the formula checks (*this cell's name, cal, hides the template cal
  of $*), and the template's bar lists it (*not by B30>A23, a cell of
  B30 named cal*, a link to it). Rename it, or remove it, for B30's
  formulas to mean the template. *Place as a copy* names its copy
  `cal-copy` for this reason, and placing an instance where a cell
  already has the template's name is refused, saying which cell.
* **Placing where a different version is in scope** asks which: OK makes
  that one the shelf's version, Cancel uses it as it is. No
  `colstats-2` is made.
* **The template bar.** While a template is shown, the grid bar says
  what editing it reaches: *Template colstats in $ · used by 5 · not by
  B9 (its own, differs, 2 uses)*, with *Use this one there* to drop that
  grid's own, and its standing against the shelf's (*Put on the shelf*,
  *Use the shelf's*). A template hiding another of its name is also a
  formula check (§26), so it's counted in the status bar's ◤; *Keep this
  one* beside it says it's on purpose. *Make template…* from inside the
  grid, and *Use $'s* or *Move up* on the template you're in, leave you
  in the template the name now means.

**Limits.** Only templates in a margin are seen from below, and only by
a bare name (`$>colstats` is a path, which reaches out as any path
does). A copy changed in a sheet isn't linked to its template's new name
after a rename on the shelf. *Move up* goes one grid at a time.

## 36. Sheet tabs

A new "sheet" in RecDep is a grid in a cell of `$`. The strip under the grid is Excel's sheet tabs for
these grids.

* **Home and Advanced.** At the strip's left, the house is *Home*: it
  goes to `$`, the top grid, or, on a sheet with a facade (§54) and
  Advanced off, to the facade, the sheet as its users see it. It goes
  to that grid itself, out of any grid inside it, and is lit while
  you're there. Beside it, on a sheet with a facade, is *Advanced*
  (stacked layers), lit while it's on: it says which sheets have tabs
  (§54). Home is never a tab of its own: `$` has none, nor has the
  facade with Advanced off.
* **The tabs.** Every grid in a cell of `$` that's drawn
  as a link (the ▦ button, `grid-link`, which a grid has unless it's
  given another renderer), in reading order. Each is labelled by its
  name, or by its address when it has none. A grid drawn in place
  (`(grid-view)`, like the tour's dashboard in C3) is part of the `$`
  page and has no tab; give it `grid-link` and it gets one. **A grid
  whose name starts with `_`** (`_loading`) has no tab either: it's
  private (§6), the data behind a sheet that presents it, reached
  through that sheet, the Outline, the name box or a link. Nothing
  about the tabs is saved: they are the grids of `$`. So an Excel
  workbook's sheets (§17) are tabs as soon as it's opened, and a grid
  put in `$` (the Inspector's *Make grid*) becomes one.
* **Which tab is lit:** the one holding the grid in view, however deep
  you've zoomed into it (the breadcrumb says how deep). Home is lit
  while you're in `$`, or in a grid drawn in place there.
* **Switching.** Click a tab to go back to where you last were in it
  (the grid in view and the selected cells), or to its grid the first
  time. Clicking the lit tab goes out to its own grid.
  Ctrl-Alt-← / → go to the previous / next tab, Home counted first,
  wrapping round (not
  while typing in the formula bar or a box). Alt-PgUp / Alt-PgDn do
  too, on keyboards that send Alt with PageUp (some laptops, where
  PageUp is a Fn key, don't). Excel's Ctrl-PgUp / PgDn work where the
  browser passes them on; Chrome and Edge keep them for their own
  tabs. ▾ lists
  every tab, Home's grid first, and ◂ ▸ scroll the strip. Switches are jumps: Back returns.
* **+ makes a sheet:** a 30×10 grid (a New sheet's size) named Sheet1,
  Sheet2 … (the first name that's free). It goes in the first empty cell
  going down the column of the last tab, or of column A when there's
  none, and `$` grows a row when that column is full. It's shown, and
  the status bar says where it went. One undo step.
* **Renaming:** double-click a tab (or right-click → *Rename*), type,
  then Enter; Esc keeps the old name. This renames the cell, as the
  Inspector does: formulas that used the old name are rewritten (§6),
  and a name that can't be used is refused, saying why.
* **Right-click a tab** (or Home, for `$`'s *New sheet*) for *Rename*, *Select its cell in $* (to give it
  a renderer, move it or comment it), *New sheet* and *Delete sheet…*.
  Deleting removes the grid with everything in it, and its name. You're
  asked first. Formulas reaching into it show `#NAME` when they use its
  name (`Sheet1>A1`) and `#REF` when they use its address (`B30>A1`).
  Undo brings it back.

**Limits.** Only the grids of `$` have tabs. Tabs can't be dragged to
reorder them: their order is their cells' order, so move a cell (cut and
paste) to move its tab. Where you were in each tab is forgotten on
reload.

## 37. The REPL: writing a library's functions

The Library tab's editor applies a
definition, but it can't try one on the sheet's values. The **REPL tab**
is a scratchpad for that: write a function, call it on cells, change it,
and Publish it into a library when it works.

* **What it sees.** The REPL sees what formulas in the active cell's
  grid see: that grid's library and those of the grids around it, My
  library and the standard library. Which library a definition goes
  into is chosen when you publish it. A library's private definitions
  (`defn-`) aren't seen.
* **Evaluating.** Type a form and press Enter: `(SUM A1:A3)`,
  `(map tax B2:B9)`. It's evaluated as a formula in the active cell
  would be: references, names and `..>` paths read the sheet from
  there, and `call` and `here` work. Not everything a cell's formula
  does: a reference to a cell in error stops it, even inside `IFERROR`
  (it names the cell); `simulate`, `compare-scenarios`, `unit`, `tick`
  and `history` work only in a cell. Enter goes on to a new
  line while a bracket is open; Ctrl/Cmd-Enter evaluates anyway;
  Shift-Enter is a new line. ↑ and ↓ (on the first and last line) bring
  back earlier inputs; click one in the history to put it back in the
  input. Esc goes back to the grid, keeping what's typed.
* **What the formula bar has, the input has:** coloured references
  with their cells outlined, coloured brackets with the pair at the
  cursor boxed, completion (with the scratch definitions), hover, F4,
  and clicking or dragging over cells to write their reference after
  `(`, a space or a comma.
* **Results.** The value is printed (long collections cut short); what
  `println` prints is shown above it in italics. A renderer's output
  (hiccup, `[:b "bold"]`, `((data-bar 25) 7 {})`) is drawn too, as a
  cell would draw it. An error says what went wrong, and a cell in
  error that the form reads is named.
* **Definitions stay in the scratchpad.** `(defn …)`, `def`,
  `defmacro` and the rest evaluated here are listed under *Scratch
  definitions*, not put in any library, with their `;;` comment lines.
  Each is marked against what formulas in the active cell's grid see
  under its name: *new* (nothing defines it), *changed · B5* (B5's
  library, or `$`'s, My library's or the standard library's, has
  another), *published · B5* (it has this one) or *doesn't load* (a
  name it uses has gone; hover for why). *Edit* puts one back in the
  input, *Discard* drops it. Defining a name the library or the standard
  library has (a fixed `tax`, your own `SUM`) shadows it in the REPL
  only: formulas keep using the library's until you publish, and a
  published one kept in the scratchpad goes on shadowing it there, so
  a later change in the Library tab shows it as *changed* until you
  discard it. The
  library's own functions keep calling the library's definitions, as
  everywhere (§14); the scratch definitions call each other's.
* **Writing into cells.** A form typed here is given `ctx`, as a
  button's handler is, so `set-cell!`, `update-cell!` and `set-self!`
  work, relative to the active cell: `(set-cell! ctx 'u0 (SUM A1:A3))`
  pins the total into the active cell as a value (it doesn't follow
  A1:A3 afterwards), `(set-cell! ctx "B5" (tax 100))` writes into B5,
  `(update-cell! ctx 'd1 inc)` adds one to the cell below. The cell to
  write is quoted (`'A20`) or a string (`"A20"`), never bare: a bare
  `A20` is A20's value, as everywhere in a formula, so
  `(set-cell! ctx A20 v)` hands set-cell! what A20 holds. The writes
  are made once the input has run, all of them as one step that Undo
  takes back, and the transcript says where they went. So the form
  itself reads the values from before them. Nothing is written if the
  input stops at an error or a write is refused (a range, a bound input,
  a value its validation refuses), and it says why. A definition can't
  use `ctx` without being given it: `(defn put [ctx v] (set-cell! ctx
  'd1 v))`, then `(put ctx 99)`.
* **A method can't be added to a library's multimethod** from the
  REPL (it would change the sheet's): evaluate the `defmulti` there too,
  then its methods.
* **A definition can't read cells**, as a library's can't:
  `(defn f [x] (* x A1))` is refused, saying so. Give it the value as an
  argument, `(f A1)`.
* **Publish…** beside a definition (or *Publish all…*) shows what would
  be written and where. *Publish to* picks the library: the active
  cell's grid's (made if it has none), one of a grid around it, or My
  library. It starts on the library the definition came from, if it
  replaces one, else the nearest library around the active cell, else
  the top grid's. Below it, the definition and every scratch definition
  it uses ("used by net"), each new or replacing the library's. Publish writes
  them in one step, which Undo takes back: a replaced definition stays
  where it was in the library, new ones go in before the first
  definition that uses them, else at the end, and each goes in after
  what it uses, since a library loads top to bottom. It's refused, with
  why, while a definition doesn't load, while one uses a name the
  library can't see (a grid's function, publishing to My library),
  and while the Library tab's *Edit all* has changes to that library
  not applied. Published definitions stay in the scratchpad, marked
  *published*, to go on working on; *Discard published* clears them.
* **From the Library tab.** *REPL* beside a definition opens it in the
  REPL's input: change it, try it, and Publish puts it back, offering
  its library first. With *Show code* on, *Adapt in the REPL…* opens a copy of a
  definition from another library (the standard library's, an outer
  grid's) to make your own version.
* **Kept in the browser.** The scratch definitions, the history (as
  printed) and the input are kept in this browser, like My library, not
  in sheets: they're there after a reload and with any sheet. *Clear
  history* empties the history; the definitions stay.

## 38. Grids that grow to fit their results

A grid in a cell grows to fit what spills in it. A result that runs off
its bottom or right-hand edge adds the rows or columns it needs, and
they go again when the result needs less. So a grid can take any grid
as its input and lay out a result as big as that input: a data-analysis
template doesn't have to be as wide as the widest data it will see.

```clojure
;; in A3 of a template whose parameter is data: one mean per column
(BYCOL data AVERAGE)
(instance means {:data B2})     ; elsewhere: given a 4-column grid, 4 means
```

* **Its own size.** The size a grid was given is its *own size*, the
  least it ever is: it can't shrink itself away. It's what is saved and
  copied (a copy, Save grid, a template from the shelf), what the
  Inspector's boxes and the format bar's rows and columns show and Resize
  sets, and what `down` and `across` fill to (so a fill can't feed its
  own growth). The Inspector says *grown to 3 × 7 by its results*
  beside it; the breadcrumb shows the size now.
* **The grown part holds results only.** Typing a formula or a value
  there is refused, saying why: it goes when the result shrinks.
  Formatting it is fine. Make the grid that big (Resize) to write there;
  *Convert spill to values* over a result that grew it does that for
  you, and so does placing anything from the Shelf tab there. Rows and columns are inserted and deleted within its own size.
* **Where to put working cells.** A grid grows at its bottom and right
  only, so keep working cells above or to the left of a result that
  grows; one in its way is `#SPILL`. Hide them with a display
  area that runs to the edge (§6): select the result's row up to the
  right-hand edge, right-click → *Show only this*.
* **Instances and calls** (§7, §21) grow for themselves: an instance
  given wider data is drawn as wide as its result, while the template
  stays as it is; `(call means {:data B2} :D3)` reads a grown cell.
* **References.** Formulas that read the grid's shape (its own cell's
  value, ranges to its edge, cells past its own size) follow it as it
  grows and shrinks; a cell that goes is `#REF` until it comes back.
* **What doesn't grow.** The top grid (the sheet, sized by hand: it
  grows when you type or paste past its edge, §12, but not to fit a
  result), and a grid made by `as-grid` (§22), already as big as its
  result. A result
  whose size follows the grid's own, such as one as wide as a range to
  the grid's edge, would grow without end: it's `#SPILL`, saying so,
  until its formula is edited.

**Limits.** Formatting in the grown part stays when the grid shrinks,
out of sight, and shows again when it grows. Excel's `BYCOL` with
`AVERAGE` over a column of text is a `#DIV/0!` for the whole result in
RecDep, not for that column only.

## 39. The solver

Goal seek (§18) changes one input to hit one target. Excel's Solver
changes several, keeping to constraints, to make an output the largest,
the smallest or a number. Here it's a formula, `optimize`, over a
template grid as `call` and `goal-seek` are, and *Solver…* on the cell
menu for the sheet itself.

```clojure
;; Plan holds a product mix: chairs, tables -> hours, wood, profit (named)
(optimize Plan :profit :max [:chairs :tables]
          [[:chairs :>= 0] [:tables :>= 0] [:hours :<= 40] [:wood :<= 11]])
;; => {:chairs 5 :tables 6}
(call Plan D1 :profit)              ; D1 holding the optimize above
(optimize Priced :margin :max :price)               ; no constraints
(optimize Priced :margin 300 :price [10 :<= :price :<= 20])  ; reach 300
(optimize Loan :amount :max [:amount :years]
          [[:payment :<= 500] [1 :<= :years :<= 7]])
```

* **`(optimize grid output goal inputs constraints)`**: `goal` is `:max`,
  `:min` or a number to reach; `inputs` a cell or a vector of cells;
  `constraints` (optional) a vector of `[cell op value]`, `op` one of
  `:<=`, `:>=`, `:=`, `value` a number or another cell; `[0 :<= :share
  :<= 1]` is two, and one constraint may be given alone. Cells are names
  or addresses in the grid, as for `call`. A constraint between an input
  and a number is a bound on that input.
* **The result** is a map of the inputs' values, keyed as you gave them,
  so it is an override map: `(call Plan (optimize …) :profit)` is the
  profit it makes, `(instance Plan (optimize …))` draws it, `(:chairs
  D1)` reads one. Like any formula it follows the template: change a
  formula in it and the answer moves. Nothing in the template changes.
* **Where it starts.** From the inputs' own values (an instance's, on an
  instance, §21: `(optimize (instance Plan {:chairs 3}) …)` starts
  there). When everything the output and the constraints read is linear
  in the inputs, the answer is exact (Excel's Simplex LP); otherwise it's
  the best near the start (Excel's GRG Nonlinear), and another start may
  find a better one. Answers are tidied: 39.99999998 is 40 when that's
  no worse.
* **When there's no answer** the cell is an error saying why: no values
  found that keep to every constraint (and the one missed most, by how
  much), an output that grows without limit ("add a constraint or a
  bound"), an output that doesn't depend on the inputs, a constraint on
  cells the inputs don't move that doesn't hold, two bounds that can't
  both hold, or a cell erring at the start (named, as `call` names it).
  The search runs within the cell's time limit (§8).
* **Solver… on the cell menu** of any formula does the same for the
  sheet, as Excel's dialog does. *Make it* the largest, the smallest or
  equal to a number; *By changing* the cells holding numbers to change
  (names, cells or ranges, between commas; it offers the numbers the
  formula depends on); *Subject to* one constraint per line, as a
  formula in that cell would write the references: `B2:B4 >= 0`, `B9 <=
  100`, `0 <= share <= 1`, `B5 = C5` (a range against a number or a cell
  is each of its cells; `<` and `>` between spaces mean ≤ and ≥, since
  `>` is also a path, `B6>C2`). Solve shows what it
  found beside what the cells hold now, and each constraint as those
  values make it (✓ or ✗); *Keep these values* types them in, one undo
  step. Refusals are said in the dialog: a cell to change that holds a
  formula or a grid, an output that doesn't depend on them, a constraint
  that can't hold, whole-number constraints. What you asked is
  remembered for that cell until the page is reloaded; to keep a model,
  write it as `optimize`.

**Limits.** No whole-number (`int`, `bin`) or all-different
constraints. The answer is a local optimum when the model isn't linear,
and the output must change smoothly near it. The page waits while the
menu's solver runs (it stops after 15 seconds, showing the best found).

## 40. Columns that make room

What's drawn whole in a cell gets the room to show all of it: its
column is as wide as it needs, and its row as tall. That's a grid shown
by `(grid-view)` (a formula's table under it is one, §22), a formula's
2-D result kept in its cell (the renderer `table` or `table-of`), and an
instance (§21). So as it
grows (§38), it pushes the columns to its right along, and the rows
below it down, instead of scrolling or being cut off inside its cell,
and as it shrinks they come back. With a grid in A1 holding three
columns, column A is as wide as three of the grid's columns; when a
result grows the grid to six, column A doubles.

* **Its own size, or what's drawn whole, whichever is more.** A column
  keeps the width you gave it (dragging its edge, or the default), a
  row the height you gave it (or fits its content). What needs more
  widens or lengthens it; a grid that needs less stretches to fill it
  (§6). The widest (or tallest) thing drawn whole decides.
* **A result instead of a spill.** A formula's 2-D result spills into
  the cells below and to its right, and runs into what's there (#SPILL,
  §13). Give it the renderer `table` (or `grid-view`) and it's kept in
  its cell as a box instead, which pushes its neighbours aside: as wide
  as its text, as many rows as it has.
* **All the way up.** A grid inside a grid inside a cell widens its own
  column in the middle grid, which widens the middle grid, which widens
  the sheet's column. A result inside a grid does the same.
* **What it takes.** A grid: its columns as drawn and its row headings,
  or, for a grid with a display area (§6), only the area's columns,
  drawn clean without headings. A result: its table as wide as its
  text. An instance: its template's grid (its display area), without
  the line over it.
* **What asks for nothing.** A grid in a window, `(grid-view rows
  cols)`, which scrolls (a result is drawn whole by either); the link
  that zooms in; anything in a span (§4), whose block you sized
  yourself; any other renderer, including one of your own.
* **Seeing why.** Hovering a held column's heading says what holds it
  and how wide it needs to be. Dragging a column's or a row's edge
  smaller than that is refused next to the edge, saying why and what to
  do for what holds it: narrow or shorten the grid, or show it in a
  window; give the result fewer columns or rows, or take its renderer
  off to let it spill; narrow the instance's template, or give the
  template a display area showing less. So is a double-click on a
  column's edge (the default width) that leaves it held. Dragging it
  bigger sets its own size as ever.
* **Results are measured when drawn.** A grid's need is worked out from
  the sheet, so it counts wherever it is. A result's or an instance's
  is its text's width, measured when its cell is drawn, so a cell in a
  part of a big grid you haven't scrolled to yet (§9) counts once it's
  been drawn; scrolled away again, it still counts.
* Nothing new is saved: the sizes you set are, and the room taken
  follows from them and what's drawn.

**Limits.** Grids side by side don't line up: two grids in the same
column, or in columns next to each other, each stretch their columns
their own way, unless the mesh (§41) is on. A result spilled into the
sheet (no renderer) still runs into its neighbours, as in Excel.

## 41. The mesh: grids that line up

Grids drawn in cells make room for themselves (§40), but each stretches
its columns to fill its cell its own way, so grids stacked in a column,
or side by side, don't line up. Press **Mesh** (the lattice) in the format bar and they
do: every column and every row, in every grid at every depth, is a whole
number of tracks, so every line falls on one lattice.

* **Tracks.** *columns … px* and *rows … px* beside the Mesh button are
  the track width and height, 24 px each at first. Rows have their own:
  a row holds a line of text, and wider columns shouldn't make every row
  two lines tall.
  A column's width is read as tracks of 24 px, rounded: the sheet's
  default 120 px is 5 tracks, a grid in a cell's 72 px is 3. Change the
  track width and every column, at every depth, resizes alike: at 30 px
  the sheet's columns are 150 px and a grid's 90.
* **Grids in cells.** A grid drawn whole sits on its cell's left edge,
  with no frame or padding, at its columns' width; it doesn't stretch to
  fill the cell, and the rest of the cell is left empty. Its headings'
  column is one track, so its own lines are on the lattice too, one
  track in. A grid drawn clean (no headings, §6) gives that track to its
  first column instead, so its lines stay where they'd be with headings:
  switching a grid to clean, or putting a clean grid beside one with
  headings, keeps them lined up. Its column makes room for it in whole
  tracks (§40). An instance (§21) is drawn the same way, its first column
  taking the track too.
* **Rows.** A row is as tall as what's in it, as ever, rounded up to
  whole row tracks: one line of text is one track, three lines of a note
  three or four, a grid drawn in a cell as many as its rows, and its
  column headings one. A height you set is read as tracks too (30 px is
  one track of 24, 40 px is two), and is still the least the row is.
  A grid drawn whole sits at its cell's top, with nothing above or below,
  so its rows fall on the sheet's: two grids side by side line up row
  for row. A grid drawn clean gives its headings' row to its first row,
  and an instance's line over it (its template's name) is one track, so
  they stay in line with a grid with headings beside them. As a grid
  grows down or shrinks, the row holding it grows or shrinks by whole
  tracks. Change the row track and every row resizes alike.
* **Lines meet.** A 3-column grid and a 5-column grid in the same column
  count their tracks from the same edge, so their lines coincide where
  both have one. A grid that grows adds whole tracks and leaves the
  others in its column as they are.
* **Dragging a column's edge** goes by whole tracks, at any depth; the
  width kept is the tracks, as 24 px each. A double-click goes back to
  the default. A row's edge goes by whole row tracks the same way.
* **Seeing the lines.** Tick *show lines* beside the track width and the
  mesh's lines are drawn over the sheet, a faint line on every track's
  edge, across and down, over the cells and the grids in them. Widen one grid's column and
  the columns after it move off the others'; drag the next edge to the
  line you want (two tracks, say) and everything after it is back in
  line. The lines are a way of looking, kept in this browser, not saved
  with the sheet; they're off at first.
* **Lining up far apart.** With the lines showing, select whole columns
  (click a column heading), in the sheet or in a grid at any depth, and
  the tracks they cover are tinted faintly down the whole sheet, over
  every cell and grid in them: what lines up with the column far above
  or below shows. Whole rows (a row heading) are tinted across the whole
  sheet the same way. Without the lines, only the selection's own cells
  are coloured, as ever.
* **Edges among the lines.** With the lines showing, every other column
  heading and every other row heading is shaded a little darker, in every
  grid, so where each column and row begins and ends still shows.
* **Off.** Press Mesh again and every column is as wide as it's stored
  again, every row as tall as its contents or its set height, and grids
  fill their cells as usual (§6). The mesh is saved with the
  sheet, and undone like any change.

**Limits.** A 2-D result's own columns and rows (§40) are as its text
needs; its column and row round up to whole tracks. A grid in a window,
`(grid-view rows cols)`, keeps its frame and scrolls, so its own lines
are off the lattice (its column and row are on it). Rows are rounded
once they're drawn: right after a change, or as a big grid scrolls rows
into view, a row may show its natural height for a moment, and while you
scroll fast the rows coming into view are rounded when you stop. Moving
one mesh line, so that every column with an edge on it follows, isn't
there: the track sizes are one each for the sheet.

## 42. Charts like Excel's

Every kind of chart Excel makes can be made here without writing a
Vega-Lite spec: select the data and the **Shelf** tab offers the charts
that suit it; give a cell a chart and its **refine panel** changes the
kind and every option. The tour's *Charts like Excel's* section has one
of every kind, three to a band.

**Using it**

* **Making one.** Select a table under headings (or names or numbers
  beside columns of numbers) and open the Shelf tab: *As one chart*
  offers what suits the shape, first to last, as Excel's *Recommended
  charts* does (the first six drawn, the rest named on a button below
  them; *All kinds* places the first and opens its refine panel, which
  draws every kind to change it to). Columns, lines, pies, stacked and 100% columns, stacked areas,
  radar, a heatmap of a crosstab, columns and a line on two axes (when
  one series is ten times the other's size), bars (many or long names),
  Pareto (none negative), a waterfall (rises and falls), a funnel
  (falling), a treemap and sunburst (two columns of names), a box and
  whisker (names repeated), a stock chart (dates beside open, high, low,
  close), a bubble chart (a third column of sizes), points with a
  trendline, a map (a column of countries), and for a long column of
  numbers a histogram and a box. Numbers whose cells have a number format
  (`$#,##0`) are offered in it, on the axis, the labels and the hover;
  a column in a date format is drawn as dates. Numbers with a unit
  (§25) are offered in it, `{:up {:in "km"}}`, and drawn in it, the unit
  in the axis's title; a chart in a cell given a unit of its own (mi)
  shows them in that.
* **Changing it.** Click the chart, then the Renderer field (◇ under
  the formula). Its refine panel has, at the top:
  * *Kind*: columns, line, area, points, pie, donut; *More kinds*:
    histogram, Pareto, box and whisker, waterfall, funnel, stock, radar,
    heatmap, treemap, sunburst, map; each drawn with your data (hover
    one to see it in the cell). The columns are kept where they fit. A kind the data
    doesn't suit says what it needs, in the tile and the cell: a stock
    chart without price columns, or with a high below a low (which row);
    a column the data hasn't (and the ones it has).
  * *Variant*: side by side, stacked, 100%; *Horizontal* turns columns
    into bars (a combination too: its second axis goes to the top).
  * The columns read (*Across*, *Up*: one, or each a series), *Summary*,
    *Colour by*, *Size by* (a bubble chart), each kind's own rows (a
    pie's hole, first slice angle, gap and small slices as *Other*; a
    histogram's bins, width, underflow and overflow; a box's whiskers,
    mean marker and line, inner points; a waterfall's totals, connector
    lines and colours; a funnel's labels; a stock chart's prices, volume,
    days and colours; a radar's fill, markers, rings and maximum; a
    heatmap's labels and contours, filled or *Lines only (wireframe)*;
    a treemap's labels, layout and parent
    banners; a map's projection and area), and *Title*.
  * Then, in groups that fold, as Excel's Format pane has them:
    * **Axes**: titles, *Up from* and *to*, *Major unit*, *Logarithmic
      scale*, *Values in reverse order*, *Display units* (thousands,
      millions, billions), *Number format*, *Gridlines*, where the labels
      go (beside the axis, the other side, none), *Label angle*, *Start
      at zero*, *Up in unit* (km, mi, $ …: the numbers shown in it;
      empty, the cell's), *Across in unit* for numbers across, *Order*
      (the rows', largest first, a to z), *Dates across* and their unit
      (days to years; columns side by side over dates are a band a day,
      month or year as the dates are spaced, unless you choose).
    * **Elements**: *Data labels* (outside the end, inside it, the
      centre, the base) and what they show (the value, the category, the
      series, or a name and the value; inside a column, dark on a pale
      fill and white on a dark one); *Trendline* (linear, exponential,
      logarithmic, polynomial and its order, power, moving average and
      its period; a forecast forward and backward, the equation and R²,
      over numbers, dates or names: over names, fitted against their
      places 1, 2, 3 … as Excel does, a forecast adding blank places);
      *Error bars* (standard error, standard deviation, 5%, 10%, or a
      column of amounts); *Lines* (drop lines; with series on one axis,
      high-low lines and up/down bars, and the bars' colours);
      *Reference line* (a value, a threshold across the chart, or a
      column of the rows, a line at each place: each column against its
      own, a capacity, a target), its *label* (in the legend), and
      *Above the line* (not marked, or marked red, orange or green: in
      that colour, or outlined in it where colours tell series apart;
      on lines, a point at each place above).
    * **Series**: a row per series (drawn as columns, a line, an area or
      points; on the left axis or the right; its line solid, dashed or
      dotted): Excel's combo chart, each axis's series with their own
      labels, trendlines and error bars. *Switch row/column*; *Show* (a tick
      per series and per category, Excel's chart filters); *Empty cells*
      (left out, gaps, as zero, joined across); *Smooth* and *Markers*;
      *Gap width*, *Overlap*; *Invert if negative*.
    * **Colours and legend**: *Palette*, each series' colour and name,
      *Legend* (where, or none) and its title, or the one colour.
  Rows that wouldn't change this chart aren't offered (a centre for
  columns side by side, an across axis for a box of one column, labels
  on contours).
* **The call, written.** Every row writes an option of `chart`'s own:
  ```clojure
  (chart :bar :month [:north :south] {:stack :stacked :labels :center})
  (chart :bar :month [:north :south] {:labels {:at :inside :show [:series :value]}})
  (chart :bar :month [:sales :margin] {:series {:margin {:as :line :right true}}})
  (chart :line :day :close {:across {:dates true :unit :month} :up {:format "$#,##0"}})
  (chart :point :x :y {:trend {:poly 2 :forward 5 :equation true :r2 true}})
  (chart :bar :month :sales {:trend {:poly 2 :forward 2}})  ; over names too
  (chart :bar :day :run {:up {:in "mi"}})            ; numbers in a unit
  (chart :bar :week :days {:color :resource :reference {:column :capacity :label "Working days" :over true}})
  (chart :line :day :load {:reference 80})           ; a threshold
  (chart :histogram :qty {:bins {:count 10}})        ; one field, then options
  (chart :waterfall :step :amount {:totals ["Subtotal" "End"]})
  (chart :stock :date [:open :high :low :close] {:volume :volume})
  (chart :treemap [:region :city] :sales)            ; a vector of levels
  (chart :heatmap :region [:jan :feb :mar])          ; a crosstab
  (chart :heatmap [:x :y] :z {:contour 8 :wireframe true})
  (chart :map :country :sales)
  ```
  The axes are named by role, `:across` and `:up`, so their options
  stay with them when columns become bars. Columns coloured by a field
  (`{:color :region}`) stand side by side, as Excel's clustered columns
  do; `{:stack :stacked}` stacks them. Anything else in the options
  map is Vega-Lite's own, merged in.
* **Maps.** Countries by name (and other names they go by: USA, UK,
  South Korea …) or ISO code (FR, FRA or 250), coloured by the numbers
  beside them,
  over the world in grey (*Area: the countries in the data* zooms in).
  Names not found are listed on the map. The boundaries ship with
  RecDep (Natural Earth's countries at 1:110m, from world-atlas, ISC
  licence), loaded the first time a map is drawn: nothing is fetched
  from elsewhere. The Shelf tab offers a map for a column of names or
  three-letter codes, not two-letter ones (CA, DE, IN are US states as
  often as countries); `(chart :map :code :sales)` draws them.
* **Sparklines.** `win-loss` (a column up for a gain, down for a loss);
  `(sparkline [:high :low :first :last :negative])` and `(bars
  [:negative])` mark those points, as Excel's markers do (the Renderer
  field's *Markers* row).

**Limits.** No 3-D charts (the 2-D kind instead), pie of pie (*Small
slices as Other* instead), exploding one slice (a gap between all),
Excel's data table under a chart, minor gridlines, an axis crossing at a
value, or a trendline's set intercept. A chart takes its numbers to be
of the unit's kind: shown in kg, metres would be numbers, not an error.
Maps know countries, not regions within them (states, provinces).


### Charting a cross-table of dates

A table with dates (or numbers) across its first row and names down its
first column, a week's figure for each person, say, is charted as Excel charts it: the dates across, a series
per name. Select it, or give its cell a chart renderer, and the Shelf
and *Show as…* offer lines, points, stacked areas and columns, side by
side and stacked. It's told apart from names beside numbers (no
headings) as Excel tells them: the corner cell is blank, or the
headings read as dates (whole numbers a day, a week or about a month
apart). Columns over dates are drawn a band each, not as slivers at
each date, labelled with their own date.

## 43. Table steps: the Data tab's verbs in a formula

A table worked out step by step, as the Data tab's steps do it, but in a
formula that stays live: filter, choose columns, sort, remove
duplicates, group, skip or keep rows, a running total, a rolling
average, the previous row's value, a rank, the rows from a date on,
clean, change type, split, merge, pivot, unpivot. The steps are the
Data tab's, with the same names and options, so a step means the same in
both. The tour's *Table steps* section is a
rolling forecast's window.

**Using it**

* **From the Shelf: Start steps.** The Shelf tab starts with a line for
  it, whatever is selected: *Steps on* a drop-down of tables, and **Start
  steps**. The first table offered is the one you've selected (a cell
  holding a table, or a range with headings), else the one you're in (a
  cell of Tasks offers Tasks), then every named table in the sheet. Start
  steps places `(steps tasks [])` as the Shelf places a view: an outline
  beside the selection (beside the table's own cell when you're inside
  the table) that Enter or a click puts down. Then the Inspector shows
  its *Steps* row, with **+ Step…** marked: choose the first step there.
* **From the Shelf: As steps.** Select a table (with headings or without)
  and the Shelf tab's *As steps* also offers, when the table has a column
  of numbers, its running total, rolling average and rank. Place one as a
  view is placed. Placed, a steps table (or a view)
  has its source's column widths: each column as wide as the source's
  column of the same name, a new one (a running total) the source's
  usual width; change them as any grid's.
* **In the Inspector.** Select a cell holding a steps formula: *Steps*
  lists them, numbered, one to a line: first the table they start from
  (0), then each step. Click a step and two things happen: the sheet
  shows the table after that step in the cell's place (step 0: before
  any step), over the cells around it, outlined, under a line saying
  which; and the step's form opens below it in the list. Click it again
  (or *show all*) for the whole table, the form closed; select another
  cell and the same. Only the drawing changes: the cell's value, and
  whatever reads it, stay as they are. In the form, besides its own
  fields: **Why**, a note saying what the step is for; **▴ Earlier** and
  **▾ Later** move it one place; **× Take out** removes it. **+ Step**
  adds one: any of the Data tab's, a worked-out column, a join or an
  append, a column in dependency order or a schedule (§49), or steps
  saved in the Data tab. It adds at the end, or *here*: click a step
  first and it reads *+ Step here, after step 3…* (click 0, the table,
  for *before step 1*), and the new step's form opens in the list at
  that place, on the table the sheet shows there. So a step can go
  before one that refuses, to fix what it refuses: added at the end
  while a step refuses, its form says which step and offers *Add it
  before step 8 instead*. Each change rewrites only that step in the
  formula, so the rest stays as you wrote it, and Undo takes it back in
  one. A step that refuses is marked in red, and why is said under the
  list, with *Change it* to open its form. An *own values* step (typing
  into the table, below) has no form: a click shows the table after it.
  Neither has a step that can't be read, nor one worked out by a
  formula rather than written as a map: under any of these, a click
  says why there's no form, with **▴ Earlier**, **▾ Later** and **×
  Take out** all the same.
* **Steps inside a formula.** The steps needn't be the whole formula:
  `(record-view (query "orders" [step …]))`, `(:total (first (query
  "order_summary" [step …])))`, `(IFERROR (count (steps orders [step
  …])) "none")` or a `let` around them all have the same *Steps* list.
  Clicking a step shows the table after it in the cell's place, worked
  out where the steps sit (with the `let`'s names; inside a `for`, the
  last time round, which the line above the table says), whatever the
  cell itself shows: a form, a number, the text IFERROR gives. A step
  that refuses inside IFERROR, which the cell doesn't show, is marked
  and said once you've clicked a step. Changing, moving, adding and
  taking out steps rewrite only the steps; the formula around them
  stays as written. A formula with several lists of steps (two queries,
  or a join's other table written as steps) names each above the list:
  click one to list it. In the *Trace* row, **Step through** gives each
  list its own **↘ its steps**, which lists it and brings it into view.
  Own values are typed only into a steps formula's own table, so they
  stay with steps that are the whole formula.
* **What a step reads.** Above an open step's form, a line names the
  cells and ranges the step reads, each a link to it: the Schedule step
  *Reads settings>level, settings>calendar*, a cost worked out from a
  rate table *Reads resources>resource, resources>rate*. A step reading
  no cell has none; the columns it picks are those of the table coming
  in, which the step above it shows.
* **Which row a worked-out column refused.** A formula of the row that
  fails on one row says which: *step 1, worked-out column drawn, the row
  with ID 13 (Testing and fixes): pert: needs lo ≤ most likely ≤ hi, got
  9, 6, 15*: the row's ID, with its task or name when it has one, else
  its place (*row 2*).
* **A name the sheet doesn't know.** A name in a step's formula that is
  neither a column, a cell's or a function's is said with the step it's
  in, as a missing column is: *step 9, worked-out column vat: there's no
  function, name or cell called vat-rate here*, the step marked in the
  list, with the columns there are and the names it nearly matches,
  columns first: *Did you mean work? The columns are :id :task …*. A
  missing column's refusal says the names it nearly matches too (a cell
  name misspelt, `vatt`: *Did you mean vat?*).
* **A note for each step.** What you type in a step's *Why* is kept in
  the step, `{:op :filter … :note "Only the weeks with something
  booked."}`, and the list shows it in the step's place, so a sequence of
  steps reads as an explanation; hover it for the step itself. Every
  form keeps it, and so do moving the step and applying its form. Clear
  the field to show the step itself again.
* **In a protected cell** (§47), the list works as ever, and the forms
  open so you can read them, with Apply, ▴ ▾ and × switched off and the
  line beside them saying to open the lock 🔒 above the grid first.
* **Dates stay dates.** A column of dates (its kind, §19: Tasks' Started
  and Finished) is still one after any step that keeps it: filtered,
  sorted, joined, chosen and renamed. A worked-out column is a date when
  its formula gives one by a date's own means: a date column, `WORKDAY`,
  `EDATE`, `DATE`, `TODAY`, the later or earlier of dates (`MAX`, `MIN`),
  a date plus or less some days, or an `if` whose every answer is one
  (`NETWORKDAYS` counts days: a number). A schedule's start and finish
  are dates, and so is a series' period. So the table after a step, the
  Inspector's *Value*, a table laid out in cells (by `steps`, `as-grid`
  or `(grid-view)`), and one kept in its cell by `table` show them as
  dates: in their column's (or the cell's) own format when it has one,
  else as `2027-01-04`. A
  table typed in as a value (`[{:start 46391}]`) has no kinds: format its
  columns.
* **A step's values.** In a step's form, a value is typed as it is
  (`North`, `40`), or in quotes as text is elsewhere in the sheet
  (`"North"`, the same); after `=` it's a formula (`=D1`, `=as-of`,
  `=(EDATE as-of -6)`), and with the field focused, clicking or
  dragging over cells writes one (`=D1`): the table then follows the
  cell.
* **A worked-out column** (+ Step → *A worked-out column*): a name and a
  formula of the row, its columns by name, `(* qty price)`, `(str
  region " " month)`; written as `(fn [{:keys [qty price]}] (* qty
  price))`, and shown on its chip, and in its form when edited again,
  as you typed it (`probe = 1`). A function typed whole, `#(* 2 (:qty
  %))`, is kept as it is. A name in the formula that isn't a cell's,
  a named cell's or a function's is a column, read from the row
  wherever the step is: `(if (> finish deadline) "late" "on time")`
  reads both from the row, so moved above a step that drops Deadline it
  works at once. A name the table hasn't got (a heading cleared, a
  column deleted) is refused, naming it and the columns there are,
  rather than read as blank: retype it in the expression (a heading
  retyped is followed, §19). When an earlier step dropped it, the
  refusal says which: *deadline was dropped by step 8 (Choose
  columns). Keep it there, or move this step above it* (renamed there:
  *renamed due … Write due*), and under the list *Change step 8* opens
  that step's form and *Move it above step 8* moves this one; if step 8
  then leaves out the column this step makes, the list says so. For a
  column only some tables have, give it a default: `(fn [{:keys [done]
  :or {done 0}}] …)`, or `(:done % 0)`. A column only some rows have
  (rows appended from another table) is a column of the table and is
  fine. After *A new column … =*, choose what it holds: **a formula of
  the row** (the above), **a value, the same for every row**, or **a
  copy of a column**. A value is typed as into a cell: `1`, `12.5%`,
  `Working days`, `15 Feb 2027`, `true`, and `"12"` in quotes for text
  that would otherwise read as a number or a date; or after `=` a cell
  or a name (`=settings>status`, or click the cell), worked out once and
  the same in every row. Under the field it says what it reads as (*the
  date Mon 15 Feb 2027*). A copy is a column's values under the new
  name, the column itself staying: written `{:op :compute :name
  :capacity :from :units}`, it reads *capacity = units* in the list, and
  a value `{:op :compute :name :resource :value "Working days"}` reads
  *resource = "Working days"*. Whichever you choose, the form writes the
  plainest step for what you typed: a formula that is a column alone
  (`units`) is a copy, a number or text alone a value; a value after `=`
  naming a column is a copy, and one reading the row's columns is a
  formula of the row. A formula calling a function (`(RAND)`) stays a
  formula, worked out row by row. **A join**: the other table (its name, or click or drag over
  its cells, headings and all), every row here or only those matching,
  or the rows with a match or without one; the pairs of columns that
  match (*compare as text* matches the number 3 with the text `3`, as
  IDs typed in two tables often differ), the columns to bring in, and a
  prefix to name them with (*baseline*: start comes in as
  *baseline-start*). Under the pairs, how the rows match: *17 of 17
  rows match*. **When a row matches several there** (a baseline kept
  twice, a customer with two addresses), each match makes a row, so
  rows repeat, and the join says so: in its form (*Every row matches
  more than one row there (2 each), so rows repeat: 17 become 34*), on
  its step in the list (⚠, and *Step 1, join: … Change it* under the
  list), and over the table when that step is shown. Its form's *When a
  row matches several there, keep* chooses **every match** (a row
  each), **the first match** or **the last match**, in the other
  table's order, as XLOOKUP does: the last of several baselines kept is
  the latest. Or filter the other table before the join (by `taken`,
  say, for another one). It can also keep **their sum**, **their
  average**, **the smallest** or **the largest**: a row for each row
  here, each column brought in summed (or so) over its matches, as
  *Group and summarise* sums; a row matching nothing gets a sum of 0
  (the others blank), and a column of text there is refused, since it
  can't be summed. **A list matches item by item**: where the column
  here holds lists, as a *one or more of* column does (Resource:
  `Ana, Ben`), each item is matched, so joining Resources on it,
  bringing in Rate with *their sum*, gives each task the sum of its
  people's day rates, and a worked-out column `(* days rate)` its
  cost. The form says so under the pairs (*resource holds lists: each
  item is matched*), and an item matching nothing is said there, and on
  the step (⚠ *1 list item matches nothing there (Bne, row 4)*), since
  it would otherwise just make the sum smaller. Items are read as
  everywhere else: ignoring case, and `3+2d` as 3 when there's a 3.
  Text with commas in a plain text column is one value. The Data tab's
  join keeps every match: there, remove the other table's duplicates
  first, or group it to sum. **An append**: the other table's
  rows after these, its columns matched by name or in order. The other
  table is a choice of the sheet's named tables, or another typed (a
  reference, `(as-table F1:H9)`, or click or drag over its cells); under
  it, its rows and columns, or why it isn't a table, and then the match:
  *3 columns matched; new, added after these: weekdays; not in it
  (blank there): role, units* (in order: *Who → resource, …*). **Apply
  records that match** in the step (`:match [:resource :from :to :why]`,
  and `:new` for the columns added). If either table's columns change
  afterwards — a heading cleared and typed differently (*Resorce*), a
  column added or taken out, two headings swapped — the step stops with
  an error saying what changed: *days-off has Resorce where Resource was
  matched*. Its line under the steps, and a box by the heading as you
  type it, offer **Rename it back** (the heading as it was) and **Accept
  the new match** (the step records the columns as they are now); each
  is one undo. Retyping a heading is followed and needs neither (§19).
  An append written by hand without `:match` works as before, and its
  step says *match not recorded*: open it and Apply to record it.
* **A series: rows made from a window** (+ Step → *A series: a row per
  day, week or month*): a new column, a row per day, week or month, from
  a first day for so many periods or until a last day. By default each
  row of the table becomes a row per period, with the period in the new
  column: a list of people × the weeks ahead, as the project plan's
  load has. *Once* gives the periods alone. A period's value is its
  first day: a week's Monday, or, given a calendar (chosen as the
  Schedule step's is), the day after its weekly days off (as `week-start`'s weeks); a month's 1st. *From* and
  *until* take a date (`2027-01-04`, `4 Jan 2027`), a cell or name (click
  it), a formula, or a column of the row by its name, which gives each
  row its own window: a booking's Start and Finish make a row for each
  day it books. As a formula: `{:op :series :name :week :every :week
  :from from :count weeks :calendar ..>settings>calendar}`, or `:to`
  for the last day, `:once? true` for the periods alone.
  * **With a calendar, a series of days is its working days**: the
    days with any capacity, for the row's resource (below), so a
    person's leave and the weekend aren't rows and a Saturday of
    overtime is. *For this many periods* then counts working days, as
    WORKDAY does.
  * **What the calendar has in each period.** Under the calendar, the
    form asks *whose*: a column of the row naming them (Resource), or
    everyone's rows. Then it can add, each named as you
    like or left blank: *the period's working days* (each day's capacity
    summed, a half day a half, as `working-days` counts them); and,
    given *what's left of it* (the calendar with the plan's bookings),
    *the days booked* (the working days less what's left) and *the days
    left* (below 0 when it's booked twice). A week is its seven days and
    a month all of it, whatever the window. As a formula:
    `{:op :series :name :week :every :week :from from :count weeks
    :calendar cal :who :resource :capacity :capacity :left left :booked
    :days}`, `:free :left` for the days left. With no calendar, the
    working days are Monday to Friday. A resource × week table of
    capacity and bookings needs no formula of the row.
  * Refused, saying what to give: days booked or left without *what's
    left*; *what's left* with neither; *whose* without a calendar, or
    naming a column with *once*; a *whose* column the table hasn't got
    (naming the columns it has); a calendar that isn't one; a series of
    working days with none in ten years for someone.
* **A pivot's headings are its values.** Each column a pivot makes is
  headed by the value it's for, as written: `4 Jan`, or a date, which
  the heading row's date format shows as one (format the row, `d mmm`,
  as any heading). A column is read by name as its heading reads,
  lower case with `-` for spaces: `by-week>_4-jan`, or for a date its
  number, `by-week>_46391`. Without `:keys` a pivot makes a column for
  every value the data has, in the order each first comes; with
  `:keys`, those (a list, or a formula giving one). A pivot of dates,
  names down its first column, charts as a cross-table does: a series
  per name, the dates across.
* **Typing into the table.** Type a value into a cell of a steps
  table and it's kept with its row's record, not typed over the result:
  the formula gains an *own values* step, `{:op :own :key "item"
  :values {"plum" {"qty" 99}}}`, keyed by the table's first column. The
  value stays with plum when rows are added to the source, reordered or
  filtered; a worked-out column gets its exception for that record.
  Empty the cell to take it out. Keys match as text does here, ignoring
  case. A key on two rows, a row without one, or typing into the key
  column itself is refused, saying why. A record filtered out keeps its
  values, doing nothing until it's back; the Inspector's Steps row says
  which, with *Drop them*.
* **Keeping each group's rows.** *Group and summarise* has, in the
  sheet, *keep each group's rows, as a table in the column …*: beside a
  group's summaries, a column holding its rows (without the columns
  grouped by), each drawn as a small table in its cell. Name the cell
  `summary` and `summary>rows` is that column. + Step → *Unnest* turns
  it back into rows: one for each row of each table, with its group's
  other columns. As a formula: `{:op :group :by [:region] :nest :rows}`
  and `{:op :unnest :column :rows}`. The Data tab's steps don't keep
  rows (its tables hold values only); its joins keep the rows with a
  match, or without one, there too.
* **As a formula:**

  ```clojure
  (steps actuals
    [{:op :filter :column :region :test := :value region}
     {:op :window :fn :running-sum :column :amount :as :to-date :order [[:month :asc]]}
     {:op :window :fn :rolling-avg :column :amount :n 3 :as :avg-3}
     {:op :limit :from {:column :month :test :>= :value as-of} :n 6}])
  ```

  A step's values can be cells or names (`region`, `as-of` above): the
  table follows them. Columns are named by keyword (`:amount`) or by
  their heading's text (`"Unit price"`). The result is a grid of its
  own with a header row: name the cell `window` and `window>to-date`
  is a column, `(last window)` its last row. Typed by hand, the cell
  shows a link to its grid (▦ … grid ›, a click zooms in); give it the
  renderer `(grid-view)` to draw the table in place, as the Shelf does.
* **The order steps.** `{:op :window :fn … :column … :as …}`, with
  `:order [[col :asc]]` (else the rows' order) and `:by [cols]`
  (restarting for each): `:previous` and `:next` (`:n` rows away),
  `:running-sum`, `:running-count`, `:running-avg`, `:running-min`,
  `:running-max`, `:rolling-sum` and `:rolling-avg` (over `:n` rows, this
  one included), `:rank` and `:dense-rank` (1 the largest), `:row-number`,
  `:share` (of the total). `{:op :limit :from {condition} :while
  {condition} :skip 2 :n 12}`: the rows from the first passing `:from`,
  then as long as `:while` holds, then the skip and the count. The Data
  tab has both too: + Step → *Running, rolling, rank…*, and *Skip or keep
  rows* → *+ from the first row where…*, *+ while…*.
* **A view as steps.** A `filter-view`, `group-view` or `pivot-view`
  cell's Inspector has *Edit as steps*: the view rewritten as the steps
  its choices stand for, to change and add to (a cross-table's Total row
  and column aren't a step, and are left out). A filter view's search is
  a filter on any column containing the text, its page a step keeping
  rows; a group view's levels are one grouping by them all (what's
  folded is how it's shown, not a step). `record-view` and
  `profile-view` aren't steps.
* Text compares ignoring case (as the views' filters do); a number or a
  date written as text (`"40"`, `"2025-07-01"`) compares as one with a
  column of numbers or dates; a number never compares with text (`"n/a"`
  isn't more than 50, nor less).

**Limits.** A SQL step and *⇄ Rows left out* are the Data tab's only.
Steps move one place at a time (▴ ▾), not by dragging. A kept group's
table can be read, not zoomed into or typed in. A value that
should start with `=` as text is typed in quotes (`"=x"`).


### Web answers: opening what's nested

A web service's answer (§56) is data: records inside lists inside
records. Steps take it as it is, and open it a step at a time, as Power
Query's *Expand* does.

* **Any data is a table.** `(steps (web url) [])` shows the answer as a
  table:
  * a record is one row, its fields its columns;
  * a list of records is a row each;
  * a list of values is a row each, in a column `value`.
* **⤢ on a heading opens a column.** Click the steps' *0* in the
  Inspector (or a step) to see the table there. A column holding lists
  or records is summed up, *{ 5 fields }* or *[ 63 records ]*, with ⤢ in
  its heading. Click ⤢ to add the step:
  * **a list of records:** a row for each record, its fields as columns,
    the row's other columns repeated;
  * **a list of values:** a row for each value;
  * **a record:** its fields as columns in its place;
  * **lists side by side,** as long as each other in every row
    (timestamps beside prices, as many services send them): opened
    together, a row for each position.

  The table after the new step is shown, so you go on clicking. On a
  web cell whose formula isn't steps yet, the Inspector's Web row offers
  *⤢ Shape the answer by steps*, which wraps it and opens the table.
* **A record keyed by names,** `{"AAPL" {…} "MSFT" {…}}`, is a row with
  a column per name. *⤢ A row for each column* (an unpivot of every
  column) makes a row per name, its record in `value`, for ⤢ to open.
* **The steps as written** (each a chip in the Inspector, changed by
  its form):
  * `{:op :unnest :column :news}`;
  * `{:op :unnest :column :chart :path [:result 0]}`, reaching in first;
  * `{:op :unnest :column :quote :keep [:close :datetime]}`, only those
    fields;
  * `{:op :unnest :column :quote :prefix "quote"}`, naming them
    `quote-close`… (a field with the name of a column already there is
    refused until it has a prefix);
  * `{:op :unnest :column [:timestamp :open :close]}`, lists side by
    side;
  * `{:op :unpivot :columns :all :name :symbol :value :quote}`.
* **Values as they should be.** `{:op :type :column [:open :close] :to
  :number}` changes several columns at once: numbers sent as text become
  numbers. `{:op :type :column :t :to :date :from :unix}` reads Unix
  times (seconds since 1970; `:unix-ms` for milliseconds) as a date and
  time where you are. In a formula, `(from-unix t)` does the same.

## 44. Repeat: a cell as one formula

Fill down (Ctrl-D) copies a formula into each cell: copies that drift apart
once one is edited. *Repeat* writes one formula in the cell that holds
them all, so an edit to it reaches every cell.

**Using it**

* Select a cell and the cells below it (or to its right), or the cell
  alone (to the grid's edge), and open the **Shelf** tab: its first
  section, *Repeat C2 as one formula*, shows what it would write and
  sketches the repeat: the cells it fills, the first holding the formula
  (ƒ), each beside the cell it reads (B2 → C2, B3 → C3, ⋮), and how many.
  A block, or the cell alone, can go either way: ↓ or →.
* **A formula**: its moving references become ranges where that gives
  what Fill would, `(* B2 2)` repeated down three cells is `(* B2:B4
  2)`; where it wouldn't (text, Clojure's functions), it's Excel's `MAP`,
  `(MAP B2:B4 (fn [b2] (str "x" b2)))`. Each reference can **move** or
  **stay**, as `$` says for Fill: the choice is on the Shelf, beside it.
* **A value**: the same value over the cells, `(down (repeat 3 "North"))`,
  or a series where it starts one, `(down (take 3 (series "Jan")))`
  (§15).
* A click writes it in place: the cell gets the formula, the other cells
  are emptied for it to spill into, with the cell's number format. One
  undo step.
* Once repeated, the Shelf says the cell already holds them as one
  formula: change it there.
* Refused, saying why: a range that moves (`(SUM $B$2:B2)`, a running
  total: Fill copies it, or use a steps formula's running total, §43, or
  make it stay); a reference moving into the cells it fills (a cycle).

**Limits.** The cell alone repeats to the grid's edge as it is now; a
formula's ranges don't grow when the grid does. A repeat's cells can't
be overtyped one by one: Fill's copies can.

## 45. Charts that pick

A chart can be a control: drag across it, or click a column, and what
you picked goes into the chart's cell, for other cells to read. The
standard example is an overview and a zoom: a long line chart you drag
across, and a second chart showing just the days you picked.

**Using it**

* Give a chart `{:pick :across}` and its data with `:data`, in an empty
  cell: `(chart :line :date :price {:data prices :pick :across :across
  {:dates true}})` as the renderer of B3. Or click the chart, then the
  Renderer field, and choose *Picking* in its refine panel: the cell's
  formula moves into `:data` for you.
* **Drag across** a chart with numbers or dates across to pick a range;
  drag the grey band to move it; click outside it to pick nothing. The
  cell holds `[from to]` (dates as date numbers, whole days when the
  data's are).
* **Click a column** of a chart with names across to pick that name (the
  others fade); click beside the columns to pick none. The cell holds
  the name.
* The cell's value is what's picked, so anything can read it.
  `(picked prices :date B3)` is the rows of prices whose date is in B3's
  range, or that hold B3's name; every row while nothing is picked. Draw
  it with an ordinary chart for the zoom: `(chart :line :date :price
  {:across {:dates true}})`. Or average it, filter a table with it, give
  it to an instance.
* Each drag or click is one undo step. Typing a range into the cell,
  `[45750 45780]`, moves the band.
* **In a panel**, bind the chart to an option as any control: `(input
  [:zoom] (chart :line :date :price {:data rows :pick :across}))`; a
  drag rewrites the panel's formula, `{:zoom [45720 45740]}`, and the
  panel can draw `(picked rows :date (:zoom opts))` beneath it.
* Columns, bars (sideways too), lines, areas and points can pick. A pie,
  a combination on two axes and the other kinds say so in the chart.

* **A view's picked rows.** A view's value is what it picks too: a
  Gantt chart's row clicked (none until one is), a record view's row, a
  filter view's rows. `picked` reads them: `(picked tasks :id chart)` is
  the rows of tasks whose ID is one of the IDs picked (4 and `"4"`
  match), none while nothing is picked, saying *Nothing picked: click a
  bar to pick its row*. Show them with `(record-view (picked tasks :id
  chart))`: the rows are the table's own, so typing in them writes
  that table (§48). A column the table, or the rows picked, haven't got
  is an error.

**Limits.** One range or one name per chart, across only. Two charts
can't share one picked cell. Clicking a row of a filter view to pick it
isn't there yet. The zoom shows the rows inside the range, so its line
stops at the last day inside rather than running to the edge. Picking
nothing writes `nil` into the cell.

## 46. Sheets on the web: examples, and links that open one

RecDep comes with example sheets beside the guided tour: worked models
to use as they are, or as a start. They're ordinary saved sheets, kept
where the app is served, and any sheet saved on a web server can be
opened the same way.

* **Examples ▾** (the light bulb at the toolbar's far right) lists them, each with a line saying what it
  is: first the **Guided tour**, the sheet RecDep opens with the first
  time (it's built into the app, so it's there offline too), then the
  others. Click one and it opens in place of your sheet, as
  Open does (Undo doesn't bring your sheet back: save it first). It's yours to change; Save asks
  where to keep it. The first is **Project plan**: tasks, resources
  and roles in, a schedule out that keeps to the links between tasks
  and to what each resource can do, named baselines compared two at a
  time, tracking against the latest with earned value, a Gantt chart,
  weekly loading, and how sure its finish is (Schedule risk), built
  from blocks. The second is **Order book**:
  a wholesaler's customers, products and orders kept in the sheet's own
  database (§53), with sheets that are forms onto it. *Find customers*
  and *Find orders* search as you fill in their boxes (leave one blank
  for any); a customer's or an order's number opens it on *A customer*
  or *An order*, where typing over a detail changes it in the database
  and Undo changes it back; an order's lines can be added and removed;
  browser Back returns to the search. *Shapes* says what each table
  holds. A green frame shows where to type: the boxes, and the columns
  of the database grids you can type over. The form sheets are
  protected (§47), so everything else on them refuses typing: their
  tabs carry a lock, and the bar above each says so. It's a
  `.recdep.db`, so Save keeps the database with it.
* **Open from a web address…** (at the bottom of Examples ▾) opens a
  sheet (`.recdep.edn`) or an Excel workbook from any address, such as
  a file in a GitHub repository (`https://raw.githubusercontent.com/…`)
  or on GitHub Pages. If it can't, it says why: *the server answered 404
  Not Found (there's no file at that address)*, or that the page couldn't
  read it, because the address is wrong, you're offline, or the server
  doesn't let other pages read its files (most file shares don't; GitHub
  does).
* **A link that opens a sheet.** Add `?open=` and the sheet's address
  to the app's address, and the page opens that sheet when it starts:
  `https://klibro.github.io/recdep/?open=examples/project-plan.recdep.edn`
  (an address relative to the app, or any full address). Over a sheet
  of yours (the one the browser kept from last time) it asks first;
  Cancel keeps yours. It opens once:
  the address bar then drops the `?open=`, so reloading the page keeps
  your edits instead of opening the sheet again. A cell to go to can
  follow, as always: `…?open=examples/project-plan.recdep.edn#B4`.

### The project plan

How the Project plan example is put together, to use it or to build
your own. It is built from **blocks** (§6, *Inputs that take a
concept*): the sheet holds the data and the dates, and each block is a
self-contained grid saying what it takes and gives, bound to them or to
another block's outputs. Select a block's cell and the Inspector's
Inputs show its bindings; its body holds only its own data and
workings. The summary at the top reads the blocks' outputs, never their
insides.

* **Settings**: the start date, the *status date* (the day the actuals
  are as of; work not done is scheduled after it) and *Level resources*
  (ticked, a resource does one task at a time).
* **Resources**: each one's name, *units* (1 full time, 0.5 half) and
  day rate. A machine is a resource too: the device lab.
* **Tasks**: an ID, the task, its phase, its *work* in days (the most
  likely), its *Low* and *High* (the least and the most it could take,
  for Schedule risk: a Low above Work, or a High below it, is refused
  where it's typed), who does it, what it waits for (*After*: IDs, `3, 4`; a lag in working days,
  `12-2d`), a *Not before* date, a *Priority*, and, as work happens,
  *Started*, *Finished* and *Done*. Who does it is asked for in three
  ways: *Resource* names people or things, *Role* asks for any one of
  a role's people (the one finishing first), and *Follows on* gives it
  whoever got another task.
* **Roles**: a row per resource per role, so Ana is a designer and a
  developer.
* **Working calendar** (a block): takes the resources, keeps its layers
  as its own tables (the working week, public holidays, company
  holidays, days off, overtime: §51) and gives them as one calendar,
  said to be capacity.
* **Timesheet** (a block): a row per person, day and task, *Spent* in
  days; Task is picked from the tasks with the ▾. It gives effort.
* **Scheduler** (a block): the Schedule step over the tasks, into the
  calendar, from the start and status dates, a role's people from the
  roles, who is on a task begun from the timesheet. It gives the plan,
  its bookings, and what they leave of the calendar.
* **Baselines** (a block): keeps baselines of the plan by name, and
  compares any two.
  * **Keep one**: type a name in *Name* (*Re-plan*) and press *Keep the
    plan now as a baseline, by that name*. A new name keeps it beside
    the others; a name kept already replaces that one, and the button
    says so (Undo brings it back). With no name, the button shows ⚠
    and says a copy needs one.
  * Its **Compare** grid lists them: *Baselines*, a row each: its name, the day it was
    taken, its tasks, its last day and its cost. The example comes with
    two: *Original*, and *Approved*, taken later, with a Style guide
    task added and a finish a week later.
  * **The baseline** is the latest kept: what Tracking and the summary
    follow (*Baseline taken*, *Baseline finish*, *Slip*). Keep another
    and they follow it.
  * **Compare … with …**: pick two from the ▾ beside each, any two
    baselines, or one and *Plan now*, the plan as it is. *Comparison*
    puts them side by side, task by task: each one's start and finish,
    how far the start and the finish moved in working days (on the
    calendar), the work and the cost, and what became of the task:
    *added* (only the second has it), *gone* (only the first), *later*,
    *sooner*, *moved* (its start, not its finish) or *same*. *As a
    Gantt* draws the second's bars over the first's (thin and grey),
    coloured by what changed; a task gone from the second shows its grey
    bar alone. Hover a bar for both dates, each under its baseline's
    name. A pick no longer kept is said beside it.
  * Its workings are steps (§43) under the comparison: the plan now and
    every baseline as one table, then the rows of each picked. The
    Compare grid is the block's own view, not what it gives (its name,
    `_compare`, is private: §6): the block gives the baseline and when it
    was taken.
* **Tracking** (a block): the plan beside the baseline, task by task:
  the slip in working days, and earned value at the status date
  (planned value, earned value, actual cost). Time spent is the
  Timesheet's effort, summed by task in its binding.
* **Gantt** (a block): the plan as a chart, the baseline's bars beneath,
  the part done darker, the status date in red. Click a bar: under the
  chart, that task's row of Tasks, to change there and then; the plan
  and the chart follow.
* **Load** and the two loading views: each resource's days booked
  beside its working days, week by week, from the week of the status
  date for *Weeks shown* weeks; one view draws it week by week, the
  other person by person. Both take the same load, so either can
  replace the other.
* **Next project** (a block): a second project's tasks, scheduled into
  what the relaunch leaves its people, and beside it on the calendar
  alone.
* **Schedule risk** (a block): how sure the finish is. It takes Tasks,
  and the Scheduler itself, which it calls as a function: it works the
  plan out again *Runs* times (300), each time with the work of every
  task that has a Low and a High drawn afresh between them, most often
  its Work (a part day leaving the rest of the day free). *Drawn tasks*
  is Tasks with the work drawn (select it for its step); on the sheet
  each draw shows its mean. *Drawn plan* is `(call scheduler {:tasks
  drawn-tasks} :plan)`: the Scheduler's plan for the drawn tasks, so
  whatever you change in the Scheduler, the runs follow. *Finishes*
  draws the runs' last days as a histogram; *Likely by*, *80% sure by*
  and *95% sure by* are the dates half, four in five and nineteen in
  twenty of the runs finish by; type the date the launch was promised
  for in *Promised by*, and *Chance of that* says how many runs make it.
  A task without Low and High, or finished, keeps its Work, so with none
  every run is the plan. The facade shows *80% sure by* under the
  forecast. More runs are steadier and slower: every change to the
  tasks works them all out again (300 take about a second).
* The summary: forecast and baseline finish, the slip, work done,
  schedule and cost performance (SPI, CPI), late tasks, the critical
  path, the days someone is booked twice, the weeks overbooked, and the
  assignments to check.

## 47. Protection

A workbook such as the project plan has places to type (tasks,
resources, settings) and places that only work things out (the summary,
the plan). Protection keeps the second kind from being changed by
accident: a value typed over a formula, a row cleared, a paste in the
wrong place. It's a guard against slips, not security: anyone can turn
it off.

* **Protect a grid**: right-click a cell of it → *Freeze and protect ▸
  Protect this grid*
  (*Protect the top grid*). Its cells then refuse typing, clearing,
  pasting, filling and cutting, and its rows and columns can't be
  inserted or deleted, nor the grid resized. Nor do the panels change
  them: a steps formula's steps in the Inspector (+ Step, ✎, ◂ ▸, ×,
  which say *🔒 Protected* above them), a renderer, number format, unit
  or style, a name, a rule on the Rules tab or a column's kind over
  a protected cell, placing from the Shelf, Repeat, Find and Replace,
  the solver's Keep, AutoSum, promoting from the Data tab, inserting a
  file, the cell menu's conversions (spill to values, freeze, sub-grid,
  instances, Make input, headers) and moving a spanned cell. Each says
  why where you pressed, and the way out: open the lock. It says so as you start: a protected cell shows
  🔒 where the formula bar's ƒ is, and the first key typed over it (or
  Enter, F2, a double-click, a click in the formula bar) gets the reason
  and the way to edit it anyway, before anything is typed. Grids inside
  it have protection of their own (or none).
* **Leave inputs open**: in a protected grid, select the cells to type
  into and right-click → *Freeze and protect ▸ Leave B1:B2 open to
  typing*; *Close … to
  typing again* undoes it. An open area is an `{:area "B1:B2" :unlocked
  true}` rule on the Rules tab.
* **Protect part of a grid**: a rule `{:area "1:1" :protect true}` on the
  Rules tab protects that area of a grid that isn't protected as a
  whole: a table's heading row, say, so its column names stay.
* **Protect while something is true**: a protect rule can have a
  condition (*Only while* on the Rules tab, `:when`), and then protects
  its cells only while the condition holds: an order's details once it
  has shipped,
  `{:area "B3:B4" :protect true :when "(not= $>order-card>order-head>B5 \"open\")"}`.
  The condition is written as a conditional format's is (§10): a formula
  that's true or false, or a function of the cell's own value. Its frames
  come and go with it, and typing into a cell it protects says why at the
  first key. Give it what to say (*Refusal says*, `:msg`), such as *a
  shipped order is kept as it was: set its status back to open to change
  it*, and that is what a refusal says; without one, it names the
  condition. A condition that fails (a cell it reads has an error)
  protects nothing. An *open* rule can have a condition too: in a
  protected grid it leaves its cells open only while the condition holds.
* **Records the database keeps**: protection is the sheet's guard against
  slips, and the lock turns it off. A rule that must hold whatever is
  typed, and wherever (another sheet, a button, the Data tab), belongs
  in the database itself, as a trigger (§53). The two go together: the
  database refuses, and a protect rule with a condition says so before
  anything is typed.
* **You can see where typing goes**: every input has a soft green frame
  round it (one round each block of them), so the places to type stand
  out, whether or not anything is protected and whether the lock is open
  or closed. An input is a cell the sheet reads that you can type into:
  a value or a blank box a formula or a button uses (a search box, an
  order's number, a quantity to add), the cells of a table a formula
  reads (not its headings), a cell of a view or a query that writes back
  to its table or the sheet's database (§48, §53), and a cell left open
  in a protected grid. Labels nothing reads and formulas aren't framed,
  nor is what's worked out (an order's total); typing into a cell that
  can't take it says why at the first key.
* **You can see what's protected**: a sheet with anything protected has
  a small lock after its name on its tab, and the bar above its grid
  (where the grid's name is) says *🔒 Protected* and *type in the green
  frames*, or *🔒 Partly protected* and *some cells here refuse typing*
  when only some of what's shown is (a heading row, a form drawn inside
  it). A grid that protects nothing shows the lock alone and *nothing
  here is protected*: everything there takes typing. (A sheet with
  nothing protected anywhere shows no lock at all.)
* **The lock** is that *🔒 Protected* in the bar above the grid. Click it
  to turn protection off while you work on the sheet (*🔓 Editing all*,
  in gold; the tabs' locks open too), and again to turn it back on.
  It's one switch for the whole sheet, so *Editing all* shows above every
  grid while it's open. It isn't saved: a sheet opens protected.
* **What isn't stopped**: the sheet's own buttons and controls (a button
  setting the baseline writes it, a tick box in an open cell ticks);
  column widths, row heights, freezing panes and what a grid shows;
  comments; libraries, scenarios, tests and templates; and protection
  itself (*Protect*, *Leave open*), the way out.
* The Project plan example protects the heading rows of its tables; the
  Order book protects its form sheets, leaving their boxes open. An
  order that has shipped (or is invoiced or cancelled) is kept as it
  was: its customer, its days and its lines no longer take typing, and
  the database refuses changes to them however they come, *Mark as
  shipped today* pressed again included. Its status and note still
  change, and setting its status back to open lets you change the rest.

## 48. Views that write back

A view of a table (a sorted or filtered copy, one row as a form, the
plan made from a list of tasks) shows values that live elsewhere, in the
table's cells. You can type over such a value where you see it, and
what you type is written into the table's cell, as if you'd gone there
and typed it: the view and everything else follow.

* **Which cells**: in a grid a formula makes (§22), a cell showing a
  value read from a grid with headers (§19), unchanged. They're marked
  with a thin line at their left, and hovering one says which cell it
  writes (*Typing here writes tasks>D5*). The formula bar shows that
  cell's own source, and a date in its own format.
* **Where it works**: `(as-grid table)` and anything that keeps the
  table's rows (sorting, filtering, `sort-by`, `filter`, `assoc`,
  `merge`); `filter-view`'s rows; `record-view`'s fields, which make it
  a form; the Project plan's details under its Gantt chart.
* **Where it doesn't**: a value worked out afresh (a total, a sum by
  group, `qty × 10`, a date the plan computes): those cells still say
  they can't be typed into, and why, as soon as you start. A value that
  matches the table's only by name isn't taken for it: it must be the
  same value.
* **New records**: a table shown as it is (a reference to it, `=tasks`
  or `$>tasks`, drawn as a grid) shows the table's empty row after its
  records too, as the last row of its grid. Type a record there and the
  table takes it, keeping an empty row after it as it does when typed
  into directly (§19), and the view follows: Enter goes on to its new
  empty row. A sorted or filtered view has no such row (a record typed
  there could land anywhere, or not show), nor does a view protected
  outright, or one of a table whose last row already holds a record:
  type into the table itself once to give it its empty row. Pasting
  several records past the end isn't done yet: paste into the empty row
  one record at a time, or into the table. Inserting a row inside the
  table is done in the table itself.
* **A column's default** (§19) shown in a view (a new task's work, 0
  until given) is typed over as in the table: the table's cell takes
  the value.
* **Controls too**: a renderer that writes its cell (a date picker,
  `date-input`; a drop-down; a slider) over such a value writes the
  table's cell, so a form of pickers edits the table.
* **The table's rules apply**: its protection (§47), validation (§16)
  and column kinds (§19); typing text, numbers and dates follows the
  table's column (§12). Each edit is one undo step; the status bar says
  which cell was written.
* **In your own functions**: rows read from a table keep where they came
  from through `filter`, `sort-by`, `assoc` and `merge`. A row built
  afresh can say it came from another with `(from-row src new)`, so its
  values under the same names are the source's; `(row-source row :qty)`
  says which cell a row's value came from. A panel names the cells its
  values came from with `:sources {[row col] cell}` (as-grid's option),
  as `record-view` does.
* **Read-only, writable, or in part**: a view writes back unless you
  say otherwise, with protection (§47). Right-click a cell of the view →
  *Freeze and protect ▸ Protect this grid* for a view that only reads;
  then select what may be typed into (a column of it, by its heading
  when zoomed in, or a block) → *Leave … open to typing* for a view that
  writes those alone. Both the view's cell and the table's must be open:
  a protected table isn't written through an open view either.
* Typing, Delete and paste write through (a pasted formula goes as its
  value); fill and cut over such cells don't: they're refused, saying
  the grid is a formula's.

## 49. Rows in dependency order

Some tables are worked out row by row, each row from the results of
other rows it names: a part's cost is its own plus its parts' costs; a
task starts after the tasks it waits for finish; a manager's head count
is their reports' plus one. In cells alone this is a circular reference
(a lookup by ID reads the column it is in), as in Excel. `in-order`
does it in one formula:

```clojure
(in-order parts :part :made-of
  (fn [p below]
    (assoc p :total (reduce + (:cost p) (map :total below)))))
```

* **The rows** are a table (a headed grid, or any list of row maps).
  `:part` is each row's ID; `:made-of` is what it waits for: IDs
  separated by commas or semicolons (`rim, spokes`), one ID, a list of
  them, or blank for none. IDs match as text, so the number `3` and the
  text `3` are one ID.
* **The function** gets the row and `below`, the results of the rows it
  waits for, in the order it lists them (a part listed twice comes
  twice). Its result is the row's result; `in-order` returns them all,
  in the rows' order. The rows can be in any order: each is worked out
  after what it waits for.
* **A plan's dates** are the same shape: start the working day after
  the latest finish among the tasks a task waits for.

  ```clojure
  (in-order tasks :id :after
    (fn [t before]
      (let [s (WORKDAY (reduce max (dec starting) (map :finish before)) 1)]
        (assoc t :start s :finish (WORKDAY s (dec (:days t)))))))
  ```

  `WORKDAY` and `NETWORKDAYS` take a list or a set of holiday dates, or
  a working calendar (§51).
* **Carrying something from row to row.** Give options with a `:state`
  and the function gets that state first and returns the row's result
  and the new state, `(fn [state row before] [result new-state])`: each
  resource's booked days, say, so a task also waits for its resource to
  be free. `:by` decides which of the rows that are ready (whatever they
  wait for is worked out) goes next, the smallest first: `{:state {} :by
  :priority}`, or a list such as `[priority date]`; ties go in the rows'
  order. A row never goes before what it waits for, whatever its `:by`.
* **Refused, with the reason**: two rows with the same ID; a row waiting
  for an ID no row has; rows that wait for each other in a loop, named
  (*These rows wait for each other in a loop, so none can be worked out:
  bike → wheel → spokes → bike. Take one link of the loop out of their
  :made-of.*).

### As a step

The same, without writing a function: in a table's steps (§43), **+ Step
› In dependency order** adds a column worked out row by row, each row
after the rows it waits for.

* **Rows named by** the ID column, **each waiting for the rows in** the
  column that names them (IDs separated by commas, or a list).
* **A new column** and what it is: a formula of the row, its columns by
  name, and `waits-for`, the rows it waits for, already worked out. A
  column of them is `(:finish waits-for)`, so `(MAX (:finish waits-for))`
  is the latest finish and `(SUM (:cost waits-for))` their total. A
  task's finish is `(WORKDAY (MAX start (:finish waits-for)) days)`;
  add its start after, as an ordinary worked-out column.
* **Upwards**, ticked, is for trees where each row names its parent: an
  org chart's Manager, the assembly a part goes into. A row is then
  worked out after the rows naming it, and those are its `waits-for`.
  Headcount by manager is `(+ 1 (SUM (:headcount waits-for)))`; a part's
  total cost `(+ cost (SUM (:total waits-for)))`.
* **Refused, naming the rows**: two rows with the same ID, a row waiting
  for an ID no row has, and a loop (*1 → 3 → 2 → 1. Take one link of the
  loop out of their after*). The step's chip is marked and the reason is
  under the chips, with *Change it ✎*.

### Scheduling: rows placed in time

**+ Step › Schedule** places each row in time: as early as what it
waits for, its earliest date and its resources allow. It suits
projects, and also production orders on machines, or lab and studio
sessions with their rooms and crews.

* **Columns to pick:** the ID; the work in days; who does it (one or
  several, `Dee, Chen`); what it waits for (IDs separated by commas, a
  lag in working days after an ID: `12-2d` starts two days before 12
  finishes, `5+3d` three days after); the earliest date (Not before):
  a column of each row's, or *a date or a cell…* for one date for every
  row, typed as you type a date in a cell (`15 Feb 2027`, `2027-02-15`)
  or a cell holding one (click it), read back beside it (*the date Mon
  15 Feb 2027*, *the cell start, for every row*); Priority (1 first when rows compete for a resource; blank last); and
  fixed dates: a row with a *finished* date is done and keeps its
  dates; one with a *started* date alone is under way, its work placed
  from its earliest date on. A row that waits for nothing needs an
  earliest date. Also **the part done** (0 to 1: the step places only
  the work left, and gives a finished row 1) and **due** (the date each
  row should finish by). A new Schedule step picks them for you from the
  columns coming in, as a schedule's rows have them: by their names
  (*ID*, *Work*, *Resource*, *After*, *Not before*, *Priority*,
  *Started*, *Finished*, *Share*, *Done*, *Due*, or names like them:
  *Days* for Work, *Who* for Resource, *Waits for* for After, *Deadline*
  for Due), then by what they hold (a column of IDs is the ID). The
  form says so: check them before Apply.
* **Starting on** and **status date**, each a date or a cell (click
  it), read back beside the field. No row still to do starts before the
  start date. The status date is the day your actuals (started,
  finished, the part done) are as of: work not done is placed after it,
  a task under way continuing from the first day after it that its
  people work, a task not started no earlier than the latest of the
  start, its own earliest date and that day. Left blank (or a blank
  cell), there are no actuals yet.
* **Either of two resources:** write `M1 | M2` in the Resource column
  for one or the other (`Ana, M1 | M2` is Ana and one of the machines).
  The step tries each and takes the one that finishes the row first (of
  two equal machines, the one free first; when they finish alike, the
  one with least booked so far, so with levelling off a role's work is
  shared rather than all given to its first person). A Resource column holding
  *one or more of* your resources takes `M1 | M2` when both are yours.
* **Roles: pools of resources.** Give the step a table of pools, a row
  per pool and member: Role and Resource (Ana, Designer; Ana,
  Developer; Chen, Designer). A row naming a pool (*Designer*) is then
  any one of its members, picked as `M1 | M2` is, and one person can be
  in several pools. Who does each row can come from two columns: *by
  Resource and Role* takes the names in both, so *Accessibility
  review* with Resource Device lab and Role Designer is the lab and a
  designer. A row asking for two of a pool (`Designer, Designer`), or
  for Ana and a Designer, gets two different people. The plan keeps
  what was asked (Role) beside who was given it (**assigned**). The
  pools table needs a Resource column (the members) and one other
  column (the pools); with more, say which is which (`:pool :role`,
  `:member :resource` in the step's formula). Refused, naming it: a
  pool whose rows have no members, a pool with the name of a member
  (an item naming it couldn't say which it means). A name in Role that
  isn't a pool is caught where it's typed, as Resource's are: give Role
  *one or more of* the pools' Role column (*Holds ▸*). The Shelf's
  *This table is ▸* offers *pools* for such a table.
* **Who is on it, once begun:** a column of the people actually on a
  row. When it isn't blank, they fill the row's choices in place of a
  pick, just as a started date is taken in place of a planned one: Cy
  on a Designer task is its designer. What the row names stays (the
  Device lab logs no time, and is still booked). Someone on it who
  isn't what the row asks for (someone else did the work) stands in
  for a choice nobody fills, or is added, and the row's **who-check**
  says so (*Ben is on it but isn't a Designer*). A choice nobody on it
  fills is still picked, and said (*nobody said to be on it is a
  Developer: Ben is a guess*), as is a task begun with nobody said to
  be on it. A timesheet gives the column: group it by task with a list
  of Resource, and join that to the tasks before the Schedule step.
  Someone who handed over part way through counts as on it for the
  whole task.
* **Follows on from:** a column naming another row, whose people this
  row takes: *Visual design* follows on from *Wireframes*, so it goes
  to whichever designer did the wireframes. Each pool the row asks for
  is filled by whoever the other row got for that pool; a row asking
  for no pool of its own takes everyone picked for the other row. When
  someone was picked for the other row, the resources it names aren't
  passed on (the lab that Accessibility review names): name them here
  too if they're needed. When nobody was (the other row names everyone),
  the row takes the resources it names. It says who, not when: a row that should also come later
  waits for the other in After as well. The other row is placed first.
  A row begun before the one it follows on from takes its own people
  (who is on it), and who-check says to give them. The pick is made
  for the first row alone: it takes whoever finishes that row first,
  not whoever would finish both first. Refused, naming the row: a row
  that names none, or itself, or several; rows following on from each
  other in a loop; asking for a pool the other row had nobody picked
  in; following on from a row with nobody on it (a milestone).
* **Place rows** as soon as possible (the usual), or **as late as
  possible, finishing by** a date or a cell: each task still to do
  finishes as late as the tasks waiting for it, its due date and that
  date allow, and is levelled backwards. Tasks done or under way stay
  where they are. When a task can't fit, the step says which, the day
  it would have to start, how many working days too early that is, and
  when the plan finishes as soon as possible; with a due column picked,
  the finishing date can be left blank.
* **The calendar** (§51) gives each resource's working days and how
  much of each; a row with several resources works the days all of them
  do. The field is a choice: the sheet's tables that have a calendar's
  columns (a Capacity or a Takes column, with Resource, From, To and
  Weekdays as they need), by name as the formula writes them
  (`settings>calendar`); *a new calendar…*; or *another*, to type a
  name, a reference or a formula. Under it, what the one chosen reads
  as: *6 resources, 13 rules: 259 working days in 2027* (everyone's, in
  the year of its first dated rule), and *month view ▸*, its months one
  at a time, a chip per resource per day: green working, amber part of
  the day, struck through off. A table that reads the step's own cell
  isn't offered (it would make the step read itself). A typed one is
  read when you press Enter or leave the field; `(calendar-for
  settings>calendar "Ana")` and a list of dates are taken too. A table
  that would be a calendar with a column renamed is listed too, greyed,
  with what it lacks (*settings>layer — has Start: rename it From*). A
  table without a calendar's columns is refused there, under the field,
  naming the columns a calendar needs and those it has; Apply refuses
  it too. A column of the right name holding something else (Capacity
  as text) is read all the same, and the reading says so (*⚠ Capacity
  holds text*): its rows are refused one by one if they can't be read. Blank (*none*): Monday to
  Friday, every day a full one.
* **A new calendar…** puts a working calendar in layers beside the
  steps cell, where the Shelf would put it (below everything when
  there's no room beside it): a grid named *work-calendar* holding a
  resources table (Resource, Units), the working week, holidays, days
  off and overtime, each a table with its columns' kinds and an example
  row, and *calendar*, the steps combining them, each step with a note.
  The field then holds `work-calendar>calendar`, and the form stays open.
  In a protected grid it's switched off until the lock is opened.
  The Shelf's *Start a table* offers the same: *working calendar, in
  layers*.
* A name in the calendar or the rows that isn't one of your resources
  is caught where it's typed: the calendar's layers and Tasks' Resource
  hold *one or more of* the resources (*Holds ▸*, §19), which refuses
  any other name (§19, *IDs, lists and lookups*).
* **Level resources:** on, rows wait for their resources; off, only
  for their links; or as a toggle cell says.
* **A resource does** one row at a time, or several, each row taking
  its share of the day from a column (0.5: half): two half-day rows fit
  on one day, and a row at half a day takes twice its work in days.
* **Part days.** Work needn't be whole days. A row of 2.5 days holds
  its third day for half of it, and the rest of that day is free: its
  resource's next row starts there, and a row waiting for it (finish to
  start, no lag) starts that afternoon, with the half day that's left.
  So 5.3 + 5.3 + 5.3 days in a row take 15.9 working days, not 18. Its
  bookings and its cost count the part it takes; a milestone after it
  passes the half day on. Dates are still days: the row's finish is the
  day its work ends, part-way through or not.
* **A row** waits until all its days are free, or **splits** its work
  round days already taken.
* **Shared work:** each resource does all of it (a day's work for Ana
  and Ben is a day for each), or it's divided between them.
* **Link types**, written after the ID as project tools write them:
  `3` (or `3FS`) finish to start, the row starting after 3 finishes;
  `3SS` start to start, starting no earlier than 3 starts; `3FF` finish
  to finish, finishing no earlier than 3 finishes; each with a lag in
  working days, `3SS+1d` a day after 3 starts, `3FF-2d` two days before
  3 finishes. An ID that itself looks like one (`ASS`, `3FF`) is that ID.
  Start to finish (`3SF`) isn't taken, and a type that isn't one
  (`3XX`) is refused, both naming the forms a link takes. Slack and
  critical follow each type: a task that finishes with another is
  critical when that one is. No task starts before the plan's first
  day, the earliest of its earliest and fixed dates: a long task
  finishing with a short one starts there at the earliest.

It adds **start**, **finish**, **days** (working days from start to
finish), **wait** (days it waited for its resources), **slack** (how far
it could slip without moving the last finish; as late as possible, how
much earlier it could start), **critical** (no slack) and, when it's
given resources, **booked** and **assigned** (who does it, each
`M1 | M2` the one picked); given started or finished dates, **state**
(*done*, *in progress* or *not started*); given due dates, **late**
(working days past due: 0 on time, blank for a task with no due date;
a task due on a Saturday and finished on the Monday is a day late);
given pools, who is on it or follows on from, **who-check** (blank, or
what to check about who the row went to). Their other columns pass
through.

**booked** is how much of each resource's time the row takes, a number
per resource in the order the row names them (`10, 5` beside `Dee,
Chen`): on each day it holds them, its share of the day times what the
resource has that day. Chen at half time on a task of 21 working days
is booked 10.5; with *divided*, each does their part; a meeting the
calendar takes out of a Monday isn't the task's. A row holds its
resources on the days from its start to its finish, except those it
waits for them; a task's last day counts whole, even when its work ends
part way through, since nothing else is placed in the rest. A row with
no resources has none (blank). Times each resource's rate, it is the
row's cost: `(SUMPRODUCT booked (XLOOKUP assigned resources>resource
resources>rate 0))` (*assigned*, so a row written `M1 | M2` is priced
at the machine it got).

**`(bookings table)`** gives what the step booked as rows a calendar
takes (§51, *Bookings*): for each row and each of its resources,
Resource, From, To and Takes, a row for each run of days with one
Takes, after the columns you name (`(bookings plan :task)` puts each
task's name first). Appended to the calendar the step was given, they
leave each person exactly what the schedule left them. They stay with
their rows through later steps (a filter keeps those of the rows it
keeps); a table not made from a Schedule step's rows is refused. A table with a
column already named one of those is refused: rename yours first, with
a Choose columns step. Refused, naming them: two rows with one ID, a row
waiting for an ID no row has (with a hint when a lag lacks its `d`),
rows waiting for each other in a loop, a row with nothing to start
from (as late as possible: nothing to place it back from), a part done
that isn't a number from 0 to 1 (50% is 0.5), a due date that isn't a
date, and a start, status or finishing date that isn't one.

### Several projects in one plan: keys of several columns

Two projects numbering their tasks 1, 2, 3… can share one plan, and
their people, without renumbering. In the Schedule step's form, **its
ID [id] within [project]**: the ID column and the column telling the
projects apart (the step then holds `:id ["project" "id"]`; typed,
`:id [:project :id]` is the same). Then:

* **After** names the row's own project's tasks: in Intranet, `3` is
  Intranet's 3, `3SS+1d` its start. Another project's task is written
  with the project first and a colon: `Web:3`, `Web:3FF-1d`. (A key of
  three columns, typed in the step as `[:client :project :id]`, writes
  `Acme:Web:3`.)
* **Levelling** spans the projects: one resource does one task at a
  time, whichever project it's in.
* **Slack** and **critical** are each project's own: how far a task
  could slip without moving its project's last finish. The project that
  finishes first has its own critical path.
* **Refused, naming them:** two rows with the same project and ID
  (*rows 1 and 18 both have project Intranet and id 1*), a task naming a
  task its project hasn't got (*project Intranet has no id 4: Web has a
  4; write Web:4 for it*), and `Web:9` when Web has no 9. With the
  whole table's IDs (no *within*), IDs that repeat are refused with a
  hint naming the column that tells them apart.

**The plan carries its key.** Later steps, views and copies use it
without being told again:

* A **Gantt** of the plan draws the links within each project and
  across (`Web:3`), and a click picks a project's task: `:pick
  "Intranet:3"`. `{:lane :project}` gives each project a line.
  `{:key :id}` written in the options means the same key.
* `(record-view (picked tasks :id chart))` shows the task picked from
  Tasks by project and ID, not every project's 3. A Tasks table without
  the project column (or with it under another heading) is refused,
  saying so; give the pairs, Tasks' column first: `(picked tasks
  [[:programme :project] [:id :id]] chart)`, or `[[:id :id]]` for the
  ID alone. Several columns are
  written as a join's: `[[:project :project] [:id :id]]`.
* A **join** matches on the pairs you give it: tracking joined to a
  baseline matches on both, `:on [[:project :project] [:id :id]]`. Joined
  on the ID alone, each task matches every project's task of that ID,
  and the join's warning says what to match on (*the rows are told
  apart by project and id: match on both, [[:project :project] [:id
  :id]]*).
* A **snapshot** (§52) keeping the ID must keep the project too, or
  the copy couldn't be matched back: the button says so before it's
  pressed.

**IDs within a column.** *Holds ▸ IDs within Project* makes a column of
IDs each once per project: Web's 1 and Intranet's 1 are two, a second
Web 1 is flagged (*this value is used by another row of the same project
too*). Retyping Project's heading keeps it.

**A limit:** a column holding *one or more of* the IDs (§19) checks
bare IDs, so `Web:3` typed into it is refused there. For a plan of
several projects, make After a column of *lists*: the Schedule step
checks every link, refusing one that names no task.

The tour's *Rows in dependency order* rolls up a bike's parts and works
out a small plan, and schedules it with a Schedule step, as the Project
plan example does; its *Two projects* rows schedule two projects
sharing two people, draw them a lane each, and track them against a
baseline.

## 51. Working calendars

Which days someone works, and how much of each, is a **calendar**: a
table with the columns *Resource*, *From*, *To*, *Weekdays* and
*Capacity*, one rule a row. For a resource and a day, **the last row
that matches wins**, as in a style sheet:

| Resource | From | To | Weekdays | Capacity | |
|---|---|---|---|---|---|
| Chen | | | | 0.5 | Chen works half days |
| | | | Sat, Sun | 0 | the weekend |
| | 1 Jan 2027 | | | 0 | a holiday |
| | 27 Dec 2027 | 31 Dec 2027 | | 0 | the company closes |
| Press 2 | 1 Mar 2027 | 3 Mar 2027 | | 0 | a machine's maintenance |
| Ben | 13 Feb 2027 | | | 1 | Ben works a Saturday |

* **Resource**: blank is everyone; one name, or several separated by
  commas (`Ana, Ben`). Names match whatever their case.
* **From** and **To**: the dates the row covers, both included. Both
  blank: always. To blank: the From day alone.
* **Weekdays** limits the row to those days of the week: `Sat, Sun`,
  `Mon-Wed`, `Fri–Sun` (round the weekend), the names in full or cut to
  three letters or more, separated by commas.
* **Capacity** is the share of the day worked: 0 off, 0.5 half, 1 full,
  more than 1 for longer days. For a pool of things (three projectors,
  twelve seats, 100 widgets in stock) it's how many there are: 3, 12,
  100.
* A day no row matches is a full working day. So a calendar holds its
  own weekends: give it a weekend row (blank Resource, `Sat, Sun`, 0).

Order matters: put a resource's half time (Chen, 0.5, no dates) before
the weekend, or Chen would work half of every Saturday.

### Layers, one table each

Keep each kind of rule in a table where it belongs (the working week
and the holidays on a settings sheet, leave on its own, maintenance
beside the machines) and join them with a `steps` formula of appends
(§43), in the order they override each other:

```clojure
(steps week
  [{:op :append :table holidays :match [:from :holiday] :new [:from :holiday]}
   {:op :clean :column :capacity :how :fill :to 0}
   {:op :append :table leave :match [:resource :from :to] :new [:to]}
   {:op :clean :column :capacity :how :fill :to 0}
   {:op :append :table overtime :match [:resource :from :weekdays :capacity]}])
```

The `:match` and `:new` are what each append's form records when you
Apply it (§43): the columns it matched. A table needn't have every column: a holidays table of *From* and
*Holiday* is fine, and the *fill blanks* step after its append gives
its rows Capacity 0 (it touches only blanks, so only the rows just
appended). Other columns, a holiday's name or a resource's rate, are
carried along and ignored. Select the formula and the Inspector shows
the layers as chips: move one, with its fill, and what overrides what
changes. Each append made with its form records the columns it matched
(§43), so a layer's heading cleared and typed wrong (*Resorce*, *Form*)
stops the calendar with an error naming the layer and the heading,
rather than quietly making someone's leave everyone's.

**Giving a layer its kinds at once.** Select a layer, and on the Shelf
tab *This table holds* → *a calendar*: each of its columns a calendar has,
by its heading, gets what that column holds (Resource one or more of
the sheet's resources, From and To dates, To on or after From,
Weekdays days of the week, Capacity numbers at least 0) and dates a
date format. A column named nearly right is renamed (*a calendar —
renaming Start to From*); a layer without Capacity is a part of one
(the steps give it Capacity), and the button and the message say so.
An empty grid has a button per way a calendar starts: of Capacity
(working time), of Takes (bookings), of Uses (a pool's bookings). A
table none of whose columns is a calendar's has no button: *… shapes
it isn't: why* says what a calendar needs and the columns it has.

### Using a calendar

Give it to Excel's date functions where they take holidays:

* `(NETWORKDAYS from to cal)` counts the working days (any day with a
  capacity above 0, a half day included); `(WORKDAY start n cal)` is the
  date n working days on. A list of holiday dates still works as in
  Excel, with Saturday and Sunday off; a calendar has no weekend but its
  own rows. With the .INTL forms, a weekend given is laid beneath the
  calendar's rows.
* `(calendar-for cal "Ana")` is Ana's calendar: her rows and
  everyone's. Give it to WORKDAY or NETWORKDAYS, or anything else that
  takes a calendar. With several, `(calendar-for cal "Dee, Chen")`, a
  day counts when all of them work, at the smallest of their
  capacities. A calendar not narrowed is everyone's rows alone.

RecDep's own date functions are all called the same way: **the dates
first, then the calendar, then who**, as WORKDAY takes its dates and
then its holidays. The calendar and who can be left out: with no
calendar, Saturday and Sunday are off and every other day is a full
one; with no who, it's everyone's rows. Who works as `calendar-for`'s
does (`"Ana"`, `"Dee, Chen"`, a list), so `(working-days from to cal
"Ana")` and `(working-days from to (calendar-for cal "Ana"))` are the
same.

* `(working-days from to cal "Ana")` sums the capacities: Chen at half
  time has 2.5 in a week of five.
* `(capacity d cal "Ana")` is one day's: 0, 0.5, 1.
* `(work-end start work cal "Ana")` is the day `work` days of work begun
  on `start` are done, each day doing its capacity: two days at half
  time take four working days.
* `(week-start d cal "Ana")` starts weeks the day after the calendar's
  weekly days off (Ana's, with her name).
* `(share-elapsed from to date cal "Ana")` is the part of the period from
  `from` to `to` gone by on `date`, in the calendar's working days: 0
  before it starts, 1 from its last day on, `3/5` on the Wednesday of a
  Monday-to-Friday period. A task's planned value by a status date, a
  subscription's revenue recognised so far, a budget phased.
* `(workdays-after from to cal "Ana")` is how many working days `to` is
  after `from`: 0 the same day, 1 the next working day, -1 the working
  day before, a day off between them counted for neither. A finish's
  slip against its baseline, in working days: `(workdays-after
  baseline-finish finish cal)`.
* `(month-view d cal "Ana")` draws the month (below).

**A blank date gives a blank**, in every one of them: a task with no
baseline yet has no slip and no planned value, rather than a number
that looks like one. Where a blank should count as 0, say so with `N`:
`(N (share-elapsed from to date cal))` is 0 for a task with no
baseline. A date that doesn't read (`2027-13-01`, *soon*) is refused,
naming it, and so is a call in another order: a calendar where a date
goes, or a name where the calendar goes, says the order to write.

### A month at a glance

`(month-view d cal who)` draws the month holding date `d` (any day of
it: a date cell, `"2027-03-01"`; it can't be blank) as a monthly
calendar, a week a row, in a grid of its own:

* each day shows its number and a chip per resource: **green** working
  the full day (×1.5 when more), **amber** part of it (½), **struck
  through** off;
* weeks start as the calendar's do: Monday, or Sunday for a
  Sunday-to-Thursday week;
* `who` is a name, several (`"Kim, Sam"`) or a list, each shown on its
  own; a calendar from `calendar-for` shows its resources;
* leave out `who` and each day is shaded by everyone's capacity; leave
  out the calendar too, `(month-view d)`, and it's a plain month with
  Saturday and Sunday off;
* a name the calendar has no rows of its own for is named beside the
  month (*no rows of their own: Lee*): a typo, or someone with no leave
  or pattern, who works everyone's days;
* a day **overbooked** by bookings (below) is **red**, with how far
  below 0 it is.

Its value is the month's days as a table, a row per day: the date, the
weekday and each resource's capacity under their name, so
`(month-view m cal "Kim, Sam")` named `january` gives `january>sam`,
Sam's day by day, and `(SUM january>kim)`, Kim's working time. Point
`d` at a date cell with a `date-input` and picking a day shows its
month.

### Bookings: what's taken, and what's left

A row can book time instead of giving it: put a **Takes** in place of
its Capacity. Takes is the share of the resource's day the booking
takes, 1 all of it, 0.5 half:

| Resource | From | To | Takes | |
|---|---|---|---|---|
| Sam | 13 Jan 2027 | 15 Jan 2027 | 1 | a task of Sam's |
| Kim, Sam | 18 Jan 2027 | 19 Jan 2027 | 1 | a task for both |
| Lee | | | 0.2 | a fifth of every day on support |

* **A share of what the day has**, not a number of days: Kim at half
  time, booked with Takes 1, loses her half day.
* **On the days all its resources work.** The task for Kim and Sam
  takes from each of them only on the days both work: a day Kim is off,
  Sam's day isn't taken. Bookings with no Resource take from everyone.
* **Bookings add up.** A day's capacity is what the Capacity rows give
  it, times 1 less the day's Takes added together. Two bookings of a
  full day leave −1: **overbooked**, by a day.
* **Bookings come after every Capacity row**, wherever they are in the
  table: an overtime row added below them doesn't give back what they
  took.
* `capacity` gives the day as it is, below 0 when overbooked, and
  `working-days` counts such a day as negative: what's left of an
  overbooked week is below 0. WORKDAY, NETWORKDAYS and `work-end`
  don't count a day at or below 0 as working.

So **what's left** of a calendar is a calendar too: the calendar, then
the bookings appended.

```clojure
(steps calendar [{:op :append :table booked :match [:resource :from :to :takes] :new [:takes]}])
```

Give it to anything that takes a calendar. A project plan's bookings
are `(bookings plan)`, what its Schedule step booked (§49): each
task's share of each of its resources' days, from one day to another,
less on a day the calendar's own Takes have taken part of (a weekly
meeting), so a plan over a calendar with bookings of its own doesn't
overbook it. The project's Loading is then the calendar's working days
less what's left's, a week at a time (a series step counts both, §43),
its double-booked days the days what's left is below 0, and a second
project scheduled on what's left goes round the days the first takes.

### Amounts: projectors, seats, stock

Takes is a share: on a pool of three projectors (Capacity 3), *Takes 1*
takes all three. To book how many, put a **Uses** in place of Takes:

| Resource | From | To | Capacity | Uses | |
|---|---|---|---|---|---|
| Projectors | | | 3 | | three of them |
| | | | | | (the weekend: Sat, Sun, 0) |
| Projectors | 3 Mar 2027 | 5 Mar 2027 | | 1 | workshop A |
| Projectors | 4 Mar 2027 | 5 Mar 2027 | | 1 | workshop B |

* `(capacity d cal "Projectors")` is how many are left that day: 3, 2
  on 3 March, 1 on the 4th, below 0 when more are used than there are
  (overbooked). `working-days` adds them up over the dates
  (projector-days); WORKDAY and NETWORKDAYS count the days with any
  left.
* Uses works for stock reserved too: Widgets with Capacity 100, a
  reservation of 30 from 1 to 10 March leaves 70 each of those days.
* A row has one of Capacity, Takes and Uses. Takes and Uses can both
  apply to one resource, from different rows: the day's share is taken
  first, then the amounts used. A meeting taking 0.2 of everyone's
  Monday takes 0.2 of the projectors too, unless it names who it's for.
* **Units.** Give the Capacity and Uses cells a unit of your own
  (`custom-units` with `"projector" :kind`, §25) and they show "3
  projector"; the calendar reads the number. Cells of one kind convert
  as usual (a stock of 5 t and a reservation of 200 kg agree). Don't
  give a share a time unit: 8 h is a third of a whole day, not a
  working day.
* A **Schedule step** shares a resource's day: a task on the pool with
  share 1 holds all three; one projector is share 1/3 (of what's left
  that day). A task needing a person and one projector can't be said
  yet: the share is the person's too.

### When a later row raises someone (a check)

The last matching row wins, so a row for everyone after a resource's
own row replaces it. Usually that's the point (a holiday takes Chen's
half day off), but a row like (Weekdays `Mon`, Capacity 0.8), meant as
"everyone loses a fifth of Monday", *raises* Chen from her 0.5 to 0.8.
Such a row gets a check's corner on the cell where its Capacity was
typed, saying whom it raises and from what:

* for less of everyone's day, make it a booking: **Takes 0.2** in place
  of Capacity 0.8 takes a fifth of each one's own day (put it in a
  table of bookings with Weekdays and Takes columns, appended to the
  calendar last: a *fill blanks* step on Capacity after it would give
  the row a Capacity too, and a row can't have both);
* if they should have 0.8 there, say so in a row of their own after it
  (that row also keeps the check from showing again);
* or Ignore it in the Inspector's Check row, as any check.

A Takes of 1 or more on a pool (a resource whose Capacity is more than
1) gets one too: it takes all of the pool; Uses 1 takes one. The
Schedule step's calendar field says the first of these after what the
calendar reads as. They're checks, not errors: the calendar is read as
written. The status bar's count of checks doesn't include them.

### Mistakes are refused, naming the row

Whatever reads the calendar (the plan, a WORKDAY) stops with a message
naming the row and saying what to change: From after To, a To without a
From, a date that isn't one, Weekdays that don't read (`"Sum" in "Sat,
Sum" isn't a day of the week`), a blank or negative Capacity, Takes or
Uses, a row with two of them (a *fill blanks* step on Capacity after the
bookings' append fills theirs: move it before), a table with none of
them, or a column named nearly right (*has Weekday: rename it Weekdays*,
*has Booked: rename it Takes*, *has Use: rename it Uses*). A row with no Resource, dates or
Weekdays and Capacity 0 would take every day off for everyone, which is
what a layer with its dates in another column makes, so it's refused
too. For a mistake to show where it's typed, give a layer's columns
validation rules (§16): `weekdays?` checks a Weekdays column. For the
rest, put a cell that reads the calendar beside its layers, such as
`(NETWORKDAYS start (EDATE start 12) calendar)`: a row that can't be
read shows there, naming it.

The tour's *Working calendars* builds one from four small tables,
shows a month of it, and books time out of it; the Project plan example
schedules people and a machine on one, and a second project into what
the first leaves (its *Next project* block).

### Any table of rules: prices, discounts, rates

A calendar is one kind of **rule table**: rows of conditions and a
value, where the last row whose conditions all hold gives the value. A
room tariff, discounts by customer tier and region, tax rates that
change on a date and a team's shifts are all rule tables.
`rule-value` reads any of them:

| Room | Weekdays | From | To | Price |
|---|---|---|---|---|
| | | | | 100 |
| | Fri, Sat | | | 140 |
| Suite | | | | 250 |
| Suite | Fri, Sat | | | 320 |
| | | 24 Dec 2027 | 26 Dec 2027 | 180 |

```clojure
(rule-value {:room "Suite" :date "2027-03-05"} tariff :price)   ; 320, a Friday
(rule-value {:room "Double" :date "2027-12-25"} tariff :price)  ; 180
(rule-value {:tier "Gold" :region "North"} discounts :discount 0)
```

* **What to look up comes first**, as a map from a column to a value,
  then the rules: the order of XLOOKUP, whose lookup value comes
  first, and of the calendar functions, whose date comes before the
  calendar. Then the column the answer comes from (`:price`, or the
  heading `"Price"`), and last what to give when no rule holds. A call
  with the rules first is refused, saying the order.
* **A blank cell holds for anything.** A cell can list several values
  (`Gold, Platinum`). Text matches whatever its case, and numbers
  match the same number typed as text.
* **`:date`** is matched against From and To (both days included, a
  blank To is the From day alone) or, with neither, a Date column. Its
  day of the week is matched against Weekdays, or against a column
  holding days of the week under another name. The date can be a date
  cell or date text.
* **Asking for a blank value** (a customer with no tier) matches only
  the rules for any tier. Ask for one value at a time: text listing
  several (`"Gold, Silver"`) is refused, since it would match no rule
  as it stands.
* A column listed in `:ignore` is not matched at all, even when it is
  From, To or Weekdays.
* **No rule holds:** you get the default when one is given. Without
  one, you get **#N/A**, saying what was asked (`IFNA` catches it).
* **The value can be anything:** a number, text, a date.

**Every column of the table is a condition**, apart from the value's.
A rule with something in a column the call doesn't name can't be
decided, so it is refused. The refusal names the row and the column:
*its Machine isn't matched by the call*. This stops a rule meant for one
machine from quietly applying to everyone. Either give the call a value
for that column (`{:machine "M1" …}`), or, when the column is only a
note, say so as a last argument: `{:ignore [:note]}`. A column the
call names that the table hasn't got is refused too, with the columns
it has. A rule with no value is refused as well.

**A table of questions at once.** `(rule-values rows rules :price)`
gives each row of another table its value, as a list, so
`(SUM (rule-values nights tariff :price))` prices a stay night by
night. Each condition is read from the row's column of the same name,
and the date from its Date. When the rows call a column something else,
say which: `{:on {:date :night}}`. A condition the rows have no column
for is refused, naming it.

**Rates that change on a date.** Put the base rate first, with no
dates, and each change after it with its From and To. For a change
with no end, use a far To such as 31 Dec 9999.

**A working calendar is the same reading.** Who is in Resource, the
value is in Capacity, and a day no rule matches is a full one:
`(capacity d cal "Ana")` gives what `(rule-value {:resource "Ana"
:date d} cal :capacity 1)` gives for a calendar of Capacity rows. The
calendar functions add what a lookup can't do: bookings that add up,
several people at once, and counting over a span of days. A calendar
still carries its other columns (a holiday's name) without asking for
them to be named.

The tour's *Working calendars* ends with a tariff.

## 52. Snapshots

A snapshot is a copy of a table kept as it was when taken: a project's
baseline, a month-end close, the versions of a forecast. A
`snapshot-button` takes one when pressed.

```clojure
(snapshot-button ..>plan {:into 'A3 :stamp 'B1 :keep [:id :task :start :finish] :mode :replace})
```

Give it to a cell as its renderer; the cell's value is the button's
label (*Set the baseline*).

* `:into` — the cell the copy goes to, quoted (`'A3`) or as text.
* `:stamp` — a cell given today's date, if you want one.
* `:keep` — the columns kept, in that order; all of them without it.
* `:mode` — `:replace` (the default): the copy replaces the last one.
  `:append`: the copy is added after the ones kept, each row stamped
  `:taken` with the date, so several versions are kept. Pressing again on
  the same day replaces that day's copy (one version a day), and the
  button says so.
* `:name` — the copy's name, a cell of text typed by whoever presses
  (`{:into 'kept :name B1}`; by its address, so *Clear all* on it, which
  takes a cell's name off too, leaves the button saying it needs one). Named
  copies are kept side by side: each row is stamped `:version` with the
  name (first) and `:taken` with the date (last), several a day if you
  like. Pressing with a name kept already replaces that copy (the new
  one goes last, as the latest), and the button says so, with the day
  the old one was taken: Undo brings it back, another name keeps both.
  A project's baselines are kept so: the Project plan's Baselines block
  (§46) keeps *Original*, *Approved* and whatever you name next, gives
  the last as the baseline, and compares any two.

One press is one undo step, the copy and the date together (presses
within a second of each other make one step, as for every button). In a
protected grid the button still writes: protection keeps the copy from
being typed over. A grid too small for the copy grows to fit it. With
`:mode :append` or `:name` the copy holds every version, so whatever
reads it by ID finds each ID once per version: filter by `:taken` or
`:version` to pick one (the latest: `(:version (last kept))`).

**Without typing the map**: click the button's cell, then the Renderer
field (◇): its refine panel has *When pressed* (replace the copy, add
to the copies kept, or, named, keep named copies), *Copy into*, *Named
by* (a cell of text: B1),
*Date in* and *Columns kept* (typed as a list: `id, task, start`).

**When it can't**, it says so on the button (⚠ after its label, the
reason on hover), and beside it when pressed, writing nothing:

* what it copies isn't a table, or has no rows (a copy of nothing would
  wipe the last one);
* a column in `:keep` the table doesn't have (the message lists the ones
  it has);
* the ID kept without its project, when the table's rows are told
  apart by both (a plan of several projects, §49): keep the project too;
* no `:into`;
* `:name` given, but its cell is blank: type a name, and press again;
* `:name` with `:mode :replace`: named copies are kept side by side,
  so leave `:mode` out;
* `:name`, and the table has a column *version* of its own, kept: leave
  it out of `:keep`;
* the `:into` cell holds something that isn't a copy (text, a number):
  it isn't overwritten. Clear it, or give `:into` an empty cell.

The copy is a value written into its cell, saved with the sheet. A
column of dates in it shows as numbers until you give it a date format.

## 53. Sheets with a database

A sheet can carry a SQLite database with it. Saved, it is one
`.recdep.db` file: the sheet and the database's tables together. The
tables show in the **Data** tab, first, as *This sheet's database*,
where you look through them, shape them with steps and promote what
you want into the sheet, as with any database file (§34).

**Giving a sheet a database.** Save, and in the save dialog pick the
file type *RecDep sheet with a database (.recdep.db)*. Where there's
no save dialog (Firefox, Safari), type a name ending `.recdep.db`. The
sheet keeps its database from then on: Save writes a `.recdep.db`
again.

**Adding tables.** Two ways, each a copy: what it copies isn't
changed, and a sheet without a database is given one first (it asks).
* **A table of the sheet**: right-click one of its headings → *Copy
  this table into the sheet's database…*. The headings are the column
  names; each column's kind (*Column holds ▸*) is its type, an IDs
  column the table's key. Dates go in as `2027-01-04`, TRUE/FALSE as 1
  and 0, blanks as nothing.
* **A table or result in the Data tab**, from any file you've opened
  there: *Copy into the database…*, beside Promote. Every row comes,
  as the step shown leaves them, without going through the sheet's
  cells: the way to bring in more rows than a sheet wants to hold.

**What the table says.** What the grid's columns say goes with them:

| In the grid | In the table |
|---|---|
| the column's kind (*Column holds ▸*) | its type: text TEXT, number REAL, whole number INTEGER, date DATE, TRUE/FALSE BOOLEAN |
| an IDs column | the table's key; without one, a column `id` is added, last, numbered by the database (said), so the table's rows can be written |
| another IDs column, or *each value once* | UNIQUE |
| what a blank reads as (a kind's default) | DEFAULT; a blank copied is that value |
| a date *on or after* another column | a CHECK (`until >= day`) |
| values listed, bounds | CHECKs |
| a lookup of another table's keys (`{:kind :one-of :of-table "customers"}`, or of a query's key column) | REFERENCES that table, indexed |

A table that is a shape of yours (§19, *What a whole table holds*: its
columns hold the shape's kinds) also takes the shape's rules: a needed
column is NOT NULL, and *one of customers' id*, where no grid of the
sheet is customers, is a lookup of the sheet's database's customers
table (a box of its keys by name, as above), so its REFERENCES too. A
row breaking a rule is refused, naming the column.

So a table is made from scratch in three steps: a headed grid, its
headings the column names; each column's kind (*Column holds ▸*, with
a default if blanks mean one, or a shape adopted, *Start a table ▸*);
then right-click a heading → *Copy* (or *Move*) *into the sheet's
database…*. Copy a table referred to first (customers before orders),
so the lookup of its keys has a table to name.

Each asks the new table's name; a name already there is refused.

**Changing tables.** Two changes are made in RecDep, each one undo
step (Undo takes it out of the database, Redo puts it back):
* **A column added**: on a query showing every column of one table
  (`select * from products`, or steps that keep its columns), Tab on its
  last heading asks for the new column's name and what it holds (*barcode
  text*, *due date*, *weight number*, *pallets whole number*, *fragile
  TRUE/FALSE*); it's added to the table, and the query shows it once
  asked again. A name already there is refused.
* **A view saved**: in the Data tab, on the sheet's own database, *Save
  as a view…* (in place of *Copy into the database…*, beside Promote)
  saves what the step shown gives as a view: its query, which follows
  the tables as they change, read as any table (`(query "leeds_customers"
  [])`).
Views written in SQL are made the same way: a SQL step in the Data
tab (`select c.id, c.name, (select count(*) from orders o where
o.customer_id = c.id) as orders from prev c`), then *Save as a view…*;
its columns that come from a table are typed into as that table's.
**Its schema, in the Data tab.** Unfold a table of *This sheet's
database* (▸ beside its name): under its columns (each marked *indexed*
when an index starts with it, and *on delete cascade* where a reference
says so) are
* **its definition**, the table's CREATE TABLE. *Edit…* opens it as
  text: add what the sheet has no words for (a CHECK across columns,
  `check (shipped is null or shipped >= placed)`; `on delete cascade`
  on a reference, so a line goes with its order; NOT NULL, UNIQUE,
  DEFAULT), and *Apply*: the table is made again from it, its rows
  copied by the columns both have, its indexes and triggers kept. A
  definition its rows break is refused in words, and nothing changes;
* **its indexes**: each with the columns it's by (× drops one; those
  its definition makes, its key's and a UNIQUE's, are changed there).
  Tick columns, in the order wanted, and *each once* for a unique one,
  then *Add an index*;
* **its triggers**: each with its SQL (× drops one). *+ Trigger…*
  writes one (`create trigger … begin … end`), for a value filled in or
  a change refused (`select raise(abort, 'what to say')`), which the
  sheet then says in those words;
* **Drop table…** (or *Drop view…*): it asks first, naming the cells
  whose queries read it, and is refused while another table refers to
  it, naming that table.
Each is one undo step. Undo brings a dropped table back with its rows,
indexes and triggers, and a definition back as it was, rows and all;
what was set aside for that stays out of the file you save. A column
renamed or dropped is a definition edited too.

Tables
can also be added with another SQLite program (DB Browser for SQLite,
the `sqlite3` command): open the `.recdep.db` there, add them, close
it, and open the file here again.

**Queries.** A formula can read the sheet's database:

```clojure
(query "select id, task, hours from tasks where hours >= ? order by id" B1)
```

The rows are laid out as a grid of their own in the cell, headed by the
columns' names, and read like any grid: `(SUM A1>hours)`,
`(count A1)`. Each `?` takes the next value after the SQL (a cell, or
any value). A cell shown as a date (by its kind or its number format)
goes in as the database's text for it, `2027-01-04`, so `(query "…
where placed >= ?" from-date)` works as it reads (for its day number
instead, pass a sum, `(+ from-date 0)`). A value that isn't a
number, text, a date or TRUE/FALSE (a function, a list) is refused,
naming it: `TRUE` written in a cell's formula is Excel's function; the
value is `true`. A column the database gives no type (a value worked
out, `max(placed)`) takes the kind its values have: ISO dates are
dates, whole numbers whole numbers.
* **Cells named in the SQL.** Instead of `?`, name the cell, as a
  formula would, after a colon:

  ```clojure
  (query "select * from orders where customer_id = :customer-id and placed >= :from-date")
  ```

  A name, a path (`:$>lists>cutoff`) or an address (`:B2`). It is a
  reference like any other: the query asks again when the cell changes,
  renaming the cell rewrites it, copying the formula moves an address
  as it moves any, and the formula bar colours it. Its value is bound,
  never written into the text, and a date cell goes in as a date. A
  name inside quotes (`':x'`) or a comment isn't one. Write spaces
  around an operator after a name (`:a > 5`, not `:a>5`, which reads as
  the path `a>5`).
* **Typing SQL**: inside a query's SQL in the formula bar, the
  database's tables, views and columns are offered as you type a word
  (Ctrl-Space shows them all), and after a colon the sheet's names.
* While the database is asked, the cell shows **#WAIT**, and what reads
  it waits too. That isn't an error: the status bar doesn't count it.
* SQLite's refusal shows as **#SQL**, with its reason: an unknown table
  or column, more than one statement, or a statement that writes
  (queries only read).
* A query asks again when its values change, and when the sheet's
  tables do (a table copied in, another database opened).
* F9 asks every query again: the way to retry after the database
  couldn't be reached.
* `query` works in a cell's own formula, not inside a function of a
  library.
* **By steps** instead of SQL: `(query "orders" [{:op :filter :column
  "customer" :test := :value B2} {:op :sort :by [["day" :desc]]}])`,
  the Data tab's steps run in the database, a step's value a cell. In
  the Data tab, on the sheet's database, *Promote live ▸* writes one.
  **A search form** is a query by steps: a filter whose value is blank
  is no condition, so each criterion is a box, left blank for any,
  with no SQL. **A record's key** is the other way round: a filter
  given `:if-blank :none` matches no row while its cell is blank, so a
  form showing one customer by `(query "customers" [{:op :filter
  :column "id" :test := :value customer-id :if-blank :none}])` shows
  none when its key is emptied, not every customer. A search over orders with their customers' names and
  totals searches a view the database has (`create view order_summary
  as …`): a view's columns that come from a table are typed into as
  that table's.
* **Its steps in the Inspector.** Select a query's cell written by
  steps, `(query "orders" [step …])`, and the Inspector lists them as
  it lists a `steps` formula's: click one and the sheet shows the table
  after it in the cell (the database is asked for it), with its form
  below, the Data tab's, to change it; *+ Step* adds one (the Data
  tab's kinds of step, a SQL step among them, which may name cells as
  above). A filter on a cell has *When it's blank: any row* (a search
  box) *or no row* (a record's key). Each change rewrites that step in
  the formula, the others as written.
* **How many it found.** A search ending in `{:op :limit :n 100}`
  shows its first 100 rows, and `(count found)` counts those;
  `(count-found found)` counts every row its other steps found, so a
  form can say *100 of 1,234*. The grid's cell says it too when you
  point at it. A total of every row found is a second query, its last
  step a `:group` (`{:op :group :by [] :aggs [{:fn :sum :column
  "total" :as "total"}]}`).
* While a query is asked again its rows are drawn dimmed.
* A query giving more than 200,000 cells shows the first rows, and its
  cell says so; a formula reading all of it (a whole column, or the
  query's cell, as `(count A1)`) is refused: work the total out in the
  query. A single cell of it reads as it is. Such a query takes a few
  seconds to lay out.

**Writing.** Rows of the sheet's database are written, never by a
formula, but:
* **by typing into a query's grid**, in a column that comes from a
  table whose key the query also selects (`select id, task, hours from
  tasks`: task and hours can be typed into; a worked-out column such as
  `hours * 2` can't, and says so). The formula bar says what typing
  there writes (*typing here writes hours in the row of tasks with id
  2*), and Enter on its last row stays there, the write said;
* **by typing a new row below a query's grid**: a query of one table's
  rows, its key selected (`select * from products`), has an empty row
  after its last. What's typed there is held, drawn in italics, and the
  cells the table can't do without (NOT NULL, with no default) say
  *needed* until they have values; the status bar says what it still
  needs. The moment they all have values the row is added to the
  table, as one undo step: Undo takes the row out of the database, and
  the empty row holds again what it held before the last value. So
  type the columns it can do without first, or type them into the
  record once it's added (a sorted query may show it elsewhere); what's
  typed into the empty row after that starts another. A row the
  database refuses (a code that's there already) stays held, its cells
  saying why: change one and it's tried again. Delete over its cells,
  or right-click it → *Delete ▸ rows*, lets it go; Delete, paste or
  fill over a column that runs through it writes the records and
  leaves it as it is. A column the table needs that the query doesn't
  show but fixes with an `=` filter of its steps (an order's lines,
  `{:op :filter :column "order_id" :test := :value order-id}`) is given
  that value: a line is typed as a product and a quantity, and joins
  the order shown. A query that doesn't
  show any other column the table needs, or whose rows are two tables', has no
  such row (an `insert-button` adds rows there). A row added that the
  query's conditions or order leave out is said, not left to vanish.
  What's held isn't a cell's value: nothing reads it, and a column read
  by its heading (`(count found>placed)`) ends above the row. *Convert
  to a plain grid* keeps what it holds, as the table's last row;
* **by typing into a view of a query**: `(record-view (query "select *
  from customers where id = ?" B1))` is a record laid out as a form, a
  field a row, and typing over a field writes the customer's row, as
  typing into the query's own grid does. So does a `record-view` or
  `filter-view` of a query's cell read by its name (`(record-view
  details)`), or of its rows picked (`picked`). A field shows a date
  as a date. A record view of one row (a query by its key) has no
  *1 of 1 ◀ ▶* or search box: it is a record as a form. A value the formula changed before showing it (`(assoc
  row :name "…")`) isn't the table's, and can't be typed over;
* **from a button or another control**, or the REPL:
  `(insert-row! ctx "tasks" {:task "Review" :hours 4})`,
  `(update-row! ctx "tasks" {:id 3} {:hours 6})`,
  `(delete-row! ctx "tasks" {:id 3})`. A row is found by its table's
  key; the keys of the map are the table's column names.

Delete, Ctrl-Enter, paste and fill down over a query's grid write those
cells' rows too, as one step (cells that aren't a table's column are
skipped). Delete blanks the cells; to delete rows, right-click them →
*Delete ▸ rows*: they go from their table (a row of a query joining two
tables is refused, saying so). A date picker or another control in a
query's cell, or in a view's field, writes its row.

**Checked before it's written.** What is typed, pasted or picked into a
query's cell, or a view's field, is checked first, as typing into any
cell is: against its column's type, and against the rules of the grid
it is in (a lookup that refuses what isn't there, values listed,
bounds), so *9999 isn't one of customers' id* is said before the
database is asked. The database's own rules hold too, and its refusals
say what broke: *customer_id 99999 isn't the id of a row of customers*;
deleting a customer who has orders, *11 rows of orders name this row of
customers (by their customer_id): delete them, or change their
customer_id first*.

**What the table says, the grid knows.** A query's columns take what
their table's own rules say: a `check (status in ('open', 'shipped'))`
gives the status cells a ▾ of those values, and anything else typed is
refused; a `check (qty > 0)` makes them whole numbers at least 1. A
column that refers to another table (`customer_id integer references
customers(id)`) is one of that table's keys: its ▾ lists them with
their names (*273 Acorn Kitchen Ltd*, found by typing *acorn*), the
cell shows the key with its name, and a key that isn't there is
refused. The name is the other table's column called name, title,
description or label; a table with none is named by the row it refers
to and its first date (an order: *Acorn Kitchen · 2026-11-20*); else
its first column of text. A table of more than 2,000 rows gives no ▾
(the database still checks). A record view of the query does the same
field by field. A rule written on the grid itself (*Column holds ▸*)
comes first.

**A box for a record's key.** A form's key, the box that says which
customer or order it shows, is one of a table's keys: give it the rule
`{:kind :one-of :of-table "customers" :refuse true}` (in the grid's
rules, as any kind). Its ▾ lists the table's rows by name, found by
typing part of one, the box shows the key with its name, and a key
that isn't there is refused in words (*isn't one of orders' id*), so
the form never shows an empty record for a number mistyped. No list
of the keys is kept in the sheet: the database is asked, and asked
again when the table changes. `:where` chooses the rows, as a query's
steps do: `:where [{:op :filter :column "discontinued" :test := :value
0}]` for the products on sale. Up to 10,000 rows; past that, the box
checks nothing (its ▾ shows the first 10,000). Copied into the
database, a column with such a rule refers to the table (REFERENCES).

**Where a table's records open.** Say it once, on the table's shape:

```clojure
(table-shape "customers" notes {:opens "$>customer-card>customer-id"})
```

Then every key of customers the sheet shows (a search's id column, an
order's customer_id shown by name, a customer's id among their orders)
is drawn as a link: a click writes the key into that cell and goes
there, as `(opens …)` does, and Back returns. A cell's own renderer
still wins; the form the key belongs to (the card's own key box, its
record's id) isn't a link to itself; a table whose shape says nothing
has no links. The ▾ of a key that is a link is still there when the
cell is selected.

**Forms without code.** The standard library has renderers for the
usual form actions, each a button labelled with its cell's value:
* `(opens "$>customer-card>customer-id")`, over a column of keys (where
  no table shape says where the table's records open, or to open them
  elsewhere):
  each id drawn as a link that writes it into that cell (a form's key)
  and goes there. Back returns to the search.
* `(insert-button "customers" new-customer {:opens
  "$>customer-card>customer-id"})` adds a row. Its values are a map
  (`{:customer_id customer-id :placed (TODAY)}`), or a form: a small
  grid with headers named as the table's columns (*Credit limit* is
  `credit_limit`: case, spaces and dashes don't matter) and one row
  typed in. Once the row is in, the form's typed cells are emptied (a
  formula in it, such as `(TODAY)`, stays), and with `:opens` the new
  row's key, the number the database gave it, goes into that cell, so
  *New order* opens the order it made. The form knows its table: the
  cells of its first row the table can't do without (NOT NULL, no
  default) say *needed* until typed into, its columns take the table's
  rules as a query's do (a CHECK's values as a ▾, its bounds, a
  reference as a ▾ by name; a rule written on the form comes first),
  and a heading that is none of the table's columns is flagged, saying
  which there are, before Add is pressed.
* `(update-button "orders" {:id order-id} {:status "shipped" :shipped
  (TODAY)})` changes the row the key finds.
* `(delete-button "customers" {:id customer-id} {:empties
  'customer-id})` deletes it; with `:empties`, that cell (the form's
  key) is emptied once the row is gone, so the form shows no record
  rather than one that isn't there.
* A button whose key is blank (no record shown) is drawn as refused,
  dashed: pointing at it says why (*Nothing to change: id is blank, so
  no row of orders is shown*), and so does pressing it.
* What the database refuses is said in its terms: a NOT NULL, a
  UNIQUE, a reference, a CHECK by its rule, and a trigger's own words
  (`raise(abort, 'that product is discontinued: pick one on sale')`).
  Rules the database keeps itself, a trigger filling a line's price
  from the product's or refusing a discontinued product, are the
  sheet's rules too: nothing in the sheet repeats them.
* **Records kept as they were**: a record that is history (a shipped
  order) is kept by triggers refusing changes to it, which hold however
  the change comes, with the lock open too. Undo still takes back what
  you did (a *Mark as shipped* undone): it puts the rows back as they
  were without running the triggers. A protect rule with a condition
  (§47) shows the same thing on the form before anything is typed: the
  order book's
  `create trigger order_kept before update of customer_id, placed, shipped on orders when old.status <> 'open' … select raise(abort, 'order … is shipped, so its details are kept as they were: set its status back to open to change them')`,
  and on its order card `{:protect true :when "(not= $>order-card>order-head>B5 \"open\")"}`.
A drop-down with no options of its own (`(dropdown)`, offering its
cell's kind) has a blank first choice, which empties the cell: a
search's criterion goes back to *any*. A lookup's ▾ over a list of a
query (a list of keys and names, `one of lists>staff>id`) shows each key with its name
(the list's column of text), as a task's ID shows its name. A cell
holding a shape says what it is: *a shape of customers: id, name, …
(9 columns, 3 needed)*. An input that a button's values read is read:
it isn't flagged as read by nothing.

Each says what it wrote, and each write is one undo step, with what
followed it (a key written into the form, a form emptied): one Undo
takes both back. For anything else, write a button's
handler: `insert-row!` takes a function given the new row's key,
`(insert-row! ctx "orders" {…} (fn [k] …))`, and `go-to!` goes to a
cell, `(go-to! ctx "$>order-card>A1")`.

Each write is one undo step: Undo takes it back out of the database,
Redo puts it in again. The queries reading the table ask again. A table
needs a primary key to be written to.

**Moving a table in.** Right-click a heading → *Move this table into the
sheet's database…*: its rows go into a new table, and the grid becomes
that table's query, keeping its name, formats and widths, so formulas
reading it (`(SUM tasks>hours)`) go on working, and typing into it
writes the table. A table without an IDs column is given an `id`
column, last. Cells holding formulas give the table their values, and
it asks first. Undo brings the grid's cells back; the table stays.

**Opening one.** Open takes a `.recdep.db` as it takes a `.recdep.edn`,
from a file or a web address (`?open=…`). A SQLite database that isn't
a sheet's is refused: open it in the Data tab instead (Open file…),
where it is read and nothing writes to it. Insert refuses a
`.recdep.db`: a grid in a cell can't hold its tables.

**Between visits.** The browser keeps the sheet's database with the
autosaved sheet, so a reload finds it again. Opening or starting
another sheet lets it go; the file you saved keeps it.

**Back to a plain sheet.** Save and pick *RecDep sheet (.recdep.edn)*.
That works while the database has no tables of yours and no cell queries
it; otherwise it is refused, naming them, since a `.recdep.edn` can't
hold them. Right-click a query's cell → *Convert to a plain grid* keeps
its rows as values.

**One tab at a time.** A sheet's database can be open in one tab of a
browser. Another tab starting the same sheet shows it, and says so in
the status bar; it can't save until the first tab is closed and it is
reloaded.

What's in the file: an ordinary SQLite database. The sheet is in a
table named `recdep_sheet`; tables whose names start with `recdep_` are
RecDep's and aren't listed in the Data tab.

## 54. Facades

A sheet that works something out is laid out for whoever built it:
working values, steps, notes to self. The person who uses it wants the
few figures they put in and the few they get out. A **facade** is a
grid made for them: it holds chosen cells of the working sheet (the
model) and your own text, laid out as you like. Its cells don't work
anything out; each shows a cell of the model, and a value typed into it
goes to that cell, so the model works it out again and the facade shows
the new results.

* **A link** is a cell whose whole formula is one reference to one cell
  of the model: `$>model>rate`, `$>model>B2`. It shows that cell's
  value. Written by name where the cells have names, it reads as what it
  is.
* **Paste link** builds them: copy cells of the model, go to the
  facade, right-click → *Paste special ▸ Paste link*. Or pick them one at a time: type `=`
  in a facade cell, go to the model's cell (the Outline, the breadcrumb,
  a sheet tab: the formula stays open, §11) and click it (point mode,
  §5); Enter brings you back to the facade. Each copied cell
  becomes a link at the same place in the block, by name where it has
  one, in its cell's number format. A label copied with its value is
  linked too, so the facade's label follows the model's when it's
  reworded. Blank cells copied are left out.
* **Drawn as what it shows**: a new link on a facade takes the renderer
  of the cell it shows, so a grid or table picked into it is drawn there
  as a grid (`(grid-view)`), a chart as its chart, without typing it. A
  link that has a renderer of its own keeps it; change it in the ⟐ field
  as any cell's. A renderer given by a rule (the Rules tab) isn't taken.
* **The facade is the grid of $ named `facade`**, with every grid
  inside it: name a grid of $ `facade` (its tab: double-click to
  rename) and it is the sheet's facade. Its own sheets are grids inside
  it, named as you like, even as the model's are: `facade>resources`
  beside the model's `resources`. So a facade can be built as grids
  within grids, with its own way around. Its tabs come first, marked ▣:
  the facade, then its sheets. The Inspector says, for a cell holding
  a grid, whether it is the facade, part of it, or not a facade, and
  for a cell on a facade, what it does there: a link and where a value
  typed into it goes, a link shown only, a result, or part of the
  layout.
* **Advanced**: with a facade, the sheet tabs have *Advanced* (stacked
  layers), beside Home (§36). On, Home is `$` and every sheet has its
  tab, the facade's first. Off, Home is the facade, only its sheets
  have tabs, there is no + for new sheets, and the sheet opens on the
  facade: the sheet as its users see it. It is saved with the sheet.
  The project plan example opens so: Advanced shows how it's worked
  out.
* **What its users can do** while the lock 🔒 above the grid is closed
  (as a sheet opens):
  * Type into a link to a value the model holds (an input): the value
    goes into the model's cell, checked there as if typed there (its
    kind, validation and protection), in one undo step, and the status
    bar says which cell was written. These cells have the soft green
    frame of open cells, and the writes-through mark; the formula bar
    shows the model's cell, and a list's ▾ or a date's picker is the
    model cell's own.
  * Delete clears the model's cell; paste writes values into the model's
    cells (a formula copied goes as its value). Cut cells aren't moved
    onto links (they would replace them): copy and paste instead.
  * Text goes through only a link left open (below): elsewhere a link
    takes numbers, dates and the like, so a word typed by mistake
    doesn't turn a number into text.
  * A slider, tick box or other control on a link changes the model's
    cell, never the link. A view's controls work the model's view: a
    click on a bar of a Gantt shown through a link picks that task in
    the model's chart (and whatever reads it, a detail of the task,
    follows), its ◀ − + move the model's window, a filter's choices
    change the model's filter.
  * A link to a query's cell, or a record view's field, of the sheet's
    database (§53) writes that row, as typing into the query does: text
    too, since a database's column holds text as data. So a facade lays
    a record out by links.
  * Type into the facade's own inputs: a cell of the facade itself,
    named, holding a value or nothing, that a formula reads (a search
    box: `find-name`, read by the query beside it). It takes typing,
    text too, with no model cell, link or *Leave open*.
  * Nothing else: a link to a result (a formula in the model) says it's
    worked out from other figures and can't be typed over; labels, text
    and the facade's layout say they're part of the facade. Only a
    value goes through a link, never a formula.
* **Choosing what can be typed into**: every link to an input takes
  typing, except links to text, which are taken for labels. A link
  left open takes text too. Select
  cells → *Keep … for showing only* for an input to show but not change;
  *Leave … open to typing* for a link to a text input (a name), or for a
  cell of the facade's own.
* **Editing the facade**: open the lock (*🔓 Editing all*). The facade is
  then an ordinary grid: retype a link to point it elsewhere, change
  labels and layout. Typing a value over a link replaces it, and the
  status bar says so, with Undo to get it back. The same is said when a
  link into another grid is typed over before its grid is a facade.
* **Hiding the model**: with Advanced off the model has no tabs. A grid
  named with a leading `_` has none even with it on (§36).
* **What it can't do yet**: a link is the whole formula: a link to a
  range or a formula using model cells shows its value but takes no
  typing. A grid or table linked and drawn as a grid (`(grid-view)`)
  takes typing in its cells, which goes to the model's cells (checked
  there, refused where they're results). A table linked as it is takes
  new records too, typed into its empty last row (§48); a sorted or
  filtered one doesn't. An error in the model shows on the
  facade as anywhere, naming the model's cell.

## 55. The toolbars

Two rows of buttons sit above the formula bar. They show icons, not
words; hold the pointer over one and its tooltip says what it does,
starting with the name this manual uses for it.

**The top toolbar** is about the sheet as a whole, left to right:

* **New** (a page with +): a blank sheet in place of this one, asked
  first.
* **Save** (a floppy disk): the whole sheet, wherever you are zoomed
  in. **Save as…** (a floppy with a pencil) appears beside it once the
  sheet has a file, to save to another. To save only the grid you're
  in, use **Save grid** in the bar above the grid (§6, *Saving grids
  and inserting files*).
* **Open** (a folder): a saved sheet or an Excel workbook in place of
  this one.
* **Insert** (an arrow into a tray): a saved sheet's grid, a CSV file
  or an Excel workbook into the selected cell, as a grid.
* **Library ▾** (books): the library file (§14).
* **Undo** and **Redo** (curved arrows).
* (The lock, §47, is in the bar above the grid, and *Advanced*, §54,
  beside Home in the sheet tabs.)
* **Light / Dark / Auto**, a drop-down (§4, *Light and dark*).
* **Examples ▾** (a light bulb), at the far right: the guided tour and
  the example sheets (§46).

**The format bar**, the second row, is about the selection and the grid
in view:

* **Bold**, **Italic**, text colour (A), fill colour, and align left,
  centre and right.
* The **number format** field (*General* when it has none, §10) and
  the **unit** field (§25). Both are framed as fields to type
  in; clicking either opens its popover. The popover's heading names
  the selection: a cell with a name by its name, its address after it
  in grey (*Number format for* **price** (B2)); otherwise by address.
  The renderer popover, opened from the formula bar's second line, does
  the same.
* **AutoSum** (Σ), **Fill down** and **Fill right** (arrows to a line;
  Ctrl-D and Ctrl-R, §15).
* The grid's size: **rows** and **columns** of the grid in view (§6),
  and **Mesh** (a lattice, §41), whose track sizes and *show lines*
  appear beside it while it's on.

A grid goes into a cell from the Inspector (*Make grid*), or from cells
already filled with right-click → *Make this a table* (§19).

## 56. Web requests and schedules

A formula can read what a web address answers: prices from a market
data service, exchange rates, anything a REST API gives as JSON. It is
asked again on a schedule, by F9, or when a cell you choose changes.
The example *Examples ▾ → Share portfolio* values shares at live prices
this way, and the tour's *Web requests and schedules* asks for exchange
rates.

* **`(web url)`, `(web url {:query {…}})`** reads `url` with GET. It
  shows `#WAIT` until the site answers.
  * A JSON answer becomes data: maps with keyword keys, and lists as
    vectors. Read it with `get-in`: `(get-in r [:rates :GBP])`. A list
    of maps is a table, so steps, views and charts take it. Any other
    answer is text. `{:as :json}` or `{:as :text}` says which, whatever
    the site says.
  * `:query` holds the parameters, each a value or a cell:
    `{:query {:symbol B2 :range "5d"}}`. They're never written into the
    URL text, so a cell's value can't break it, and a URL with `?…`
    written in is refused, saying so. A list is joined with commas:
    `{:symbols ["AAPL" "MSFT"]}` sends `AAPL,MSFT`.
  * **GET only.** A formula runs again whenever RecDep needs its value,
    so it never changes anything on a site.
  * Several requests in one formula are asked at once:
    `(for [s symbols] (web url {:query {:symbol s}}))`. The cell waits
    until all are answered.
* **When it's asked again.**
  * Without `:fresh`: once, then by F9. F9 (or ↻ F9) asks every request
    again, and the cells show their last answers, dimmed, until the new
    ones come.
  * **`:fresh x`**: the answer is good while `x` is what it was asked
    with. When `x` changes, the request is asked again, the last answer
    dimmed meanwhile. `x` can be any value: a schedule, or a
    `counter-button`'s cell to ask again at a click.
* **`(every schedule)`** is a schedule as a value: the date and time of
  its latest slot. The cell runs again when the next comes.
  * `"15m"`, `"1h"`, `"30s"` and `"1d"` are slots on the clock: 10:00,
    10:15… for `"15m"`; midnight for `"1d"`.
  * Cron's five fields, minute hour day month weekday:
    `"*/15 9-16 * * 1-5"` is every quarter hour from 9:00 to 16:45,
    Monday to Friday. Lists, ranges, steps and names (`MON-FRI`, `JAN`)
    are taken.
  * `{:zone "America/New_York"}` reads the schedule in that zone (the
    browser's otherwise), so a London sheet follows New York's market
    hours.
  * `(web url {:fresh (every "15m")})` asks every quarter hour. One
    `every` cell can be the schedule of a whole sheet.
  * **A page isn't a server.** Nothing runs while it's closed, and a
    hidden tab waits. Opening the sheet, or coming back to the tab, asks
    once for the slot that has come round. Slots missed meanwhile aren't
    made up.
* **Keys.** A service that wants a key is given `(secret "name")`:
  `{:query {:token (secret "finnhub")}}`.
  * The key itself is kept in this browser, never in the sheet, its
    file, a value or a message.
  * **Where it's given.** A sheet that uses keys gives each a field:
    a cell holding the key's name, `"finnhub"`, drawn as `key-input`.
    Paste the key there and press *Keep* (or Enter); the field then
    says *✓ kept in this browser*, with *Change* and *Forget*. The share
    portfolio's are in its Settings, under the provider.
  * **🔑 in the status bar** counts the places a key is wanted, as the
    error and comment counts do (*🔑 1 of 3 keys needed*, and how many
    are in the grid in view):
    * every key field whose key isn't kept in this browser yet,
      wherever it is;
    * for a key that requests wait for and no field names, the cell
      asking (its Inspector **Web** row takes the key too, and opens on
      arriving).

    Click it to go to the next, Shift-click the previous. A cell
    showing an old answer while it waits says which key in its tooltip.
    Once a key is given, the requests waiting for it are asked.
  * A key is sent only to the site it was given for. A request taking
    it elsewhere is refused, saying where it was given.
  * **Change** and **Forget**, in the field or the Web row, replace or
    drop a key.
* **The Inspector's Web row** lists a cell's requests: the address,
  its parameters (a key by its name, «finnhub»), and how each stands:
  answered at 14:02, asking, waiting its turn, waiting for a key or for
  a site that said it was asked too often, or refused (with the site's
  words).
* **Answers are kept in this browser.**
  * A sheet opened again shows its last answers at once, dimmed, while
    it asks.
  * A request whose `:fresh` value hasn't changed since (the same slot
    of a schedule) isn't asked again, which spares a service's daily
    allowance.
  * A sheet file can carry answers to start from (the example's prices,
    recorded when it was built). They're shown dimmed until its own
    requests are answered.
* **When a site doesn't answer**, the cell is `#WEB`, in RecDep's words:
  * *answered 401 …*, with the start of what the site said;
  * *answered a web page (HTML), not JSON*, often a check that a person
    is asking;
  * *took more than 30 s*;
  * *Couldn't read …: the site doesn't let web pages read it, or it
    can't be reached*.

  A site that says it was asked too often (429) is waited for, as long
  as it asks or from 5 s doubling, the cell saying until when. At most
  two requests go to one site at a time.
* **Sites a page can't read.** A web page may only read another site's
  answers if that site allows it (CORS). Market data services meant for
  apps (Finnhub, Twelve Data, Alpha Vantage) and Frankfurter's
  exchange rates do. Yahoo Finance doesn't. Such a site is read through
  a **relay** of your own, a small program on a server that asks the
  site and answers with the header that allows the page. Help's *Web
  requests* has one to copy (a Cloudflare Worker: free, one file). The
  share portfolio's Settings take its address.
* **In the share portfolio,** every share's symbol is a link,
  `(opens "$>share>ticker")` on its column: a click picks that share on
  the share sheet and goes there (Back returns).
* **A request for each item.** `(web url {:query {:symbol symbols} :each
  :symbol})` asks once for each value of `:symbol` (a list, or a
  one-column range such as `(UNIQUE (SORT trades>symbol))`). Its value
  is a table, a row per answer, `symbol` first, then the answer's
  fields. All are asked at once, and the cell waits for all of them.
* **A parameter in the address.** `(web "https://…/chart/{symbol}"
  {:query {:symbol "AAPL" :range "5d"}})` puts the symbol in the path,
  so APIs that take it there work with `:each` too.
* **A list parameter** may be a one-column range or a table's column:
  its values are joined with commas.
* **Answers that move with the date.** A request whose parameters
  differ only in dates written as text (news from a week ago to today)
  shows its last answer, dimmed, while the new day's is asked.
* **Answers as text: `(markdown "template")`.** The `markdown` renderer,
  given a template, draws a record, or each row of a table, through it,
  and the cell keeps the data:

  ```
  (markdown {:template "#### [{headline}]({url})
  *{source} · {datetime:d mmm, hh:mm}*

  {summary}" :between "---"})
  ```

  * `{field}` is the field's value; `{field:format}` is it in a number
    format.
  * Values are escaped (brackets and stars show as written), except as a
    link's or image's address.
  * A line whose fields are all blank is left out.
  * An image or link with no address is taken out, with a `·`, `|` or
    `–` it leaves dangling.
  * `:between` sets Markdown between rows (`---` is a rule); `(markdown
    "…")` is the same without it.
  * An image's size goes after a bar in its text: `![{name}|56]({logo})`
    is 56 pixels wide, `|56x40` both ways. Without one, an image in
    Markdown is kept to a few lines' height, so a large one (a service's
    1000-pixel logo) shows whole.

  `(fill-markdown template data)` gives the text itself, for a formula.
* `web` and `every` must be written in the cell's own formula, not in a
  library function a cell calls, as `query` and `tick` are.

## 57. Server databases

A sheet can read a database that lives on a server: Postgres, or a
SQLite file on that machine.
RecDep itself still needs nothing installed. Beside the database runs
**recdep-proxy**, a small program that RecDep talks to over the web.
It can also run on your own machine, to reach a database there, or one
your machine can reach. The database's password stays with the proxy;
RecDep only ever holds the proxy's token.

**Running the proxy.** It needs Node 22 or later. Build it once from
RecDep's source with `npx shadow-cljs release proxy`, which writes
`out/recdep-proxy.js`. Then write its settings in a file,
`recdep-proxy.edn`:

```clojure
{:listen {:host "127.0.0.1" :port 8787}
 :origins ["https://klibro.github.io"]        ; the pages it answers
 :databases {"shop" {:dbms :postgres :url {:env "SHOP_DATABASE_URL"}}}}
```

`:url` is the database's address, `postgres://user:password@host/db`,
here read from the environment variable `SHOP_DATABASE_URL` so the
password isn't in the file. Start it:

```bash
node out/recdep-proxy.js serve recdep-proxy.edn
```

It prints its **token**: RecDep asks for it once. `node
out/recdep-proxy.js token` prints it again, and `node
out/recdep-proxy.js check recdep-proxy.edn` says whether each database
can be reached, and what it is. Give the database login the proxy uses
the right to read, and to write only the tables it may write (below),
and nothing more.

**In the sheet.** Put the database's address in a cell, and name the
cell (say `shop`):

```clojure
(connection "http://127.0.0.1:8787/shop")
```

It shows as *shop (recdep-proxy at http://127.0.0.1:8787)*. Then read
from it with `query`, the connection first:

```clojure
(query shop "select id, name, city from customers where city = $1 order by name" B2)
```

`$1`, `$2` … are the values after the SQL, in order: cells or any
values, never written into the text. A cell showing a date goes in as
a date. The rows are laid out as a grid in the cell, read as any grid
(`(SUM A1>credit_limit)`), exactly as a query of the sheet's own
database (§53). The first query asks for the proxy's token. RecDep
keeps it in this browser, not in the sheet, so a sheet you pass on
carries no password and no token. Cancel, and the query says it wasn't
read; F9 asks again.

**What comes back.** Each column holds what its Postgres type says:
* whole numbers and numbers (exact decimals become numbers);
* text, and an enum's values;
* TRUE/FALSE;
* dates;
* a time with a zone in your browser's time zone;
* JSON and arrays as their text.

At most 200,000 cells come back. A formula reading more than that is
refused, as for the sheet's own database.

**When the database changes.** A change made through RecDep, in any
browser, shows at once in every other sheet reading the same tables.
A change made by another program is noticed within about half a
minute while the page is open, and the queries reading those tables
ask again (their old rows show dimmed meanwhile). F9 asks every query
again at once.

To have other programs' changes show at once too, the database's
administrator runs the triggers the proxy prints, once:

```bash
node out/recdep-proxy.js triggers shop recdep-proxy.edn
```

Then mark the database `:changes :notify` in the proxy's file. A table
added later needs them printed and run again.

**When it goes wrong.** A cell shows `#SQL` saying why:
* the database's own words for a mistake in the SQL;
* *queries only read* for anything that would write;
* that RecDep can't reach the proxy (is it running, with this page
  among its `:origins`?);
* that the proxy refused the token (F9 asks for it again).

`#WAIT` while the database is asked isn't an error.

**Writing.** Give the proxy the right to write, and the tables it may
write:

```clojure
{"shop" {:dbms :postgres :url {:env "SHOP_DATABASE_URL"}
         :mode :write :writable ["customers" "orders" "order_lines"]}}
```

Then a server database's query grids are written to exactly as the
sheet's own database's are (§53):
* typing into a cell writes its row (the query must select the table's
  key, such as `id`);
* Delete, paste and fill write the cells' rows;
* a row typed below the grid is added once its needed columns have
  values;
* right-click → *Delete ▸ rows* deletes them.

A column's rules come from the database: an enum's or a CHECK's values
give a ▾, and a column naming another table's rows offers them by name.
In buttons and the REPL, a table of a server database is written with
its connection: `(insert-button [shop "customers"] new-customer)`,
`(update-row! ctx [shop "orders"] {:id 3} {:status "shipped"})`.

Each write is one Undo step. Undo asks the database to put the rows
back, and puts the sheet back once it has. When someone else has
changed the row since, Undo says so and leaves both as they are, so
nobody's change is lost. Deleting a row the database deletes others
with (an order's lines) takes them too, and Undo brings them all back.
What the database refuses is said in words: a name already there, a
column left blank, a customer that eleven orders name.

Writes go to the database at once: there's nothing to save. A table
the proxy isn't given to write can't be typed into, and says why.

**Queries by steps.** The Data tab's steps work on a server database
too, given the connection first:

```clojure
(query shop "customers" [{:op :filter :column "city" :test := :value city}
                         {:op :filter :column "name" :test :contains :value search}
                         {:op :sort :by [["name" :asc]]} {:op :limit :n 20}])
```

They run in the database, live, as a search form wants: a blank
criterion is no condition, a date cell goes in as a date, and typing
into a row of the result writes it. Every step works there but three,
which say so:
* a filter by kind;
* an append;
* a join to a table of another source.

**In the Data tab.** *Open a URL…* and the database's address
(`http://127.0.0.1:8787/shop`). It's listed with its tables, as a file
is, and everything the Data tab does works on it:
* scroll through a table;
* *Details + data* for each column's counts, range, spread and
  histogram;
* *Correlations*, with all three measures, and a pair's cross-plot
  when you click a cell;
* steps to filter and clean;
* *Promote*, *Promote live* and *Copy into the database…*.

None of it is copied. Postgres works out each step, each column's
details and each screenful of rows, and only what you see is sent, so
the Data tab and your queries always agree, however large the table.
The steps only shape what you look at; the database is never changed
by them. Correlations are worked out by recdep-proxy, beside the
database, and only the matrix comes back. As with a file, past 200,000
rows they are worked out from an even sample of about 100,000, and the
note under the matrix says so. A cross-plot draws at most 5,000 points.

**A SQLite file behind the proxy.** A SQLite database on the proxy's
machine is served the same way, named by its file:

```clojure
{"notes" {:dbms :sqlite :file "/srv/data/notes.db" :mode :read}}
```

A path that doesn't start at `/` is from the settings file's folder.
It needs Node 22.13 or later. Everything above works on it as on
Postgres:
* queries, written in SQLite's own SQL, with `?` for each value;
* steps;
* writing, with Undo and Redo;
* the Data tab's tables, rows and steps.

Where they differ:
* a column read through a view is written back to the view's table;
* a change made by another program is noticed by asking (it can't be
  told), so `:changes :notify` isn't for SQLite;
* in the Data tab, a column's details and *Correlations* say, under
  the grid, that they aren't worked out for SQLite there. Open the file
  itself in the Data tab for them, or promote the rows into the sheet.
  A filter's list of values and the step builder's previews work.

A query that runs past the proxy's time limit (`:limits {:time-ms
10000}`) is stopped, and the next one is answered as usual.

**Serving others.** Run beside a server database, the proxy serves
everyone who should reach it, each with a token of their own and the
databases they may use. Three things change in its file:

```clojure
{:listen {:host "0.0.0.0" :port 8787 :tls {:cert "fullchain.pem" :key "privkey.pem"}}
 :origins ["https://klibro.github.io"]
 :auth {:mode :users}
 :databases {"shop" {:dbms :postgres :url {:env "SHOP_DATABASE_URL"}
                     :mode :write :writable ["customers" "orders" "order_lines"]}
             "notes" {:dbms :sqlite :file "/srv/data/notes.db"}}}
```

* **HTTPS.** `:tls` names the certificate and its key (from Let's
  Encrypt, say). Off this machine the proxy won't listen without it,
  since its tokens and your data would cross the network unencrypted,
  and a page served over HTTPS can't reach it anyway. When a server in
  front of the proxy (nginx, a load balancer) does HTTPS, write `:tls
  :in-front`. A renewed certificate is taken up by sending the proxy
  `SIGHUP` (`kill -HUP`), with no restart.
* **A token per user.** `:auth {:mode :users}` keeps the users in
  `recdep-proxy-users.edn` beside the settings file. Add one, saying
  what they may use (a database alone is read, `:write` writes it too):

  ```bash
  node out/recdep-proxy.js add-user alice shop:write notes recdep-proxy.edn
  ```

  It prints Alice's token once: give it to her alone. Only a hash of it
  is kept, so a lost token is replaced by running `add-user` again,
  which also changes what she may use. `remove-user alice` takes her
  token away, and `users` lists who may use what. Each takes effect at
  once, without a restart.
* **What each user sees.** A user is shown only their databases.
  Writing needs both the database's `:mode :write` and their `:write`;
  someone who may only read is told so, and told who can change it.

In RecDep nothing changes: the first query asks for the token, here
the one you were given.

**In a container.** `scripts/recdep-proxy.Dockerfile` makes an image of
the proxy:

```bash
docker build -f scripts/recdep-proxy.Dockerfile -t recdep-proxy out
```

Its settings, users and token live in `/config`, a folder of the host's
mounted there. The proxy runs as the image's `node` user (uid 1000), so
that folder must be writable by it (`chown 1000 /srv/recdep-proxy`). A
SQLite file is reached through a mount of its own, its `:file` the path
inside the container. Its `:listen` host is `"0.0.0.0"`, usually with
`:tls :in-front`:

```bash
docker run -d -p 8787:8787 -v /srv/recdep-proxy:/config -e SHOP_DATABASE_URL=postgres://… recdep-proxy
```

```bash
docker run --rm -v /srv/recdep-proxy:/config recdep-proxy node recdep-proxy.js add-user alice shop:write /config/recdep-proxy.edn
```
