# Logic Test 3 – Season / Determinant Logic (Common Property version)

Source workbook: `PQSpecshop_Validation_V4.xlsx`

Logic Test 3 answers one question: **for a shipment from a Source, of a Product, to a
Destination, arriving on an ETA date, what Min / Max limits apply to each IQDM
property?** The limits come from the Refinery DB spec sheet and the Terminal DB.

## 1. The sheets and their roles

| Sheet | Role |
|---|---|
| **Refinery DB** (4,097 rows) | Refinery spec limits: SpecificationSheet (A), Source (B), Product (F), HasSeason (I), Season (J), Tests/quality (M), SeasonStart (O), SeasonEnd (P), Min (R), Max (S) |
| **Terminal DB** (32,596 rows) | Terminal spec limits. Source (B) is the **terminal** (the destination). Product (F), HasSeason (I), Season (J), Tests (L), SeasonStart (N), SeasonEnd (O), Min (Q), Max (R) |
| **Terminals Specsheet Master Data** | From row 7: Source (C), Product (D), SpecificationSheet (E), up to 16 destination terminals (F:U). Source and Product are left blank on continuation rows. |
| **Property Mapping** | From row 7: IQDM Tag (A), IQDM Description (B), Common Property Name (C, only on the first row of each group), Quality Name (D), Source DB (E). It translates every raw quality name to one Common Property and lists the IQDM tags for each property. |
| **Product Mapping** | Spec-sheet product → Terminal DB product. The override list in J:K wins, e.g. `No. 2 MV 15 / Heating Oil` → `No. 2 MV 15`. |
| **Location Mapping** | IQDM ↔ PQ Specshop location names. Logic Test 3 does not use it. |
| **Logic Calc 3** | Hidden helper columns. All the row-level work happens here. |
| **Logic Test 3** | Inputs (B4:B7), checks (B8:B11) and the results table (row 13 onwards). |
| **Logic BPMN 3** | Process diagram plus a 43-step register. The step numbers below follow it. |

## 2. Step-by-step logic

### Stage 1 – Inputs (steps 1–2)
1. The user picks **Source**, **Product**, **Destination** and **ETA**. The dropdown lists come from `Logic Calc 3` P2#, Q2# and R2#.
2. **ETA key** = `MONTH(ETA)*100 + DAY(ETA)`. For example, 29‑Sep → **929**. The year is ignored, so a season such as 5/1–9/15 applies every year.

The Refinery path (stages 3–4) and the Terminal path (stage 5) then run independently from the same inputs.

### Stage 2 – Find the SpecificationSheet (steps 3–6)
3. In the master data, blank Source and Product cells are filled down from the row above (cols T:U). The 16 destination columns are joined into one key, `|dest1|dest2|…|` (col W).
4. A master row matches (col X = 1) when the spec is not blank, Source and Product equal the inputs, and `|Destination|` appears in the key.
5. **SpecificationSheet** (B8) is the spec on the **first** matching row. If no row matches, it is `Not found in master data`.
6. **Data check** (B10) reports one of three messages:
   - no spec sheet for this Source / Product / Destination,
   - the spec sheet was found but has no Refinery DB rows,
   - `n data rows matched; m applicable on ETA`, plus a warning when several spec sheets share the destination (the first one is used).

### Stage 3 – Refinery DB rows (steps 7–9)
7. **InputMatch** (K) = 1 when the trimmed SpecificationSheet, Source and Product all equal the inputs.
8. **Common Property Name** (D): look up `"Refinery DB|" & quality` in Property Mapping, where col C is carried down over blank cells. If the quality is not mapped, the raw quality name is used. Because of this, several different test methods (for example API Gravity by D1298, D287 and D4052) become one property.
9. Values read from each row:
   - **Min** (R) and **Max** (S) are kept only if they are numbers. Text such as `None` or `NULL` is ignored.
   - **Season** (J) is kept only when HasSeason = 1 and the value is not `NULL`.
   - **SeasonStart** (O) and **SeasonEnd** (P) are text in `m/d` form. Each becomes `month*100 + day`, so `" 5/1"` → 501 and `" 9/15"` → 915.

### Stage 4 – Season status per row (steps 10–17)
| Condition | SeasonStatus | Applied? |
|---|---|---|
| HasSeason ≠ 1 | **No season** | Yes – limits used as they are |
| HasSeason = 1 and Start or End missing or unreadable | **Dates missing** | No – flagged in "Season dates missing?" |
| Start ≤ End and Start ≤ ETAkey ≤ End | **In season** | Yes |
| Start > End (the window wraps the year end, e.g. 11/1–2/29) and (ETAkey ≥ Start **or** ETAkey ≤ End) | **In season** | Yes |
| otherwise | **Out of season** | No |

**Applicable** (M) = 1 when InputMatch = 1 and the status is *No season* or *In season*.

### Stage 5 – Terminal DB rows (steps 18–29)
18. The Terminal DB product is resolved through Product Mapping. The override list J:K is checked first; otherwise the input product name is used.
19. **TDB Match** (Z) = 1 when Terminal DB col B = Destination and col F = the resolved product.
20. The quality is translated with `"Terminal DB|" & quality`, the same way as step 8.
21–29. The same numeric filtering, season-key parsing, status rules and Applicable flag as stages 3–4 (cols AB:AH, AK).

### Stage 6 – Build the property list (steps 30–35)
30. **Refinery properties**: the unique Common Properties of matched Refinery rows, in the order they first appear (col O).
31. **Terminal-only properties**: the unique Common Properties of matched Terminal rows that are not already in the Refinery list (col AI).
32. The two lists are combined, Refinery first, then **stably sorted** so that non-seasonal properties come first and seasonal ones second (BG:BH). A property counts as seasonal if any of its matched rows has HasSeason = 1. Refinery rows decide this; Terminal rows decide it for Terminal-only properties.
33–34. For each property, find the Property Mapping rows whose carried-down Common Property matches and that have an IQDM Tag or Description. A property with no such row is **dropped**.
35. Each property is expanded to **one result row per Tag/Description pair** (AW:BC). For example, *API Gravity at 60F* produces AP1, AP2, AP3 and AP5.

### Stage 7 – Results per row (steps 36–43)
All columns are computed over every row that maps to the Common Property, so several qualities or methods are pooled together.

| Column | Rule |
|---|---|
| **D Min** | `MAX` of numeric Mins over applicable Refinery rows; blank if none |
| **E Max** | `MIN` of numeric Maxes over applicable Refinery rows; blank if none |
| **F/G Terminal Min/Max** | Same rule over applicable Terminal rows |
| **H Overlap?** | `Yes` if more than one applicable **seasonal** Refinery row; blank for Terminal-only properties |
| **I Rows applied** | Count of applicable Refinery rows |
| **K Season dates missing?** | `Yes` if any matched Refinery row has status *Dates missing* |
| **L / M Min / Max taken from season** | Season of the first applicable row (in data order) whose Min (or Max) equals the final value; `N/A` if that row has no season |
| **N Property source** | `Refinery DB`, or `Terminal DB only` |

Taking the MAX of Mins and the MIN of Maxes means that when several rows apply (overlapping season windows, or several pooled qualities), the **strictest** limits win.

**A – Season**
- *Terminal-only property* (step 39): if no Terminal row applies → `Dates missing` when any Terminal row has that status, otherwise `Out of season`. If rows apply → the unique Terminal season names joined with commas, or `N/A` if none.
- *Refinery property, Rows applied = 0* (step 41): `Dates missing` if K = Yes, otherwise `Out of season`.
- *Refinery property, rows apply* (step 42): the Max's season (M) if a Max exists, otherwise the Min's season (L), otherwise `N/A` if no applicable row is seasonal, otherwise all applicable season names joined with commas.

**J – Rule applied** (checked in this order)
1. Terminal-only → *"Terminal DB only – property not in the Refinery DB spec list; limits taken from Terminal DB"*
2. Rows applied = 0 → *"Season dates not in data – cannot be evaluated for ETA"* when K = Yes, otherwise *"No season window covers the ETA – no limits apply"*
3. Overlap = Yes → *"Overlap: Min = MAX of Mins, Max = MIN of Maxes; Season = row giving the final value"*, plus *"(Min from X, Max from Y)"* when the two seasons differ
4. Rows applied > 1 → *"Several qualities / rows pooled – Min = MAX of Mins, Max = MIN of Maxes"*
5. Exactly one applicable seasonal row → *"Single season window – limits used as is"*
6. Otherwise → *"No season – limits used as is"*

## 3. Worked example (inputs saved in the workbook)

Source `Canton, OH : MPC Refinery`, Product `84 Blend Grade Conventional`,
Destination `East Sparta, OH : MPL Tank Farm`, ETA 29‑Sep‑2026:

- ETA key = 929.
- Spec sheet = `87 Octane - Marathon & Buckeye Pipelines (OH, PA, WV)`.
- Refinery DB: 71 rows matched, 47 applicable. Terminal DB: 187 rows matched, 59 applicable.
- 50 result rows. Examples:
  - **RVP2 Vapor Pressure, EPA**: the only window is 5/1–9/15, and 929 is outside it, so the result is `Out of season` with Rows applied = 0.
  - **T50E 50% Evaporated (E10)**: season C‑3 covers 929, so Min 150 / Max 240 (Terminal 149.5 / 239.5) and *Single season window*.
  - **VL20E**: two seasonal windows apply, so Overlap = Yes and Min = MAX(Mins) = 116.
  - **FBP / TEP / DFBP**: these properties exist only in the Terminal DB, so the Terminal Max of 430 is used.

A Python re-implementation of the rules above reproduced all 50 Excel result rows exactly.

## 4. SQL implementation

- `sql/01_tables.sql` – staging tables, one per source sheet. `RowId` must keep the sheet row order.
- `sql/02_usp_LogicTest3_SeasonDeterminant.sql` – `dbo.usp_LogicTest3_SeasonDeterminant @Source, @Product, @Destination, @ETA`. It returns:
  1. a header (SpecificationSheet, ETAKey, TerminalDBProduct, DataCheck, TerminalDBCheck),
  2. the results table with the same columns as Logic Test 3 A:N.

Notes when comparing the procedure with Excel:
- Load Min and Max as `FLOAT`, using NULL for any non-numeric value. This matches Excel's `ISNUMBER` rule.
- Text matching depends on a case-insensitive collation, like Excel. SQL `TRIM` removes only leading and trailing spaces, while Excel `TRIM` also collapses double spaces inside the text.
- The workbook formulas cover only Terminal DB rows 2–31,967. The procedure reads the whole table.
