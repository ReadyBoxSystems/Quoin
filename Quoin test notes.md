Quoin notes:

Use this sequence in the running app at http://localhost:3000.

  1. Basic Values
  Import:

  import-test-01-basic-values.xlsx

  Expected:

  - One sheet: Basic Values
  - Imports as a new configuration
  - Plain labels, numbers, boolean, and date values appear in the grid
  - No formulas
  - No review items

  2. Basic Formulas
  Import:

  import-test-02-basic-formulas.xlsx

  Expected:

  - One sheet: Basic Formulas
  - Formula cells import with visible =... formulas
  - B5, B6, B7 should calculate
  - No review items expected

  3. Named Cells
  Import:

  import-test-03-named-cells.xlsx

  Expected:

  - Safe names promote to Smart Cells:
      - design_span
      - design_plf

  - Named range input_block becomes a review item
  - Unsafe name Design Span Unsafe becomes a review item
  - Formula using names should remain visible

  4. Multi-Sheet
  Import:

  import-test-04-multi-sheet-cross-sheet.xlsx

  Test twice:

  - First choose Inputs
      - Should import plain inputs
      - input_span should promote
      - No major review items

  - Then import same file again and choose Calculator
      - Should import formulas with Inputs!B3 / Inputs!B4
      - Should show cross-sheet formula review warnings
      - Formulas stay visible, even if Quoin cannot calculate them yet

  5. Reference Table Style
  Import:

  import-test-05-reference-table-style.xlsx

  Expected:

  - Visible lookup-table-like grid imports
  - selected_span should promote to a Smart Cell
  - beam_table named range should show as a review item
  - This is the fixture for the future reference-table direction

  6. Review Warnings
  Import:

  import-test-06-review-warnings.xlsx

  Expected:

  - Imports formulas unchanged
  - Review list should flag:
      - cross-sheet reference
      - structured table reference
      - external workbook reference
      - semicolon argument separator
      - spill marker / dynamic array marker

  - This file is meant to produce warnings

  After Each Import
  Check:

  - A new local configuration was created
  - Existing configurations were not overwritten
  - Grid dimensions expanded enough for the imported cells
  - Formula bar shows the imported formula when selecting formula cells
  - Inspector shows Smart Cell names where expected
  - Runner Preview stays mostly empty unless imported named cells are later surfaced manually



1-seemed to pass no problem
2-seemed to pass no problem
3-seemed to pass no problem
4/5-without a way to generate other sheets, Quoin doesn't know what to do with the information. we need a "tabs" feature or something similar. 
6-lots of reviews, which seemed to be the point.