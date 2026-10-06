# Example models (Phase 0)

Two finished example models, laid out the way the add-in will hand models over, and the screenshots of them. Kelvin asked for the contents wired in as Modano's, updating as sheets are added, with links back to the contents and the checks, and for screenshots of a budget and actuals model and a development model.

```
python prototypes/reports/build.py out.xlsx                                       # budget and actuals example
HFG_TEMPLATE=/path/to/Budget_Template.xlsx python prototypes/models/development.py out.xlsx
HFG_TEMPLATE=... python prototypes/models/screenshots.py budget|development out_dir   # PNGs and one captioned PDF
pytest tests/test_models.py                                                        # 5 tests; the development one needs HFG_TEMPLATE
```

Workbooks, screenshots and the template stay out of git.

## Contents, section covers and links (navigation.py)

The same layout as Modano's, written into a built workbook:

- **Contents**, first: the model name with the error count, notes, then the table of contents. Sections are numbered and link to their covers, sheets are lettered within each section, and the headings on each sheet sit under it marked "-". The error and alert totals close the page.
- **A cover sheet before each section**: title, "Section N.", model name, a link to the contents, links to the sheets either side, notes.
- **Every other sheet**: A1 links to the contents, A2 to the checks and shows a tick while the error checks are clear.
- **Links are HYPERLINK formulas to names** (`HL_Home`, `HL_Err_Chk`, `HL_Sheet_<sheet>`, `HL_Toc_<sheet>_<n>`), and entries read each sheet's title cell, so renames keep them right.

Sections are read from the workbook's order: a cover starts a section and the sheets after it belong to it. So `refresh` after a sheet is added, moved or removed files it where it sits and rewrites the contents, the letters and the cover links; the test adds a sheet after Inputs and moves Lookups, and both land where they sit. In the add-in the change plan does this on every structural change (proven live in `prototypes/assembly`).

## The two examples

| Example | Sections | Built from |
|---|---|---|
| Budget and actuals | Dashboards (four summaries), Financial Model (Time, Assumptions, Scenarios, Seasonality, Inputs, Statements), Reports (five reports), Appendices (Chart register, Lookups, Checks) | `prototypes/reports/build.py`: fictional building and maintenance business, actuals to September 2026, a budget phased by seasonality, three scenarios, 95 native charts |
| Development | Dashboards (Development summary), Development Model (sites, costs, sales and exit, GST, funding, statements, returns), Appendices (Time, Lookups, checks) | `development.py`: the development test example built into Kelvin's template, then finished as New model will: only the frame sheets it uses and the development sheets kept, a dashboard with key figures and four native charts, contents, covers and links |

Finishing the development example drops the template's other sheets, the names that pointed at them, and freezes to values any formula on a kept template sheet that read a removed one. The template's Lookups sheet carries the Income Statement Summary title, which is corrected. Two LET formulas on the template's Time sheet are Excel 365 only, so LibreOffice 24.2 shows them as errors; Excel calculates them.

## Screenshots

`screenshots.py` recalculates in LibreOffice, freezes values, and exports chosen sheet areas with row and column headers, then trims each image and puts them into a PDF with a caption per view. LibreOffice stands in for Excel, so fonts and chart details differ slightly from what Excel draws.
