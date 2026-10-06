# HFG frame standard

Every HFG model sheet follows this layout, so models look and behave the same whatever their content. The code version is `hfgmodels/frame.py`; this page explains it.

## Workbook

- Built from a copy of Kelvin's template (`Budget_Template.xlsx`), so the theme, named styles, scenario manager, form controls, summary pages and Python in Excel charts stay intact.
- Sections, each with a cover sheet in the template: Dashboards, Financial Model, Reports, Appendices. New model sheets go into the Financial Model section; their checks sheet goes into Appendices.
- Contents, section covers and links follow Modano's layout and are generated, never typed (decided 6 October 2026): a Contents sheet first with sections numbered 1, 2, 3 linking to their covers, sheets lettered a., b., c. within each section and the headings on each sheet marked "-"; a cover sheet before each section (title, "Section N.", model name, links to the contents and the sheets either side, notes); and on every other sheet A1 linking to the contents and A2 to the checks. Links are cell hyperlinks to `HL_` names with a screen tip, as in Modano's example (changed 7 October 2026 from the HYPERLINK formulas `prototypes/models/navigation.py` still writes; the probe's `links` test checks a link to a name follows it), rebuilt with every structural change. The template's typed Cover contents is replaced by the generated one.
- The template's Time sheet holds the timeline settings: start month (April 2025 in the template), term in months (`Ts_Term`), last actual month (`DD_Ts_Last_Hist_Mth`).

## Sheet layout

| Area | Use |
|---|---|
| Column A | Navigation: A1 back to the table of contents (`HL_Home`), A2 to the checks (`HL_Err_Chk`) |
| Columns B to G | Labels. B for section bars and sub-headings; C, D, E for indented items. B to F are narrow (2.5), G is wide (34) so labels can run across |
| Column H | Units (`$`, `%`, `flag`) on time series sheets; the value on assumption sheets |
| Column I | Row total (sum, maximum or closing value) on time series sheets; units on assumption sheets |
| Column J on | Monthly timeline, one column per month |
| Rows 1 to 3 | Sheet title, model name line (`Model_Name`: the title with the error, scenario and alert phrases), entity name (Cover B1) |
| Rows 5 to 15 | Timeline block (time series sheets): month ending, month with (H) or (F), period start and end, counter, financial year, active column, month number, month, quarter and half keys |
| Row 16 | Frozen panes sit below the timeline and right of column I |
| Row 17 on | Content in sections |

Content and heading rows are 15 high (proposed; Kelvin to confirm), with short spacer rows as in Modano's example: 3 points between the drivers of one line, 6 before checks and between groups of lines, 9 before a block's summary. Gridlines are off.

## Content blocks

- **Section bar** (Heading 1., theme accent 1) for each major block: Portfolio, then one per site.
- **Sub-heading** (Heading 2., grey band) for groups of rows inside a section.
- **Portfolio first**: every time series sheet starts with a portfolio block that adds the sites flagged "Include in portfolio", followed by one block per site in the same order as the inputs sheet. This is the category pattern: adding a site means adding one block to each sheet.
- **Totals** use the bold total format with a top border.
- **Checks** sit next to what they test (balance check rows, cash checks) and roll up to the model's checks sheet.

## Cell formats

Formats come from the template's named styles, so changing the theme recolours every sheet.

| Key | Template style | Use |
|---|---|---|
| `in_num`, `in_pct`, `in_text`, `in_date` | Assumption Number., Assumption Percentage., Assumption Heading. | Inputs (light accent fill, thin border) |
| `num`, `num1`, `pct`, `date` | Number., Percentage., Date. | Calculations |
| `total` | Number. bold with top border | Totals |
| `section`, `sub`, `h3`, `label` | Heading 1. to 4. | Structure |
| `check`, `check_flag` | Number. in dark red | Check results |

Formats the template lacks (input dates, two-decimal inputs, check colours, multiples) are derived from an existing format and appended to `styles.xml`. Existing formats are never changed.

## Naming

- Time: the template's `Ts_` names (`Ts_Start_Date`, `Ts_Term`, `Ts_Mths_In_Yr`) and `DD_` drop-down links (`DD_Ts_Last_Hist_Mth`).
- Model-wide inputs: model prefix, for example `Dev_GST_Rate`, `Dev_Fund_LVR`.
- Site inputs and derived values: site code prefix, for example `S01_Sale_Date`, `S01_Limit`. Every scalar input and derived value has a name, so formulas read like sentences.
- Check totals: `Dev_Err_Chk` (errors, should be nil) and `Dev_Alt_Chk` (alerts to review).

## Historical then forecast

Months up to `DD_Ts_Last_Hist_Mth` are actuals; later months are forecast. Until the Home Hub actuals feed exists, actual months carry no flows and each site starts from an opening position (WIP, facility balance) at the end of the last actual month.

## Formula rules

- No circular references and no iterative calculation. Interest is charged on opening balances.
- No volatile functions in model logic (INDIRECT, OFFSET, TODAY).
- Functions must work in Excel 365 and LibreOffice. Newer functions are stored with their `_xlfn.` prefix.
- Every calculated row is reproduced by the reference calculation in Python and tested.

## Look and wiring (7 October 2026)

Kelvin asked that models look and behave like Modano's example. The spec's Look and wiring tab is the full standard; the rules the engine follows are:

- Titles: `Model_Name` is the title plus error, scenario and alert phrases, each switchable in Settings; summary and report titles are formulas of name, denomination and period; every label downstream of its source is a formula of the source label.
- Styles: every style refers to theme slots only (accent 1 for section bars, inputs and the first chart series; the hyperlink colours for links; Segoe UI as both theme fonts), so a model takes its entity's brand from the theme. Inputs: accent 1 at 80% tint fill with a thin accent 1 at 40% tint border all round.
- Emphasis: the last item of a list has a dashed rule beneath; major results are bold with a thin rule above; detail and ratio lines on summaries are italic; check cells turn bold red when not zero.
- Grouping: a module's heading bar at level 0, its body at level 1, working rows at level 2; statements detail at level 2 under totals at level 1; the timeline block rows at level 2; hidden historical and inactive columns grouped at level 2.
- Conditional formats travel with their rows: inputs switch off in actual months, unused method inputs grey out, schedule inputs show only in scheduled months, the selected period and active scenario are highlighted.
- Charts read a source block at the foot of their summary (selector, period rows, two-line axis labels, data rows by INDEX), coloured by theme slot.
- A new model comes from the New model wizard: entity (name, logo, theme), model, timeline, display, scenarios; the package writer builds the file.
