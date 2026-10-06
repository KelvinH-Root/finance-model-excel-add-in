/*
  The command registry: every ribbon button, menu item, right-click item and shortcut the
  HFG add-in is designed to have, mapped to Modano's ribbon. One source for the manifest
  (tools/build-ribbon.mjs writes manifest.xml from it), the runtime Build tab and the
  placeholder views in the task pane. Labels and icons are HFG's own.

  In the probe every command opens a placeholder view that says what it will do, which
  Modano command it matches and the phase it lands in. Nothing here is built yet.

  Fields: key (short id), label (ribbon text), ctxLabel (right-click text, if different), glyph (icon), tip (supertip, 250 characters
  at most), modano (the Modano command it matches, or "New"), phase (1 to 3), view (what the
  command will do, shown in the pane), items (menu entries, same fields).
*/
(function (root) {
  'use strict';

  const MAIN = [
    {
      id: 'model', label: 'Model', modano: 'Modular Workbooks', controls: [
        { key: 'model-new', label: 'New model', glyph: '✚', phase: 1, modano: 'Modular Workbooks > New',
          tip: 'Build a new model from the HFG template: choose a recipe or modules, the entity, start month and term, preview, then build.',
          view: 'Opens the new model wizard. Pick a model type from the catalogue (or modules by hand), the entity and its palette, the start month, term and last actual month. The preview lists the sheets, modules and checks before the package writer builds the file and opens it.' },
        { key: 'model-adopt', label: 'Adopt workbook', glyph: '⇪', phase: 2, modano: 'Modular Workbooks > Open',
          tip: 'Bring an existing Modano-style workbook, such as the budget template or BUD25, under management by reading its frame, names and checks.',
          view: 'Reads the open workbook: its sheets, the frame (header rows, timeline block, columns), defined names (Ts_, DD_, CB_, LU_, HL_, Err_), checks and form controls. Shows what it recognised and what it could not, then writes the model metadata so the other commands work on it.' },
        { key: 'model-explorer', label: 'Explorer', glyph: '☰', phase: 1, modano: 'Build > Project Manager (Project Explorer; Composition, Links and Properties)',
          tip: 'The model in one pane: sections, sheets and modules coloured by area, and for the module selected its rows, its links in and out, its settings and checks.',
          view: 'Modano\'s Project Manager in one pane: a tree of sections, sheets and modules (coloured by area, with check status), and tabs for the module selected: Composition, Links (what it takes from and sends to, as a diagram), Properties and Checks. Selecting a module selects it in the workbook.' },
        { key: 'model-checks', label: 'Checks', glyph: '✔', phase: 1, modano: 'Checks sheet and header status',
          tip: 'Every error, alert and sensitivity check in the model, which ones are failing, and a go-to for each failing cell.',
          view: 'Lists every check by module with its status and include toggle. A failing check jumps to the cells that fail. The same totals drive the status shown in every sheet header.' },
        { key: 'model-props', label: 'Properties', glyph: 'ℹ', phase: 1, modano: 'Modular Workbooks > Workbook',
          tip: 'Model properties: id, recipe, entity, GST status, version, owner and change history.',
          view: 'Shows and edits the model properties held in the metadata: id, recipe, entity, GST registration, frame and template versions, owner, and the history of structural changes.' },
        { key: 'model-group', label: 'Group', glyph: '⋔', phase: 2, modano: 'Consolidations: entities summary, adding and deleting entities', type: 'menu',
          tip: 'The group as a tree with who owns what, each group\'s figures rolled up, and planned entities added or changed with their share held.',
          view: 'The group\'s entities as a tree (share held, owned by the top, outside investors, the groups each rolls into) and each group\'s surplus and net assets rolled up the tree.',
          items: [
            { key: 'grp-structure', label: 'Group structure', tip: 'Org tree of the entities with share held, owned by the top and outside investors, the groups each rolls into, and each group\'s figures rolled up.',
              view: 'The tree and the roll-up for the year shown: each group\'s members, its sub-groups consolidated and the eliminations made in it. Goes with the Group structure sheet in the model.' },
            { key: 'grp-add', label: 'Add entity', tip: 'Add a planned entity (a future LP or fund) under its parent with its share held, start date and capital; it stays flagged until Home Hub has it.',
              view: 'Code, name, parent, share held, member from, capital and type, with a preview of where it goes in the tree, the groups it rolls into and what the change plan writes.' },
            { key: 'grp-ownership', label: 'Change ownership', tip: 'Change an entity\'s share held or parent from a date; groups, eliminations and NCI follow from that date and earlier years keep the old structure.',
              view: 'Pick the entity, the new parent or share held and the date, and see the effect on ownership, outside investors and the groups it rolls into before applying it.' },
            { key: 'grp-remove', label: 'Remove entity', tip: 'Take out a planned entity, or end an entity that is sold or wound up from a date; its figures stay in the years it belonged.',
              view: 'A planned entity with no figures is removed with its rows. An actual entity is ended from a date instead: it leaves its groups from then, and earlier years keep it.' },
            { key: 'grp-refresh', label: 'Refresh from Home Hub', tip: 'Pull the entity register and ownership from Home Hub; differences show before anything changes, and planned entities stay until Home Hub has them.',
              view: 'Compares the model\'s register with Home Hub\'s: new entities, ownership changes and planned entities Home Hub now has (they become Actual). Applies through a change plan with a preview.' }
          ] }
      ]
    },
    {
      id: 'modules', label: 'Modules', modano: 'Modules', controls: [
        { key: 'mod-insert', label: 'Insert', glyph: '⊞', phase: 1, modano: 'Modules > Insert',
          tip: 'Insert a module from the library. Filter by model type, variant, GST treatment, timeline and version; preview its rows and links first.',
          view: 'The module picker: library on the left, filters for model type, variant, GST treatment, timeline and version, a preview of the rows, inputs, outputs and links. Insert places it on its area\'s sheet and links it to the rest of the model.' },
        { key: 'mod-replace', label: 'Replace', glyph: '⇄', phase: 2, modano: 'Modules > Replace',
          tip: 'Swap a module for another variant, for example revenue from amounts to price times volume. Inputs carry across by key and links are kept.',
          view: 'Choose the new variant. The preview shows which inputs carry across, which are new and which go, and confirms every link stays connected.' },
        { key: 'mod-duplicate', label: 'Duplicate', glyph: '❐', phase: 1, modano: 'Modules > Duplicate',
          tip: 'Copy a module with its own names and links.',
          view: 'Copies the selected module as a new instance with the next number, its own names and its links resolved again.' },
        { key: 'mod-mirror', label: 'Mirror', glyph: '⇋', phase: 2, modano: 'Modules > Mirror',
          tip: 'Add a module that repeats for every sender of a link, for example debtors for every revenue line.',
          view: 'Inserts a mirror module: one block for every module that sends the chosen link, added and removed with them.' },
        { key: 'mod-rename', label: 'Rename', glyph: '✎', phase: 1, modano: 'Modules > Rename',
          tip: 'Rename a module; its names, labels and contents entry follow.',
          view: 'Renames the module and re-prefixes its defined names, labels and contents entry.' },
        { key: 'mod-delete', label: 'Delete', glyph: '⊟', phase: 1, modano: 'Modules > Delete',
          tip: 'Remove a module, its rows, names, links and checks, after a warning about anything that depends on it.',
          view: 'Lists every row, name, link and check that would go and every module that depends on it, then removes them and rewires the totals.' },
        { key: 'mod-links', label: 'Links', glyph: '∞', phase: 1, modano: 'Right-click Link In and Link Out; Module Links',
          tip: 'The Explorer on its Links tab: what the module takes from and sends to, inputs still waiting for a sender, and link or unlink.',
          view: 'Opens the Explorer on the Links tab for the module at the cursor: what it takes from and sends to, named, inputs still waiting for a sender with the candidates, and outputs nothing takes. Link and unlink from here.' }
      ]
    },
    {
      id: 'categories', label: 'Categories', modano: 'Categories', controls: [
        { key: 'cat-add', ctxLabel: 'Add category', label: 'Add', glyph: '⊕', phase: 1, modano: 'Categories > Insert',
          tip: 'Add a category (a site, a revenue line, a facility) at the end of the block; totals, summaries and checks extend.',
          view: 'Adds a category below the last one. Every module that collects it (summaries, statements, working capital, lookups, charts) grows with it.' },
        { key: 'cat-above', ctxLabel: 'Add category above', label: 'Add above', glyph: '⇧', phase: 1, modano: 'Categories > Add Above',
          tip: 'Add a category above the selected one.',
          view: 'Adds a category above the selected category and renumbers the ones below.' },
        { key: 'cat-many', label: 'Add many', glyph: '⋮', phase: 1, modano: 'Categories > Add Multiple',
          tip: 'Add several categories at once from a pasted list.',
          view: 'Paste or type a list of names; each becomes a category in one step.' },
        { key: 'cat-remove', ctxLabel: 'Remove category', label: 'Remove', glyph: '⊖', phase: 1, modano: 'Categories > Delete',
          tip: 'Remove a category; totals and summaries close up.',
          view: 'Removes the category everywhere it appears, after a preview.' },
        { key: 'cat-subtotal', ctxLabel: 'Insert subtotal', label: 'Subtotal', glyph: 'Σ', phase: 2, modano: 'Right-click Insert Subtotal',
          tip: 'Group categories under a subtotal inside the block.',
          view: 'Adds a subtotal row over the selected categories and keeps the grand total right.' }
      ]
    },
    {
      id: 'data', label: 'Data', modano: 'Get Data', controls: [
        { key: 'data-pull', label: 'Pull actuals', glyph: '⇩', phase: 2, modano: 'Get Data (Xero and other ledgers)',
          tip: 'Pull actuals from Home Hub by group account and month, mapped and checked.',
          view: 'Signs in with Entra, pulls hh_in_actuals for the entity and months, maps them through the group chart and writes them to the actual months. Shows the last pull time and the tie-out checks.' },
        { key: 'data-push', label: 'Push', glyph: '⇧', phase: 2, modano: 'New',
          tip: 'Send results to Home Hub Planning: to plan lines this model is the source for, or as a proposed scenario.',
          view: 'Previews what goes to Home Hub under the push rule: lines this model owns are written; anything else goes as a proposed scenario for someone to accept.' },
        { key: 'data-import', label: 'Import', glyph: '⇲', phase: 2, modano: 'Get Data > Import Assumptions',
          tip: 'Import assumptions from a file or another sheet in the hh_in shape.',
          view: 'Choose a file or sheet laid out in the hh_in shape; the import maps it to input rows and shows what changed.' },
        { key: 'data-map', label: 'Map accounts', glyph: '⇆', phase: 2, modano: 'Mapping files',
          tip: 'Map ledger accounts to model lines.',
          view: 'The mapping table between group accounts and model lines, with unmapped accounts flagged.' }
      ]
    },
    {
      id: 'timeline', label: 'Timeline', modano: 'Time Series', controls: [
        { key: 'time-extend', label: 'Extend', glyph: '⟷', phase: 1, modano: 'Time Series > Extend or Shorten',
          tip: 'Extend or shorten the timeline, with a preview of every sheet, column, name and check it touches.',
          view: 'Set the new end month. The preview lists each sheet and the columns, names, charts and checks that change.' },
        { key: 'time-roll', label: 'Roll forward', glyph: '⇥', phase: 2, modano: 'Model roll forward',
          tip: 'Month-end roll forward: pull actuals, move the last actual month and keep inputs on their dates.',
          view: 'Moves the last actual month on by one, pulls that month\'s actuals and keeps every input on its date.' },
        { key: 'time-settings', label: 'Time settings', glyph: '◷', phase: 1, modano: 'Time sheet controls',
          tip: 'Model start, term, financial year end, last actual month, first budget month and denomination.',
          view: 'The Time sheet settings and their drop-downs (DD_Ts_ names): model start month, term, financial year end, last actual month, first budget month, denomination.' }
      ]
    },
    {
      id: 'analysis', label: 'Analysis', modano: 'Scenario manager (model content)', controls: [
        { key: 'an-scenarios', label: 'Scenarios', glyph: '◫', phase: 2, modano: 'Scenario manager',
          tip: 'Pick the active scenario, compare scenarios side by side and run every scenario at once.',
          view: 'The scenario manager: active scenario, per-block overrides, scenario names, and a run that stores each scenario\'s results with the inputs that made them.' },
        { key: 'an-sensitivity', label: 'Sensitivity', glyph: '⚖', phase: 2, modano: 'Sensitivity methods',
          tip: 'Flex chosen inputs by a step, a percentage or a low-high range and record the outputs, with a tornado.',
          view: 'Choose inputs and a method (step, percentage, low-high). The run records each output and draws a tornado.' },
        { key: 'an-simulation', label: 'Simulation', glyph: '⚄', phase: 2, modano: 'New (Monte Carlo scenario in the EXL demos)',
          tip: 'Run the Monte Carlo: trials, seed and sampling; time the run; inspect any trial.',
          view: 'Shows the trial count the model was built with, the seed and sampling, the estimated run time, then runs the data table and refreshes the distribution report. Inspect replays one trial through the model.' },
        { key: 'an-freeze', label: 'Freeze results', glyph: '❄', phase: 2, modano: 'New',
          tip: 'Store scenario or simulation results with their settings and an input fingerprint, so a stale result shows.',
          view: 'Copies the results to the store with seed, trials, versions, input fingerprint, time and user, and sets the stale flag check.' },
        { key: 'an-versions', label: 'Versions', glyph: '❒', phase: 1, modano: 'Budget module values and Budget Variance Analysis (reforecasting)', type: 'menu',
          tip: 'Save the budget or this month\'s reforecast as a version, approve and lock budgets, and compare the reports against any saved version.',
          view: 'Saved versions: the register of budgets and monthly reforecasts kept as values, and what the budget reports compare against.',
          items: [
            { key: 'an-ver-save', label: 'Save version', tip: 'Save the statements as values: this month\'s reforecast (as at the last actual month) or the budget being built.',
              view: 'Checks must be clear. Suggests the type and label (Reforecast Sep 2026 (6+6), Budget FY2027) and writes the values to the version store.' },
            { key: 'an-ver-approve', label: 'Approve budget', tip: 'Make a saved budget the approved budget for its year; the previous one becomes Superseded.',
              view: 'Sets the budget\'s status to Approved and locks it; the reports\' Budget choice reads it for that year.' },
            { key: 'an-ver-lock', label: 'Lock or unlock', tip: 'Locked versions cannot be replaced or deleted.',
              view: 'Locks or unlocks the selected version; unlocking an approved budget needs a reason, which is logged.' },
            { key: 'an-ver-compare', label: 'Compare versions', tip: 'Open the Version comparison: month, year to date and full year against two saved versions.',
              view: 'Goes to the Version comparison module (inserting it if the model has none) and sets the two comparisons.' },
            { key: 'an-ver-manage', label: 'Manage versions', tip: 'Rename, delete or push saved versions, and check their checksums.',
              view: 'The register of saved versions: rename, delete unlocked ones, push one to Home Hub Planning as a proposed version, or load the approved budget from Planning.' }
          ] },
        { key: 'an-impacts', label: 'Impacts', glyph: '⇶', phase: 2, modano: 'Financial Statement Impacts Analyser (a separate Modano model)', type: 'menu',
          tip: 'What a change or a transaction does to the income statement, balance sheet and cash flow, in this model\'s own lines.',
          view: 'Impacts: the effect of changing an input on every statement line, with the ties and the links that carried it, or Impacts sheets for each kind of transaction the model holds.',
          items: [
            { key: 'imp-live', label: 'Impact of a change', tip: 'Change one input, see every statement line that moves, the ties and why; the model is put back as it was.',
              view: 'Pick an input, a new value and a month. The add-in writes the value, recalculates, reads the statements and puts the value back in one step, then shows each line that moved, the ties and the chain of links.' },
            { key: 'imp-sheets', label: 'Impacts sheets', tip: 'Add an Impacts section: one formula sheet per kind of transaction this model holds, in its own accounts and entities.',
              view: 'Preview and add the Impacts sheets this model supports: entries, the effect on each statement by entity and, for intergroup items, the eliminations and the group, with switches and checks.' }
          ] }
      ]
    },
    {
      id: 'system', label: 'System', modano: 'System and Mode', controls: [
        { key: 'sys-library', label: 'Library', glyph: '▤', phase: 1, modano: 'System > Content Libraries',
          tip: 'Browse the module library: model types, modules, versions and what changed.',
          view: 'The library bundle in use, its version, the model types and modules in it, and release notes.' },
        { key: 'sys-settings', label: 'Settings', glyph: '⚙', phase: 2, modano: 'System > Options',
          tip: 'Your defaults and this model\'s settings: contents, links, number formats, names, timeline, checks, page setup, styles.',
          view: 'User defaults and model settings, as on Modano\'s options dialog.' },
        { key: 'sys-builder', label: 'Builder tools', glyph: '⚒', phase: 1, modano: 'Mode > Automate; Build tab',
          tip: 'Show or hide the HFG Build tab with the builder commands.',
          view: 'Shows the HFG Build tab (structure, styles, content, charts, review, finish). In the probe this also tests contextual tabs.' },
        { key: 'sys-help', label: 'Help', glyph: '?', phase: 1, modano: 'System > Help',
          tip: 'Help for the command or view in front of you.',
          view: 'Help for the current view, written with the command and released with it.' }
      ]
    }
  ];

  const BUILD = [
    {
      id: 'manage', label: 'Manage', modano: 'Build > Manage', controls: [
        { key: 'b-explorer', label: 'Explorer', glyph: '☰', phase: 1, modano: 'Build > Project Manager',
          tip: 'Model explorer with builder detail: areas, components, link records and metadata.',
          view: 'The explorer with builder detail: areas and their order, each module\'s components and sheets, its link records, and the metadata drift check.' },
        { key: 'b-linkmap', label: 'Link map', glyph: '⇢', phase: 2, modano: 'Module Structure > Module Links',
          tip: 'A map of every link between modules, with unmet and unused links highlighted.',
          view: 'Draws the modules and the links between them, as in the spec\'s assembly diagram, with gaps highlighted.' }
      ]
    },
    {
      id: 'structure', label: 'Structure', modano: 'Module Structure', controls: [
        { key: 'b-sheet', label: 'Insert sheet', glyph: '▭', phase: 1, modano: 'Insert base sheet',
          tip: 'Insert a sheet by type and section, with header rows, timeline, navigation links and a contents entry.',
          view: 'Choose the sheet type (timeline, non-timeline, report, presentation) and section; the sheet arrives in the frame with its contents entry.' },
        { key: 'b-section', label: 'Section', glyph: '§', phase: 1, modano: 'Module Structure > Section',
          tip: 'Add a section bar or sub-heading in the frame styles.',
          view: 'Adds a section bar (Heading 1) or sub-heading (Heading 2) with its outline group.' },
        { key: 'b-rows', label: 'Insert rows', glyph: '⇟', phase: 1, modano: 'Module Structure > Cell Blocks',
          tip: 'Insert rows inside a block and keep its formulas, names and totals.',
          view: 'Inserts rows inside the selected block, copies the row formulas and widens totals.' },
        { key: 'b-catblock', label: 'Category block', glyph: '▥', phase: 2, modano: 'Module Structure > Category Blocks',
          tip: 'Make a block repeat for every category in a set.',
          view: 'Turns the selected rows into a category block that repeats for each member of a category set.' },
        { key: 'b-names', label: 'Names', glyph: '#', phase: 2, modano: 'Tools > Names',
          tip: 'Name a range by convention or from its row label, add named formulas, list and locate names.',
          view: 'The names panel: name by convention (module prefix plus item), named formulas, the full list with where each is used, and broken names.' }
      ]
    },
    {
      id: 'styles', label: 'Styles', modano: 'Styles & Formats', controls: [
        { key: 'b-style', label: 'Style', glyph: '¶', phase: 1, modano: 'Styles & Formats > Assumptions, Outputs', type: 'menu',
          tip: 'Apply an HFG style by purpose.',
          view: 'Applies a named workbook style by purpose.',
          items: [
            { key: 'b-style-input', label: 'Input', tip: 'Input style: light fill, thin border, blue text.', view: 'Applies the Input style to the selection.' },
            { key: 'b-style-link', label: 'Link', tip: 'Link style for values brought from another sheet.', view: 'Applies the Link style.' },
            { key: 'b-style-calc', label: 'Calculation', tip: 'Calculation style.', view: 'Applies the Calculation style.' },
            { key: 'b-style-total', label: 'Total', tip: 'Total style: bold with a top border.', view: 'Applies the Total style.' },
            { key: 'b-style-check', label: 'Check', tip: 'Check style.', view: 'Applies the Check style.' },
            { key: 'b-style-heading', label: 'Heading', tip: 'Section bar or sub-heading.', view: 'Applies Heading 1 or Heading 2.' }
          ] },
        { key: 'b-present', label: 'Presentation', glyph: '▣', phase: 1, modano: 'Styles & Formats > Presentations',
          tip: 'Switch a sheet to the presentation style set for reports.',
          view: 'Sets the presentation flag on the sheet and applies the output styles.' },
        { key: 'b-format', label: 'Number format', glyph: '±', phase: 1, modano: 'Styles & Formats > Decimal', type: 'menu',
          tip: 'Decimals, thousands, millions, percentages and multiples.',
          view: 'Number format commands.',
          items: [
            { key: 'b-fmt-more', label: 'More decimals', tip: 'Add a decimal place.', view: 'Adds a decimal place.' },
            { key: 'b-fmt-less', label: 'Fewer decimals', tip: 'Remove a decimal place.', view: 'Removes a decimal place.' },
            { key: 'b-fmt-thousands', label: 'Thousands', tip: 'Show in thousands.', view: 'Shows the selection in thousands.' },
            { key: 'b-fmt-millions', label: 'Millions', tip: 'Show in millions.', view: 'Shows the selection in millions.' },
            { key: 'b-fmt-percent', label: 'Percentage', tip: 'Percentage format.', view: 'Applies the percentage format.' }
          ] },
        { key: 'b-colour', label: 'Colour by content', glyph: '◐', phase: 2, modano: 'Content-based font colour',
          tip: 'Colour fonts by content: inputs, formulas, links and unique formulas.',
          view: 'Colours each cell by what it holds and shows the key.' }
      ]
    },
    {
      id: 'content', label: 'Content', modano: 'Content (Time Series, Forms, Checks)', controls: [
        { key: 'b-timeline', label: 'Timeline block', glyph: '◷', phase: 1, modano: 'Content > Time Series',
          tip: 'Add or remove the timeline block on a sheet.',
          view: 'Adds the timeline block (rows 5 to 15) and timeline columns to a sheet, or removes them.' },
        { key: 'b-controls', label: 'Controls', glyph: '☑', phase: 1, modano: 'Content > Forms', type: 'menu',
          tip: 'Add a drop-down, checkbox or other control linked to a named cell, with its list on Lookups.',
          view: 'Adds a control linked to a DD_ or CB_ name, with its list on the Lookups sheet.',
          items: [
            { key: 'b-ctl-list', label: 'Drop-down list', tip: 'In-cell drop-down from a Lookups list, linked to a DD_ name.', view: 'Adds an in-cell drop-down fed by an LU_ list and names the cell DD_.' },
            { key: 'b-ctl-check', label: 'Checkbox', tip: 'In-cell checkbox linked to a CB_ name.', view: 'Adds an in-cell checkbox and names the cell CB_.' },
            { key: 'b-ctl-option', label: 'Option list', tip: 'Choose one of several options, stored as a number.', view: 'Adds a drop-down that stores the chosen option\'s number, as an option button group would.' },
            { key: 'b-ctl-spin', label: 'Number input', tip: 'Whole-number input with limits and a step, in place of a spin button.', view: 'Adds a validated whole-number input with a minimum, maximum and step.' },
            { key: 'b-ctl-form', label: 'Form control (rebuild)', tip: 'Classic combo box or check box, written by the package writer when the model is rebuilt.', view: 'Queues a classic form control (combo box or check box) for the next package build; Office.js cannot create one in an open workbook.' }
          ] },
        { key: 'b-check', ctxLabel: 'Add check', label: 'Add check', glyph: '✔', phase: 1, modano: 'Content > Checks',
          tip: 'Add an error, alert or sensitivity check on the selected row; it rolls up to Checks.',
          view: 'Adds a check row, its include toggle and its entry on the Checks sheet.' },
        { key: 'b-lookup', label: 'Lookup list', glyph: '▾', phase: 1, modano: 'Lookup tables in modules',
          tip: 'Create a list on Lookups that grows with a category set, for drop-downs to use.',
          view: 'Creates an LU_ list on the Lookups sheet. A list built from a category set grows and shrinks with it.' },
        { key: 'b-driver', label: 'Driver method', glyph: '⇄', phase: 2, modano: 'Modules > Replace (driver variant)',
          tip: 'Change how a line is driven: amount, growth, price times volume, share of another line, days.',
          view: 'Shows the driver methods the module allows for the selected category and switches it, writing the new driver rows.' }
      ]
    },
    {
      id: 'charts', label: 'Charts', modano: 'New (Modano has category charts and dashboards)', controls: [
        { key: 'ch-ibcs', label: 'IBCS column', glyph: '▮', phase: 2, modano: 'New',
          tip: 'Actual, prior year, plan and forecast in IBCS notation: solid, grey, outlined and hatched.',
          view: 'Inserts an IBCS column chart from the selected rows: actual solid, prior year grey, plan outlined, forecast hatched, with a scale shared across charts.' },
        { key: 'ch-variance', label: 'Variance', glyph: '±', phase: 2, modano: 'New',
          tip: 'Absolute and relative variance bars, green for good and red for bad, by line type.',
          view: 'Inserts IBCS variance charts (absolute and percentage) against plan or prior year; colours follow whether the line is income or cost.' },
        { key: 'ch-waterfall', label: 'Waterfall', glyph: '▙', phase: 2, modano: 'New',
          tip: 'A bridge from one total to another: P&L walk, budget to actual, cash bridge.',
          view: 'Inserts a waterfall from the selected lines with totals marked, increases green, decreases red and connector lines.' },
        { key: 'ch-z', label: 'Z chart', glyph: 'Z', phase: 2, modano: 'New',
          tip: 'Monthly actual and forecast bars, year to date and moving annual total lines, against budget.',
          view: 'Inserts a Z chart for a line: monthly bars (actual solid, forecast hatched, budget grey), cumulative actual then forecast, cumulative budget, moving annual total, and the year to date and full year variances.' },
        { key: 'ch-dist', label: 'Distribution', glyph: '∩', phase: 2, modano: 'New', type: 'menu',
          tip: 'Charts for simulation and sensitivity results.',
          view: 'Simulation and sensitivity charts.',
          items: [
            { key: 'ch-hist', label: 'Histogram', tip: 'Distribution of a result across trials, with percentiles marked.', view: 'Inserts a histogram with P10, P50 and P90 marked.' },
            { key: 'ch-scurve', label: 'S-curve', tip: 'Cumulative probability of a result.', view: 'Inserts a cumulative probability curve.' },
            { key: 'ch-fan', label: 'Fan chart', tip: 'Percentile bands of a balance by month.', view: 'Inserts a fan chart of P10, P50 and P90 by month.' },
            { key: 'ch-tornado', label: 'Tornado', tip: 'Which inputs move a result most.', view: 'Inserts a tornado from the sensitivity run.' }
          ] },
        { key: 'ch-more', label: 'More charts', glyph: '▦', phase: 2, modano: 'Category charts, dashboard templates', type: 'menu',
          tip: 'Standard charts and dashboard pieces.',
          view: 'Standard charts.',
          items: [
            { key: 'ch-line', label: 'Line', tip: 'Line chart from the selected rows.', view: 'Inserts a line chart.' },
            { key: 'ch-combo', label: 'Column and line', tip: 'Columns with a line on a second axis.', view: 'Inserts a combo chart.' },
            { key: 'ch-kpi', label: 'KPI tiles', tip: 'Tiles for key results with variance.', view: 'Inserts KPI tiles linked to named results.' },
            { key: 'ch-spark', label: 'Sparklines', tip: 'Small trend lines beside each row.', view: 'Adds sparklines next to the rows.' }
          ] },
        { key: 'ch-refresh', label: 'Refresh charts', glyph: '↻', phase: 2, modano: 'Category charts re-pointed',
          tip: 'Re-point charts after categories or the timeline change.',
          view: 'Rebinds every chart to its rows by key after categories or the timeline change, and lists any chart it could not fix.' }
      ]
    },
    {
      id: 'review', label: 'Review', modano: 'Finalization > Auditing; WIP', controls: [
        { key: 'r-trace', label: 'Trace', glyph: '↯', phase: 2, modano: 'Traverse Formula',
          tip: 'Precedents and dependents with labels and values; click to go and back to start.',
          view: 'The trace pane for the selected cell.' },
        { key: 'r-errors', label: 'Errors', glyph: '⚠', phase: 2, modano: 'Locate errors',
          tip: 'Find formula errors by sheet, with the root errors marked.',
          view: 'Lists formula errors by sheet and marks the ones that cause the rest.' },
        { key: 'r-scan', label: 'Consistency', glyph: '≡', phase: 3, modano: 'Shade unique formulas',
          tip: 'Find formulas that break the row pattern and numbers typed into formula rows.',
          view: 'Scans each row for unique formulas and typed values and shades them.' },
        { key: 'r-inspect', label: 'Inspector', glyph: '⊙', phase: 2, modano: 'Audit cell content',
          tip: 'What the selected cell is: module, row, unit, style, name and flags.',
          view: 'The cell inspector.' },
        { key: 'r-wip', ctxLabel: 'Mark work in progress', label: 'Work in progress', glyph: '✱', phase: 3, modano: 'WIP > Apply',
          tip: 'Mark cells as work in progress; finalise is blocked until they are cleared.',
          view: 'Marks or clears work in progress and shows the register.' },
        { key: 'r-notes', label: 'Review notes', glyph: '✉', phase: 3, modano: 'Finalization > Issues Tracking',
          tip: 'Review notes tied to cells, with owner and status.',
          view: 'The review notes register.' }
      ]
    },
    {
      id: 'finish', label: 'Finish', modano: 'Finalization', controls: [
        { key: 'f-tidy', label: 'Tidy view', glyph: '✧', phase: 2, modano: 'Finalization > Workbook View',
          tip: 'A1 selected, panes frozen, gridlines off and outline set on every sheet.',
          view: 'Tidies every sheet for handover.' },
        { key: 'f-page', label: 'Page setup', glyph: '▤', phase: 2, modano: 'Page setups',
          tip: 'Page setup from the frame: orientation, fit to width, print titles, footer.',
          view: 'Applies the frame page setup.' },
        { key: 'f-pack', label: 'Report pack', glyph: '⇲', phase: 3, modano: 'Print workbook from report definitions',
          tip: 'Export a named set of report sheets with cover and contents to PDF.',
          view: 'Builds the report pack PDF.' },
        { key: 'f-final', label: 'Finalise', glyph: '✓', phase: 3, modano: 'Validation and finalisation',
          tip: 'Checks clear, no work in progress, no errors, tidy view, protection on, version stamped.',
          view: 'Runs the finalise checklist.' },
        { key: 'f-release', label: 'Release copy', glyph: '⇪', phase: 3, modano: 'Unmanage workbook',
          tip: 'A copy for outside recipients with the metadata removed.',
          view: 'Saves a release copy without metadata.' }
      ]
    }
  ];

  // Right-click cell menu: one HFG submenu (Office allows one level).
  const CONTEXT = ['cat-add', 'cat-above', 'cat-remove', 'cat-subtotal', 'ctx-linkin', 'ctx-linkout', 'r-trace',
                   'ctx-impacts', 'b-check', 'r-wip', 'ctx-stats', 'model-explorer', 'sys-help'];
  const CONTEXT_ONLY = [
    { key: 'ctx-linkin', label: 'Link in', glyph: '⇢', phase: 1, modano: 'Right-click Link In',
      tip: 'Link the selected row to a sender.', view: 'Choose a sender for the selected input row.' },
    { key: 'ctx-linkout', label: 'Link out', glyph: '⇢', phase: 1, modano: 'Right-click Link Out',
      tip: 'Send the selected row as a link item.', view: 'Choose the link item the selected row sends.' },
    { key: 'ctx-impacts', label: 'Show impacts', glyph: '⇶', phase: 2, modano: 'New (Modano\'s analyser is a separate model)',
      tip: 'The Impacts pane for the input under the cursor: what changing it does to the statements.', view: 'Opens Impact of a change with the selected input.' },
    { key: 'ctx-stats', label: 'Range stats', glyph: 'Σ', phase: 2, modano: 'Right-click Range Stats',
      tip: 'Sum, average, minimum, maximum and count of the selection.', view: 'Range statistics for the selection.' }
  ];

  const SHORTCUTS = [
    { action: 'HFG_EXPLORER', key: 'model-explorer', name: 'Open the model explorer', win: 'Ctrl+Alt+Shift+E', mac: 'Command+Alt+Shift+E' },
    { action: 'HFG_INSERT', key: 'mod-insert', name: 'Insert a module', win: 'Ctrl+Alt+Shift+I', mac: 'Command+Alt+Shift+I' },
    { action: 'HFG_ADDCAT', key: 'cat-add', name: 'Add a category', win: 'Ctrl+Alt+Shift+A', mac: 'Command+Alt+Shift+A' },
    { action: 'HFG_TRACE', key: 'r-trace', name: 'Trace the selected cell', win: 'Ctrl+Alt+Shift+T', mac: 'Command+Alt+Shift+T' },
    { action: 'HFG_CHECKS', key: 'model-checks', name: 'Open the checks panel', win: 'Ctrl+Alt+Shift+K', mac: 'Command+Alt+Shift+K' },
    { action: 'HFG_GROUP', key: 'grp-structure', name: 'Open the group structure', win: 'Ctrl+Alt+Shift+G', mac: 'Command+Alt+Shift+G' }
  ];

  function all() {
    const out = [];
    for (const [tab, groups] of [['main', MAIN], ['build', BUILD]]) {
      for (const g of groups) {
        for (const c of g.controls) {
          out.push(Object.assign({ tab, group: g.label }, c));
          (c.items || []).forEach(i => out.push(Object.assign({ tab, group: g.label, parent: c.key, glyph: c.glyph, phase: c.phase, modano: c.modano }, i)));
        }
      }
    }
    CONTEXT_ONLY.forEach(c => out.push(Object.assign({ tab: 'context', group: 'Right-click' }, c)));
    return out;
  }

  function find(key) { return all().find(c => c.key === key) || null; }

  const api = {
    MAIN, BUILD, CONTEXT, CONTEXT_ONLY, SHORTCUTS, all, find,
    MAIN_TAB: { id: 'HFG.Probe.Tab', label: 'HFG Model' },
    BUILD_TAB: { id: 'HFG.Build.Tab', label: 'HFG Build' },
    controlId: key => 'HFG.' + key
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.HfgCommands = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
