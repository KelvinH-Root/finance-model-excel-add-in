// Writes manifest.xml and shortcuts.json from the command registry (src/commands.js).
//
//   node tools/build-ribbon.mjs
//
// The XML manifest allows one custom tab, so the main tab (HFG Model) is defined here and the
// Build tab is created at runtime as a contextual tab (see buildTabDefinition in probe.js).
// The probe's own test controls keep their ids so the existing probes still work.

import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
await import(path.join(ROOT, 'src/commands.js'));
const C = globalThis.HfgCommands;
const BASE = 'https://localhost:3000';

const shortStrings = new Map();
const longStrings = new Map();
const images = new Set();
const problems = [];

const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function shortRes(id, text) {
  if (id.length > 32) problems.push(`resource id over 32 characters: ${id}`);
  if (text.length > 125) problems.push(`short string over 125 characters: ${id}`);
  shortStrings.set(id, text);
  return id;
}
function longRes(id, text) {
  if (id.length > 32) problems.push(`resource id over 32 characters: ${id}`);
  if (text.length > 250) problems.push(`supertip over 250 characters: ${id} (${text.length})`);
  longStrings.set(id, text);
  return id;
}
function icon(key, pad) {
  images.add(key);
  return [`${pad}<Icon>`,
    ...[16, 32, 80].map(s => `${pad}  <bt:Image size="${s}" resid="I${s}.${key}"/>`),
    `${pad}</Icon>`].join('\n');
}
function supertip(c, pad, prefix = '') {
  return [`${pad}<Supertip>`,
    `${pad}  <Title resid="${shortRes(prefix + 'L.' + c.key, c.label)}"/>`,
    `${pad}  <Description resid="${longRes(prefix + 'T.' + c.key, c.tip)}"/>`,
    `${pad}</Supertip>`].join('\n');
}
function action(pad) {
  return [`${pad}<Action xsi:type="ExecuteFunction">`, `${pad}  <FunctionName>openView</FunctionName>`, `${pad}</Action>`].join('\n');
}
function item(c, idPrefix, iconKey, pad, context = false) {
  const label = context && c.ctxLabel ? shortRes('CL.' + c.key, c.ctxLabel) : shortRes('L.' + c.key, c.label);
  return [`${pad}<Item id="${idPrefix}${c.key}">`,
    `${pad}  <Label resid="${label}"/>`,
    supertip(c, pad + '  '),
    icon(iconKey, pad + '  '),
    action(pad + '  '),
    `${pad}</Item>`].join('\n');
}
function control(c, pad) {
  if (c.type === 'menu') {
    return [`${pad}<Control xsi:type="Menu" id="${C.controlId(c.key)}">`,
      `${pad}  <Label resid="${shortRes('L.' + c.key, c.label)}"/>`,
      supertip(c, pad + '  '),
      icon(c.key, pad + '  '),
      `${pad}  <Items>`,
      ...c.items.map(i => item(i, 'HFG.', c.key, pad + '    ')),
      `${pad}  </Items>`,
      `${pad}</Control>`].join('\n');
  }
  return [`${pad}<Control xsi:type="Button" id="${C.controlId(c.key)}">`,
    `${pad}  <Label resid="${shortRes('L.' + c.key, c.label)}"/>`,
    supertip(c, pad + '  '),
    icon(c.key, pad + '  '),
    action(pad + '  '),
    `${pad}</Control>`].join('\n');
}
function group(g, pad) {
  return [`${pad}<Group id="HFG.G.${g.id}">`,
    `${pad}  <Label resid="${shortRes('G.' + g.id, g.label)}"/>`,
    icon(g.controls[0].key, pad + '  '),
    ...g.controls.map(c => control(c, pad + '  ')),
    `${pad}</Group>`].join('\n');
}

// The probe's own test group, unchanged so the ribbon, menu and command probes still work.
const PROBE_GROUP = `              <Group id="HFG.Probe.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="HFG.Probe.Open">
                  <Label resid="Open.Label"/>
                  <Supertip>
                    <Title resid="Open.Label"/>
                    <Description resid="Open.Tip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
                <Control xsi:type="Button" id="HFG.Probe.Mark">
                  <Label resid="Mark.Label"/>
                  <Supertip>
                    <Title resid="Mark.Label"/>
                    <Description resid="Mark.Tip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>ribbonMark</FunctionName>
                  </Action>
                </Control>
                <Control xsi:type="Button" id="HFG.Probe.Toggle">
                  <Label resid="Toggle.Label"/>
                  <Supertip>
                    <Title resid="Toggle.Label"/>
                    <Description resid="Toggle.Tip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>ribbonToggle</FunctionName>
                  </Action>
                </Control>
                <Control xsi:type="Menu" id="HFG.Probe.Menu">
                  <Label resid="Menu.Label"/>
                  <Supertip>
                    <Title resid="Menu.Label"/>
                    <Description resid="Menu.Tip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Items>
                    <Item id="HFG.Probe.Menu.Mark">
                      <Label resid="MenuMark.Label"/>
                      <Supertip>
                        <Title resid="MenuMark.Label"/>
                        <Description resid="Mark.Tip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="Icon.16"/>
                        <bt:Image size="32" resid="Icon.32"/>
                        <bt:Image size="80" resid="Icon.80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>menuMark</FunctionName>
                      </Action>
                    </Item>
                  </Items>
                </Control>
              </Group>`;

const PROBE_CTX_ITEM = `                  <Item id="HFG.Probe.Ctx.Mark">
                    <Label resid="CtxMark.Label"/>
                    <Supertip>
                      <Title resid="CtxMark.Label"/>
                      <Description resid="Ctx.Tip"/>
                    </Supertip>
                    <Icon>
                      <bt:Image size="16" resid="Icon.16"/>
                      <bt:Image size="32" resid="Icon.32"/>
                      <bt:Image size="80" resid="Icon.80"/>
                    </Icon>
                    <Action xsi:type="ExecuteFunction">
                      <FunctionName>contextMark</FunctionName>
                    </Action>
                  </Item>`;

const mainGroups = C.MAIN.map(g => group(g, '              ')).join('\n');
const contextItems = C.CONTEXT.map(k => {
  const c = C.find(k);
  if (!c) throw new Error(`context command ${k} is not in the registry`);
  return item(c, 'HFG.ctx.', c.key, '                  ', true);
}).join('\n');

const probeShort = [
  ['GetStarted.Title', 'HFG Probe is loaded'], ['Tab.Label', C.MAIN_TAB.label], ['Group.Label', 'Probe'],
  ['Open.Label', 'Open probe'], ['Mark.Label', 'Mark cell'], ['Toggle.Label', 'Toggle test'], ['Menu.Label', 'Menu test'],
  ['MenuMark.Label', 'Mark from menu'], ['Ctx.Label', 'HFG'], ['CtxMark.Label', 'Mark this cell (probe)']
];
const probeLong = [
  ['GetStarted.Description', 'Open the HFG Model tab and choose Open probe.'],
  ['Open.Tip', 'Opens the probe task pane.'],
  ['Mark.Tip', 'Writes a timestamp next to the selected cell without opening the task pane.'],
  ['Toggle.Tip', 'The probe enables and disables this button to test ribbon updates.'],
  ['Menu.Tip', 'Tests a menu on the ribbon.'],
  ['Ctx.Tip', 'HFG commands for the selected cell.']
];

const imageRes = [...images].sort().flatMap(k => [16, 32, 80].map(s =>
  `        <bt:Image id="I${s}.${k}" DefaultValue="${BASE}/assets/cmd/${k}-${s}.png"/>`)).join('\n');
const shortRes_ = [...probeShort, ...shortStrings].map(([k, v]) => `        <bt:String id="${k}" DefaultValue="${esc(v)}"/>`).join('\n');
const longRes_ = [...probeLong, ...longStrings].map(([k, v]) => `        <bt:String id="${k}" DefaultValue="${esc(v)}"/>`).join('\n');

const manifest = `<?xml version="1.0" encoding="UTF-8"?>
<!--
  HFG add-in probe: a throwaway add-in that tests what Office.js can do in real Excel
  before the HFG add-in is built. Served from https://localhost:3000 (see README.md).

  GENERATED by tools/build-ribbon.mjs from src/commands.js. Edit the registry, not this file.
  The HFG Model tab and the right-click menu show the designed commands; each opens a
  placeholder view. The Build tab is created at runtime (one custom tab per add-in).
-->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>67e07eed-1d14-413c-81da-c08e5aa30b22</Id>
  <Version>1.1.0.0</Version>
  <ProviderName>Home Foundation Group</ProviderName>
  <DefaultLocale>en-NZ</DefaultLocale>
  <DisplayName DefaultValue="HFG Probe"/>
  <Description DefaultValue="Tests what Office.js can do in this copy of Excel, and shows the designed HFG ribbon."/>
  <IconUrl DefaultValue="${BASE}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${BASE}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="${BASE}/src/help.html"/>
  <AppDomains>
    <AppDomain>${BASE}</AppDomain>
    <AppDomain>https://login.microsoftonline.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${BASE}/src/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long"/>
        </Runtimes>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="Readme.Url"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="${C.MAIN_TAB.id}">
${mainGroups}
${PROBE_GROUP}
              <Label resid="Tab.Label"/>
            </CustomTab>
          </ExtensionPoint>
          <ExtensionPoint xsi:type="ContextMenu">
            <OfficeMenu id="ContextMenuCell">
              <Control xsi:type="Menu" id="HFG.Probe.Ctx">
                <Label resid="Ctx.Label"/>
                <Supertip>
                  <Title resid="Ctx.Label"/>
                  <Description resid="Ctx.Tip"/>
                </Supertip>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Items>
${contextItems}
${PROBE_CTX_ITEM}
                </Items>
              </Control>
            </OfficeMenu>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${BASE}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="${BASE}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="${BASE}/assets/icon-80.png"/>
${imageRes}
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane.Url" DefaultValue="${BASE}/src/taskpane.html"/>
        <bt:Url id="Readme.Url" DefaultValue="${BASE}/src/help.html"/>
      </bt:Urls>
      <bt:ShortStrings>
${shortRes_}
      </bt:ShortStrings>
      <bt:LongStrings>
${longRes_}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
  <ExtendedOverrides Url="${BASE}/shortcuts.json"/>
</OfficeApp>
`;

const shortcuts = {
  actions: [
    { id: 'SHORTCUTMARK', type: 'ExecuteFunction', name: 'Mark cell (probe shortcut)' },
    ...C.SHORTCUTS.map(s => ({ id: s.action, type: 'ExecuteFunction', name: s.name }))
  ],
  shortcuts: [
    { action: 'SHORTCUTMARK', key: { default: 'Ctrl+Alt+Shift+J', mac: 'Command+Alt+Shift+J' } },
    ...C.SHORTCUTS.map(s => ({ action: s.action, key: { default: s.win, mac: s.mac } }))
  ]
};

if (problems.length) {
  console.error(problems.join('\n'));
  process.exit(1);
}
writeFileSync(path.join(ROOT, 'manifest.xml'), manifest);
writeFileSync(path.join(ROOT, 'shortcuts.json'), JSON.stringify(shortcuts, null, 2) + '\n');
const buttons = C.MAIN.reduce((n, g) => n + g.controls.length, 0);
console.log(`manifest.xml: ${C.MAIN.length} groups and ${buttons} controls on ${C.MAIN_TAB.label}, plus the probe group; ` +
  `${C.CONTEXT.length} right-click items; ${images.size} icons. shortcuts.json: ${shortcuts.shortcuts.length} shortcuts.`);
