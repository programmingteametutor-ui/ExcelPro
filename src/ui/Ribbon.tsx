import { useState, type ReactNode } from 'react'
import type { CellStyle, CFRule, TotalFn } from '../model/types'
import type { RibbonContext } from './useFeatures'
import { TABLE_STYLES } from '../engine/features/tables'
import { CHART_TYPES } from './charts'
import { FORMAT_PRESETS } from '../engine/numberFormat'
import type { Commands, ViewState } from './commands'
import { Check, ColorGrid, MenuItem, MenuLabel, MenuSep, Popover, ToolButton } from './components/ui'
import { Icon } from './icons'
import { cn, shortcut } from './lib/utils'

const TABS = ['Home', 'Insert', 'Page Layout', 'Formulas', 'Data', 'Review', 'View', 'Automate', 'Help'] as const
type Tab = (typeof TABS)[number] | 'Table Design'

const FN_CATEGORIES: [string, string[]][] = [
  ['Financial', ['PMT', 'FV', 'PV', 'NPV', 'IRR', 'RATE', 'NPER', 'IPMT', 'PPMT', 'XNPV', 'XIRR', 'SLN', 'DB', 'DDB']],
  ['Logical', ['IF', 'IFS', 'AND', 'OR', 'NOT', 'XOR', 'IFERROR', 'IFNA', 'SWITCH', 'TRUE', 'FALSE']],
  ['Text', ['CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE', 'TRIM', 'CLEAN', 'UPPER', 'LOWER', 'PROPER', 'TEXT', 'VALUE', 'TEXTBEFORE', 'TEXTAFTER', 'REPT', 'EXACT']],
  ['Date & Time', ['TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'WEEKNUM', 'EDATE', 'EOMONTH', 'NETWORKDAYS', 'WORKDAY', 'DATEDIF', 'DAYS', 'DATEVALUE']],
  ['Lookup & Reference', ['XLOOKUP', 'VLOOKUP', 'HLOOKUP', 'LOOKUP', 'INDEX', 'MATCH', 'XMATCH', 'CHOOSE', 'OFFSET', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'TRANSPOSE']],
  ['Math & Trig', ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'MOD', 'ABS', 'POWER', 'SQRT', 'SEQUENCE', 'RANDARRAY', 'RAND', 'RANDBETWEEN', 'SUBTOTAL', 'AGGREGATE', 'PRODUCT', 'CEILING', 'FLOOR']],
  ['Statistical', ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'MAX', 'MIN', 'MAXIFS', 'MINIFS', 'MEDIAN', 'MODE', 'STDEV', 'VAR', 'LARGE', 'SMALL', 'RANK', 'PERCENTILE', 'QUARTILE', 'CORREL', 'FORECAST']],
]
const RECENT_KEY = 'excelpro.recentFns'
function recentFns(): string[] {
  try {
    return JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]') as string[]
  } catch {
    return []
  }
}
export function rememberFn(name: string) {
  try {
    localStorage.setItem(RECENT_KEY, JSON.stringify([name, ...recentFns().filter((n) => n !== name)].slice(0, 10)))
  } catch {
    /* storage unavailable */
  }
}

const FONTS = ['Calibri', 'Arial', 'Aptos', 'Cambria', 'Consolas', 'Courier New', 'Georgia', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana']
const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72]

function Group({ label, children, className }: { label: string; children: ReactNode; className?: string }) {
  return (
    <div className={cn('flex shrink-0 flex-col border-r border-[var(--border)] px-2 last:border-r-0', className)} role="group" aria-label={label}>
      <div className="flex flex-1 flex-wrap content-start items-center gap-0.5">{children}</div>
      <div className="mt-0.5 hidden text-center text-[10.5px] leading-3 text-[var(--muted)] md:block">{label}</div>
    </div>
  )
}

function Planned({ phase, items }: { phase: string; items: string }) {
  return (
    <div className="flex max-w-[520px] shrink-0 items-center gap-2 self-center rounded-md border border-dashed border-[var(--border-strong)] px-3 py-1.5 text-[11.5px] leading-4 text-[var(--muted)]">
      <Icon name="info" size={14} className="shrink-0" />
      <span>
        <b className="font-semibold">{phase}:</b> {items}
      </span>
    </div>
  )
}

function Dd({ label, tip, icon, children, wide }: { label?: string; tip: string; icon?: Parameters<typeof Icon>[0]['name']; children: ReactNode | ((close: () => void) => ReactNode); wide?: boolean }) {
  return (
    <Popover
      trigger={({ toggle, open }) => (
        <ToolButton tip={tip} icon={icon} label={label} onClick={toggle} active={open} wide={wide}>
          <Icon name="chevronDown" size={12} className="opacity-60" />
        </ToolButton>
      )}
    >
      {children}
    </Popover>
  )
}

const selectCls = 'h-7 rounded-[5px] border border-[var(--border)] bg-[var(--bg)] px-1.5 text-[12.5px] text-[var(--text)] outline-none hover:border-[var(--border-strong)] focus:border-[var(--accent)]'

export function Ribbon({
  cmd,
  style,
  view,
  setView,
  painterActive,
  hasClip,
  ctx,
}: {
  cmd: Commands
  style: CellStyle | undefined
  view: ViewState
  setView: (v: Partial<ViewState>) => void
  painterActive: boolean
  hasClip: boolean
  ctx: RibbonContext
}) {
  const [tabState, setTab] = useState<Tab>('Home')
  const tab: Tab = tabState === 'Table Design' && !ctx.table ? 'Home' : tabState
  const f = cmd.feature
  const ps = ctx.pageSetup
  const insertFn = (name: string) => {
    rememberFn(name)
    cmd.insertFunctionNamed(name)
  }
  const cf = (rule: CFRule) => f({ a: 'cfNew', rule })
  const s = style ?? {}
  const currentFmt = FORMAT_PRESETS.find((p) => p.fmt === s.numFmt)?.id ?? (s.numFmt ? 'custom' : 'general')

  return (
    <div className="shrink-0 border-b border-[var(--border)] bg-[var(--panel)]">
      <div role="tablist" className="xp-scroll-thin flex gap-0.5 overflow-x-auto px-2 pt-1">
        {[...TABS, ...(ctx.table ? (['Table Design'] as const) : [])].map((t) => (
          <button
            key={t}
            role="tab"
            aria-selected={tab === t}
            onClick={() => setTab(t)}
            className={cn(
              'relative h-7 shrink-0 rounded-t-md px-3 text-[12.5px] hover:bg-[var(--hover)]',
              tab === t && 'font-semibold text-[var(--accent)] after:absolute after:inset-x-2 after:bottom-0 after:h-[2px] after:rounded after:bg-[var(--accent)]',
              t === 'Table Design' && 'text-[var(--accent)]',
            )}
          >
            {t}
          </button>
        ))}
      </div>
      <div role="tabpanel" aria-label={tab} className="xp-scroll-thin flex min-h-[70px] items-stretch overflow-x-auto px-1 py-1.5">
        {tab === 'Home' && (
          <>
            <Group label="Clipboard">
              <div className="flex flex-col">
                <div className="flex">
                  <ToolButton tip={`Paste (${shortcut('Ctrl+V')})`} icon="paste" onClick={() => cmd.paste()} />
                  <Dd tip="Paste options">
                    <MenuLabel>Paste</MenuLabel>
                    <MenuItem icon="paste" onSelect={() => cmd.paste('all')} shortcut={shortcut('Ctrl+V')}>
                      Paste
                    </MenuItem>
                    <MenuItem onSelect={() => cmd.paste('values')} disabled={!hasClip}>
                      Paste values
                    </MenuItem>
                    <MenuItem onSelect={() => cmd.paste('formulas')} disabled={!hasClip}>
                      Paste formulas
                    </MenuItem>
                    <MenuItem onSelect={() => cmd.paste('formats')} disabled={!hasClip}>
                      Paste formatting
                    </MenuItem>
                  </Dd>
                </div>
                <div className="flex">
                  <ToolButton tip={`Cut (${shortcut('Ctrl+X')})`} icon="cut" onClick={cmd.cut} />
                  <ToolButton tip={`Copy (${shortcut('Ctrl+C')})`} icon="copy" onClick={cmd.copy} />
                  <ToolButton tip="Format Painter: copy formatting to the next selection" icon="brush" active={painterActive} onClick={cmd.formatPainter} />
                </div>
              </div>
            </Group>
            <Group label="Font" className="max-w-[260px]">
              <div className="flex gap-1">
                <select aria-label="Font" className={cn(selectCls, 'w-[118px]')} value={s.fontFamily ?? 'Calibri'} onChange={(e) => cmd.setStyle({ fontFamily: e.target.value === 'Calibri' ? undefined : e.target.value }, 'Font')}>
                  {[...new Set([...FONTS, s.fontFamily ?? 'Calibri'])].map((f) => (
                    <option key={f}>{f}</option>
                  ))}
                </select>
                <select aria-label="Font size" className={cn(selectCls, 'w-[54px]')} value={s.fontSize ?? 11} onChange={(e) => cmd.setStyle({ fontSize: +e.target.value === 11 ? undefined : +e.target.value }, 'Font size')}>
                  {[...new Set([...SIZES, s.fontSize ?? 11])].sort((a, b) => a - b).map((n) => (
                    <option key={n}>{n}</option>
                  ))}
                </select>
              </div>
              <div className="flex">
                <ToolButton tip={`Bold (${shortcut('Ctrl+B')})`} icon="bold" active={!!s.bold} onClick={() => cmd.toggle('bold')} />
                <ToolButton tip={`Italic (${shortcut('Ctrl+I')})`} icon="italic" active={!!s.italic} onClick={() => cmd.toggle('italic')} />
                <ToolButton tip={`Underline (${shortcut('Ctrl+U')})`} icon="underline" active={!!s.underline} onClick={() => cmd.toggle('underline')} />
                <ToolButton tip={`Strikethrough (${shortcut('Ctrl+5')})`} icon="strike" active={!!s.strike} onClick={() => cmd.toggle('strike')} />
                <Dd tip="Borders" icon="borders">
                  <MenuLabel>Borders</MenuLabel>
                  <MenuItem icon="borders" onSelect={() => cmd.borders('all')}>All borders</MenuItem>
                  <MenuItem icon="borderNone" onSelect={() => cmd.borders('outside')}>Outside borders</MenuItem>
                  <MenuItem icon="borderNone" onSelect={() => cmd.borders('thickOutside')}>Thick outside borders</MenuItem>
                  <MenuSep />
                  <MenuItem onSelect={() => cmd.borders('top')}>Top border</MenuItem>
                  <MenuItem onSelect={() => cmd.borders('bottom')}>Bottom border</MenuItem>
                  <MenuItem onSelect={() => cmd.borders('left')}>Left border</MenuItem>
                  <MenuItem onSelect={() => cmd.borders('right')}>Right border</MenuItem>
                  <MenuSep />
                  <MenuItem icon="x" onSelect={() => cmd.borders('none')}>No border</MenuItem>
                </Dd>
                <Popover
                  trigger={({ toggle, open }) => (
                    <ToolButton tip="Fill color" onClick={toggle} active={open}>
                      <span className="flex flex-col items-center">
                        <Icon name="fill" size={14} />
                        <span className="mt-[1px] h-[3px] w-4 rounded-sm" style={{ background: s.fill ?? '#ffe599' }} />
                      </span>
                    </ToolButton>
                  )}
                >
                  <ColorGrid noneLabel="No fill" onPick={(c) => cmd.setStyle({ fill: c }, 'Fill color')} />
                </Popover>
                <Popover
                  trigger={({ toggle, open }) => (
                    <ToolButton tip="Font color" onClick={toggle} active={open}>
                      <span className="flex flex-col items-center">
                        <Icon name="fontColor" size={14} />
                        <span className="mt-[1px] h-[3px] w-4 rounded-sm" style={{ background: s.color ?? '#c00000' }} />
                      </span>
                    </ToolButton>
                  )}
                >
                  <ColorGrid noneLabel="Automatic" onPick={(c) => cmd.setStyle({ color: c }, 'Font color')} />
                </Popover>
              </div>
            </Group>
            <Group label="Alignment" className="max-w-[200px]">
              <div className="flex">
                <ToolButton tip="Top align" icon="alignTop" active={s.vAlign === 'top'} onClick={() => cmd.setAlign(undefined, 'top')} />
                <ToolButton tip="Middle align" icon="alignMiddle" active={s.vAlign === 'middle'} onClick={() => cmd.setAlign(undefined, 'middle')} />
                <ToolButton tip="Bottom align" icon="alignBottom" active={!s.vAlign || s.vAlign === 'bottom'} onClick={() => cmd.setAlign(undefined, 'bottom')} />
                <ToolButton tip="Wrap text" icon="wrap" active={!!s.wrap} onClick={() => cmd.toggle('wrap')} />
              </div>
              <div className="flex">
                <ToolButton tip="Align left" icon="alignLeft" active={s.hAlign === 'left'} onClick={() => cmd.setAlign('left')} />
                <ToolButton tip="Center" icon="alignCenter" active={s.hAlign === 'center'} onClick={() => cmd.setAlign('center')} />
                <ToolButton tip="Align right" icon="alignRight" active={s.hAlign === 'right'} onClick={() => cmd.setAlign('right')} />
                <ToolButton tip="Decrease indent" label="⇤" onClick={() => f({ a: 'indent', delta: -1 })} />
                <ToolButton tip="Increase indent" label="⇥" onClick={() => f({ a: 'indent', delta: 1 })} />
                <Dd tip="Merge cells" icon="merge">
                  <MenuItem onSelect={() => cmd.merge('center')}>Merge &amp; Center</MenuItem>
                  <MenuItem onSelect={() => cmd.merge('merge')}>Merge cells</MenuItem>
                  <MenuItem onSelect={() => cmd.merge('unmerge')}>Unmerge cells</MenuItem>
                </Dd>
              </div>
            </Group>
            <Group label="Number" className="max-w-[200px]">
              <select
                aria-label="Number format"
                className={cn(selectCls, 'w-[150px]')}
                value={currentFmt}
                onChange={(e) => cmd.numFmt(FORMAT_PRESETS.find((p) => p.id === e.target.value)?.fmt)}
              >
                {FORMAT_PRESETS.map((p) => (
                  <option key={p.id} value={p.id}>
                    {p.label}
                  </option>
                ))}
                {currentFmt === 'custom' && <option value="custom">Custom: {s.numFmt}</option>}
              </select>
              <div className="flex">
                <Popover
                  trigger={({ toggle, open }) => (
                    <ToolButton tip="Currency format" onClick={toggle} active={open} label="$">
                      <Icon name="chevronDown" size={12} className="opacity-60" />
                    </ToolButton>
                  )}
                >
                  {(['$', '₹', '€', '£', '¥'] as const).map((sym) => (
                    <MenuItem key={sym} onSelect={() => cmd.numFmt(`${sym}#,##0.00`)}>
                      {sym} {({ $: 'Dollar', '₹': 'Indian Rupee', '€': 'Euro', '£': 'Pound', '¥': 'Yen' } as Record<string, string>)[sym]}
                    </MenuItem>
                  ))}
                </Popover>
                <ToolButton tip="Percent style" icon="percent" onClick={() => cmd.numFmt('0%')} />
                <ToolButton tip="Comma style (thousands separator)" label="," className="font-bold" onClick={() => cmd.numFmt('#,##0.00')} />
                <ToolButton tip="Increase decimal" icon="decInc" onClick={() => cmd.decimals(1)} />
                <ToolButton tip="Decrease decimal" icon="decDec" onClick={() => cmd.decimals(-1)} />
              </div>
            </Group>
            <Group label="Styles">
              <div className="flex flex-col">
                <Dd tip="Conditional formatting: highlight cells, data bars, color scales and icon sets" icon="palette" label="Conditional" wide>
                  <MenuLabel>Highlight cells rules</MenuLabel>
                  <MenuItem onSelect={() => cf({ type: 'cellIs', op: 'gt', v1: '' })}>Greater than…</MenuItem>
                  <MenuItem onSelect={() => cf({ type: 'cellIs', op: 'lt', v1: '' })}>Less than…</MenuItem>
                  <MenuItem onSelect={() => cf({ type: 'cellIs', op: 'between', v1: '', v2: '' })}>Between…</MenuItem>
                  <MenuItem onSelect={() => cf({ type: 'cellIs', op: 'eq', v1: '' })}>Equal to…</MenuItem>
                  <MenuItem onSelect={() => cf({ type: 'text', op: 'contains', text: '' })}>Text that contains…</MenuItem>
                  <MenuItem onSelect={() => cf({ type: 'duplicate' })}>Duplicate values…</MenuItem>
                  <MenuSep />
                  <MenuLabel>Top / bottom rules</MenuLabel>
                  <MenuItem onSelect={() => cf({ type: 'top', n: 10 })}>Top 10 items…</MenuItem>
                  <MenuItem onSelect={() => cf({ type: 'top', n: 10, bottom: true })}>Bottom 10 items…</MenuItem>
                  <MenuItem onSelect={() => cf({ type: 'average', above: true })}>Above average…</MenuItem>
                  <MenuSep />
                  <MenuLabel>Visual</MenuLabel>
                  <MenuItem onSelect={() => f({ a: 'cfQuick', rule: { type: 'dataBar', color: '#638ec6' } })}>Data bars (blue)</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'cfQuick', rule: { type: 'dataBar', color: '#63be7b' } })}>Data bars (green)</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'cfQuick', rule: { type: 'colorScale', colors: ['#f8696b', '#ffeb84', '#63be7b'] } })}>Color scale: red–yellow–green</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'cfQuick', rule: { type: 'colorScale', colors: ['#ffffff', '#63be7b'] } })}>Color scale: white–green</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'cfQuick', rule: { type: 'iconSet', set: 'traffic' } })}>Icon set: traffic lights</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'cfQuick', rule: { type: 'iconSet', set: 'arrows' } })}>Icon set: arrows</MenuItem>
                  <MenuSep />
                  <MenuItem icon="plus" onSelect={() => f({ a: 'cfNew' })}>New rule…</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'cfClear', scope: 'selection' })}>Clear rules from selected cells</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'cfClear', scope: 'sheet' })}>Clear rules from entire sheet</MenuItem>
                  <MenuItem icon="settings" onSelect={() => f({ a: 'cfManage' })}>Manage rules…</MenuItem>
                </Dd>
                <ToolButton tip={`Format as Table (${shortcut('Ctrl+T')})`} icon="table" label="Table" wide onClick={() => f({ a: 'table' })} />
                <ToolButton tip={`Format Cells (${shortcut('Ctrl+1')})`} icon="palette" label="Format Cells" wide onClick={() => f({ a: 'formatCells' })} />
              </div>
            </Group>
            <Group label="Cells">
              <div className="flex flex-col">
                <Dd tip="Insert cells, rows, columns or sheets" icon="rowsInsert" label="Insert" wide>
                  <MenuItem icon="rowsInsert" onSelect={() => cmd.insert('rowsAbove')}>Insert rows above</MenuItem>
                  <MenuItem onSelect={() => cmd.insert('rowsBelow')}>Insert rows below</MenuItem>
                  <MenuItem icon="colsInsert" onSelect={() => cmd.insert('colsLeft')}>Insert columns left</MenuItem>
                  <MenuItem onSelect={() => cmd.insert('colsRight')}>Insert columns right</MenuItem>
                  <MenuSep />
                  <MenuItem icon="sheet" onSelect={() => cmd.insert('sheet')} shortcut={shortcut('Shift+F11')}>Insert sheet</MenuItem>
                </Dd>
                <Dd tip="Delete rows, columns or sheets" icon="trash" label="Delete" wide>
                  <MenuItem onSelect={() => cmd.remove('rows')}>Delete rows</MenuItem>
                  <MenuItem onSelect={() => cmd.remove('cols')}>Delete columns</MenuItem>
                  <MenuSep />
                  <MenuItem danger onSelect={() => cmd.remove('sheet')}>Delete sheet</MenuItem>
                </Dd>
              </div>
              <Dd tip="Row height, column width, AutoFit, hide and unhide" icon="settings" label="Format" wide>
                <MenuLabel>Cell size</MenuLabel>
                <MenuItem onSelect={() => cmd.sizeDialog('row')}>Row height…</MenuItem>
                <MenuItem onSelect={() => cmd.autoFit('row')}>AutoFit row height</MenuItem>
                <MenuItem onSelect={() => cmd.sizeDialog('col')}>Column width…</MenuItem>
                <MenuItem onSelect={() => cmd.autoFit('col')}>AutoFit column width</MenuItem>
                <MenuSep />
                <MenuLabel>Visibility</MenuLabel>
                <MenuItem onSelect={() => cmd.hide('row', true)}>Hide rows</MenuItem>
                <MenuItem onSelect={() => cmd.hide('row', false)}>Unhide rows</MenuItem>
                <MenuItem onSelect={() => cmd.hide('col', true)}>Hide columns</MenuItem>
                <MenuItem onSelect={() => cmd.hide('col', false)}>Unhide columns</MenuItem>
              </Dd>
            </Group>
            <Group label="Editing">
              <div className="flex flex-col">
                <div className="flex">
                  <ToolButton tip="AutoSum (Alt+=)" icon="sigma" label="AutoSum" wide onClick={() => cmd.autoSum('SUM')} />
                  <Dd tip="AutoSum functions">
                    {(['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as const).map((f) => (
                      <MenuItem key={f} onSelect={() => cmd.autoSum(f)}>
                        {f === 'SUM' ? 'Sum' : f === 'AVERAGE' ? 'Average' : f === 'COUNT' ? 'Count numbers' : f === 'MAX' ? 'Max' : 'Min'}
                      </MenuItem>
                    ))}
                  </Dd>
                </div>
                <div className="flex">
                  <Dd tip="Clear" icon="eraser" label="Clear" wide>
                    <MenuItem onSelect={() => cmd.clear('all')}>Clear all</MenuItem>
                    <MenuItem onSelect={() => cmd.clear('contents')} shortcut="Del">Clear contents</MenuItem>
                    <MenuItem onSelect={() => cmd.clear('formats')}>Clear formats</MenuItem>
                  </Dd>
                  <ToolButton tip={`Find & Replace (${shortcut('Ctrl+H')})`} icon="search" label="Find" wide onClick={() => cmd.find(false)} />
                </div>
                <Dd tip="Sort & Filter" icon="filter" label="Sort & Filter" wide>
                  <MenuItem icon="sortAsc" onSelect={() => f({ a: 'sort', desc: false })}>Sort A to Z</MenuItem>
                  <MenuItem icon="sortDesc" onSelect={() => f({ a: 'sort', desc: true })}>Sort Z to A</MenuItem>
                  <MenuItem onSelect={() => f({ a: 'sortCustom' })}>Custom sort…</MenuItem>
                  <MenuSep />
                  <MenuItem icon="filter" checked={ctx.hasFilter} shortcut={shortcut('Ctrl+Shift+L')} onSelect={() => f({ a: 'filterToggle' })}>Filter</MenuItem>
                  <MenuItem disabled={!ctx.filterActive} onSelect={() => f({ a: 'filterClear' })}>Clear</MenuItem>
                  <MenuItem disabled={!ctx.hasFilter} onSelect={() => f({ a: 'filterReapply' })}>Reapply</MenuItem>
                </Dd>
              </div>
            </Group>
          </>
        )}
        {tab === 'Insert' && (
          <>
            <Group label="Functions">
              <ToolButton tip="Insert a function" icon="fx" label="Function" wide onClick={cmd.insertFunction} />
            </Group>
            <Group label="Structure">
              <ToolButton tip="Insert rows above the selection" icon="rowsInsert" label="Rows" wide onClick={() => cmd.insert('rowsAbove')} />
              <ToolButton tip="Insert columns left of the selection" icon="colsInsert" label="Columns" wide onClick={() => cmd.insert('colsLeft')} />
              <ToolButton tip="Insert a new worksheet" icon="sheet" label="Sheet" wide onClick={() => cmd.insert('sheet')} />
            </Group>
            <Group label="Tables">
              <ToolButton tip="Create a PivotTable to summarize the selected data" icon="pivot" label="PivotTable" wide onClick={() => f({ a: 'pivot' })} />
              <ToolButton tip={`Insert a table (${shortcut('Ctrl+T')})`} icon="table" label="Table" wide onClick={() => f({ a: 'table' })} />
            </Group>
            <Group label="Charts">
              <div className="flex flex-col">
                <div className="flex">
                  {CHART_TYPES.slice(0, 4).map((t) => <ToolButton key={t.type} tip={`Insert ${t.label.toLowerCase()} chart`} icon={t.icon} label={t.label} wide onClick={() => f({ a: 'chart', type: t.type })} />)}
                </div>
                <div className="flex">
                  {CHART_TYPES.slice(4).map((t) => <ToolButton key={t.type} tip={`Insert ${t.label.toLowerCase()} chart`} icon={t.icon} label={t.label.replace(' (XY)', '')} wide onClick={() => f({ a: 'chart', type: t.type })} />)}
                </div>
              </div>
            </Group>
            <Group label="Links & notes">
              <ToolButton tip={`Insert link (${shortcut('Ctrl+K')})`} icon="link" label="Link" wide onClick={() => f({ a: 'hyperlink' })} />
              <ToolButton tip="Insert note (Shift+F2)" icon="note" label="Note" wide onClick={() => f({ a: 'note' })} />
            </Group>
            <Group label="Not yet available">
              <Planned phase="Planned" items="Sparklines, images, shapes, text boxes and checkboxes." />
            </Group>
          </>
        )}
        {tab === 'Page Layout' && (
          <>
            <Group label="Sheet options">
              <div className="flex flex-col">
                <Check label="View gridlines" checked={view.showGrid} onChange={(v) => setView({ showGrid: v })} />
                <Check label="View headings" checked={view.showHeadings} onChange={(v) => setView({ showHeadings: v })} />
              </div>
            </Group>
            <Group label="Page setup">
              <Dd tip="Margins" icon="page" label="Margins" wide>
                <MenuItem checked={ps.margins.top === 0.75 && ps.margins.left === 0.7} onSelect={() => f({ a: 'page', patch: { margins: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7 } } })}>Normal</MenuItem>
                <MenuItem checked={ps.margins.left === 0.25} onSelect={() => f({ a: 'page', patch: { margins: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25 } } })}>Narrow</MenuItem>
                <MenuItem checked={ps.margins.left === 1} onSelect={() => f({ a: 'page', patch: { margins: { top: 1, bottom: 1, left: 1, right: 1 } } })}>Wide</MenuItem>
                <MenuSep />
                <MenuItem onSelect={() => f({ a: 'pageSetup' })}>Custom margins…</MenuItem>
              </Dd>
              <Dd tip="Orientation" icon="page" label={ps.orientation === 'portrait' ? 'Portrait' : 'Landscape'} wide>
                <MenuItem checked={ps.orientation === 'portrait'} onSelect={() => f({ a: 'page', patch: { orientation: 'portrait' } })}>Portrait</MenuItem>
                <MenuItem checked={ps.orientation === 'landscape'} onSelect={() => f({ a: 'page', patch: { orientation: 'landscape' } })}>Landscape</MenuItem>
              </Dd>
              <Dd tip="Paper size" icon="page" label={ps.paper} wide>
                {(['A4', 'Letter', 'Legal'] as const).map((p) => <MenuItem key={p} checked={ps.paper === p} onSelect={() => f({ a: 'page', patch: { paper: p } })}>{p}</MenuItem>)}
              </Dd>
              <Dd tip="Print area" icon="grid" label="Print Area" wide>
                <MenuItem onSelect={() => f({ a: 'printArea', set: true })}>Set print area</MenuItem>
                <MenuItem disabled={!ps.printArea} onSelect={() => f({ a: 'printArea', set: false })}>Clear print area</MenuItem>
              </Dd>
              <Dd tip="Page breaks" icon="split" label="Breaks" wide>
                <MenuItem onSelect={() => f({ a: 'pageBreak', axis: 'row' })}>Insert page break above row</MenuItem>
                <MenuItem onSelect={() => f({ a: 'pageBreak', axis: 'col' })}>Insert page break left of column</MenuItem>
                <MenuItem disabled={!ps.rowBreaks.length && !ps.colBreaks.length} onSelect={() => f({ a: 'pageBreak', axis: 'reset' })}>Reset all page breaks</MenuItem>
              </Dd>
              <ToolButton tip="Page setup: margins, scaling, header and footer" icon="settings" label="Page Setup" wide onClick={() => f({ a: 'pageSetup' })} />
            </Group>
            <Group label="Scale to fit">
              <div className="flex flex-col gap-0.5 text-[12px]">
                <Check label="Fit all columns on one page" checked={ps.fitToWidth} onChange={(v) => f({ a: 'page', patch: { fitToWidth: v } })} />
                <label className="flex items-center gap-1.5">
                  Scale
                  <select aria-label="Print scale" className={cn(selectCls, 'w-[70px]')} disabled={ps.fitToWidth} value={ps.scale} onChange={(e) => f({ a: 'page', patch: { scale: +e.target.value } })}>
                    {[...new Set([50, 60, 70, 75, 80, 90, 100, 110, 125, 150, 200, ps.scale])].sort((a, b) => a - b).map((n) => <option key={n} value={n}>{n}%</option>)}
                  </select>
                </label>
              </div>
            </Group>
            <Group label="Print">
              <div className="flex flex-col">
                <Check label="Print gridlines" checked={ps.gridlines} onChange={(v) => f({ a: 'page', patch: { gridlines: v } })} />
                <Check label="Print headings" checked={ps.headings} onChange={(v) => f({ a: 'page', patch: { headings: v } })} />
              </div>
              <ToolButton tip={`Print preview (${shortcut('Ctrl+P')})`} icon="print" label="Preview" wide onClick={() => f({ a: 'printPreview' })} />
              <ToolButton tip="Download the active sheet as PDF" icon="download" label="PDF" wide onClick={() => f({ a: 'pdf' })} />
            </Group>
          </>
        )}
        {tab === 'Formulas' && (
          <>
            <Group label="Function library">
              <ToolButton tip="Insert Function: browse and search all supported functions" icon="fx" label="Insert Function" wide onClick={cmd.insertFunction} />
              <div className="flex flex-col">
                <div className="flex">
                  {FN_CATEGORIES.slice(0, 4).map(([cat, fns]) => (
                    <Dd key={cat} tip={`${cat} functions`} label={cat.split(' ')[0]} wide>
                      <div className="max-h-80 overflow-auto">{fns.map((n) => <MenuItem key={n} onSelect={() => insertFn(n)}>{n}</MenuItem>)}</div>
                    </Dd>
                  ))}
                </div>
                <div className="flex">
                  {FN_CATEGORIES.slice(4).map(([cat, fns]) => (
                    <Dd key={cat} tip={`${cat} functions`} label={cat.split(' ')[0]} wide>
                      <div className="max-h-80 overflow-auto">{fns.map((n) => <MenuItem key={n} onSelect={() => insertFn(n)}>{n}</MenuItem>)}</div>
                    </Dd>
                  ))}
                  <Dd tip="Recently used functions" label="Recent" wide>
                    {recentFns().length ? recentFns().map((n) => <MenuItem key={n} onSelect={() => insertFn(n)}>{n}</MenuItem>) : <div className="px-2 py-1 text-[12px] text-[var(--muted)]">Functions you insert appear here.</div>}
                  </Dd>
                </div>
              </div>
              <div className="flex">
                <ToolButton tip="AutoSum" icon="sigma" label="AutoSum" wide onClick={() => cmd.autoSum('SUM')} />
                <Dd tip="AutoSum functions">
                  {(['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'] as const).map((f) => (
                    <MenuItem key={f} onSelect={() => cmd.autoSum(f)}>
                      {f}
                    </MenuItem>
                  ))}
                </Dd>
              </div>
            </Group>
            <Group label="Defined names">
              <ToolButton tip="Define a named range from the selection" icon="tag" label="Define Name" wide onClick={cmd.defineName} />
              <ToolButton tip="Name Manager: edit and delete names" icon="book" label="Name Manager" wide onClick={() => f({ a: 'nameManager' })} />
            </Group>
            <Group label="Auditing">
              <ToolButton tip={`Show formulas (${shortcut('Ctrl+`')})`} icon="eye" label="Show Formulas" wide active={view.showFormulas} onClick={() => setView({ showFormulas: !view.showFormulas })} />
              <ToolButton tip="Which Excel functions ExcelPro supports" icon="book" label="Compatibility" wide onClick={cmd.showCompat} />
            </Group>
            <Group label="Calculation">
              <div className="px-1 text-[11.5px] leading-4 text-[var(--muted)]">
                Automatic
                <br />
                (recalculates on every change)
              </div>
            </Group>
            <Group label="Not yet available">
              <Planned phase="Planned" items="Trace precedents/dependents and manual calculation mode." />
            </Group>
          </>
        )}
        {tab === 'Data' && (
          <>
            <Group label="Get data">
              <ToolButton tip="Import an .xlsx, .csv or .tsv file as a new workbook" icon="upload" label="Import" wide onClick={cmd.importFile} />
              <ToolButton tip="Download as .xlsx, .csv or .tsv" icon="download" label="Export" wide onClick={cmd.exportDialog} />
            </Group>
            <Group label="Fill">
              <ToolButton tip={`Fill down (${shortcut('Ctrl+D')})`} label="Fill Down" wide onClick={cmd.fillDown} />
              <ToolButton tip={`Fill right (${shortcut('Ctrl+R')})`} label="Fill Right" wide onClick={cmd.fillRight} />
            </Group>
            <Group label="Sort & Filter">
              <div className="flex flex-col">
                <ToolButton tip="Sort A to Z (smallest to largest)" icon="sortAsc" label="A→Z" wide onClick={() => f({ a: 'sort', desc: false })} />
                <ToolButton tip="Sort Z to A (largest to smallest)" icon="sortDesc" label="Z→A" wide onClick={() => f({ a: 'sort', desc: true })} />
              </div>
              <ToolButton tip="Sort by several columns" icon="sortAsc" label="Sort…" wide onClick={() => f({ a: 'sortCustom' })} />
              <ToolButton tip={`Filter (${shortcut('Ctrl+Shift+L')})`} icon="filter" label="Filter" wide active={ctx.hasFilter} onClick={() => f({ a: 'filterToggle' })} />
              <div className="flex flex-col">
                <ToolButton tip="Clear all filters on this sheet" label="Clear" wide disabled={!ctx.filterActive} onClick={() => f({ a: 'filterClear' })} />
                <ToolButton tip="Reapply the filter to changed data" icon="refresh" label="Reapply" wide disabled={!ctx.hasFilter} onClick={() => f({ a: 'filterReapply' })} />
              </div>
            </Group>
            <Group label="Data tools">
              <div className="flex flex-col">
                <ToolButton tip="Split one column into several by a delimiter or fixed width" icon="split" label="Text to Columns" wide onClick={() => f({ a: 'textToColumns' })} />
                <ToolButton tip={`Flash Fill: fill the column from your examples (${shortcut('Ctrl+E')})`} icon="wand" label="Flash Fill" wide onClick={() => f({ a: 'flashFill' })} />
              </div>
              <div className="flex flex-col">
                <ToolButton tip="Remove duplicate rows" icon="dup" label="Remove Duplicates" wide onClick={() => f({ a: 'removeDuplicates' })} />
                <Dd tip="Data validation" icon="check2" label="Validation" wide>
                  <MenuItem onSelect={() => f({ a: 'validation' })}>Data validation…</MenuItem>
                  <MenuItem checked={ctx.showInvalid} onSelect={() => f({ a: 'circleInvalid', on: true })}>Circle invalid data</MenuItem>
                  <MenuItem disabled={!ctx.showInvalid} onSelect={() => f({ a: 'circleInvalid', on: false })}>Clear validation circles</MenuItem>
                </Dd>
              </div>
              <Dd tip="Clean text in the selected cells" icon="eraser" label="Clean" wide>
                <MenuItem onSelect={() => f({ a: 'clean', kind: 'trim' })}>Trim extra spaces</MenuItem>
                <MenuItem onSelect={() => f({ a: 'clean', kind: 'clean' })}>Remove non-printing characters</MenuItem>
                <MenuItem onSelect={() => f({ a: 'clean', kind: 'toNumber' })}>Convert text to numbers</MenuItem>
                <MenuSep />
                <MenuItem onSelect={() => f({ a: 'clean', kind: 'upper' })}>UPPERCASE</MenuItem>
                <MenuItem onSelect={() => f({ a: 'clean', kind: 'lower' })}>lowercase</MenuItem>
                <MenuItem onSelect={() => f({ a: 'clean', kind: 'proper' })}>Proper Case</MenuItem>
              </Dd>
            </Group>
            <Group label="Forecast & pivots">
              <ToolButton tip="Goal Seek: find the input that gives a target result" icon="target" label="Goal Seek" wide onClick={() => f({ a: 'goalSeek' })} />
              <ToolButton tip="Create a PivotTable" icon="pivot" label="PivotTable" wide onClick={() => f({ a: 'pivot' })} />
              <ToolButton tip="Refresh all PivotTables" icon="refresh" label="Refresh All" wide disabled={!ctx.pivotCount} onClick={() => f({ a: 'refreshPivots' })} />
            </Group>
          </>
        )}
        {tab === 'Review' && (
          <>
            <Group label="Notes">
              <ToolButton tip="New or edit note (Shift+F2)" icon="note" label={ctx.hasNote ? 'Edit Note' : 'New Note'} wide onClick={() => f({ a: 'note' })} />
              <ToolButton tip="Delete the note on the active cell" icon="trash" label="Delete Note" wide disabled={!ctx.hasNote} onClick={() => f({ a: 'deleteNote' })} />
            </Group>
            <Group label="Data">
              <ToolButton tip="Circle cells that break their data validation rule" icon="check2" label="Circle Invalid" wide active={ctx.showInvalid} onClick={() => f({ a: 'circleInvalid', on: !ctx.showInvalid })} />
              <ToolButton tip={`Find & Replace (${shortcut('Ctrl+H')})`} icon="search" label="Find" wide onClick={() => cmd.find(true)} />
            </Group>
            <Group label="Not yet available">
              <Planned phase="Phase 4" items="Threaded comments, sheet & workbook protection, unlocking cells and version history (these need accounts and cloud storage)." />
            </Group>
          </>
        )}
        {tab === 'View' && (
          <>
            <Group label="Show">
              <div className="flex flex-col">
                <Check label="Gridlines" checked={view.showGrid} onChange={(v) => setView({ showGrid: v })} />
                <Check label="Headings" checked={view.showHeadings} onChange={(v) => setView({ showHeadings: v })} />
              </div>
              <div className="flex flex-col">
                <Check label="Formula bar" checked={view.showFormulaBar} onChange={(v) => setView({ showFormulaBar: v })} />
                <Check label="Formulas in cells" checked={view.showFormulas} onChange={(v) => setView({ showFormulas: v })} />
              </div>
            </Group>
            <Group label="Zoom">
              <ToolButton tip="Zoom out" icon="minus" onClick={() => setView({ zoom: Math.max(0.5, +(view.zoom - 0.1).toFixed(2)) })} />
              <ToolButton tip="Reset zoom to 100%" label={`${Math.round(view.zoom * 100)}%`} wide onClick={() => setView({ zoom: 1 })} />
              <ToolButton tip="Zoom in" icon="plus" onClick={() => setView({ zoom: Math.min(2, +(view.zoom + 0.1).toFixed(2)) })} />
            </Group>
            <Group label="Window">
              <ToolButton tip="Dark mode" icon={view.dark ? 'sun' : 'moon'} label={view.dark ? 'Light' : 'Dark'} wide onClick={() => setView({ dark: !view.dark })} />
              <ToolButton
                tip="Full screen"
                icon="expand"
                label="Full Screen"
                wide
                onClick={() => {
                  if (document.fullscreenElement) void document.exitFullscreen()
                  else void document.documentElement.requestFullscreen?.().catch(() => {})
                }}
              />
            </Group>
            <Group label="Panes">
              <Dd tip="Freeze panes" icon="freeze" label="Freeze" wide>
                <MenuItem onSelect={() => f({ a: 'freeze', mode: 'panes' })}>Freeze panes (above and left of the active cell)</MenuItem>
                <MenuItem checked={ctx.freeze?.rows === 1 && !ctx.freeze.cols} onSelect={() => f({ a: 'freeze', mode: 'top' })}>Freeze top row</MenuItem>
                <MenuItem checked={ctx.freeze?.cols === 1 && !ctx.freeze.rows} onSelect={() => f({ a: 'freeze', mode: 'first' })}>Freeze first column</MenuItem>
                <MenuSep />
                <MenuItem disabled={!ctx.freeze} onSelect={() => f({ a: 'freeze', mode: 'none' })}>Unfreeze panes</MenuItem>
              </Dd>
            </Group>
            <Group label="Workbook views">
              <ToolButton tip="Show where pages break when printed" icon="split" label="Page Breaks" wide active={ctx.pageBreakPreview} onClick={() => f({ a: 'pageBreakPreview' })} />
              <ToolButton tip="Print preview" icon="print" label="Print Preview" wide onClick={() => f({ a: 'printPreview' })} />
            </Group>
          </>
        )}
        {tab === 'Table Design' && ctx.table && (
          <>
            <Group label="Properties">
              <label className="flex items-center gap-1.5 text-[12px]">
                Name
                <input
                  key={ctx.table.id + ctx.table.name}
                  aria-label="Table name"
                  defaultValue={ctx.table.name}
                  className={cn(selectCls, 'w-[120px]')}
                  onBlur={(e) => e.target.value !== ctx.table!.name && f({ a: 'tableUpdate', patch: { name: e.target.value } })}
                  onKeyDown={(e) => e.key === 'Enter' && (e.target as HTMLInputElement).blur()}
                />
              </label>
            </Group>
            <Group label="Tools">
              <ToolButton tip="Summarize with PivotTable" icon="pivot" label="PivotTable" wide onClick={() => f({ a: 'pivot' })} />
              <ToolButton tip="Remove duplicate rows from the table" icon="dup" label="Remove Duplicates" wide onClick={() => f({ a: 'removeDuplicates' })} />
              <ToolButton tip="Convert to a normal range (keeps formatting)" icon="grid" label="Convert to Range" wide onClick={() => f({ a: 'tableConvert' })} />
            </Group>
            <Group label="Style options">
              <div className="flex flex-col">
                <Check label="Total row" checked={ctx.table.totalRow} onChange={(v) => f({ a: 'tableUpdate', patch: { totalRow: v } })} />
                <Check label="Banded rows" checked={ctx.table.banded} onChange={(v) => f({ a: 'tableUpdate', patch: { banded: v } })} />
              </div>
              {ctx.table.totalRow && (
                <label className="flex items-center gap-1.5 text-[12px]">
                  Total for this column
                  <select
                    aria-label="Total function"
                    className={selectCls}
                    value={ctx.table.totals[ctx.anchorCol - ctx.table.range.c1] ?? 'none'}
                    onChange={(e) => f({ a: 'tableTotal', fn: e.target.value as TotalFn })}
                  >
                    <option value="none">None</option>
                    <option value="sum">Sum</option>
                    <option value="average">Average</option>
                    <option value="count">Count</option>
                    <option value="countNumbers">Count numbers</option>
                    <option value="min">Min</option>
                    <option value="max">Max</option>
                  </select>
                </label>
              )}
            </Group>
            <Group label="Table styles">
              <div className="flex gap-1">
                {(Object.keys(TABLE_STYLES) as (keyof typeof TABLE_STYLES)[]).map((k) => (
                  <button key={k} type="button" aria-label={`${k} table style`} title={k} onClick={() => f({ a: 'tableUpdate', patch: { style: k } })} className={cn('w-10 overflow-hidden rounded border-2', ctx.table!.style === k ? 'border-[var(--accent)]' : 'border-transparent')}>
                    <div className="h-2" style={{ background: TABLE_STYLES[k].header }} />
                    <div className="h-1.5" style={{ background: TABLE_STYLES[k].band }} />
                    <div className="h-1.5 bg-white" />
                    <div className="h-1.5" style={{ background: TABLE_STYLES[k].band }} />
                  </button>
                ))}
              </div>
            </Group>
          </>
        )}
        {tab === 'Automate' && (
          <Group label="Coming next">
            <Planned phase="Phase 5" items="Reusable actions, automation templates and repeatable cleaning steps, run in a sandbox with explicit permissions (no server-side script execution)." />
          </Group>
        )}
        {tab === 'Help' && (
          <>
            <Group label="Help">
              <ToolButton tip="Keyboard shortcuts" icon="keyboard" label="Shortcuts" wide onClick={cmd.showShortcuts} />
              <ToolButton tip="Search functions by name or purpose" icon="fx" label="Formula search" wide onClick={cmd.insertFunction} />
              <ToolButton tip="Function compatibility list" icon="book" label="Compatibility" wide onClick={cmd.showCompat} />
              <ToolButton tip="What works in this release, file limits and troubleshooting" icon="help" label="About & help" wide onClick={cmd.showAbout} />
            </Group>
          </>
        )}
      </div>
    </div>
  )
}
