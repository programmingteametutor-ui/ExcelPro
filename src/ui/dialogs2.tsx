/* Phase 2–3 dialogs: format cells, conditional formatting, sort, validation, data tools, names, links, notes, page setup, tables. */
import { useMemo, useState } from 'react'
import type { Workbook } from '../engine/workbook'
import type { BorderStyle, CellBorders, CellStyle, CFRule, CFStyle, CompareOp, CondFormat, PageSetup, Range, TableDef, Validation } from '../model/types'
import { DEFAULT_PAGE_SETUP } from '../model/types'
import { colToLetters, parseRangeA1, quoteSheetName, rangeToA1 } from '../model/address'
import { FORMAT_PRESETS, formatNumber } from '../engine/numberFormat'
import { CF_PRESET_STYLES } from '../engine/features/condFormat'
import { TABLE_STYLES } from '../engine/features/tables'
import { changeNamedRange, removeNamedRange, type SortKey } from '../engine/features/data'
import { Button, Check, Dialog, Field, inputClass } from './components/ui'
import { Icon } from './icons'
import { cn } from './lib/utils'

const sel = cn(inputClass, 'h-8')

// ------------------------------------------------------------------ Format Cells
const BORDER_STYLES: BorderStyle[] = ['thin', 'medium', 'thick', 'dashed', 'dotted', 'double']

export function FormatCellsDialog({ initial, sample, onClose, onApply }: { initial: CellStyle | undefined; sample: unknown; onClose: () => void; onApply: (s: CellStyle) => void }) {
  const [tab, setTab] = useState<'number' | 'alignment' | 'font' | 'border' | 'fill'>('number')
  const [s, setS] = useState<CellStyle>({ ...initial })
  const [bStyle, setBStyle] = useState<BorderStyle>('thin')
  const [bColor, setBColor] = useState('#000000')
  const up = (p: Partial<CellStyle>) => setS((x) => ({ ...x, ...p }))
  const preview = typeof sample === 'number' ? formatNumber(sample, s.numFmt).text : String(sample ?? '')
  const setSide = (side: keyof CellBorders | 'none' | 'outline' | 'all') => {
    const e = { style: bStyle, color: bColor }
    if (side === 'none') up({ borders: undefined })
    else if (side === 'outline' || side === 'all') up({ borders: { top: e, right: e, bottom: e, left: e } })
    else up({ borders: { ...s.borders, [side]: s.borders?.[side] ? undefined : e } })
  }
  return (
    <Dialog
      title="Format cells"
      onClose={onClose}
      width={560}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" onClick={() => onApply(s)}>OK</Button>
        </>
      }
    >
      <div className="mb-3 flex gap-1 border-b border-[var(--border)]">
        {(['number', 'alignment', 'font', 'border', 'fill'] as const).map((t) => (
          <button key={t} type="button" onClick={() => setTab(t)} className={cn('-mb-px border-b-2 px-3 py-1.5 text-[12.5px] capitalize', tab === t ? 'border-[var(--accent)] font-semibold text-[var(--accent)]' : 'border-transparent')}>
            {t}
          </button>
        ))}
      </div>
      {tab === 'number' && (
        <div className="grid grid-cols-[170px_1fr] gap-3">
          <div className="max-h-64 overflow-auto rounded-md border border-[var(--border)]">
            {FORMAT_PRESETS.map((p) => (
              <button key={p.id} type="button" className={cn('block w-full px-2 py-1 text-left text-[12.5px] hover:bg-[var(--hover)]', s.numFmt === p.fmt && 'bg-[var(--active)]')} onClick={() => up({ numFmt: p.fmt })}>
                {p.label}
              </button>
            ))}
          </div>
          <div>
            <Field label="Format code" hint="Excel format codes, e.g. #,##0.00;[Red]-#,##0.00 or dd-mmm-yyyy or 0.0%">
              <input className={cn(inputClass, 'font-mono')} value={s.numFmt ?? 'General'} onChange={(e) => up({ numFmt: e.target.value === 'General' || !e.target.value ? undefined : e.target.value })} aria-label="Format code" />
            </Field>
            <div className="rounded-md bg-[var(--panel)] p-2 text-[12.5px]">
              <span className="text-[var(--muted)]">Sample: </span>
              <span className="font-mono">{preview || '—'}</span>
            </div>
          </div>
        </div>
      )}
      {tab === 'alignment' && (
        <div className="grid grid-cols-2 gap-3">
          <Field label="Horizontal">
            <select className={sel} value={s.hAlign ?? ''} onChange={(e) => up({ hAlign: (e.target.value || undefined) as CellStyle['hAlign'] })}>
              <option value="">General</option>
              <option value="left">Left</option>
              <option value="center">Center</option>
              <option value="right">Right</option>
            </select>
          </Field>
          <Field label="Vertical">
            <select className={sel} value={s.vAlign ?? 'bottom'} onChange={(e) => up({ vAlign: e.target.value === 'bottom' ? undefined : (e.target.value as CellStyle['vAlign']) })}>
              <option value="top">Top</option>
              <option value="middle">Middle</option>
              <option value="bottom">Bottom</option>
            </select>
          </Field>
          <Field label="Indent">
            <input className={inputClass} type="number" min={0} max={15} value={s.indent ?? 0} onChange={(e) => up({ indent: +e.target.value || undefined })} />
          </Field>
          <div className="pt-5">
            <Check label="Wrap text" checked={!!s.wrap} onChange={(v) => up({ wrap: v || undefined })} />
          </div>
        </div>
      )}
      {tab === 'font' && (
        <div className="grid grid-cols-2 gap-3">
          <Field label="Font">
            <input className={inputClass} value={s.fontFamily ?? 'Calibri'} onChange={(e) => up({ fontFamily: e.target.value === 'Calibri' ? undefined : e.target.value })} />
          </Field>
          <Field label="Size">
            <input className={inputClass} type="number" min={6} max={96} value={s.fontSize ?? 11} onChange={(e) => up({ fontSize: +e.target.value === 11 ? undefined : +e.target.value })} />
          </Field>
          <div className="flex flex-col">
            <Check label="Bold" checked={!!s.bold} onChange={(v) => up({ bold: v || undefined })} />
            <Check label="Italic" checked={!!s.italic} onChange={(v) => up({ italic: v || undefined })} />
            <Check label="Underline" checked={!!s.underline} onChange={(v) => up({ underline: v || undefined })} />
            <Check label="Strikethrough" checked={!!s.strike} onChange={(v) => up({ strike: v || undefined })} />
          </div>
          <Field label="Color">
            <div className="flex items-center gap-2">
              <input type="color" value={s.color ?? '#000000'} onChange={(e) => up({ color: e.target.value })} className="h-8 w-12 rounded border border-[var(--border)]" aria-label="Font color" />
              <Button size="sm" variant="ghost" onClick={() => up({ color: undefined })}>Automatic</Button>
            </div>
          </Field>
        </div>
      )}
      {tab === 'border' && (
        <div className="grid grid-cols-[1fr_180px] gap-4">
          <div>
            <div className="mb-2 flex flex-wrap gap-1">
              <Button size="sm" variant="outline" onClick={() => setSide('none')}>None</Button>
              <Button size="sm" variant="outline" onClick={() => setSide('outline')}>All sides</Button>
              {(['top', 'bottom', 'left', 'right'] as const).map((sd) => (
                <Button key={sd} size="sm" variant={s.borders?.[sd] ? 'default' : 'outline'} onClick={() => setSide(sd)} className="capitalize">
                  {sd}
                </Button>
              ))}
            </div>
            <div className="flex h-24 items-center justify-center rounded-md bg-[var(--panel)]">
              <div
                className="flex h-14 w-32 items-center justify-center bg-white text-[12px] text-black"
                style={{
                  borderTop: s.borders?.top ? `2px ${s.borders.top.style === 'dashed' ? 'dashed' : s.borders.top.style === 'dotted' ? 'dotted' : 'solid'} ${s.borders.top.color}` : '1px dotted #ccc',
                  borderBottom: s.borders?.bottom ? `2px solid ${s.borders.bottom.color}` : '1px dotted #ccc',
                  borderLeft: s.borders?.left ? `2px solid ${s.borders.left.color}` : '1px dotted #ccc',
                  borderRight: s.borders?.right ? `2px solid ${s.borders.right.color}` : '1px dotted #ccc',
                }}
              >
                Text
              </div>
            </div>
          </div>
          <div>
            <Field label="Line style">
              <select className={sel} value={bStyle} onChange={(e) => setBStyle(e.target.value as BorderStyle)}>
                {BORDER_STYLES.map((b) => <option key={b}>{b}</option>)}
              </select>
            </Field>
            <Field label="Line color">
              <input type="color" value={bColor} onChange={(e) => setBColor(e.target.value)} className="h-8 w-full rounded border border-[var(--border)]" aria-label="Border color" />
            </Field>
          </div>
        </div>
      )}
      {tab === 'fill' && (
        <Field label="Background color">
          <div className="flex items-center gap-2">
            <input type="color" value={s.fill ?? '#ffffff'} onChange={(e) => up({ fill: e.target.value })} className="h-8 w-16 rounded border border-[var(--border)]" aria-label="Fill color" />
            <Button size="sm" variant="ghost" onClick={() => up({ fill: undefined })}>No fill</Button>
          </div>
        </Field>
      )}
    </Dialog>
  )
}

// ------------------------------------------------------------------ Conditional formatting
type RuleKind = 'cellIs' | 'text' | 'duplicate' | 'unique' | 'blank' | 'notBlank' | 'errors' | 'formula' | 'top' | 'bottom' | 'above' | 'below' | 'colorScale2' | 'colorScale3' | 'dataBar' | 'iconSet'

export function CFRuleDialog({ initial, range, onClose, onSave }: { initial?: CondFormat; range: Range; onClose: () => void; onSave: (cf: Omit<CondFormat, 'id'>) => void }) {
  const r0 = initial?.rule
  const startKind: RuleKind = !r0 ? 'cellIs' : r0.type === 'top' ? (r0.bottom ? 'bottom' : 'top') : r0.type === 'average' ? (r0.above ? 'above' : 'below') : r0.type === 'colorScale' ? (r0.colors.length >= 3 ? 'colorScale3' : 'colorScale2') : (r0.type as RuleKind)
  const [kind, setKind] = useState<RuleKind>(startKind)
  const [op, setOp] = useState<CompareOp>(r0?.type === 'cellIs' ? r0.op : 'gt')
  const [v1, setV1] = useState(r0?.type === 'cellIs' ? r0.v1 : '')
  const [v2, setV2] = useState(r0?.type === 'cellIs' ? r0.v2 ?? '' : '')
  const [textOp, setTextOp] = useState<'contains' | 'notContains' | 'begins' | 'ends'>(r0?.type === 'text' ? r0.op : 'contains')
  const [text, setText] = useState(r0?.type === 'text' ? r0.text : '')
  const [formula, setFormula] = useState(r0?.type === 'formula' ? r0.formula : `=${colToLetters(range.c1)}${range.r1 + 1}>0`)
  const [n, setN] = useState(r0?.type === 'top' ? r0.n : 10)
  const [percent, setPercent] = useState(r0?.type === 'top' ? !!r0.percent : false)
  const [colors, setColors] = useState<string[]>(r0?.type === 'colorScale' ? r0.colors : ['#f8696b', '#ffeb84', '#63be7b'])
  const [barColor, setBarColor] = useState(r0?.type === 'dataBar' ? r0.color : '#638ec6')
  const [iconSet, setIconSet] = useState<'arrows' | 'traffic' | 'flags'>(r0?.type === 'iconSet' ? r0.set : 'traffic')
  const [style, setStyle] = useState<CFStyle>(initial?.style ?? CF_PRESET_STYLES[0].style)
  const [rangeText, setRangeText] = useState(rangeToA1(initial?.range ?? range))
  const [stop, setStop] = useState(!!initial?.stopIfTrue)
  const parsed = parseRangeA1(rangeText)
  const needsStyle = !['colorScale2', 'colorScale3', 'dataBar', 'iconSet'].includes(kind)
  const build = (): CFRule => {
    switch (kind) {
      case 'cellIs': return { type: 'cellIs', op, v1, v2: op === 'between' || op === 'notBetween' ? v2 : undefined }
      case 'text': return { type: 'text', op: textOp, text }
      case 'formula': return { type: 'formula', formula }
      case 'top': return { type: 'top', n, percent }
      case 'bottom': return { type: 'top', n, percent, bottom: true }
      case 'above': return { type: 'average', above: true }
      case 'below': return { type: 'average', above: false }
      case 'colorScale2': return { type: 'colorScale', colors: [colors[0], colors[2]] }
      case 'colorScale3': return { type: 'colorScale', colors }
      case 'dataBar': return { type: 'dataBar', color: barColor }
      case 'iconSet': return { type: 'iconSet', set: iconSet }
      default: return { type: kind } as CFRule
    }
  }
  return (
    <Dialog
      title={initial?.id ? 'Edit formatting rule' : 'New formatting rule'}
      onClose={onClose}
      width={540}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" disabled={!parsed} onClick={() => parsed && onSave({ range: parsed, rule: build(), style: needsStyle ? style : undefined, stopIfTrue: stop || undefined })}>OK</Button>
        </>
      }
    >
      <Field label="Applies to">
        <input className={cn(inputClass, 'font-mono')} value={rangeText} onChange={(e) => setRangeText(e.target.value)} aria-label="Applies to" />
      </Field>
      <Field label="Format cells where">
        <select className={sel} value={kind} onChange={(e) => setKind(e.target.value as RuleKind)} aria-label="Rule type">
          <optgroup label="Highlight cells">
            <option value="cellIs">Cell value is…</option>
            <option value="text">Text that contains…</option>
            <option value="duplicate">Duplicate values</option>
            <option value="unique">Unique values</option>
            <option value="blank">Blanks</option>
            <option value="notBlank">No blanks</option>
            <option value="errors">Errors</option>
            <option value="formula">A formula is true…</option>
          </optgroup>
          <optgroup label="Top / bottom">
            <option value="top">Top N</option>
            <option value="bottom">Bottom N</option>
            <option value="above">Above average</option>
            <option value="below">Below average</option>
          </optgroup>
          <optgroup label="Visual">
            <option value="colorScale2">2-color scale</option>
            <option value="colorScale3">3-color scale</option>
            <option value="dataBar">Data bars</option>
            <option value="iconSet">Icon set</option>
          </optgroup>
        </select>
      </Field>
      {kind === 'cellIs' && (
        <div className="mb-3 flex flex-wrap items-center gap-2">
          <select className={sel} value={op} onChange={(e) => setOp(e.target.value as CompareOp)} aria-label="Operator">
            <option value="gt">greater than</option>
            <option value="lt">less than</option>
            <option value="gte">greater than or equal to</option>
            <option value="lte">less than or equal to</option>
            <option value="eq">equal to</option>
            <option value="ne">not equal to</option>
            <option value="between">between</option>
            <option value="notBetween">not between</option>
          </select>
          <input className={cn(inputClass, 'w-28')} value={v1} onChange={(e) => setV1(e.target.value)} placeholder="value or =A1" aria-label="Value" />
          {(op === 'between' || op === 'notBetween') && (
            <>
              and
              <input className={cn(inputClass, 'w-28')} value={v2} onChange={(e) => setV2(e.target.value)} aria-label="Second value" />
            </>
          )}
        </div>
      )}
      {kind === 'text' && (
        <div className="mb-3 flex gap-2">
          <select className={sel} value={textOp} onChange={(e) => setTextOp(e.target.value as typeof textOp)}>
            <option value="contains">contains</option>
            <option value="notContains">does not contain</option>
            <option value="begins">begins with</option>
            <option value="ends">ends with</option>
          </select>
          <input className={cn(inputClass, 'flex-1')} value={text} onChange={(e) => setText(e.target.value)} aria-label="Text" />
        </div>
      )}
      {kind === 'formula' && (
        <Field label={`Formula (relative to ${rangeText.split(':')[0]})`} hint="Use $ to lock columns, e.g. =$C2>1000 formats whole rows.">
          <input className={cn(inputClass, 'font-mono')} value={formula} onChange={(e) => setFormula(e.target.value)} aria-label="Formula" />
        </Field>
      )}
      {(kind === 'top' || kind === 'bottom') && (
        <div className="mb-3 flex items-center gap-2">
          <input className={cn(inputClass, 'w-20')} type="number" min={1} value={n} onChange={(e) => setN(+e.target.value || 1)} aria-label="N" />
          <Check label="% of the range" checked={percent} onChange={setPercent} />
        </div>
      )}
      {(kind === 'colorScale2' || kind === 'colorScale3') && (
        <div className="mb-3 flex items-center gap-3 text-[12.5px]">
          {['Lowest', 'Middle', 'Highest'].map((lbl, i) =>
            kind === 'colorScale2' && i === 1 ? null : (
              <label key={lbl} className="flex items-center gap-1.5">
                {lbl}
                <input type="color" value={colors[i]} onChange={(e) => setColors((c) => c.map((x, j) => (j === i ? e.target.value : x)))} className="h-7 w-10 rounded border border-[var(--border)]" />
              </label>
            ),
          )}
        </div>
      )}
      {kind === 'dataBar' && (
        <label className="mb-3 flex items-center gap-2 text-[12.5px]">
          Bar color <input type="color" value={barColor} onChange={(e) => setBarColor(e.target.value)} className="h-7 w-10 rounded border border-[var(--border)]" />
        </label>
      )}
      {kind === 'iconSet' && (
        <Field label="Icons">
          <select className={sel} value={iconSet} onChange={(e) => setIconSet(e.target.value as typeof iconSet)}>
            <option value="traffic">Traffic lights</option>
            <option value="arrows">Arrows</option>
            <option value="flags">Flags</option>
          </select>
        </Field>
      )}
      {needsStyle && (
        <>
          <Field label="Format with">
            <select className={sel} value={CF_PRESET_STYLES.findIndex((p) => JSON.stringify(p.style) === JSON.stringify(style))} onChange={(e) => +e.target.value >= 0 && setStyle(CF_PRESET_STYLES[+e.target.value].style)}>
              {CF_PRESET_STYLES.map((p, i) => <option key={p.label} value={i}>{p.label}</option>)}
              <option value={-1}>Custom…</option>
            </select>
          </Field>
          <div className="mb-3 flex flex-wrap items-center gap-3 text-[12.5px]">
            <label className="flex items-center gap-1.5">Fill <input type="color" value={style.fill ?? '#ffffff'} onChange={(e) => setStyle({ ...style, fill: e.target.value })} className="h-7 w-10 rounded border border-[var(--border)]" /></label>
            <label className="flex items-center gap-1.5">Text <input type="color" value={style.color ?? '#000000'} onChange={(e) => setStyle({ ...style, color: e.target.value })} className="h-7 w-10 rounded border border-[var(--border)]" /></label>
            <Check label="Bold" checked={!!style.bold} onChange={(v) => setStyle({ ...style, bold: v || undefined })} />
            <Check label="Italic" checked={!!style.italic} onChange={(v) => setStyle({ ...style, italic: v || undefined })} />
            <span className="rounded px-2 py-0.5" style={{ background: style.fill, color: style.color, fontWeight: style.bold ? 700 : undefined, fontStyle: style.italic ? 'italic' : undefined }}>AaBbCc 123</span>
          </div>
          <Check label="Stop if true (lower rules don't apply when this one matches)" checked={stop} onChange={setStop} />
        </>
      )}
    </Dialog>
  )
}

function describeRule(r: CFRule): string {
  const ops: Record<string, string> = { gt: '>', lt: '<', gte: '≥', lte: '≤', eq: '=', ne: '≠', between: 'between', notBetween: 'not between' }
  switch (r.type) {
    case 'cellIs': return `Cell value ${ops[r.op]} ${r.v1}${r.v2 ? ` and ${r.v2}` : ''}`
    case 'text': return `Text ${r.op === 'notContains' ? 'does not contain' : r.op === 'begins' ? 'begins with' : r.op === 'ends' ? 'ends with' : 'contains'} “${r.text}”`
    case 'formula': return `Formula: ${r.formula}`
    case 'top': return `${r.bottom ? 'Bottom' : 'Top'} ${r.n}${r.percent ? '%' : ''}`
    case 'average': return r.above ? 'Above average' : 'Below average'
    case 'colorScale': return `${r.colors.length}-color scale`
    case 'dataBar': return 'Data bar'
    case 'iconSet': return `Icon set (${r.set})`
    default: return { duplicate: 'Duplicate values', unique: 'Unique values', blank: 'Blanks', notBlank: 'No blanks', errors: 'Errors' }[r.type]
  }
}

export function CFManagerDialog({ rules, onClose, onEdit, onDelete, onMove, onNew }: { rules: CondFormat[]; onClose: () => void; onEdit: (cf: CondFormat) => void; onDelete: (id: string) => void; onMove: (id: string, dir: -1 | 1) => void; onNew: () => void }) {
  return (
    <Dialog title="Conditional formatting rules" onClose={onClose} width={640} footer={<Button variant="default" onClick={onClose}>Done</Button>}>
      <div className="mb-2 flex items-center justify-between">
        <p className="text-[12px] text-[var(--muted)]">Rules higher in the list take priority.</p>
        <Button size="sm" variant="outline" onClick={onNew}><Icon name="plus" /> New rule</Button>
      </div>
      {rules.length === 0 ? (
        <p className="rounded-md bg-[var(--panel)] p-4 text-center text-[12.5px] text-[var(--muted)]">This sheet has no conditional formatting rules.</p>
      ) : (
        <div className="max-h-[50vh] overflow-auto rounded-md border border-[var(--border)]">
          {rules.map((cf, i) => (
            <div key={cf.id} className="flex items-center gap-2 border-t border-[var(--border)] px-2 py-1.5 text-[12.5px] first:border-t-0">
              <span className="flex h-6 w-14 shrink-0 items-center justify-center rounded text-[11px]" style={cf.rule.type === 'colorScale' ? { background: `linear-gradient(90deg, ${cf.rule.colors.join(',')})` } : cf.rule.type === 'dataBar' ? { background: `linear-gradient(90deg, ${cf.rule.color} 60%, transparent 60%)` } : { background: cf.style?.fill, color: cf.style?.color, fontWeight: cf.style?.bold ? 700 : undefined }}>
                {cf.rule.type === 'iconSet' ? '●▲' : 'Aa'}
              </span>
              <span className="min-w-0 flex-1 truncate">{describeRule(cf.rule)}</span>
              <span className="font-mono text-[11.5px] text-[var(--muted)]">{rangeToA1(cf.range)}</span>
              <Button size="icon" aria-label="Move up" disabled={i === 0} onClick={() => onMove(cf.id, -1)}><Icon name="chevronDown" className="rotate-180" size={14} /></Button>
              <Button size="icon" aria-label="Move down" disabled={i === rules.length - 1} onClick={() => onMove(cf.id, 1)}><Icon name="chevronDown" size={14} /></Button>
              <Button size="icon" aria-label="Edit rule" onClick={() => onEdit(cf)}><Icon name="edit" size={14} /></Button>
              <Button size="icon" aria-label="Delete rule" onClick={() => onDelete(cf.id)}><Icon name="trash" size={14} /></Button>
            </div>
          ))}
        </div>
      )}
    </Dialog>
  )
}

// ------------------------------------------------------------------ Sort
export function SortDialog({ wb, sheetId, range, headerGuess, onClose, onSort }: { wb: Workbook; sheetId: string; range: Range; headerGuess: boolean; onClose: () => void; onSort: (keys: SortKey[], hasHeader: boolean) => void }) {
  const [hasHeader, setHasHeader] = useState(headerGuess)
  const [keys, setKeys] = useState<SortKey[]>([{ col: range.c1, desc: false }])
  const colName = (c: number) => (hasHeader ? wb.getDisplay(sheetId, range.r1, c).text || `Column ${colToLetters(c)}` : `Column ${colToLetters(c)}`)
  const cols = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i)
  return (
    <Dialog
      title="Sort"
      onClose={onClose}
      width={520}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" onClick={() => onSort(keys, hasHeader)}>Sort</Button>
        </>
      }
    >
      <div className="mb-2 flex items-center justify-between">
        <span className="text-[12.5px] text-[var(--muted)]">Range {rangeToA1(range)}</span>
        <Check label="My data has headers" checked={hasHeader} onChange={setHasHeader} />
      </div>
      {keys.map((k, i) => (
        <div key={i} className="mb-2 flex items-center gap-2">
          <span className="w-16 text-[12px] text-[var(--muted)]">{i === 0 ? 'Sort by' : 'Then by'}</span>
          <select className={cn(sel, 'flex-1')} value={k.col} onChange={(e) => setKeys((ks) => ks.map((x, j) => (j === i ? { ...x, col: +e.target.value } : x)))} aria-label={`Sort column ${i + 1}`}>
            {cols.map((c) => <option key={c} value={c}>{colName(c)}</option>)}
          </select>
          <select className={sel} value={k.desc ? 'desc' : 'asc'} onChange={(e) => setKeys((ks) => ks.map((x, j) => (j === i ? { ...x, desc: e.target.value === 'desc' } : x)))} aria-label={`Order ${i + 1}`}>
            <option value="asc">A to Z / smallest first</option>
            <option value="desc">Z to A / largest first</option>
          </select>
          <Button size="icon" aria-label="Remove level" disabled={keys.length === 1} onClick={() => setKeys((ks) => ks.filter((_, j) => j !== i))}><Icon name="x" size={14} /></Button>
        </div>
      ))}
      <Button size="sm" variant="outline" disabled={keys.length >= cols.length} onClick={() => setKeys((ks) => [...ks, { col: cols.find((c) => !ks.some((x) => x.col === c)) ?? range.c1, desc: false }])}><Icon name="plus" /> Add level</Button>
    </Dialog>
  )
}

// ------------------------------------------------------------------ Data validation
export function ValidationDialog({ initial, range, onClose, onSave, onClear }: { initial?: Validation; range: Range; onClose: () => void; onSave: (v: Omit<Validation, 'id'>) => void; onClear: () => void }) {
  const [tab, setTab] = useState<'settings' | 'input' | 'error'>('settings')
  const [v, setV] = useState<Omit<Validation, 'id'>>(
    initial ?? { range, type: 'list', op: 'between', list: '', v1: '', v2: '', allowBlank: true, showDropdown: true, errorStyle: 'stop' },
  )
  const up = (p: Partial<Validation>) => setV((x) => ({ ...x, ...p }))
  const numeric = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(v.type)
  return (
    <Dialog
      title="Data validation"
      onClose={onClose}
      width={520}
      footer={
        <>
          <Button variant="ghost" onClick={onClear} className="mr-auto">Clear all</Button>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" onClick={() => onSave({ ...v, range: initial?.range ?? range })}>OK</Button>
        </>
      }
    >
      <div className="mb-3 flex gap-1 border-b border-[var(--border)]">
        {(['settings', 'input', 'error'] as const).map((t) => (
          <button key={t} type="button" onClick={() => setTab(t)} className={cn('-mb-px border-b-2 px-3 py-1.5 text-[12.5px]', tab === t ? 'border-[var(--accent)] font-semibold text-[var(--accent)]' : 'border-transparent')}>
            {t === 'settings' ? 'Settings' : t === 'input' ? 'Input message' : 'Error alert'}
          </button>
        ))}
      </div>
      {tab === 'settings' && (
        <>
          <p className="mb-2 text-[12px] text-[var(--muted)]">Applies to {rangeToA1(initial?.range ?? range)}</p>
          <Field label="Allow">
            <select className={sel} value={v.type} onChange={(e) => up({ type: e.target.value as Validation['type'] })} aria-label="Allow">
              <option value="list">List</option>
              <option value="whole">Whole number</option>
              <option value="decimal">Decimal</option>
              <option value="date">Date</option>
              <option value="time">Time</option>
              <option value="textLength">Text length</option>
              <option value="custom">Custom formula</option>
            </select>
          </Field>
          {v.type === 'list' && (
            <>
              <Field label="Source" hint="Comma-separated items (Open, Closed) or a range such as =$H$2:$H$10">
                <input className={inputClass} value={v.list ?? ''} onChange={(e) => up({ list: e.target.value })} aria-label="Source" />
              </Field>
              <Check label="Show in-cell dropdown" checked={v.showDropdown} onChange={(x) => up({ showDropdown: x })} />
            </>
          )}
          {numeric && (
            <div className="flex flex-wrap items-end gap-2">
              <Field label="Data">
                <select className={sel} value={v.op} onChange={(e) => up({ op: e.target.value as CompareOp })}>
                  <option value="between">between</option>
                  <option value="notBetween">not between</option>
                  <option value="eq">equal to</option>
                  <option value="ne">not equal to</option>
                  <option value="gt">greater than</option>
                  <option value="lt">less than</option>
                  <option value="gte">greater than or equal to</option>
                  <option value="lte">less than or equal to</option>
                </select>
              </Field>
              <Field label={v.op === 'between' || v.op === 'notBetween' ? 'Minimum' : 'Value'}>
                <input className={cn(inputClass, 'w-32')} value={v.v1 ?? ''} onChange={(e) => up({ v1: e.target.value })} placeholder={v.type === 'date' ? '2025-01-01' : ''} />
              </Field>
              {(v.op === 'between' || v.op === 'notBetween') && (
                <Field label="Maximum">
                  <input className={cn(inputClass, 'w-32')} value={v.v2 ?? ''} onChange={(e) => up({ v2: e.target.value })} />
                </Field>
              )}
            </div>
          )}
          {v.type === 'custom' && (
            <Field label="Formula" hint="Must return TRUE for valid entries, relative to the top-left cell of the range.">
              <input className={cn(inputClass, 'font-mono')} value={v.list ?? ''} onChange={(e) => up({ list: e.target.value })} placeholder="=MOD(A1,5)=0" />
            </Field>
          )}
          <Check label="Ignore blank" checked={v.allowBlank} onChange={(x) => up({ allowBlank: x })} />
        </>
      )}
      {tab === 'input' && (
        <>
          <Field label="Title"><input className={inputClass} value={v.inputTitle ?? ''} onChange={(e) => up({ inputTitle: e.target.value || undefined })} /></Field>
          <Field label="Input message" hint="Shown when the cell is selected.">
            <textarea className={cn(inputClass, 'h-20 py-1')} value={v.inputMessage ?? ''} onChange={(e) => up({ inputMessage: e.target.value || undefined })} />
          </Field>
        </>
      )}
      {tab === 'error' && (
        <>
          <Field label="Style">
            <select className={sel} value={v.errorStyle} onChange={(e) => up({ errorStyle: e.target.value as Validation['errorStyle'] })}>
              <option value="stop">Stop (reject the value)</option>
              <option value="warning">Warning (ask to continue)</option>
              <option value="information">Information (allow, with a notice)</option>
            </select>
          </Field>
          <Field label="Title"><input className={inputClass} value={v.errorTitle ?? ''} onChange={(e) => up({ errorTitle: e.target.value || undefined })} /></Field>
          <Field label="Error message">
            <textarea className={cn(inputClass, 'h-20 py-1')} value={v.errorMessage ?? ''} onChange={(e) => up({ errorMessage: e.target.value || undefined })} />
          </Field>
        </>
      )}
    </Dialog>
  )
}

// ------------------------------------------------------------------ Text to columns / remove duplicates / goal seek
export function TextToColumnsDialog({ onClose, onApply }: { onClose: () => void; onApply: (d: string | { widths: number[] }) => void }) {
  const [mode, setMode] = useState<'delim' | 'fixed'>('delim')
  const [delim, setDelim] = useState(',')
  const [other, setOther] = useState('')
  const [widths, setWidths] = useState('5,5')
  return (
    <Dialog
      title="Text to columns"
      onClose={onClose}
      width={420}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" onClick={() => onApply(mode === 'fixed' ? { widths: widths.split(',').map((x) => +x.trim()).filter((x) => x > 0) } : delim === 'other' ? other || ',' : delim)}>Split</Button>
        </>
      }
    >
      <div className="mb-3 flex gap-3 text-[12.5px]">
        <label className="flex items-center gap-1.5"><input type="radio" checked={mode === 'delim'} onChange={() => setMode('delim')} className="accent-[var(--accent)]" />Delimited</label>
        <label className="flex items-center gap-1.5"><input type="radio" checked={mode === 'fixed'} onChange={() => setMode('fixed')} className="accent-[var(--accent)]" />Fixed width</label>
      </div>
      {mode === 'delim' ? (
        <Field label="Delimiter">
          <div className="flex gap-2">
            <select className={sel} value={delim} onChange={(e) => setDelim(e.target.value)} aria-label="Delimiter">
              <option value=",">Comma</option>
              <option value=";">Semicolon</option>
              <option value="\t">Tab</option>
              <option value=" ">Space</option>
              <option value="other">Other…</option>
            </select>
            {delim === 'other' && <input className={cn(inputClass, 'w-16')} maxLength={3} value={other} onChange={(e) => setOther(e.target.value)} aria-label="Other delimiter" />}
          </div>
        </Field>
      ) : (
        <Field label="Column widths (characters)" hint="e.g. 3,8,4 splits after the 3rd and 11th characters">
          <input className={inputClass} value={widths} onChange={(e) => setWidths(e.target.value)} />
        </Field>
      )}
    </Dialog>
  )
}

export function RemoveDuplicatesDialog({ wb, sheetId, range, headerGuess, onClose, onApply }: { wb: Workbook; sheetId: string; range: Range; headerGuess: boolean; onClose: () => void; onApply: (cols: number[], hasHeader: boolean) => void }) {
  const cols = Array.from({ length: range.c2 - range.c1 + 1 }, (_, i) => range.c1 + i)
  const [hasHeader, setHasHeader] = useState(headerGuess)
  const [picked, setPicked] = useState<number[]>(cols)
  return (
    <Dialog
      title="Remove duplicates"
      onClose={onClose}
      width={420}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" disabled={!picked.length} onClick={() => onApply(picked, hasHeader)}>Remove</Button>
        </>
      }
    >
      <p className="mb-2 text-[12.5px] text-[var(--muted)]">Rows in {rangeToA1(range)} that repeat the values in the checked columns are removed (the first occurrence stays).</p>
      <Check label="My data has headers" checked={hasHeader} onChange={setHasHeader} />
      <div className="mt-2 max-h-52 overflow-auto rounded-md border border-[var(--border)] p-1">
        {cols.map((c) => (
          <Check key={c} label={hasHeader ? wb.getDisplay(sheetId, range.r1, c).text || `Column ${colToLetters(c)}` : `Column ${colToLetters(c)}`} checked={picked.includes(c)} onChange={(v) => setPicked((p) => (v ? [...p, c] : p.filter((x) => x !== c)))} />
        ))}
      </div>
    </Dialog>
  )
}

export function GoalSeekDialog({ initialSet, onClose, onRun }: { initialSet: string; onClose: () => void; onRun: (setCell: string, target: number, changing: string) => string | null }) {
  const [setCell, setSetCell] = useState(initialSet)
  const [target, setTarget] = useState('')
  const [changing, setChanging] = useState('')
  const [msg, setMsg] = useState<string | null>(null)
  return (
    <Dialog
      title="Goal Seek"
      onClose={onClose}
      width={380}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Close</Button>
          <Button variant="default" disabled={!setCell || target === '' || !changing} onClick={() => setMsg(onRun(setCell, Number(target), changing))}>Solve</Button>
        </>
      }
    >
      <Field label="Set cell (must contain a formula)"><input className={cn(inputClass, 'font-mono')} value={setCell} onChange={(e) => setSetCell(e.target.value.toUpperCase())} aria-label="Set cell" /></Field>
      <Field label="To value"><input className={inputClass} value={target} onChange={(e) => setTarget(e.target.value)} inputMode="decimal" aria-label="To value" /></Field>
      <Field label="By changing cell"><input className={cn(inputClass, 'font-mono')} value={changing} onChange={(e) => setChanging(e.target.value.toUpperCase())} aria-label="By changing cell" /></Field>
      {msg && <p className="rounded-md bg-[var(--panel)] p-2 text-[12.5px]" role="status">{msg}</p>}
    </Dialog>
  )
}

// ------------------------------------------------------------------ Name manager
export function NameManagerDialog({ wb, onClose, onNew, onChanged }: { wb: Workbook; onClose: () => void; onNew: () => void; onChanged: () => void }) {
  const [editing, setEditing] = useState<string | null>(null)
  const [expr, setExpr] = useState('')
  const [err, setErr] = useState<string | null>(null)
  const names = wb.namedRanges()
  return (
    <Dialog title="Name Manager" onClose={onClose} width={600} footer={<Button variant="default" onClick={onClose}>Close</Button>}>
      <div className="mb-2 flex justify-end">
        <Button size="sm" variant="outline" onClick={onNew}><Icon name="plus" /> New name</Button>
      </div>
      {names.length === 0 ? (
        <p className="rounded-md bg-[var(--panel)] p-4 text-center text-[12.5px] text-[var(--muted)]">No names defined. Select cells and choose New name.</p>
      ) : (
        <div className="max-h-[50vh] overflow-auto rounded-md border border-[var(--border)]">
          {names.map((n) => (
            <div key={n.name} className="flex items-center gap-2 border-t border-[var(--border)] px-2 py-1.5 text-[12.5px] first:border-t-0">
              <span className="w-32 shrink-0 truncate font-medium">{n.name}</span>
              {editing === n.name ? (
                <>
                  <input className={cn(inputClass, 'h-7 flex-1 font-mono')} value={expr} onChange={(e) => (setExpr(e.target.value), setErr(null))} aria-label="Refers to" />
                  <Button size="sm" variant="default" onClick={() => {
                    const e = changeNamedRange(wb, n.name, expr.startsWith('=') ? expr : '=' + expr)
                    if (e) setErr(e)
                    else (setEditing(null), onChanged())
                  }}>Save</Button>
                </>
              ) : (
                <>
                  <span className="flex-1 truncate font-mono text-[var(--muted)]">{n.expression}</span>
                  <Button size="icon" aria-label={`Edit ${n.name}`} onClick={() => (setEditing(n.name), setExpr(n.expression))}><Icon name="edit" size={14} /></Button>
                  <Button size="icon" aria-label={`Delete ${n.name}`} onClick={() => (removeNamedRange(wb, n.name), onChanged())}><Icon name="trash" size={14} /></Button>
                </>
              )}
            </div>
          ))}
        </div>
      )}
      {err && <p className="mt-2 text-[12.5px] text-[var(--danger)]">{err}</p>}
    </Dialog>
  )
}

// ------------------------------------------------------------------ hyperlink & note
export function HyperlinkDialog({ initialUrl, initialText, onClose, onSave, onRemove }: { initialUrl: string; initialText: string; onClose: () => void; onSave: (url: string, text: string) => void; onRemove?: () => void }) {
  const [url, setUrl] = useState(initialUrl)
  const [text, setText] = useState(initialText)
  return (
    <Dialog
      title="Insert hyperlink"
      onClose={onClose}
      width={440}
      footer={
        <>
          {onRemove && <Button variant="ghost" className="mr-auto" onClick={onRemove}>Remove link</Button>}
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" disabled={!url.trim()} onClick={() => onSave(url.trim(), text)}>OK</Button>
        </>
      }
    >
      <Field label="Address" hint="A web address, or #Sheet2!A1 to jump to a cell."><input className={inputClass} value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://" aria-label="Address" /></Field>
      <Field label="Text to display"><input className={inputClass} value={text} onChange={(e) => setText(e.target.value)} aria-label="Text to display" /></Field>
    </Dialog>
  )
}

export function NoteDialog({ initial, onClose, onSave, onDelete }: { initial: string; onClose: () => void; onSave: (t: string) => void; onDelete?: () => void }) {
  const [t, setT] = useState(initial)
  return (
    <Dialog
      title={initial ? 'Edit note' : 'New note'}
      onClose={onClose}
      width={380}
      footer={
        <>
          {onDelete && <Button variant="ghost" className="mr-auto" onClick={onDelete}>Delete note</Button>}
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" disabled={!t.trim()} onClick={() => onSave(t.trim())}>Save</Button>
        </>
      }
    >
      <textarea className={cn(inputClass, 'h-28 w-full py-1.5')} value={t} onChange={(e) => setT(e.target.value)} aria-label="Note" />
    </Dialog>
  )
}

// ------------------------------------------------------------------ page setup
export function PageSetupDialog({ initial, selection, sheetName, onClose, onSave }: { initial: PageSetup | undefined; selection: Range; sheetName: string; onClose: () => void; onSave: (p: PageSetup) => void }) {
  const [p, setP] = useState<PageSetup>({ ...DEFAULT_PAGE_SETUP, ...initial })
  const up = (x: Partial<PageSetup>) => setP((c) => ({ ...c, ...x }))
  void sheetName
  return (
    <Dialog
      title="Page setup"
      onClose={onClose}
      width={520}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" onClick={() => onSave(p)}>OK</Button>
        </>
      }
    >
      <div className="grid grid-cols-2 gap-3">
        <Field label="Orientation">
          <select className={sel} value={p.orientation} onChange={(e) => up({ orientation: e.target.value as PageSetup['orientation'] })}>
            <option value="portrait">Portrait</option>
            <option value="landscape">Landscape</option>
          </select>
        </Field>
        <Field label="Paper size">
          <select className={sel} value={p.paper} onChange={(e) => up({ paper: e.target.value as PageSetup['paper'] })}>
            <option value="A4">A4</option>
            <option value="Letter">Letter</option>
            <option value="Legal">Legal</option>
          </select>
        </Field>
        <Field label="Scaling">
          <select className={sel} value={p.fitToWidth ? 'fit' : 'pct'} onChange={(e) => up({ fitToWidth: e.target.value === 'fit' })}>
            <option value="pct">Adjust to %</option>
            <option value="fit">Fit all columns on one page</option>
          </select>
        </Field>
        <Field label="Scale (%)">
          <input className={inputClass} type="number" min={10} max={400} disabled={p.fitToWidth} value={p.scale} onChange={(e) => up({ scale: Math.max(10, Math.min(400, +e.target.value || 100)) })} />
        </Field>
      </div>
      <div className="mb-3 grid grid-cols-4 gap-2">
        {(['top', 'bottom', 'left', 'right'] as const).map((m) => (
          <Field key={m} label={`${m[0].toUpperCase()}${m.slice(1)} (in)`}>
            <input className={inputClass} type="number" step={0.05} min={0} max={3} value={p.margins[m]} onChange={(e) => up({ margins: { ...p.margins, [m]: Math.max(0, +e.target.value) } })} />
          </Field>
        ))}
      </div>
      <div className="mb-3 flex gap-4">
        <Check label="Print gridlines" checked={p.gridlines} onChange={(v) => up({ gridlines: v })} />
        <Check label="Print row and column headings" checked={p.headings} onChange={(v) => up({ headings: v })} />
      </div>
      <div className="grid grid-cols-2 gap-3">
        <Field label="Header" hint="&P page, &N pages, &D date, &A sheet, &F file"><input className={inputClass} value={p.header ?? ''} onChange={(e) => up({ header: e.target.value || undefined })} /></Field>
        <Field label="Footer"><input className={inputClass} value={p.footer ?? ''} onChange={(e) => up({ footer: e.target.value || undefined })} /></Field>
      </div>
      <div className="flex items-center gap-2 text-[12.5px]">
        <span className="text-[var(--muted)]">Print area:</span>
        <span className="font-mono">{p.printArea ? rangeToA1(p.printArea) : 'Used range'}</span>
        <Button size="sm" variant="outline" onClick={() => up({ printArea: selection })}>Use selection ({rangeToA1(selection)})</Button>
        {p.printArea && <Button size="sm" variant="ghost" onClick={() => up({ printArea: undefined })}>Clear</Button>}
      </div>
    </Dialog>
  )
}

// ------------------------------------------------------------------ create table
export function CreateTableDialog({ range, headerGuess, onClose, onCreate }: { range: Range; headerGuess: boolean; onClose: () => void; onCreate: (range: Range, hasHeader: boolean, style: TableDef['style']) => void }) {
  const [text, setText] = useState(rangeToA1(range))
  const [hasHeader, setHasHeader] = useState(headerGuess)
  const [style, setStyle] = useState<TableDef['style']>('teal')
  const parsed = parseRangeA1(text)
  return (
    <Dialog
      title="Create table"
      onClose={onClose}
      width={420}
      footer={
        <>
          <Button variant="outline" onClick={onClose}>Cancel</Button>
          <Button variant="default" disabled={!parsed} onClick={() => parsed && onCreate(parsed, hasHeader, style)}>OK</Button>
        </>
      }
    >
      <Field label="Where is the data for your table?"><input className={cn(inputClass, 'font-mono')} value={text} onChange={(e) => setText(e.target.value)} aria-label="Table range" /></Field>
      <Check label="My table has headers" checked={hasHeader} onChange={setHasHeader} />
      <div className="mt-3 flex flex-wrap gap-2">
        {(Object.keys(TABLE_STYLES) as TableDef['style'][]).map((k) => (
          <button key={k} type="button" aria-label={`${k} style`} onClick={() => setStyle(k)} className={cn('w-16 overflow-hidden rounded border-2', style === k ? 'border-[var(--accent)]' : 'border-transparent')}>
            <div className="h-3" style={{ background: TABLE_STYLES[k].header }} />
            <div className="h-2" style={{ background: TABLE_STYLES[k].band }} />
            <div className="h-2 bg-white" />
            <div className="h-2" style={{ background: TABLE_STYLES[k].band }} />
          </button>
        ))}
      </div>
    </Dialog>
  )
}

export function useSheetRefText(wb: Workbook, sheetId: string, rg: Range) {
  return useMemo(() => `=${quoteSheetName(wb.meta(sheetId).name)}!$${colToLetters(rg.c1)}$${rg.r1 + 1}:$${colToLetters(rg.c2)}$${rg.r2 + 1}`, [wb, sheetId, rg])
}
