/**
 * ExcelPro function plugin for HyperFormula: fills gaps in the engine's Excel compatibility.
 * - TEXT is replaced so it uses the same Excel number-format renderer as cell display.
 * - Adds CONCAT, NUMBERVALUE, TYPE, ERROR.TYPE, RANK.EQ/RANK, MODE.SNGL/MODE.
 */
import { CellError, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue } from 'hyperformula'
import type { ImplementedFunctions } from 'hyperformula'
import { formatNumber } from './numberFormat'
import { registerExcelProFunctions2 } from './functions2'

// HyperFormula does not export its AST/state types from the package root.
type Ast = any // eslint-disable-line @typescript-eslint/no-explicit-any
type State = any // eslint-disable-line @typescript-eslint/no-explicit-any

function flatten(v: unknown): unknown[] {
  if (v instanceof SimpleRangeValue) return v.valuesFromTopLeftCorner()
  return [v]
}

export class ExcelProFunctions extends FunctionPlugin {
  static implementedFunctions: ImplementedFunctions = {
    TEXT: {
      method: 'text',
      parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.STRING }],
    },
    CONCAT: {
      method: 'concat',
      parameters: [{ argumentType: FunctionArgumentType.STRING }],
      repeatLastArgs: 1,
      expandRanges: true,
    },
    NUMBERVALUE: {
      method: 'numbervalue',
      parameters: [
        { argumentType: FunctionArgumentType.STRING },
        { argumentType: FunctionArgumentType.STRING, optionalArg: true, defaultValue: '.' },
        { argumentType: FunctionArgumentType.STRING, optionalArg: true, defaultValue: ',' },
      ],
    },
    TYPE: {
      method: 'type',
      parameters: [{ argumentType: FunctionArgumentType.ANY }],
    },
    'ERROR.TYPE': {
      method: 'errortype',
      parameters: [{ argumentType: FunctionArgumentType.SCALAR }],
    },
    'RANK.EQ': {
      method: 'rankeq',
      parameters: [
        { argumentType: FunctionArgumentType.NUMBER },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0 },
      ],
    },
    'MODE.SNGL': {
      method: 'modesngl',
      parameters: [{ argumentType: FunctionArgumentType.ANY }],
      repeatLastArgs: 1,
    },
  }

  static aliases = { RANK: 'RANK.EQ', MODE: 'MODE.SNGL' }

  text(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('TEXT'), (n: number, fmt: string) => formatNumber(n, fmt).text)
  }

  concat(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('CONCAT'), (...args: string[]) => args.join(''))
  }

  numbervalue(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('NUMBERVALUE'), (text: string, dec: string, group: string) => {
      let s = text.replace(/\s+/g, '')
      if (group) s = s.split(group[0]).join('')
      if (dec && dec[0] !== '.') s = s.replace(dec[0], '.')
      let pct = 0
      while (s.endsWith('%')) {
        pct++
        s = s.slice(0, -1)
      }
      if (s === '') return 0
      const n = Number(s)
      if (isNaN(n)) return new CellError(ErrorType.VALUE, 'NUMBERVALUE could not parse the text.')
      return n / Math.pow(100, pct)
    })
  }

  type(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('TYPE'), (v: unknown) => {
      if (v instanceof SimpleRangeValue) return v.width() * v.height() > 1 ? 64 : this.typeOf(v.valuesFromTopLeftCorner()[0])
      return this.typeOf(v)
    })
  }

  private typeOf(v: unknown): number {
    if (v instanceof CellError) return 16
    if (typeof v === 'boolean') return 4
    if (typeof v === 'string') return 2
    return 1
  }

  errortype(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('ERROR.TYPE'), (v: unknown) => {
      if (!(v instanceof CellError)) return new CellError(ErrorType.NA)
      const map: Partial<Record<ErrorType, number>> = {
        [ErrorType.DIV_BY_ZERO]: 2, [ErrorType.VALUE]: 3, [ErrorType.REF]: 4,
        [ErrorType.NAME]: 5, [ErrorType.NUM]: 6, [ErrorType.NA]: 7,
      }
      return map[v.type] ?? new CellError(ErrorType.NA)
    })
  }

  rankeq(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('RANK.EQ'), (n: number, range: SimpleRangeValue, order: number) => {
      const vals = range.valuesFromTopLeftCorner().filter((x): x is number => typeof x === 'number')
      if (!vals.includes(n)) return new CellError(ErrorType.NA, 'Value not found in the reference list.')
      const better = order ? vals.filter((x) => x < n).length : vals.filter((x) => x > n).length
      return better + 1
    })
  }

  modesngl(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('MODE.SNGL'), (...args: unknown[]) => {
      const counts = new Map<number, number>()
      const order: number[] = []
      for (const a of args) {
        for (const v of flatten(a)) {
          if (v instanceof CellError) return v
          if (typeof v !== 'number') continue
          if (!counts.has(v)) order.push(v)
          counts.set(v, (counts.get(v) ?? 0) + 1)
        }
      }
      let best: number | undefined
      let bestCount = 1
      for (const v of order) {
        const c = counts.get(v)!
        if (c > bestCount) {
          best = v
          bestCount = c
        }
      }
      return best === undefined ? new CellError(ErrorType.NA, 'No value repeats.') : best
    })
  }
}

let registered = false
/** Registers ExcelPro's functions once per process (HyperFormula registry is global). */
export function registerExcelProFunctions() {
  if (registered) return
  registered = true
  try {
    HyperFormula.unregisterFunction('TEXT')
  } catch {
    /* not registered */
  }
  const names = ['TEXT', 'CONCAT', 'NUMBERVALUE', 'TYPE', 'ERROR.TYPE', 'RANK.EQ', 'MODE.SNGL', 'RANK', 'MODE']
  const translations = { enGB: Object.fromEntries(names.map((n) => [n, n])) }
  HyperFormula.registerFunctionPlugin(ExcelProFunctions, translations)
  registerExcelProFunctions2()
}
