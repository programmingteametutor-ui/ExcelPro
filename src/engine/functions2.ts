/**
 * ExcelPro function plugin, part 2 (Phase 2/3): AVERAGEIFS, XMATCH, LOOKUP, AGGREGATE,
 * FREQUENCY, TAKE, DROP, CHOOSECOLS, CHOOSEROWS, SORTBY.
 *
 * Array-returning functions must tell the engine their result size before evaluation. When a
 * size depends on arguments that aren't literal numbers, the input array's size is used as an
 * upper bound and unused cells of the spill area are left empty.
 */
import { ArraySize, CellError, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue } from 'hyperformula'
import type { ImplementedFunctions } from 'hyperformula'

type Ast = any // eslint-disable-line @typescript-eslint/no-explicit-any
type State = any // eslint-disable-line @typescript-eslint/no-explicit-any
type V = unknown

function literal(ast: Ast): number | undefined {
  if (!ast) return undefined
  if (ast.type === 'NUMBER') return ast.value
  if (ast.type === 'MINUS_UNARY_OP' && ast.value?.type === 'NUMBER') return -ast.value.value
  if (ast.type === 'PLUS_UNARY_OP' && ast.value?.type === 'NUMBER') return ast.value.value
  return undefined
}

const values2d = (r: V): V[][] => (r instanceof SimpleRangeValue ? (r.data as V[][]) : [[r]])

/** Excel criteria: ">5", "<>x", "=a*", "apple" with * and ? wildcards (case-insensitive). */
export function criteriaMatcher(crit: V): (v: V) => boolean {
  if (typeof crit === 'number' || typeof crit === 'boolean') return (v) => v === crit
  const s = String(crit ?? '')
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(s)!
  const op = m[1] ?? '='
  const rhs = m[2]
  const n = rhs.trim() !== '' && !isNaN(Number(rhs)) ? Number(rhs) : null
  if (n !== null && op !== '=' && op !== '<>') {
    return (v) => typeof v === 'number' && (op === '<' ? v < n : op === '>' ? v > n : op === '<=' ? v <= n : v >= n)
  }
  if (n !== null) return (v) => (op === '=' ? v === n || (typeof v === 'string' && v === rhs) : v !== n)
  if (rhs === '') return op === '=' ? (v) => v === null || v === undefined || v === '' : (v) => !(v === null || v === undefined || v === '')
  const re = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0000').replace(/~\?/g, '\u0001').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0000/g, '\\*').replace(/\u0001/g, '\\?') + '$', 'is')
  if (op === '<>') return (v) => !(typeof v === 'string' && re.test(v))
  if (op === '=') return (v) => typeof v === 'string' && re.test(v)
  return (v) => typeof v === 'string' && (op === '<' ? v.toLowerCase() < rhs.toLowerCase() : op === '>' ? v.toLowerCase() > rhs.toLowerCase() : op === '<=' ? v.toLowerCase() <= rhs.toLowerCase() : v.toLowerCase() >= rhs.toLowerCase())
}

const eq = (a: V, b: V) => (typeof a === 'string' && typeof b === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b)
const cmp = (a: V, b: V): number => {
  if (typeof a === 'number' && typeof b === 'number') return a - b
  if (typeof a === 'string' && typeof b === 'string') return a.localeCompare(b, undefined, { sensitivity: 'base' })
  const rank = (x: V) => (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : typeof x === 'boolean' ? 2 : 3)
  return rank(a) - rank(b)
}

function stats(fn: number, nums: number[]): number | CellError {
  const n = nums.length
  const sum = nums.reduce((a, b) => a + b, 0)
  const mean = sum / (n || 1)
  const ss = nums.reduce((a, b) => a + (b - mean) ** 2, 0)
  switch (fn) {
    case 1: return n ? mean : new CellError(ErrorType.DIV_BY_ZERO)
    case 2: return n
    case 4: return n ? Math.max(...nums) : 0
    case 5: return n ? Math.min(...nums) : 0
    case 6: return nums.reduce((a, b) => a * b, 1)
    case 7: return n > 1 ? Math.sqrt(ss / (n - 1)) : new CellError(ErrorType.DIV_BY_ZERO)
    case 8: return n ? Math.sqrt(ss / n) : new CellError(ErrorType.DIV_BY_ZERO)
    case 9: return sum
    case 10: return n > 1 ? ss / (n - 1) : new CellError(ErrorType.DIV_BY_ZERO)
    case 11: return n ? ss / n : new CellError(ErrorType.DIV_BY_ZERO)
    case 12: {
      if (!n) return new CellError(ErrorType.NUM)
      const s = [...nums].sort((a, b) => a - b)
      return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2
    }
    case 13: {
      const counts = new Map<number, number>()
      let best: number | undefined
      let bc = 1
      for (const x of nums) {
        const c = (counts.get(x) ?? 0) + 1
        counts.set(x, c)
        if (c > bc) (bc = c), (best = x)
      }
      return best ?? new CellError(ErrorType.NA)
    }
  }
  return new CellError(ErrorType.VALUE, 'AGGREGATE function number must be 1–19.')
}

export class ExcelProFunctions2 extends FunctionPlugin {
  static implementedFunctions: ImplementedFunctions = {
    AVERAGEIFS: {
      method: 'averageifs',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.NOERROR }],
      repeatLastArgs: 2,
    },
    XMATCH: {
      method: 'xmatch',
      parameters: [
        { argumentType: FunctionArgumentType.NOERROR },
        { argumentType: FunctionArgumentType.RANGE },
        { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 0 },
        { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 1 },
      ],
    },
    LOOKUP: {
      method: 'lookup',
      parameters: [{ argumentType: FunctionArgumentType.NOERROR }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE, optionalArg: true }],
    },
    AGGREGATE: {
      method: 'aggregate',
      parameters: [{ argumentType: FunctionArgumentType.INTEGER }, { argumentType: FunctionArgumentType.INTEGER }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.ANY, optionalArg: true }],
    },
    FREQUENCY: {
      method: 'frequency',
      sizeOfResultArrayMethod: 'frequencySize',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }],
    },
    TAKE: {
      method: 'take',
      sizeOfResultArrayMethod: 'takeSize',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.INTEGER }, { argumentType: FunctionArgumentType.INTEGER, optionalArg: true }],
    },
    DROP: {
      method: 'drop',
      sizeOfResultArrayMethod: 'dropSize',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.INTEGER }, { argumentType: FunctionArgumentType.INTEGER, optionalArg: true }],
    },
    CHOOSECOLS: {
      method: 'choosecols',
      sizeOfResultArrayMethod: 'choosecolsSize',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.INTEGER }],
      repeatLastArgs: 1,
    },
    CHOOSEROWS: {
      method: 'chooserows',
      sizeOfResultArrayMethod: 'chooserowsSize',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.INTEGER }],
      repeatLastArgs: 1,
    },
    SORTBY: {
      method: 'sortby',
      sizeOfResultArrayMethod: 'sameAsFirst',
      parameters: [{ argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.RANGE }, { argumentType: FunctionArgumentType.INTEGER, optionalArg: true, defaultValue: 1 }],
      repeatLastArgs: 2,
    },
  }

  averageifs(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('AVERAGEIFS'), (avg: SimpleRangeValue, ...pairs: V[]) => {
      const vals = avg.data as V[][]
      const conds: { data: V[][]; m: (v: V) => boolean }[] = []
      for (let i = 0; i + 1 < pairs.length; i += 2) {
        const rg = pairs[i] as SimpleRangeValue
        if (rg.width() !== avg.width() || rg.height() !== avg.height()) return new CellError(ErrorType.VALUE, 'Criteria ranges must be the same size as the average range.')
        conds.push({ data: rg.data as V[][], m: criteriaMatcher(pairs[i + 1]) })
      }
      let sum = 0
      let n = 0
      for (let r = 0; r < vals.length; r++)
        for (let c = 0; c < vals[r].length; c++) {
          if (!conds.every((k) => k.m(k.data[r][c]))) continue
          const v = vals[r][c]
          if (typeof v === 'number') {
            sum += v
            n++
          }
        }
      return n ? sum / n : new CellError(ErrorType.DIV_BY_ZERO, 'No cells meet the criteria.')
    })
  }

  xmatch(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('XMATCH'), (look: V, arr: SimpleRangeValue, mode: number, search: number) => {
      const data = arr.data as V[][]
      const list = arr.height() === 1 ? data[0] : data.map((r) => r[0])
      const order = search < 0 ? [...list.keys()].reverse() : [...list.keys()]
      const wild = mode === 2 ? criteriaMatcher(look) : null
      if (mode === 0 || mode === 2) {
        for (const i of order) if (wild ? wild(list[i]) : eq(list[i], look)) return i + 1
        return new CellError(ErrorType.NA, 'Value not found.')
      }
      let best = -1
      for (const i of order) {
        const c = cmp(list[i], look)
        if (c === 0) return i + 1
        if (mode === -1 && c < 0 && (best < 0 || cmp(list[i], list[best]) > 0)) best = i
        if (mode === 1 && c > 0 && (best < 0 || cmp(list[i], list[best]) < 0)) best = i
      }
      return best >= 0 ? best + 1 : new CellError(ErrorType.NA, 'Value not found.')
    })
  }

  lookup(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('LOOKUP'), (look: V, vec: SimpleRangeValue, res?: SimpleRangeValue) => {
      const d = vec.data as V[][]
      let keys: V[]
      let outs: V[]
      if (res) {
        keys = vec.height() === 1 ? d[0] : d.map((r) => r[0])
        const rd = res.data as V[][]
        outs = res.height() === 1 ? rd[0] : rd.map((r) => r[0])
      } else if (vec.width() > vec.height()) {
        keys = d[0]
        outs = d[d.length - 1]
      } else {
        keys = d.map((r) => r[0])
        outs = d.map((r) => r[r.length - 1])
      }
      let found = -1
      for (let i = 0; i < keys.length; i++) {
        if (keys[i] === null || keys[i] === undefined) continue
        if (typeof keys[i] !== typeof look) continue
        if (cmp(keys[i], look) <= 0) found = i
        else break
      }
      return found >= 0 ? (outs[found] as never) ?? new CellError(ErrorType.NA) : new CellError(ErrorType.NA, 'No value is less than or equal to the lookup value.')
    })
  }

  aggregate(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('AGGREGATE'), (fn: number, opt: number, ref: SimpleRangeValue, k?: V) => {
      const ignoreErrors = [2, 3, 6, 7].includes(opt)
      const nums: number[] = []
      let countA = 0
      for (const row of ref.data as V[][])
        for (const v of row) {
          if (v instanceof CellError) {
            if (ignoreErrors) continue
            return v
          }
          if (v !== null && v !== undefined && v !== '') countA++
          if (typeof v === 'number') nums.push(v)
        }
      if (fn === 3) return countA
      if (fn >= 14 && fn <= 19) {
        const kk = typeof k === 'number' ? k : NaN
        const s = [...nums].sort((a, b) => a - b)
        if (fn === 14) return s[s.length - kk] ?? new CellError(ErrorType.NUM)
        if (fn === 15) return s[kk - 1] ?? new CellError(ErrorType.NUM)
        const p = fn === 17 || fn === 19 ? kk / 4 : kk
        if (!(p >= 0 && p <= 1) || !s.length) return new CellError(ErrorType.NUM)
        const idx = (s.length - 1) * p
        const lo = Math.floor(idx)
        return s[lo] + (s[Math.ceil(idx)] - s[lo]) * (idx - lo)
      }
      return stats(fn, nums)
    })
  }

  frequency(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('FREQUENCY'), (data: SimpleRangeValue, bins: SimpleRangeValue) => {
      const nums = (data.data as V[][]).flat().filter((v): v is number => typeof v === 'number')
      const b = (bins.data as V[][]).flat().filter((v): v is number => typeof v === 'number')
      const sorted = b.map((x, i) => ({ x, i })).sort((p, q) => p.x - q.x)
      const counts = new Array(b.length + 1).fill(0)
      for (const v of nums) {
        const slot = sorted.find((s) => v <= s.x)
        counts[slot ? slot.i : b.length]++
      }
      return SimpleRangeValue.onlyValues(counts.map((c) => [c]))
    })
  }
  frequencySize(ast: Ast, state: State) {
    const bins = this.arraySizeForAst(ast.args[1], state)
    return new ArraySize(1, bins.width * bins.height + 1)
  }

  take(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('TAKE'), (arr: SimpleRangeValue, rows: number, cols?: number) => {
      let d = arr.data as V[][]
      if (rows === 0 || cols === 0) return new CellError(ErrorType.VALUE, 'TAKE cannot return an empty array.')
      d = rows > 0 ? d.slice(0, rows) : d.slice(rows)
      if (cols !== undefined) d = d.map((r) => (cols > 0 ? r.slice(0, cols) : r.slice(cols)))
      return SimpleRangeValue.onlyValues(d as never)
    })
  }
  takeSize(ast: Ast, state: State) {
    const s = this.arraySizeForAst(ast.args[0], state)
    const r = literal(ast.args[1])
    const c = literal(ast.args[2])
    return new ArraySize(c === undefined ? s.width : Math.min(s.width, Math.abs(c)), r === undefined ? s.height : Math.min(s.height, Math.abs(r)))
  }

  drop(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('DROP'), (arr: SimpleRangeValue, rows: number, cols?: number) => {
      let d = arr.data as V[][]
      d = rows >= 0 ? d.slice(rows) : d.slice(0, rows)
      if (cols !== undefined) d = d.map((r) => (cols >= 0 ? r.slice(cols) : r.slice(0, cols)))
      if (!d.length || !d[0].length) return new CellError(ErrorType.VALUE, 'DROP removed every row or column.')
      return SimpleRangeValue.onlyValues(d as never)
    })
  }
  dropSize(ast: Ast, state: State) {
    const s = this.arraySizeForAst(ast.args[0], state)
    const r = literal(ast.args[1])
    const c = literal(ast.args[2])
    return new ArraySize(Math.max(1, c === undefined ? s.width : s.width - Math.abs(c)), Math.max(1, r === undefined ? s.height : s.height - Math.abs(r)))
  }

  choosecols(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('CHOOSECOLS'), (arr: SimpleRangeValue, ...idx: number[]) => {
      const d = arr.data as V[][]
      const w = arr.width()
      const pick = idx.map((i) => (i < 0 ? w + i : i - 1))
      if (pick.some((i) => i < 0 || i >= w)) return new CellError(ErrorType.VALUE, 'Column index out of range.')
      return SimpleRangeValue.onlyValues(d.map((r) => pick.map((i) => r[i])) as never)
    })
  }
  choosecolsSize(ast: Ast, state: State) {
    const s = this.arraySizeForAst(ast.args[0], state)
    return new ArraySize(ast.args.length - 1, s.height)
  }

  chooserows(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('CHOOSEROWS'), (arr: SimpleRangeValue, ...idx: number[]) => {
      const d = arr.data as V[][]
      const h = arr.height()
      const pick = idx.map((i) => (i < 0 ? h + i : i - 1))
      if (pick.some((i) => i < 0 || i >= h)) return new CellError(ErrorType.VALUE, 'Row index out of range.')
      return SimpleRangeValue.onlyValues(pick.map((i) => d[i]) as never)
    })
  }
  chooserowsSize(ast: Ast, state: State) {
    const s = this.arraySizeForAst(ast.args[0], state)
    return new ArraySize(s.width, ast.args.length - 1)
  }

  sortby(ast: Ast, state: State) {
    return this.runFunction(ast.args, state, this.metadata('SORTBY'), (arr: SimpleRangeValue, ...rest: V[]) => {
      const d = arr.data as V[][]
      const keys: { vals: V[]; dir: number }[] = []
      for (let i = 0; i < rest.length; i += 2) {
        const by = rest[i] as SimpleRangeValue
        const bd = by.data as V[][]
        const vals = by.width() === 1 ? bd.map((r) => r[0]) : bd[0]
        if (vals.length !== d.length) return new CellError(ErrorType.VALUE, 'SORTBY arrays must have the same number of rows.')
        keys.push({ vals, dir: (rest[i + 1] as number | undefined) === -1 ? -1 : 1 })
      }
      const order = d.map((_, i) => i).sort((a, b) => {
        for (const k of keys) {
          const c = cmp(k.vals[a], k.vals[b])
          if (c) return c * k.dir
        }
        return a - b
      })
      return SimpleRangeValue.onlyValues(order.map((i) => d[i]) as never)
    })
  }
  sameAsFirst(ast: Ast, state: State) {
    return this.arraySizeForAst(ast.args[0], state)
  }
}

void values2d

let registered = false
export function registerExcelProFunctions2() {
  if (registered) return
  registered = true
  const names = Object.keys(ExcelProFunctions2.implementedFunctions)
  HyperFormula.registerFunctionPlugin(ExcelProFunctions2, { enGB: Object.fromEntries(names.map((n) => [n, n])) })
}
