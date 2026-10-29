/**
 * Function compatibility registry. Status is derived from what the formula engine actually has
 * registered at runtime (HyperFormula + the ExcelPro plugin), plus hand-maintained notes for
 * functions that work but differ from Excel in documented ways.
 */
import { HyperFormula } from 'hyperformula'
import { registerExcelProFunctions } from './functions'

registerExcelProFunctions()

export type CompatStatus = 'supported' | 'partial' | 'unsupported'

export const SPEC_FUNCTIONS: Record<string, string[]> = {
  'Math & Trig': 'SUM SUMIF SUMIFS PRODUCT ABS ROUND ROUNDUP ROUNDDOWN INT MOD POWER SQRT SIGN RAND RANDBETWEEN SUBTOTAL AGGREGATE CEILING FLOOR TRUNC PI SIN COS TAN LOG LOG10 EXP FACT'.split(' '),
  Statistical: 'AVERAGE AVERAGEIF AVERAGEIFS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS MAX MIN MEDIAN MODE.SNGL STDEV.S STDEV.P VAR.S VAR.P LARGE SMALL RANK.EQ PERCENTILE.INC QUARTILE.INC CORREL FREQUENCY'.split(' '),
  Logical: 'IF IFS AND OR NOT XOR IFERROR IFNA TRUE FALSE SWITCH LET'.split(' '),
  Text: 'LEFT RIGHT MID LEN TRIM CLEAN UPPER LOWER PROPER CONCAT CONCATENATE TEXTJOIN FIND SEARCH SUBSTITUTE REPLACE EXACT TEXT VALUE NUMBERVALUE CHAR CODE REPT TEXTSPLIT'.split(' '),
  'Date & Time': 'TODAY NOW DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND DAYS DATEDIF EDATE EOMONTH WEEKDAY WEEKNUM WORKDAY NETWORKDAYS YEARFRAC'.split(' '),
  'Lookup & Reference': 'VLOOKUP HLOOKUP XLOOKUP LOOKUP INDEX MATCH XMATCH CHOOSE ROW ROWS COLUMN COLUMNS ADDRESS INDIRECT OFFSET TRANSPOSE FILTER SORT SORTBY UNIQUE TAKE DROP CHOOSECOLS CHOOSEROWS'.split(' '),
  Financial: 'PMT PV FV NPV IRR RATE NPER IPMT PPMT SLN DB DDB EFFECT'.split(' '),
  Information: 'ISBLANK ISNUMBER ISTEXT ISLOGICAL ISERROR ISERR ISNA TYPE CELL ERROR.TYPE'.split(' '),
  Database: 'DSUM DAVERAGE DCOUNT DCOUNTA DMAX DMIN DGET DPRODUCT DSTDEV DVAR'.split(' '),
  Advanced: 'LAMBDA SEQUENCE'.split(' '),
}

/** Functions handled outside the function registry (parser-level). */
const PARSER_LEVEL = new Set(['OFFSET'])

export const PARTIAL_NOTES: Record<string, string> = {
  OFFSET: 'Supported when height/width arguments are constants; evaluated by the parser, so arguments cannot be other formulas returning references.',
  TEXT: 'Uses ExcelPro\'s number-format renderer (common codes: 0 # , . % E+ dates/times, [Color], "literals"). Locale-specific codes are not supported.',
  CEILING: 'Follows Excel 2010+ semantics; negative-number edge cases may differ from Excel.',
  FLOOR: 'Follows Excel 2010+ semantics; negative-number edge cases may differ from Excel.',
  SUBTOTAL: 'Hidden rows are not excluded for function numbers 101–111 (row hiding is not visible to the engine).',
  RAND: 'Recalculates on every edit (volatile), like Excel.',
  NOW: 'Uses the browser clock and time zone.',
  TODAY: 'Uses the browser clock and time zone.',
  FILTER: 'Dynamic-array spill supported; results cannot overlap non-empty cells (#SPILL!).',
  SORT: 'Dynamic-array spill supported; results cannot overlap non-empty cells (#SPILL!).',
  UNIQUE: 'Dynamic-array spill supported; results cannot overlap non-empty cells (#SPILL!).',
  SEQUENCE: 'Dynamic-array spill supported.',
  TRANSPOSE: 'Returns a spilled array.',
  'RANK.EQ': 'Implemented by ExcelPro (RANK is an alias).',
  AGGREGATE: 'Implemented by ExcelPro. Options that ignore errors work; options that ignore hidden rows are treated like their non-hidden equivalents.',
  TAKE: 'Implemented by ExcelPro. When the row/column counts are cell references rather than numbers, the spill area is sized to the whole input.',
  DROP: 'Implemented by ExcelPro. When the row/column counts are cell references rather than numbers, the spill area is sized to the whole input.',
  XMATCH: 'Implemented by ExcelPro. Binary search modes (±2) are treated as linear search.',
  AVERAGEIFS: 'Implemented by ExcelPro.',
  LOOKUP: 'Implemented by ExcelPro (vector and array forms).',
  FREQUENCY: 'Implemented by ExcelPro (spills).',
  SORTBY: 'Implemented by ExcelPro (spills).',
  CHOOSECOLS: 'Implemented by ExcelPro (spills).',
  CHOOSEROWS: 'Implemented by ExcelPro (spills).',
  'MODE.SNGL': 'Implemented by ExcelPro (MODE is an alias).',
  CONCAT: 'Implemented by ExcelPro.',
  NUMBERVALUE: 'Implemented by ExcelPro.',
  TYPE: 'Implemented by ExcelPro.',
  'ERROR.TYPE': 'Implemented by ExcelPro. #NULL! is not produced by the engine.',
  TRUE: 'Typed as TRUE; stored internally as TRUE() and converted back for display and export.',
  FALSE: 'Typed as FALSE; stored internally as FALSE() and converted back for display and export.',
}

/** Functions that work identically enough that the note is informational, not a limitation. */
const INFO_ONLY = new Set(['AVERAGEIFS', 'LOOKUP', 'FREQUENCY', 'SORTBY', 'CHOOSECOLS', 'CHOOSEROWS', 'RANK.EQ', 'MODE.SNGL', 'CONCAT', 'NUMBERVALUE', 'TYPE', 'RAND', 'TRUE', 'FALSE', 'SEQUENCE', 'TRANSPOSE'])

export interface CompatEntry {
  name: string
  category: string
  status: CompatStatus
  note?: string
}

let cached: Set<string> | null = null
export function engineFunctions(): Set<string> {
  if (!cached) cached = new Set([...HyperFormula.getRegisteredFunctionNames('enGB'), ...PARSER_LEVEL])
  return cached
}

export function compatOf(name: string, category = ''): CompatEntry {
  const up = name.toUpperCase()
  const has = engineFunctions().has(up)
  if (!has) return { name: up, category, status: 'unsupported', note: 'Not available in the formula engine. Cells using it show #NAME? and keep their formula.' }
  const note = PARTIAL_NOTES[up]
  const status: CompatStatus = note && !INFO_ONLY.has(up) ? 'partial' : 'supported'
  return { name: up, category, status, note }
}

export function compatibilityMatrix(): CompatEntry[] {
  return Object.entries(SPEC_FUNCTIONS).flatMap(([cat, names]) => names.map((n) => compatOf(n, cat)))
}

/** Short help for the Insert Function dialog (common functions). */
export const FUNCTION_HELP: Record<string, { syntax: string; description: string; category: string }> = {
  SUM: { syntax: 'SUM(number1, [number2], ...)', description: 'Adds numbers and cell ranges.', category: 'Math & Trig' },
  SUMIF: { syntax: 'SUMIF(range, criteria, [sum_range])', description: 'Adds cells that meet one condition.', category: 'Math & Trig' },
  SUMIFS: { syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, ...)', description: 'Adds cells that meet several conditions.', category: 'Math & Trig' },
  ROUND: { syntax: 'ROUND(number, num_digits)', description: 'Rounds a number to a number of digits.', category: 'Math & Trig' },
  ABS: { syntax: 'ABS(number)', description: 'Absolute value of a number.', category: 'Math & Trig' },
  PRODUCT: { syntax: 'PRODUCT(number1, [number2], ...)', description: 'Multiplies its arguments.', category: 'Math & Trig' },
  MOD: { syntax: 'MOD(number, divisor)', description: 'Remainder after division.', category: 'Math & Trig' },
  AVERAGE: { syntax: 'AVERAGE(number1, [number2], ...)', description: 'Calculates the average (arithmetic mean).', category: 'Statistical' },
  AVERAGEIF: { syntax: 'AVERAGEIF(range, criteria, [average_range])', description: 'Average of cells that meet a condition.', category: 'Statistical' },
  COUNT: { syntax: 'COUNT(value1, [value2], ...)', description: 'Counts cells that contain numbers.', category: 'Statistical' },
  COUNTA: { syntax: 'COUNTA(value1, [value2], ...)', description: 'Counts cells that are not empty.', category: 'Statistical' },
  COUNTBLANK: { syntax: 'COUNTBLANK(range)', description: 'Counts empty cells.', category: 'Statistical' },
  COUNTIF: { syntax: 'COUNTIF(range, criteria)', description: 'Counts cells that meet a condition (e.g. count completed tasks).', category: 'Statistical' },
  COUNTIFS: { syntax: 'COUNTIFS(criteria_range1, criteria1, ...)', description: 'Counts cells that meet several conditions.', category: 'Statistical' },
  MAX: { syntax: 'MAX(number1, [number2], ...)', description: 'Largest value.', category: 'Statistical' },
  MIN: { syntax: 'MIN(number1, [number2], ...)', description: 'Smallest value.', category: 'Statistical' },
  MEDIAN: { syntax: 'MEDIAN(number1, [number2], ...)', description: 'Middle value.', category: 'Statistical' },
  LARGE: { syntax: 'LARGE(array, k)', description: 'k-th largest value.', category: 'Statistical' },
  SMALL: { syntax: 'SMALL(array, k)', description: 'k-th smallest value.', category: 'Statistical' },
  IF: { syntax: 'IF(logical_test, value_if_true, [value_if_false])', description: 'Returns one value if a condition is true and another if false.', category: 'Logical' },
  IFS: { syntax: 'IFS(test1, value1, [test2, value2], ...)', description: 'Returns the value for the first true condition.', category: 'Logical' },
  AND: { syntax: 'AND(logical1, [logical2], ...)', description: 'TRUE if all arguments are true.', category: 'Logical' },
  OR: { syntax: 'OR(logical1, [logical2], ...)', description: 'TRUE if any argument is true.', category: 'Logical' },
  NOT: { syntax: 'NOT(logical)', description: 'Reverses a logical value.', category: 'Logical' },
  IFERROR: { syntax: 'IFERROR(value, value_if_error)', description: 'Returns a fallback when a formula errors.', category: 'Logical' },
  IFNA: { syntax: 'IFNA(value, value_if_na)', description: 'Returns a fallback when a formula returns #N/A.', category: 'Logical' },
  SWITCH: { syntax: 'SWITCH(expression, value1, result1, ..., [default])', description: 'Matches a value against a list of cases.', category: 'Logical' },
  LEFT: { syntax: 'LEFT(text, [num_chars])', description: 'First characters of a text.', category: 'Text' },
  RIGHT: { syntax: 'RIGHT(text, [num_chars])', description: 'Last characters of a text.', category: 'Text' },
  MID: { syntax: 'MID(text, start_num, num_chars)', description: 'Characters from the middle of a text.', category: 'Text' },
  LEN: { syntax: 'LEN(text)', description: 'Number of characters.', category: 'Text' },
  TRIM: { syntax: 'TRIM(text)', description: 'Removes extra spaces.', category: 'Text' },
  UPPER: { syntax: 'UPPER(text)', description: 'Converts text to upper case.', category: 'Text' },
  LOWER: { syntax: 'LOWER(text)', description: 'Converts text to lower case.', category: 'Text' },
  PROPER: { syntax: 'PROPER(text)', description: 'Capitalizes each word.', category: 'Text' },
  CONCAT: { syntax: 'CONCAT(text1, [text2], ...)', description: 'Joins text from several values or ranges.', category: 'Text' },
  TEXTJOIN: { syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, ...)', description: 'Joins text with a delimiter.', category: 'Text' },
  SUBSTITUTE: { syntax: 'SUBSTITUTE(text, old_text, new_text, [instance_num])', description: 'Replaces text by matching.', category: 'Text' },
  TEXT: { syntax: 'TEXT(value, format_text)', description: 'Formats a number as text using a number format.', category: 'Text' },
  VALUE: { syntax: 'VALUE(text)', description: 'Converts text to a number.', category: 'Text' },
  FIND: { syntax: 'FIND(find_text, within_text, [start_num])', description: 'Position of text (case-sensitive).', category: 'Text' },
  SEARCH: { syntax: 'SEARCH(find_text, within_text, [start_num])', description: 'Position of text (case-insensitive, wildcards).', category: 'Text' },
  TODAY: { syntax: 'TODAY()', description: "Today's date.", category: 'Date & Time' },
  NOW: { syntax: 'NOW()', description: 'Current date and time.', category: 'Date & Time' },
  DATE: { syntax: 'DATE(year, month, day)', description: 'Builds a date from parts.', category: 'Date & Time' },
  YEAR: { syntax: 'YEAR(serial_number)', description: 'Year of a date.', category: 'Date & Time' },
  MONTH: { syntax: 'MONTH(serial_number)', description: 'Month of a date.', category: 'Date & Time' },
  DAY: { syntax: 'DAY(serial_number)', description: 'Day of a date.', category: 'Date & Time' },
  EOMONTH: { syntax: 'EOMONTH(start_date, months)', description: 'Last day of the month, n months away.', category: 'Date & Time' },
  NETWORKDAYS: { syntax: 'NETWORKDAYS(start_date, end_date, [holidays])', description: 'Working days between two dates.', category: 'Date & Time' },
  VLOOKUP: { syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', description: 'Looks up a value in the first column of a table (e.g. find a matching employee ID).', category: 'Lookup & Reference' },
  HLOOKUP: { syntax: 'HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])', description: 'Looks up a value in the first row of a table.', category: 'Lookup & Reference' },
  XLOOKUP: { syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])', description: 'Finds a match and returns the corresponding item.', category: 'Lookup & Reference' },
  INDEX: { syntax: 'INDEX(array, row_num, [column_num])', description: 'Value at a position in a range.', category: 'Lookup & Reference' },
  MATCH: { syntax: 'MATCH(lookup_value, lookup_array, [match_type])', description: 'Position of a value in a range.', category: 'Lookup & Reference' },
  CHOOSE: { syntax: 'CHOOSE(index_num, value1, [value2], ...)', description: 'Picks a value from a list by position.', category: 'Lookup & Reference' },
  FILTER: { syntax: 'FILTER(array, include, [if_empty])', description: 'Filters a range by a condition (spills).', category: 'Lookup & Reference' },
  UNIQUE: { syntax: 'UNIQUE(array)', description: 'Unique values of a range (spills).', category: 'Lookup & Reference' },
  SORT: { syntax: 'SORT(array, [sort_index], [sort_order])', description: 'Sorts a range (spills).', category: 'Lookup & Reference' },
  PMT: { syntax: 'PMT(rate, nper, pv, [fv], [type])', description: 'Loan payment for constant payments and rate.', category: 'Financial' },
  FV: { syntax: 'FV(rate, nper, pmt, [pv], [type])', description: 'Future value of an investment.', category: 'Financial' },
  PV: { syntax: 'PV(rate, nper, pmt, [fv], [type])', description: 'Present value of an investment.', category: 'Financial' },
  NPV: { syntax: 'NPV(rate, value1, [value2], ...)', description: 'Net present value of cash flows.', category: 'Financial' },
  IRR: { syntax: 'IRR(values, [guess])', description: 'Internal rate of return.', category: 'Financial' },
  ISBLANK: { syntax: 'ISBLANK(value)', description: 'TRUE if the cell is empty.', category: 'Information' },
  ISNUMBER: { syntax: 'ISNUMBER(value)', description: 'TRUE if the value is a number.', category: 'Information' },
  ISERROR: { syntax: 'ISERROR(value)', description: 'TRUE if the value is any error.', category: 'Information' },
}

/** Additional syntax help (name|arguments|description) used by autocomplete and Insert Function. */
const MORE_HELP = `
SUMPRODUCT|array1, [array2], ...|Multiplies ranges element by element and adds the results.|Math & Trig
ROUNDUP|number, num_digits|Rounds a number up, away from zero.|Math & Trig
ROUNDDOWN|number, num_digits|Rounds a number down, toward zero.|Math & Trig
INT|number|Rounds a number down to the nearest integer.|Math & Trig
POWER|number, power|Raises a number to a power.|Math & Trig
SQRT|number|Square root.|Math & Trig
SIGN|number|Sign of a number: 1, 0 or -1.|Math & Trig
RAND||Random number between 0 and 1.|Math & Trig
RANDBETWEEN|bottom, top|Random integer between two numbers.|Math & Trig
SUBTOTAL|function_num, ref1, [ref2], ...|Subtotal using SUM, AVERAGE, COUNT… by function number.|Math & Trig
AGGREGATE|function_num, options, ref1, [k]|Aggregate that can ignore errors (options 2, 3, 6, 7).|Math & Trig
CEILING|number, significance|Rounds up to the nearest multiple.|Math & Trig
FLOOR|number, significance|Rounds down to the nearest multiple.|Math & Trig
TRUNC|number, [num_digits]|Truncates a number.|Math & Trig
PI||The number π.|Math & Trig
SIN|number|Sine of an angle in radians.|Math & Trig
COS|number|Cosine of an angle in radians.|Math & Trig
TAN|number|Tangent of an angle in radians.|Math & Trig
LOG|number, [base]|Logarithm to a base (default 10).|Math & Trig
LOG10|number|Base-10 logarithm.|Math & Trig
LN|number|Natural logarithm.|Math & Trig
EXP|number|e raised to a power.|Math & Trig
FACT|number|Factorial.|Math & Trig
AVERAGEIFS|average_range, criteria_range1, criteria1, ...|Average of cells that meet several conditions.|Statistical
MODE.SNGL|number1, [number2], ...|Most frequent value.|Statistical
STDEV.S|number1, [number2], ...|Standard deviation of a sample.|Statistical
STDEV.P|number1, [number2], ...|Standard deviation of a population.|Statistical
VAR.S|number1, [number2], ...|Variance of a sample.|Statistical
VAR.P|number1, [number2], ...|Variance of a population.|Statistical
RANK.EQ|number, ref, [order]|Rank of a number in a list.|Statistical
PERCENTILE.INC|array, k|k-th percentile (inclusive).|Statistical
QUARTILE.INC|array, quart|Quartile (inclusive).|Statistical
CORREL|array1, array2|Correlation coefficient.|Statistical
FREQUENCY|data_array, bins_array|Counts values in each bin (spills).|Statistical
MAXIFS|max_range, criteria_range1, criteria1, ...|Largest value meeting conditions.|Statistical
MINIFS|min_range, criteria_range1, criteria1, ...|Smallest value meeting conditions.|Statistical
XOR|logical1, [logical2], ...|TRUE if an odd number of arguments are true.|Logical
TRUE||The logical value TRUE.|Logical
FALSE||The logical value FALSE.|Logical
CLEAN|text|Removes non-printable characters.|Text
CONCATENATE|text1, [text2], ...|Joins text items.|Text
REPLACE|old_text, start_num, num_chars, new_text|Replaces part of a text by position.|Text
EXACT|text1, text2|TRUE if two texts are identical (case-sensitive).|Text
NUMBERVALUE|text, [decimal_separator], [group_separator]|Converts text to a number, locale-independent.|Text
CHAR|number|Character for a code.|Text
CODE|text|Code of the first character.|Text
REPT|text, number_times|Repeats text.|Text
TIME|hour, minute, second|Builds a time.|Date & Time
HOUR|serial_number|Hour of a time.|Date & Time
MINUTE|serial_number|Minute of a time.|Date & Time
SECOND|serial_number|Second of a time.|Date & Time
DAYS|end_date, start_date|Days between two dates.|Date & Time
DATEDIF|start_date, end_date, unit|Difference between dates in "Y", "M", "D"…|Date & Time
EDATE|start_date, months|Date n months away.|Date & Time
WEEKDAY|serial_number, [return_type]|Day of the week.|Date & Time
WEEKNUM|serial_number, [return_type]|Week number of the year.|Date & Time
WORKDAY|start_date, days, [holidays]|Date n working days away.|Date & Time
YEARFRAC|start_date, end_date, [basis]|Fraction of a year between dates.|Date & Time
XMATCH|lookup_value, lookup_array, [match_mode], [search_mode]|Position of a match (exact, next smaller/larger, wildcard).|Lookup & Reference
LOOKUP|lookup_value, lookup_vector, [result_vector]|Approximate lookup in a sorted vector.|Lookup & Reference
ROW|[reference]|Row number of a reference.|Lookup & Reference
ROWS|array|Number of rows.|Lookup & Reference
COLUMN|[reference]|Column number of a reference.|Lookup & Reference
COLUMNS|array|Number of columns.|Lookup & Reference
ADDRESS|row_num, column_num, [abs_num], [a1], [sheet_text]|Builds a cell address as text.|Lookup & Reference
OFFSET|reference, rows, cols, [height], [width]|Reference offset from a starting cell.|Lookup & Reference
TRANSPOSE|array|Swaps rows and columns (spills).|Lookup & Reference
SORTBY|array, by_array1, [sort_order1], ...|Sorts a range by other ranges (spills).|Lookup & Reference
TAKE|array, rows, [columns]|First or last rows/columns of an array (spills).|Lookup & Reference
DROP|array, rows, [columns]|Removes rows/columns from the start or end (spills).|Lookup & Reference
CHOOSECOLS|array, col_num1, [col_num2], ...|Picks columns from an array (spills).|Lookup & Reference
CHOOSEROWS|array, row_num1, [row_num2], ...|Picks rows from an array (spills).|Lookup & Reference
SEQUENCE|rows, [columns], [start], [step]|Generates a sequence of numbers (spills).|Lookup & Reference
RATE|nper, pmt, pv, [fv], [type], [guess]|Interest rate per period.|Financial
NPER|rate, pmt, pv, [fv], [type]|Number of periods.|Financial
IPMT|rate, per, nper, pv, [fv], [type]|Interest part of a payment.|Financial
PPMT|rate, per, nper, pv, [fv], [type]|Principal part of a payment.|Financial
SLN|cost, salvage, life|Straight-line depreciation.|Financial
DB|cost, salvage, life, period, [month]|Fixed-declining balance depreciation.|Financial
DDB|cost, salvage, life, period, [factor]|Double-declining balance depreciation.|Financial
EFFECT|nominal_rate, npery|Effective annual interest rate.|Financial
ISTEXT|value|TRUE if the value is text.|Information
ISLOGICAL|value|TRUE if the value is TRUE or FALSE.|Information
ISERR|value|TRUE for any error except #N/A.|Information
ISNA|value|TRUE if the value is #N/A.|Information
TYPE|value|Type code: 1 number, 2 text, 4 logical, 16 error, 64 array.|Information
ERROR.TYPE|error_val|Number for an error value.|Information
DSUM|database, field, criteria|Sum of database records meeting criteria.|Database
DAVERAGE|database, field, criteria|Average of matching database records.|Database
DCOUNT|database, field, criteria|Counts numbers in matching records.|Database
DCOUNTA|database, field, criteria|Counts non-empty matching records.|Database
DMAX|database, field, criteria|Largest value in matching records.|Database
DMIN|database, field, criteria|Smallest value in matching records.|Database
`
for (const line of MORE_HELP.trim().split('\n')) {
  const [name, args, description, category] = line.split('|')
  if (!FUNCTION_HELP[name]) FUNCTION_HELP[name] = { syntax: `${name}(${args})`, description, category }
}

/** Syntax help for a function, falling back to a bare signature. */
export function helpFor(name: string): { syntax: string; description: string; category: string } | undefined {
  return FUNCTION_HELP[name.toUpperCase()]
}
