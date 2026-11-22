import type { NamedRange, PureSheetsDocument } from '../types'
import { cellKey, cellsInRange, normalizeRangeToken, parseCellKey } from './workbookModel'

export type FormulaCategory =
  | 'Math'
  | 'Stats'
  | 'Lookup'
  | 'Logical'
  | 'Text'
  | 'Date'
  | 'Financial'
  | 'Array'
  | 'Info'

export interface FormulaFunctionDoc {
  name: string
  category: FormulaCategory
  params: string[]
  signature: string
  description: string
}

const RAW_FUNCTIONS: Array<[string, FormulaCategory, string[], string]> = [
  // Core Math
  ['SUM', 'Math', ['range'], 'Add the numbers in a range'],
  ['SUMIF', 'Math', ['range', 'criteria', '[sum_range]'], 'Sum cells matching a criterion'],
  ['SUMIFS', 'Math', ['sum_range', 'criteria_range1', 'criteria1', '…'], 'Sum cells matching multiple criteria'],
  ['SUMPRODUCT', 'Math', ['array1', '[array2]', '…'], 'Sum of products of corresponding array elements'],
  ['ROUND', 'Math', ['number', 'num_digits'], 'Round a number to specified digits'],
  ['ROUNDUP', 'Math', ['number', 'num_digits'], 'Round a number up away from zero'],
  ['ROUNDDOWN', 'Math', ['number', 'num_digits'], 'Round a number down toward zero'],
  ['MROUND', 'Math', ['number', 'multiple'], 'Round to the nearest multiple'],
  ['CEILING', 'Math', ['number', 'significance'], 'Round up to the nearest multiple'],
  ['FLOOR', 'Math', ['number', 'significance'], 'Round down to the nearest multiple'],
  ['INT', 'Math', ['number'], 'Round down to the nearest integer'],
  ['TRUNC', 'Math', ['number', '[num_digits]'], 'Truncate a number to an integer or digits'],
  ['ABS', 'Math', ['number'], 'Absolute value of a number'],
  ['SIGN', 'Math', ['number'], 'Sign of a number (-1, 0, or 1)'],
  ['MOD', 'Math', ['number', 'divisor'], 'Remainder after division'],
  ['QUOTIENT', 'Math', ['numerator', 'denominator'], 'Integer portion of division'],
  ['POWER', 'Math', ['number', 'power'], 'Raise a number to a power'],
  ['SQRT', 'Math', ['number'], 'Positive square root'],
  ['EXP', 'Math', ['number'], 'e raised to the power of number'],
  ['LN', 'Math', ['number'], 'Natural logarithm'],
  ['LOG', 'Math', ['number', '[base]'], 'Logarithm to a specified base'],
  ['LOG10', 'Math', ['number'], 'Base-10 logarithm'],
  ['PRODUCT', 'Math', ['number1', '[number2]', '…'], 'Multiply all given numbers'],
  ['PI', 'Math', [], 'Mathematical constant π (3.14159…)'],
  ['RAND', 'Math', [], 'Random number between 0 and 1'],
  ['RANDBETWEEN', 'Math', ['bottom', 'top'], 'Random integer between two bounds'],
  ['SUBTOTAL', 'Math', ['function_num', 'ref1', '…'], 'Aggregate ignoring hidden or filtered rows'],
  ['AGGREGATE', 'Math', ['function_num', 'options', 'ref1', '…'], 'Aggregate with error/hidden-row control'],
  ['GCD', 'Math', ['number1', '[number2]', '…'], 'Greatest common divisor'],
  ['LCM', 'Math', ['number1', '[number2]', '…'], 'Least common multiple'],
  ['FACT', 'Math', ['number'], 'Factorial of a number'],
  ['COMBIN', 'Math', ['number', 'number_chosen'], 'Number of combinations'],
  ['SIN', 'Math', ['number'], 'Sine of an angle in radians'],
  ['COS', 'Math', ['number'], 'Cosine of an angle in radians'],
  ['TAN', 'Math', ['number'], 'Tangent of an angle in radians'],
  ['RADIANS', 'Math', ['angle'], 'Convert degrees to radians'],
  ['DEGREES', 'Math', ['angle'], 'Convert radians to degrees'],

  // Statistics
  ['AVERAGE', 'Stats', ['number1', '[number2]', '…'], 'Arithmetic mean of numbers'],
  ['AVERAGEA', 'Stats', ['value1', '[value2]', '…'], 'Mean including text and logicals'],
  ['AVERAGEIF', 'Stats', ['range', 'criteria', '[average_range]'], 'Mean of cells matching a criterion'],
  ['AVERAGEIFS', 'Stats', ['average_range', 'criteria_range1', 'criteria1', '…'], 'Mean of cells matching multiple criteria'],
  ['MEDIAN', 'Stats', ['number1', '[number2]', '…'], 'Middle value of a numeric set'],
  ['MODE', 'Stats', ['number1', '[number2]', '…'], 'Most frequently occurring number'],
  ['MIN', 'Stats', ['number1', '[number2]', '…'], 'Smallest number in a range'],
  ['MINIFS', 'Stats', ['min_range', 'criteria_range1', 'criteria1', '…'], 'Smallest value matching criteria'],
  ['MAX', 'Stats', ['number1', '[number2]', '…'], 'Largest number in a range'],
  ['MAXIFS', 'Stats', ['max_range', 'criteria_range1', 'criteria1', '…'], 'Largest value matching criteria'],
  ['LARGE', 'Stats', ['array', 'k'], 'k-th largest value in a dataset'],
  ['SMALL', 'Stats', ['array', 'k'], 'k-th smallest value in a dataset'],
  ['COUNT', 'Stats', ['value1', '[value2]', '…'], 'Count numeric cells'],
  ['COUNTA', 'Stats', ['value1', '[value2]', '…'], 'Count non-empty cells'],
  ['COUNTBLANK', 'Stats', ['range'], 'Count empty cells in a range'],
  ['COUNTIF', 'Stats', ['range', 'criteria'], 'Count cells matching a criterion'],
  ['COUNTIFS', 'Stats', ['criteria_range1', 'criteria1', '…'], 'Count cells matching multiple criteria'],
  ['RANK', 'Stats', ['number', 'ref', '[order]'], 'Rank of a number within a range'],
  ['RANK.EQ', 'Stats', ['number', 'ref', '[order]'], 'Rank of a number (tied values share rank)'],
  ['RANK.AVG', 'Stats', ['number', 'ref', '[order]'], 'Rank of a number (tied values average rank)'],
  ['STDEV', 'Stats', ['number1', '[number2]', '…'], 'Sample standard deviation'],
  ['STDEV.S', 'Stats', ['number1', '[number2]', '…'], 'Sample standard deviation'],
  ['STDEV.P', 'Stats', ['number1', '[number2]', '…'], 'Population standard deviation'],
  ['VAR.S', 'Stats', ['number1', '[number2]', '…'], 'Sample variance'],
  ['VAR.P', 'Stats', ['number1', '[number2]', '…'], 'Population variance'],
  ['PERCENTILE', 'Stats', ['array', 'k'], 'k-th percentile of values (0–1)'],
  ['PERCENTILE.INC', 'Stats', ['array', 'k'], 'Inclusive k-th percentile'],
  ['QUARTILE', 'Stats', ['array', 'quart'], 'Quartile of a dataset (0–4)'],
  ['CORREL', 'Stats', ['array1', 'array2'], 'Correlation coefficient between two ranges'],
  ['SLOPE', 'Stats', ['known_y', 'known_x'], 'Slope of linear regression line'],
  ['INTERCEPT', 'Stats', ['known_y', 'known_x'], 'Intercept of linear regression line'],
  ['FORECAST.LINEAR', 'Stats', ['x', 'known_y', 'known_x'], 'Predict a future value along a linear trend'],
  ['GEOMEAN', 'Stats', ['number1', '[number2]', '…'], 'Geometric mean'],
  ['TRIMMEAN', 'Stats', ['array', 'percent'], 'Mean excluding extreme tails'],

  // Lookup & Reference
  ['XLOOKUP', 'Lookup', ['lookup_value', 'lookup_array', 'return_array', '[if_not_found]', '[match_mode]', '[search_mode]'], 'Modern lookup in any direction with default fallback'],
  ['XMATCH', 'Lookup', ['lookup_value', 'lookup_array', '[match_mode]', '[search_mode]'], 'Relative position of an item in an array'],
  ['VLOOKUP', 'Lookup', ['lookup_value', 'table_array', 'col_index_num', '[range_lookup]'], 'Look up a value in the first column'],
  ['HLOOKUP', 'Lookup', ['lookup_value', 'table_array', 'row_index_num', '[range_lookup]'], 'Look up a value in the top row'],
  ['INDEX', 'Lookup', ['array', 'row_num', '[column_num]'], 'Return value at a row and column in a range'],
  ['MATCH', 'Lookup', ['lookup_value', 'lookup_array', '[match_type]'], 'Find relative position of a value in a range'],
  ['CHOOSE', 'Lookup', ['index_num', 'value1', '[value2]', '…'], 'Pick a value from a list by index'],
  ['OFFSET', 'Lookup', ['reference', 'rows', 'cols', '[height]', '[width]'], 'Reference offset from a starting cell'],
  ['INDIRECT', 'Lookup', ['ref_text', '[a1]'], 'Evaluate a reference specified as text'],
  ['ROW', 'Lookup', ['[reference]'], 'Row number of a reference'],
  ['ROWS', 'Lookup', ['array'], 'Number of rows in a reference'],
  ['COLUMN', 'Lookup', ['[reference]'], 'Column number of a reference'],
  ['COLUMNS', 'Lookup', ['array'], 'Number of columns in a reference'],
  ['ADDRESS', 'Lookup', ['row_num', 'column_num', '[abs_num]'], 'Create a cell address as text'],
  ['TRANSPOSE', 'Lookup', ['array'], 'Flip rows and columns of a range'],
  ['HYPERLINK', 'Lookup', ['link_location', '[friendly_name]'], 'Create a clickable link'],

  // Logical
  ['IF', 'Logical', ['logical_test', 'value_if_true', '[value_if_false]'], 'Return one value if true and another if false'],
  ['IFS', 'Logical', ['test1', 'value1', '[test2]', '[value2]', '…'], 'Check multiple conditions in order'],
  ['SWITCH', 'Logical', ['expression', 'value1', 'result1', '…', '[default]'], 'Match an expression against a list of values'],
  ['IFERROR', 'Logical', ['value', 'value_if_error'], 'Return a fallback if a formula evaluates to an error'],
  ['IFNA', 'Logical', ['value', 'value_if_na'], 'Return a fallback if a formula evaluates to #N/A'],
  ['AND', 'Logical', ['logical1', '[logical2]', '…'], 'TRUE if all arguments are TRUE'],
  ['OR', 'Logical', ['logical1', '[logical2]', '…'], 'TRUE if any argument is TRUE'],
  ['NOT', 'Logical', ['logical'], 'Reverse a boolean value'],
  ['XOR', 'Logical', ['logical1', '[logical2]', '…'], 'Logical exclusive OR'],
  ['TRUE', 'Logical', [], 'Logical value TRUE'],
  ['FALSE', 'Logical', [], 'Logical value FALSE'],
  ['LET', 'Logical', ['name1', 'value1', 'calculation'], 'Assign names to intermediate calculation results'],

  // Text
  ['CONCAT', 'Text', ['text1', '[text2]', '…'], 'Join values or ranges into one text string'],
  ['CONCATENATE', 'Text', ['text1', '[text2]', '…'], 'Join several text strings into one'],
  ['TEXTJOIN', 'Text', ['delimiter', 'ignore_empty', 'text1', '…'], 'Join text with a delimiter, skipping blanks'],
  ['TEXTSPLIT', 'Text', ['text', 'col_delimiter', '[row_delimiter]'], 'Split text by delimiters into cells'],
  ['TEXTBEFORE', 'Text', ['text', 'delimiter', '[instance_num]'], 'Text appearing before a delimiter'],
  ['TEXTAFTER', 'Text', ['text', 'delimiter', '[instance_num]'], 'Text appearing after a delimiter'],
  ['LEFT', 'Text', ['text', '[num_chars]'], 'Leading characters from text'],
  ['RIGHT', 'Text', ['text', '[num_chars]'], 'Trailing characters from text'],
  ['MID', 'Text', ['text', 'start_num', 'num_chars'], 'Characters from the middle of text'],
  ['LEN', 'Text', ['text'], 'Number of characters in text'],
  ['TRIM', 'Text', ['text'], 'Remove extra leading, trailing, and double spaces'],
  ['CLEAN', 'Text', ['text'], 'Remove non-printable characters'],
  ['LOWER', 'Text', ['text'], 'Convert text to lowercase'],
  ['UPPER', 'Text', ['text'], 'Convert text to uppercase'],
  ['PROPER', 'Text', ['text'], 'Capitalize the first letter of each word'],
  ['SUBSTITUTE', 'Text', ['text', 'old_text', 'new_text', '[instance_num]'], 'Replace matching text substrings'],
  ['REPLACE', 'Text', ['old_text', 'start_num', 'num_chars', 'new_text'], 'Replace characters at a given position'],
  ['FIND', 'Text', ['find_text', 'within_text', '[start_num]'], 'Case-sensitive position of a substring'],
  ['SEARCH', 'Text', ['find_text', 'within_text', '[start_num]'], 'Case-insensitive position of a substring'],
  ['TEXT', 'Text', ['value', 'format_text'], 'Format a number or date as text'],
  ['VALUE', 'Text', ['text'], 'Convert numeric text to a number'],
  ['EXACT', 'Text', ['text1', 'text2'], 'Check if two strings are identical (case-sensitive)'],
  ['REPT', 'Text', ['text', 'number_times'], 'Repeat text a given number of times'],
  ['CHAR', 'Text', ['number'], 'Character specified by a code number'],
  ['CODE', 'Text', ['text'], 'Numeric code for the first character'],
  ['REGEXMATCH', 'Text', ['text', 'regular_expression'], 'Whether text matches a regular expression'],
  ['REGEXEXTRACT', 'Text', ['text', 'regular_expression'], 'Extract matching substring via regex'],
  ['REGEXREPLACE', 'Text', ['text', 'regular_expression', 'replacement'], 'Replace matching text via regex'],

  // Date & Time
  ['TODAY', 'Date', [], "Today's date"],
  ['NOW', 'Date', [], 'Current date and time'],
  ['DATE', 'Date', ['year', 'month', 'day'], 'Create a date from year, month, and day'],
  ['DATEVALUE', 'Date', ['date_text'], 'Convert a date string into a date serial'],
  ['TIME', 'Date', ['hour', 'minute', 'second'], 'Create a time value'],
  ['YEAR', 'Date', ['serial_number'], 'Extract the year from a date'],
  ['MONTH', 'Date', ['serial_number'], 'Extract the month (1–12) from a date'],
  ['DAY', 'Date', ['serial_number'], 'Extract the day of the month (1–31)'],
  ['HOUR', 'Date', ['serial_number'], 'Extract the hour (0–23)'],
  ['MINUTE', 'Date', ['serial_number'], 'Extract the minute (0–59)'],
  ['SECOND', 'Date', ['serial_number'], 'Extract the second (0–59)'],
  ['WEEKDAY', 'Date', ['serial_number', '[return_type]'], 'Day of the week as a number'],
  ['WEEKNUM', 'Date', ['serial_number', '[return_type]'], 'Week number in the year'],
  ['ISOWEEKNUM', 'Date', ['date'], 'ISO week number of the year'],
  ['DAYS', 'Date', ['end_date', 'start_date'], 'Number of days between two dates'],
  ['DATEDIF', 'Date', ['start_date', 'end_date', 'unit'], 'Difference between dates in Y, M, or D'],
  ['EDATE', 'Date', ['start_date', 'months'], 'Date a given number of months before/after'],
  ['EOMONTH', 'Date', ['start_date', 'months'], 'Last day of the month N months away'],
  ['NETWORKDAYS', 'Date', ['start_date', 'end_date', '[holidays]'], 'Working days between two dates'],
  ['WORKDAY', 'Date', ['start_date', 'days', '[holidays]'], 'Date after a given number of workdays'],
  ['YEARFRAC', 'Date', ['start_date', 'end_date', '[basis]'], 'Fraction of the year between two dates'],

  // Financial
  ['PMT', 'Financial', ['rate', 'nper', 'pv', '[fv]', '[type]'], 'Periodic payment for a loan or annuity'],
  ['PV', 'Financial', ['rate', 'nper', 'pmt', '[fv]', '[type]'], 'Present value of an investment'],
  ['FV', 'Financial', ['rate', 'nper', 'pmt', '[pv]', '[type]'], 'Future value of an investment'],
  ['NPV', 'Financial', ['rate', 'value1', '[value2]', '…'], 'Net present value of cash flows'],
  ['XNPV', 'Financial', ['rate', 'values', 'dates'], 'Net present value for irregular dates'],
  ['IRR', 'Financial', ['values', '[guess]'], 'Internal rate of return for periodic cash flows'],
  ['XIRR', 'Financial', ['values', 'dates', '[guess]'], 'Internal rate of return for irregular dates'],
  ['RATE', 'Financial', ['nper', 'pmt', 'pv', '[fv]', '[type]'], 'Interest rate per period of an annuity'],
  ['NPER', 'Financial', ['rate', 'pmt', 'pv', '[fv]', '[type]'], 'Number of payment periods'],
  ['IPMT', 'Financial', ['rate', 'per', 'nper', 'pv'], 'Interest payment for a given period'],
  ['PPMT', 'Financial', ['rate', 'per', 'nper', 'pv'], 'Principal payment for a given period'],
  ['SLN', 'Financial', ['cost', 'salvage', 'life'], 'Straight-line depreciation per period'],
  ['EFFECT', 'Financial', ['nominal_rate', 'npery'], 'Effective annual interest rate'],
  ['NOMINAL', 'Financial', ['effect_rate', 'npery'], 'Nominal annual interest rate'],

  // Dynamic Arrays
  ['FILTER', 'Array', ['array', 'include', '[if_empty]'], 'Filter a range dynamically by condition'],
  ['SORT', 'Array', ['array', '[sort_index]', '[sort_order]', '[by_col]'], 'Sort a range dynamically'],
  ['SORTBY', 'Array', ['array', 'by_array1', '[sort_order1]', '…'], 'Sort a range by another range'],
  ['UNIQUE', 'Array', ['array', '[by_col]', '[exactly_once]'], 'Return unique rows or values from a range'],
  ['SEQUENCE', 'Array', ['rows', '[columns]', '[start]', '[step]'], 'Generate a sequential array of numbers'],
  ['VSTACK', 'Array', ['array1', '[array2]', '…'], 'Stack arrays vertically'],
  ['HSTACK', 'Array', ['array1', '[array2]', '…'], 'Stack arrays horizontally'],
  ['TAKE', 'Array', ['array', 'rows', '[columns]'], 'Take first or last N rows/columns'],
  ['DROP', 'Array', ['array', 'rows', '[columns]'], 'Drop first or last N rows/columns'],
  ['CHOOSECOLS', 'Array', ['array', 'col_num1', '…'], 'Select specific columns from an array'],
  ['CHOOSEROWS', 'Array', ['array', 'row_num1', '…'], 'Select specific rows from an array'],
  ['TOCOL', 'Array', ['array', '[ignore]', '[scan_by_column]'], 'Flatten an array into one column'],
  ['TOROW', 'Array', ['array', '[ignore]', '[scan_by_column]'], 'Flatten an array into one row'],

  // Information
  ['ISBLANK', 'Info', ['value'], 'TRUE if cell is empty'],
  ['ISNUMBER', 'Info', ['value'], 'TRUE if value is a number'],
  ['ISTEXT', 'Info', ['value'], 'TRUE if value is text'],
  ['ISERROR', 'Info', ['value'], 'TRUE if value is any error'],
  ['ISNA', 'Info', ['value'], 'TRUE if value is #N/A'],
  ['ISFORMULA', 'Info', ['reference'], 'TRUE if cell contains a formula'],
  ['ISLOGICAL', 'Info', ['value'], 'TRUE if value is TRUE or FALSE'],
  ['ISEVEN', 'Info', ['number'], 'TRUE if number is even'],
  ['ISODD', 'Info', ['number'], 'TRUE if number is odd'],
  ['FORMULATEXT', 'Info', ['reference'], 'Return a cell’s formula as text'],
  ['TYPE', 'Info', ['value'], 'Numeric code indicating data type'],
  ['NA', 'Info', [], 'Return the #N/A error value'],
]

export const FORMULA_CATALOG: FormulaFunctionDoc[] = RAW_FUNCTIONS.map(
  ([name, category, params, description]) => ({
    name,
    category,
    params,
    signature: `${name}(${params.join(', ')})`,
    description,
  }),
)

const CATALOG_BY_NAME = new Map<string, FormulaFunctionDoc>(
  FORMULA_CATALOG.map(fn => [fn.name, fn]),
)

export function lookupFormulaFunction(name: string): FormulaFunctionDoc | undefined {
  return CATALOG_BY_NAME.get(name.toUpperCase())
}

export interface ActiveCallContext {
  fn: FormulaFunctionDoc
  argIndex: number
}

/**
 * Given a formula string and caret offset, find the innermost enclosing
 * function call and which 0-based argument the cursor is currently inside.
 * Ignores commas and parentheses inside quoted strings.
 */
export function activeCallContext(
  formula: string,
  cursorPos: number,
): ActiveCallContext | null {
  if (!formula.trimStart().startsWith('=')) return null
  const clamped = Math.max(0, Math.min(cursorPos, formula.length))
  const stack: Array<{ name: string | null; argIndex: number }> = []
  let inString = false

  for (let i = 0; i < clamped; i += 1) {
    const ch = formula[i]
    if (ch === '"') {
      inString = !inString
      continue
    }
    if (inString) continue
    if (ch === '(') {
      const before = formula.slice(0, i)
      const match = /([A-Z][A-Z0-9._]*)$/i.exec(before)
      stack.push({ name: match ? match[1].toUpperCase() : null, argIndex: 0 })
    } else if (ch === ')') {
      stack.pop()
    } else if (ch === ',' && stack.length > 0) {
      stack[stack.length - 1].argIndex += 1
    }
  }

  for (let i = stack.length - 1; i >= 0; i -= 1) {
    const frame = stack[i]
    if (!frame.name) continue
    const fn = CATALOG_BY_NAME.get(frame.name)
    if (fn) return { fn, argIndex: frame.argIndex }
  }
  return null
}

export interface FormulaErrorExplanation {
  code: string
  title: string
  detail: string
  suggestedFix?: {
    label: string
    formula: string
  }
}

/**
 * Plain-English explanation + one-click fix suggestion for standard Excel
 * formula errors.
 */
export function explainFormulaError(
  errorCode: string,
  formula: string,
): FormulaErrorExplanation {
  const code = errorCode.trim().toUpperCase()
  const bareExpr = formula.trim().replace(/^=/, '')
  const canWrapIferror =
    formula.trim().startsWith('=') && !/^=IFERROR\(/i.test(formula.trim())

  if (code === '#DIV/0!') {
    return {
      code,
      title: 'Division by zero',
      detail: 'A divisor in this formula evaluates to 0 or an empty cell.',
      suggestedFix: canWrapIferror
        ? { label: 'Wrap in IFERROR(…, 0)', formula: `=IFERROR(${bareExpr}, 0)` }
        : undefined,
    }
  }
  if (code === '#NAME?') {
    return {
      code,
      title: 'Unrecognized name',
      detail:
        'A function name, named range, or unquoted text label could not be resolved.',
    }
  }
  if (code === '#REF!') {
    return {
      code,
      title: 'Invalid cell reference',
      detail:
        'A referenced cell, row, column, or sheet was deleted or creates a circular loop.',
    }
  }
  if (code === '#VALUE!' || code === '#VALUE') {
    return {
      code,
      title: 'Wrong value type',
      detail:
        'An operator or function received text where a number or date was expected.',
      suggestedFix: canWrapIferror
        ? { label: 'Wrap in IFERROR(…, "")', formula: `=IFERROR(${bareExpr}, "")` }
        : undefined,
    }
  }
  if (code === '#N/A') {
    return {
      code,
      title: 'Value not found',
      detail:
        'A lookup function (VLOOKUP, XLOOKUP, MATCH) could not find a matching entry.',
      suggestedFix: canWrapIferror
        ? { label: 'Wrap in IFERROR(…, "Not found")', formula: `=IFERROR(${bareExpr}, "Not found")` }
        : undefined,
    }
  }
  if (code === '#NUM!') {
    return {
      code,
      title: 'Invalid numeric result',
      detail:
        'A calculation produced an out-of-range or impossible number (e.g. square root of a negative number).',
    }
  }
  if (code === '#CYCLE!' || code === '#CIRC!') {
    return {
      code,
      title: 'Circular reference',
      detail: 'This cell references itself directly or through another cell.',
    }
  }
  return {
    code,
    title: 'Formula error',
    detail: `This formula evaluates to ${code}. Check its syntax and referenced cells.`,
    suggestedFix: canWrapIferror
      ? { label: 'Wrap in IFERROR(…, "")', formula: `=IFERROR(${bareExpr}, "")` }
      : undefined,
  }
}

export interface FormulaReference {
  sheetName: string
  range: string
  token: string
}

const REF_PATTERN =
  /(?:'([^']+)'!|([A-Za-z0-9_]+)!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)(?![A-Z0-9_(])/gi

/**
 * Extract all unique cell/range precedents referenced by a formula.
 */
export function extractFormulaPrecedents(
  formula: string,
  activeSheetName: string,
): FormulaReference[] {
  if (!formula.trimStart().startsWith('=')) return []
  const seen = new Set<string>()
  const refs: FormulaReference[] = []
  const segments = formula.split('"')

  segments.forEach((segment, idx) => {
    if (idx % 2 === 1) return // skip string literals
    for (const match of segment.matchAll(REF_PATTERN)) {
      const sheetName = (match[1] ?? match[2] ?? activeSheetName).trim()
      const cleanRange = match[3].toUpperCase().replace(/\$/g, '')
      const id = `${sheetName.toLowerCase()}!${cleanRange}`
      if (seen.has(id)) continue
      seen.add(id)
      refs.push({
        sheetName,
        range: cleanRange,
        token: match[0],
      })
    }
  })
  return refs
}

export interface DependentRef {
  sheetId: string
  sheetName: string
  cell: string
  formula: string
}

/**
 * Find every formula cell across the workbook that references `targetCell` on
 * `targetSheetId`.
 */
export function traceDependents(
  workbook: PureSheetsDocument['workbook'],
  targetSheetId: string,
  targetCell: string,
): DependentRef[] {
  const targetSheet = workbook.sheets.find(s => s.id === targetSheetId)
  if (!targetSheet) return []
  const targetPos = parseCellKey(targetCell)
  if (!targetPos) return []
  const targetKey = cellKey(targetPos.row, targetPos.column)
  const targetSheetLower = targetSheet.name.toLowerCase()
  const results: DependentRef[] = []

  for (const sheet of workbook.sheets) {
    if (sheet.agentDraft) continue
    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (cell.kind !== 'formula' || !cell.value.startsWith('=')) continue
      const precedents = extractFormulaPrecedents(cell.value, sheet.name)
      const hits = precedents.some(ref => {
        if (ref.sheetName.toLowerCase() !== targetSheetLower) return false
        if (!ref.range.includes(':')) return ref.range === targetKey
        return cellsInRange(ref.range, 10_000).includes(targetKey)
      })
      if (hits) {
        results.push({
          sheetId: sheet.id,
          sheetName: sheet.name,
          cell: key,
          formula: cell.value,
        })
      }
    }
  }
  return results
}

// ---------------------------------------------------------------------------
// Named ranges
// ---------------------------------------------------------------------------

let namedRangeCounter = 1

export function isValidNamedRangeName(name: string): boolean {
  const trimmed = name.trim()
  if (!/^[A-Za-z_][A-Za-z0-9_.]{0,63}$/.test(trimmed)) return false
  // Disallow valid A1 cell references like A1, BC12
  if (/^[A-Za-z]{1,3}\d+$/.test(trimmed)) return false
  return true
}

export function upsertNamedRange(
  document: PureSheetsDocument,
  input: { name: string; sheetId: string; range: string; comment?: string },
): PureSheetsDocument {
  const name = input.name.trim()
  const range = normalizeRangeToken(input.range)
  if (!isValidNamedRangeName(name) || !range) return document
  const existing = document.workbook.namedRanges ?? []
  const lower = name.toLowerCase()
  const match = existing.find(nr => nr.name.toLowerCase() === lower)
  const nextItem: NamedRange = {
    id: match?.id ?? `nr-${Date.now()}-${namedRangeCounter++}`,
    name,
    sheetId: input.sheetId,
    range,
    comment: input.comment?.trim() || undefined,
  }
  const nextList = match
    ? existing.map(nr => (nr.id === match.id ? nextItem : nr))
    : [...existing, nextItem]
  return {
    ...document,
    workbook: {
      ...document.workbook,
      namedRanges: nextList,
    },
  }
}

export function deleteNamedRange(
  document: PureSheetsDocument,
  namedRangeId: string,
): PureSheetsDocument {
  const existing = document.workbook.namedRanges ?? []
  const remaining = existing.filter(nr => nr.id !== namedRangeId)
  if (remaining.length === existing.length) return document
  return {
    ...document,
    workbook: {
      ...document.workbook,
      namedRanges: remaining.length ? remaining : undefined,
    },
  }
}

/**
 * Expand workbook Named Ranges inside a formula into explicit A1/Sheet!A1
 * references so both the fallback evaluator and Univer compute them reliably.
 */
export function resolveNamedRangesInFormula(
  formula: string,
  namedRanges: readonly NamedRange[] | undefined,
  sheets: readonly PureSheetsDocument['workbook']['sheets'][number][],
  activeSheetId: string,
): string {
  if (!namedRanges?.length || !formula.trimStart().startsWith('=')) return formula
  const byName = new Map(namedRanges.map(nr => [nr.name.toLowerCase(), nr]))
  const sheetNameById = new Map(sheets.map(s => [s.id, s.name]))
  const segments = formula.split('"')
  return segments
    .map((segment, idx) => {
      if (idx % 2 === 1) return segment
      return segment.replace(
        /(?<![A-Za-z0-9_!'.])([A-Za-z_][A-Za-z0-9_.]*)(?![A-Za-z0-9_(])/g,
        token => {
          const hit = byName.get(token.toLowerCase())
          if (!hit) return token
          if (hit.sheetId === activeSheetId) return hit.range
          const targetSheetName = sheetNameById.get(hit.sheetId)
          if (!targetSheetName) return token
          const quoted = /[^A-Za-z0-9_]/.test(targetSheetName)
            ? `'${targetSheetName.replace(/'/g, "''")}'`
            : targetSheetName
          return `${quoted}!${hit.range}`
        },
      )
    })
    .join('"')
}
