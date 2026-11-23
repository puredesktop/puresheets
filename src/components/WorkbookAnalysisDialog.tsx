import { useId, useLayoutEffect, useMemo, useRef, useState } from 'react'
import styled from 'styled-components'
import { Button } from '@purescience/platform-ui/components/common/buttons/Button'
import {
  applyRows, auditFormula, cleanedRows, goalSeek, pivotRows, readSelection, selectionFields, traceFormulaErrors,
  type AnalysisPivotSummary,
} from '../lib/workbookAnalysis'
import { computePivotSourceSignature } from '../lib/pivotEngine'
import { validRangeName } from '../lib/rangeNames'
import {
  addSheet, cellKey, cloneWithUpdatedSheet, computedCellValue, isFormulaError, parseCellKey, rangesIntersect, upsertCell,
} from '../lib/workbookModel'
import { withFreshUniverSnapshot } from '../lib/univerAdapter'
import type { PivotDefinition, PureSheetsDocument, SheetCell, WorkbookSheet } from '../types'

export const analysisOptions = [
  { value: 'cleanup', label: 'Clean up data' },
  { value: 'pivot', label: 'Pivot table' },
  { value: 'audit', label: 'Trace formulas' },
  { value: 'names', label: 'Named ranges' },
  { value: 'goal', label: 'Goal Seek' },
] as const

export type AnalysisMode = typeof analysisOptions[number]['value']

type Props = {
  mode: AnalysisMode
  document: PureSheetsDocument
  sheet: WorkbookSheet
  range: string
  cell: string
  onClose: () => void
  onApply: (document: PureSheetsDocument, message: string) => void
  onNavigate: (sheetId: string, range: string) => void
}

export function WorkbookAnalysisDialog(props: Props): React.ReactElement {
  const ref = useRef<HTMLDialogElement>(null)
  const id = useId()
  const [error, setError] = useState('')
  useLayoutEffect(() => { ref.current?.showModal() }, [])
  const apply: Props['onApply'] = (next, message) => {
    try { props.onApply(next, message) } catch (cause) {
      setError(cause instanceof Error ? cause.message : String(cause))
    }
  }
  return (
    <Panel ref={ref} aria-labelledby={id} onClose={props.onClose}>
      <header><h2 id={id}>{analysisOptions.find(option => option.value === props.mode)?.label}</h2>
        <Button onClick={props.onClose}>Close</Button></header>
      {props.mode === 'cleanup' || props.mode === 'pivot' ? <TableTool {...props} onApply={apply} /> : null}
      {props.mode === 'audit' ? <Audit {...props} /> : null}
      {props.mode === 'names' ? <Names {...props} onApply={apply} /> : null}
      {props.mode === 'goal' ? <Goal {...props} onApply={apply} /> : null}
      {error ? <p className="error" role="alert">{error}</p> : null}
    </Panel>
  )
}

function TableTool({ mode, document, sheet, range, onApply }: Props): React.ReactElement {
  const pivot = mode === 'pivot' ? sheet.pivot : undefined
  const source = pivot ? document.workbook.sheets.find(sheet => sheet.id === pivot.sourceSheetId) : sheet
  const baseColumn = parseCellKey((pivot?.sourceRange ?? range).split(':')[0])?.column ?? 0
  const [sourceRange, setSourceRange] = useState(pivot?.sourceRange ?? range)
  const [cleanup, setCleanup] = useState<'dedupe' | 'split'>('dedupe')
  const [separator, setSeparator] = useState(',')
  const [header, setHeader] = useState(true)
  const [excluded, setExcluded] = useState<number[]>([])
  const [row, setRow] = useState(pivot?.rows[0] ? Math.max(0, pivot.rows[0].columnIndex - baseColumn) : 0)
  const [column, setColumn] = useState(pivot?.columns[0] ? Math.max(0, pivot.columns[0].columnIndex - baseColumn) : -1)
  const [value, setValue] = useState(pivot?.values[0] ? Math.max(0, pivot.values[0].columnIndex - baseColumn) : 1)
  const initialSummary: AnalysisPivotSummary =
    pivot?.values[0]?.aggregate === 'avg' ? 'average' : pivot?.values[0]?.aggregate === 'count' ? 'count' : 'sum'
  const [summary, setSummary] = useState<AnalysisPivotSummary>(initialSummary)
  const [preview, setPreview] = useState<SheetCell[][] | null>(null)
  const [problem, setProblem] = useState('')
  const selection = useMemo(() => {
    if (!source) return null
    try { return readSelection(source, sourceRange, document.workbook) } catch { return null }
  }, [source, sourceRange, document.workbook])
  const fields = selection ? selectionFields(selection) : []
  const reset = () => { setPreview(null); setProblem('') }
  const generate = () => {
    try {
      if (!selection) throw new Error('Choose a source range on this sheet with up to 25,000 cells.')
      if (mode === 'cleanup') {
        setPreview(cleanedRows(selection, {
          mode: cleanup, separator: separator === '\\t' ? '\t' : separator, header,
          columns: fields.flatMap((_, index) => excluded.includes(index) ? [] : [index]),
        }))
      } else {
        setPreview(pivotRows(selection, { row, column: column < 0 ? null : column, value, summary }))
      }
      setProblem('')
    } catch (cause) { setPreview(null); setProblem(String((cause as Error).message)) }
  }
  const commit = () => {
    if (!preview || !source || !selection) return
    try {
      if (mode === 'cleanup') {
        const next = addSheet(document, uniqueSheetName(document, 'Cleaned data'))
        const output = next.workbook.sheets.at(-1)!
        onApply(cloneWithUpdatedSheet(next, applyRows(output, preview)), 'Created cleaned data sheet')
        return
      }
      const base = pivot ? document : addSheet(document, uniqueSheetName(document, 'Pivot'))
      const output = pivot ? sheet : base.workbook.sheets.at(-1)!
      const outputCell = pivot?.output ?? 'A1'
      if (pivot && source.id === output.id) {
        const origin = parseCellKey(outputCell)!
        const after = cellKey(origin.row + Math.max(preview.length, pivot.outputRows ?? preview.length) - 1, origin.column + Math.max(preview[0].length, pivot.outputColumns ?? preview[0].length) - 1)
        if (rangesIntersect(sourceRange, `${outputCell}:${after}`))
          throw new Error('The source range overlaps the pivot output. Choose a different source range.')
      }
      const result = applyRows(output, preview, outputCell, pivot)
      const colOffset = selection.first.column
      const definition: PivotDefinition = {
        id: pivot?.id ?? `pivot-${Date.now()}`,
        sourceSheetId: source.id,
        sourceRange,
        rows: [{ columnIndex: colOffset + row, label: fields[row] ?? `Column ${row + 1}` }],
        columns: column < 0 ? [] : [{ columnIndex: colOffset + column, label: fields[column] ?? `Column ${column + 1}` }],
        values: [{
          columnIndex: colOffset + value,
          label: fields[value] ?? `Column ${value + 1}`,
          aggregate: summary === 'average' ? 'avg' : summary,
        }],
        refreshedAt: new Date().toISOString(),
        sourceSignature: computePivotSourceSignature(source, sourceRange, document.workbook),
        output: outputCell,
        outputRows: preview.length,
        outputColumns: preview[0].length,
      }
      onApply(cloneWithUpdatedSheet(base, { ...result, pivot: definition }), pivot ? 'Updated pivot table' : 'Created pivot table')
    } catch (cause) { setProblem(String((cause as Error).message)) }
  }
  return (
    <div className="body" onChange={reset}>
      <label>Source range {source ? `on ${source.name}` : ''}<input value={sourceRange} onChange={event => setSourceRange(event.target.value)} /></label>
      {mode === 'cleanup' ? <>
        <label>Operation<select value={cleanup} onChange={event => setCleanup(event.target.value as typeof cleanup)}>
          <option value="dedupe">Remove duplicates</option><option value="split">Split text into columns</option>
        </select></label>
        <label className="check"><input type="checkbox" checked={header} onChange={event => setHeader(event.target.checked)} /> First row has headers</label>
        {cleanup === 'split' ? <label>Delimiter<input value={separator} onChange={event => setSeparator(event.target.value)} /><small>Use <code>{String.raw`\t`}</code> for a tab. Select a single source column.</small></label> :
          <fieldset><legend>Compare these columns</legend><div className="checks">{fields.map((field, index) =>
            <label className="check" key={index}><input type="checkbox" checked={!excluded.includes(index)} onChange={event => setExcluded(event.target.checked ? excluded.filter(value => value !== index) : [...excluded, index])} />
              {header ? field : `Column ${index + 1}`}
            </label>
          )}</div></fieldset>}
        <p>The preview writes values to a new sheet.</p>
      </> : <>
        <div className="fields">
          <FieldPicker label="Rows" fields={fields} value={row} onChange={setRow} />
          <FieldPicker label="Columns" fields={fields} value={column} onChange={setColumn} optional />
          <FieldPicker label="Values" fields={fields} value={value} onChange={setValue} />
          <label>Summarize by<select value={summary} onChange={event => setSummary(event.target.value as typeof summary)}><option value="sum">Sum</option><option value="count">Count</option><option value="average">Average</option></select></label>
        </div>
        <p>{pivot ? 'Preview the new result, then replace the current pivot output.' : 'The selection must include headers. The pivot opens on a new sheet.'}</p>
      </>}
      {problem ? <p className="error" role="alert">{problem}</p> : null}
      {preview ? <>
        <p role="status">{preview.length} rows in the result{mode === 'cleanup' && cleanup === 'dedupe' && selection ? ` · ${selection.rows.length - preview.length} duplicates removed` : ''}</p>
        <div className="table"><table aria-label="Result preview"><tbody>{preview.slice(0, 8).map((line, index) =>
          <tr key={index}>{line.slice(0, 12).map((cell, col) => <td key={col}>{cell.value}</td>)}</tr>
        )}</tbody></table></div>
      </> : null}
      <footer><Button onClick={generate}>{preview ? 'Update preview' : 'Preview'}</Button>
        <Button variant="primary" disabled={!preview} onClick={commit}>{mode === 'cleanup' ? 'Create cleaned sheet' : pivot ? 'Update pivot' : 'Create pivot'}</Button></footer>
    </div>
  )
}

function uniqueSheetName(document: PureSheetsDocument, base: string): string {
  const names = new Set(document.workbook.sheets.map(sheet => sheet.name.toLowerCase()))
  let name = base
  for (let index = 2; names.has(name.toLowerCase()); index++) name = `${base} ${index}`
  return name
}

function FieldPicker({ label, fields, value, onChange, optional }: {
  label: string; fields: string[]; value: number; onChange: (value: number) => void; optional?: boolean
}) {
  return <label>{label}<select value={value} onChange={event => onChange(Number(event.target.value))}>
    {optional ? <option value={-1}>None</option> : null}
    {fields.map((name, index) => <option value={index} key={index}>{name}</option>)}
  </select></label>
}

function Audit({ document, sheet, cell, onNavigate }: Props): React.ReactElement {
  const [selected, setSelected] = useState({ sheet, cell })
  const current = selected.sheet.cells[selected.cell]
  const result = computedCellValue(selected.sheet, selected.cell, document.workbook)
  const trace = useMemo(() => auditFormula(document.workbook, selected.sheet, selected.cell), [document.workbook, selected])
  const errors = isFormulaError(result) ? traceFormulaErrors(document.workbook, trace.precedents) : []
  function group(title: string, references: typeof trace.precedents) {
    return <section><h3>{title}</h3>
      {references.length ? <ul>{references.map(({ sheetId, range }, index) => {
        const at = document.workbook.sheets.find(sheet => sheet.id === sheetId)
        if (!at) return null
        const result = range.includes(':') ? '' : computedCellValue(at, range, document.workbook)
        return <li key={`${sheetId}:${range}:${index}`}>
          <Button onClick={() => onNavigate(sheetId, range)}>{at.name}!{range}</Button>
          <span>{result}</span>
          {at.cells[range]?.kind === 'formula' && <Button onClick={() => setSelected({ sheet: at, cell: range })}>Trace</Button>}
        </li>
      })}</ul> : <p>None found</p>}
    </section>
  }
  return <div className="body">
    <p><strong>{selected.sheet.name}!{selected.cell}</strong> {current?.kind === 'formula' && <code>{current.value}</code>}</p>
    <p>Result <strong>{result || '(empty)'}</strong></p>
    {isFormulaError(result) && <p className="error">Follow a formula below to trace where the error begins.</p>}
    {errors.length ? group('Referenced formulas with errors', errors) : null}
    {group('Reads from', trace.precedents)}{group('Used by', trace.dependents)}
    {trace.limited ? <p>Showing the first 200 dependent formulas.</p> : null}
    <p>Shows explicit cell and named-range references. References produced by functions such as INDIRECT may be omitted.</p>
  </div>
}

function Names({ document, sheet, range, onApply, onNavigate }: Props): React.ReactElement {
  const [name, setName] = useState('')
  const names = document.workbook.namedRanges ?? []
  const key = name.trim()
  const existing = names.find(entry => entry.name.toLowerCase() === key.toLowerCase())
  const save = (next: typeof names, message: string) => onApply(withFreshUniverSnapshot({
    ...document, metadata: { ...document.metadata, updatedAt: new Date().toISOString() },
    workbook: { ...document.workbook, namedRanges: next.length ? next : undefined },
  }), message)
  return <div className="body">
    <form onSubmit={event => {
      event.preventDefault()
      if (validRangeName(key)) save([...names.filter(entry => entry !== existing), {
        id: existing?.id ?? `nr-${Date.now()}`,
        name: key,
        sheetId: sheet.id,
        range,
        ...(existing?.comment ? { comment: existing.comment } : {}),
      }], `Saved named range ${key}`)
    }}>
      <label>Name for {sheet.name}!{range}<input value={name} onChange={event => setName(event.target.value)} placeholder="TaxRate" /></label>
      <Button variant="primary" type="submit" disabled={!validRangeName(key)}>{existing ? 'Update name to this range' : 'Save name'}</Button>
    </form>
    <p>Use a name in formulas, for example <code>=SUM(Sales)</code>.</p>
    {names.length ? <ul>{names.map(entry => <li key={entry.name}>
      <Button onClick={() => onNavigate(entry.sheetId, entry.range)}>{entry.name}</Button>
      <span>{document.workbook.sheets.find(sheet => sheet.id === entry.sheetId)?.name}!{entry.range}</span>
      <Button onClick={() => save(names.filter(name => name !== entry), `Removed ${entry.name}`)}>Remove</Button>
    </li>)}</ul> : null}
  </div>
}

function Goal({ document, sheet, cell, onApply }: Props): React.ReactElement {
  const [formula, setFormula] = useState(cell)
  const [variable, setVariable] = useState('')
  const [target, setTarget] = useState('')
  const [preview, setPreview] = useState<ReturnType<typeof goalSeek> | null>(null)
  const [problem, setProblem] = useState('')
  const address = (value: string) => value.trim().replace(/\$/g, '').toUpperCase()
  return <div className="body">
    <form onChange={() => { setPreview(null); setProblem('') }} onSubmit={event => {
      event.preventDefault()
      try { setPreview(goalSeek(document.workbook, sheet, address(formula), address(variable), target.trim() ? Number(target) : NaN)); setProblem('') }
      catch (cause) { setPreview(null); setProblem(String((cause as Error).message)) }
    }}>
      <label>Formula cell<input value={formula} onChange={event => setFormula(event.target.value)} /></label>
      <label>Target value<input value={target} inputMode="decimal" onChange={event => setTarget(event.target.value)} /></label>
      <label>Input cell to change<input value={variable} placeholder="B2" onChange={event => setVariable(event.target.value)} /></label>
      <Button type="submit">Find value</Button>
    </form>
    {problem ? <p className="error" role="alert">{problem}</p> : null}
    {preview ? <p role="status">Change {address(variable)} from <strong>{preview.initial}</strong> to <strong>{preview.value}</strong>. The calculated result is <strong>{preview.result}</strong>.</p> : null}
    <footer><Button variant="primary" disabled={!preview} onClick={() => {
      if (!preview) return
      const position = address(variable)
      onApply(cloneWithUpdatedSheet(document, upsertCell(sheet, position, String(preview.value))), `Goal Seek changed ${position}`)
    }}>Apply value</Button></footer>
  </div>
}

const Panel = styled.dialog`
  width: min(556px, calc(100% - 24px));
  max-height: calc(100% - 32px);
  box-sizing: border-box;
  padding: 0;
  border: 1px solid var(--n-line);
  border-radius: 3px;
  background: var(--n-sheet);
  color: var(--n-ink);
  font: inherit;
  &::backdrop { background: rgb(0 0 0 / 42%); }
  header { position: sticky; top: 0; z-index: 1; background: var(--n-sheet); border-bottom: 1px solid var(--n-line); padding: 12px 16px; display: flex; align-items: center; justify-content: space-between; gap: 10px; }
  h2 { margin: 0; font-size: 16px; }
  h3 { margin: 0 0 6px; font-size: 13px; }
  .body { display: grid; gap: 12px; padding: 16px; }
  p { margin: 0; overflow-wrap: anywhere; color: var(--n-soft); }
  .error { color: var(--platform-colors-danger); margin: 0; }
  .body > form { display: flex; flex-wrap: wrap; gap: 8px; align-items: end; }
  .body > form > label { flex: 1 1 125px; }
  label { display: grid; gap: 4px; }
  input:not([type='checkbox']), select { width: 100%; min-height: 32px; box-sizing: border-box; padding: 4px 8px; border: 1px solid var(--n-line); border-radius: 3px; color: var(--n-ink); background: var(--n-sheet); font: inherit; }
  .fields, .checks { display: grid; grid-template-columns: repeat(auto-fit, minmax(145px, 1fr)); gap: 8px; }
  .checks { max-height: 140px; overflow: auto; }
  .check { display: flex; align-items: center; gap: 6px; overflow-wrap: anywhere; }
  fieldset { border: 1px solid var(--n-line); min-width: 0; padding: 8px; }
  .table { overflow: auto; max-height: 190px; border: 1px solid var(--n-line); }
  table { width: 100%; border-collapse: collapse; }
  td { padding: 4px 8px; border: 1px solid var(--n-fence); max-width: 240px; overflow-wrap: anywhere; }
  footer { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 8px; }
  ul { display: grid; gap: 4px; list-style: none; margin: 0; padding: 0; }
  li { display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
`
