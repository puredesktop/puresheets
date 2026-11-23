import JSZip from 'jszip'
import { SaxesParser } from 'saxes'
import { cellKey, computedCellValue, displayCellValue, parseCellKey, rangesIntersect, type FormulaWorkbook } from './workbookModel'
import { parseQualifiedRange, qualifiedRange, validRangeName } from './rangeNames'
import { normalizePivotDefinition } from './sheetMetadataNormalize'
import type { NamedRange, PivotAggregate, WorkbookChart, WorkbookSheet } from '../types'

const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const drawingNs = 'http://schemas.openxmlformats.org/drawingml/2006'
const relDocNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
const header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

export type Element = { name: string; attrs: Record<string, string>; children: Element[]; text: string }

const local = (name: string) => name.slice(name.lastIndexOf(':') + 1)
export const child = (element: Element | undefined, name: string) => element?.children.find(node => node.name === name)
export const children = (element: Element | undefined, name: string) => element?.children.filter(node => node.name === name) ?? []
const attr = (element: Element | undefined, name: string) => element?.attrs[name]

async function part(zip: JSZip, path: string | undefined, keep?: Set<string>): Promise<Element | undefined> {
  if (!path) return
  const source = await zip.file(path)?.async('string')
  if (!source) return
  const root: Element = { name: '', attrs: {}, children: [], text: '' }
  let skipping = 0
  const stack = [root]
  const parser = new SaxesParser({ xmlns: false })
  parser.on('opentag', tag => {
    if (skipping || (keep && stack.length === 2 && !keep.has(local(tag.name)))) { skipping++; return }
    const node = { name: local(tag.name), attrs: tag.attributes, children: [], text: '' }
    stack.at(-1)!.children.push(node)
    stack.push(node)
  })
  parser.on('closetag', () => { if (skipping) skipping--; else stack.pop() })
  parser.on('text', value => { if (!skipping) stack.at(-1)!.text += value })
  try {
    parser.write(source).close()
    return root.children[0]
  } catch {
    return undefined
  }
}

function relativePart(owner: string, target: string): string {
  const segments = target.startsWith('/') ? [] : owner.split('/').slice(0, -1)
  for (const segment of target.split('/')) {
    if (segment === '..') segments.pop()
    else if (segment && segment !== '.') segments.push(segment)
  }
  return segments.join('/')
}

function relsPath(path: string): string {
  const end = path.lastIndexOf('/')
  return `${path.slice(0, end + 1)}_rels/${path.slice(end + 1)}.rels`
}

async function relationships(zip: JSZip, owner: string) {
  return children(await part(zip, relsPath(owner)), 'Relationship').flatMap(({ attrs }) =>
    attrs.Target && attrs.Id && attrs.TargetMode !== 'External'
      ? [{ id: attrs.Id, type: attrs.Type?.split('/').at(-1), path: relativePart(owner, attrs.Target) }]
      : [],
  )
}

const idRef = (element: Element) => Object.entries(element.attrs).find(([name]) => name.endsWith(':id'))?.[1]
const bookPath = 'xl/workbook.xml'

export async function readWorkbookParts(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes)
  const workbook = await part(zip, bookPath)
  const rels = await relationships(zip, bookPath)
  const worksheets = new Map(children(child(workbook, 'sheets'), 'sheet').map(node => [node.attrs.name, rels.find(rel => rel.id === idRef(node))?.path]))
  const cacheRelations = new Map(children(child(workbook, 'pivotCaches'), 'pivotCache').map(node => [node.attrs.cacheId, rels.find(rel => rel.id === idRef(node))?.path]))
  return { zip, workbook, worksheets, cacheRelations }
}

export function importRangeNames(workbook: Element | undefined, sheets: WorkbookSheet[]) {
  const names: NamedRange[] = []
  const seen = new Set<string>()
  let skipped = 0
  for (const entry of children(child(workbook, 'definedNames'), 'definedName')) {
    const name = entry.attrs.name ?? ''
    if (name.startsWith('_xlnm.')) continue
    const ref = parseQualifiedRange(entry.text)
    const sheet = ref && sheets.find(sheet => sheet.name.toLowerCase() === ref.sheet.toLowerCase())
    const key = name.toUpperCase()
    if (entry.attrs.localSheetId !== undefined || !validRangeName(name) || !ref || !sheet || seen.has(key)) { skipped++; continue }
    names.push({ id: `nr-xlsx-${names.length + 1}`, name, sheetId: sheet.id, range: ref.range })
    seen.add(key)
  }
  return { names, skipped }
}

export async function readWorksheetObjects(parts: Awaited<ReturnType<typeof readWorkbookParts>>, sheets: WorkbookSheet[]) {
  let chartCount = 0
  let pivotCount = 0
  const reachableCharts = new Set<string>()
  for (const sheet of sheets) {
    const path = parts.worksheets.get(sheet.name)
    if (!path) continue
    const rels = await relationships(parts.zip, path)
    for (const rel of rels) {
      if (rel.type === 'drawing') {
        for (const chartRel of await relationships(parts.zip, rel.path)) {
          if (chartRel.type !== 'chart' || reachableCharts.has(chartRel.path)) continue
          reachableCharts.add(chartRel.path)
          const chart = importChart(await part(parts.zip, chartRel.path), chartRel.path, sheet)
          if (!chart) continue
          ;(sheet.charts ??= []).push(chart)
          chartCount++
        }
      }
      if (rel.type === 'pivotTable' && !sheet.pivot) {
        const pivot = await importPivot(parts, rel.path, sheet, sheets)
        if (!pivot) continue
        sheet.pivot = pivot
        pivotCount++
      }
    }
  }
  return { chartCount, pivotCount }
}

const chartTypes = { barChart: 'bar', lineChart: 'line', areaChart: 'area', pieChart: 'pie', doughnutChart: 'donut', scatterChart: 'scatter' } as const
function pointValues(node: Element | undefined): string[] {
  const container = node?.children[0]
  const cache = container?.children.find(item => item.name === 'strCache' || item.name === 'numCache') ?? container
  const points = children(cache, 'pt')
  const values: string[] = []
  for (const entry of points) {
    const index = Number(entry.attrs.idx)
    if (Number.isInteger(index) && index >= 0 && index < 25_000) values[index] = child(entry, 'v')?.text ?? ''
  }
  return Array.from({ length: values.length }, (_, index) => values[index] ?? '')
}
function titleOf(node: Element | undefined): string {
  if (!node) return ''
  if (node.name === 't' || node.name === 'v') return node.text
  return node.children.map(titleOf).join('')
}
function importChart(root: Element | undefined, id: string, sheet: WorkbookSheet): WorkbookChart | undefined {
  const chart = child(root, 'chart')
  const plot = child(chart, 'plotArea')
  const plots = plot?.children.filter(item => item.name.endsWith('Chart')) ?? []
  if (plots.length !== 1) return
  const plotType = plots[0]
  const type = chartTypes[plotType.name as keyof typeof chartTypes]
  const seriesList = children(plotType, 'ser')
  if (!type || seriesList.length !== 1) return
  const series = seriesList[0]
  const x = child(series, type === 'scatter' ? 'xVal' : 'cat')
  const y = child(series, type === 'scatter' ? 'yVal' : 'val')
  const labels = pointValues(x)
  const values = pointValues(y).map(value => value === '' ? NaN : Number(value))
  if (!values.length || !values.every(Number.isFinite)) return
  const xRange = parseQualifiedRange(child(x?.children[0], 'f')?.text ?? '')
  const yRange = parseQualifiedRange(child(y?.children[0], 'f')?.text ?? '')
  const sameSheet = (ref: typeof xRange) => ref?.sheet.toLowerCase() === sheet.name.toLowerCase()
  const range = xRange && yRange && sameSheet(xRange) && sameSheet(yRange)
    ? `${xRange.range.split(':')[0]}:${yRange.range.split(':').at(-1)}`
    : ''
  const legend = child(chart, 'legend')
  const axes = [...children(plot, 'catAx'), ...children(plot, 'valAx')]
  return {
    id, title: titleOf(child(chart, 'title')) || 'Chart', type, range,
    labels: values.map((_, index) => labels[index] || String(index + 1)), values,
    legend: legend ? attr(child(legend, 'legendPos'), 'val') === 'b' ? 'bottom' : 'right' : 'none',
    showAxes: axes.some(axis => attr(child(axis, 'delete'), 'val') !== '1'),
    xLabel: titleOf(child(axes[0], 'title')) || undefined,
    yLabel: titleOf(child(axes[1], 'title')) || undefined,
  }
}

async function importPivot(parts: Awaited<ReturnType<typeof readWorkbookParts>>, path: string, sheet: WorkbookSheet, sheets: WorkbookSheet[]) {
  const def = await part(parts.zip, path)
  const row = children(child(def, 'rowFields'), 'field')
  const col = children(child(def, 'colFields'), 'field').filter(node => node.attrs.x !== '-2')
  const data = children(child(def, 'dataFields'), 'dataField')
  const fields = children(child(def, 'pivotFields'), 'pivotField')
  if (!def || row.length !== 1 || col.length > 1 || data.length !== 1 ||
    child(def, 'pageFields') || child(def, 'filters') || fields.some(field => children(child(field, 'items'), 'item').some(item => item.attrs.h === '1'))) return
  const summary = data[0].attrs.subtotal ?? 'sum'
  const aggregate: PivotAggregate | undefined =
    summary === 'avg' || summary === 'average'
      ? 'avg'
      : summary === 'sum' || summary === 'count' || summary === 'min' || summary === 'max'
      ? summary
      : undefined
  if (!aggregate) return
  const cachePath = (await relationships(parts.zip, path)).find(rel => rel.type === 'pivotCacheDefinition')?.path ?? parts.cacheRelations.get(def.attrs.cacheId)
  const cache = await part(parts.zip, cachePath)
  const source = child(child(cache, 'cacheSource'), 'worksheetSource')
  const input = source && sheets.find(sheet => sheet.name === source.attrs.sheet)
  const range = source?.attrs.ref
  const outputRange = child(def, 'location')?.attrs.ref
  const [origin, end = origin] = (outputRange ?? '').split(':').map(parseCellKey)
  const sourceOrigin = range ? parseCellKey(range.split(':')[0]) : null
  const rowIndex = Number(row[0].attrs.x)
  const colIndex = col.length ? Number(col[0].attrs.x) : null
  const valueIndex = Number(data[0].attrs.fld)
  if (!range || !input || !origin || !end || !outputRange || !sourceOrigin ||
    !Number.isInteger(rowIndex) || rowIndex < 0 ||
    !Number.isInteger(valueIndex) || valueIndex < 0 ||
    (colIndex !== null && (!Number.isInteger(colIndex) || colIndex < 0))) return
  const rowsCount = end.row - origin.row + 1
  const colsCount = end.column - origin.column + 1
  if (rowsCount < 1 || colsCount < 1 || rowsCount * colsCount > 25_000) return
  const headerOf = (offset: number) =>
    input.cells[cellKey(sourceOrigin.row, sourceOrigin.column + offset)]?.value.trim() || `Column ${offset + 1}`
  const result = normalizePivotDefinition({
    id: `pivot-xlsx-${path.replace(/[^A-Za-z0-9_-]/g, '-')}`,
    sourceSheetId: input.id,
    sourceRange: range,
    rows: [{ columnIndex: sourceOrigin.column + rowIndex, label: headerOf(rowIndex) }],
    columns: colIndex !== null ? [{ columnIndex: sourceOrigin.column + colIndex, label: headerOf(colIndex) }] : [],
    values: [{ columnIndex: sourceOrigin.column + valueIndex, label: headerOf(valueIndex), aggregate }],
    refreshedAt: new Date().toISOString(),
    output: outputRange.split(':')[0],
    outputRows: rowsCount,
    outputColumns: colsCount,
  })
  if (result && input.id === sheet.id && rangesIntersect(range, outputRange)) return
  return result
}

export async function worksheetValidations(parts: Awaited<ReturnType<typeof readWorkbookParts>>, sheet: string) {
  const xml = await part(parts.zip, parts.worksheets.get(sheet), new Set(['dataValidations']))
  return children(child(xml, 'dataValidations'), 'dataValidation')
}

export const xml = (value: string | number) => String(value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
  .replace(/[&<>"']/g, match => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[match]!)

export async function addChartsToXlsx(bytes: Uint8Array, sheets: { sheet: WorkbookSheet; name: string }[], workbook: FormulaWorkbook): Promise<Uint8Array> {
  if (!sheets.some(({ sheet }) => sheet.charts?.length)) return bytes
  const { zip, worksheets } = await readWorkbookParts(bytes)
  const content = zip.file('[Content_Types].xml')
  if (!content) throw new Error('The Excel package has no content type list.')
  let contentTypes = await content.async('string')
  const override = (name: string, type: string) => {
    contentTypes = contentTypes.replace('</Types>', `<Override PartName="/${name}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/></Types>`)
  }
  let index = 0
  for (const [sheetIndex, { sheet, name }] of sheets.entries()) {
    if (!sheet.charts?.length) continue
    const owner = worksheets.get(name)
    if (!owner) throw new Error(`Could not attach charts to ${name}.`)
    const drawingName = `drawing${sheetIndex + 1}.xml`
    const drawingPath = `xl/drawings/${drawingName}`
    const anchors: string[] = []
    const refs: string[] = []
    const baseColumn = Math.min(25, Math.max(4, ...sheet.charts.map(chart => (parseCellKey(chart.range.split(':').at(-1) ?? '')?.column ?? 0) + 2)))
    for (const [position, chart] of sheet.charts.entries()) {
      index++
      const id = `rId${index}`
      const chartPath = `xl/charts/chart${index}.xml`
      zip.file(chartPath, chartXml(chart, chartReferences(sheet, name, chart, workbook)))
      override(chartPath, 'drawingml.chart')
      refs.push(`<Relationship Id="${id}" Type="${relNs}/chart" Target="../charts/chart${index}.xml"/>`)
      anchors.push(`<xdr:oneCellAnchor><xdr:from><xdr:col>${baseColumn}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${position * 16}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="5700000" cy="2900000"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${index}" name="${xml(chart.title || 'Chart')}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${drawingNs}/chart"><c:chart xmlns:c="${drawingNs}/chart" r:id="${id}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`)
    }
    zip.file(drawingPath, `${header}<xdr:wsDr xmlns:xdr="${drawingNs}/spreadsheetDrawing" xmlns:a="${drawingNs}/main" xmlns:r="${relNs}">${anchors.join('')}</xdr:wsDr>`)
    override(drawingPath, 'drawing')
    zip.file(relsPath(drawingPath), `${header}<Relationships xmlns="${relDocNs}">${refs.join('')}</Relationships>`)
    const ownerRels = relsPath(owner)
    const previous = await zip.file(ownerRels)?.async('string') ?? `${header}<Relationships xmlns="${relDocNs}"></Relationships>`
    const relation = '<Relationship Id="rIdPureSheetsDrawing" Type="' + relNs + '/drawing" Target="../drawings/' + drawingName + '"/>'
    zip.file(ownerRels, previous.replace('</Relationships>', relation + '</Relationships>'))
    const sheetXml = await zip.file(owner)!.async('string')
    zip.file(owner, sheetXml.replace(/<legacyDrawing\b|<extLst\b|<\/worksheet>/, point => `<drawing xmlns:r="${relNs}" r:id="rIdPureSheetsDrawing"/>${point}`))
  }
  zip.file('[Content_Types].xml', contentTypes)
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

function chartReferences(sheet: WorkbookSheet, name: string, chart: WorkbookChart, workbook: FormulaWorkbook) {
  const [start, end = start] = chart.range.split(':').map(parseCellKey)
  if (!start || !end || end.column <= start.column || !chart.values.length) return {}
  // Link a chart to its sheet when its plotted points match contiguous cells.
  // Scalar chart snapshots and gapped data keep their saved series values.
  const row = end.row - chart.values.length + 1
  if (row < start.row) return {}
  const compatible = chart.values.every((amount, index) => {
    const y = computedCellValue(sheet, cellKey(row + index, start.column + 1), workbook)
    const x = displayCellValue(sheet, cellKey(row + index, start.column), workbook)
    return y.trim() !== '' && Number(y) === amount && x === chart.labels[index]
  })
  if (!compatible) return {}
  const address = (column: number) => qualifiedRange(name, cellKey(row, column) + ':' + cellKey(end.row, column))
  return { x: address(start.column), y: address(start.column + 1) }
}

function chartXml(chart: WorkbookChart, refs: { x?: string; y?: string }): string {
  const tag = Object.entries(chartTypes).find(([, type]) => type === chart.type)![0]
  const pie = chart.type === 'pie' || chart.type === 'donut'
  const scatter = chart.type === 'scatter'
  const seriesValues = (values: Array<string | number>, numeric: boolean, ref?: string) => {
    const data = (tag: string) => `<c:${tag}>${numeric ? '<c:formatCode>General</c:formatCode>' : ''}<c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${xml(value)}</c:v></c:pt>`).join('')}</c:${tag}>`
    const type = numeric ? 'num' : 'str'
    return ref ? `<c:${type}Ref><c:f>${xml(ref)}</c:f>${data(type + 'Cache')}</c:${type}Ref>` : data(type + 'Lit')
  }
  const title = (value?: string) => value ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xml(value)}</a:t></a:r></a:p></c:rich></c:tx></c:title>` : ''
  const axis = (id: number, cross: number, x: boolean) => {
    const tag = x && !scatter ? 'catAx' : 'valAx'
    return `<c:${tag}><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${chart.showAxes === false ? 1 : 0}"/><c:axPos val="${x ? 'b' : 'l'}"/>${title(x ? chart.xLabel : chart.yLabel)}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/></c:${tag}>`
  }
  const plotOptions = chart.type === 'bar' ? '<c:barDir val="col"/><c:grouping val="clustered"/>' :
    chart.type === 'line' || chart.type === 'area' ? '<c:grouping val="standard"/>' : scatter ? '<c:scatterStyle val="marker"/>' : ''
  const domain = scatter
    ? `<c:xVal>${seriesValues(chart.labels.map((label, index) => Number.isFinite(Number(label)) ? Number(label) : index + 1), true, refs.x)}</c:xVal>`
    : `<c:cat>${seriesValues(chart.labels, false, refs.x)}</c:cat>`
  const values = scatter ? 'yVal' : 'val'
  const series = `<c:ser><c:idx val="0"/><c:order val="0"/>${domain}<c:${values}>${seriesValues(chart.values, true, refs.y)}</c:${values}></c:ser>`
  const ids = pie ? '' : '<c:axId val="10"/><c:axId val="20"/>'
  const extra = chart.type === 'donut' ? '<c:holeSize val="55"/>' : ''
  return `${header}<c:chartSpace xmlns:c="${drawingNs}/chart" xmlns:a="${drawingNs}/main"><c:chart>${title(chart.title)}<c:plotArea><c:layout/><c:${tag}>${plotOptions}${series}${extra}${ids}</c:${tag}>${pie ? '' : axis(10, 20, true) + axis(20, 10, false)}</c:plotArea>${chart.legend && chart.legend !== 'none' ? `<c:legend><c:legendPos val="${chart.legend === 'bottom' ? 'b' : 'r'}"/></c:legend>` : ''}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`
}
