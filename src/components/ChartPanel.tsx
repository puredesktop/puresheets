import { useEffect, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import { Maximize2, Pencil, X } from 'lucide-react'
import { ToolIconButton } from './controls'
import type { WorkbookChart, WorkbookChartType } from '../types'

/**
 * Chart shelf for the active sheet (Phase S4): bar, line, area, pie, donut,
 * and scatter charts with an inline editor (range, title, type, legend
 * placement, axis labels). Chart data lives in the `.sheets` document —
 * PureSheets-owned metadata, persisted alongside the snapshot.
 */

const SERIES_COLORS = [
  '#377d5a',
  '#1d6fa5',
  '#8a5aa8',
  '#b06b1f',
  '#a8385a',
  '#5b7161',
  '#946f43',
  '#3f6f8f',
]

export type ChartUpdatePatch = Partial<
  Pick<
    WorkbookChart,
    'title' | 'type' | 'range' | 'legend' | 'showAxes' | 'xLabel' | 'yLabel'
  >
>

const CHART_TYPE_OPTIONS: Array<{ value: WorkbookChartType; label: string }> = [
  { value: 'bar', label: 'Bar' },
  { value: 'line', label: 'Line' },
  { value: 'area', label: 'Area' },
  { value: 'pie', label: 'Pie' },
  { value: 'donut', label: 'Donut' },
  { value: 'scatter', label: 'Scatter' },
]

const LEGEND_OPTIONS: Array<{
  value: NonNullable<WorkbookChart['legend']>
  label: string
}> = [
  { value: 'right', label: 'Legend right' },
  { value: 'bottom', label: 'Legend bottom' },
  { value: 'none', label: 'No legend' },
]

export function ChartPanel({
  charts,
  onDeleteChart,
  onUpdateChart,
}: {
  charts: WorkbookChart[]
  onDeleteChart: (chartId: string) => void
  onUpdateChart: (chartId: string, patch: ChartUpdatePatch) => void
}): React.ReactElement | null {
  if (!charts.length) return null
  return (
    <ChartShelf aria-label="Workbook charts">
      {charts.map(chart => (
        <ChartCard key={chart.id}>
          <ChartCardBody
            chart={chart}
            onDeleteChart={onDeleteChart}
            onUpdateChart={onUpdateChart}
          />
        </ChartCard>
      ))}
    </ChartShelf>
  )
}

function ChartCardBody({
  chart,
  onDeleteChart,
  onUpdateChart,
}: {
  chart: WorkbookChart
  onDeleteChart: (chartId: string) => void
  onUpdateChart: (chartId: string, patch: ChartUpdatePatch) => void
}): React.ReactElement {
  const [editing, setEditing] = useState(false)
  const [maximized, setMaximized] = useState(false)

  const legend = chart.legend ?? (isRadial(chart.type) ? 'right' : 'none')

  return (
    <>
      <ChartHeader>
        <div>
          <ChartTitle>{chart.title}</ChartTitle>
          <ChartMeta>
            {chart.range} · {chart.type}
          </ChartMeta>
        </div>
        <ChartHeaderActions>
          <ToolIconButton
            title="Maximize chart"
            onClick={() => setMaximized(true)}
            icon={Maximize2}
          />
          <ToolIconButton
            title="Edit chart"
            onClick={() => setEditing(value => !value)}
            icon={Pencil}
          />
          <ToolIconButton
            title="Remove chart"
            onClick={() => onDeleteChart(chart.id)}
            icon={X}
          />
        </ChartHeaderActions>
      </ChartHeader>

      {maximized && (
        <ChartLightbox
          chart={chart}
          onUpdateChart={onUpdateChart}
          onClose={() => setMaximized(false)}
        />
      )}

      <ChartWithLegend $legend={legend}>
        <MiniChart chart={chart} />
        {legend !== 'none' && (
          <ChartLegend
            $placement={legend === 'right' ? 'right' : 'bottom'}
            aria-label="Chart legend"
          >
            {chart.labels.slice(0, 8).map((label, index) => (
              <LegendItem key={`${label}-${index}`}>
                <LegendSwatch
                  $color={SERIES_COLORS[index % SERIES_COLORS.length]}
                />
                <span>{label}</span>
              </LegendItem>
            ))}
          </ChartLegend>
        )}
      </ChartWithLegend>

      {(chart.xLabel || chart.yLabel) && chart.showAxes !== false && (
        <AxisCaption>
          {chart.xLabel ? `x: ${chart.xLabel}` : ''}
          {chart.xLabel && chart.yLabel ? ' · ' : ''}
          {chart.yLabel ? `y: ${chart.yLabel}` : ''}
        </AxisCaption>
      )}

      {editing && (
        <ChartEditorForm
          chart={chart}
          onUpdateChart={onUpdateChart}
          onDone={() => setEditing(false)}
        />
      )}
    </>
  )
}

// Shared range/title/type/legend/axis editor, rendered both under the shelf
// card and inside the maximized lightbox. It mounts fresh whenever the editor
// is opened, so the draft fields seed from the current chart each time.
function ChartEditorForm({
  chart,
  onUpdateChart,
  onDone,
}: {
  chart: WorkbookChart
  onUpdateChart: (chartId: string, patch: ChartUpdatePatch) => void
  onDone: () => void
}): React.ReactElement {
  const [titleDraft, setTitleDraft] = useState(chart.title)
  const [rangeDraft, setRangeDraft] = useState(chart.range)
  const [xLabelDraft, setXLabelDraft] = useState(chart.xLabel ?? '')
  const [yLabelDraft, setYLabelDraft] = useState(chart.yLabel ?? '')

  const legend = chart.legend ?? (isRadial(chart.type) ? 'right' : 'none')

  function applyEdits(): void {
    onUpdateChart(chart.id, {
      title: titleDraft.trim() || chart.title,
      range: rangeDraft.trim() || chart.range,
      xLabel: xLabelDraft.trim() || undefined,
      yLabel: yLabelDraft.trim() || undefined,
    })
    onDone()
  }

  return (
    <ChartEditor aria-label={`Edit ${chart.title}`}>
      <EditorRow>
        <EditorInput
          value={titleDraft}
          onChange={event => setTitleDraft(event.currentTarget.value)}
          placeholder="Chart title"
          aria-label="Chart title"
        />
        <EditorInput
          value={rangeDraft}
          onChange={event => setRangeDraft(event.currentTarget.value)}
          placeholder="Range (e.g. A2:B6)"
          aria-label="Chart range"
          $mono
        />
      </EditorRow>
      <OptionRow role="radiogroup" aria-label="Chart type">
        {CHART_TYPE_OPTIONS.map(option => (
          <OptionButton
            key={option.value}
            type="button"
            role="radio"
            aria-checked={chart.type === option.value}
            $active={chart.type === option.value}
            onClick={() => onUpdateChart(chart.id, { type: option.value })}
          >
            {option.label}
          </OptionButton>
        ))}
      </OptionRow>
      <OptionRow role="radiogroup" aria-label="Legend placement">
        {LEGEND_OPTIONS.map(option => (
          <OptionButton
            key={option.value}
            type="button"
            role="radio"
            aria-checked={legend === option.value}
            $active={legend === option.value}
            onClick={() => onUpdateChart(chart.id, { legend: option.value })}
          >
            {option.label}
          </OptionButton>
        ))}
        <OptionButton
          type="button"
          aria-pressed={chart.showAxes !== false}
          $active={chart.showAxes !== false}
          onClick={() =>
            onUpdateChart(chart.id, {
              showAxes: chart.showAxes === false,
            })
          }
        >
          Axis labels
        </OptionButton>
      </OptionRow>
      <EditorRow>
        <EditorInput
          value={xLabelDraft}
          onChange={event => setXLabelDraft(event.currentTarget.value)}
          placeholder="X axis label"
          aria-label="X axis label"
        />
        <EditorInput
          value={yLabelDraft}
          onChange={event => setYLabelDraft(event.currentTarget.value)}
          placeholder="Y axis label"
          aria-label="Y axis label"
        />
      </EditorRow>
      <EditorActions>
        <ApplyButton type="button" onClick={applyEdits}>
          Apply
        </ApplyButton>
        <QuietButton type="button" onClick={onDone}>
          Close
        </QuietButton>
      </EditorActions>
    </ChartEditor>
  )
}

function ChartLightbox({
  chart,
  onUpdateChart,
  onClose,
}: {
  chart: WorkbookChart
  onUpdateChart: (chartId: string, patch: ChartUpdatePatch) => void
  onClose: () => void
}): React.ReactElement {
  // Enlarged view of a chart (#276): the shelf card is small, so let the user
  // pop a chart open at a legible size. Escape / backdrop click / the close
  // button all dismiss it; the chart is redrawn at a larger geometry so bars,
  // points, and axis labels spread out instead of just being scaled up.
  const [editing, setEditing] = useState(false)

  useEffect(() => {
    function onKey(event: KeyboardEvent): void {
      if (event.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  const legend = chart.legend ?? (isRadial(chart.type) ? 'right' : 'none')

  return (
    <LightboxBackdrop
      role="dialog"
      aria-modal="true"
      aria-label={`${chart.title} enlarged`}
      onClick={onClose}
    >
      <LightboxCard onClick={event => event.stopPropagation()}>
        <LightboxHeader>
          <div>
            <LightboxTitle>{chart.title}</LightboxTitle>
            <ChartMeta>
              {chart.range} · {chart.type}
            </ChartMeta>
          </div>
          <ChartHeaderActions>
            <ToolIconButton
              title="Edit chart"
              onClick={() => setEditing(value => !value)}
              icon={Pencil}
            />
            <ToolIconButton title="Close" onClick={onClose} icon={X} />
          </ChartHeaderActions>
        </LightboxHeader>

        <LightboxBody $legend={legend}>
          <MiniChart chart={chart} width={760} height={440} />
          {legend !== 'none' && (
            <ChartLegend
              $placement={legend === 'right' ? 'right' : 'bottom'}
              aria-label="Chart legend"
            >
              {chart.labels.slice(0, 16).map((label, index) => (
                <LegendItem key={`${label}-${index}`}>
                  <LegendSwatch
                    $color={SERIES_COLORS[index % SERIES_COLORS.length]}
                  />
                  <span>{label}</span>
                </LegendItem>
              ))}
            </ChartLegend>
          )}
        </LightboxBody>

        {(chart.xLabel || chart.yLabel) && chart.showAxes !== false && (
          <AxisCaption>
            {chart.xLabel ? `x: ${chart.xLabel}` : ''}
            {chart.xLabel && chart.yLabel ? ' · ' : ''}
            {chart.yLabel ? `y: ${chart.yLabel}` : ''}
          </AxisCaption>
        )}

        {editing && (
          <ChartEditorForm
            chart={chart}
            onUpdateChart={onUpdateChart}
            onDone={() => setEditing(false)}
          />
        )}
      </LightboxCard>
    </LightboxBackdrop>
  )
}

function isRadial(type: WorkbookChartType): boolean {
  return type === 'pie' || type === 'donut'
}

export function MiniChart({
  chart,
  width = 280,
  height = 104,
}: {
  chart: WorkbookChart
  width?: number
  height?: number
}): React.ReactElement {
  if (isRadial(chart.type)) {
    return <RadialChart chart={chart} width={width} height={height} />
  }
  if (chart.type === 'scatter') {
    return <ScatterChart chart={chart} width={width} height={height} />
  }
  const max = Math.max(...chart.values, 1)
  const points = chart.values.map((value, index) => {
    const x =
      chart.values.length <= 1
        ? width / 2
        : 18 + (index / (chart.values.length - 1)) * (width - 36)
    const y = height - 16 - (value / max) * (height - 28)
    return { x, y, value, label: chart.labels[index] ?? `${index + 1}` }
  })
  const showAxisText = chart.showAxes !== false
  if (chart.type === 'line' || chart.type === 'area') {
    const linePoints = points.map(point => `${point.x},${point.y}`).join(' ')
    return (
      <ChartSvg
        viewBox={`0 0 ${width} ${height}`}
        role="img"
        aria-label={chart.title}
      >
        {chart.type === 'area' && points.length > 1 && (
          <polygon
            fill="#377d5a"
            opacity="0.18"
            points={`${points[0].x},${height - 16} ${linePoints} ${
              points[points.length - 1].x
            },${height - 16}`}
          />
        )}
        <polyline
          fill="none"
          stroke="#377d5a"
          strokeWidth="3"
          points={linePoints}
        />
        {points.map(point => (
          <g key={`${point.label}-${point.x}`}>
            <circle cx={point.x} cy={point.y} r="4" fill="#377d5a" />
            {showAxisText && (
              <text x={point.x} y={height - 2} textAnchor="middle">
                {point.label}
              </text>
            )}
          </g>
        ))}
      </ChartSvg>
    )
  }
  // bar
  const barWidth = Math.max(
    14,
    Math.min(34, (width - 32) / Math.max(points.length, 1) - 8),
  )
  return (
    <ChartSvg
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label={chart.title}
    >
      {points.map(point => {
        // Measure the bar from the same baseline the point `y` uses
        // (height - 16); the old `height - 18` was 2px lower, so tiny/zero
        // bars produced a slightly negative height and an invalid <rect>.
        // Clamp keeps negative data values from re-introducing it.
        const barHeight = Math.max(0, height - 16 - point.y)
        return (
          <g key={`${point.label}-${point.x}`}>
            <rect
              x={point.x - barWidth / 2}
              y={point.y}
              width={barWidth}
              height={barHeight}
              rx="3"
              fill="#377d5a"
            />
            {showAxisText && (
              <text x={point.x} y={height - 2} textAnchor="middle">
                {point.label}
              </text>
            )}
          </g>
        )
      })}
    </ChartSvg>
  )
}

function RadialChart({
  chart,
  width,
  height,
}: {
  chart: WorkbookChart
  width: number
  height: number
}): React.ReactElement {
  const total = chart.values.reduce(
    (sum, value) => sum + Math.max(0, value),
    0,
  )
  const cx = width / 2
  const cy = height / 2
  const radius = Math.min(width, height) / 2 - 6
  const innerRadius = chart.type === 'donut' ? radius * 0.55 : 0
  let angle = -Math.PI / 2
  const slices = chart.values.map((value, index) => {
    const share = total > 0 ? Math.max(0, value) / total : 0
    const start = angle
    const end = angle + share * Math.PI * 2
    angle = end
    return {
      start,
      end,
      color: SERIES_COLORS[index % SERIES_COLORS.length],
      label: chart.labels[index] ?? `${index + 1}`,
    }
  })
  return (
    <ChartSvg
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label={chart.title}
    >
      {slices.map((slice, index) => (
        <path
          key={`${slice.label}-${index}`}
          d={arcPath(cx, cy, radius, innerRadius, slice.start, slice.end)}
          fill={slice.color}
        >
          <title>{slice.label}</title>
        </path>
      ))}
    </ChartSvg>
  )
}

function arcPath(
  cx: number,
  cy: number,
  outer: number,
  inner: number,
  start: number,
  end: number,
): string {
  const clampedEnd = Math.min(end, start + Math.PI * 2 - 0.0001)
  const largeArc = clampedEnd - start > Math.PI ? 1 : 0
  const sx = cx + outer * Math.cos(start)
  const sy = cy + outer * Math.sin(start)
  const ex = cx + outer * Math.cos(clampedEnd)
  const ey = cy + outer * Math.sin(clampedEnd)
  if (inner <= 0) {
    return `M ${cx} ${cy} L ${sx} ${sy} A ${outer} ${outer} 0 ${largeArc} 1 ${ex} ${ey} Z`
  }
  const isx = cx + inner * Math.cos(clampedEnd)
  const isy = cy + inner * Math.sin(clampedEnd)
  const iex = cx + inner * Math.cos(start)
  const iey = cy + inner * Math.sin(start)
  return [
    `M ${sx} ${sy}`,
    `A ${outer} ${outer} 0 ${largeArc} 1 ${ex} ${ey}`,
    `L ${isx} ${isy}`,
    `A ${inner} ${inner} 0 ${largeArc} 0 ${iex} ${iey}`,
    'Z',
  ].join(' ')
}

function ScatterChart({
  chart,
  width,
  height,
}: {
  chart: WorkbookChart
  width: number
  height: number
}): React.ReactElement {
  // x comes from the label column (numeric); fall back to the index.
  const points = chart.values.map((value, index) => {
    const numericLabel = Number(chart.labels[index])
    return {
      x: Number.isFinite(numericLabel) ? numericLabel : index,
      y: value,
    }
  })
  const xValues = points.map(point => point.x)
  const yValues = points.map(point => point.y)
  const xMin = Math.min(...xValues, 0)
  const xMax = Math.max(...xValues, 1)
  const yMin = Math.min(...yValues, 0)
  const yMax = Math.max(...yValues, 1)
  const plotX = (x: number): number =>
    14 + ((x - xMin) / Math.max(xMax - xMin, 1e-9)) * (width - 28)
  const plotY = (y: number): number =>
    height - 14 - ((y - yMin) / Math.max(yMax - yMin, 1e-9)) * (height - 26)
  return (
    <ChartSvg
      viewBox={`0 0 ${width} ${height}`}
      role="img"
      aria-label={chart.title}
    >
      {chart.showAxes !== false && (
        <>
          <line
            x1={14}
            y1={height - 14}
            x2={width - 10}
            y2={height - 14}
            stroke="#c7cfc0"
          />
          <line x1={14} y1={6} x2={14} y2={height - 14} stroke="#c7cfc0" />
        </>
      )}
      {points.map((point, index) => (
        <circle
          key={index}
          cx={plotX(point.x)}
          cy={plotY(point.y)}
          r="3.5"
          fill="#1d6fa5"
          opacity="0.85"
        />
      ))}
    </ChartSvg>
  )
}

const ChartShelf = styled.section`
  display: flex;
  gap: 10px;
  overflow-x: auto;
  padding: 10px 12px;
  border-top: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-well);
`

const ChartCard = styled.article`
  display: grid;
  gap: 8px;
  min-width: 260px;
  max-width: 340px;
  padding: 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius);
  background: var(--pure-chrome-surface);
`

const LightboxBackdrop = styled.div`
  position: fixed;
  inset: 0;
  z-index: 1000;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 24px;
  background: rgba(15, 20, 16, 0.55);
`

const LightboxCard = styled.article`
  display: grid;
  gap: 12px;
  width: min(920px, 100%);
  max-height: 100%;
  padding: 18px 20px;
  overflow: auto;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 12px;
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-lg, var(--platform-shadow-md));
`

const LightboxHeader = styled.div`
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: start;
  gap: 12px;
`

const LightboxTitle = styled.div`
  color: var(--platform-colors-text);
  font-size: 16px;
  font-weight: 700;
`

const LightboxBody = styled.div<{ $legend: 'right' | 'bottom' | 'none' }>`
  --chart-svg-height: min(60vh, 440px);
  display: flex;
  flex-direction: ${({ $legend }) => ($legend === 'right' ? 'row' : 'column')};
  gap: 16px;
  align-items: ${({ $legend }) =>
    $legend === 'right' ? 'center' : 'stretch'};
`

const ChartHeader = styled.div`
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  align-items: start;
  gap: 8px;
`

const ChartHeaderActions = styled.div`
  display: inline-flex;
  gap: 4px;
`

const ChartTitle = styled.div`
  overflow: hidden;
  color: var(--platform-colors-text);
  font-size: var(--pure-chrome-ui-size);
  font-weight: 600;
  text-overflow: ellipsis;
  white-space: nowrap;
`

const ChartMeta = styled(MetaText)`
  display: block;
  white-space: normal;
`

const ChartWithLegend = styled.div<{ $legend: 'right' | 'bottom' | 'none' }>`
  display: flex;
  flex-direction: ${({ $legend }) => ($legend === 'right' ? 'row' : 'column')};
  gap: 8px;
  align-items: ${({ $legend }) =>
    $legend === 'right' ? 'center' : 'stretch'};
`

const ChartLegend = styled.div<{ $placement: 'right' | 'bottom' }>`
  display: flex;
  flex-direction: ${({ $placement }) =>
    $placement === 'right' ? 'column' : 'row'};
  flex-wrap: wrap;
  gap: 4px 10px;
  max-width: ${({ $placement }) => ($placement === 'right' ? '96px' : 'none')};
`

const LegendItem = styled.span`
  display: inline-flex;
  gap: 5px;
  align-items: center;
  min-width: 0;
  color: var(--platform-colors-text-muted, #687064);
  font-size: 10px;

  > span {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
`

const LegendSwatch = styled.span<{ $color: string }>`
  flex: 0 0 auto;
  width: 8px;
  height: 8px;
  border-radius: 2px;
  background: ${({ $color }) => $color};
`

const AxisCaption = styled.div`
  color: var(--platform-colors-text-muted, #687064);
  font-size: 10px;
`

const ChartEditor = styled.div`
  display: grid;
  gap: 6px;
  padding-top: 6px;
  border-top: 1px solid var(--pure-chrome-line);
`

const EditorRow = styled.div`
  display: flex;
  gap: 6px;
`

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const fieldChrome: Record<string, string> = { 'data-chrome': 'field' }

const EditorInput = styled.input.attrs(fieldChrome)<{ $mono?: boolean }>`
  width: 0;
  flex: 1 1 0;
  ${({ $mono }) => ($mono ? 'font-family: var(--sheets-mono);' : '')}
`

const OptionRow = styled.div`
  display: flex;
  flex-wrap: wrap;
  gap: 4px;
`

const OptionButton = styled.button<{ $active?: boolean }>`
  height: var(--pure-chrome-control-height);
  padding: 0 8px;
  border: 1px solid
    ${({ $active }) => ($active ? 'var(--pure-chrome-soft)' : 'var(--pure-chrome-line)')};
  border-radius: 7px;
  background: ${({ $active }) =>
    $active ? 'var(--pure-chrome-selection)' : 'var(--pure-chrome-surface)'};
  color: var(--platform-colors-text);
  font-size: 12px;
  font-weight: ${({ $active }) => ($active ? 600 : 500)};
`

const EditorActions = styled.div`
  display: flex;
  gap: 6px;
`

const ApplyButton = styled.button`
  height: var(--pure-chrome-control-height);
  padding: 0 10px;
  border: 1px solid var(--pure-chrome-accent);
  border-radius: 7px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-size: 12px;
  font-weight: 600;
`

const QuietButton = styled.button`
  height: var(--pure-chrome-control-height);
  padding: 0 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 7px;
  background: var(--pure-chrome-surface);
  color: var(--platform-colors-text);
  font-size: 12px;
  font-weight: 500;
`

const ChartSvg = styled.svg`
  width: 100%;
  height: var(--chart-svg-height, 104px);
  flex: 1 1 auto;
  min-width: 0;

  text {
    fill: var(--platform-colors-text-muted, #687064);
    font-size: 9px;
  }
`
