import { useMemo, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import { normalizeRangeToken } from '../lib/workbookModel'

/**
 * Formula bar with the name box: active cell reference, editable range token,
 * formula/value editor with function autocomplete and cell/range reference
 * highlighting, and a computed-value preview (errors surface in red).
 *
 * Reference highlighting uses a mirror layer: while the value is a formula,
 * the input's own text is transparent (caret stays visible) and a
 * typographically identical mirror renders behind it with each unique
 * reference colored from a small palette.
 */

interface FormulaFunctionDoc {
  name: string
  signature: string
  description: string
}

/** The verified 80/20 function set (plus shim-supported extras). */
export const FORMULA_FUNCTIONS: FormulaFunctionDoc[] = [
  { name: 'SUM', signature: 'SUM(range)', description: 'Add the numbers in a range' },
  { name: 'AVERAGE', signature: 'AVERAGE(range)', description: 'Mean of the numbers in a range' },
  { name: 'MIN', signature: 'MIN(range)', description: 'Smallest number in a range' },
  { name: 'MAX', signature: 'MAX(range)', description: 'Largest number in a range' },
  { name: 'COUNT', signature: 'COUNT(range)', description: 'Count numeric cells' },
  { name: 'COUNTA', signature: 'COUNTA(range)', description: 'Count non-empty cells' },
  { name: 'COUNTIF', signature: 'COUNTIF(range, criteria)', description: 'Count cells matching a criterion' },
  { name: 'SUMIF', signature: 'SUMIF(range, criteria, [sum_range])', description: 'Sum cells matching a criterion' },
  { name: 'AVERAGEIF', signature: 'AVERAGEIF(range, criteria, [avg_range])', description: 'Average cells matching a criterion' },
  { name: 'IF', signature: 'IF(condition, then, else)', description: 'Conditional value' },
  { name: 'IFERROR', signature: 'IFERROR(value, fallback)', description: 'Fallback when a value errors' },
  { name: 'ROUND', signature: 'ROUND(number, places)', description: 'Round to a number of places' },
  { name: 'RANK', signature: 'RANK(number, range, [order])', description: 'Rank of a number within a range' },
  { name: 'CONCAT', signature: 'CONCAT(value1, value2, …)', description: 'Join values into one text' },
  { name: 'LEFT', signature: 'LEFT(text, count)', description: 'Leading characters of a text' },
  { name: 'RIGHT', signature: 'RIGHT(text, count)', description: 'Trailing characters of a text' },
  { name: 'VLOOKUP', signature: 'VLOOKUP(key, range, column, [exact])', description: 'Look up a value in the first column' },
  { name: 'TODAY', signature: 'TODAY()', description: "Today's date" },
  { name: 'NOW', signature: 'NOW()', description: 'Current date and time' },
]

const REFERENCE_COLORS = [
  '#1d6fa5',
  '#377d5a',
  '#8a5aa8',
  '#b06b1f',
  '#a8385a',
]

/** Split a formula into text and reference tokens, coloring each unique
 * reference (cell or range, `$` allowed) by order of first appearance. */
export function tokenizeFormulaReferences(
  value: string,
): Array<{ text: string; colorIndex: number | null }> {
  const tokens: Array<{ text: string; colorIndex: number | null }> = []
  const colorByRef = new Map<string, number>()
  const referencePattern =
    /(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)(?![A-Z0-9(])/gi
  const segments = value.split('"')
  segments.forEach((segment, segmentIndex) => {
    if (segmentIndex % 2 === 1) {
      // Inside a quoted string — never highlighted; restore the quotes.
      tokens.push({ text: `"${segment}"`, colorIndex: null })
      return
    }
    let lastIndex = 0
    for (const match of segment.matchAll(referencePattern)) {
      const index = match.index ?? 0
      if (index > lastIndex) {
        tokens.push({ text: segment.slice(lastIndex, index), colorIndex: null })
      }
      const normalized = match[1].toUpperCase().replace(/\$/g, '')
      if (!colorByRef.has(normalized)) {
        colorByRef.set(
          normalized,
          colorByRef.size % REFERENCE_COLORS.length,
        )
      }
      tokens.push({
        text: match[1],
        colorIndex: colorByRef.get(normalized) ?? 0,
      })
      lastIndex = index + match[1].length
    }
    if (lastIndex < segment.length) {
      tokens.push({ text: segment.slice(lastIndex), colorIndex: null })
    }
  })
  return tokens.filter(token => token.text.length > 0)
}

/** Trailing function-name token being typed (e.g. "=SU" → "SU"). */
function autocompleteToken(value: string): string | null {
  if (!value.trimStart().startsWith('=')) return null
  const match = /(?:^=|[\s(,+\-*/&])([A-Z][A-Z0-9]*)$/i.exec(value)
  return match ? match[1] : null
}

export function FormulaBar({
  selectedCell,
  selectedRange,
  editValue,
  selectedDisplay,
  selectedFormulaError,
  onRangeChange,
  onRangeInvalid,
  onEditValueChange,
  onCommitSelectedCell,
}: {
  selectedCell: string
  selectedRange: string
  editValue: string
  selectedDisplay: string
  selectedFormulaError: boolean
  onRangeChange: (range: string) => void
  onRangeInvalid: (message: string) => void
  onEditValueChange: (value: string) => void
  onCommitSelectedCell: (value: string) => void
}): React.ReactElement {
  const mirrorRef = useRef<HTMLDivElement | null>(null)
  const [suggestionIndex, setSuggestionIndex] = useState(0)
  const [suppressSuggestions, setSuppressSuggestions] = useState(false)
  const [focused, setFocused] = useState(false)

  const isFormula = editValue.trimStart().startsWith('=')
  const referenceTokens = useMemo(
    () => (isFormula ? tokenizeFormulaReferences(editValue) : []),
    [editValue, isFormula],
  )

  const token = suppressSuggestions ? null : autocompleteToken(editValue)
  const suggestions = useMemo(() => {
    if (!token || !focused) return []
    const upper = token.toUpperCase()
    return FORMULA_FUNCTIONS.filter(
      fn => fn.name.startsWith(upper) && fn.name !== upper,
    ).slice(0, 8)
  }, [token, focused])
  const activeSuggestion =
    suggestions[Math.min(suggestionIndex, suggestions.length - 1)]

  function acceptSuggestion(name: string): void {
    if (!token) return
    const next = `${editValue.slice(0, editValue.length - token.length)}${name}(`
    setSuppressSuggestions(false)
    setSuggestionIndex(0)
    onEditValueChange(next)
  }

  function handleChange(value: string): void {
    setSuppressSuggestions(false)
    setSuggestionIndex(0)
    onEditValueChange(value)
  }

  return (
    <Bar aria-label="Formula bar">
      <CellReference aria-label="Name box">{selectedCell}</CellReference>
      <RangeInput
        value={selectedRange}
        onChange={event => onRangeChange(event.currentTarget.value)}
        onBlur={() => {
          const normalized = normalizeRangeToken(selectedRange)
          if (normalized) onRangeChange(normalized)
          else onRangeInvalid(`Invalid range ${selectedRange}`)
        }}
        aria-label="Selected range"
      />
      <FormulaStack>
        {isFormula && (
          <HighlightMirror ref={mirrorRef} aria-hidden="true">
            {referenceTokens.map((refToken, index) =>
              refToken.colorIndex === null ? (
                <span key={index}>{refToken.text}</span>
              ) : (
                <ReferenceSpan
                  key={index}
                  data-formula-reference
                  $color={REFERENCE_COLORS[refToken.colorIndex]}
                >
                  {refToken.text}
                </ReferenceSpan>
              ),
            )}
          </HighlightMirror>
        )}
        <FormulaInput
          value={editValue}
          $mirrored={isFormula}
          onChange={event => handleChange(event.currentTarget.value)}
          onFocus={() => setFocused(true)}
          onScroll={event => {
            if (mirrorRef.current) {
              mirrorRef.current.scrollLeft = event.currentTarget.scrollLeft
            }
          }}
          onBlur={() => {
            setFocused(false)
            onCommitSelectedCell(editValue)
          }}
          onKeyDown={event => {
            if (suggestions.length) {
              if (event.key === 'ArrowDown') {
                event.preventDefault()
                setSuggestionIndex(index =>
                  Math.min(index + 1, suggestions.length - 1),
                )
                return
              }
              if (event.key === 'ArrowUp') {
                event.preventDefault()
                setSuggestionIndex(index => Math.max(index - 1, 0))
                return
              }
              if (event.key === 'Tab' || event.key === 'Enter') {
                if (activeSuggestion) {
                  event.preventDefault()
                  acceptSuggestion(activeSuggestion.name)
                  return
                }
              }
              if (event.key === 'Escape') {
                event.preventDefault()
                setSuppressSuggestions(true)
                return
              }
            }
            if (event.key === 'Enter') {
              onCommitSelectedCell(editValue)
              event.currentTarget.blur()
            }
          }}
          aria-label="Formula or value"
          aria-autocomplete="list"
          aria-expanded={suggestions.length > 0}
          autoComplete="off"
          spellCheck={false}
        />
        {suggestions.length > 0 && (
          <SuggestionList role="listbox" aria-label="Formula suggestions">
            {suggestions.map((fn, index) => (
              <SuggestionItem
                key={fn.name}
                type="button"
                role="option"
                aria-selected={fn.name === activeSuggestion?.name}
                $active={fn.name === activeSuggestion?.name}
                // preventDefault so the input keeps focus (blur would commit)
                onMouseDown={event => {
                  event.preventDefault()
                  acceptSuggestion(fn.name)
                }}
                onMouseEnter={() => setSuggestionIndex(index)}
              >
                <SuggestionName>{fn.signature}</SuggestionName>
                <SuggestionDescription>{fn.description}</SuggestionDescription>
              </SuggestionItem>
            ))}
          </SuggestionList>
        )}
      </FormulaStack>
      <FormulaPreview
        $error={selectedFormulaError}
        title={
          selectedFormulaError
            ? `This formula evaluates to ${selectedDisplay}`
            : undefined
        }
      >
        {selectedDisplay !== editValue ? selectedDisplay : ''}
      </FormulaPreview>
    </Bar>
  )
}

const Bar = styled.div`
  display: grid;
  grid-template-columns: 76px 132px minmax(0, 1fr) minmax(120px, 220px) auto;
  gap: 8px;
  align-items: center;
  padding: 6px var(--pure-chrome-inset);
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-bar);

  @media (max-width: 720px) {
    grid-template-columns: 64px minmax(96px, 128px) minmax(160px, 1fr) auto;
  }
`

/** Typed loosely on purpose: styled-components' attrs rejects data-* literals. */
const fieldChrome: Record<string, string> = { 'data-chrome': 'field' }

// The cell reference, formula and range boxes are platform fields (32px,
// hairline, surface) set in the mono face because their content is code.
const CellReference = styled.div.attrs(fieldChrome)`
  border-radius: 0;
  font-family: var(--sheets-mono);
  font-weight: 700;
`

const FormulaStack = styled.div`
  position: relative;
  display: grid;
  min-width: 0;
`

const FormulaInput = styled.input.attrs(fieldChrome)<{ $mirrored?: boolean }>`
  border-radius: 0;
  background: ${({ $mirrored }) =>
    $mirrored ? 'transparent' : 'var(--pure-chrome-surface)'};
  color: ${({ $mirrored }) => ($mirrored ? 'transparent' : 'var(--sheets-ink)')};
  caret-color: var(--sheets-ink);
  font-family: var(--sheets-mono);
  font-size: var(--pure-chrome-ui-size);
  position: relative;
  z-index: 1;

  &::selection {
    background: color-mix(in srgb, var(--pure-chrome-accent) 24%, transparent);
  }
`

/** Behind the input: identical metrics so glyphs line up exactly. */
const HighlightMirror = styled.div`
  position: absolute;
  inset: 0;
  z-index: 0;
  display: flex;
  align-items: center;
  height: var(--pure-chrome-field-height);
  padding: 0 10px;
  overflow: hidden;
  border: 1px solid transparent;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-family: var(--sheets-mono);
  font-size: var(--pure-chrome-ui-size);
  white-space: pre;
  pointer-events: none;
`

const ReferenceSpan = styled.span<{ $color: string }>`
  color: ${({ $color }) => $color};
  font-weight: 700;
`

const RangeInput = styled.input.attrs(fieldChrome)`
  border-radius: 0;
  font-family: var(--sheets-mono);
  font-weight: 700;
`

const SuggestionList = styled.div`
  position: absolute;
  top: calc(100% + 4px);
  left: 0;
  z-index: 30;
  display: grid;
  min-width: 280px;
  max-width: 420px;
  padding: 4px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 0;
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
`

const SuggestionItem = styled.button<{ $active?: boolean }>`
  display: grid;
  gap: 1px;
  padding: 5px 8px;
  border: 0;
  border-radius: 4px;
  background: ${({ $active }) =>
    $active ? 'var(--pure-chrome-selection)' : 'transparent'};
  text-align: left;

  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const SuggestionName = styled.span`
  color: var(--platform-colors-text);
  font-family: var(--sheets-mono);
  font-size: var(--pure-chrome-ui-size);
  font-weight: 700;
`

const SuggestionDescription = styled.span`
  color: var(--pure-chrome-soft);
  font-size: var(--pure-chrome-ui-size);
`

// The live result of the formula is meta: mono 11, muted (red when it errs).
const FormulaPreview = styled(MetaText)<{ $error?: boolean }>`
  overflow: hidden;
  color: ${({ $error }) =>
    $error ? 'var(--platform-colors-danger-text)' : 'var(--pure-chrome-muted)'};
  font-weight: ${({ $error }) => ($error ? 700 : 400)};
  text-overflow: ellipsis;
  white-space: nowrap;

  @media (max-width: 720px) {
    display: none;
  }
`
