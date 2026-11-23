import { useMemo, useRef, useState } from 'react'
import type React from 'react'
import styled from 'styled-components'
import { ChevronDown, ChevronUp, Sparkles } from 'lucide-react'
import { PlatformIcon } from '@purescience/platform-ui/components/chrome/PlatformIcon'
import { MetaText } from '@purescience/platform-ui/components/common/containers/AppChrome'
import type { NamedRange } from '../types'
import {
  activeCallContext,
  explainFormulaError,
  FORMULA_CATALOG,
  type FormulaFunctionDoc,
} from '../lib/formulaCatalog'
import { normalizeRangeToken } from '../lib/workbookModel'

/**
 * Next-gen Formula Bar (Wave 2):
 * - Name Box + Range input + optional Named Range quick-jump
 * - Expandable multiline formula editor for complex nested formulas
 * - Full 180+ Excel/Univer function catalog autocomplete with category badges
 * - Live active-parameter signature pill highlighting the argument being typed
 * - Plain-English error diagnosis popover with one-click IFERROR fix
 * - Reference color highlighting (preserves mirror layer and all existing aria/test hooks)
 */
export const FORMULA_FUNCTIONS: FormulaFunctionDoc[] = FORMULA_CATALOG

const REFERENCE_COLORS = [
  '#1d6fa5',
  '#377d5a',
  '#8a5aa8',
  '#b06b1f',
  '#a8385a',
]

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

function autocompleteToken(value: string): string | null {
  if (!value.trimStart().startsWith('=')) return null
  const match = /(?:^=|[\s(,+\-*/&])([A-Z][A-Z0-9.]*)$/i.exec(value)
  return match ? match[1] : null
}

export function FormulaBar({
  selectedCell,
  selectedRange,
  editValue,
  selectedDisplay,
  selectedFormulaError,
  selectedHyperlink,
  selectedCheckbox,
  onToggleCheckboxValue,
  namedRanges = [],
  onSelectNamedRange,
  onRangeChange,
  onRangeInvalid,
  onEditValueChange,
  onCommitSelectedCell,
  onCancelEdit,
}: {
  selectedCell: string
  selectedRange: string
  editValue: string
  selectedDisplay: string
  selectedFormulaError: boolean
  selectedHyperlink?: string
  selectedCheckbox?: boolean
  onToggleCheckboxValue?: () => void
  namedRanges?: NamedRange[]
  onSelectNamedRange?: (nr: NamedRange) => void
  onRangeChange: (range: string) => void
  onRangeInvalid: (message: string) => void
  onEditValueChange: (value: string) => void
  onCommitSelectedCell: (value: string) => void
  onCancelEdit?: () => void
}): React.ReactElement {
  const mirrorRef = useRef<HTMLDivElement | null>(null)
  const skipNextBlurCommitRef = useRef(false)
  const rangeBeforeFocusRef = useRef(selectedRange)
  const skipNextRangeBlurRef = useRef(false)
  const [suggestionIndex, setSuggestionIndex] = useState(0)
  const [suppressSuggestions, setSuppressSuggestions] = useState(false)
  const [focused, setFocused] = useState(false)
  const [expanded, setExpanded] = useState(false)
  const [cursorPos, setCursorPos] = useState(editValue.length)
  const [namesOpen, setNamesOpen] = useState(false)
  const [errorPopoverOpen, setErrorPopoverOpen] = useState(false)

  const isFormula = editValue.trimStart().startsWith('=')
  const referenceTokens = useMemo(
    () => (isFormula ? tokenizeFormulaReferences(editValue) : []),
    [editValue, isFormula],
  )

  const token = suppressSuggestions ? null : autocompleteToken(editValue)
  const suggestions = useMemo(() => {
    if (!token || !focused) return []
    const upper = token.toUpperCase()
    const fnMatches = FORMULA_FUNCTIONS.filter(
      fn => fn.name.startsWith(upper) && fn.name !== upper,
    ).slice(0, 8)
    return fnMatches
  }, [token, focused])
  const activeSuggestion =
    suggestions[Math.min(suggestionIndex, suggestions.length - 1)]

  const callContext = useMemo(
    () =>
      focused && isFormula && suggestions.length === 0
        ? activeCallContext(editValue, cursorPos)
        : null,
    [focused, isFormula, suggestions.length, editValue, cursorPos],
  )

  const errorExplanation = useMemo(
    () =>
      selectedFormulaError
        ? explainFormulaError(selectedDisplay, editValue)
        : null,
    [selectedFormulaError, selectedDisplay, editValue],
  )

  function acceptSuggestion(name: string): void {
    if (!token) return
    const next = `${editValue.slice(0, editValue.length - token.length)}${name}(`
    setSuppressSuggestions(false)
    setSuggestionIndex(0)
    setCursorPos(next.length)
    onEditValueChange(next)
  }

  function handleChange(value: string, pos?: number): void {
    setSuppressSuggestions(false)
    setSuggestionIndex(0)
    setCursorPos(pos ?? value.length)
    onEditValueChange(value)
  }

  function handleKeyDown(
    event: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>,
  ): void {
    if (suggestions.length) {
      if (event.key === 'ArrowDown') {
        event.preventDefault()
        setSuggestionIndex(index => Math.min(index + 1, suggestions.length - 1))
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
    if (event.key === 'Enter' && (!expanded || !event.shiftKey)) {
      event.preventDefault()
      skipNextBlurCommitRef.current = true
      onCommitSelectedCell(editValue)
      event.currentTarget.blur()
    } else if (event.key === 'Escape') {
      event.preventDefault()
      skipNextBlurCommitRef.current = true
      onCancelEdit?.()
      event.currentTarget.blur()
    }
  }

  return (
    <Bar aria-label="Formula bar" $expanded={expanded}>
      <NameBoxWrap>
        <CellReference aria-label="Name box">{selectedCell}</CellReference>
        {namedRanges.length > 0 && onSelectNamedRange && (
          <NamesMenuWrap>
            <NamesTrigger
              type="button"
              title="Jump to named range"
              aria-label="Named ranges"
              aria-expanded={namesOpen}
              onClick={() => setNamesOpen(o => !o)}
            >
              <PlatformIcon icon={ChevronDown} size={12} />
            </NamesTrigger>
            {namesOpen && (
              <NamesList role="listbox" aria-label="Workbook named ranges">
                {namedRanges.map(nr => (
                  <NamesItem
                    key={nr.id}
                    type="button"
                    role="option"
                    onClick={() => {
                      onSelectNamedRange(nr)
                      setNamesOpen(false)
                    }}
                  >
                    <strong>{nr.name}</strong>
                    <span>{nr.range}</span>
                  </NamesItem>
                ))}
              </NamesList>
            )}
          </NamesMenuWrap>
        )}
      </NameBoxWrap>

      <RangeInput
        value={selectedRange}
        onFocus={() => {
          rangeBeforeFocusRef.current = selectedRange
        }}
        onChange={event => onRangeChange(event.currentTarget.value)}
        onKeyDown={event => {
          if (event.key === 'Enter') {
            event.preventDefault()
            skipNextRangeBlurRef.current = true
            const normalized = normalizeRangeToken(selectedRange)
            if (normalized) onRangeChange(normalized)
            else onRangeInvalid(`Invalid range ${selectedRange}`)
            event.currentTarget.blur()
          } else if (event.key === 'Escape') {
            event.preventDefault()
            skipNextRangeBlurRef.current = true
            onRangeChange(rangeBeforeFocusRef.current)
            event.currentTarget.blur()
          }
        }}
        onBlur={() => {
          if (skipNextRangeBlurRef.current) {
            skipNextRangeBlurRef.current = false
            return
          }
          const normalized = normalizeRangeToken(selectedRange)
          if (normalized) onRangeChange(normalized)
          else onRangeInvalid(`Invalid range ${selectedRange}`)
        }}
        aria-label="Selected range"
      />

      <FormulaStack>
        <FormulaMark aria-hidden="true">ƒx</FormulaMark>
        {!expanded && isFormula && (
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
        {expanded ? (
          <FormulaTextarea
            value={editValue}
            rows={3}
            onChange={event =>
              handleChange(
                event.currentTarget.value,
                event.currentTarget.selectionStart ?? undefined,
              )
            }
            onSelect={event =>
              setCursorPos(
                event.currentTarget.selectionStart ?? editValue.length,
              )
            }
            onFocus={() => {
              skipNextBlurCommitRef.current = false
              setFocused(true)
            }}
            onBlur={() => {
              setFocused(false)
              if (skipNextBlurCommitRef.current) {
                skipNextBlurCommitRef.current = false
                return
              }
              onCommitSelectedCell(editValue)
            }}
            onKeyDown={handleKeyDown}
            placeholder="Enter a value or formula (Shift+Enter for new line)"
            aria-label="Formula or value"
            spellCheck={false}
          />
        ) : (
          <FormulaInput
            value={editValue}
            $mirrored={isFormula}
            onChange={event =>
              handleChange(
                event.currentTarget.value,
                event.currentTarget.selectionStart ?? undefined,
              )
            }
            onSelect={event =>
              setCursorPos(
                event.currentTarget.selectionStart ?? editValue.length,
              )
            }
            onFocus={() => {
              skipNextBlurCommitRef.current = false
              setFocused(true)
            }}
            onScroll={event => {
              if (mirrorRef.current) {
                mirrorRef.current.scrollLeft = event.currentTarget.scrollLeft
              }
            }}
            onBlur={() => {
              setFocused(false)
              if (skipNextBlurCommitRef.current) {
                skipNextBlurCommitRef.current = false
                return
              }
              onCommitSelectedCell(editValue)
            }}
            onKeyDown={handleKeyDown}
            placeholder="Enter a value or formula"
            aria-label="Formula or value"
            aria-autocomplete="list"
            aria-expanded={suggestions.length > 0}
            autoComplete="off"
            spellCheck={false}
          />
        )}

        <ExpandBtn
          type="button"
          title={expanded ? 'Collapse formula bar' : 'Expand formula bar'}
          aria-label={expanded ? 'Collapse formula bar' : 'Expand formula bar'}
          onClick={() => setExpanded(e => !e)}
        >
          <PlatformIcon icon={expanded ? ChevronUp : ChevronDown} size={13} />
        </ExpandBtn>

        {suggestions.length > 0 && (
          <SuggestionList role="listbox" aria-label="Formula suggestions">
            {suggestions.map((fn, index) => (
              <SuggestionItem
                key={fn.name}
                type="button"
                role="option"
                aria-selected={fn.name === activeSuggestion?.name}
                $active={fn.name === activeSuggestion?.name}
                onMouseDown={event => {
                  event.preventDefault()
                  acceptSuggestion(fn.name)
                }}
                onMouseEnter={() => setSuggestionIndex(index)}
              >
                <SuggestionHead>
                  <SuggestionName>{fn.signature}</SuggestionName>
                  <CatBadge>{fn.category}</CatBadge>
                </SuggestionHead>
                <SuggestionDescription>{fn.description}</SuggestionDescription>
              </SuggestionItem>
            ))}
          </SuggestionList>
        )}

        {callContext && (
          <ParamHint role="status" aria-label="Function argument hint">
            <ParamSig>
              <strong>{callContext.fn.name}(</strong>
              {callContext.fn.params.map((param, idx) => (
                <span key={`${param}-${idx}`}>
                  {idx > 0 ? ', ' : ''}
                  <ParamToken $active={idx === callContext.argIndex}>
                    {param}
                  </ParamToken>
                </span>
              ))}
              <strong>)</strong>
            </ParamSig>
            <ParamDesc>{callContext.fn.description}</ParamDesc>
          </ParamHint>
        )}
      </FormulaStack>

      <PreviewWrap>
        {selectedFormulaError && errorExplanation ? (
          <>
            <ErrorPill
              type="button"
              onClick={() => setErrorPopoverOpen(o => !o)}
              title={`${errorExplanation.title} — click for details`}
            >
              {selectedDisplay}
            </ErrorPill>
            {errorPopoverOpen && (
              <ErrorPopover aria-label="Formula error explanation">
                <ErrorTitle>
                  {errorExplanation.code} · {errorExplanation.title}
                </ErrorTitle>
                <ErrorDetail>{errorExplanation.detail}</ErrorDetail>
                {errorExplanation.suggestedFix && (
                  <FixBtn
                    type="button"
                    onClick={() => {
                      const fixed = errorExplanation.suggestedFix!.formula
                      onEditValueChange(fixed)
                      onCommitSelectedCell(fixed)
                      setErrorPopoverOpen(false)
                    }}
                  >
                    <PlatformIcon icon={Sparkles} size={12} />
                    {errorExplanation.suggestedFix.label}
                  </FixBtn>
                )}
              </ErrorPopover>
            )}
          </>
        ) : selectedCheckbox && onToggleCheckboxValue ? (
          <SmartChipBtn
            type="button"
            onClick={onToggleCheckboxValue}
            title="Toggle checkbox value"
          >
            {editValue.trim().toUpperCase() === 'TRUE' ? '☑ TRUE' : '☐ FALSE'}
          </SmartChipBtn>
        ) : selectedHyperlink ? (
          <LinkChip
            href={selectedHyperlink}
            target="_blank"
            rel="noopener noreferrer"
            title={selectedHyperlink}
          >
            ↗ Open link
          </LinkChip>
        ) : (
          <FormulaPreview $error={false}>
            {selectedDisplay !== editValue ? selectedDisplay : ''}
          </FormulaPreview>
        )}
      </PreviewWrap>
    </Bar>
  )
}

const Bar = styled.div<{ $expanded?: boolean }>`
  display: grid;
  grid-template-columns: auto 104px minmax(0, 1fr) minmax(0, 160px);
  gap: 8px;
  align-items: ${({ $expanded }) => ($expanded ? 'start' : 'center')};
  padding: 8px var(--pure-chrome-inset);
  border-bottom: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-bar);

  @media (max-width: 720px) {
    grid-template-columns: auto 84px minmax(0, 1fr);
  }
`

const fieldChrome: Record<string, string> = { 'data-chrome': 'field' }

const NameBoxWrap = styled.div`
  position: relative;
  display: inline-flex;
  align-items: center;
  min-width: 52px;
`

const CellReference = styled.div.attrs(fieldChrome)`
  && {
    border-color: transparent;
    background: transparent;
    color: var(--pure-chrome-soft);
  }
  border-radius: var(--pure-chrome-radius, 3px);
  font-family: var(--sheets-mono);
  font-weight: 700;
`

const NamesMenuWrap = styled.div`
  position: relative;
  display: inline-flex;
`

const NamesTrigger = styled.button`
  display: inline-grid;
  place-items: center;
  width: 20px;
  height: 24px;
  border: 0;
  background: transparent;
  color: var(--pure-chrome-soft);
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const NamesList = styled.div`
  position: absolute;
  top: calc(100% + 4px);
  left: -48px;
  z-index: 40;
  width: 220px;
  padding: 4px;
  border: 1px solid var(--pure-chrome-line);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
  display: grid;
  gap: 2px;
`

const NamesItem = styled.button`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  padding: 5px 8px;
  border: 0;
  background: transparent;
  color: var(--sheets-ink);
  font-size: 12px;
  cursor: pointer;
  span {
    font-family: var(--sheets-mono);
    font-size: 11px;
    color: var(--pure-chrome-muted);
  }
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const FormulaMark = styled.span`
  position: absolute;
  left: 11px;
  top: 15px;
  transform: translateY(-50%);
  z-index: 2;
  color: var(--pure-chrome-muted);
  font: italic 15px Georgia, serif;
  pointer-events: none;
`

const FormulaStack = styled.div`
  position: relative;
  display: grid;
  min-width: 0;
`

const FormulaInput = styled.input.attrs(fieldChrome)<{ $mirrored?: boolean }>`
  && {
    border-radius: var(--pure-chrome-radius, 3px);
    padding-left: 36px;
    padding-right: 28px;
    background: ${({ $mirrored }) =>
      $mirrored ? 'transparent' : 'var(--pure-chrome-surface)'};
    color: ${({ $mirrored }) =>
      $mirrored ? 'transparent' : 'var(--sheets-ink)'};
    caret-color: var(--sheets-ink);
    font-family: var(--sheets-mono);
    font-size: var(--pure-chrome-ui-size);
    position: relative;
    z-index: 1;
  }

  &::selection {
    background: color-mix(in srgb, var(--pure-chrome-accent) 24%, transparent);
  }
`

const FormulaTextarea = styled.textarea`
  box-sizing: border-box;
  width: 100%;
  min-height: 68px;
  padding: 6px 28px 6px 36px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius, 3px);
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-family: var(--sheets-mono);
  font-size: var(--pure-chrome-ui-size);
  line-height: 1.45;
  resize: vertical;
  outline: none;
  &:focus {
    border-color: var(--sheets-accent);
  }
`

const ExpandBtn = styled.button`
  position: absolute;
  right: 4px;
  top: 4px;
  z-index: 3;
  display: inline-grid;
  place-items: center;
  width: 22px;
  height: 22px;
  border: 0;
  background: transparent;
  color: var(--pure-chrome-muted);
  cursor: pointer;
  &:hover {
    color: var(--sheets-ink);
    background: var(--pure-chrome-hover);
  }
`

const HighlightMirror = styled.div`
  position: absolute;
  inset: 0;
  z-index: 0;
  display: flex;
  align-items: center;
  height: var(--pure-chrome-field-height);
  padding: 0 28px 0 36px;
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
  font-weight: 400;
`

const RangeInput = styled.input.attrs(fieldChrome)`
  border-radius: var(--pure-chrome-radius, 3px);
  font-family: var(--sheets-mono);
  font-weight: 700;
`

const SuggestionList = styled.div`
  position: absolute;
  top: calc(100% + 4px);
  left: 0;
  z-index: 35;
  display: grid;
  width: min(440px, calc(100vw - 190px));
  min-width: 0;
  padding: 4px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: var(--pure-chrome-radius, 3px);
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
`

const SuggestionItem = styled.button<{ $active?: boolean }>`
  display: grid;
  gap: 1px;
  padding: 5px 8px;
  border: 0;
  border-radius: var(--pure-chrome-radius, 3px);
  background: ${({ $active }) =>
    $active ? 'var(--pure-chrome-selection)' : 'transparent'};
  text-align: left;
  cursor: pointer;

  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const SuggestionHead = styled.div`
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
`

const SuggestionName = styled.span`
  color: var(--platform-colors-text);
  font-family: var(--sheets-mono);
  font-size: var(--pure-chrome-ui-size);
  font-weight: 700;
`

const CatBadge = styled.span`
  font-size: 10px;
  font-weight: 600;
  text-transform: uppercase;
  color: var(--pure-chrome-muted);
`

const SuggestionDescription = styled.span`
  color: var(--pure-chrome-soft);
  font-size: var(--pure-chrome-ui-size);
`

const ParamHint = styled.div`
  position: absolute;
  top: calc(100% + 4px);
  left: 0;
  z-index: 34;
  display: grid;
  gap: 2px;
  max-width: min(540px, calc(100vw - 160px));
  padding: 6px 10px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
  color: var(--sheets-ink);
  font-size: 11.5px;
  line-height: 1.35;
`

const ParamSig = styled.div`
  font-family: var(--sheets-mono);
  word-break: break-word;
`

const ParamToken = styled.span<{ $active?: boolean }>`
  font-weight: ${({ $active }) => ($active ? 700 : 400)};
  color: ${({ $active }) => ($active ? 'var(--sheets-accent)' : 'inherit')};
  text-decoration: ${({ $active }) => ($active ? 'underline' : 'none')};
`

const ParamDesc = styled.div`
  font-family: var(--platform-typography-font-family);
  color: var(--pure-chrome-soft);
`

const PreviewWrap = styled.div`
  position: relative;
  min-width: 0;

  @media (max-width: 720px) {
    display: none;
  }
`

const FormulaPreview = styled(MetaText)<{ $error?: boolean }>`
  overflow: hidden;
  color: ${({ $error }) =>
    $error ? 'var(--platform-colors-danger-text)' : 'var(--pure-chrome-muted)'};
  font-weight: ${({ $error }) => ($error ? 700 : 400)};
  text-overflow: ellipsis;
  white-space: nowrap;
`

const ErrorPill = styled.button`
  padding: 2px 7px;
  border: 1px solid var(--platform-colors-danger-text);
  border-radius: 2px;
  background: color-mix(in srgb, var(--platform-colors-danger-text) 10%, transparent);
  color: var(--platform-colors-danger-text);
  font-family: var(--sheets-mono);
  font-size: 11px;
  font-weight: 700;
  cursor: pointer;
`

const ErrorPopover = styled.div`
  position: absolute;
  top: calc(100% + 6px);
  right: 0;
  z-index: 45;
  width: 260px;
  display: grid;
  gap: 6px;
  padding: 10px;
  border: 1px solid var(--platform-colors-danger-text);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  box-shadow: var(--platform-shadow-md);
`

const ErrorTitle = styled.div`
  font-size: 12px;
  font-weight: 700;
  color: var(--platform-colors-danger-text);
`

const ErrorDetail = styled.div`
  font-size: 12px;
  line-height: 1.4;
  color: var(--sheets-ink);
`

const FixBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 5px;
  height: 26px;
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-accent);
  border-radius: 2px;
  background: var(--pure-chrome-accent);
  color: var(--pure-chrome-on-accent);
  font-size: 11.5px;
  font-weight: 600;
  cursor: pointer;
`

const SmartChipBtn = styled.button`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  height: 24px;
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-ink);
  font-family: var(--sheets-mono);
  font-size: 11px;
  font-weight: 600;
  cursor: pointer;
  &:hover {
    background: var(--pure-chrome-hover);
  }
`

const LinkChip = styled.a`
  display: inline-flex;
  align-items: center;
  gap: 4px;
  height: 24px;
  padding: 0 8px;
  border: 1px solid var(--pure-chrome-line);
  border-radius: 2px;
  background: var(--pure-chrome-surface);
  color: var(--sheets-accent);
  font-size: 11.5px;
  font-weight: 600;
  text-decoration: none;
  &:hover {
    background: var(--pure-chrome-hover);
    text-decoration: underline;
  }
`
