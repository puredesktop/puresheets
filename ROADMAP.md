# puresheets roadmap

## Scope

Keep Univer-backed workbook editing, existing cells and formulas, worksheet tabs and supported file formats.

These are proposed, incremental improvements, not a release schedule or a list of missing core features. Keep each change small and preserve existing file formats, user data and app workflows.

## Improvements

1. **Formula suggestion examples.** Add a short example and argument hint to each existing formula suggestion so users can choose the right function without leaving the formula bar.

2. **Formula error explanation.** Show a concise explanation for common displayed formula errors alongside the cell reference, without replacing the stored formula.

3. **Long formula editing.** Allow the existing formula input to expand for long expressions while preserving its commit and cancel behavior.

4. **Range selection clarity.** Display the selected range's row and column counts next to its address so large operations have a visible scope.

5. **Sheet rename validation.** Explain blank, duplicate or unsupported sheet names in the rename control and retain the attempted name for correction.

6. **Sheet delete context.** Include the sheet name and whether it contains data in the existing delete confirmation.

7. **Duplicate sheet naming.** Choose a predictable unused copy name and focus the new tab after the existing duplicate-sheet action.

8. **Overflowing tab navigation.** Keep the active sheet tab visible when many worksheets overflow the available width, including after rename or duplicate.

9. **Comment location context.** Show the sheet and cell reference beside entries in the existing comments list so similarly worded comments can be distinguished.

10. **Comment empty-text handling.** Treat a blank comment consistently in the editor and list, explaining when an existing comment will be removed.

11. **Validation bounds feedback.** Explain when a data rule's minimum exceeds its maximum beside the rule fields before applying the rule.

12. **Validation list cleanup.** Show a preview of dropdown choices after trimming accidental whitespace and identify repeated choices before saving a rule.

13. **Validation range recap.** Repeat the target sheet and cell range in the data-rule panel so users know which cells will be affected.

14. **Import sheet summary.** After workbook import, show the imported sheet names and any compatibility notices already reported by the importer.

15. **Export format guidance.** Explain the distinction between editable .sheets, portable .sheets.html and supported interchange exports beside the existing choices.

16. **CSV export scope wording.** Make the active sheet and output scope explicit when exporting CSV so users do not assume it includes the whole workbook.

17. **Number format previews.** Show a representative value under each existing number-format choice, using the workbook's current locale settings.

18. **Chart source range visibility.** Display the source sheet and range prominently in the existing chart panel to make stale or incorrect selections easier to spot.

19. **Agent activity range summaries.** Include affected sheet names and ranges beside existing agent log summaries without expanding every before-and-after cell value.

20. **Compatibility mode explanation.** When the compatibility grid is active, explain which editing controls are unavailable and how that differs from the full spreadsheet surface.

## References

- [App guide](docs/app-guide.md)
- [Development guide](docs/development.md)
- [Current implementation](src/components/PureSheetsShell.tsx)
