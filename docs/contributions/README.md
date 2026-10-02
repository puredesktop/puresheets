# puresheets contribution roadmap

[View roadmap issues](https://github.com/puredesktop/puresheets/issues?q=is%3Aissue%20label%3Aroadmap)

Build something you can see and try in the app. The first five items are **good first contributions**: bounded changes with a concrete demonstration. Choose a feature below, fix a bug, or propose your own improvement.

## Scope

Keep Univer-backed workbook editing, existing cells and formulas, worksheet tabs and supported file formats.

Size describes scope, not a promised completion time: **Small** = one focused interface change; **Medium** = coordinated interface/state work; **Large** = a feature across several flows, storage or export paths. All items are proposals, not claims that existing features are absent. Check the current code and extend what is there. Maintainers review code and tests before merging. Attribution is your choice.

## Good first contributions

1. **[Expand the formula editor.](https://github.com/puredesktop/puresheets/issues/3)** Allow the existing formula input to expand for long expressions while preserving its commit and cancel behavior.
   <!-- contribution: {"id": "long-formula-editing", "size": "small", "goodFirstIssue": true, "guide": "docs/contributions/long-formula-editing.md"} -->
   [Small · Good first contribution · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/long-formula-editing.md)

2. **[See the size of a selected range.](https://github.com/puredesktop/puresheets/issues/4)** Display the selected range's row and column counts next to its address so large operations have a visible scope.
   <!-- contribution: {"id": "range-selection-clarity", "size": "small", "goodFirstIssue": true, "guide": "docs/contributions/range-selection-clarity.md"} -->
   [Small · Good first contribution · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/range-selection-clarity.md)

3. **[See where a comment belongs.](https://github.com/puredesktop/puresheets/issues/5)** Show the sheet and cell reference beside entries in the existing comments list so similarly worded comments can be distinguished.
   <!-- contribution: {"id": "comment-location-context", "size": "small", "goodFirstIssue": true, "guide": "docs/contributions/comment-location-context.md"} -->
   [Small · Good first contribution · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/comment-location-context.md)

4. **[Know which sheet a CSV export includes.](https://github.com/puredesktop/puresheets/issues/6)** Make the active sheet and output scope explicit when exporting CSV so users do not assume it includes the whole workbook.
   <!-- contribution: {"id": "csv-export-scope-wording", "size": "small", "goodFirstIssue": true, "guide": "docs/contributions/csv-export-scope-wording.md"} -->
   [Small · Good first contribution · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/csv-export-scope-wording.md)

5. **[See the cells behind a chart.](https://github.com/puredesktop/puresheets/issues/7)** Display the source sheet and range prominently in the existing chart panel to make stale or incorrect selections easier to spot.
   <!-- contribution: {"id": "chart-source-range-visibility", "size": "small", "goodFirstIssue": true, "guide": "docs/contributions/chart-source-range-visibility.md"} -->
   [Small · Good first contribution · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/chart-source-range-visibility.md)

## More improvements

6. **[Choose formulas with examples.](https://github.com/puredesktop/puresheets/issues/8)** Add a short example and argument hint to each existing formula suggestion so users can choose the right function without leaving the formula bar.
   <!-- contribution: {"id": "formula-suggestion-examples", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/formula-suggestion-examples.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/formula-suggestion-examples.md)

7. **[Understand a formula error.](https://github.com/puredesktop/puresheets/issues/9)** Show a concise explanation for common displayed formula errors alongside the cell reference, without replacing the stored formula.
   <!-- contribution: {"id": "formula-error-explanation", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/formula-error-explanation.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/formula-error-explanation.md)

8. **[Correct a sheet name inline.](https://github.com/puredesktop/puresheets/issues/10)** Explain blank, duplicate or unsupported sheet names in the rename control and retain the attempted name for correction.
   <!-- contribution: {"id": "sheet-rename-validation", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/sheet-rename-validation.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/sheet-rename-validation.md)

9. **[Check which sheet you are deleting.](https://github.com/puredesktop/puresheets/issues/11)** Include the sheet name and whether it contains data in the existing delete confirmation.
   <!-- contribution: {"id": "sheet-delete-context", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/sheet-delete-context.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/sheet-delete-context.md)

10. **[Find a newly duplicated sheet.](https://github.com/puredesktop/puresheets/issues/12)** Choose a predictable unused copy name and focus the new tab after the existing duplicate-sheet action.
   <!-- contribution: {"id": "duplicate-sheet-naming", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/duplicate-sheet-naming.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/duplicate-sheet-naming.md)

11. **[Keep the active sheet tab visible.](https://github.com/puredesktop/puresheets/issues/13)** Keep the active sheet tab visible when many worksheets overflow the available width, including after rename or duplicate.
   <!-- contribution: {"id": "overflowing-tab-navigation", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/overflowing-tab-navigation.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/overflowing-tab-navigation.md)

12. **[Understand what clearing a comment does.](https://github.com/puredesktop/puresheets/issues/14)** Treat a blank comment consistently in the editor and list, explaining when an existing comment will be removed.
   <!-- contribution: {"id": "comment-empty-text-handling", "size": "small", "goodFirstIssue": false, "guide": "docs/contributions/comment-empty-text-handling.md"} -->
   [Small · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/comment-empty-text-handling.md)

13. **[Correct data-rule bounds.](https://github.com/puredesktop/puresheets/issues/15)** Explain when a data rule's minimum exceeds its maximum beside the rule fields before applying the rule.
   <!-- contribution: {"id": "validation-bounds-feedback", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/validation-bounds-feedback.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/validation-bounds-feedback.md)

14. **[Preview dropdown choices before saving.](https://github.com/puredesktop/puresheets/issues/16)** Show a preview of dropdown choices after trimming accidental whitespace and identify repeated choices before saving a rule.
   <!-- contribution: {"id": "validation-list-cleanup", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/validation-list-cleanup.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/validation-list-cleanup.md)

15. **[Check which cells a rule affects.](https://github.com/puredesktop/puresheets/issues/17)** Repeat the target sheet and cell range in the data-rule panel so users know which cells will be affected.
   <!-- contribution: {"id": "validation-range-recap", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/validation-range-recap.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/validation-range-recap.md)

16. **[Review imported sheets and compatibility notices.](https://github.com/puredesktop/puresheets/issues/18)** After workbook import, show the imported sheet names and any compatibility notices already reported by the importer.
   <!-- contribution: {"id": "import-sheet-summary", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/import-sheet-summary.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/import-sheet-summary.md)

17. **[Choose the right workbook export format.](https://github.com/puredesktop/puresheets/issues/19)** Explain the distinction between editable .sheets, portable .sheets.html and supported interchange exports beside the existing choices.
   <!-- contribution: {"id": "export-format-guidance", "size": "small", "goodFirstIssue": false, "guide": "docs/contributions/export-format-guidance.md"} -->
   [Small · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/export-format-guidance.md)

18. **[Preview number formats.](https://github.com/puredesktop/puresheets/issues/20)** Show a representative value under each existing number-format choice, using the workbook's current locale settings.
   <!-- contribution: {"id": "number-format-previews", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/number-format-previews.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/number-format-previews.md)

19. **[See which ranges an agent changed.](https://github.com/puredesktop/puresheets/issues/21)** Include affected sheet names and ranges beside existing agent log summaries without expanding every before-and-after cell value.
   <!-- contribution: {"id": "agent-activity-range-summaries", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/agent-activity-range-summaries.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/agent-activity-range-summaries.md)

20. **[Understand the compatibility grid’s limits.](https://github.com/puredesktop/puresheets/issues/22)** When the compatibility grid is active, explain which editing controls are unavailable and how that differs from the full spreadsheet surface.
   <!-- contribution: {"id": "compatibility-mode-explanation", "size": "medium", "goodFirstIssue": false, "guide": "docs/contributions/compatibility-mode-explanation.md"} -->
   [Medium · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/compatibility-mode-explanation.md)

21. **[Preview CSV imports before changing cells.](https://github.com/puredesktop/puresheets/issues/23)** Add a CSV import preview with delimiter and header choices plus the destination range. Show the cells that will be overwritten and require Apply; Cancel leaves the workbook unchanged.
   <!-- contribution: {"id": "preview-csv-imports-before-changing-cells", "size": "large", "goodFirstIssue": false, "guide": "docs/contributions/preview-csv-imports-before-changing-cells.md"} -->
   [Large · Implementation brief](https://github.com/puredesktop/puresheets/blob/main/docs/contributions/preview-csv-imports-before-changing-cells.md)

## References

- [App guide](https://github.com/puredesktop/puresheets/blob/main/docs/app-guide.md)
- [Development guide](https://github.com/puredesktop/puresheets/blob/main/docs/development.md)
- [Contributing](https://github.com/puredesktop/puresheets/blob/main/CONTRIBUTING.md)
