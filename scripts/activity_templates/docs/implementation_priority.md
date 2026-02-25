# Activity Template Implementation Priority Matrix

This document tracks implementation priority based on usage frequency and complexity.

## Priority Levels

- **HIGH**: Critical for workflow generation, high usage
- **MEDIUM**: Common usage, moderate complexity
- **LOW**: Rare usage or niche scenarios

---

## Implementation Status

### HIGH Priority (Completed)

| # | Activity | Usages | Files | Complexity | Status |
|---|----------|--------|-------|------------|--------|
| 1 | FlowStep | 19 | 2 | Complex | Done |
| 2 | Catch | 10 | 8 | Complex | Done |
| 3 | FlowDecision | 5 | 2 | Complex | Done |
| 4 | Flowchart | - | - | Complex | Done |

### MEDIUM Priority (Completed)

| # | Activity | Usages | Files | Complexity | Status |
|---|----------|--------|-------|------------|--------|
| 5 | Throw | 7 | 5 | Simple | Done |
| 6 | Return | 7 | 2 | Simple | Done |
| 7 | Continue | 6 | 3 | Simple | Done |
| 8 | InvokeVBAArgumentX | 6 | 3 | Simple | Done |
| 9 | InvokeCode | 7 | 4 | Moderate | Done |
| 10 | PointOffset | 8 | 1 | Simple | Done |
| 11 | InputDialog | 4 | 3 | Moderate | Done |
| 12 | ExecuteMacroX | 4 | 3 | Moderate | Done |
| 13 | TargetApp | 4 | 4 | Moderate | Done |
| 14 | Parallel | 3 | 3 | Moderate | Done |
| 15 | ClearRangeX | 3 | 2 | Simple | Done |
| 16 | InvokeVBAX | 3 | 3 | Moderate | Done |

### LOW Priority (Completed)

| # | Activity | Usages | Files | Category | Status |
|---|----------|--------|-------|----------|--------|
| 17 | SetToClipboard | 2 | 1 | Utilities | Done |
| 18 | NMouseScroll | 2 | 1 | UI Automation | Done |
| 19 | SearchedElement | 2 | 1 | UI Automation | Done |
| 20 | InterruptibleWhile | 1 | 1 | Control Flow | Done |
| 21 | FilterX | 1 | 1 | Excel | Done |
| 22 | FindFirstLastDataRowX | 1 | 1 | Excel | Done |
| 23 | DeleteFileX | 1 | 1 | File Operations | Done |
| 24 | ReadTextFile | 1 | 1 | File Operations | Done |
| 25 | BuildDataTable | 1 | 1 | Data Manipulation | Done |
| 26 | AddDataRow | 1 | 1 | Data Manipulation | Done |

*All LOW priority templates implemented in Phase 4-7.*

### Excluded (Custom/Third-Party - 3 Activities)

| Activity | Namespace | Reason |
|----------|-----------|--------|
| Report_and_Log | sl | Third-party (Swift Flow Tools) |
| Log_in_as_Robot | ss | Third-party (Swift Flow Tools) |
| Open_transaction | ss | Third-party (Swift Flow Tools) |

---

## Complexity Definitions

### Simple
- No nested structures
- Self-closing element
- 1-2 required attributes
- No scoping requirements

### Moderate
- Contains body/child elements
- ActivityAction pattern (optional)
- 2-4 required attributes
- Basic scoping requirements

### Complex
- Reference ID management
- Multiple nested levels
- ActivityAction with DelegateInArgument
- Strict scoping requirements

---

## Implementation Notes

### FlowStep/FlowDecision
- Requires unique `x:Name` generation (__ReferenceID#)
- References between nodes via `<x:Reference>`
- ViewState for visual positioning

### Catch
- TypeArguments must match between Catch and ActivityAction
- DelegateInArgument provides exception variable name
- Handler is typically a Sequence

### Excel VBA Activities
- Must be within ExcelApplicationCard scope
- Workbook variable references parent scope
- InvokeVBAArgumentX passes parameters in order

### UI Automation
- TargetApp nested in NApplicationCard.TargetApp
- PointOffset often empty element
- Selector identifies target application

---

## Gap Analysis Reference

See `scripts/activity_gap_analysis.md` for detailed gap analysis.
See `scripts/activity_templates/LOW_PRIORITY_TEMPLATES.md` for remaining template specifications.
See `scripts/TEMPLATE_IMPLEMENTATION_ROADMAP.md` for implementation timeline.

### Coverage Summary

| Phase | Templated | Coverage |
|-------|-----------|----------|
| Baseline (Before Phase 3) | 30/57 | 53% |
| After Phase 3 | 46/57 | 81% |
| After Phase 4-7 (Final) | 57/57 | **100%** |

*Note: 57 eligible activities = 59 total discovered - 3 excluded custom + 1 new data category activity*
*All 26 HIGH, MEDIUM, and LOW priority activities are now implemented.*
