# Activity Templates

This directory contains JSON templates for all supported UiPath activities.

## Structure

Each template file contains:
- `type`: Activity type name (matches xaml_syntaxer handler)
- `displayName`: Human-readable name
- `description`: What the activity does
- `namespace`: XML namespace prefix (ui, ueab, uix, or default)
- `requiredAttributes`: List of required JSON fields
- `optionalAttributes`: List of optional JSON fields
- `template`: Minimal working example in JSON format

## Adding New Activities

1. Create new JSON file in appropriate category directory
2. Follow existing template structure
3. Ensure `type` matches the activity handler in xaml_syntaxer.py
4. Test with `python xaml_constructor.py --mode list`
5. Validate build with sample JSON

## Categories (8 categories, 57 templates)

- **core/** (9): Sequence, Assign, If, LogMessage, InvokeWorkflowFile, Flowchart, FlowStep, FlowDecision, Catch
- **control_flow/** (13): Switch, TryCatch, ForEach, ForEachRow, While, Rethrow, Delay, Throw, Return, Continue, Break, Parallel, InterruptibleWhile
- **data/** (2): BuildDataTable, AddDataRow
- **excel/** (13): ExcelProcessScopeX, ExcelApplicationCard, ReadRangeX, SaveExcelFileX, WriteCellX, WriteRangeX, CopyPasteRangeX, ClearRangeX, ExecuteMacroX, InvokeVBAX, InvokeVBAArgumentX, FilterX, FindFirstLastDataRowX
- **file_ops/** (6): CreateDirectory, MoveFile, PathExists, ReadRange, DeleteFileX, ReadTextFile
- **process/** (1): KillProcess
- **utilities/** (5): CommentOut, RetryScope, InputDialog, InvokeCode, SetToClipboard
- **ui_automation/** (8): NApplicationCard, NClick, NTypeInto, NCheckState, PointOffset, TargetApp, NMouseScroll, SearchedElement

## Template Format Example

```json
{
  "type": "Assign",
  "displayName": "Assign",
  "description": "Assigns a value to a variable or argument",
  "namespace": "default",
  "requiredAttributes": ["displayName", "to", "value"],
  "optionalAttributes": ["hintSize", "idRef"],
  "template": {
    "type": "Assign",
    "displayName": "Assign",
    "to": "variableName",
    "value": "\"value\""
  }
}
```

## Namespace Mappings

| Prefix | URI |
|--------|-----|
| default | (no namespace) |
| ui | http://schemas.uipath.com/workflow/activities |
| ueab | clr-namespace:UiPath.Excel.Activities.Business;assembly=UiPath.Excel.Activities |
| uix | http://schemas.uipath.com/workflow/activities/uix |

## Usage

List all activities:
```bash
python xaml_constructor.py --mode list
```

Get template for specific activity:
```bash
python xaml_constructor.py --mode template --type Assign
```

Build activity from JSON:
```bash
python xaml_constructor.py --mode build --input '{"type":"Assign","displayName":"Set Var","to":"x","value":"1"}'
```
