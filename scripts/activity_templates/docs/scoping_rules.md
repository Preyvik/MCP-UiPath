# Activity Scoping Rules

This document consolidates all scoping requirements for UiPath activities.

## Excel Activity Scoping Hierarchy

All Excel activities **MUST** be placed within the correct scope containers:

```
ExcelProcessScopeX (outermost - manages Excel process lifecycle)
  └─ ExcelApplicationCard (manages workbook)
      └─ Excel Activities (ReadRangeX, WriteCellX, etc.)
          └─ InvokeVBAX (VBA code execution)
              └─ InvokeVBAArgumentX (in Body - passes arguments)
```

### Affected Excel Activities

| Activity | Scope Required |
|----------|----------------|
| ReadRangeX | ExcelApplicationCard |
| WriteCellX | ExcelApplicationCard |
| WriteRangeX | ExcelApplicationCard |
| ClearRangeX | ExcelApplicationCard |
| SaveExcelFileX | ExcelApplicationCard |
| CopyPasteRangeX | ExcelApplicationCard |
| ExecuteMacroX | ExcelApplicationCard |
| InvokeVBAX | ExcelApplicationCard |
| InvokeVBAArgumentX | InvokeVBAX.Body |

### Valid Excel Structure Pattern

```json
{
  "type": "ExcelProcessScopeX",
  "displayName": "Use Excel Process Scope",
  "body": {
    "variableName": "ExcelProcessScopeTag",
    "variableType": "ui:IExcelProcess",
    "activity": {
      "type": "Sequence",
      "activities": [
        {
          "type": "ExcelApplicationCard",
          "displayName": "Use Excel File",
          "workbook": "\"C:\\path\\to\\file.xlsx\"",
          "body": {
            "variableName": "Excel",
            "variableType": "ue:IWorkbookQuickHandle",
            "activity": {
              "type": "Sequence",
              "activities": [
                { "type": "ReadRangeX", "..." },
                { "type": "WriteCellX", "..." }
              ]
            }
          }
        }
      ]
    }
  }
}
```

---

## Flowchart Activity Scoping

Flowchart activities use reference IDs for navigation between nodes.

```
Flowchart
  ├─ StartNode → <x:Reference>__ReferenceID#</x:Reference>
  ├─ FlowStep (x:Name="__ReferenceID#")
  │   ├─ Activity (child - Sequence, Assign, etc.)
  │   └─ Next → <x:Reference>__ReferenceID#</x:Reference>
  └─ FlowDecision (x:Name="__ReferenceID#")
      ├─ Condition (boolean expression)
      ├─ True → <x:Reference>__ReferenceID#</x:Reference>
      └─ False → <x:Reference>__ReferenceID#</x:Reference>
```

### Reference ID Rules

- Format: `x:Name="__ReferenceID#"` where # is sequential (0, 1, 2, ...)
- Must be unique within the Flowchart
- Linking uses: `<x:Reference>__ReferenceID#</x:Reference>`
- StartNode references first FlowStep

---

## Loop Activity Scoping

Loop control activities must be within their parent loop.

| Activity | Must Be Within |
|----------|----------------|
| Continue | ForEach, ForEachRow, While, InterruptibleWhile |
| Break | ForEach, ForEachRow, While, InterruptibleWhile |

### Valid Loop Structure

```json
{
  "type": "ForEach",
  "displayName": "For Each Item",
  "body": {
    "variableName": "item",
    "variableType": "x:Object",
    "activity": {
      "type": "Sequence",
      "activities": [
        {
          "type": "If",
          "condition": "[skipCondition]",
          "then": { "type": "Continue" }
        }
      ]
    }
  }
}
```

---

## Error Handling Scoping

Exception handling activities have specific placement requirements.

| Activity | Must Be Within |
|----------|----------------|
| Catch | TryCatch.Catches collection |
| Rethrow | Catch handler |
| Throw | Anywhere (no restriction) |

### Valid TryCatch Structure

```json
{
  "type": "TryCatch",
  "displayName": "Try Catch",
  "try": { "type": "Sequence", "activities": [...] },
  "catches": [
    {
      "type": "Catch",
      "x:TypeArguments": "s:Exception",
      "activityAction": {
        "argument": { "name": "exception" },
        "handler": { "type": "Sequence", "activities": [...] }
      }
    }
  ],
  "finally": null
}
```

---

## UI Automation Scoping

UI automation activities have nested targeting structures.

| Activity | Scope/Container |
|----------|-----------------|
| TargetApp | NApplicationCard.TargetApp property |
| PointOffset | TargetAnchorable.PointOffset property |
| NClick, NTypeInto, etc. | NApplicationCard body |

### Valid UI Automation Structure

```json
{
  "type": "NApplicationCard",
  "displayName": "Use Application/Browser",
  "targetApp": {
    "type": "TargetApp",
    "selector": "<wnd app='notepad.exe' />"
  },
  "body": {
    "type": "Sequence",
    "activities": [
      {
        "type": "NClick",
        "displayName": "Click",
        "targetAnchorable": {
          "selector": "...",
          "pointOffset": { "type": "PointOffset" }
        }
      }
    ]
  }
}
```

---

## Common Scoping Violations

| Violation | Error Message | Fix |
|-----------|---------------|-----|
| Excel activity outside scope | "scoping_violation" | Wrap in ExcelApplicationCard within ExcelProcessScopeX |
| Continue outside loop | Runtime error | Place inside ForEach/While body |
| Catch outside TryCatch | Parse error | Place in TryCatch.Catches array |
| Rethrow outside Catch | Compile error | Place inside Catch handler |
| FlowStep outside Flowchart | Parse error | Place inside Flowchart container |

---

## Scoping Validation

The Constructor agent validates scoping rules automatically:
1. Identifies activity types requiring scope containers
2. Checks ancestor hierarchy for required containers
3. Returns structured error for violations
4. Provides `retry_suggestion` with correct structure

See: `uipath-activity-scoping-rules` skill for implementation details.

---

## Parallel Activity Scoping

Parallel activities execute multiple branches concurrently.

### Structure

```
Parallel
  ├─ CompletionCondition (optional boolean)
  └─ Branches (collection of activities)
      ├─ Branch 1 (Sequence or single activity)
      ├─ Branch 2 (Sequence or single activity)
      └─ Branch N (...)
```

### Variable Scope Isolation

| Rule | Description |
|------|-------------|
| Shared read | All branches can READ parent-scope variables |
| No cross-write | Branches CANNOT write to variables another branch reads |
| Local preferred | Use local variables within each branch for write operations |

### Valid Parallel Structure

```json
{
  "type": "Parallel",
  "displayName": "Parallel Processing",
  "completionCondition": null,
  "branches": [
    {
      "type": "Sequence",
      "displayName": "Branch 1",
      "activities": [
        { "type": "LogMessage", "message": "\"Branch 1 executing\"" }
      ]
    },
    {
      "type": "Sequence",
      "displayName": "Branch 2",
      "activities": [
        { "type": "LogMessage", "message": "\"Branch 2 executing\"" }
      ]
    }
  ]
}
```

### Completion Condition

- When `null` or not set: Waits for ALL branches to complete
- When set to expression: Stops when condition evaluates to `True`
- Example: `[completedCount >= 3]` stops after 3 branches finish

---

## InvokeCode Scoping

InvokeCode executes custom VB.NET or C# code within the workflow.

### Argument Visibility

| Argument Type | Code Access | Direction |
|---------------|-------------|-----------|
| InArgument | Read-only | Into code |
| OutArgument | Write-only | Out of code |
| InOutArgument | Read/Write | Both |

### Namespace Imports

Code can reference:
- Standard .NET namespaces (System, System.Collections.Generic)
- UiPath namespaces (if assembly referenced)
- Custom namespaces via explicit Import property

### Variable Access Pattern

```json
{
  "type": "InvokeCode",
  "displayName": "Calculate Result",
  "language": "VisualBasic",
  "code": "result = input * multiplier",
  "arguments": [
    {
      "type": "InArgument",
      "x:TypeArguments": "x:Int32",
      "x:Key": "input",
      "value": "[inputVariable]"
    },
    {
      "type": "InArgument",
      "x:TypeArguments": "x:Int32",
      "x:Key": "multiplier",
      "value": "[2]"
    },
    {
      "type": "OutArgument",
      "x:TypeArguments": "x:Int32",
      "x:Key": "result",
      "value": "[outputVariable]"
    }
  ]
}
```

### Scoping Rules

| Rule | Description |
|------|-------------|
| No direct variable access | Must use Arguments collection |
| Argument names in code | Match x:Key values exactly |
| Type safety | x:TypeArguments must match variable types |

---

## InputDialog Scoping

InputDialog prompts users for input with optional choices.

### Result Variable Type

| Property | Type | Notes |
|----------|------|-------|
| Result | x:Object | Always Object type, cast as needed |
| SelectedIndex | Int32 | Index of selected option (if using options) |

### Options Handling

Two mutually exclusive approaches:

| Property | Type | Use Case |
|----------|------|----------|
| Options | String[] | Array of option strings |
| OptionsString | String | Newline-separated options |

### Valid InputDialog Structure

```json
{
  "type": "InputDialog",
  "displayName": "Get User Choice",
  "title": "\"Select Option\"",
  "label": "\"Please choose an option:\"",
  "options": "[{\"Option A\", \"Option B\", \"Option C\"}]",
  "result": "[userChoice]",
  "selectedIndex": "[selectedIdx]"
}
```

### Scoping Rules

| Rule | Description |
|------|-------------|
| No container required | Can be placed anywhere in workflow |
| Result is Object | Cast to String: `CStr(userChoice)` |
| Options array format | VB.NET array literal: `{\"A\", \"B\", \"C\"}` |

---

## VBA Argument Passing

InvokeVBAArgumentX passes arguments to VBA code in InvokeVBAX.

### Ordering Requirements

Arguments are passed POSITIONALLY to VBA:

```
InvokeVBAX
  └─ Body
      └─ Sequence
          ├─ InvokeVBAArgumentX (Arg1 - first parameter)
          ├─ InvokeVBAArgumentX (Arg2 - second parameter)
          └─ InvokeVBAArgumentX (Arg3 - third parameter)
```

### Parameter Type Conversions

| VB.NET Type | VBA Type | Notes |
|-------------|----------|-------|
| String | String | Direct mapping |
| Int32 | Long | VBA Long is 32-bit |
| Double | Double | Direct mapping |
| Boolean | Boolean | Direct mapping |
| Object | Variant | For dynamic types |

### Valid VBA Invocation Structure

```json
{
  "type": "InvokeVBAX",
  "displayName": "Run VBA Function",
  "codeFilePath": "VBA\\MyFunction.txt",
  "entryMethodName": "ProcessData",
  "workbook": "[Excel]",
  "body": {
    "activityAction": {
      "activity": {
        "type": "Sequence",
        "displayName": "VBA Arguments",
        "activities": [
          {
            "type": "InvokeVBAArgumentX",
            "displayName": "Arg1 - SheetName",
            "argumentValue": "[sheetName]"
          },
          {
            "type": "InvokeVBAArgumentX",
            "displayName": "Arg2 - StartRow",
            "argumentValue": "[startRow]"
          },
          {
            "type": "InvokeVBAArgumentX",
            "displayName": "Arg3 - EndRow",
            "argumentValue": "[endRow]"
          }
        ]
      }
    }
  }
}
```

### Scoping Rules

| Rule | Description |
|------|-------------|
| Must be in InvokeVBAX.Body | Cannot be standalone |
| Order matters | Arguments passed positionally to VBA |
| Parent must be in ExcelApplicationCard | InvokeVBAX requires Excel scope |

---

## Validation Rules Summary

| Activity | Required Container | Error if Violated |
|----------|-------------------|-------------------|
| ReadRangeX, WriteCellX, etc. | ExcelApplicationCard | scoping_violation |
| ExcelApplicationCard | ExcelProcessScopeX | scoping_violation |
| InvokeVBAArgumentX | InvokeVBAX.Body | parse_error |
| InvokeVBAX | ExcelApplicationCard | scoping_violation |
| FlowStep, FlowDecision | Flowchart | parse_error |
| Catch | TryCatch.Catches | parse_error |
| Rethrow | Catch handler | compile_error |
| Continue, Break | Loop body (ForEach/While) | runtime_error |
| PointOffset | TargetAnchorable | parse_error |
| TargetApp | NApplicationCard | attribute_error |

### Validation Error Response Format

```json
{
  "status": "error",
  "error_type": "scoping_violation",
  "rule": "Excel activities must be within ExcelApplicationCard scope",
  "invalid_activities": [
    {
      "type": "ReadRangeX",
      "displayName": "Read Data",
      "current_parent": "Sequence"
    }
  ],
  "fix": "Wrap Excel activities in ExcelApplicationCard container",
  "retry_suggestion": "ExcelProcessScopeX -> ExcelApplicationCard -> Sequence -> [activities]"
}
```

---

## Cross-Reference

- [Common Patterns](./common_patterns.md) - Pattern examples with code
- [Implementation Priority](./implementation_priority.md) - Template status tracking
- [Quick Reference](../TEMPLATE_QUICK_REFERENCE.md) - Fast lookup guide
