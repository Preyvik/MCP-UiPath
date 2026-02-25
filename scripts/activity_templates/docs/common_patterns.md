# Common Activity Patterns and Best Practices

This document highlights recurring structures and implementation patterns across the 56 activity templates.

---

## Pattern 1: Simple Self-Closing Activities

Used by activities that require no nested children or body elements.

### Examples
- DeleteFileX, ReadTextFile, SetToClipboard, AddDataRow
- KillProcess, Delay, Rethrow, Return, Continue, Break

### Template Structure

```json
{
  "type": "DeleteFileX",
  "displayName": "Delete File X",
  "description": "Deletes a file at the specified path",
  "namespace": "ui",
  "requiredAttributes": ["displayName", "path"],
  "optionalAttributes": ["hintSize", "idRef"],
  "template": {
    "type": "DeleteFileX",
    "displayName": "Delete File X",
    "path": "\"C:\\\\Path\\\\To\\\\File.txt\""
  }
}
```

### XAML Output

```xml
<ui:DeleteFileX DisplayName="Delete File X" Path="&quot;C:\Path\To\File.txt&quot;" />
```

### Characteristics
- 2-4 required attributes, minimal optional attributes
- Handler inherits from base handler, implements `build_xaml()` with self-closing element
- No body, no nested activities
- Simplest handler pattern

---

## Pattern 2: Activities with XML Schema (BuildDataTable)

Used for activities that embed XML schema definitions.

### Template Structure

```json
{
  "type": "BuildDataTable",
  "displayName": "Build Data Table",
  "namespace": "ui",
  "requiredAttributes": ["displayName", "dataTable", "tableInfo"],
  "template": {
    "type": "BuildDataTable",
    "displayName": "Build Data Table",
    "dataTable": "[dtOutput]",
    "tableInfo": "<NewDataSet>...schema...</NewDataSet>"
  }
}
```

### XAML Output

```xml
<ui:BuildDataTable DisplayName="Build Data Table" DataTable="[dtOutput]"
  TableInfo="&lt;NewDataSet&gt;&#xA;  &lt;xs:schema ...&gt;...&lt;/xs:schema&gt;&#xA;&lt;/NewDataSet&gt;" />
```

### Schema Structure

The `tableInfo` attribute contains an HTML-encoded XML schema with:
- `<xs:schema>` root with XML Schema namespace
- `<xs:element>` for the DataSet and table
- `<xs:sequence>` containing column definitions
- Each column: `<xs:element name="ColumnName" type="xs:string" />`

### Column Type Mapping

| UiPath Type | XSD Type |
|-------------|----------|
| String | `xs:string` |
| Int32 | `xs:int` |
| Boolean | `xs:boolean` |
| DateTime | `xs:dateTime` |
| Double | `xs:double` |

### Characteristics
- Special handling for nested XML schema definition
- `tableInfo` is HTML-entity-encoded in XAML
- `with-schema` pattern with DataTable.Columns collection
- Handler must properly encode/decode the schema string

---

## Pattern 3: Excel Scoped Activities

Used by Excel activities that require ExcelApplicationCard scope.

### Examples
- FilterX, FindFirstLastDataRowX
- ReadRangeX, WriteCellX, WriteRangeX, SaveExcelFileX
- ClearRangeX, CopyPasteRangeX, ExecuteMacroX, InvokeVBAX

### Template Structure (FilterX)

```json
{
  "type": "FilterX",
  "displayName": "Filter Range",
  "namespace": "ueab",
  "requiredAttributes": ["displayName", "range", "columnName"],
  "optionalAttributes": ["filterArgument", "clearFilter", "hintSize", "idRef"],
  "template": {
    "type": "FilterX",
    "displayName": "Filter Range",
    "range": "[Excel.Sheet(\"Sheet1\").Range(\"A1:E100\")]",
    "columnName": "\"Status\"",
    "filterArgument": "\"Active\"",
    "clearFilter": false
  }
}
```

### Scoping Requirement

All Excel activities MUST be within ExcelApplicationCard:

```
ExcelProcessScopeX
  -> Sequence
    -> ExcelApplicationCard
      -> Sequence
        -> FilterX
        -> FindFirstLastDataRowX
        -> ReadRangeX
```

### Registration

New Excel activities must be added to the `EXCEL_SCOPED_ACTIVITIES` set in `xaml_constructor.py`:

```python
EXCEL_SCOPED_ACTIVITIES = {
    "ReadRangeX", "WriteCellX", "WriteRangeX", "SaveExcelFileX",
    "ClearRangeX", "CopyPasteRangeX", "ExecuteMacroX", "InvokeVBAX",
    "InvokeVBAArgumentX", "FilterX", "FindFirstLastDataRowX"
}
```

### Characteristics
- Standard Excel activity handler with `ueab` namespace
- Constructor automatically validates scoping during build
- Scoping violation returns structured error for auto-retry
- Must also update `skills/uipath-activity-scoping-rules/SKILL.md`

---

## Pattern 4: UI Automation with Target

Used by UI automation activities that interact with on-screen elements.

### Examples
- NMouseScroll, SearchedElement
- NClick, NTypeInto, NCheckState

### Template Structure (NMouseScroll)

```json
{
  "type": "NMouseScroll",
  "displayName": "Mouse Scroll",
  "namespace": "uix",
  "requiredAttributes": ["displayName"],
  "optionalAttributes": ["target", "direction", "amount", "movementUnits", "searchedElement", "hintSize", "idRef"],
  "template": {
    "type": "NMouseScroll",
    "displayName": "Mouse Scroll",
    "direction": "Down",
    "amount": 3,
    "movementUnits": "Lines",
    "target": {
      "Selector": "<webctrl tag='div' class='scroll-container' />",
      "Timeout": 30000
    }
  }
}
```

### Template Structure (SearchedElement)

```json
{
  "type": "SearchedElement",
  "displayName": "Searched Element",
  "namespace": "uix",
  "requiredAttributes": [],
  "optionalAttributes": ["target", "timeout", "outUiElement"],
  "template": {
    "type": "SearchedElement",
    "target": {
      "Selector": "<webctrl tag='input' name='search' />",
      "Timeout": 30000
    },
    "timeout": "00:00:30"
  }
}
```

### XAML Structure

```xml
<uix:NMouseScroll DisplayName="Mouse Scroll" Direction="Down" Amount="3">
  <uix:NMouseScroll.TargetAnchorable>
    <uix:TargetAnchorable>
      <uix:TargetAnchorable.Selector>
        <webctrl tag='div' class='scroll-container' />
      </uix:TargetAnchorable.Selector>
      <uix:TargetAnchorable.PointOffset>
        <uix:PointOffset />
      </uix:TargetAnchorable.PointOffset>
    </uix:TargetAnchorable>
  </uix:NMouseScroll.TargetAnchorable>
</uix:NMouseScroll>
```

### Characteristics
- Uses `with-target` structure leveraging TargetAnchorable
- Nested elements: TargetAnchorable contains selector and optional PointOffset
- Should be within NApplicationCard for application context
- `uix` namespace for all UI automation activities

---

## Pattern 5: Complex Control Flow with ActivityAction (InterruptibleWhile)

Used by activities that provide loop variables to child activities via ActivityAction.

### Examples
- InterruptibleWhile
- ForEach, ForEachRow (similar pattern)
- RetryScope, While (simpler body pattern)

### Template Structure

```json
{
  "type": "InterruptibleWhile",
  "displayName": "Interruptible While",
  "namespace": "ui",
  "requiredAttributes": ["displayName", "condition"],
  "optionalAttributes": ["body", "interruptCondition", "maxIterations", "currentIndex", "hintSize", "idRef"],
  "template": {
    "type": "InterruptibleWhile",
    "displayName": "Interruptible While",
    "condition": "True",
    "interruptCondition": null,
    "maxIterations": -1,
    "currentIndex": null,
    "body": {
      "variableName": "argument",
      "variableType": "Boolean",
      "activity": null
    }
  },
  "scopingRules": {
    "canContain": ["Continue", "Break"],
    "description": "Supports Continue and Break activities within body"
  }
}
```

### XAML Structure

```xml
<ui:InterruptibleWhile Condition="[condition]" MaxIterations="-1">
  <ui:InterruptibleWhile.Body>
    <ActivityAction x:TypeArguments="x:Boolean">
      <ActivityAction.Argument>
        <DelegateInArgument x:TypeArguments="x:Boolean" Name="argument" />
      </ActivityAction.Argument>
      <Sequence DisplayName="Loop Body">
        <!-- Child activities, including Continue/Break -->
      </Sequence>
    </ActivityAction>
  </ui:InterruptibleWhile.Body>
</ui:InterruptibleWhile>
```

### Characteristics
- `with-body` structure with ActivityAction pattern
- ActivityAction contains DelegateInArgument for loop variable
- Body supports Continue and Break activities within scope
- `scopingRules` field in template defines allowed child activities
- More complex than simple While (which uses direct body without ActivityAction)

---

## ActivityAction Pattern

Used by scoping activities to provide variables to child activities.

### Usage
- `ExcelProcessScopeX` (provides ExcelProcessScopeTag)
- `ExcelApplicationCard` (provides Excel workbook handle)
- `Catch` (provides exception variable)
- `ForEach` (provides iteration item)
- `InterruptibleWhile` (provides Boolean argument)
- `ExecuteMacroX` / `InvokeVBAX` (provides body scope)

### JSON Structure

```json
{
  "body": {
    "variableName": "scopeVariable",
    "variableType": "TypeName",
    "activity": null
  }
}
```

### XAML Structure

```xml
<ActivityAction x:TypeArguments="TypeName">
  <ActivityAction.Argument>
    <DelegateInArgument x:TypeArguments="TypeName" Name="variableName" />
  </ActivityAction.Argument>
  <Sequence>
    <!-- Child activities can use [variableName] -->
  </Sequence>
</ActivityAction>
```

### Catch-Specific Pattern

```json
{
  "type": "Catch",
  "x:TypeArguments": "s:Exception",
  "activityAction": {
    "x:TypeArguments": "s:Exception",
    "argument": {
      "type": "DelegateInArgument",
      "x:TypeArguments": "s:Exception",
      "name": "exception"
    },
    "handler": {
      "type": "Sequence",
      "activities": [...]
    }
  }
}
```

---

## Reference ID Pattern

Used by Flowchart for node navigation.

### Format
- Attribute: `x:Name="__ReferenceID#"`
- Reference: `<x:Reference>__ReferenceID#</x:Reference>`
- Sequential numbering: 0, 1, 2, ...

### JSON Structure

```json
{
  "type": "Flowchart",
  "displayName": "Main Flowchart",
  "startNode": { "reference": "__ReferenceID0" },
  "nodes": [
    {
      "type": "FlowStep",
      "x:Name": "__ReferenceID0",
      "activity": { "type": "Sequence", "..." },
      "next": { "reference": "__ReferenceID1" }
    },
    {
      "type": "FlowDecision",
      "x:Name": "__ReferenceID1",
      "condition": "[condition]",
      "true": { "reference": "__ReferenceID2" },
      "false": { "reference": "__ReferenceID3" }
    }
  ]
}
```

---

## Argument Collection Pattern

Used for passing data to activities with dynamic arguments.

### Usage
- `InvokeCode` (VB.NET code arguments)
- `InvokeWorkflowFile` (workflow arguments)

### JSON Structure

```json
{
  "type": "InvokeCode",
  "displayName": "Run Custom Code",
  "code": "result = input * 2",
  "arguments": [
    {
      "type": "InArgument",
      "x:TypeArguments": "x:Int32",
      "x:Key": "input",
      "value": "[inputVariable]"
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

---

## Excel Scope Pattern

Required structure for all Excel operations.

### JSON Structure

```json
{
  "type": "ExcelProcessScopeX",
  "displayName": "Use Excel Process Scope",
  "body": {
    "variableName": "ExcelProcessScopeTag",
    "variableType": "ui:IExcelProcess",
    "activity": {
      "type": "Sequence",
      "displayName": "Excel Operations",
      "activities": [
        {
          "type": "ExcelApplicationCard",
          "displayName": "Use Excel File",
          "workbook": "\"C:\\\\Data\\\\workbook.xlsx\"",
          "body": {
            "variableName": "Excel",
            "variableType": "ue:IWorkbookQuickHandle",
            "activity": {
              "type": "Sequence",
              "displayName": "File Operations",
              "activities": [
                { "type": "ReadRangeX", "..." },
                { "type": "FilterX", "..." },
                { "type": "FindFirstLastDataRowX", "..." },
                { "type": "WriteCellX", "..." },
                { "type": "SaveExcelFileX", "..." }
              ]
            }
          }
        }
      ]
    }
  }
}
```

### Key Points
- `ExcelProcessScopeX` provides `ExcelProcessScopeTag` variable
- `ExcelApplicationCard` provides `Excel` workbook handle
- All Excel activities use `[Excel.Sheet(...)]` or `[Excel.Cell(...)]` syntax
- VBA activities (ExecuteMacroX, InvokeVBAX) also require these scopes
- FilterX and FindFirstLastDataRowX follow the same scoping rules

---

## VBA Execution Pattern

For running Excel VBA macros.

### ExecuteMacroX Pattern

```json
{
  "type": "ExecuteMacroX",
  "displayName": "Run Macro",
  "macroName": "'Workbook.xlsm'!MacroName",
  "workbook": "[Excel]",
  "result": "[macroResult]",
  "body": {
    "activityAction": {
      "activity": null
    }
  }
}
```

### InvokeVBAX with Arguments Pattern

```json
{
  "type": "InvokeVBAX",
  "displayName": "Invoke Custom VBA",
  "codeFilePath": "VBA\\CustomCode.txt",
  "entryMethodName": "ProcessData",
  "workbook": "[Excel]",
  "body": {
    "activityAction": {
      "activity": {
        "type": "Sequence",
        "activities": [
          {
            "type": "InvokeVBAArgumentX",
            "displayName": "Arg1 - SheetName",
            "argumentValue": "[sheetName]"
          }
        ]
      }
    }
  }
}
```

---

## UI Automation Pattern

For application automation.

### NApplicationCard Structure

```json
{
  "type": "NApplicationCard",
  "displayName": "Use Application",
  "targetApp": {
    "type": "TargetApp",
    "selector": "<wnd app='notepad.exe' cls='Notepad' />",
    "filePath": "C:\\Windows\\notepad.exe",
    "version": "V2"
  },
  "body": {
    "type": "Sequence",
    "activities": [
      {
        "type": "NClick",
        "displayName": "Click Menu",
        "targetAnchorable": {
          "selector": "<ctrl name='File' role='menu item' />",
          "pointOffset": { "type": "PointOffset" }
        }
      },
      {
        "type": "NMouseScroll",
        "displayName": "Scroll Down",
        "direction": "Down",
        "amount": 3,
        "target": {
          "Selector": "<webctrl tag='div' class='content' />",
          "Timeout": 30000
        }
      }
    ]
  }
}
```

---

## Handler Registration Pattern

All handlers are registered in the ACTIVITY_HANDLERS dictionary in `xaml_syntaxer.py`.

### Registration

```python
ACTIVITY_HANDLERS = {
    # Key: Activity type name (matches JSON template "type" field)
    # Value: Handler class instance
    "Sequence": SequenceHandler(),
    "Assign": AssignHandler(),
    "DeleteFileX": DeleteFileXHandler(),
    "BuildDataTable": BuildDataTableHandler(),
    "FilterX": FilterXHandler(),
    "NMouseScroll": NMouseScrollHandler(),
    "InterruptibleWhile": InterruptibleWhileHandler(),
    # ... etc.
}
```

### Adding a New Handler

1. Create handler class inheriting from appropriate base
2. Implement `build_xaml()` method for write mode
3. Implement parsing logic for read mode
4. Register in ACTIVITY_HANDLERS dictionary
5. Create corresponding template JSON in `activity_templates/`

---

## Template JSON Structure

### Standard Fields

| Field | Required | Description |
|-------|----------|-------------|
| `type` | Yes | Activity type name (matches handler key) |
| `displayName` | Yes | Human-readable default name |
| `description` | Yes | What this activity does |
| `namespace` | Yes | One of: `default`, `ui`, `ueab`, `uix` |
| `requiredAttributes` | Yes | Must be present for validation to pass |
| `optionalAttributes` | Yes | May be provided but not required |
| `template` | Yes | Minimal working example with default values |
| `scopingRules` | No | Activities allowed as children (e.g., Continue/Break) |

### Namespace Values

Mapped in `scripts/namespace_mapping.json`:

| Value | XAML Prefix | Usage |
|-------|-------------|-------|
| `default` | (none) | Core .NET activities (Sequence, Assign, If) |
| `ui` | `ui:` | UiPath common activities (LogMessage, ForEach, DeleteFileX) |
| `ueab` | `ueab:` | Excel Business activities (ReadRangeX, FilterX) |
| `uix` | `uix:` | UI automation activities (NClick, NMouseScroll) |

---

## Best Practices

### 1. DisplayName Convention
- Use descriptive action verbs: "Read Customer Data", "Write Results"
- Include context: "Log - Process Started", "Assign - Initialize Counter"

### 2. Expression Formatting
- VB.NET expressions in brackets: `[variableName]`
- String literals escaped: `"\"value\""`
- XML entities in XAML: `&quot;` for quotes

### 3. Scoping Order
- Always outer scope first, then inner
- Excel: ExcelProcessScopeX -> ExcelApplicationCard -> Activities
- UI: NApplicationCard -> NClick/NTypeInto/NMouseScroll
- Loops: While/ForEach -> Continue/Break

### 4. Error Handling
- Wrap risky operations in TryCatch
- Use specific exception types in Catch when possible
- Include meaningful error messages in Throw

### 5. Variable Naming
- Prefix by type: `str_`, `int_`, `dt_`, `bool_`
- Use PascalCase or camelCase consistently
- Avoid Hungarian notation in modern workflows
