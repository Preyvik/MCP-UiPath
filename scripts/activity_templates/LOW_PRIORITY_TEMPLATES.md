# Low Priority Template Specifications

This document provides specifications for the 10 remaining low-priority activities identified in the gap analysis. These templates can be implemented as needed to achieve 100% coverage of eligible activities.

---

## Overview

| Activity | Namespace | Usages | Complexity | Est. Hours |
|----------|-----------|--------|------------|------------|
| SetToClipboard | ui | 2 | Simple | 1-2 |
| NMouseScroll | uix | 2 | Moderate | 3-4 |
| SearchedElement | uix | 2 | Moderate | 3-4 |
| InterruptibleWhile | ui | 1 | Complex | 6-8 |
| FilterX | ueab | 1 | Moderate | 3-4 |
| FindFirstLastDataRowX | ueab | 1 | Moderate | 3-4 |
| DeleteFileX | ui | 1 | Simple | 1-2 |
| ReadTextFile | ui | 1 | Simple | 1-2 |
| BuildDataTable | ui | 1 | Moderate | 3-4 |
| AddDataRow | ui | 1 | Simple | 1-2 |

**Total Estimated Effort:** ~31 hours

*Note: 56 eligible activities = 59 total - 3 excluded custom (Report_and_Log, Log_in_as_Robot, Open_transaction)*

---

## Tier 1: Quick Wins (Simple, 1-2 hours each)

### 1. SetToClipboard

**Namespace:** `http://schemas.uipath.com/workflow/activities` (ui:)
**Usages:** 2 | **Files:** 1
**Complexity:** Simple

**Description:** Copies text or data to the system clipboard.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| Text | InArgument<string> | The text to copy to clipboard |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Set To Clipboard" | Activity display name |

**Template JSON:**
```json
{
  "type": "SetToClipboard",
  "namespace": "ui",
  "category": "utilities",
  "complexity": "simple",
  "requiredAttributes": ["text"],
  "optionalAttributes": ["displayName"],
  "template": {
    "type": "SetToClipboard",
    "displayName": "Set To Clipboard",
    "text": null
  },
  "xamlStructure": "self-closing"
}
```

**XAML Example:**
```xml
<ui:SetToClipboard DisplayName="Copy Result" Text="[resultText]" />
```

**Scoping:** None required - can be used anywhere.

---

### 2. DeleteFileX

**Namespace:** `http://schemas.uipath.com/workflow/activities` (ui:)
**Usages:** 1 | **Files:** 1
**Complexity:** Simple

**Description:** Deletes a file from the file system.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| Path | InArgument<string> | Full path to the file to delete |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Delete File" | Activity display name |
| ContinueOnError | bool | false | Continue if file not found |

**Template JSON:**
```json
{
  "type": "DeleteFileX",
  "namespace": "ui",
  "category": "file_ops",
  "complexity": "simple",
  "requiredAttributes": ["path"],
  "optionalAttributes": ["displayName", "continueOnError"],
  "template": {
    "type": "DeleteFileX",
    "displayName": "Delete File",
    "path": null,
    "continueOnError": false
  },
  "xamlStructure": "self-closing"
}
```

**XAML Example:**
```xml
<ui:DeleteFileX DisplayName="Delete Temp File" Path="[tempFilePath]" ContinueOnError="True" />
```

**Scoping:** None required - can be used anywhere.

---

### 3. ReadTextFile

**Namespace:** `http://schemas.uipath.com/workflow/activities` (ui:)
**Usages:** 1 | **Files:** 1
**Complexity:** Simple

**Description:** Reads all text content from a file.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| FileName | InArgument<string> | Full path to the file to read |
| Content | OutArgument<string> | Variable to store file contents |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Read Text File" | Activity display name |
| Encoding | string | "UTF-8" | Text encoding |

**Template JSON:**
```json
{
  "type": "ReadTextFile",
  "namespace": "ui",
  "category": "file_ops",
  "complexity": "simple",
  "requiredAttributes": ["fileName", "content"],
  "optionalAttributes": ["displayName", "encoding"],
  "template": {
    "type": "ReadTextFile",
    "displayName": "Read Text File",
    "fileName": null,
    "content": null,
    "encoding": "\"UTF-8\""
  },
  "xamlStructure": "self-closing"
}
```

**XAML Example:**
```xml
<ui:ReadTextFile DisplayName="Read Config" FileName="[configPath]" Content="[fileContent]" Encoding="UTF-8" />
```

**Scoping:** None required - can be used anywhere.

---

### 4. AddDataRow

**Namespace:** `http://schemas.uipath.com/workflow/activities` (ui:)
**Usages:** 1 | **Files:** 1
**Complexity:** Simple

**Description:** Adds a new row to a DataTable.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| DataTable | InArgument<DataTable> | The DataTable to add row to |
| ArrayRow | InArgument<Object[]> | Array of values for the new row |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Add Data Row" | Activity display name |
| DataRow | InArgument<DataRow> | null | Existing DataRow to add |

**Template JSON:**
```json
{
  "type": "AddDataRow",
  "namespace": "ui",
  "category": "data",
  "complexity": "simple",
  "requiredAttributes": ["dataTable", "arrayRow"],
  "optionalAttributes": ["displayName", "dataRow"],
  "template": {
    "type": "AddDataRow",
    "displayName": "Add Data Row",
    "dataTable": null,
    "arrayRow": null
  },
  "xamlStructure": "self-closing"
}
```

**XAML Example:**
```xml
<ui:AddDataRow DisplayName="Add Customer Row" DataTable="[dt_Customers]" ArrayRow="[{customerName, customerEmail, customerPhone}]" />
```

**Scoping:** None required - can be used anywhere.

---

## Tier 2: Moderate Effort (3-5 hours each)

### 5. NMouseScroll

**Namespace:** `http://schemas.uipath.com/workflow/activities/uix` (uix:)
**Usages:** 2 | **Files:** 1
**Complexity:** Moderate

**Description:** Scrolls within a UI element using mouse wheel.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| Target | Target | UI element to scroll within |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Mouse Scroll" | Activity display name |
| Direction | ScrollDirection | Down | Scroll direction (Up/Down/Left/Right) |
| Amount | int | 3 | Number of scroll units |
| MovementUnits | MovementUnits | Lines | Unit type (Lines/Pages) |

**Template JSON:**
```json
{
  "type": "NMouseScroll",
  "namespace": "uix",
  "category": "ui_automation",
  "complexity": "moderate",
  "requiredAttributes": ["target"],
  "optionalAttributes": ["displayName", "direction", "amount", "movementUnits"],
  "template": {
    "type": "NMouseScroll",
    "displayName": "Mouse Scroll",
    "direction": "Down",
    "amount": 3,
    "movementUnits": "Lines",
    "targetAnchorable": {
      "selector": null,
      "pointOffset": { "type": "PointOffset" }
    }
  },
  "xamlStructure": "with-target"
}
```

**XAML Example:**
```xml
<uix:NMouseScroll DisplayName="Scroll Down" Direction="Down" Amount="5" MovementUnits="Lines">
  <uix:NMouseScroll.Target>
    <uix:TargetAnchorable>
      <uix:TargetAnchorable.Selector>&lt;webctrl tag='div' class='scroll-container' /&gt;</uix:TargetAnchorable.Selector>
    </uix:TargetAnchorable>
  </uix:NMouseScroll.Target>
</uix:NMouseScroll>
```

**Scoping:** Should be within NApplicationCard body.

---

### 6. SearchedElement

**Namespace:** `http://schemas.uipath.com/workflow/activities/uix` (uix:)
**Usages:** 2 | **Files:** 1
**Complexity:** Moderate

**Description:** Configures element search parameters for UI automation.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| Selector | string | UI element selector |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| SearchDepth | int | 10 | Max search depth |
| Timeout | TimeSpan | 30s | Search timeout |
| WaitForReady | WaitType | Interactive | Wait strategy |

**Template JSON:**
```json
{
  "type": "SearchedElement",
  "namespace": "uix",
  "category": "ui_automation",
  "complexity": "moderate",
  "requiredAttributes": ["selector"],
  "optionalAttributes": ["searchDepth", "timeout", "waitForReady"],
  "template": {
    "type": "SearchedElement",
    "selector": null,
    "searchDepth": 10,
    "timeout": "00:00:30",
    "waitForReady": "Interactive"
  },
  "xamlStructure": "nested-element"
}
```

**XAML Example:**
```xml
<uix:SearchedElement>
  <uix:SearchedElement.Selector>&lt;webctrl tag='input' name='search' /&gt;</uix:SearchedElement.Selector>
</uix:SearchedElement>
```

**Scoping:** Nested within UI automation activity's Target property.

---

### 7. FilterX

**Namespace:** `clr-namespace:UiPath.Excel.Activities.Business` (ueab:)
**Usages:** 1 | **Files:** 1
**Complexity:** Moderate

**Description:** Applies a filter to an Excel range.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| Range | InArgument<IExcelRange> | The range to filter |
| ColumnName | InArgument<string> | Column to filter on |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Filter Range" | Activity display name |
| FilterArgument | InArgument<string> | null | Filter criteria |
| ClearFilter | bool | false | Clear existing filters first |

**Template JSON:**
```json
{
  "type": "FilterX",
  "namespace": "ueab",
  "category": "excel",
  "complexity": "moderate",
  "requiredAttributes": ["range", "columnName"],
  "optionalAttributes": ["displayName", "filterArgument", "clearFilter"],
  "template": {
    "type": "FilterX",
    "displayName": "Filter Range",
    "range": null,
    "columnName": null,
    "filterArgument": null,
    "clearFilter": false
  },
  "xamlStructure": "self-closing"
}
```

**XAML Example:**
```xml
<ueab:FilterX DisplayName="Filter by Status" Range="[Excel.Sheet(&quot;Data&quot;).Range(&quot;A1:E100&quot;)]" ColumnName="&quot;Status&quot;" FilterArgument="&quot;Active&quot;" />
```

**Scoping:** Must be within ExcelApplicationCard scope.

---

### 8. FindFirstLastDataRowX

**Namespace:** `clr-namespace:UiPath.Excel.Activities.Business` (ueab:)
**Usages:** 1 | **Files:** 1
**Complexity:** Moderate

**Description:** Finds the first and/or last data row in an Excel range.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| Range | InArgument<IExcelRange> | The range to search |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Find Data Rows" | Activity display name |
| ColumnName | InArgument<string> | null | Specific column to check |
| FirstRowIndex | OutArgument<int> | null | Output first row index |
| LastRowIndex | OutArgument<int> | null | Output last row index |

**Template JSON:**
```json
{
  "type": "FindFirstLastDataRowX",
  "namespace": "ueab",
  "category": "excel",
  "complexity": "moderate",
  "requiredAttributes": ["range"],
  "optionalAttributes": ["displayName", "columnName", "firstRowIndex", "lastRowIndex"],
  "template": {
    "type": "FindFirstLastDataRowX",
    "displayName": "Find First Last Data Row",
    "range": null,
    "columnName": null,
    "firstRowIndex": null,
    "lastRowIndex": null
  },
  "xamlStructure": "self-closing"
}
```

**XAML Example:**
```xml
<ueab:FindFirstLastDataRowX DisplayName="Find Data Bounds" Range="[Excel.Sheet(&quot;Data&quot;).Range(&quot;A:A&quot;)]" FirstRowIndex="[firstRow]" LastRowIndex="[lastRow]" />
```

**Scoping:** Must be within ExcelApplicationCard scope.

---

### 9. BuildDataTable

**Namespace:** `http://schemas.uipath.com/workflow/activities` (ui:)
**Usages:** 1 | **Files:** 1
**Complexity:** Moderate

**Description:** Creates a DataTable with defined schema (columns).

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| DataTable | OutArgument<DataTable> | Output DataTable variable |
| TableInfo | string | XML schema definition |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Build Data Table" | Activity display name |

**Template JSON:**
```json
{
  "type": "BuildDataTable",
  "namespace": "ui",
  "category": "data",
  "complexity": "moderate",
  "requiredAttributes": ["dataTable", "tableInfo"],
  "optionalAttributes": ["displayName"],
  "template": {
    "type": "BuildDataTable",
    "displayName": "Build Data Table",
    "dataTable": null,
    "tableInfo": null
  },
  "xamlStructure": "with-schema"
}
```

**XAML Example:**
```xml
<ui:BuildDataTable DisplayName="Build Customer Table" DataTable="[dt_Customers]">
  <ui:BuildDataTable.TableInfo>
    &lt;TableInfo&gt;
      &lt;Column ColumnName="Name" DataType="System.String" /&gt;
      &lt;Column ColumnName="Email" DataType="System.String" /&gt;
      &lt;Column ColumnName="Age" DataType="System.Int32" /&gt;
    &lt;/TableInfo&gt;
  </ui:BuildDataTable.TableInfo>
</ui:BuildDataTable>
```

**Implementation Notes:**
- TableInfo is XML-encoded within the XAML
- Column definitions include ColumnName, DataType, and optionally MaxLength, AllowDBNull
- Studio provides a visual designer for this activity

**Scoping:** None required - can be used anywhere.

---

## Tier 3: Complex (6-8 hours)

### 10. InterruptibleWhile

**Namespace:** `http://schemas.uipath.com/workflow/activities` (ui:)
**Usages:** 1 | **Files:** 1
**Complexity:** Complex

**Description:** A while loop that can be interrupted by external triggers or conditions.

**Required Attributes:**
| Attribute | Type | Description |
|-----------|------|-------------|
| Condition | InArgument<bool> | Loop continuation condition |
| Body | Activity | Activity to execute each iteration |

**Optional Attributes:**
| Attribute | Type | Default | Description |
|-----------|------|---------|-------------|
| DisplayName | string | "Interruptible While" | Activity display name |
| InterruptCondition | InArgument<bool> | null | Condition to interrupt loop |
| MaxIterations | int | -1 | Maximum iterations (-1 = unlimited) |

**Template JSON:**
```json
{
  "type": "InterruptibleWhile",
  "namespace": "ui",
  "category": "control_flow",
  "complexity": "complex",
  "requiredAttributes": ["condition", "body"],
  "optionalAttributes": ["displayName", "interruptCondition", "maxIterations"],
  "template": {
    "type": "InterruptibleWhile",
    "displayName": "Interruptible While",
    "condition": null,
    "interruptCondition": null,
    "maxIterations": -1,
    "body": {
      "variableName": null,
      "variableType": null,
      "activity": null
    }
  },
  "xamlStructure": "with-body"
}
```

**XAML Example:**
```xml
<ui:InterruptibleWhile DisplayName="Process Until Done" Condition="[hasMoreItems]" InterruptCondition="[userCancelled]">
  <ui:InterruptibleWhile.Body>
    <ActivityAction>
      <Sequence DisplayName="Process Item">
        <!-- Processing activities -->
        <ui:LogMessage DisplayName="Log Progress" Message="[&quot;Processing item &quot; &amp; currentItem]" />
      </Sequence>
    </ActivityAction>
  </ui:InterruptibleWhile.Body>
</ui:InterruptibleWhile>
```

**Implementation Notes:**
- Supports Continue and Break activities within body
- InterruptCondition checked at start of each iteration
- May use ActivityAction pattern for body (verify from XAML examples)

**Scoping:**
- Continue and Break must be within the InterruptibleWhile body
- Similar scoping rules to standard While loop

---

## Implementation Guidelines

### Template Format

All templates should follow this JSON structure:

```json
{
  "type": "ActivityName",
  "namespace": "prefix",
  "category": "category_name",
  "complexity": "simple|moderate|complex",
  "requiredAttributes": ["attr1", "attr2"],
  "optionalAttributes": ["attr3", "attr4"],
  "template": {
    "type": "ActivityName",
    "displayName": "Default Display Name",
    // ... attributes with default or null values
  },
  "xamlStructure": "self-closing|with-body|with-target|nested-element|with-schema"
}
```

### Testing Checklist

For each template:

- [ ] Validate with `xaml_constructor.py --mode template`
- [ ] Build test XAML with `xaml_constructor.py --mode build`
- [ ] Load generated XAML in UiPath Studio
- [ ] Execute workflow with sample data
- [ ] Verify all attributes render correctly
- [ ] Test edge cases (empty values, special characters)

### Documentation Updates

After implementing each template:

1. Add entry to `docs/implementation_priority.md`
2. Update category coverage in `RESEARCH_METRICS.md`
3. Add scoping rules to `docs/scoping_rules.md` if applicable
4. Update `TEMPLATE_QUICK_REFERENCE.md` lookup table

---

## Priority Order Recommendation

1. **DeleteFileX** - Simple, completes file operations
2. **ReadTextFile** - Simple, completes file operations
3. **SetToClipboard** - Simple, commonly needed
4. **AddDataRow** - Simple, enables data manipulation
5. **BuildDataTable** - Moderate, pairs with AddDataRow
6. **FilterX** - Moderate, improves Excel coverage
7. **FindFirstLastDataRowX** - Moderate, improves Excel coverage
8. **NMouseScroll** - Moderate, expands UI capabilities
9. **SearchedElement** - Moderate, expands UI capabilities
10. **InterruptibleWhile** - Complex, specialized use case

---

*Document version: 1.0*
*Last updated: February 5, 2026*
