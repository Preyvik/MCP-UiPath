# MCP-UiPath Smart Tooling — End-to-End Test Report

**Date:** 2026-02-24
**Tested by:** Claude Opus 4.6 (automated)
**Build:** `npm run build` — clean (0 errors)
**MCP Tools:** 7 registered

---

## Test Results

| Test | File | Status | Errors | Warnings | HR Compliance | Info Issues |
|------|------|--------|--------|----------|---------------|-------------|
| 1 | `Test1_DataValidation.xaml` | **PASS** | 0 | 0 | All PASS | None |
| 2 | `Test2_OrderRouter.xaml` | **PASS** | 0 | 0 | All PASS | HR-3 INFO (nested Sequences) |
| 3 | `Test3_FileProcessor.xaml` | **PASS** | 0 | 0 | All PASS | HR-3 INFO (4 nested Sequences) |
| 4 | `Test4_ExcelReport.xaml` | **PASS** | 0 | 0 | All PASS | None |
| 5 | `Test5_StressTest.xaml` | **PASS** | 0 | 0 | All PASS | HR-3 INFO (6 nested Sequences) |
| 6 | `Test6_RoundTrip.xaml` | **PASS** | 0 | 0 | All PASS | HR-3 INFO (5 nested Sequences) |
| 7 | `Test7_NormalizerPipeline.xaml` | **PASS** | 0 | 0 | All PASS | HR-3 INFO (3 nested Sequences) |

**Overall: 7/7 PASS — 0 critical errors, 0 warnings, 5 INFO-level notices (all HR-3 nested container advisories)**

---

## Test Details

### Test 1 — Data Validation Pipeline (`Test1_DataValidation.xaml`)

**Purpose:** Validate typed Assign generation and type inference across 4 data types.

**Activities:** Sequence, Assign x5, LogMessage x2

**Arguments:**
- `in_RawName` (In, String)
- `in_RawAge` (In, Int32)
- `out_IsValid` (Out, Boolean)

**Variables:** `cleanName` (String), `ageValue` (Int32), `score` (Double), `isAdult` (Boolean), `validationMsg` (String)

**Verification:**
- 5 Assign activities with typed `OutArgument`/`InArgument` elements covering String, Int32, Double, Boolean
- 2 LogMessage activities with bracketed expressions
- Arguments render with correct types in XAML

---

### Test 2 — Order Router (`Test2_OrderRouter.xaml`)

**Purpose:** Validate Switch with multiple cases, If branching, and activities-to-children rename.

**Activities:** Sequence, Switch (3 cases + default), If, LogMessage x4, Assign x7

**Arguments:**
- `in_OrderType` (In, String)
- `in_OrderAmount` (In, Double)
- `out_RoutingResult` (Out, String)

**Switch Cases:**
- `Standard` → Sequence (Assign + LogMessage)
- `Express` → Sequence (Assign + LogMessage)
- `Priority` → single Assign
- `default` → single Assign

**Verification:**
- All 3 keyed cases present in XAML (`x:Key="Standard"`, `x:Key="Express"`, `x:Key="Priority"`)
- Default case present in `Switch.Default`
- If condition correctly bracketed and XML-escaped

---

### Test 3 — Resilient File Processor (`Test3_FileProcessor.xaml`)

**Purpose:** Validate TryCatch, ForEach, While, and nested Assigns at 3+ depth levels.

**Activities:** Sequence, ForEach, TryCatch, While, Assign x7, LogMessage x4

**Arguments:**
- `in_FilePaths` (In, `scg:List(x:String)`)
- `in_MaxRetries` (In, Int32, default=3)
- `out_ProcessedCount` (Out, Int32)

**Nesting:** ForEach > TryCatch > Sequence > While > Sequence > Assign

**Verification:**
- TryCatch with `Catch(System.Exception)` and exception variable `ex`
- ForEach with `x:TypeArguments="x:String"` and `values="[in_FilePaths]"`
- While with compound condition `[keepRetrying AndAlso retryCount < in_MaxRetries]`
- Generic `List(String)` argument type correctly rendered

---

### Test 4 — Excel Report Generator (`Test4_ExcelReport.xaml`)

**Purpose:** Validate Excel scoping, Throw, Delay, and passthrough of already-correct writer-format JSON.

**Activities:** Sequence, If, Throw, ExcelProcessScopeX, ReadRangeX, WriteRangeX, Delay, Assign x3, LogMessage x3

**Arguments:**
- `in_ExcelPath` (In, String)
- `in_SheetName` (In, String, default="Sheet1")
- `out_RowCount` (Out, Int32)

**Excel Scoping:** `ExcelProcessScopeX` > `ReadRangeX` + `Assign` + `WriteRangeX`

**Verification:**
- Excel activities properly nested inside `ExcelProcessScopeX`
- Throw with `New ArgumentException(...)` expression
- Delay with `TimeSpan.FromSeconds(2)` duration
- DataTable variable declared for `dtData`

---

### Test 5 — Multi-Format Stress Test (`Test5_StressTest.xaml`)

**Purpose:** Exercise every core activity type in a single workflow.

**Activities:** Sequence x6, Assign x8, LogMessage x8, ForEach, Switch (3 cases + default), While, If, TryCatch (2 catches + finally), Delay

**Arguments:**
- `in_InputList` (In, `scg:List(x:String)`)
- `in_Mode` (In, String, default="Normal")
- `out_Summary` (Out, String)
- `io_Counter` (InOut, Int32)

**TryCatch:** 2 catches (`Exception` + `ArgumentException`) + `finally` block

**Verification:**
- All 4 argument directions (In, Out, InOut) represented
- 2 distinct catch blocks with different exception types
- Finally block present
- Switch with 3 cases (Fast, Normal, Slow) + default
- ForEach iterating `scg:List(x:String)`
- While with compound Boolean condition

---

### Test 6 — Round-Trip (`Test6_RoundTrip.xaml`)

**Purpose:** Verify read-modify-write fidelity — read existing XAML, write it back, read again, compare.

**Source:** `TestWorkflow_CustomerStatusCheck.xaml`

**Comparison Results:**

| Aspect | Match |
|--------|-------|
| Activity tree shape | Identical |
| displayNames | Identical |
| Expressions (all VB.NET) | Identical |
| Variables (names, types, defaults) | Identical |
| Arguments (names, directions, types) | Identical |
| hintSize values | Preserved |
| idRef values | Preserved |
| viewState (IsExpanded, IsPinned) | Preserved |

---

### Test 7 — Full Normalizer Pipeline (`Test7_NormalizerPipeline.xaml`)

**Purpose:** Exercise every normalizer conversion path end-to-end via `create_workflow` (tests 1-5 used `write_workflow` which bypasses the normalizer). Validates both bug fixes (Switch cases dropped, argument double-wrapping) and all format normalization paths.

**Pipeline:** constructor-format JSON → normalize → write → xmlns fix → lint

**Activities:** Sequence x3, Assign x12, LogMessage x6, If, ForEach, TryCatch, Switch (3 cases + default), While

**Arguments (pre-wrapped — tests Bug 2 fix):**
- `in_Name` — `InArgument(x:String)` → decomposed to direction=In, type=x:String
- `in_Items` — `InArgument(scg:List(x:String))` → decomposed to direction=In, type=scg:List(x:String)
- `out_Count` — `OutArgument(x:Int32)` → decomposed to direction=Out, type=x:Int32
- `io_Flag` — `InOutArgument(x:Boolean)` → decomposed to direction=InOut, type=x:Boolean

**Variables:** `greeting` (String), `counter` (Int32, default: 0), `ratio` (Double), `isReady` (Boolean), `keepGoing` (Boolean, default: True), `priority` (String, default: "Medium")

**Normalizer Path Coverage:**

| # | Path | Input (constructor) | Output (writer) | Status |
|---|------|---------------------|-----------------|--------|
| 1 | Assign flat → typed | `to: "greeting"`, `value: "\"Hello \" & in_Name"` | `to: {type: "String", value: "[greeting]"}` | PASS (8 Assigns converted) |
| 2 | `activities` → `children` | Root Sequence uses `"activities"` key | Renamed to `"children"` in normalized JSON | PASS |
| 3 | Argument decomposition (Bug 2) | `"InArgument(x:String)"` | `direction: "In", type: "x:String"` | PASS (4 args, no double-wrapping) |
| 4a | Switch object → array (Bug 1) | `cases: {"High": ..., "Medium": ..., "Low": ...}` | `[{key: "High", activity: ...}, ...]` — all 3 cases + default in XAML | PASS |
| 5 | TryCatch constructor catches | `{type: "Catch", "x:TypeArguments": "s:Exception", activityAction: {...}}` | `{exceptionType: "Exception", variableName: "ex", handler: {...}}` | PASS |
| 6 | ForEach bare values | `values: "in_Items"` (no brackets) | `values: "[in_Items]"` | PASS |
| 7 | ForEach x: prefix strip | `typeArguments: "x:String"` | `typeArguments: "String"` | PASS |
| 8 | Bare expression bracketing | `condition: "isReady"`, `expression: "priority"`, etc. | `[isReady]`, `[priority]`, `["Starting pipeline"]` | PASS (7 expressions bracketed) |
| 9 | Metadata auto-fill | `namespaces: []`, `assemblyReferences: []` | 28 namespaces, 12 assembly references | PASS |
| 10 | Type inference | `"42"`, `"3.14"`, `"True"`, `"\"Hello \" & in_Name"` | Int32, Double, Boolean, String | PASS |
| 11 | Validation auto-run | — | `{is_valid: true, error_count: 0, warning_count: 0}` | PASS |

**Normalizer Warnings (35 total):**
- 2× metadata auto-fill (namespaces, assemblyReferences)
- 4× argument decomposition (in_Name, in_Items, out_Count, io_Flag)
- 1× `activities` → `children` rename
- 16× flat Assign conversion (8 `to` + 8 `value`)
- 7× bare expression bracketing (4 LogMessage, 1 If, 1 Switch, 1 While)
- 1× ForEach `x:` prefix stripping
- 1× constructor-format catch conversion
- 1× Switch object→array conversion
- 2× ForEach values bracket addition (implicit, within expression bracketing)

**XAML Verification (via `read_workflow`):**
- Arguments: `in_Name` (In/String), `in_Items` (In/List<String>), `out_Count` (Out/Int32), `io_Flag` (InOut/Boolean) — no double-wrapping
- Switch: 3 keyed cases (`x:Key="High"`, `x:Key="Medium"`, `x:Key="Low"`) + `Switch.Default`
- ForEach: `x:TypeArguments="x:String"`, `Values="[in_Items]"`
- TryCatch: `<Catch x:TypeArguments="Exception">` with `DelegateInArgument Name="ex"`
- All expressions XML-escaped: `&amp;` for `&`, `&lt;` for `<`
- While condition: `[keepGoing AndAlso counter &lt; 100]`

**Bug Discovered During Testing:**

**Bug 3: Normalizer doesn't recurse into Switch case activities**

The normalizer converts Switch cases from object format `{"High": activity}` to array format `[{key: "High", activity: {...}}]`, but does **not** recurse into the nested `activity` objects to normalize flat Assign strings. This causes `'str' object has no attribute 'get'` when the writer processes un-normalized Assign `to`/`value` fields.

- **Workaround:** Provide Switch case activities in writer format (typed `to`/`value` objects)
- **Scope:** Only affects Assign activities directly inside Switch cases that use flat strings
- **Severity:** Medium — the `create_workflow` call fails rather than silently producing bad output
- **File:** `src/format-normalizer.ts` — `normalizeSwitch()` converts case structure but doesn't call `normalizeAssign()` on nested activities

**Paths not tested (plan scope reduction):**
- 4b: Switch `[{key, value}]` array format — not included (only object format tested)
- 4c: Switch `[{case, activity}]` array format — not included (only object format tested)

---

## Activity Coverage Matrix

| Activity Type | T1 | T2 | T3 | T4 | T5 | T6 | T7 |
|---------------|:--:|:--:|:--:|:--:|:--:|:--:|:--:|
| Sequence | x | x | x | x | x | x | x |
| Assign | x | x | x | x | x | x | x |
| LogMessage | x | x | x | x | x | x | x |
| If | | x | | x | x | x | x |
| Switch | | x | | | x | | x |
| TryCatch | | | x | | x | x | x |
| ForEach | | | x | | x | | x |
| While | | | x | | x | | x |
| Throw | | | | x | | | |
| Delay | | | | x | x | | |
| ExcelProcessScopeX | | | | x | | | |
| ReadRangeX | | | | x | | | |
| WriteRangeX | | | | x | | | |

---

## Bugs Found & Fixed During Testing

### Bug 1: Switch cases silently dropped

**Symptom:** `create_workflow` produces XAML with empty Switch (only `Switch.Default`, no keyed cases).

**Root cause:** `normalizeSwitch()` in `src/format-normalizer.ts` converted array cases `[{key, value}]` to object format `{key: value}`, but the Python writer (`xaml_syntaxer.py:3257`) expects array format `[{key, activity}]`. When the writer iterated over a dict, it got string keys and then failed with `"string indices must be integers, not 'str'"` or silently produced 0 cases.

**Fix:** Rewrote `normalizeSwitch()` to normalize all input formats to the writer-expected `[{key, activity}]` array format. Handles 4 input variants:
- `[{key, activity}]` — pass-through (already correct)
- `[{key, value}]` — rename `value` → `activity`
- `[{case, activity}]` — rename `case` → `key`
- `{key: activity, ...}` — object → array conversion

**File:** `src/format-normalizer.ts` lines 282-318

### Bug 2: Argument types double-wrapped

**Symptom:** XAML shows `Type="InArgument(InArgument(x:String))"` instead of `Type="InArgument(x:String)"`.

**Root cause:** The writer's `apply_arguments()` always wraps the type in `InArgument(...)` based on direction. But when the JSON input uses the user-facing format `"type": "InArgument(x:String)"`, the type is already wrapped, producing double-wrapping.

**Fix:** Added argument decomposition in the normalizer: `"InArgument(x:String)"` → `{direction: "In", type: "x:String"}` before passing to the writer. Matches the format the writer's reader (`_parse_argument`) outputs.

**File:** `src/format-normalizer.ts` lines 87-97

### Bug 3: Normalizer doesn't recurse into Switch case activities (found in Test 7)

**Symptom:** `create_workflow` fails with `'str' object has no attribute 'get'` when Switch cases contain Assign activities with flat string `to`/`value` fields.

**Root cause:** `normalizeSwitch()` in `src/format-normalizer.ts` converts the case _structure_ (object → array, key renaming) but does not recurse into the `activity` objects within each case to apply `normalizeAssign()` or other activity-level normalizations. The writer then receives un-normalized Assign activities and calls `.get()` on a string, causing the crash.

**Workaround:** Provide activities inside Switch cases in writer format (typed objects for `to`/`value`).

**Severity:** Medium — the pipeline fails loudly (no silent data loss), and only affects Assign activities directly inside Switch cases when using flat strings.

**File:** `src/format-normalizer.ts` — `normalizeSwitch()` needs to call `normalizeActivity()` recursively on each case activity.

### Status

Bug 1 and Bug 2 fixes are compiled to `build/` and confirmed working end-to-end in Test 7 via `create_workflow`.
Bug 3 is newly discovered — workaround documented, fix pending.
Tests 1-6 were executed using `write_workflow` (bypasses normalizer). Test 7 is the first end-to-end normalizer pipeline test via `create_workflow`.

---

## Test Methodology

- Tests 1-5 generated via `write_workflow` (low-level XAML writer) with writer-native format JSON
- Test 6 generated via `read_workflow` → `write_workflow` round-trip
- **Test 7 generated via `create_workflow` (full pipeline)** with deliberately constructor-format JSON to exercise all normalizer conversion paths
- All 7 XAMLs validated with `validate_workflow` linter
- Test 2 Switch cases verified via Python XML element inspection
- Test 7 verified via `read_workflow` round-trip read to confirm structural fidelity

---

## XAML Files for UiPath Studio Verification

```
test-outputs/Test1_DataValidation.xaml    — Sequence, 5 Assigns (4 types), 2 LogMessages
test-outputs/Test2_OrderRouter.xaml       — Switch (3 cases), If, Assigns, LogMessages
test-outputs/Test3_FileProcessor.xaml     — ForEach, While, TryCatch, nested Assigns
test-outputs/Test4_ExcelReport.xaml       — ExcelProcessScopeX, ReadRangeX, WriteRangeX, Throw, Delay
test-outputs/Test5_StressTest.xaml        — All core types, 2 catches + finally, 4 args (In/Out/InOut)
test-outputs/Test6_RoundTrip.xaml         — Mirror of TestWorkflow_CustomerStatusCheck.xaml
test-outputs/Test7_NormalizerPipeline.xaml — Full normalizer pipeline (create_workflow), 12 Assigns, Switch 3-case, ForEach+TryCatch
```

### Studio Verification Checklist
- [ ] No load errors on any XAML
- [ ] Activity tree renders correctly
- [ ] Variables and arguments appear with correct types
- [ ] Expressions compile (no red underlines)
- [ ] Excel activities properly scoped (Test 4)
- [ ] Switch cases visible (Test 2, Test 5, Test 7)
- [ ] Test 7 arguments show correct types (no double-wrapping)
- [ ] Test 7 ForEach iterates with bracketed values
