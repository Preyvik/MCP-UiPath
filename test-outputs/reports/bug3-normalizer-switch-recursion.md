# Bug 3: Normalizer Doesn't Recurse Into Switch Case or Default Activities

**Reported:** 2026-02-24
**Found in:** Test 7 — Full Normalizer Pipeline
**Severity:** Medium
**Status:** Fixed
**File:** `src/format-normalizer.ts`

---

## Symptom

`create_workflow` fails with:

```
Error: 'str' object has no attribute 'get'
```

when a Switch activity contains Assign activities using flat constructor-format strings for `to`/`value`.

## Reproduction

Pass this JSON through `create_workflow`:

```json
{
  "type": "Switch",
  "displayName": "Route",
  "typeArguments": "String",
  "expression": "status",
  "cases": {
    "High": {
      "type": "Assign",
      "displayName": "Set High",
      "to": "counter",
      "value": "counter * 10"
    }
  },
  "default": {
    "type": "Assign",
    "displayName": "Set Default",
    "to": "counter",
    "value": "counter"
  }
}
```

**Expected:** Normalizer converts flat Assign strings to typed objects before the writer processes them.

**Actual:** The Assign activities inside `cases` and `default` are passed un-normalized to the writer. The Python writer calls `.get()` on `to` (a plain string), which crashes.

## Root Cause

`normalizeActivity()` in `src/format-normalizer.ts` (lines 147-188) recursively normalizes child activities for:
- `children` (Sequence) — line 148
- `then` / `else` (If) — lines 155-160
- `try` / `finally` (TryCatch) — lines 163-168
- `catches[].handler` (TryCatch) — lines 171-177
- `body.activity` / `body` (ForEach, While) — lines 180-188

**Missing recursion targets:**
- `Switch.cases[].activity` — never recursed into
- `Switch.default` — never recursed into

After `normalizeSwitch()` converts the case structure (object → array), the activities inside each case are left un-normalized. The normalizer returns, and the writer receives raw constructor-format Assign objects.

## Affected Code

```typescript
// src/format-normalizer.ts, after line 188
// These recursion paths are MISSING:

// Recurse into Switch cases
if (Array.isArray(activity.cases)) {
  for (const c of activity.cases) {
    if (c.activity) {
      c.activity = normalizeActivity(c.activity, warnings, "Switch");
    }
  }
}

// Recurse into Switch default
if (activity.default && typeof activity.default === "object" && activity.default.type) {
  activity.default = normalizeActivity(activity.default, warnings, "Switch");
}
```

## Impact

- **Switch + flat Assign** → crash (`'str' object has no attribute 'get'`)
- **Switch + flat LogMessage** → likely crash or missing bracket normalization on nested `message` fields
- **Switch + any nested container** (Sequence, If, TryCatch inside a case) → children of those containers also un-normalized
- Any activity type nested inside a Switch case that requires normalization will fail

## Workaround

Provide activities inside Switch `cases` and `default` in writer format (typed objects):

```json
{
  "cases": {
    "High": {
      "type": "Assign",
      "displayName": "Set High",
      "to": {"type": "Int32", "value": "[counter]"},
      "value": {"type": "Int32", "value": "[counter * 10]"}
    }
  }
}
```

## Proposed Fix

Add two recursion blocks to `normalizeActivity()` after line 188 in `src/format-normalizer.ts`:

```typescript
// Recurse into Switch cases
if (Array.isArray(activity.cases)) {
  for (const c of activity.cases) {
    if (c.activity) {
      c.activity = normalizeActivity(c.activity, warnings, "Switch");
    }
  }
}

// Recurse into Switch default
if (activity.default && typeof activity.default === "object" && activity.default.type) {
  activity.default = normalizeActivity(activity.default, warnings, "Switch");
}
```

**Placement:** After the existing `body` recursion block (line 188), before the `return activity;` statement (line 190).

**Note:** The recursion must happen **after** `normalizeSwitch()` is called (line 140), because `normalizeSwitch()` converts the case structure first. The existing code order (type-specific normalization at lines 129-141, then recursion at lines 147+) already satisfies this — the new blocks just need to be added to the recursion section.

## Test Validation

After fix, re-run Test 7 with flat Assign strings inside Switch cases (remove the writer-format workaround) and confirm:
1. `create_workflow` returns `success: true`
2. Normalizer warnings include Assign conversions for case activities
3. `read_workflow` shows typed `to`/`value` objects in Switch case Assigns
