/**
 * Format Normalizer — converts constructor-format JSON to writer-compatible format.
 *
 * Handles:
 * - Assign to/value: flat string → typed object {type, value}
 * - Sequence key: "activities" → "children"
 * - TryCatch catches: {type, x:TypeArguments, activityAction} → {exceptionType, variableName, handler}
 * - Expression brackets: bare `myVar` → `[myVar]`
 * - Metadata defaults: adds empty namespaces, assemblyReferences, arguments if missing
 */

export interface NormalizeResult {
  normalized: any;
  warnings: string[];
}

// Default namespaces every UiPath workflow needs
const DEFAULT_NAMESPACES = [
  "Microsoft.VisualBasic",
  "Microsoft.VisualBasic.Activities",
  "System",
  "System.Activities",
  "System.Activities.Expressions",
  "System.Activities.Statements",
  "System.Activities.Validation",
  "System.Activities.XamlIntegration",
  "System.Collections",
  "System.Collections.Generic",
  "System.Collections.ObjectModel",
  "System.Data",
  "System.Diagnostics",
  "System.Drawing",
  "System.IO",
  "System.Linq",
  "System.Net.Mail",
  "System.Windows",
  "System.Windows.Markup",
  "System.Xml",
  "System.Xml.Linq",
  "UiPath.Core",
  "UiPath.Core.Activities",
];

const DEFAULT_ASSEMBLY_REFERENCES = [
  "Microsoft.VisualBasic",
  "System",
  "System.Activities",
  "System.Core",
  "System.Data",
  "System.Drawing",
  "System.Linq",
  "System.Private.CoreLib",
  "System.Xml",
  "System.Xml.Linq",
  "UiPath.System.Activities",
  "mscorlib",
];

/**
 * Entry point: normalize a full workflow JSON object.
 */
export function normalizeWorkflowJson(input: any): NormalizeResult {
  const warnings: string[] = [];
  const normalized = structuredClone(input);

  // Ensure metadata exists with defaults
  if (!normalized.metadata) {
    normalized.metadata = {};
    warnings.push("Added missing metadata block");
  }
  const meta = normalized.metadata;

  if (!meta.class) {
    meta.class = "";
  }
  if (!meta.namespaces || !Array.isArray(meta.namespaces) || meta.namespaces.length === 0) {
    meta.namespaces = [...DEFAULT_NAMESPACES];
    warnings.push("Added default namespaces");
  }
  if (!meta.assemblyReferences || !Array.isArray(meta.assemblyReferences) || meta.assemblyReferences.length === 0) {
    meta.assemblyReferences = [...DEFAULT_ASSEMBLY_REFERENCES];
    warnings.push("Added default assemblyReferences");
  }
  if (!meta.arguments) {
    meta.arguments = [];
  }

  // Decompose argument types: "InArgument(x:String)" → {direction: "In", type: "x:String"}
  for (const arg of meta.arguments) {
    if (arg.type && !arg.direction) {
      const match = arg.type.match(/^(In|Out|InOut)Argument\((.+)\)$/);
      if (match) {
        arg.direction = match[1];
        arg.type = match[2];
        warnings.push(`Decomposed argument "${arg.name}" type: ${match[0]} → direction=${match[1]}, type=${match[2]}`);
      }
    }
  }

  // Ensure top-level variables array exists
  if (!normalized.variables) {
    normalized.variables = [];
  }

  // Normalize the workflow tree
  if (normalized.workflow) {
    normalized.workflow = normalizeActivity(normalized.workflow, warnings);
  }

  return { normalized, warnings };
}

/**
 * Recursively normalize an activity and its children.
 */
function normalizeActivity(activity: any, warnings: string[], parentType?: string): any {
  if (!activity || typeof activity !== "object") return activity;

  const type = activity.type;

  // Rename "activities" → "children" for Sequence-like types
  if (activity.activities && !activity.children) {
    activity.children = activity.activities;
    delete activity.activities;
    warnings.push(`Renamed "activities" to "children" in ${type || "unknown"}`);
  }

  // Type-specific normalizations
  switch (type) {
    case "Assign":
      normalizeAssign(activity, warnings);
      break;
    case "TryCatch":
      normalizeTryCatch(activity, warnings);
      break;
    case "ForEach":
      normalizeForEach(activity, warnings);
      break;
    case "Switch":
      normalizeSwitch(activity, warnings);
      break;
  }

  // Normalize expression fields that commonly hold VB expressions
  normalizeExpressionFields(activity, type, warnings);

  // Recurse into children
  if (Array.isArray(activity.children)) {
    activity.children = activity.children.map((child: any) =>
      normalizeActivity(child, warnings, type)
    );
  }

  // Recurse into If branches
  if (activity.then) {
    activity.then = normalizeActivity(activity.then, warnings, type);
  }
  if (activity.else) {
    activity.else = normalizeActivity(activity.else, warnings, type);
  }

  // Recurse into TryCatch branches
  if (activity.try) {
    activity.try = normalizeActivity(activity.try, warnings, type);
  }
  if (activity.finally) {
    activity.finally = normalizeActivity(activity.finally, warnings, type);
  }

  // Recurse into catch handlers
  if (Array.isArray(activity.catches)) {
    for (const c of activity.catches) {
      if (c.handler) {
        c.handler = normalizeActivity(c.handler, warnings, "Catch");
      }
    }
  }

  // Recurse into ForEach/While body
  if (activity.body && typeof activity.body === "object") {
    if (activity.body.activity) {
      activity.body.activity = normalizeActivity(activity.body.activity, warnings, type);
    }
    // If body is directly an activity (not the {variableName, activity} structure)
    if (activity.body.type && !activity.body.activity) {
      activity.body = normalizeActivity(activity.body, warnings, type);
    }
  }

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

  return activity;
}

/**
 * Normalize Assign activity: flat string to/value → typed objects.
 */
function normalizeAssign(activity: any, warnings: string[]): void {
  if (activity.to !== undefined && typeof activity.to === "string") {
    const varName = activity.to;
    activity.to = normalizeTypedField(varName, "String", true);
    warnings.push(`Converted Assign.to from flat string "${varName}" to typed object`);
  }
  if (activity.value !== undefined && typeof activity.value === "string") {
    const rawValue = activity.value;
    const inferredType = activity.to?.type || inferType(rawValue);
    activity.value = normalizeTypedField(rawValue, inferredType, false);
    warnings.push(`Converted Assign.value from flat string to typed object`);
  }
}

/**
 * Normalize TryCatch: convert constructor-format catches to writer format.
 */
function normalizeTryCatch(activity: any, warnings: string[]): void {
  if (!Array.isArray(activity.catches)) return;

  activity.catches = activity.catches.map((c: any) => {
    // Already in writer format
    if (c.exceptionType && c.handler) return c;

    // Constructor format: {type: "Catch", "x:TypeArguments": "s:Exception", activityAction: {...}}
    if (c.type === "Catch" || c["x:TypeArguments"] || c.activityAction) {
      const typeArg = c["x:TypeArguments"] || c.typeArguments || "s:Exception";
      // Strip namespace prefix (s:Exception → Exception)
      const exceptionType = typeArg.includes(":") ? typeArg.split(":")[1] : typeArg;

      let variableName = "ex";
      let handler = null;

      if (c.activityAction) {
        // Extract from full ActivityAction structure
        const action = c.activityAction;
        if (action.argument?.name) {
          variableName = action.argument.name;
        }
        handler = action.handler || action.activity || null;
      } else if (c.handler) {
        handler = c.handler;
      }

      if (c.variableName) {
        variableName = c.variableName;
      }

      warnings.push(`Converted constructor-format catch (${exceptionType}) to writer format`);

      return {
        exceptionType,
        variableName,
        handler,
      };
    }

    // Already has exceptionType but maybe missing handler key
    if (c.exceptionType) {
      return {
        exceptionType: c.exceptionType,
        variableName: c.variableName || "ex",
        handler: c.handler || c.body || null,
      };
    }

    return c;
  });
}

/**
 * Normalize ForEach: ensure body has the right structure.
 */
function normalizeForEach(activity: any, warnings: string[]): void {
  // If values is a flat string, wrap in brackets
  if (typeof activity.values === "string") {
    activity.values = ensureBrackets(activity.values);
  }

  // Normalize typeArguments: strip "x:" prefix if present
  if (activity.typeArguments) {
    const ta = activity.typeArguments;
    if (ta.startsWith("x:")) {
      activity.typeArguments = ta.substring(2);
      warnings.push(`Stripped "x:" prefix from ForEach typeArguments`);
    }
  }

  // Ensure body structure
  if (!activity.body) {
    activity.body = {
      variableName: "item",
      variableType: activity.typeArguments || "String",
      activity: null,
    };
  }
}

/**
 * Normalize Switch: ensure cases is [{key, activity}] array (writer format).
 *
 * Accepted inputs:
 *   - [{key, value}]        → [{key, activity: value}]
 *   - [{case, activity}]    → [{key: case, activity}]
 *   - [{key, activity}]     → pass-through (already correct)
 *   - {key: activity, ...}  → [{key, activity}, ...]  (object→array)
 */
function normalizeSwitch(activity: any, warnings: string[]): void {
  if (Array.isArray(activity.cases)) {
    let needsConversion = false;
    const normalized: Array<{ key: string; activity: any }> = [];
    for (const c of activity.cases) {
      if (c.key !== undefined && c.activity !== undefined) {
        // Already in writer format
        normalized.push({ key: String(c.key), activity: c.activity });
      } else if (c.key !== undefined && c.value !== undefined) {
        normalized.push({ key: String(c.key), activity: c.value });
        needsConversion = true;
      } else if (c.case !== undefined && c.activity !== undefined) {
        normalized.push({ key: String(c.case), activity: c.activity });
        needsConversion = true;
      }
    }
    if (needsConversion) {
      activity.cases = normalized;
      warnings.push("Converted Switch cases from array to writer format");
    }
  } else if (
    activity.cases &&
    typeof activity.cases === "object" &&
    !Array.isArray(activity.cases)
  ) {
    // Object format {key: activity} → [{key, activity}]
    const normalized: Array<{ key: string; activity: any }> = [];
    for (const [key, value] of Object.entries(activity.cases)) {
      normalized.push({ key, activity: value });
    }
    activity.cases = normalized;
    warnings.push("Converted Switch cases from object to writer format");
  }
}

/**
 * Convert a flat string value to a typed object {type, value}.
 */
function normalizeTypedField(
  value: string,
  defaultType: string,
  isVariable: boolean
): { type: string; value: string } {
  const type = isVariable ? defaultType : inferType(value, defaultType);
  const bracketedValue = ensureBrackets(value);
  return { type, value: bracketedValue };
}

/**
 * Ensure a value expression is wrapped in brackets.
 * - Already bracketed: [myVar] → [myVar] (no change)
 * - Bare variable: myVar → [myVar]
 * - String literal: "Hello" → ["Hello"]
 * - Number literal: 42 → [42]
 * - Empty string: "" → "" (no change)
 */
export function ensureBrackets(value: string): string {
  if (!value || value.length === 0) return value;

  // Already has brackets
  if (value.startsWith("[") && value.endsWith("]")) return value;

  return `[${value}]`;
}

/**
 * Infer the VB.NET type from a value string.
 */
function inferType(value: string, defaultType = "String"): string {
  if (!value) return defaultType;

  // Strip brackets for analysis
  const raw = value.startsWith("[") && value.endsWith("]")
    ? value.slice(1, -1)
    : value;

  // Boolean
  if (raw === "True" || raw === "False") return "Boolean";

  // Integer (whole number, no decimal)
  if (/^-?\d+$/.test(raw)) return "Int32";

  // Double (has decimal point)
  if (/^-?\d+\.\d+$/.test(raw)) return "Double";

  // String literal (quoted)
  if (raw.startsWith('"') && raw.endsWith('"')) return "String";

  return defaultType;
}

/**
 * Normalize expression-holding fields on activities.
 * Ensures fields like condition, message, expression are bracketed.
 */
function normalizeExpressionFields(activity: any, type: string | undefined, warnings: string[]): void {
  // Fields that hold VB expressions and need brackets
  const expressionFields: Record<string, string[]> = {
    If: ["condition"],
    While: ["condition"],
    DoWhile: ["condition"],
    Switch: ["expression"],
    LogMessage: ["message"],
  };

  if (!type) return;

  const fields = expressionFields[type];
  if (!fields) return;

  for (const field of fields) {
    if (typeof activity[field] === "string" && activity[field].length > 0) {
      const original = activity[field];
      activity[field] = ensureBrackets(original);
      if (activity[field] !== original) {
        warnings.push(`Added brackets to ${type}.${field}`);
      }
    }
  }
}
