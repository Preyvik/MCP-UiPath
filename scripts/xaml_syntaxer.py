#!/usr/bin/env python3
"""
XAML Syntaxer - Bidirectional XAML-JSON Conversion for UiPath Workflows

This script provides conversion between UiPath XAML workflow files and a simplified
JSON representation. It supports read mode (XAML to JSON) and write mode (JSON to XAML).

Supported Activities:

Core Activities:
- Sequence          - Container for sequential activities
- Assign            - Variable assignment
- If                - Conditional branching
- LogMessage        - Logging messages
- InvokeWorkflowFile - Invoke external workflows
- Flowchart         - Flowchart container with nodes and connections
- FlowStep          - Flowchart step node wrapping an activity
- FlowDecision      - Flowchart decision node with True/False branches

Control Flow:
- Switch            - Multi-way branching based on expression
- TryCatch          - Exception handling with try/catch/finally
- ForEach           - Iterate over collections
- ForEachRow        - Iterate over DataTable rows
- Rethrow           - Re-throw caught exceptions
- While             - Loop while condition is true
- InterruptibleWhile - While loop with interrupt condition support
- Continue          - Skip to next loop iteration
- Break             - Exit current loop immediately
- Delay             - Pause execution for a duration

Excel Activities (Modern):
- ExcelProcessScopeX    - Excel process scope container
- ExcelApplicationCard  - Open and manage Excel workbooks
- ReadRangeX            - Read data from Excel range
- SaveExcelFileX        - Save Excel file
- WriteCellX            - Write to a single cell
- WriteRangeX           - Write DataTable to range
- CopyPasteRangeX       - Copy and paste range

File Operations:
- CreateDirectory   - Create a directory
- MoveFile          - Move or rename a file
- ReadRange         - Legacy Excel read range
- PathExists        - Check if file or folder exists
- DeleteFileX       - Delete a file at a specified path
- ReadTextFile      - Read all text from a file

Data Activities:
- AddDataRow        - Add a row to a DataTable

Process Activities:
- KillProcess       - Terminate a process by name

Utilities:
- CommentOut        - Comment out (disable) activities
- RetryScope        - Retry activities on failure
- SetToClipboard    - Set text to the system clipboard

UI Automation Activities (Modern):
- NApplicationCard  - Use Application/Browser scope (with TargetApp, OCREngine)
- NClick            - Click UI element (with TargetAnchorable selectors)
- NTypeInto         - Type into UI element (with TargetAnchorable selectors)
- NCheckState       - Check App State (with IfExists/IfNotExists branches)
- NMouseScroll      - Mouse scroll within UI element (with TargetAnchorable, SearchedElement)

Usage:
    Read Mode:
        python xaml_syntaxer.py --mode read --input workflow.xaml --output workflow.json

    Write Mode:
        python xaml_syntaxer.py --mode write --input workflow.json --output workflow.xaml

Author: Claude Code
Version: 2.1.0
"""

import xml.etree.ElementTree as ET
import json
import argparse
import sys
import re
import copy
from pathlib import Path
from abc import ABC, abstractmethod
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field


# =============================================================================
# Namespace Constants
# =============================================================================

NAMESPACES = {
    '': 'http://schemas.microsoft.com/netfx/2009/xaml/activities',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mva': 'clr-namespace:Microsoft.VisualBasic.Activities;assembly=System.Activities',
    's': 'clr-namespace:System;assembly=System.Private.CoreLib',
    'sap': 'http://schemas.microsoft.com/netfx/2009/xaml/activities/presentation',
    'sap2010': 'http://schemas.microsoft.com/netfx/2010/xaml/activities/presentation',
    'scg': 'clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib',
    'sco': 'clr-namespace:System.Collections.ObjectModel;assembly=System.Private.CoreLib',
    'sd': 'clr-namespace:System.Data;assembly=System.Data.Common',
    'sd1': 'clr-namespace:System.Drawing;assembly=System.Drawing.Primitives',
    'sd2': 'clr-namespace:System.Drawing;assembly=System.Drawing.Common',
    'ui': 'http://schemas.uipath.com/workflow/activities',
    'ue': 'clr-namespace:UiPath.Excel;assembly=UiPath.Excel.Activities',
    'ueab': 'clr-namespace:UiPath.Excel.Activities.Business;assembly=UiPath.Excel.Activities',
    'uix': 'http://schemas.uipath.com/workflow/activities/uix',
    'x': 'http://schemas.microsoft.com/winfx/2006/xaml',
    'av': 'http://schemas.microsoft.com/winfx/2006/xaml/presentation',
}

# Reverse mapping for namespace URI to prefix
NS_URI_TO_PREFIX = {v: k for k, v in NAMESPACES.items()}

# Baseline namespace prefixes that must always be present in Writer output
DEFAULT_REQUIRED_NAMESPACES = {
    '',        # default namespace (activities)
    'ui',      # UiPath activities
    'x',       # XAML core
    'sap',     # presentation
    'sap2010', # presentation 2010
    'mc',      # markup compatibility
    'mva',     # VisualBasic.Activities
    'sco',     # System.Collections.ObjectModel (for namespace/assembly collections)
    's',       # System (DateTime, TimeSpan)
    'sd',      # System.Data (DataTable, DataRow)
    'scg',     # System.Collections.Generic (List, Dictionary)
    'ue',      # UiPath.Excel
    'ueab',    # UiPath.Excel.Activities.Business
    'uix',     # UiPath.UIAutomationNext
    'av',      # System.Windows (presentation)
    'sd1',     # System.Drawing.Primitives
    'sd2',     # System.Drawing.Common
}

# Mapping from xmlns prefix to CLR namespace strings for TextExpression.NamespacesForImplementation
PREFIX_TO_CLR_NAMESPACES = {
    's': ['System'],
    'sd': ['System.Data'],
    'scg': ['System.Collections.Generic'],
    'sco': ['System.Collections.ObjectModel'],
    'ue': ['UiPath.Excel', 'UiPath.Excel.Activities', 'UiPath.Excel.Activities.Business'],
    'ueab': ['UiPath.Excel.Activities.Business'],
    'ui': ['UiPath.Core', 'UiPath.Core.Activities'],
    'sd1': ['System.Drawing'],
    'sd2': ['System.Drawing'],
    'uix': ['UiPath.UIAutomationNext.Activities', 'UiPath.UIAutomationNext.Enums'],
    'av': ['System.Windows', 'System.Windows.Markup'],
}

# Mapping from xmlns prefix to required assembly names for TextExpression.ReferencesForImplementation
PREFIX_TO_ASSEMBLIES = {
    's': ['System.Private.CoreLib'],
    'sd': ['System.Data.Common', 'System.Data'],
    'scg': ['System.Private.CoreLib'],
    'sco': ['System.Private.CoreLib'],
    'ue': ['UiPath.Excel.Activities', 'UiPath.Excel'],
    'ueab': ['UiPath.Excel.Activities'],
    'ui': ['UiPath.System.Activities'],
    'uix': ['UiPath.UIAutomation.Activities'],
    'mva': ['System.Activities'],
    'sd1': ['System.Drawing.Primitives'],
    'sd2': ['System.Drawing.Common'],
    'av': ['PresentationFramework', 'PresentationCore', 'WindowsBase'],
    'x': [],
    'mc': [],
    'sap': [],
    'sap2010': [],
}

# Baseline CLR namespaces that must always appear in TextExpression.NamespacesForImplementation
# even when metadata is empty. These are required for VB expression resolution.
BASELINE_CLR_NAMESPACES = [
    'System',
    'System.Collections.Generic',
    'System.Collections.ObjectModel',
    'System.Data',
    'System.Drawing',
    'System.Linq',
    'UiPath.Core',
    'UiPath.Core.Activities',
    'UiPath.Excel',
    'UiPath.Excel.Activities',
    'UiPath.Excel.Activities.Business',
]

# Mapping of CLR namespaces to their assembly names for VisualBasicImportReference generation
CLR_NAMESPACE_TO_ASSEMBLY = {
    'Microsoft.VisualBasic': 'Microsoft.VisualBasic',
    'Microsoft.VisualBasic.Activities': 'System.Activities',
    'System': 'mscorlib',
    'System.Activities': 'System.Activities',
    'System.Activities.Expressions': 'System.Activities',
    'System.Activities.Statements': 'System.Activities',
    'System.Activities.Validation': 'System.Activities',
    'System.Activities.XamlIntegration': 'System.Activities',
    'System.Collections': 'mscorlib',
    'System.Collections.Generic': 'mscorlib',
    'System.Collections.ObjectModel': 'mscorlib',
    'System.Data': 'System.Data',
    'System.Diagnostics': 'System',
    'System.Drawing': 'System.Drawing',
    'System.IO': 'mscorlib',
    'System.Linq': 'System.Core',
    'System.Net.Mail': 'System.Net.Mail',
    'System.Windows': 'PresentationFramework',
    'System.Windows.Markup': 'PresentationFramework',
    'System.Xml': 'System.Xml',
    'System.Xml.Linq': 'System.Xml.Linq',
    'UiPath.Core': 'UiPath.System.Activities',
    'UiPath.Core.Activities': 'UiPath.System.Activities',
    'UiPath.Excel': 'UiPath.Excel.Activities',
    'UiPath.Excel.Activities': 'UiPath.Excel.Activities',
    'UiPath.Excel.Activities.Business': 'UiPath.Excel.Activities',
    'UiPath.UIAutomationNext.Activities': 'UiPath.UIAutomation.Activities',
    'UiPath.UIAutomationNext.Enums': 'UiPath.UIAutomation.Activities',
}

# Default HintSize values per activity type
DEFAULT_HINT_SIZES = {
    # Core activities
    'Sequence': '400,200',
    'Assign': '262,60',
    'If': '464,200',
    'LogMessage': '262,60',
    'InvokeWorkflowFile': '318,88',

    # Control flow
    'Switch': '497,354',
    'TryCatch': '456,713',
    'ForEach': '518,1098',
    'ForEachRow': '502,1167',
    'Rethrow': '382,48',
    'Catch': '422,528',

    # Excel activities
    'ExcelProcessScopeX': '580,1701',
    'ExcelApplicationCard': '512,1522',
    'ReadRangeX': '444,201',
    'SaveExcelFileX': '444,108',
    'WriteCellX': '444,191',
    'WriteRangeX': '444,191',
    'CopyPasteRangeX': '444,272',
    'ClearRangeX': '444,191',
    'FilterX': '444,191',
    'FindFirstLastDataRowX': '444,150',

    # File operations
    'CreateDirectory': '334,90',
    'MoveFile': '450,182',
    'ReadRange': '450,120',

    # Utilities
    'CommentOut': '580,84',
    'RetryScope': '580,1800',

    # Control flow additions
    'While': '514,707',
    'InterruptibleWhile': '660,1200',
    'Delay': '434,122',
    'Continue': '262,60',
    'Break': '262,60',
    'Throw': '382,48',
    'Return': '262,60',

    # File operations additions
    'PathExists': '450,84',
    'KillProcess': '552,156',

    # File operations additions (new)
    'DeleteFileX': '382,48',
    'ReadTextFile': '586,124',

    # Data activities
    'AddDataRow': '334,186',
    'BuildDataTable': '586,92',

    # UI Automation activities
    'NApplicationCard': '552,1503',
    'NClick': '484,189',
    'NTypeInto': '450,240',
    'NCheckState': '484,639',
    'NMouseScroll': '416,299',

    # Utility additions
    'SetToClipboard': '434,83',
    'InputDialog': '444,191',
    'InvokeCode': '434,191',

    # Flowchart activities
    'Flowchart': '614,636',
    'FlowStep': '110,70',
    'FlowDecision': '60,60',
}

# Default assembly references for NEW workflows (when metadata has no valid refs)
DEFAULT_ASSEMBLY_REFERENCES = [
    'Microsoft.CSharp',
    'Microsoft.VisualBasic',
    'mscorlib',
    'PresentationCore',
    'PresentationFramework',
    'System',
    'System.Activities',
    'System.ComponentModel.Composition',
    'System.ComponentModel.TypeConverter',
    'System.Core',
    'System.Data',
    'System.Data.Common',
    'System.Data.DataSetExtensions',
    'System.Drawing',
    'System.Drawing.Common',
    'System.Drawing.Primitives',
    'System.Linq',
    'System.Memory',
    'System.ObjectModel',
    'System.Private.CoreLib',
    'System.Private.ServiceModel',
    'System.Runtime.Serialization',
    'System.ServiceModel',
    'System.ServiceModel.Activities',
    'System.Xaml',
    'System.Xml',
    'System.Xml.Linq',
    'UiPath.Excel',
    'UiPath.Excel.Activities',
    'UiPath.Mail.Activities',
    'UiPath.System.Activities',
    'UiPath.UIAutomation.Activities',
    'WindowsBase',
]

# =============================================================================
# VB Expression and Literal Detection Patterns
# =============================================================================

# Regex patterns matching VB.NET expressions that require bracket-wrapping
VB_EXPRESSION_PATTERNS = [
    r'If\(',                          # If() function
    r'New\s+\w+',                     # New object creation
    r'(?:CType|CInt|CStr|CDate|CDbl|CBool)\(',  # Cast functions
    r'DirectCast\(',                  # DirectCast
    r'\w+\.\w+',                      # Property/method access (e.g., dt.Rows)
    r'[+\-*/&]',                      # Arithmetic/concatenation operators
    r'[<>=]',                         # Comparison operators
    r'\b(?:And|Or|Not|Mod|AndAlso|OrElse)\b',  # Logical keywords
    r'\.(?:Count|Length|Rows|Columns|ToString)\b',  # Common property suffixes
]

# Regex patterns matching literal values that must NOT be wrapped
LITERAL_PATTERNS = [
    r'^".*"$',                        # String literals
    r'^-?\d+(\.\d+)?$',              # Numeric literals
    r'^(?:True|False|Nothing)$',      # Boolean/Nothing
    r'^\[[\w_]+\]$',                  # Already-bracketed simple vars
]


# =============================================================================
# Canonicalization Context (module-level, set during parse_file)
# =============================================================================

# These are set by XamlParser.parse_file() and used by all activity handlers
# during parsing to canonicalize type strings.
_canon_xmlns_bindings: Dict[str, str] = {}
_canon_uri_to_canonical: Dict[str, str] = {}


def canonicalize_type(xaml_type: str) -> str:
    """Canonicalize a XAML type string using the current parsing context.

    This is a convenience wrapper around TypeMapper.canonicalize_type_string()
    that uses the module-level canonicalization context set by XamlParser.parse_file().

    Activity handlers should call this on any raw XAML type string before passing
    it to TypeMapper.xaml_to_json_type().
    """
    return TypeMapper.canonicalize_type_string(
        xaml_type, _canon_xmlns_bindings, _canon_uri_to_canonical)


# =============================================================================
# Utility Functions
# =============================================================================

def setup_namespaces():
    """Register namespaces with ElementTree to preserve prefixes."""
    for prefix, uri in NAMESPACES.items():
        if prefix:  # Skip empty prefix
            ET.register_namespace(prefix, uri)
    # Register default namespace
    ET.register_namespace('', NAMESPACES[''])


def get_ns_tag(prefix: str, local_name: str) -> str:
    """Create a fully qualified tag name with namespace."""
    uri = NAMESPACES.get(prefix, '')
    if uri:
        return f'{{{uri}}}{local_name}'
    return local_name


def parse_tag(tag: str) -> Tuple[str, str]:
    """Parse a tag into (namespace_uri, local_name)."""
    if tag.startswith('{'):
        ns_end = tag.index('}')
        return tag[1:ns_end], tag[ns_end + 1:]
    return '', tag


def get_activity_type(element: ET.Element) -> str:
    """Extract activity type from element tag, stripping namespace."""
    _, local_name = parse_tag(element.tag)
    return local_name


def escape_expression(expr: str) -> str:
    """DEPRECATED: Manual entity encoding causes double-encoding with ElementTree.
    ElementTree handles encoding automatically during serialization.
    This function is preserved for reference only - do not call in build methods."""
    if expr is None:
        return ''
    expr = expr.replace('&', '&amp;')
    expr = expr.replace('<', '&lt;')
    expr = expr.replace('>', '&gt;')
    expr = expr.replace('"', '&quot;')
    return expr


def unescape_expression(expr: str) -> str:
    """Convert XML entities back to plain text."""
    if expr is None:
        return ''
    expr = expr.replace('&quot;', '"')
    expr = expr.replace('&lt;', '<')
    expr = expr.replace('&gt;', '>')
    expr = expr.replace('&amp;', '&')
    return expr


# =============================================================================
# Auto-Corrector
# =============================================================================

@dataclass
class CorrectionContext:
    """Tracks corrections applied during auto-correction pass."""
    used_prefixes: set = field(default_factory=set)
    used_types: set = field(default_factory=set)
    corrections_applied: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


class WorkflowAutoCorrector:
    """Applies automatic corrections to workflow JSON before XAML generation.

    Corrects VB expression wrapping, type normalization, and argument traversal
    for both InvokeCode and InvokeWorkflowFile argument schemas.
    """

    def correct(self, workflow_json: Dict[str, Any]) -> Tuple[Dict[str, Any], CorrectionContext]:
        """Apply all auto-corrections to a workflow JSON structure.

        Args:
            workflow_json: The workflow dict (not the top-level JSON with metadata)

        Returns:
            Tuple of (corrected_workflow_copy, correction_context)
        """
        corrected = copy.deepcopy(workflow_json)
        context = CorrectionContext()
        self._correct_activity(corrected, context)
        return corrected, context

    @staticmethod
    def _is_literal(value: str) -> bool:
        """Return True if value matches any LITERAL_PATTERNS."""
        if not isinstance(value, str):
            return False
        for pattern in LITERAL_PATTERNS:
            if re.match(pattern, value):
                return True
        return False

    @classmethod
    def _is_vb_expression(cls, value: str) -> bool:
        """Return True if value looks like a VB expression needing bracket-wrapping."""
        if not isinstance(value, str) or not value:
            return False
        # Already wrapped
        if value.startswith('[') and value.endswith(']'):
            return False
        # Literals should not be wrapped
        if cls._is_literal(value):
            return False
        # Check against VB expression patterns
        for pattern in VB_EXPRESSION_PATTERNS:
            if re.search(pattern, value):
                return True
        return False

    @classmethod
    def _correct_expression_value(cls, value: str, type_args: str,
                                   context: CorrectionContext) -> str:
        """Correct a single expression value, wrapping in brackets if needed.

        Args:
            value: The expression string to check
            type_args: The XAML type hint (e.g., 'x:String', 'x:Int32')
            context: CorrectionContext to log changes
        """
        if not isinstance(value, str) or not value:
            return value
        # Already wrapped
        if value.startswith('[') and value.endswith(']'):
            return value
        # Check if it's a VB expression
        if cls._is_vb_expression(value):
            corrected = f'[{value}]'
            context.corrections_applied.append({
                'type': 'expression_wrap',
                'before': value,
                'after': corrected,
            })
            return corrected
        # Non-string safety net: if type is not x:String and value is not a literal,
        # wrap it (catches bare variable references like 'myVar' in Int32 fields)
        if type_args and type_args != 'x:String' and not cls._is_literal(value):
            corrected = f'[{value}]'
            context.corrections_applied.append({
                'type': 'safety_net_wrap',
                'before': value,
                'after': corrected,
                'type_hint': type_args,
            })
            return corrected
        return value

    @staticmethod
    def _correct_type_reference(type_str: str, context: CorrectionContext) -> str:
        """Correct a type reference using TypeMapper.normalize_type_reference().

        Args:
            type_str: The type string to normalize
            context: CorrectionContext to log changes
        """
        if not isinstance(type_str, str) or not type_str:
            return type_str
        result = TypeMapper.normalize_type_reference(type_str, context)
        if result != type_str:
            context.corrections_applied.append({
                'type': 'type_normalize',
                'before': type_str,
                'after': result,
            })
        # Fallback tracking: ensure all type references are recorded,
        # even when unchanged (e.g., simple short names like "String")
        context.used_types.add(result)
        if ':' in result:
            context.used_prefixes.add(result.split(':', 1)[0])
        return result

    @classmethod
    def _correct_activity(cls, activity: Dict[str, Any], context: CorrectionContext):
        """Recursively correct a single activity dict in-place.

        Traversal paths are schema-accurate, matching the exact JSON keys
        consumed by each handler (FlowchartHandler, TryCatchHandler, etc.).
        """
        if not isinstance(activity, dict):
            return

        # 1. Standard expression fields
        for expr_key in ('value', 'condition', 'expression'):
            if expr_key in activity:
                val = activity[expr_key]
                if isinstance(val, str):
                    activity[expr_key] = cls._correct_expression_value(
                        val, activity.get('x:TypeArguments', 'x:String'), context)
                elif isinstance(val, dict) and 'value' in val:
                    type_hint = val.get('type', 'x:String')
                    val['value'] = cls._correct_expression_value(
                        val['value'], type_hint, context)

        # Handle 'to' field (Assign target)
        if 'to' in activity:
            val = activity['to']
            if isinstance(val, str):
                activity['to'] = cls._correct_expression_value(
                    val, activity.get('x:TypeArguments', 'x:String'), context)
            elif isinstance(val, dict) and 'value' in val:
                type_hint = val.get('type', 'x:String')
                val['value'] = cls._correct_expression_value(
                    val['value'], type_hint, context)

        # 2. Type fields on this activity dict
        for type_key in ('type', 'x:TypeArguments', 'variableType', 'argumentType',
                         'exceptionType', 'typeArguments', 'typeArgument'):
            if type_key in activity and isinstance(activity[type_key], str):
                activity[type_key] = cls._correct_type_reference(
                    activity[type_key], context)

        # 3. Variables list
        for var in activity.get('variables', []):
            if isinstance(var, dict):
                if 'type' in var and isinstance(var['type'], str):
                    var['type'] = cls._correct_type_reference(var['type'], context)
                if 'default' in var and isinstance(var['default'], str):
                    type_hint = TypeMapper.json_to_xaml_type(var.get('type', 'String'))
                    var['default'] = cls._correct_expression_value(
                        var['default'], type_hint, context)

        # 4. Schema-accurate child traversal
        # -- Sequence / generic containers --
        for child in activity.get('children', []):
            if isinstance(child, dict):
                cls._correct_activity(child, context)

        # -- Flowchart: nodes[] --
        for node in activity.get('nodes', []):
            if isinstance(node, dict):
                cls._correct_activity(node, context)

        # -- FlowStep: activity (inline dict), next (string ref or inline dict) --
        # These keys are handled at the general level so inline-nested FlowSteps
        # (e.g., inside FlowDecision true/false) are also traversed.
        if isinstance(activity.get('activity'), dict):
            cls._correct_activity(activity['activity'], context)
        if isinstance(activity.get('next'), dict):
            cls._correct_activity(activity['next'], context)

        # -- FlowDecision: true/false (string ref or inline dict) --
        if isinstance(activity.get('true'), dict):
            cls._correct_activity(activity['true'], context)
        if isinstance(activity.get('false'), dict):
            cls._correct_activity(activity['false'], context)

        # -- If: then/else (activity dicts) --
        if isinstance(activity.get('then'), dict):
            cls._correct_activity(activity['then'], context)
        if isinstance(activity.get('else'), dict):
            cls._correct_activity(activity['else'], context)

        # -- TryCatch: try (activity dict), catches[], finally (activity dict) --
        if isinstance(activity.get('try'), dict):
            cls._correct_activity(activity['try'], context)
        if isinstance(activity.get('finally'), dict):
            cls._correct_activity(activity['finally'], context)
        for catch in activity.get('catches', []):
            if not isinstance(catch, dict):
                continue
            # Normalize exceptionType on the catch dict itself
            if isinstance(catch.get('exceptionType'), str):
                catch['exceptionType'] = cls._correct_type_reference(
                    catch['exceptionType'], context)
            # handler is an activity dict
            if isinstance(catch.get('handler'), dict):
                cls._correct_activity(catch['handler'], context)

        # -- RetryScope: activityBody (activity dict), condition.activity --
        if isinstance(activity.get('activityBody'), dict):
            cls._correct_activity(activity['activityBody'], context)
        cond = activity.get('condition')
        if isinstance(cond, dict) and isinstance(cond.get('activity'), dict):
            cls._correct_activity(cond['activity'], context)

        # -- ActivityAction wrappers (ForEach, ForEachRow, InterruptibleWhile): body.activity --
        body = activity.get('body')
        if isinstance(body, dict):
            if isinstance(body.get('activity'), dict):
                cls._correct_activity(body['activity'], context)
            elif 'activity' not in body:
                # While handler uses body as a direct activity dict (not ActivityAction)
                cls._correct_activity(body, context)

        # -- Switch: cases[].activity, default (activity dict) --
        for case in activity.get('cases', []):
            if isinstance(case, dict) and isinstance(case.get('activity'), dict):
                cls._correct_activity(case['activity'], context)
        if isinstance(activity.get('default'), dict):
            cls._correct_activity(activity['default'], context)

        # -- Scope containers: ifExists / ifNotExists --
        if isinstance(activity.get('ifExists'), dict):
            cls._correct_activity(activity['ifExists'], context)
        if isinstance(activity.get('ifNotExists'), dict):
            cls._correct_activity(activity['ifNotExists'], context)

        # 5. Arguments list (Gap A - InvokeCode and InvokeWorkflowFile schemas)
        for arg in activity.get('arguments', []):
            if not isinstance(arg, dict):
                continue

            # Determine XAML type hint based on argument schema
            if 'x:TypeArguments' in arg:
                # InvokeCode schema: {direction, x:TypeArguments, x:Key, value}
                xaml_type_hint = arg.get('x:TypeArguments', 'x:String') or 'x:String'
            else:
                # InvokeWorkflowFile schema: {key, direction, type, value}
                json_type = arg.get('type', 'String') or 'String'
                xaml_type_hint = TypeMapper.json_to_xaml_type(json_type)

            # Correct expression value
            if 'value' in arg and isinstance(arg['value'], str):
                arg['value'] = cls._correct_expression_value(
                    arg['value'], xaml_type_hint, context)

            # Correct type references
            if 'x:TypeArguments' in arg and isinstance(arg['x:TypeArguments'], str):
                arg['x:TypeArguments'] = cls._correct_type_reference(
                    arg['x:TypeArguments'], context)
            if 'type' in arg and isinstance(arg['type'], str):
                arg['type'] = cls._correct_type_reference(
                    arg['type'], context)


# =============================================================================
# Type Mapper
# =============================================================================

class TypeMapper:
    """Utility class for mapping between JSON type strings and XAML x:TypeArguments format."""

    # Map from simple and fully-qualified type names to XAML prefixed format.
    # Fully-qualified entries (System.*) enable canonical-first lookup in
    # normalize_type_reference before falling back to namespace resolution.
    TYPE_MAP = {
        # Simple names (used by json_to_xaml_type)
        'String': 'x:String',
        'Int32': 'x:Int32',
        'Int64': 'x:Int64',
        'Boolean': 'x:Boolean',
        'Double': 'x:Double',
        'Decimal': 'x:Decimal',
        'DateTime': 's:DateTime',
        'TimeSpan': 's:TimeSpan',
        'Object': 'x:Object',
        'DataTable': 'sd:DataTable',
        'DataRow': 'sd:DataRow',
        'Exception': 's:Exception',
        # Fully-qualified XAML language primitives (x:* namespace)
        'System.String': 'x:String',
        'System.Int32': 'x:Int32',
        'System.Int64': 'x:Int64',
        'System.Boolean': 'x:Boolean',
        'System.Double': 'x:Double',
        'System.Decimal': 'x:Decimal',
        'System.Object': 'x:Object',
        # Fully-qualified System types (s:* namespace)
        'System.DateTime': 's:DateTime',
        'System.TimeSpan': 's:TimeSpan',
        'System.Exception': 's:Exception',
        # Fully-qualified System.Data types (sd:* namespace)
        'System.Data.DataTable': 'sd:DataTable',
        'System.Data.DataRow': 'sd:DataRow',
    }

    # Reverse map for XAML to JSON — two-pass build ensures short names win
    REVERSE_TYPE_MAP = {v: k for k, v in TYPE_MAP.items()}
    # Pass 2: overwrite with short-name entries (keys without '.') so they always win
    REVERSE_TYPE_MAP.update({v: k for k, v in TYPE_MAP.items() if '.' not in k})

    @classmethod
    def json_to_xaml_type(cls, json_type: str) -> str:
        """Convert JSON type string to XAML x:TypeArguments format."""
        # Check for generic types like List<String>
        generic_match = re.match(r'(\w+)<(.+)>', json_type)
        if generic_match:
            container = generic_match.group(1)
            inner_type = generic_match.group(2)
            inner_xaml = cls.json_to_xaml_type(inner_type)
            if container == 'List':
                return f'scg:List({inner_xaml})'
            elif container == 'Dictionary':
                # Handle Dictionary<K,V>
                parts = inner_type.split(',')
                if len(parts) == 2:
                    key_type = cls.json_to_xaml_type(parts[0].strip())
                    val_type = cls.json_to_xaml_type(parts[1].strip())
                    return f'scg:Dictionary({key_type}, {val_type})'

        return cls.TYPE_MAP.get(json_type, json_type)

    @classmethod
    def xaml_to_json_type(cls, xaml_type: str) -> str:
        """Convert XAML x:TypeArguments format to JSON type string."""
        # Check for generic types like scg:List(x:String)
        generic_match = re.match(r'scg:(\w+)\((.+)\)', xaml_type)
        if generic_match:
            container = generic_match.group(1)
            inner_xaml = generic_match.group(2)
            inner_json = cls.xaml_to_json_type(inner_xaml)
            return f'{container}<{inner_json}>'

        return cls.REVERSE_TYPE_MAP.get(xaml_type, xaml_type)

    @classmethod
    def canonicalize_type_string(cls, xaml_type: str, xmlns_bindings: Dict[str, str],
                                  uri_to_canonical: Dict[str, str]) -> str:
        """Canonicalize a prefixed XAML type string using URI-based lookup.

        Translates document-specific prefixes to the framework's canonical prefixes
        by looking up the prefix's URI and finding the canonical prefix for that URI.

        Args:
            xaml_type: The XAML type string (e.g., 'sd:Image', 'scg:List(sd:Image)')
            xmlns_bindings: Document's prefix->URI mapping
            uri_to_canonical: URI->canonical_prefix mapping from NAMESPACES registry

        Returns:
            Canonicalized type string (e.g., 'sd2:Image' if sd maps to System.Drawing.Common)
        """
        if not xaml_type or not xmlns_bindings or not uri_to_canonical:
            return xaml_type

        # Handle generic types with prefixed container like scg:List(sd:Image) or scg:Dictionary(x:String, sd:Image)
        generic_match = re.match(r'(\w+:\w+)\((.+)\)', xaml_type)
        if generic_match:
            container = generic_match.group(1)
            inner = generic_match.group(2)
            # Canonicalize the container prefix
            canon_container = cls._canonicalize_single_prefix(container, xmlns_bindings, uri_to_canonical)
            # Canonicalize inner type(s) - may be comma-separated for Dictionary
            parts = cls._split_type_args(inner)
            canon_parts = [cls.canonicalize_type_string(p.strip(), xmlns_bindings, uri_to_canonical) for p in parts]
            return f'{canon_container}({", ".join(canon_parts)})'

        # Handle unprefixed generic containers and argument wrappers like
        # InArgument(sd:Image), OutArgument(scg:List(sd:Image)), Dictionary(x:String, sd:Image)
        unprefixed_generic_match = re.match(r'(\w+)\((.+)\)', xaml_type)
        if unprefixed_generic_match:
            wrapper = unprefixed_generic_match.group(1)
            inner = unprefixed_generic_match.group(2)
            # Canonicalize inner type(s) - may be comma-separated
            parts = cls._split_type_args(inner)
            canon_parts = [cls.canonicalize_type_string(p.strip(), xmlns_bindings, uri_to_canonical) for p in parts]
            return f'{wrapper}({", ".join(canon_parts)})'

        # Handle comma-separated types (e.g., 'x:String, x:Object')
        if ',' in xaml_type and '(' not in xaml_type:
            parts = xaml_type.split(',')
            canon_parts = [cls.canonicalize_type_string(p.strip(), xmlns_bindings, uri_to_canonical) for p in parts]
            return ', '.join(canon_parts)

        # Simple prefixed type like sd:Image
        return cls._canonicalize_single_prefix(xaml_type, xmlns_bindings, uri_to_canonical)

    @classmethod
    def _canonicalize_single_prefix(cls, prefixed_type: str, xmlns_bindings: Dict[str, str],
                                     uri_to_canonical: Dict[str, str]) -> str:
        """Canonicalize a single prefixed type (no generics)."""
        if ':' not in prefixed_type:
            return prefixed_type

        prefix, local_name = prefixed_type.split(':', 1)
        uri = xmlns_bindings.get(prefix)
        if uri is None:
            return prefixed_type  # Unknown prefix, preserve as-is

        canonical_prefix = uri_to_canonical.get(uri)
        if canonical_prefix is None:
            return prefixed_type  # URI not in canonical registry, preserve as-is

        if canonical_prefix == prefix:
            return prefixed_type  # Already canonical

        return f'{canonical_prefix}:{local_name}'

    @staticmethod
    def _split_type_args(inner: str) -> List[str]:
        """Split type arguments respecting nested parentheses.

        E.g., 'scg:List(x:String), x:Object' -> ['scg:List(x:String)', 'x:Object']
        """
        parts = []
        depth = 0
        current = []
        for ch in inner:
            if ch == '(':
                depth += 1
                current.append(ch)
            elif ch == ')':
                depth -= 1
                current.append(ch)
            elif ch == ',' and depth == 0:
                parts.append(''.join(current))
                current = []
            else:
                current.append(ch)
        if current:
            parts.append(''.join(current))
        return parts

    @classmethod
    def normalize_type_reference(cls, type_str: str, context: 'CorrectionContext' = None) -> str:
        """Convert a fully-qualified type string to namespace-prefixed form.

        Uses canonical-first strategy with this precedence:
        1. Already prefixed (contains ':') -> return as-is
        2. Simple short name (no '.' and no ':') -> return as-is
        3. Canonical TYPE_MAP hit (fully-qualified names) -> return mapped value
        4. Fully-qualified non-primitive (contains '.') -> reverse-lookup namespace
           in PREFIX_TO_CLR_NAMESPACES; if exactly one prefix matches -> prefix:ShortName
        5. Ambiguous (multiple prefixes) -> return original, warn
        6. Unmapped (no prefix found) -> return original, warn

        Args:
            type_str: The type string to normalize
            context: Optional CorrectionContext to record ambiguity/unmapped warnings
        """
        if not type_str:
            return type_str

        # 1. Already prefixed
        if ':' in type_str:
            if context is not None:
                prefix = type_str.split(':')[0]
                context.used_prefixes.add(prefix)
                context.used_types.add(type_str)
            return type_str

        # 2. Simple short name (no dots, no colon) - pass through unchanged
        if '.' not in type_str:
            return type_str

        # 3. Canonical TYPE_MAP hit (fully-qualified names like 'System.String')
        if type_str in cls.TYPE_MAP:
            result = cls.TYPE_MAP[type_str]
            if context is not None and ':' in result:
                context.used_prefixes.add(result.split(':')[0])
                context.used_types.add(result)
            return result

        # 4-6. Fully-qualified non-primitive: reverse-lookup namespace
        last_dot = type_str.rfind('.')
        namespace = type_str[:last_dot]
        type_name = type_str[last_dot + 1:]

        # Find which prefix maps to this namespace
        matching_prefixes = []
        for prefix, clr_namespaces in PREFIX_TO_CLR_NAMESPACES.items():
            if namespace in clr_namespaces:
                matching_prefixes.append(prefix)

        # 4. Exactly one match -> use it
        if len(matching_prefixes) == 1:
            result = f'{matching_prefixes[0]}:{type_name}'
            if context is not None:
                context.used_prefixes.add(matching_prefixes[0])
                context.used_types.add(result)
            return result

        # 5. Ambiguous -> return original, record warning
        if len(matching_prefixes) > 1:
            warning = f"Ambiguous type: {type_str} matches prefixes {sorted(matching_prefixes)}"
            if context is not None:
                context.warnings.append(warning)
            return type_str

        # 6. Unmapped -> return original, record warning
        warning = f"Unmapped type: {type_str} has no matching namespace prefix"
        if context is not None:
            context.warnings.append(warning)
        return type_str


# =============================================================================
# IdRef Generator
# =============================================================================

class IdRefGenerator:
    """Generates unique IdRef values for activities."""

    def __init__(self):
        self._counters: Dict[str, int] = {}

    def generate(self, activity_type: str) -> str:
        """Generate a unique IdRef for the given activity type."""
        if activity_type not in self._counters:
            self._counters[activity_type] = 0
        self._counters[activity_type] += 1
        return f'{activity_type}_{self._counters[activity_type]}'

    def reset(self):
        """Reset all counters."""
        self._counters.clear()


# =============================================================================
# ViewState Builder
# =============================================================================

class ViewStateBuilder:
    """Builds ViewState dictionaries for activities."""

    @staticmethod
    def create_viewstate_element(viewstate_dict: Dict[str, Any]) -> ET.Element:
        """Create a ViewState dictionary element."""
        # Create the WorkflowViewStateService.ViewState wrapper
        viewstate_tag = get_ns_tag('sap', 'WorkflowViewStateService.ViewState')
        viewstate_elem = ET.Element(viewstate_tag)

        # Create the Dictionary element
        dict_tag = get_ns_tag('scg', 'Dictionary')
        dict_elem = ET.SubElement(viewstate_elem, dict_tag)
        dict_elem.set(get_ns_tag('x', 'TypeArguments'), 'x:String, x:Object')

        # Add entries
        for key, value in viewstate_dict.items():
            if key == 'IsExpanded' and isinstance(value, bool):
                bool_tag = get_ns_tag('x', 'Boolean')
                bool_elem = ET.SubElement(dict_elem, bool_tag)
                bool_elem.set(get_ns_tag('x', 'Key'), key)
                bool_elem.text = str(value).lower()
            elif key == 'IsPinned' and isinstance(value, bool):
                bool_tag = get_ns_tag('x', 'Boolean')
                bool_elem = ET.SubElement(dict_elem, bool_tag)
                bool_elem.set(get_ns_tag('x', 'Key'), key)
                bool_elem.text = str(value).lower()

        return viewstate_elem

    @staticmethod
    def parse_viewstate(element: ET.Element) -> Dict[str, Any]:
        """Parse ViewState from an element."""
        result = {}

        # Find the ViewState element
        viewstate_tag = get_ns_tag('sap', 'WorkflowViewStateService.ViewState')
        viewstate_elem = element.find(viewstate_tag)

        if viewstate_elem is None:
            return result

        # Find the Dictionary element
        dict_tag = get_ns_tag('scg', 'Dictionary')
        dict_elem = viewstate_elem.find(dict_tag)

        if dict_elem is None:
            return result

        # Parse entries
        for child in dict_elem:
            key_attr = get_ns_tag('x', 'Key')
            key = child.get(key_attr)
            if key:
                _, local = parse_tag(child.tag)
                if local == 'Boolean':
                    result[key] = child.text.lower() == 'true' if child.text else False
                else:
                    result[key] = child.text

        return result

    @staticmethod
    def create_flowchart_viewstate(viewstate_dict: Dict[str, Any]) -> ET.Element:
        """Create a ViewState element with flowchart-specific entries (ShapeLocation, ShapeSize, connectors)."""
        viewstate_tag = get_ns_tag('sap', 'WorkflowViewStateService.ViewState')
        viewstate_elem = ET.Element(viewstate_tag)

        dict_tag = get_ns_tag('scg', 'Dictionary')
        dict_elem = ET.SubElement(viewstate_elem, dict_tag)
        dict_elem.set(get_ns_tag('x', 'TypeArguments'), 'x:String, x:Object')

        key_attr = get_ns_tag('x', 'Key')

        for key, value in viewstate_dict.items():
            if key in ('IsExpanded', 'IsPinned') and isinstance(value, bool):
                bool_tag = get_ns_tag('x', 'Boolean')
                bool_elem = ET.SubElement(dict_elem, bool_tag)
                bool_elem.set(key_attr, key)
                bool_elem.text = str(value).lower()
            elif key == 'ShapeLocation':
                point_tag = get_ns_tag('av', 'Point')
                point_elem = ET.SubElement(dict_elem, point_tag)
                point_elem.set(key_attr, key)
                point_elem.text = str(value)
            elif key == 'ShapeSize':
                size_tag = get_ns_tag('av', 'Size')
                size_elem = ET.SubElement(dict_elem, size_tag)
                size_elem.set(key_attr, key)
                size_elem.text = str(value)
            elif key in ('ConnectorLocation', 'TrueConnector', 'FalseConnector'):
                pc_tag = get_ns_tag('av', 'PointCollection')
                pc_elem = ET.SubElement(dict_elem, pc_tag)
                pc_elem.set(key_attr, key)
                pc_elem.text = str(value)

        return viewstate_elem

    @staticmethod
    def parse_flowchart_viewstate(element: ET.Element) -> Dict[str, Any]:
        """Parse flowchart-specific ViewState from an element (ShapeLocation, ShapeSize, connectors)."""
        result = {}

        viewstate_tag = get_ns_tag('sap', 'WorkflowViewStateService.ViewState')
        viewstate_elem = element.find(viewstate_tag)

        if viewstate_elem is None:
            return result

        dict_tag = get_ns_tag('scg', 'Dictionary')
        dict_elem = viewstate_elem.find(dict_tag)

        if dict_elem is None:
            return result

        for child in dict_elem:
            key_attr_name = get_ns_tag('x', 'Key')
            key = child.get(key_attr_name)
            if key:
                _, local = parse_tag(child.tag)
                if local == 'Boolean':
                    result[key] = child.text.lower() == 'true' if child.text else False
                else:
                    # Point, Size, PointCollection all stored as string
                    result[key] = child.text if child.text else ''

        return result

    @staticmethod
    def viewstate_to_xaml(viewstate_dict: Dict[str, Any], is_flowchart: bool = False) -> ET.Element:
        """Single entrypoint for converting a viewstate dictionary to a XAML element.

        Delegates to create_flowchart_viewstate() when is_flowchart is True or
        when flowchart-specific keys (ShapeLocation, ShapeSize, or connector
        keys) are detected in the dictionary.  Falls back to
        create_viewstate_element() otherwise.
        """
        flowchart_keys = {'ShapeLocation', 'ShapeSize', 'ConnectorLocation',
                          'TrueConnector', 'FalseConnector'}
        if is_flowchart or (viewstate_dict and flowchart_keys & viewstate_dict.keys()):
            return ViewStateBuilder.create_flowchart_viewstate(viewstate_dict)
        return ViewStateBuilder.create_viewstate_element(viewstate_dict)


# =============================================================================
# Metadata Manager
# =============================================================================

class MetadataManager:
    """Handles XAML metadata extraction and application."""

    @staticmethod
    def extract_xmlns_bindings(root: ET.Element) -> Dict[str, str]:
        """Extract all xmlns prefix->URI bindings from the root element.

        Note: ElementTree strips xmlns declarations from element.attrib during parsing,
        so this method checks both the standard attrib approach and falls back
        to searching for namespace URIs in the element's tag and children.

        For reliable extraction from files, use extract_xmlns_bindings_from_file() instead.
        """
        bindings = {}
        for attr_name, attr_value in root.attrib.items():
            if attr_name.startswith('{'):
                ns_uri, local = attr_name[1:].split('}', 1)
                if ns_uri == 'http://www.w3.org/2000/xmlns/':
                    bindings[local] = attr_value
            elif attr_name == 'xmlns':
                bindings[''] = attr_value
            elif attr_name.startswith('xmlns:'):
                prefix = attr_name[6:]
                bindings[prefix] = attr_value
        return bindings

    @staticmethod
    def extract_xmlns_bindings_from_file(filepath: str) -> Dict[str, str]:
        """Extract all xmlns prefix->URI bindings from a XAML file using iterparse.

        Uses ET.iterparse with 'start-ns' events to capture namespace declarations
        before they are consumed by the parser. This is the reliable way to get
        document-level xmlns bindings.

        Args:
            filepath: Path to the XAML file

        Returns:
            Dictionary mapping prefix to URI for all namespace declarations in the file
        """
        bindings = {}
        for event, data in ET.iterparse(filepath, events=['start-ns']):
            prefix, uri = data
            bindings[prefix] = uri
        return bindings

    @staticmethod
    def build_uri_to_canonical_prefix(xmlns_bindings: Dict[str, str]) -> Dict[str, str]:
        """Build a URI->canonical_prefix mapping using the framework's NS_URI_TO_PREFIX registry.

        For each URI in the document's xmlns bindings, look up the canonical prefix
        from NS_URI_TO_PREFIX. This enables translating document prefixes to canonical
        prefixes via URI lookup.

        Args:
            xmlns_bindings: Document's prefix->URI mapping from extract_xmlns_bindings()

        Returns:
            Dictionary mapping URI->canonical_prefix for all URIs found in the registry
        """
        uri_to_canonical = {}
        for uri in xmlns_bindings.values():
            canonical = NS_URI_TO_PREFIX.get(uri)
            if canonical is not None:
                uri_to_canonical[uri] = canonical
        return uri_to_canonical

    def extract_metadata(self, root: ET.Element, xmlns_bindings: Optional[Dict[str, str]] = None,
                         uri_to_canonical: Optional[Dict[str, str]] = None) -> Dict[str, Any]:
        """Extract metadata from the root Activity element."""
        metadata = {
            'class': '',
            'namespaces': [],
            'assemblyReferences': [],
            'arguments': [],
            'xmlnsBindings': {},
        }

        # Store canonicalization mappings for use in _parse_argument
        self._xmlns_bindings = xmlns_bindings or {}
        self._uri_to_canonical = uri_to_canonical or {}

        # Extract x:Class attribute
        class_attr = get_ns_tag('x', 'Class')
        if class_attr in root.attrib:
            metadata['class'] = root.get(class_attr)

        # Extract namespaces from TextExpression.NamespacesForImplementation
        namespaces_tag = get_ns_tag('', 'TextExpression.NamespacesForImplementation')
        namespaces_elem = root.find(namespaces_tag)
        if namespaces_elem is not None:
            # Traverse into Collection wrapper to get actual x:String elements
            for child in namespaces_elem:
                # child is the Collection element; iterate its children
                for ns_elem in child:
                    if ns_elem.text and ns_elem.text.strip():
                        metadata['namespaces'].append(ns_elem.text.strip())
                # If Collection itself has text (shouldn't normally), skip it
            # Fallback: if no namespaces found via Collection, try direct children
            if not metadata['namespaces']:
                for ns_elem in namespaces_elem:
                    if ns_elem.text and ns_elem.text.strip():
                        metadata['namespaces'].append(ns_elem.text.strip())

        # Extract assembly references from TextExpression.ReferencesForImplementation
        refs_tag = get_ns_tag('', 'TextExpression.ReferencesForImplementation')
        refs_elem = root.find(refs_tag)
        if refs_elem is not None:
            # Traverse into Collection wrapper to get actual AssemblyReference elements
            for child in refs_elem:
                # child is the Collection element; iterate its children
                for ref_elem in child:
                    # Get the text content (assembly name or full reference)
                    if ref_elem.text and ref_elem.text.strip():
                        metadata['assemblyReferences'].append(ref_elem.text.strip())
                    else:
                        # Check for Assembly attribute
                        assembly_attr = ref_elem.get('Assembly')
                        if assembly_attr and assembly_attr.strip():
                            metadata['assemblyReferences'].append(assembly_attr.strip())
            # Fallback: if no refs found via Collection, try direct children
            if not metadata['assemblyReferences']:
                for ref_elem in refs_elem:
                    if ref_elem.text and ref_elem.text.strip():
                        metadata['assemblyReferences'].append(ref_elem.text.strip())
                    else:
                        assembly_attr = ref_elem.get('Assembly')
                        if assembly_attr and assembly_attr.strip():
                            metadata['assemblyReferences'].append(assembly_attr.strip())

        # Extract arguments from x:Members
        members_tag = get_ns_tag('x', 'Members')
        members_elem = root.find(members_tag)
        if members_elem is not None:
            for prop_elem in members_elem:
                _, local = parse_tag(prop_elem.tag)
                if local == 'Property':
                    arg = self._parse_argument(prop_elem)
                    if arg:
                        metadata['arguments'].append(arg)

        # Extract custom xmlns bindings (prefixes not in canonical NAMESPACES)
        effective_bindings = xmlns_bindings or {}
        for prefix, uri in effective_bindings.items():
            if prefix and prefix not in NAMESPACES:
                metadata['xmlnsBindings'][prefix] = uri

        return metadata

    def _parse_argument(self, prop_elem: ET.Element) -> Optional[Dict[str, str]]:
        """Parse a single argument from x:Property element."""
        name = prop_elem.get('Name')
        type_str = prop_elem.get('Type')

        if not name or not type_str:
            return None

        # Parse type to determine direction
        # InArgument(x:String) -> direction=In, type=String
        # OutArgument(x:String) -> direction=Out, type=String
        # InOutArgument(x:String) -> direction=InOut, type=String
        direction = 'In'
        inner_type = type_str

        if 'InOutArgument' in type_str:
            direction = 'InOut'
            match = re.search(r'InOutArgument\((.+)\)', type_str)
            if match:
                inner_type = match.group(1)
        elif 'OutArgument' in type_str:
            direction = 'Out'
            match = re.search(r'OutArgument\((.+)\)', type_str)
            if match:
                inner_type = match.group(1)
        elif 'InArgument' in type_str:
            direction = 'In'
            match = re.search(r'InArgument\((.+)\)', type_str)
            if match:
                inner_type = match.group(1)

        # Canonicalize inner type before converting to JSON type
        inner_type = TypeMapper.canonicalize_type_string(
            inner_type, self._xmlns_bindings, self._uri_to_canonical)

        # Convert XAML type to JSON type
        json_type = TypeMapper.xaml_to_json_type(inner_type)

        return {
            'name': name,
            'direction': direction,
            'type': json_type,
        }

    def create_root_element(self, metadata: Dict[str, Any]) -> ET.Element:
        """Create root Activity element with all xmlns declarations."""
        # Create Activity element with default namespace
        root = ET.Element(get_ns_tag('', 'Activity'))

        # Add mc:Ignorable attribute
        root.set(get_ns_tag('mc', 'Ignorable'), 'sap sap2010')

        # Add x:Class attribute
        if metadata.get('class'):
            root.set(get_ns_tag('x', 'Class'), metadata['class'])

        return root

    def apply_namespaces(self, root: ET.Element, namespaces: List[str]):
        """Add TextExpression.NamespacesForImplementation section."""
        if not namespaces:
            return

        # Create the container element
        ns_impl_tag = get_ns_tag('', 'TextExpression.NamespacesForImplementation')
        ns_impl_elem = ET.SubElement(root, ns_impl_tag)

        # Create the collection
        coll_tag = get_ns_tag('sco', 'Collection')
        coll_elem = ET.SubElement(ns_impl_elem, coll_tag)
        coll_elem.set(get_ns_tag('x', 'TypeArguments'), 'x:String')

        # Add each namespace
        for ns in namespaces:
            str_tag = get_ns_tag('x', 'String')
            str_elem = ET.SubElement(coll_elem, str_tag)
            str_elem.text = ns

    def apply_assembly_refs(self, root: ET.Element, refs: List[str]):
        """Add TextExpression.ReferencesForImplementation section."""
        if not refs:
            return

        # Create the container element
        refs_impl_tag = get_ns_tag('', 'TextExpression.ReferencesForImplementation')
        refs_impl_elem = ET.SubElement(root, refs_impl_tag)

        # Create the collection
        coll_tag = get_ns_tag('sco', 'Collection')
        coll_elem = ET.SubElement(refs_impl_elem, coll_tag)
        coll_elem.set(get_ns_tag('x', 'TypeArguments'), 'AssemblyReference')

        # Add each reference
        for ref in refs:
            ref_tag = get_ns_tag('', 'AssemblyReference')
            ref_elem = ET.SubElement(coll_elem, ref_tag)
            ref_elem.text = ref

    def apply_arguments(self, root: ET.Element, arguments: List[Dict[str, str]]):
        """Create x:Members section for workflow arguments."""
        if not arguments:
            return

        # Create x:Members element
        members_tag = get_ns_tag('x', 'Members')
        members_elem = ET.SubElement(root, members_tag)

        # Add each argument as x:Property
        for arg in arguments:
            prop_tag = get_ns_tag('x', 'Property')
            prop_elem = ET.SubElement(members_elem, prop_tag)
            prop_elem.set('Name', arg['name'])

            # Build type string based on direction
            xaml_type = TypeMapper.json_to_xaml_type(arg['type'])
            direction = arg.get('direction', 'In')

            if direction == 'Out':
                type_str = f'OutArgument({xaml_type})'
            elif direction == 'InOut':
                type_str = f'InOutArgument({xaml_type})'
            else:
                type_str = f'InArgument({xaml_type})'

            prop_elem.set('Type', type_str)

    @staticmethod
    def detect_required_namespaces(workflow_json: Dict[str, Any]) -> set:
        """Recursively traverse workflow JSON to auto-detect required namespace prefixes.

        Scans variable types, activity TypeArguments, expression types, and property
        values for namespace-prefixed type references (e.g., 'sd:DataTable' -> 'sd').
        """
        prefixes = set()
        prefix_pattern = re.compile(r'(\w+):[\w\[\]]+')

        def scan_value(value: Any):
            """Extract namespace prefixes from a string value.

            Also resolves unprefixed type names (e.g., 'DataTable') through
            TypeMapper.TYPE_MAP to include the required namespace prefix.
            """
            if isinstance(value, str):
                for match in prefix_pattern.finditer(value):
                    candidate = match.group(1)
                    if candidate in NAMESPACES:
                        prefixes.add(candidate)
                # Also check for unprefixed types that map to prefixed XAML types
                stripped = value.strip()
                if stripped in TypeMapper.TYPE_MAP:
                    xaml_type = TypeMapper.TYPE_MAP[stripped]
                    if ':' in xaml_type:
                        prefix = xaml_type.split(':')[0]
                        if prefix in NAMESPACES:
                            prefixes.add(prefix)

        def scan_dict(d: Dict[str, Any]):
            """Recursively scan a dict for namespace-prefixed type strings."""
            for key, value in d.items():
                if key in ('type', 'typeArgument', 'TypeArguments', 'x:TypeArguments',
                           'argumentType', 'variableType', 'exceptionType'):
                    scan_value(value)
                elif key == 'variables' and isinstance(value, list):
                    for var in value:
                        if isinstance(var, dict):
                            scan_value(var.get('type', ''))
                            scan_value(var.get('default', ''))
                elif key in ('children', 'activities', 'body', 'then', 'else',
                             'catches', 'finally', 'cases', 'default',
                             'trueBody', 'falseBody', 'nodes', 'ifExists',
                             'ifNotExists'):
                    if isinstance(value, list):
                        for item in value:
                            if isinstance(item, dict):
                                scan_dict(item)
                    elif isinstance(value, dict):
                        scan_dict(value)
                elif key in ('to', 'value', 'condition', 'expression'):
                    if isinstance(value, dict):
                        scan_value(value.get('type', ''))
                        scan_dict(value)
                    elif isinstance(value, str):
                        scan_value(value)
                elif isinstance(value, dict):
                    scan_dict(value)
                elif isinstance(value, str):
                    scan_value(value)
                elif isinstance(value, list):
                    for item in value:
                        if isinstance(item, dict):
                            scan_dict(item)
                        elif isinstance(item, str):
                            scan_value(item)

        if isinstance(workflow_json, dict):
            scan_dict(workflow_json)

        return prefixes

    @staticmethod
    def generate_namespace_strings(prefixes: set, existing_namespaces: List[str]) -> List[str]:
        """Generate CLR namespace strings for TextExpression.NamespacesForImplementation.

        Combines auto-detected CLR namespaces from prefixes with valid existing
        namespace strings from metadata (filtering out whitespace-only entries).
        When metadata namespaces are empty, seeds baseline CLR namespaces to
        ensure VB expression resolution works.
        """
        clr_namespaces = set()

        # Add CLR namespaces from detected prefixes
        for prefix in prefixes:
            if prefix in PREFIX_TO_CLR_NAMESPACES:
                for ns in PREFIX_TO_CLR_NAMESPACES[prefix]:
                    clr_namespaces.add(ns)

        # Preserve valid existing namespace strings from metadata
        valid_metadata = [ns.strip() for ns in existing_namespaces
                          if isinstance(ns, str) and ns.strip()]
        for ns in valid_metadata:
            clr_namespaces.add(ns)

        # Seed baseline CLR namespaces when metadata is empty to ensure
        # VB expression resolution works for common types
        if not valid_metadata:
            for ns in BASELINE_CLR_NAMESPACES:
                clr_namespaces.add(ns)

        return sorted(clr_namespaces)

    @staticmethod
    def detect_all_used_prefixes(workflow_json: Dict[str, Any]) -> set:
        """Detect all namespace prefixes used in the workflow JSON.

        Delegates to detect_required_namespaces() which recursively scans
        all type-bearing fields (x:TypeArguments, variableType, type, etc.).

        Args:
            workflow_json: The workflow portion of the JSON data

        Returns:
            Set of prefix strings (e.g., {'sd', 'ui', 'uix'})
        """
        return MetadataManager.detect_required_namespaces(workflow_json)

    @staticmethod
    def generate_minimal_assembly_refs(used_prefixes: set, existing_refs: List[str]) -> List[str]:
        """Generate minimal assembly references from used prefixes and existing refs.

        Three-step logic:
        1. Seed with existing refs (always preserved, never dropped)
        2. Add prefix-derived refs from PREFIX_TO_ASSEMBLIES
        3. Fallback to DEFAULT_ASSEMBLY_REFERENCES only when combined set is empty

        Args:
            used_prefixes: Set of namespace prefixes used in the workflow
            existing_refs: Assembly references from metadata (may be empty)

        Returns:
            Sorted, deduplicated list of assembly reference strings
        """
        # Step 1: Seed with existing refs
        assembly_set = {ref.strip() for ref in existing_refs
                        if ref and isinstance(ref, str) and ref.strip()}

        # Step 2: Add prefix-derived refs
        for prefix in used_prefixes:
            for assembly in PREFIX_TO_ASSEMBLIES.get(prefix, []):
                if assembly:
                    assembly_set.add(assembly)

        # Step 3: Fallback only when combined set is empty
        if not assembly_set:
            return list(DEFAULT_ASSEMBLY_REFERENCES)

        return sorted(assembly_set)

    def apply_xmlns_to_root(self, root: ET.Element, namespace_prefixes: set):
        """Add xmlns attributes directly to the root Activity element.

        This ensures all required namespace declarations appear in the output XAML,
        regardless of whether they were present in the input JSON metadata.
        """
        for prefix in namespace_prefixes:
            if prefix not in NAMESPACES:
                continue
            uri = NAMESPACES[prefix]
            if prefix == '':
                root.set('xmlns', uri)
            else:
                root.set(f'xmlns:{prefix}', uri)

    @staticmethod
    def filter_used_custom_xmlns(custom_bindings: Dict[str, str],
                                  workflow_json: Dict[str, Any]) -> Dict[str, str]:
        """Filter custom xmlns bindings to only those actually referenced in the workflow.

        Canonical prefixes (those present in the NAMESPACES registry) are always
        silently ignored regardless of usage, so they can never override canonical URIs.

        Args:
            custom_bindings: Dict of prefix->URI for non-canonical xmlns bindings
            workflow_json: The workflow JSON structure to scan

        Returns:
            Dict of prefix->URI for bindings that are actually used
        """
        if not custom_bindings:
            return {}

        def scan_node(node: Any, prefix: str) -> bool:
            """Recursively scan for usage of a namespace prefix."""
            prefix_pattern = re.compile(r'\b' + re.escape(prefix) + r':')
            if isinstance(node, str):
                return bool(prefix_pattern.search(node))
            elif isinstance(node, dict):
                for key, val in node.items():
                    if scan_node(val, prefix):
                        return True
            elif isinstance(node, list):
                for item in node:
                    if scan_node(item, prefix):
                        return True
            return False

        result = {}
        for prefix, uri in custom_bindings.items():
            if prefix in NAMESPACES:   # hard guard: canonical prefixes are immutable
                continue
            if scan_node(workflow_json, prefix):
                result[prefix] = uri
        return result


# =============================================================================
# Activity Handler Base Class
# =============================================================================

class ActivityHandler(ABC):
    """Abstract base class for activity handlers."""

    @abstractmethod
    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse an activity element into JSON structure."""
        pass

    @abstractmethod
    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build an activity element from JSON structure."""
        pass


# =============================================================================
# Sequence Handler
# =============================================================================

class SequenceHandler(ActivityHandler):
    """Handler for Sequence activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Sequence element into JSON structure."""
        result = {
            'type': 'Sequence',
            'displayName': element.get('DisplayName', ''),
            'variables': [],
            'children': [],
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse variables
        vars_tag = get_ns_tag('', 'Sequence.Variables')
        vars_elem = element.find(vars_tag)
        if vars_elem is not None:
            for var_elem in vars_elem:
                var_info = self._parse_variable(var_elem)
                if var_info:
                    result['variables'].append(var_info)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        # Parse child activities
        for child in element:
            _, local = parse_tag(child.tag)
            # Skip metadata elements
            if local in ['Sequence.Variables', 'WorkflowViewStateService.ViewState']:
                continue
            if child.tag.endswith('.ViewState'):
                continue

            # Parse child activity
            child_json = parse_activity(child)
            if child_json:
                result['children'].append(child_json)

        return result

    def _parse_variable(self, var_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """Parse a Variable element."""
        name = var_elem.get('Name')
        if not name:
            return None

        type_args_attr = get_ns_tag('x', 'TypeArguments')
        xaml_type = canonicalize_type(var_elem.get(type_args_attr, 'x:String'))

        default = var_elem.get('Default', '')

        return {
            'name': name,
            'type': TypeMapper.xaml_to_json_type(xaml_type),
            'default': unescape_expression(default) if default else '',
        }

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Sequence element from JSON structure."""
        seq_elem = ET.Element(get_ns_tag('', 'Sequence'))

        # Set attributes
        if activity_json.get('displayName'):
            seq_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Sequence', '400,200'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        seq_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Sequence')
        seq_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add variables
        if activity_json.get('variables'):
            vars_elem = ET.SubElement(seq_elem, get_ns_tag('', 'Sequence.Variables'))
            for var_info in activity_json['variables']:
                var_elem = ET.SubElement(vars_elem, get_ns_tag('', 'Variable'))
                var_elem.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(var_info['type']))
                var_elem.set('Name', var_info['name'])
                if var_info.get('default'):
                    var_elem.set('Default', var_info['default'])

        # Add child activities
        for child_json in activity_json.get('children', []):
            child_elem = build_activity(child_json, id_gen)
            if child_elem is not None:
                seq_elem.append(child_elem)

        # Add ViewState
        viewstate = activity_json.get('viewState', {'IsExpanded': True})
        viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
        seq_elem.append(viewstate_elem)

        return seq_elem


# =============================================================================
# Flowchart Handler
# =============================================================================

class FlowchartHandler(ActivityHandler):
    """Handler for Flowchart activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Flowchart element into JSON structure."""
        result = {
            'type': 'Flowchart',
            'displayName': element.get('DisplayName', ''),
            'startNode': None,
            'nodes': [],
            'variables': [],
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse variables
        vars_tag = get_ns_tag('', 'Flowchart.Variables')
        vars_elem = element.find(vars_tag)
        if vars_elem is not None:
            for var_elem in vars_elem:
                var_info = self._parse_variable(var_elem)
                if var_info:
                    result['variables'].append(var_info)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_flowchart_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        # Parse StartNode
        start_tag = get_ns_tag('', 'Flowchart.StartNode')
        start_elem = element.find(start_tag)
        if start_elem is not None:
            ref_tag = get_ns_tag('x', 'Reference')
            ref_elem = start_elem.find(ref_tag)
            if ref_elem is not None and ref_elem.text:
                result['startNode'] = ref_elem.text.strip()

        # Parse nodes (FlowStep, FlowDecision children)
        # Track x:Name references we've already seen as inline nodes
        ref_tag = get_ns_tag('x', 'Reference')
        for child in element:
            _, local = parse_tag(child.tag)
            # Skip metadata elements
            if local in ('Flowchart.Variables', 'Flowchart.StartNode',
                         'WorkflowViewStateService.ViewState'):
                continue
            if child.tag.endswith('.ViewState'):
                continue
            # Skip trailing x:Reference registration elements
            if child.tag == ref_tag:
                continue

            # Parse FlowStep or FlowDecision node
            child_json = parse_activity(child)
            if child_json:
                result['nodes'].append(child_json)

        return result

    def _parse_variable(self, var_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """Parse a Variable element."""
        name = var_elem.get('Name')
        if not name:
            return None

        type_args_attr = get_ns_tag('x', 'TypeArguments')
        xaml_type = canonicalize_type(var_elem.get(type_args_attr, 'x:String'))

        default = var_elem.get('Default', '')

        return {
            'name': name,
            'type': TypeMapper.xaml_to_json_type(xaml_type),
            'default': unescape_expression(default) if default else '',
        }

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Flowchart element from JSON structure."""
        fc_elem = ET.Element(get_ns_tag('', 'Flowchart'))

        # Set DisplayName
        if activity_json.get('displayName'):
            fc_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Flowchart', '614,636'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        fc_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Flowchart')
        fc_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add variables
        if activity_json.get('variables'):
            vars_elem = ET.SubElement(fc_elem, get_ns_tag('', 'Flowchart.Variables'))
            for var_info in activity_json['variables']:
                var_elem = ET.SubElement(vars_elem, get_ns_tag('', 'Variable'))
                var_elem.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(var_info['type']))
                var_elem.set('Name', var_info['name'])
                if var_info.get('default'):
                    var_elem.set('Default', var_info['default'])

        # Add ViewState — auto-generate when missing, preserve when present
        viewstate = activity_json.get('viewState')
        if not viewstate or 'ShapeLocation' not in viewstate or 'ShapeSize' not in viewstate:
            viewstate = {
                'ShapeLocation': '330,10',
                'ShapeSize': '50,50',
            }
            # Compute ConnectorLocation targeting start node's top-center
            _start = activity_json.get('startNode')
            if _start:
                _nodes = activity_json.get('nodes', [])
                _node_by_ref = {n.get('x:Name'): n for n in _nodes}
                _idx_by_ref = {n.get('x:Name'): i for i, n in enumerate(_nodes)}
                _target = _node_by_ref.get(_start)
                if _target:
                    _target_idx = _idx_by_ref.get(_start, 0)
                    _target_type = _target.get('type', '')
                    if _target_type == 'FlowStep':
                        _cx = 300 + 110 // 2  # 355
                        _ty = 200 + _target_idx * 100
                    elif _target_type == 'FlowDecision':
                        _cx = 325 + 60 // 2  # 355
                        _ty = 200 + _target_idx * 100
                    else:
                        _cx = None
                        _ty = None
                    if _cx is not None:
                        viewstate['ConnectorLocation'] = f"355,60 {_cx},{_ty}"
        viewstate_elem = ViewStateBuilder.create_flowchart_viewstate(viewstate)
        fc_elem.append(viewstate_elem)

        # Add StartNode
        start_node = activity_json.get('startNode')
        if start_node:
            start_elem = ET.SubElement(fc_elem, get_ns_tag('', 'Flowchart.StartNode'))
            ref_elem = ET.SubElement(start_elem, get_ns_tag('x', 'Reference'))
            ref_elem.text = start_node

        # Build and add nodes — inject index/sibling context for default ViewState
        all_nodes = activity_json.get('nodes', [])
        for idx, node_json in enumerate(all_nodes):
            node_json['_node_index'] = idx
            node_json['_parent_nodes'] = all_nodes
        node_names = []
        for node_json in all_nodes:
            node_elem = build_activity(node_json, id_gen)
            if node_elem is not None:
                fc_elem.append(node_elem)
                # Collect x:Name for trailing references
                x_name = node_elem.get(get_ns_tag('x', 'Name'))
                if x_name:
                    node_names.append(x_name)
                # Also collect names from nested inline nodes
                self._collect_nested_names(node_elem, node_names)

        # Add trailing x:Reference registrations for all nodes
        for name in node_names:
            ref_elem = ET.SubElement(fc_elem, get_ns_tag('x', 'Reference'))
            ref_elem.text = name

        return fc_elem

    def _collect_nested_names(self, element: ET.Element, names: List[str]):
        """Recursively collect x:Name attributes from nested FlowStep/FlowDecision elements."""
        x_name_attr = get_ns_tag('x', 'Name')
        for child in element:
            x_name = child.get(x_name_attr)
            if x_name and x_name not in names:
                _, local = parse_tag(child.tag)
                if local in ('FlowStep', 'FlowDecision'):
                    names.append(x_name)
            self._collect_nested_names(child, names)


# =============================================================================
# FlowStep Handler
# =============================================================================

class FlowStepHandler(ActivityHandler):
    """Handler for FlowStep activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse FlowStep element into JSON structure."""
        result = {
            'type': 'FlowStep',
            'activity': None,
            'next': None,
        }

        # Extract x:Name (required for reference ID)
        x_name = element.get(get_ns_tag('x', 'Name'))
        if x_name:
            result['x:Name'] = x_name

        # Extract DisplayName (optional for FlowStep)
        display_name = element.get('DisplayName')
        if display_name:
            result['displayName'] = display_name

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_flowchart_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        # Parse child activity (first non-metadata child)
        viewstate_tag = get_ns_tag('sap', 'WorkflowViewStateService.ViewState')
        next_tag = get_ns_tag('', 'FlowStep.Next')
        for child in element:
            if child.tag == viewstate_tag:
                continue
            if child.tag == next_tag:
                continue
            if child.tag.endswith('.ViewState'):
                continue
            # This is the activity child
            child_json = parse_activity(child)
            if child_json:
                result['activity'] = child_json
                break

        # Parse FlowStep.Next
        next_elem = element.find(next_tag)
        if next_elem is not None:
            # Check for x:Reference (back-reference to existing node)
            ref_tag = get_ns_tag('x', 'Reference')
            ref_elem = next_elem.find(ref_tag)
            if ref_elem is not None and ref_elem.text:
                result['next'] = ref_elem.text.strip()
            else:
                # Check for inline nested FlowStep or FlowDecision
                for child in next_elem:
                    child_json = parse_activity(child)
                    if child_json:
                        result['next'] = child_json
                        break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build FlowStep element from JSON structure."""
        fs_elem = ET.Element(get_ns_tag('', 'FlowStep'))

        # Set x:Name (required)
        x_name = activity_json.get('x:Name')
        if x_name:
            fs_elem.set(get_ns_tag('x', 'Name'), x_name)

        # Set DisplayName (optional)
        if activity_json.get('displayName'):
            fs_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        if activity_json.get('hintSize'):
            sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
            fs_elem.set(sap_hint, activity_json['hintSize'])

        # Set IdRef
        if activity_json.get('idRef'):
            fs_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), activity_json['idRef'])

        # Build ViewState — auto-generate when missing, merge defaults for partial
        viewstate = activity_json.get('viewState')
        idx = activity_json.get('_node_index', 0)
        x, y, width, height = 300, 200 + idx * 100, 110, 70
        if not viewstate:
            viewstate = {
                'ShapeLocation': f'{x},{y}',
                'ShapeSize': f'{width},{height}',
            }
        else:
            # Merge defaults for missing keys without overwriting provided keys
            if 'ShapeLocation' not in viewstate:
                viewstate['ShapeLocation'] = f'{x},{y}'
            if 'ShapeSize' not in viewstate:
                viewstate['ShapeSize'] = f'{width},{height}'
        # Compute ConnectorLocation to next node when absent
        if 'ConnectorLocation' not in viewstate:
            next_ref = activity_json.get('next')
            if isinstance(next_ref, str):
                parent_nodes = activity_json.get('_parent_nodes', [])
                _nb = {n.get('x:Name'): n for n in parent_nodes}
                _ib = {n.get('x:Name'): i for i, n in enumerate(parent_nodes)}
                _tgt = _nb.get(next_ref)
                if _tgt:
                    _ti = _ib.get(next_ref, 0)
                    _tt = _tgt.get('type', '')
                    if _tt == 'FlowStep':
                        _nx, _ny, _nw = 300, 200 + _ti * 100, 110
                    elif _tt == 'FlowDecision':
                        _nx, _ny, _nw = 325, 200 + _ti * 100, 60
                    else:
                        _nx, _ny, _nw = None, None, None
                    if _nx is not None:
                        cx_bottom = x + width // 2
                        cy_bottom = y + height
                        next_cx_top = _nx + _nw // 2
                        viewstate['ConnectorLocation'] = f"{cx_bottom},{cy_bottom} {next_cx_top},{_ny}"
        viewstate_elem = ViewStateBuilder.create_flowchart_viewstate(viewstate)
        fs_elem.append(viewstate_elem)

        # Build activity child
        activity = activity_json.get('activity')
        if activity:
            activity_elem = build_activity(activity, id_gen)
            if activity_elem is not None:
                fs_elem.append(activity_elem)

        # Build FlowStep.Next
        next_val = activity_json.get('next')
        if next_val is not None:
            next_elem = ET.SubElement(fs_elem, get_ns_tag('', 'FlowStep.Next'))
            if isinstance(next_val, str):
                # x:Reference to another node
                ref_elem = ET.SubElement(next_elem, get_ns_tag('x', 'Reference'))
                ref_elem.text = next_val
            elif isinstance(next_val, dict):
                # Inline nested node
                nested_elem = build_activity(next_val, id_gen)
                if nested_elem is not None:
                    next_elem.append(nested_elem)

        # Clean up injected context keys
        activity_json.pop('_node_index', None)
        activity_json.pop('_parent_nodes', None)

        return fs_elem


# =============================================================================
# FlowDecision Handler
# =============================================================================

class FlowDecisionHandler(ActivityHandler):
    """Handler for FlowDecision activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse FlowDecision element into JSON structure."""
        result = {
            'type': 'FlowDecision',
            'condition': '',
            'true': None,
            'false': None,
        }

        # Extract x:Name (required for reference ID)
        x_name = element.get(get_ns_tag('x', 'Name'))
        if x_name:
            result['x:Name'] = x_name

        # Extract Condition
        condition = element.get('Condition', '')
        if condition:
            result['condition'] = unescape_expression(condition)

        # Extract DisplayName
        display_name = element.get('DisplayName')
        if display_name:
            result['displayName'] = display_name

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_flowchart_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        # Parse FlowDecision.True
        true_tag = get_ns_tag('', 'FlowDecision.True')
        true_elem = element.find(true_tag)
        if true_elem is not None:
            result['true'] = self._parse_branch(true_elem)

        # Parse FlowDecision.False
        false_tag = get_ns_tag('', 'FlowDecision.False')
        false_elem = element.find(false_tag)
        if false_elem is not None:
            result['false'] = self._parse_branch(false_elem)

        return result

    def _parse_branch(self, branch_elem: ET.Element):
        """Parse a True or False branch element. Returns string reference or inline node dict."""
        ref_tag = get_ns_tag('x', 'Reference')
        ref_elem = branch_elem.find(ref_tag)
        if ref_elem is not None and ref_elem.text:
            return ref_elem.text.strip()

        # Inline FlowStep or FlowDecision
        for child in branch_elem:
            child_json = parse_activity(child)
            if child_json:
                return child_json

        return None

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build FlowDecision element from JSON structure."""
        fd_elem = ET.Element(get_ns_tag('', 'FlowDecision'))

        # Set x:Name (required)
        x_name = activity_json.get('x:Name')
        if x_name:
            fd_elem.set(get_ns_tag('x', 'Name'), x_name)

        # Set Condition
        condition = activity_json.get('condition', '')
        if condition:
            fd_elem.set('Condition', condition)

        # Set DisplayName
        if activity_json.get('displayName'):
            fd_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        if activity_json.get('hintSize'):
            sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
            fd_elem.set(sap_hint, activity_json['hintSize'])

        # Set IdRef
        if activity_json.get('idRef'):
            fd_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), activity_json['idRef'])

        # Build ViewState — auto-generate when missing, merge defaults for partial
        viewstate = activity_json.get('viewState')
        idx = activity_json.get('_node_index', 0)
        x, y, width, height = 325, 200 + idx * 100, 60, 60
        cy = y + height // 2
        if not viewstate:
            viewstate = {
                'ShapeLocation': f'{x},{y}',
                'ShapeSize': f'{width},{height}',
            }
        else:
            # Merge defaults for missing keys without overwriting provided keys
            if 'ShapeLocation' not in viewstate:
                viewstate['ShapeLocation'] = f'{x},{y}'
            if 'ShapeSize' not in viewstate:
                viewstate['ShapeSize'] = f'{width},{height}'
        # Compute TrueConnector/FalseConnector when absent
        parent_nodes = activity_json.get('_parent_nodes', [])
        _nb = {n.get('x:Name'): n for n in parent_nodes}
        _ib = {n.get('x:Name'): i for i, n in enumerate(parent_nodes)}

        # TrueConnector
        if 'TrueConnector' not in viewstate:
            true_ref = activity_json.get('true')
            if isinstance(true_ref, str) and true_ref in _nb:
                _tgt = _nb[true_ref]
                _ti = _ib.get(true_ref, 0)
                _tt = _tgt.get('type', '')
                if _tt == 'FlowStep':
                    _ty = 200 + _ti * 100
                elif _tt == 'FlowDecision':
                    _ty = 200 + _ti * 100
                else:
                    _ty = None
                if _ty is not None:
                    viewstate['TrueConnector'] = f"{x},{cy} 150,{cy} 150,{_ty}"

        # FalseConnector
        if 'FalseConnector' not in viewstate:
            false_ref = activity_json.get('false')
            if isinstance(false_ref, str) and false_ref in _nb:
                _tgt = _nb[false_ref]
                _ti = _ib.get(false_ref, 0)
                _tt = _tgt.get('type', '')
                if _tt == 'FlowStep':
                    _ty = 200 + _ti * 100
                elif _tt == 'FlowDecision':
                    _ty = 200 + _ti * 100
                else:
                    _ty = None
                if _ty is not None:
                    viewstate['FalseConnector'] = f"{x + width},{cy} 560,{cy} 560,{_ty}"
        viewstate_elem = ViewStateBuilder.create_flowchart_viewstate(viewstate)
        fd_elem.append(viewstate_elem)

        # Build FlowDecision.True
        true_val = activity_json.get('true')
        if true_val is not None:
            true_elem = ET.SubElement(fd_elem, get_ns_tag('', 'FlowDecision.True'))
            self._build_branch(true_elem, true_val, id_gen)

        # Build FlowDecision.False
        false_val = activity_json.get('false')
        if false_val is not None:
            false_elem = ET.SubElement(fd_elem, get_ns_tag('', 'FlowDecision.False'))
            self._build_branch(false_elem, false_val, id_gen)

        # Clean up injected context keys
        activity_json.pop('_node_index', None)
        activity_json.pop('_parent_nodes', None)

        return fd_elem

    def _build_branch(self, parent_elem: ET.Element, branch_val, id_gen: IdRefGenerator):
        """Build a True or False branch. branch_val is string reference or inline node dict."""
        if isinstance(branch_val, str):
            ref_elem = ET.SubElement(parent_elem, get_ns_tag('x', 'Reference'))
            ref_elem.text = branch_val
        elif isinstance(branch_val, dict):
            nested_elem = build_activity(branch_val, id_gen)
            if nested_elem is not None:
                parent_elem.append(nested_elem)


# =============================================================================
# Assign Handler
# =============================================================================

class AssignHandler(ActivityHandler):
    """Handler for Assign activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Assign element into JSON structure."""
        result = {
            'type': 'Assign',
            'displayName': element.get('DisplayName', ''),
            'to': {},
            'value': {},
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse Assign.To
        to_tag = get_ns_tag('', 'Assign.To')
        to_elem = element.find(to_tag)
        if to_elem is not None:
            out_arg = to_elem.find(get_ns_tag('', 'OutArgument'))
            if out_arg is not None:
                type_args_attr = get_ns_tag('x', 'TypeArguments')
                result['to'] = {
                    'type': TypeMapper.xaml_to_json_type(canonicalize_type(out_arg.get(type_args_attr, 'x:String'))),
                    'value': unescape_expression(out_arg.text or ''),
                }

        # Parse Assign.Value
        value_tag = get_ns_tag('', 'Assign.Value')
        value_elem = element.find(value_tag)
        if value_elem is not None:
            in_arg = value_elem.find(get_ns_tag('', 'InArgument'))
            if in_arg is not None:
                type_args_attr = get_ns_tag('x', 'TypeArguments')
                result['value'] = {
                    'type': TypeMapper.xaml_to_json_type(canonicalize_type(in_arg.get(type_args_attr, 'x:String'))),
                    'value': unescape_expression(in_arg.text or ''),
                }

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Assign element from JSON structure."""
        assign_elem = ET.Element(get_ns_tag('', 'Assign'))

        # Set attributes
        if activity_json.get('displayName'):
            assign_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Assign', '262,60'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        assign_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Assign')
        assign_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add Assign.To
        to_info = activity_json.get('to', {})
        to_elem = ET.SubElement(assign_elem, get_ns_tag('', 'Assign.To'))
        out_arg = ET.SubElement(to_elem, get_ns_tag('', 'OutArgument'))
        out_arg.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(to_info.get('type', 'String')))
        out_arg.text = to_info.get('value', '')

        # Add Assign.Value
        value_info = activity_json.get('value', {})
        value_elem = ET.SubElement(assign_elem, get_ns_tag('', 'Assign.Value'))
        in_arg = ET.SubElement(value_elem, get_ns_tag('', 'InArgument'))
        in_arg.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(value_info.get('type', 'String')))
        in_arg.text = value_info.get('value', '')

        return assign_elem


# =============================================================================
# If Handler
# =============================================================================

class IfHandler(ActivityHandler):
    """Handler for If activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse If element into JSON structure."""
        result = {
            'type': 'If',
            'condition': unescape_expression(element.get('Condition', '')),
            'then': None,
            'else': None,
        }

        # Extract DisplayName if present
        if element.get('DisplayName'):
            result['displayName'] = element.get('DisplayName')

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse If.Then
        then_tag = get_ns_tag('', 'If.Then')
        then_elem = element.find(then_tag)
        if then_elem is not None and len(then_elem) > 0:
            result['then'] = parse_activity(then_elem[0])

        # Parse If.Else
        else_tag = get_ns_tag('', 'If.Else')
        else_elem = element.find(else_tag)
        if else_elem is not None and len(else_elem) > 0:
            result['else'] = parse_activity(else_elem[0])

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build If element from JSON structure."""
        if_elem = ET.Element(get_ns_tag('', 'If'))

        # Set Condition
        if_elem.set('Condition', activity_json.get('condition', ''))

        # Set DisplayName if present
        if activity_json.get('displayName'):
            if_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('If', '464,200'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('If')
        if_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add If.Then
        if activity_json.get('then'):
            then_elem = ET.SubElement(if_elem, get_ns_tag('', 'If.Then'))
            then_child = build_activity(activity_json['then'], id_gen)
            if then_child is not None:
                then_elem.append(then_child)

        # Add If.Else
        if activity_json.get('else'):
            else_elem = ET.SubElement(if_elem, get_ns_tag('', 'If.Else'))
            else_child = build_activity(activity_json['else'], id_gen)
            if else_child is not None:
                else_elem.append(else_child)

        # Add ViewState
        viewstate = activity_json.get('viewState', {'IsExpanded': True})
        viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
        if_elem.append(viewstate_elem)

        return if_elem


# =============================================================================
# LogMessage Handler
# =============================================================================

class LogMessageHandler(ActivityHandler):
    """Handler for LogMessage activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse LogMessage element into JSON structure."""
        result = {
            'type': 'LogMessage',
            'displayName': element.get('DisplayName', ''),
            'level': element.get('Level', 'Info'),
            'message': unescape_expression(element.get('Message', '')),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build LogMessage element from JSON structure."""
        log_elem = ET.Element(get_ns_tag('ui', 'LogMessage'))

        # Set attributes
        if activity_json.get('displayName'):
            log_elem.set('DisplayName', activity_json['displayName'])

        log_elem.set('Level', activity_json.get('level', 'Info'))
        log_elem.set('Message', activity_json.get('message', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('LogMessage', '262,60'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        log_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('LogMessage')
        log_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            log_elem.append(viewstate_elem)

        return log_elem


# =============================================================================
# InvokeWorkflowFile Handler
# =============================================================================

class InvokeWorkflowFileHandler(ActivityHandler):
    """Handler for InvokeWorkflowFile activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse InvokeWorkflowFile element into JSON structure."""
        result = {
            'type': 'InvokeWorkflowFile',
            'displayName': element.get('DisplayName', ''),
            'fileName': element.get('WorkflowFileName', ''),
            'arguments': [],
        }

        # Extract optional attributes
        if element.get('UnSafe'):
            result['unSafe'] = element.get('UnSafe').lower() == 'true'
        if element.get('ContinueOnError'):
            result['continueOnError'] = element.get('ContinueOnError').lower() == 'true'

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse Arguments
        args_tag = get_ns_tag('ui', 'InvokeWorkflowFile.Arguments')
        args_elem = element.find(args_tag)
        if args_elem is not None:
            for arg_elem in args_elem:
                arg_info = self._parse_argument(arg_elem)
                if arg_info:
                    result['arguments'].append(arg_info)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def _parse_argument(self, arg_elem: ET.Element) -> Optional[Dict[str, str]]:
        """Parse a single argument element."""
        _, local = parse_tag(arg_elem.tag)

        # Determine direction from tag
        if local == 'InArgument':
            direction = 'In'
        elif local == 'OutArgument':
            direction = 'Out'
        elif local == 'InOutArgument':
            direction = 'InOut'
        else:
            return None

        # Get key
        key_attr = get_ns_tag('x', 'Key')
        key = arg_elem.get(key_attr)
        if not key:
            return None

        # Get type
        type_args_attr = get_ns_tag('x', 'TypeArguments')
        xaml_type = canonicalize_type(arg_elem.get(type_args_attr, 'x:String'))

        return {
            'key': key,
            'direction': direction,
            'type': TypeMapper.xaml_to_json_type(xaml_type),
            'value': unescape_expression(arg_elem.text or ''),
        }

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build InvokeWorkflowFile element from JSON structure."""
        invoke_elem = ET.Element(get_ns_tag('ui', 'InvokeWorkflowFile'))

        # Set attributes
        if activity_json.get('displayName'):
            invoke_elem.set('DisplayName', activity_json['displayName'])

        invoke_elem.set('WorkflowFileName', activity_json.get('fileName', ''))

        if activity_json.get('unSafe'):
            invoke_elem.set('UnSafe', str(activity_json['unSafe']).lower())

        if activity_json.get('continueOnError'):
            invoke_elem.set('ContinueOnError', str(activity_json['continueOnError']).lower())

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('InvokeWorkflowFile', '318,88'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        invoke_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('InvokeWorkflowFile')
        invoke_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add Arguments
        arguments = activity_json.get('arguments', [])
        if arguments:
            args_elem = ET.SubElement(invoke_elem, get_ns_tag('ui', 'InvokeWorkflowFile.Arguments'))
            for arg_info in arguments:
                self._build_argument(args_elem, arg_info)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            invoke_elem.append(viewstate_elem)

        return invoke_elem

    def _build_argument(self, parent: ET.Element, arg_info: Dict[str, str]):
        """Build a single argument element."""
        direction = arg_info.get('direction', 'In')

        if direction == 'Out':
            tag = get_ns_tag('', 'OutArgument')
        elif direction == 'InOut':
            tag = get_ns_tag('', 'InOutArgument')
        else:
            tag = get_ns_tag('', 'InArgument')

        arg_elem = ET.SubElement(parent, tag)
        arg_elem.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(arg_info.get('type', 'String')))
        arg_elem.set(get_ns_tag('x', 'Key'), arg_info.get('key', ''))
        arg_elem.text = arg_info.get('value', '')


# =============================================================================
# Helper Classes for Complex Structures
# =============================================================================

class ActivityActionParser:
    """Static methods to parse ActivityAction structures."""

    @staticmethod
    def parse_activity_action(element: ET.Element, expected_arg_type: Optional[str] = None) -> Dict[str, Any]:
        """
        Parse an ActivityAction element.

        Args:
            element: The ActivityAction element
            expected_arg_type: Expected type argument (e.g., 's:Exception', 'sd:DataRow')

        Returns:
            Dictionary with variableName, variableType, and activity
        """
        result = {
            'variableName': '',
            'variableType': '',
            'activity': None,
        }

        # Get type arguments from ActivityAction
        type_args_attr = get_ns_tag('x', 'TypeArguments')
        if type_args_attr in element.attrib:
            result['variableType'] = TypeMapper.xaml_to_json_type(canonicalize_type(element.get(type_args_attr)))

        # Find DelegateInArgument
        arg_elem_tag = get_ns_tag('', 'ActivityAction.Argument')
        arg_elem = element.find(arg_elem_tag)
        if arg_elem is not None:
            delegate_tag = get_ns_tag('', 'DelegateInArgument')
            delegate_elem = arg_elem.find(delegate_tag)
            if delegate_elem is not None:
                result['variableName'] = delegate_elem.get('Name', '')
                # Get type from DelegateInArgument if not already set
                if not result['variableType'] and type_args_attr in delegate_elem.attrib:
                    result['variableType'] = TypeMapper.xaml_to_json_type(canonicalize_type(delegate_elem.get(type_args_attr)))

        # Parse nested activity (first non-metadata child)
        for child in element:
            _, local = parse_tag(child.tag)
            if local not in ['ActivityAction.Argument']:
                child_activity = parse_activity(child)
                if child_activity:
                    result['activity'] = child_activity
                    break

        return result


class ActivityActionBuilder:
    """Static methods to build ActivityAction structures."""

    @staticmethod
    def build_activity_action(
        parent_elem: ET.Element,
        tag_name: str,
        action_json: Dict[str, Any],
        id_gen: IdRefGenerator,
        type_arg: Optional[str] = None
    ) -> ET.Element:
        """
        Build an ActivityAction element.

        Args:
            parent_elem: Parent element to append to
            tag_name: Tag name for the wrapper (e.g., 'ForEach.Body')
            action_json: JSON data with variableName, variableType, activity
            id_gen: IdRef generator
            type_arg: Type argument for ActivityAction (optional override)

        Returns:
            The created ActivityAction element
        """
        # Create wrapper element (e.g., ForEach.Body)
        wrapper = ET.SubElement(parent_elem, get_ns_tag('', tag_name))

        # Create ActivityAction element
        activity_action = ET.SubElement(wrapper, get_ns_tag('', 'ActivityAction'))

        # Set TypeArguments
        var_type = type_arg or action_json.get('variableType', 'String')
        xaml_type = TypeMapper.json_to_xaml_type(var_type)
        activity_action.set(get_ns_tag('x', 'TypeArguments'), xaml_type)

        # Create ActivityAction.Argument with DelegateInArgument
        arg_wrapper = ET.SubElement(activity_action, get_ns_tag('', 'ActivityAction.Argument'))
        delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
        delegate.set(get_ns_tag('x', 'TypeArguments'), xaml_type)
        delegate.set('Name', action_json.get('variableName', 'item'))

        # Build nested activity
        if action_json.get('activity'):
            activity_elem = build_activity(action_json['activity'], id_gen)
            if activity_elem is not None:
                activity_action.append(activity_elem)

        return wrapper

    @staticmethod
    def build_simple_activity_action(
        parent_elem: ET.Element,
        tag_name: str,
        activity_json: Optional[Dict[str, Any]],
        id_gen: IdRefGenerator
    ) -> ET.Element:
        """
        Build a simple ActivityAction without type arguments (for RetryScope.ActivityBody).

        Args:
            parent_elem: Parent element to append to
            tag_name: Tag name for the wrapper
            activity_json: Nested activity JSON
            id_gen: IdRef generator

        Returns:
            The created wrapper element
        """
        # Create wrapper element
        wrapper = ET.SubElement(parent_elem, tag_name)

        # Create ActivityAction element (no TypeArguments)
        activity_action = ET.SubElement(wrapper, get_ns_tag('', 'ActivityAction'))

        # Build nested activity
        if activity_json:
            activity_elem = build_activity(activity_json, id_gen)
            if activity_elem is not None:
                activity_action.append(activity_elem)

        return wrapper


class ActivityFuncBuilder:
    """Static methods to build ActivityFunc structures."""

    @staticmethod
    def build_activity_func(
        parent_elem: ET.Element,
        tag_name: str,
        type_args: str = 'x:Boolean'
    ) -> ET.Element:
        """
        Build an empty ActivityFunc element (used for RetryScope.Condition).

        Args:
            parent_elem: Parent element to append to
            tag_name: Tag name for the wrapper
            type_args: Type arguments for the ActivityFunc

        Returns:
            The created wrapper element
        """
        wrapper = ET.SubElement(parent_elem, tag_name)
        activity_func = ET.SubElement(wrapper, get_ns_tag('', 'ActivityFunc'))
        activity_func.set(get_ns_tag('x', 'TypeArguments'), type_args)
        return wrapper


# =============================================================================
# UI Automation Helper Classes
# =============================================================================

class TargetParser:
    """Static methods to parse Target (anchor) elements from UI automation activities."""

    # All known attributes for Target elements
    TARGET_ATTRIBUTES = [
        'ContentHash', 'CVScreenId', 'CvTextArea', 'CvTextArgument', 'CvType', 'CvElementArea',
        'DesignTimeRectangle', 'ElementType', 'FuzzySelectorArgument', 'Guid',
        'SearchSteps', 'TargetType',
    ]

    @staticmethod
    def parse_target(element: ET.Element) -> Dict[str, Any]:
        """
        Parse a Target element (used as anchor in TargetAnchorable).

        Args:
            element: The uix:Target element

        Returns:
            Dictionary with all target attributes
        """
        result = {}
        for attr in TargetParser.TARGET_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                result[attr] = val
        return result


class TargetBuilder:
    """Static methods to build Target (anchor) elements for UI automation activities."""

    @staticmethod
    def build_target(target_json: Dict[str, Any]) -> ET.Element:
        """
        Build a uix:Target element from JSON.

        Args:
            target_json: Dictionary with target attributes

        Returns:
            ET.Element for the Target
        """
        target_elem = ET.Element(get_ns_tag('uix', 'Target'))
        for attr in TargetParser.TARGET_ATTRIBUTES:
            if attr in target_json:
                target_elem.set(attr, target_json[attr])
        return target_elem


class TargetAnchorableParser:
    """Static methods to parse TargetAnchorable structures from UI automation activities."""

    # All known attributes for TargetAnchorable elements
    TARGETANCHORABLE_ATTRIBUTES = [
        'CVScreenId', 'ContentHash', 'CvTextArea', 'CvTextArgument', 'CvType',
        'CvElementArea', 'DesignTimeRectangle', 'DesignTimeScaleFactor', 'ElementType',
        'ElementVisibilityArgument', 'FullSelectorArgument', 'FuzzySelectorArgument',
        'Guid', 'InformativeScreenshot', 'IsResponsive', 'Reference', 'ScopeSelectorArgument',
        'SearchSteps', 'TargetType', 'Version', 'WaitForReadyArgument',
    ]

    @staticmethod
    def parse_target_anchorable(element: ET.Element) -> Dict[str, Any]:
        """
        Parse a TargetAnchorable element.

        Args:
            element: The uix:TargetAnchorable element

        Returns:
            Dictionary with all attributes and nested anchors
        """
        result = {}

        # Extract all known attributes
        for attr in TargetAnchorableParser.TARGETANCHORABLE_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                result[attr] = val

        # Parse TargetAnchorable.Anchors if present
        anchors_tag = get_ns_tag('uix', 'TargetAnchorable.Anchors')
        anchors_elem = element.find(anchors_tag)
        if anchors_elem is not None:
            # Find the scg:List element
            list_tag = get_ns_tag('scg', 'List')
            list_elem = anchors_elem.find(list_tag)
            if list_elem is not None:
                anchors = []
                # Parse each uix:Target child
                target_tag = get_ns_tag('uix', 'Target')
                for target_elem in list_elem.findall(target_tag):
                    anchors.append(TargetParser.parse_target(target_elem))
                if anchors:
                    result['anchors'] = anchors

        # Parse TargetAnchorable.PointOffset if present
        offset_tag = get_ns_tag('uix', 'TargetAnchorable.PointOffset')
        offset_elem = element.find(offset_tag)
        if offset_elem is not None:
            # Store the raw text or InArgument structure
            result['pointOffset'] = ET.tostring(offset_elem, encoding='unicode')

        return result


class TargetAnchorableBuilder:
    """Static methods to build TargetAnchorable structures for UI automation activities."""

    @staticmethod
    def build_target_anchorable(target_json: Dict[str, Any]) -> ET.Element:
        """
        Build a uix:TargetAnchorable element from JSON.

        Args:
            target_json: Dictionary with attributes and optional anchors

        Returns:
            ET.Element for the TargetAnchorable
        """
        target_elem = ET.Element(get_ns_tag('uix', 'TargetAnchorable'))

        # Set all known attributes
        for attr in TargetAnchorableParser.TARGETANCHORABLE_ATTRIBUTES:
            if attr in target_json:
                target_elem.set(attr, target_json[attr])

        # Build TargetAnchorable.Anchors if present
        anchors = target_json.get('anchors')
        if anchors:
            anchors_wrapper = ET.SubElement(target_elem, get_ns_tag('uix', 'TargetAnchorable.Anchors'))
            list_elem = ET.SubElement(anchors_wrapper, get_ns_tag('scg', 'List'))
            list_elem.set(get_ns_tag('x', 'TypeArguments'), 'uix:ITarget')
            list_elem.set('Capacity', str(len(anchors)))
            for anchor_json in anchors:
                anchor_elem = TargetBuilder.build_target(anchor_json)
                list_elem.append(anchor_elem)

        # Restore PointOffset if present (stored as raw XML)
        point_offset = target_json.get('pointOffset')
        if point_offset:
            try:
                offset_elem = ET.fromstring(point_offset)
                target_elem.append(offset_elem)
            except ET.ParseError:
                pass

        return target_elem


class SearchedElementParser:
    """Static methods to parse SearchedElement structures from UI automation activities."""

    @staticmethod
    def parse_searched_element(element: ET.Element) -> Dict[str, Any]:
        """
        Parse a SearchedElement element.

        Args:
            element: The uix:SearchedElement element

        Returns:
            Dictionary with target and optional properties
        """
        result = {}

        # Parse SearchedElement.Target containing TargetAnchorable
        target_tag = get_ns_tag('uix', 'SearchedElement.Target')
        target_elem = element.find(target_tag)
        if target_elem is not None:
            target_anchorable_tag = get_ns_tag('uix', 'TargetAnchorable')
            target_anchorable = target_elem.find(target_anchorable_tag)
            if target_anchorable is not None:
                result['target'] = TargetAnchorableParser.parse_target_anchorable(target_anchorable)

        # Parse SearchedElement.Timeout if present
        timeout_tag = get_ns_tag('uix', 'SearchedElement.Timeout')
        timeout_elem = element.find(timeout_tag)
        if timeout_elem is not None:
            # Store as raw XML (InArgument structure)
            result['timeout'] = ET.tostring(timeout_elem, encoding='unicode')

        # Parse SearchedElement.OutUiElement if present
        out_elem_tag = get_ns_tag('uix', 'SearchedElement.OutUiElement')
        out_elem = element.find(out_elem_tag)
        if out_elem is not None:
            # Store as raw XML (OutArgument structure)
            result['outUiElement'] = ET.tostring(out_elem, encoding='unicode')

        return result


class SearchedElementBuilder:
    """Static methods to build SearchedElement structures for UI automation activities."""

    @staticmethod
    def build_searched_element(searched_json: Dict[str, Any]) -> ET.Element:
        """
        Build a uix:SearchedElement element from JSON.

        Args:
            searched_json: Dictionary with target and optional properties

        Returns:
            ET.Element for the SearchedElement
        """
        searched_elem = ET.Element(get_ns_tag('uix', 'SearchedElement'))

        # Build SearchedElement.OutUiElement if present (before Target)
        out_ui_elem = searched_json.get('outUiElement')
        if out_ui_elem:
            if isinstance(out_ui_elem, str) and out_ui_elem.strip().startswith('<'):
                # Raw XML path - existing behavior for callers providing full XML
                try:
                    out_elem = ET.fromstring(out_ui_elem)
                    searched_elem.append(out_elem)
                except ET.ParseError:
                    pass
            else:
                # Plain string expression - build proper OutArgument structure
                out_wrapper = ET.SubElement(searched_elem, get_ns_tag('uix', 'SearchedElement.OutUiElement'))
                out_arg = ET.SubElement(out_wrapper, get_ns_tag('', 'OutArgument'))
                out_arg.set(get_ns_tag('x', 'TypeArguments'), 'ui:UiElement')
                if out_ui_elem:
                    out_arg.text = str(out_ui_elem)

        # Build SearchedElement.Target with TargetAnchorable
        target_json = searched_json.get('target')
        if target_json:
            target_wrapper = ET.SubElement(searched_elem, get_ns_tag('uix', 'SearchedElement.Target'))
            target_anchorable = TargetAnchorableBuilder.build_target_anchorable(target_json)
            target_wrapper.append(target_anchorable)

        # Build SearchedElement.Timeout if present
        timeout = searched_json.get('timeout')
        if timeout:
            if isinstance(timeout, str) and timeout.strip().startswith('<'):
                # Raw XML path - existing behavior for callers providing full XML
                try:
                    timeout_elem = ET.fromstring(timeout)
                    searched_elem.append(timeout_elem)
                except ET.ParseError:
                    pass
            else:
                # Plain string expression - build proper InArgument structure
                timeout_wrapper = ET.SubElement(searched_elem, get_ns_tag('uix', 'SearchedElement.Timeout'))
                in_arg = ET.SubElement(timeout_wrapper, get_ns_tag('', 'InArgument'))
                in_arg.set(get_ns_tag('x', 'TypeArguments'), 'x:Double')
                if timeout:
                    in_arg.text = str(timeout)

        return searched_elem


class TargetAppParser:
    """Static methods to parse TargetApp structures from NApplicationCard."""

    # All known attributes for TargetApp elements
    TARGETAPP_ATTRIBUTES = [
        'Area', 'Arguments', 'ContentHash', 'FilePath', 'IconBase64',
        'InformativeScreenshot', 'Reference', 'Selector', 'Version', 'WorkingDirectory',
    ]

    @staticmethod
    def parse_target_app(element: ET.Element) -> Dict[str, Any]:
        """
        Parse a TargetApp element.

        Args:
            element: The uix:TargetApp element

        Returns:
            Dictionary with all target app attributes
        """
        result = {}

        # Extract all known attributes
        for attr in TargetAppParser.TARGETAPP_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                result[attr] = val

        # Parse TargetApp.Arguments if present (InArgument)
        args_tag = get_ns_tag('uix', 'TargetApp.Arguments')
        args_elem = element.find(args_tag)
        if args_elem is not None:
            in_arg = args_elem.find(get_ns_tag('', 'InArgument'))
            if in_arg is not None:
                type_args = canonicalize_type(in_arg.get(get_ns_tag('x', 'TypeArguments'), 'x:String'))
                val = in_arg.text or ''
                result['argumentsValue'] = val
                result['argumentsType'] = type_args

        # Parse TargetApp.WorkingDirectory if present (InArgument)
        wd_tag = get_ns_tag('uix', 'TargetApp.WorkingDirectory')
        wd_elem = element.find(wd_tag)
        if wd_elem is not None:
            in_arg = wd_elem.find(get_ns_tag('', 'InArgument'))
            if in_arg is not None:
                type_args = canonicalize_type(in_arg.get(get_ns_tag('x', 'TypeArguments'), 'x:String'))
                val = in_arg.text or ''
                result['workingDirectoryValue'] = val
                result['workingDirectoryType'] = type_args

        return result


class TargetAppBuilder:
    """Static methods to build TargetApp structures for NApplicationCard."""

    @staticmethod
    def build_target_app(target_app_json: Dict[str, Any]) -> ET.Element:
        """
        Build a uix:TargetApp element from JSON.

        Args:
            target_app_json: Dictionary with target app attributes

        Returns:
            ET.Element for the TargetApp
        """
        target_app = ET.Element(get_ns_tag('uix', 'TargetApp'))

        # Set all known attributes
        for attr in TargetAppParser.TARGETAPP_ATTRIBUTES:
            if attr in target_app_json:
                target_app.set(attr, target_app_json[attr])

        # Build TargetApp.Arguments only if value present
        args_val = target_app_json.get('argumentsValue')
        if args_val:
            args_type = target_app_json.get('argumentsType', 'x:String')
            args_wrapper = ET.SubElement(target_app, get_ns_tag('uix', 'TargetApp.Arguments'))
            in_arg = ET.SubElement(args_wrapper, get_ns_tag('', 'InArgument'))
            in_arg.set(get_ns_tag('x', 'TypeArguments'), args_type)
            in_arg.text = args_val

        # Build TargetApp.WorkingDirectory only if value present
        wd_val = target_app_json.get('workingDirectoryValue')
        if wd_val:
            wd_type = target_app_json.get('workingDirectoryType', 'x:String')
            wd_wrapper = ET.SubElement(target_app, get_ns_tag('uix', 'TargetApp.WorkingDirectory'))
            wd_in_arg = ET.SubElement(wd_wrapper, get_ns_tag('', 'InArgument'))
            wd_in_arg.set(get_ns_tag('x', 'TypeArguments'), wd_type)
            wd_in_arg.text = wd_val

        return target_app


# =============================================================================
# Switch Handler
# =============================================================================

class SwitchHandler(ActivityHandler):
    """Handler for Switch activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Switch element into JSON structure."""
        result = {
            'type': 'Switch',
            'displayName': element.get('DisplayName', ''),
            'typeArguments': '',
            'expression': '',
            'default': None,
            'cases': [],
        }

        # Extract TypeArguments
        type_args_attr = get_ns_tag('x', 'TypeArguments')
        if type_args_attr in element.attrib:
            result['typeArguments'] = TypeMapper.xaml_to_json_type(canonicalize_type(element.get(type_args_attr)))

        # Extract Expression
        result['expression'] = unescape_expression(element.get('Expression', ''))

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse Switch.Default
        default_tag = get_ns_tag('', 'Switch.Default')
        default_elem = element.find(default_tag)
        if default_elem is not None and len(default_elem) > 0:
            result['default'] = parse_activity(default_elem[0])

        # Parse keyed cases (activities with x:Key attribute)
        key_attr = get_ns_tag('x', 'Key')
        for child in element:
            _, local = parse_tag(child.tag)
            if local == 'Switch.Default':
                continue
            if child.tag.endswith('.ViewState'):
                continue

            # Check if this is a keyed case
            if key_attr in child.attrib:
                case_key = child.get(key_attr)
                case_activity = parse_activity(child)
                if case_activity:
                    result['cases'].append({
                        'key': case_key,
                        'activity': case_activity,
                    })

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Switch element from JSON structure."""
        switch_elem = ET.Element(get_ns_tag('', 'Switch'))

        # Set TypeArguments
        type_args = activity_json.get('typeArguments', 'String')
        switch_elem.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(type_args))

        # Set DisplayName
        if activity_json.get('displayName'):
            switch_elem.set('DisplayName', activity_json['displayName'])

        # Set Expression
        switch_elem.set('Expression', activity_json.get('expression', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Switch', '497,354'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        switch_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Switch')
        switch_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add Switch.Default
        if activity_json.get('default'):
            default_elem = ET.SubElement(switch_elem, get_ns_tag('', 'Switch.Default'))
            default_activity = build_activity(activity_json['default'], id_gen)
            if default_activity is not None:
                default_elem.append(default_activity)

        # Add keyed cases
        for case_info in activity_json.get('cases', []):
            case_activity = build_activity(case_info['activity'], id_gen)
            if case_activity is not None:
                case_activity.set(get_ns_tag('x', 'Key'), case_info['key'])
                switch_elem.append(case_activity)

        return switch_elem


# =============================================================================
# TryCatch Handler
# =============================================================================

class TryCatchHandler(ActivityHandler):
    """Handler for TryCatch activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse TryCatch element into JSON structure."""
        result = {
            'type': 'TryCatch',
            'displayName': element.get('DisplayName', ''),
            'try': None,
            'catches': [],
            'finally': None,
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse TryCatch.Try
        try_tag = get_ns_tag('', 'TryCatch.Try')
        try_elem = element.find(try_tag)
        if try_elem is not None and len(try_elem) > 0:
            result['try'] = parse_activity(try_elem[0])

        # Parse TryCatch.Catches
        catches_tag = get_ns_tag('', 'TryCatch.Catches')
        catches_elem = element.find(catches_tag)
        if catches_elem is not None:
            for catch_elem in catches_elem:
                catch_info = self._parse_catch(catch_elem)
                if catch_info:
                    result['catches'].append(catch_info)

        # Parse TryCatch.Finally
        finally_tag = get_ns_tag('', 'TryCatch.Finally')
        finally_elem = element.find(finally_tag)
        if finally_elem is not None and len(finally_elem) > 0:
            result['finally'] = parse_activity(finally_elem[0])

        return result

    def _parse_catch(self, catch_elem: ET.Element) -> Optional[Dict[str, Any]]:
        """Parse a Catch element."""
        result = {
            'exceptionType': 'Exception',
            'variableName': 'ex',
            'handler': None,
            'hintSize': None,
            'idRef': None,
        }

        # Extract exception type from TypeArguments
        type_args_attr = get_ns_tag('x', 'TypeArguments')
        if type_args_attr in catch_elem.attrib:
            result['exceptionType'] = TypeMapper.xaml_to_json_type(canonicalize_type(catch_elem.get(type_args_attr)))

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in catch_elem.attrib:
            result['hintSize'] = catch_elem.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in catch_elem.attrib:
            result['idRef'] = catch_elem.get(id_ref_attr)

        # Parse ViewState for Catch
        viewstate = ViewStateBuilder.parse_viewstate(catch_elem)
        if viewstate:
            result['viewState'] = viewstate

        # Find ActivityAction
        for child in catch_elem:
            _, local = parse_tag(child.tag)
            if local == 'ActivityAction':
                action_info = ActivityActionParser.parse_activity_action(child)
                result['variableName'] = action_info.get('variableName', 'ex')
                result['handler'] = action_info.get('activity')
                break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build TryCatch element from JSON structure."""
        trycatch_elem = ET.Element(get_ns_tag('', 'TryCatch'))

        # Set DisplayName
        if activity_json.get('displayName'):
            trycatch_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('TryCatch', '456,713'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        trycatch_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('TryCatch')
        trycatch_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add TryCatch.Try
        if activity_json.get('try'):
            try_elem = ET.SubElement(trycatch_elem, get_ns_tag('', 'TryCatch.Try'))
            try_activity = build_activity(activity_json['try'], id_gen)
            if try_activity is not None:
                try_elem.append(try_activity)

        # Add TryCatch.Catches
        if activity_json.get('catches'):
            catches_elem = ET.SubElement(trycatch_elem, get_ns_tag('', 'TryCatch.Catches'))
            for catch_info in activity_json['catches']:
                self._build_catch(catches_elem, catch_info, id_gen)

        # Add TryCatch.Finally
        if activity_json.get('finally'):
            finally_elem = ET.SubElement(trycatch_elem, get_ns_tag('', 'TryCatch.Finally'))
            finally_activity = build_activity(activity_json['finally'], id_gen)
            if finally_activity is not None:
                finally_elem.append(finally_activity)

        return trycatch_elem

    def _build_catch(self, parent: ET.Element, catch_info: Dict[str, Any], id_gen: IdRefGenerator):
        """Build a Catch element."""
        catch_elem = ET.SubElement(parent, get_ns_tag('', 'Catch'))

        # Set exception type
        exc_type = catch_info.get('exceptionType', 'Exception')
        xaml_exc_type = TypeMapper.json_to_xaml_type(exc_type)
        catch_elem.set(get_ns_tag('x', 'TypeArguments'), xaml_exc_type)

        # Set HintSize
        hint_size = catch_info.get('hintSize', DEFAULT_HINT_SIZES.get('Catch', '422,528'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        catch_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = catch_info.get('idRef') or id_gen.generate('Catch')
        catch_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ViewState for Catch
        viewstate = catch_info.get('viewState', {'IsExpanded': True, 'IsPinned': False})
        viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
        catch_elem.append(viewstate_elem)

        # Create ActivityAction with DelegateInArgument
        activity_action = ET.SubElement(catch_elem, get_ns_tag('', 'ActivityAction'))
        activity_action.set(get_ns_tag('x', 'TypeArguments'), xaml_exc_type)

        # Add DelegateInArgument
        arg_wrapper = ET.SubElement(activity_action, get_ns_tag('', 'ActivityAction.Argument'))
        delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
        delegate.set(get_ns_tag('x', 'TypeArguments'), xaml_exc_type)
        delegate.set('Name', catch_info.get('variableName', 'ex'))

        # Build handler activity
        if catch_info.get('handler'):
            handler_elem = build_activity(catch_info['handler'], id_gen)
            if handler_elem is not None:
                activity_action.append(handler_elem)


# =============================================================================
# ForEach Handler
# =============================================================================

class ForEachHandler(ActivityHandler):
    """Handler for ForEach activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ForEach element into JSON structure."""
        result = {
            'type': 'ForEach',
            'displayName': element.get('DisplayName', ''),
            'typeArguments': '',
            'values': '',
            'currentIndex': None,
            'body': None,
        }

        # Extract TypeArguments
        type_args_attr = get_ns_tag('x', 'TypeArguments')
        if type_args_attr in element.attrib:
            result['typeArguments'] = TypeMapper.xaml_to_json_type(canonicalize_type(element.get(type_args_attr)))

        # Extract Values
        result['values'] = unescape_expression(element.get('Values', ''))

        # Extract CurrentIndex (may be {x:Null} or an expression)
        current_index = element.get('CurrentIndex', '')
        if current_index and current_index != '{x:Null}':
            result['currentIndex'] = unescape_expression(current_index)

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ForEach.Body
        body_tag = get_ns_tag('ui', 'ForEach.Body')
        body_elem = element.find(body_tag)
        if body_elem is not None:
            # Find ActivityAction
            for child in body_elem:
                _, local = parse_tag(child.tag)
                if local == 'ActivityAction':
                    action_info = ActivityActionParser.parse_activity_action(child)
                    result['body'] = action_info
                    break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ForEach element from JSON structure."""
        foreach_elem = ET.Element(get_ns_tag('ui', 'ForEach'))

        # Set TypeArguments
        type_args = activity_json.get('typeArguments', 'String')
        foreach_elem.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(type_args))

        # Set DisplayName
        if activity_json.get('displayName'):
            foreach_elem.set('DisplayName', activity_json['displayName'])

        # Set CurrentIndex
        current_index = activity_json.get('currentIndex')
        if current_index:
            foreach_elem.set('CurrentIndex', current_index)
        else:
            foreach_elem.set('CurrentIndex', '{x:Null}')

        # Set Values
        foreach_elem.set('Values', activity_json.get('values', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ForEach', '518,1098'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        foreach_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ForEach')
        foreach_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ForEach.Body
        body_json = activity_json.get('body')
        if body_json:
            # Body format detection: Accept both Format A (wrapper) and Format B (direct)
            # Format A: {variableName, variableType, activity: {type, ...}}
            # Format B: {type, displayName, ...}
            if 'type' in body_json and 'activity' not in body_json:
                body_json = {
                    'variableName': 'item',
                    'variableType': type_args,
                    'activity': body_json
                }

            body_wrapper = ET.SubElement(foreach_elem, get_ns_tag('ui', 'ForEach.Body'))

            # Create ActivityAction
            activity_action = ET.SubElement(body_wrapper, get_ns_tag('', 'ActivityAction'))
            activity_action.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(type_args))

            # Add DelegateInArgument
            arg_wrapper = ET.SubElement(activity_action, get_ns_tag('', 'ActivityAction.Argument'))
            delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
            delegate.set(get_ns_tag('x', 'TypeArguments'), TypeMapper.json_to_xaml_type(type_args))
            delegate.set('Name', body_json.get('variableName', 'item'))

            # Build nested activity
            if body_json.get('activity'):
                activity_elem = build_activity(body_json['activity'], id_gen)
                if activity_elem is not None:
                    activity_action.append(activity_elem)

        return foreach_elem


# =============================================================================
# ForEachRow Handler
# =============================================================================

class ForEachRowHandler(ActivityHandler):
    """Handler for ForEachRow activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ForEachRow element into JSON structure."""
        result = {
            'type': 'ForEachRow',
            'displayName': element.get('DisplayName', ''),
            'dataTable': '',
            'columnNames': None,
            'currentIndex': None,
            'body': None,
        }

        # Extract DataTable
        result['dataTable'] = unescape_expression(element.get('DataTable', ''))

        # Extract ColumnNames (may be {x:Null})
        column_names = element.get('ColumnNames', '')
        if column_names and column_names != '{x:Null}':
            result['columnNames'] = unescape_expression(column_names)

        # Extract CurrentIndex
        current_index = element.get('CurrentIndex', '')
        if current_index and current_index != '{x:Null}':
            result['currentIndex'] = unescape_expression(current_index)

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ForEachRow.Body
        body_tag = get_ns_tag('ui', 'ForEachRow.Body')
        body_elem = element.find(body_tag)
        if body_elem is not None:
            for child in body_elem:
                _, local = parse_tag(child.tag)
                if local == 'ActivityAction':
                    action_info = ActivityActionParser.parse_activity_action(child)
                    result['body'] = action_info
                    break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ForEachRow element from JSON structure."""
        foreach_elem = ET.Element(get_ns_tag('ui', 'ForEachRow'))

        # Set DisplayName
        if activity_json.get('displayName'):
            foreach_elem.set('DisplayName', activity_json['displayName'])

        # Set DataTable
        foreach_elem.set('DataTable', activity_json.get('dataTable', ''))

        # Set ColumnNames
        column_names = activity_json.get('columnNames')
        if column_names:
            foreach_elem.set('ColumnNames', column_names)
        else:
            foreach_elem.set('ColumnNames', '{x:Null}')

        # Set CurrentIndex
        current_index = activity_json.get('currentIndex')
        if current_index:
            foreach_elem.set('CurrentIndex', current_index)

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ForEachRow', '502,1167'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        foreach_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ForEachRow')
        foreach_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ForEachRow.Body
        body_json = activity_json.get('body')
        if body_json:
            # Body format detection: Accept both Format A (wrapper) and Format B (direct)
            # Format A: {variableName, variableType, activity: {type, ...}}
            # Format B: {type, displayName, ...}
            if 'type' in body_json and 'activity' not in body_json:
                body_json = {
                    'variableName': 'row',
                    'variableType': 'DataRow',
                    'activity': body_json
                }

            body_wrapper = ET.SubElement(foreach_elem, get_ns_tag('ui', 'ForEachRow.Body'))

            # Create ActivityAction with DataRow type
            activity_action = ET.SubElement(body_wrapper, get_ns_tag('', 'ActivityAction'))
            activity_action.set(get_ns_tag('x', 'TypeArguments'), 'sd:DataRow')

            # Add DelegateInArgument
            arg_wrapper = ET.SubElement(activity_action, get_ns_tag('', 'ActivityAction.Argument'))
            delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
            delegate.set(get_ns_tag('x', 'TypeArguments'), 'sd:DataRow')
            delegate.set('Name', body_json.get('variableName', 'row'))

            # Build nested activity
            if body_json.get('activity'):
                activity_elem = build_activity(body_json['activity'], id_gen)
                if activity_elem is not None:
                    activity_action.append(activity_elem)

        return foreach_elem


# =============================================================================
# Rethrow Handler
# =============================================================================

class RethrowHandler(ActivityHandler):
    """Handler for Rethrow activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Rethrow element into JSON structure."""
        result = {
            'type': 'Rethrow',
            'displayName': element.get('DisplayName', ''),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Rethrow element from JSON structure."""
        rethrow_elem = ET.Element(get_ns_tag('', 'Rethrow'))

        # Set DisplayName
        if activity_json.get('displayName'):
            rethrow_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Rethrow', '382,48'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        rethrow_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Rethrow')
        rethrow_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return rethrow_elem


class ThrowHandler(ActivityHandler):
    """Handler for Throw activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Throw element into JSON structure."""
        result = {
            'type': 'Throw',
            'displayName': element.get('DisplayName', ''),
            'exception': unescape_expression(element.get('Exception', '')),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Throw element from JSON structure."""
        throw_elem = ET.Element(get_ns_tag('', 'Throw'))

        # Set DisplayName
        if activity_json.get('displayName'):
            throw_elem.set('DisplayName', activity_json['displayName'])

        # Set Exception
        throw_elem.set('Exception', activity_json.get('exception', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Throw', '382,48'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        throw_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Throw')
        throw_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return throw_elem


class ReturnHandler(ActivityHandler):
    """Handler for Return activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Return element into JSON structure."""
        result = {
            'type': 'Return',
            'displayName': element.get('DisplayName', ''),
        }

        # Parse Result (OutArgument)
        result_tag = get_ns_tag('ui', 'Return.Result')
        result_elem = element.find(result_tag)
        if result_elem is not None:
            out_arg = result_elem.find(get_ns_tag('', 'OutArgument'))
            if out_arg is not None:
                result['result'] = {
                    'outArgument': {
                        'x:TypeArguments': canonicalize_type(out_arg.get(get_ns_tag('x', 'TypeArguments'), 'x:Object')),
                        'value': unescape_expression(out_arg.text or ''),
                    }
                }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Return element from JSON structure."""
        return_elem = ET.Element(get_ns_tag('ui', 'Return'))

        # Set DisplayName
        if activity_json.get('displayName'):
            return_elem.set('DisplayName', activity_json['displayName'])

        # Build Result (OutArgument)
        result_info = activity_json.get('result')
        if result_info and result_info.get('outArgument'):
            result_wrapper = ET.SubElement(return_elem, get_ns_tag('ui', 'Return.Result'))
            out_arg_info = result_info['outArgument']
            out_arg = ET.SubElement(result_wrapper, get_ns_tag('', 'OutArgument'))
            out_arg.set(get_ns_tag('x', 'TypeArguments'), out_arg_info.get('x:TypeArguments', 'x:Object'))
            out_arg.text = out_arg_info.get('value', '')

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Return', '262,60'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        return_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Return')
        return_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return return_elem


# =============================================================================
# Excel Activity Handlers
# =============================================================================

class ExcelProcessScopeXHandler(ActivityHandler):
    """Handler for ExcelProcessScopeX activities."""

    # Attributes that are typically {x:Null}
    NULL_ATTRS = [
        'DisplayAlerts', 'ExistingProcessAction', 'FileConflictResolution',
        'LaunchMethod', 'LaunchTimeout', 'MacroSettings', 'ProcessMode', 'ShowExcelWindow'
    ]

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ExcelProcessScopeX element into JSON structure."""
        result = {
            'type': 'ExcelProcessScopeX',
            'displayName': element.get('DisplayName', ''),
            'processTagName': 'ExcelProcessScopeTag',
            'body': None,
        }

        # Extract optional attributes
        for attr in self.NULL_ATTRS:
            val = element.get(attr)
            if val and val != '{x:Null}':
                result[attr.lower()] = val

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ExcelProcessScopeX.Body
        body_tag = get_ns_tag('ueab', 'ExcelProcessScopeX.Body')
        body_elem = element.find(body_tag)
        if body_elem is not None:
            for child in body_elem:
                _, local = parse_tag(child.tag)
                if local == 'ActivityAction':
                    action_info = ActivityActionParser.parse_activity_action(child)
                    result['processTagName'] = action_info.get('variableName', 'ExcelProcessScopeTag')
                    result['body'] = action_info.get('activity')
                    break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ExcelProcessScopeX element from JSON structure."""
        scope_elem = ET.Element(get_ns_tag('ueab', 'ExcelProcessScopeX'))

        # Set optional attributes - use parsed value if present, otherwise {x:Null}
        for attr in self.NULL_ATTRS:
            json_key = attr[0].lower() + attr[1:]  # Convert to camelCase for JSON lookup
            val = activity_json.get(json_key) or activity_json.get(attr.lower())
            if val:
                scope_elem.set(attr, val)
            else:
                scope_elem.set(attr, '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            scope_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ExcelProcessScopeX', '580,1701'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        scope_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ExcelProcessScopeX')
        scope_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ExcelProcessScopeX.Body
        body_wrapper = ET.SubElement(scope_elem, get_ns_tag('ueab', 'ExcelProcessScopeX.Body'))

        # Create ActivityAction with IExcelProcess type
        activity_action = ET.SubElement(body_wrapper, get_ns_tag('', 'ActivityAction'))
        activity_action.set(get_ns_tag('x', 'TypeArguments'), 'ui:IExcelProcess')

        # Add DelegateInArgument
        arg_wrapper = ET.SubElement(activity_action, get_ns_tag('', 'ActivityAction.Argument'))
        delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
        delegate.set(get_ns_tag('x', 'TypeArguments'), 'ui:IExcelProcess')
        delegate.set('Name', activity_json.get('processTagName', 'ExcelProcessScopeTag'))

        # Body format detection: Accept both Format A (wrapper) and Format B (direct)
        # Format A: {variableName, variableType, activity: {type, ...}}
        # Format B: {type, displayName, ...}
        body_json = activity_json.get('body')
        if body_json:
            if 'activity' in body_json and isinstance(body_json['activity'], dict):
                activity_to_build = body_json['activity']
            elif 'type' in body_json:
                activity_to_build = body_json
            else:
                activity_to_build = None

            if activity_to_build:
                activity_elem = build_activity(activity_to_build, id_gen)
                if activity_elem is not None:
                    activity_action.append(activity_elem)

        return scope_elem


class ExcelApplicationCardHandler(ActivityHandler):
    """Handler for ExcelApplicationCard activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ExcelApplicationCard element into JSON structure."""
        result = {
            'type': 'ExcelApplicationCard',
            'displayName': element.get('DisplayName', ''),
            'workbookPath': unescape_expression(element.get('WorkbookPath', '')),
            'autoSave': element.get('AutoSave', 'False').lower() == 'true',
            'createNewFile': element.get('CreateNewFile', 'False').lower() == 'true',
            'keepExcelFileOpen': element.get('KeepExcelFileOpen', 'True').lower() == 'true',
            'resizeWindow': element.get('ResizeWindow', 'None'),
            'sensitivityOperation': element.get('SensitivityOperation', 'None'),
            'excelHandleName': 'Excel',
            'body': None,
        }

        # Extract optional {x:Null} attributes
        for attr in ['Password', 'ReadFormatting', 'SensitivityLabel']:
            val = element.get(attr)
            if val and val != '{x:Null}':
                result[attr.lower()] = val

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ExcelApplicationCard.Body
        body_tag = get_ns_tag('ueab', 'ExcelApplicationCard.Body')
        body_elem = element.find(body_tag)
        if body_elem is not None:
            for child in body_elem:
                _, local = parse_tag(child.tag)
                if local == 'ActivityAction':
                    action_info = ActivityActionParser.parse_activity_action(child)
                    result['excelHandleName'] = action_info.get('variableName', 'Excel')
                    result['body'] = action_info.get('activity')
                    break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ExcelApplicationCard element from JSON structure."""
        card_elem = ET.Element(get_ns_tag('ueab', 'ExcelApplicationCard'))

        # Set optional nullable attributes - use parsed value if present, otherwise {x:Null}
        for attr in ['Password', 'ReadFormatting', 'SensitivityLabel']:
            val = activity_json.get(attr.lower())
            if val:
                card_elem.set(attr, val)
            else:
                card_elem.set(attr, '{x:Null}')

        # Set boolean attributes
        card_elem.set('AutoSave', str(activity_json.get('autoSave', False)))
        card_elem.set('CreateNewFile', str(activity_json.get('createNewFile', False)))
        card_elem.set('KeepExcelFileOpen', str(activity_json.get('keepExcelFileOpen', True)))
        card_elem.set('ResizeWindow', activity_json.get('resizeWindow', 'None'))
        card_elem.set('SensitivityOperation', activity_json.get('sensitivityOperation', 'None'))

        # Set DisplayName
        if activity_json.get('displayName'):
            card_elem.set('DisplayName', activity_json['displayName'])

        # Set WorkbookPath
        card_elem.set('WorkbookPath', activity_json.get('workbookPath', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ExcelApplicationCard', '512,1522'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        card_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ExcelApplicationCard')
        card_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ExcelApplicationCard.Body
        body_wrapper = ET.SubElement(card_elem, get_ns_tag('ueab', 'ExcelApplicationCard.Body'))

        # Create ActivityAction with IWorkbookQuickHandle type
        activity_action = ET.SubElement(body_wrapper, get_ns_tag('', 'ActivityAction'))
        activity_action.set(get_ns_tag('x', 'TypeArguments'), 'ue:IWorkbookQuickHandle')

        # Add DelegateInArgument
        arg_wrapper = ET.SubElement(activity_action, get_ns_tag('', 'ActivityAction.Argument'))
        delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
        delegate.set(get_ns_tag('x', 'TypeArguments'), 'ue:IWorkbookQuickHandle')
        delegate.set('Name', activity_json.get('excelHandleName', 'Excel'))

        # Body format detection: Accept both Format A (wrapper) and Format B (direct)
        # Format A: {variableName, variableType, activity: {type, ...}}
        # Format B: {type, displayName, ...}
        body_json = activity_json.get('body')
        if body_json:
            if 'activity' in body_json and isinstance(body_json['activity'], dict):
                activity_to_build = body_json['activity']
            elif 'type' in body_json:
                activity_to_build = body_json
            else:
                activity_to_build = None

            if activity_to_build:
                activity_elem = build_activity(activity_to_build, id_gen)
                if activity_elem is not None:
                    activity_action.append(activity_elem)

        return card_elem


class ReadRangeXHandler(ActivityHandler):
    """Handler for ReadRangeX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ReadRangeX element into JSON structure."""
        result = {
            'type': 'ReadRangeX',
            'displayName': element.get('DisplayName', ''),
            'range': unescape_expression(element.get('Range', '')),
            'saveTo': unescape_expression(element.get('SaveTo', '')),
            'hasHeaders': element.get('HasHeaders', 'False').lower() == 'true',
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ReadRangeX element from JSON structure."""
        read_elem = ET.Element(get_ns_tag('ueab', 'ReadRangeX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            read_elem.set('DisplayName', activity_json['displayName'])

        # Set Range
        read_elem.set('Range', activity_json.get('range', ''))

        # Set SaveTo
        read_elem.set('SaveTo', activity_json.get('saveTo', ''))

        # Set HasHeaders
        if activity_json.get('hasHeaders'):
            read_elem.set('HasHeaders', 'True')

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ReadRangeX', '444,201'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        read_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ReadRangeX')
        read_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return read_elem


class SaveExcelFileXHandler(ActivityHandler):
    """Handler for SaveExcelFileX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse SaveExcelFileX element into JSON structure."""
        result = {
            'type': 'SaveExcelFileX',
            'displayName': element.get('DisplayName', ''),
            'workbook': unescape_expression(element.get('Workbook', '')),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build SaveExcelFileX element from JSON structure."""
        save_elem = ET.Element(get_ns_tag('ueab', 'SaveExcelFileX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            save_elem.set('DisplayName', activity_json['displayName'])

        # Set Workbook
        save_elem.set('Workbook', activity_json.get('workbook', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('SaveExcelFileX', '444,108'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        save_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('SaveExcelFileX')
        save_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return save_elem


class WriteCellXHandler(ActivityHandler):
    """Handler for WriteCellX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse WriteCellX element into JSON structure."""
        result = {
            'type': 'WriteCellX',
            'displayName': element.get('DisplayName', ''),
            'cell': unescape_expression(element.get('Cell', '')),
            'value': unescape_expression(element.get('Value', '')),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build WriteCellX element from JSON structure."""
        write_elem = ET.Element(get_ns_tag('ueab', 'WriteCellX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            write_elem.set('DisplayName', activity_json['displayName'])

        # Set Cell
        write_elem.set('Cell', activity_json.get('cell', ''))

        # Set Value
        write_elem.set('Value', activity_json.get('value', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('WriteCellX', '444,191'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        write_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('WriteCellX')
        write_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return write_elem


class WriteRangeXHandler(ActivityHandler):
    """Handler for WriteRangeX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse WriteRangeX element into JSON structure."""
        result = {
            'type': 'WriteRangeX',
            'displayName': element.get('DisplayName', ''),
            'destination': unescape_expression(element.get('Destination', '')),
            'source': unescape_expression(element.get('Source', '')),
            'excludeHeaders': element.get('ExcludeHeaders', 'True').lower() == 'true',
            'ignoreEmptySource': element.get('IgnoreEmptySource', 'False').lower() == 'true',
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build WriteRangeX element from JSON structure."""
        write_elem = ET.Element(get_ns_tag('ueab', 'WriteRangeX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            write_elem.set('DisplayName', activity_json['displayName'])

        # Set Destination
        write_elem.set('Destination', activity_json.get('destination', ''))

        # Set Source
        write_elem.set('Source', activity_json.get('source', ''))

        # Set ExcludeHeaders
        write_elem.set('ExcludeHeaders', str(activity_json.get('excludeHeaders', True)))

        # Set IgnoreEmptySource
        if activity_json.get('ignoreEmptySource'):
            write_elem.set('IgnoreEmptySource', 'True')

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('WriteRangeX', '444,191'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        write_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('WriteRangeX')
        write_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return write_elem


class CopyPasteRangeXHandler(ActivityHandler):
    """Handler for CopyPasteRangeX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse CopyPasteRangeX element into JSON structure."""
        result = {
            'type': 'CopyPasteRangeX',
            'displayName': element.get('DisplayName', ''),
            'sourceRange': unescape_expression(element.get('SourceRange', '')),
            'destinationRange': unescape_expression(element.get('DestinationRange', '')),
            'pasteOptions': element.get('PasteOptions', 'All'),
            'transpose': element.get('Transpose', 'False').lower() == 'true',
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build CopyPasteRangeX element from JSON structure."""
        copy_elem = ET.Element(get_ns_tag('ueab', 'CopyPasteRangeX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            copy_elem.set('DisplayName', activity_json['displayName'])

        # Set SourceRange
        copy_elem.set('SourceRange', activity_json.get('sourceRange', ''))

        # Set DestinationRange
        copy_elem.set('DestinationRange', activity_json.get('destinationRange', ''))

        # Set PasteOptions
        copy_elem.set('PasteOptions', activity_json.get('pasteOptions', 'All'))

        # Set Transpose
        copy_elem.set('Transpose', str(activity_json.get('transpose', False)))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('CopyPasteRangeX', '444,272'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        copy_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('CopyPasteRangeX')
        copy_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return copy_elem


class ClearRangeXHandler(ActivityHandler):
    """Handler for ClearRangeX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ClearRangeX element into JSON structure."""
        result = {
            'type': 'ClearRangeX',
            'displayName': element.get('DisplayName', ''),
            'targetRange': unescape_expression(element.get('TargetRange', '')),
            'hasHeaders': element.get('HasHeaders', 'False').lower() == 'true',
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ClearRangeX element from JSON structure."""
        clear_elem = ET.Element(get_ns_tag('ueab', 'ClearRangeX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            clear_elem.set('DisplayName', activity_json['displayName'])

        # Set TargetRange
        clear_elem.set('TargetRange', activity_json.get('targetRange', ''))

        # Set HasHeaders
        clear_elem.set('HasHeaders', str(activity_json.get('hasHeaders', False)))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ClearRangeX', '444,191'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        clear_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ClearRangeX')
        clear_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return clear_elem


class FilterXHandler(ActivityHandler):
    """Handler for FilterX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse FilterX element into JSON structure."""
        result = {
            'type': 'FilterX',
            'displayName': element.get('DisplayName', ''),
            'range': unescape_expression(element.get('Range', '')),
            'columnName': unescape_expression(element.get('ColumnName', '')),
        }

        # Extract optional FilterArgument
        filter_arg = element.get('FilterArgument')
        if filter_arg and filter_arg != '{x:Null}':
            result['filterArgument'] = unescape_expression(filter_arg)

        # Extract ClearFilter
        clear_filter = element.get('ClearFilter', 'False')
        result['clearFilter'] = clear_filter.lower() == 'true'

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build FilterX element from JSON structure."""
        filter_elem = ET.Element(get_ns_tag('ueab', 'FilterX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            filter_elem.set('DisplayName', activity_json['displayName'])

        # Set Range
        filter_elem.set('Range', activity_json.get('range', ''))

        # Set ColumnName
        filter_elem.set('ColumnName', activity_json.get('columnName', ''))

        # Set FilterArgument if provided
        if 'filterArgument' in activity_json:
            filter_elem.set('FilterArgument', activity_json['filterArgument'])

        # Set ClearFilter
        filter_elem.set('ClearFilter', str(activity_json.get('clearFilter', False)))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('FilterX', '444,191'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        filter_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('FilterX')
        filter_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return filter_elem


class FindFirstLastDataRowXHandler(ActivityHandler):
    """Handler for FindFirstLastDataRowX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse FindFirstLastDataRowX element into JSON structure."""
        result = {
            'type': 'FindFirstLastDataRowX',
            'displayName': element.get('DisplayName', ''),
            'range': unescape_expression(element.get('Range', '')),
        }

        # Extract optional ColumnName
        col_name = element.get('ColumnName')
        if col_name and col_name != '{x:Null}':
            result['columnName'] = unescape_expression(col_name)

        # Extract optional FirstRowIndex
        first_row = element.get('FirstRowIndex')
        if first_row and first_row != '{x:Null}':
            result['firstRowIndex'] = unescape_expression(first_row)

        # Extract optional LastRowIndex
        last_row = element.get('LastRowIndex')
        if last_row and last_row != '{x:Null}':
            result['lastRowIndex'] = unescape_expression(last_row)

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build FindFirstLastDataRowX element from JSON structure."""
        find_elem = ET.Element(get_ns_tag('ueab', 'FindFirstLastDataRowX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            find_elem.set('DisplayName', activity_json['displayName'])

        # Set Range
        find_elem.set('Range', activity_json.get('range', ''))

        # Set optional ColumnName
        if 'columnName' in activity_json:
            find_elem.set('ColumnName', activity_json['columnName'])
        else:
            find_elem.set('ColumnName', '{x:Null}')

        # Set optional FirstRowIndex
        if 'firstRowIndex' in activity_json:
            find_elem.set('FirstRowIndex', activity_json['firstRowIndex'])
        else:
            find_elem.set('FirstRowIndex', '{x:Null}')

        # Set optional LastRowIndex
        if 'lastRowIndex' in activity_json:
            find_elem.set('LastRowIndex', activity_json['lastRowIndex'])
        else:
            find_elem.set('LastRowIndex', '{x:Null}')

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('FindFirstLastDataRowX', '444,150'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        find_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('FindFirstLastDataRowX')
        find_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return find_elem


# =============================================================================
# File Operation Handlers
# =============================================================================

class CreateDirectoryHandler(ActivityHandler):
    """Handler for CreateDirectory activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse CreateDirectory element into JSON structure."""
        result = {
            'type': 'CreateDirectory',
            'displayName': element.get('DisplayName', ''),
            'path': unescape_expression(element.get('Path', '')),
        }

        # Extract optional attributes
        cont_err = element.get('ContinueOnError')
        if cont_err and cont_err != '{x:Null}':
            result['continueOnError'] = cont_err.lower() == 'true'

        output = element.get('Output')
        if output and output != '{x:Null}':
            result['output'] = unescape_expression(output)

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build CreateDirectory element from JSON structure."""
        create_elem = ET.Element(get_ns_tag('ui', 'CreateDirectory'))

        # Set null attributes
        create_elem.set('ContinueOnError', '{x:Null}')
        create_elem.set('Output', '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            create_elem.set('DisplayName', activity_json['displayName'])

        # Set Path
        create_elem.set('Path', activity_json.get('path', ''))

        # Override ContinueOnError if set
        if 'continueOnError' in activity_json:
            create_elem.set('ContinueOnError', str(activity_json['continueOnError']))

        # Override Output if set
        if 'output' in activity_json:
            create_elem.set('Output', activity_json['output'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('CreateDirectory', '334,90'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        create_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('CreateDirectory')
        create_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return create_elem


class MoveFileHandler(ActivityHandler):
    """Handler for MoveFile activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse MoveFile element into JSON structure."""
        result = {
            'type': 'MoveFile',
            'displayName': element.get('DisplayName', ''),
            'path': unescape_expression(element.get('Path', '')),
            'destination': unescape_expression(element.get('Destination', '')),
            'overwrite': element.get('Overwrite', 'True').lower() == 'true',
        }

        # Extract ContinueOnError as boolean
        continue_on_error = element.get('ContinueOnError')
        if continue_on_error and continue_on_error != '{x:Null}':
            result['continueOnError'] = continue_on_error.lower() == 'true'

        # Extract optional {x:Null} string attributes
        for attr in ['PathResource', 'DestinationResource']:
            val = element.get(attr)
            if val and val != '{x:Null}':
                result[attr[0].lower() + attr[1:]] = val  # Use camelCase key

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build MoveFile element from JSON structure."""
        move_elem = ET.Element(get_ns_tag('ui', 'MoveFile'))

        # Handle ContinueOnError separately (it's a boolean)
        if 'continueOnError' in activity_json:
            move_elem.set('ContinueOnError', str(activity_json['continueOnError']))
        else:
            move_elem.set('ContinueOnError', '{x:Null}')

        # Set optional nullable string attributes - use parsed value if present, otherwise {x:Null}
        for attr in ['PathResource', 'DestinationResource']:
            json_key = attr[0].lower() + attr[1:]  # Convert to camelCase
            val = activity_json.get(json_key)
            if val:
                move_elem.set(attr, val)
            else:
                move_elem.set(attr, '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            move_elem.set('DisplayName', activity_json['displayName'])

        # Set Path
        move_elem.set('Path', activity_json.get('path', ''))

        # Set Destination
        move_elem.set('Destination', activity_json.get('destination', ''))

        # Set Overwrite
        move_elem.set('Overwrite', str(activity_json.get('overwrite', True)))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('MoveFile', '450,182'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        move_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('MoveFile')
        move_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return move_elem


class DeleteFileXHandler(ActivityHandler):
    """Handler for DeleteFileX activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse DeleteFileX element into JSON structure."""
        result = {
            'type': 'DeleteFileX',
            'displayName': element.get('DisplayName', ''),
            'path': unescape_expression(element.get('Path', '')),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build DeleteFileX element from JSON structure."""
        delete_elem = ET.Element(get_ns_tag('ui', 'DeleteFileX'))

        # Set DisplayName
        if activity_json.get('displayName'):
            delete_elem.set('DisplayName', activity_json['displayName'])

        # Set Path
        delete_elem.set('Path', activity_json.get('path', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('DeleteFileX', '382,48'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        delete_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('DeleteFileX')
        delete_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return delete_elem


class ReadTextFileHandler(ActivityHandler):
    """Handler for ReadTextFile activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ReadTextFile element into JSON structure."""
        result = {
            'type': 'ReadTextFile',
            'displayName': element.get('DisplayName', ''),
            'fileName': unescape_expression(element.get('FileName', '')),
            'content': unescape_expression(element.get('Content', '')),
        }

        # Extract optional File attribute
        file_attr = element.get('File')
        if file_attr and file_attr != '{x:Null}':
            result['file'] = unescape_expression(file_attr)

        # Extract optional Encoding
        encoding = element.get('Encoding')
        if encoding and encoding != '{x:Null}':
            result['encoding'] = encoding

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ReadTextFile element from JSON structure."""
        read_elem = ET.Element(get_ns_tag('ui', 'ReadTextFile'))

        # Set File to {x:Null} by default
        read_elem.set('File', '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            read_elem.set('DisplayName', activity_json['displayName'])

        # Set FileName
        read_elem.set('FileName', activity_json.get('fileName', ''))

        # Set Content
        read_elem.set('Content', activity_json.get('content', ''))

        # Set Encoding if provided
        if 'encoding' in activity_json:
            read_elem.set('Encoding', activity_json['encoding'])

        # Override File if provided
        if 'file' in activity_json:
            read_elem.set('File', activity_json['file'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ReadTextFile', '586,124'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        read_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ReadTextFile')
        read_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return read_elem


class ReadRangeHandler(ActivityHandler):
    """Handler for ReadRange (legacy Excel) activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse ReadRange element into JSON structure."""
        result = {
            'type': 'ReadRange',
            'displayName': element.get('DisplayName', ''),
            'workbookPath': unescape_expression(element.get('WorkbookPath', '')),
            'sheetName': element.get('SheetName', 'Sheet1'),
            'dataTable': unescape_expression(element.get('DataTable', '')),
            'addHeaders': element.get('AddHeaders', 'True').lower() == 'true',
        }

        # Extract optional range
        range_val = element.get('Range')
        if range_val and range_val != '{x:Null}':
            result['range'] = range_val

        # Extract optional {x:Null} attributes
        wb_resource = element.get('WorkbookPathResource')
        if wb_resource and wb_resource != '{x:Null}':
            result['workbookPathResource'] = wb_resource

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build ReadRange element from JSON structure."""
        read_elem = ET.Element(get_ns_tag('ui', 'ReadRange'))

        # Set null attributes
        read_elem.set('Range', '{x:Null}')
        read_elem.set('WorkbookPathResource', '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            read_elem.set('DisplayName', activity_json['displayName'])

        # Set WorkbookPath
        read_elem.set('WorkbookPath', activity_json.get('workbookPath', ''))

        # Set SheetName
        read_elem.set('SheetName', activity_json.get('sheetName', 'Sheet1'))

        # Set DataTable
        read_elem.set('DataTable', activity_json.get('dataTable', ''))

        # Set AddHeaders
        read_elem.set('AddHeaders', str(activity_json.get('addHeaders', True)))

        # Override Range if set
        if 'range' in activity_json:
            read_elem.set('Range', activity_json['range'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('ReadRange', '450,120'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        read_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('ReadRange')
        read_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return read_elem


# =============================================================================
# Data Activity Handlers
# =============================================================================

class AddDataRowHandler(ActivityHandler):
    """Handler for AddDataRow activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse AddDataRow element into JSON structure."""
        result = {
            'type': 'AddDataRow',
            'displayName': element.get('DisplayName', ''),
            'dataTable': unescape_expression(element.get('DataTable', '')),
        }

        # Extract optional DataRow attribute
        data_row = element.get('DataRow')
        if data_row and data_row != '{x:Null}':
            result['dataRow'] = unescape_expression(data_row)

        # Extract optional ArrayRow attribute
        array_row = element.get('ArrayRow')
        if array_row and array_row != '{x:Null}':
            result['arrayRow'] = unescape_expression(array_row)

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build AddDataRow element from JSON structure."""
        add_row_elem = ET.Element(get_ns_tag('ui', 'AddDataRow'))

        # Set DataRow to {x:Null} by default
        add_row_elem.set('DataRow', '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            add_row_elem.set('DisplayName', activity_json['displayName'])

        # Set DataTable
        add_row_elem.set('DataTable', activity_json.get('dataTable', ''))

        # Set ArrayRow if provided
        if 'arrayRow' in activity_json:
            add_row_elem.set('ArrayRow', activity_json['arrayRow'])

        # Override DataRow if provided
        if 'dataRow' in activity_json:
            add_row_elem.set('DataRow', activity_json['dataRow'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('AddDataRow', '334,186'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        add_row_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('AddDataRow')
        add_row_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return add_row_elem


class BuildDataTableHandler(ActivityHandler):
    """Handler for BuildDataTable activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse BuildDataTable element into JSON structure."""
        result = {
            'type': 'BuildDataTable',
            'displayName': element.get('DisplayName', ''),
            'dataTable': unescape_expression(element.get('DataTable', '')),
            'tableInfo': element.get('TableInfo', ''),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build BuildDataTable element from JSON structure."""
        build_elem = ET.Element(get_ns_tag('ui', 'BuildDataTable'))

        # Set DisplayName
        if activity_json.get('displayName'):
            build_elem.set('DisplayName', activity_json['displayName'])

        # Set DataTable
        build_elem.set('DataTable', activity_json.get('dataTable', ''))

        # Set TableInfo (already XML-encoded in JSON, set directly)
        build_elem.set('TableInfo', activity_json.get('tableInfo', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('BuildDataTable', '586,92'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        build_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('BuildDataTable')
        build_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return build_elem


# =============================================================================
# Utility Activity Handlers
# =============================================================================

class CommentOutHandler(ActivityHandler):
    """Handler for CommentOut activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse CommentOut element into JSON structure."""
        result = {
            'type': 'CommentOut',
            'displayName': element.get('DisplayName', ''),
            'body': None,
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse CommentOut.Body
        body_tag = get_ns_tag('ui', 'CommentOut.Body')
        body_elem = element.find(body_tag)
        if body_elem is not None and len(body_elem) > 0:
            result['body'] = parse_activity(body_elem[0])

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build CommentOut element from JSON structure."""
        comment_elem = ET.Element(get_ns_tag('ui', 'CommentOut'))

        # Set DisplayName
        if activity_json.get('displayName'):
            comment_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('CommentOut', '580,84'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        comment_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('CommentOut')
        comment_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add CommentOut.Body
        if activity_json.get('body'):
            body_wrapper = ET.SubElement(comment_elem, get_ns_tag('ui', 'CommentOut.Body'))
            body_activity = build_activity(activity_json['body'], id_gen)
            if body_activity is not None:
                body_wrapper.append(body_activity)

        # Add ViewState
        viewstate = activity_json.get('viewState', {'IsExpanded': False, 'IsPinned': False})
        viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
        comment_elem.append(viewstate_elem)

        return comment_elem


class RetryScopeHandler(ActivityHandler):
    """Handler for RetryScope activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse RetryScope element into JSON structure."""
        result = {
            'type': 'RetryScope',
            'displayName': element.get('DisplayName', ''),
            'numberOfRetries': int(element.get('NumberOfRetries', '2')),
            'logRetriedExceptions': element.get('LogRetriedExceptions', 'True').lower() == 'true',
            'retriedExceptionsLogLevel': element.get('RetriedExceptionsLogLevel', 'Info'),
            'activityBody': None,
            'condition': None,  # Stores parsed condition content
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse RetryScope.ActivityBody
        body_tag = get_ns_tag('ui', 'RetryScope.ActivityBody')
        body_elem = element.find(body_tag)
        if body_elem is not None:
            for child in body_elem:
                _, local = parse_tag(child.tag)
                if local == 'ActivityAction':
                    # For RetryScope, ActivityAction has no TypeArguments
                    for activity_child in child:
                        activity = parse_activity(activity_child)
                        if activity:
                            result['activityBody'] = activity
                            break
                    break

        # Parse RetryScope.Condition (ActivityFunc with TypeArguments x:Boolean)
        condition_tag = get_ns_tag('ui', 'RetryScope.Condition')
        condition_elem = element.find(condition_tag)
        if condition_elem is not None:
            for child in condition_elem:
                _, local = parse_tag(child.tag)
                if local == 'ActivityFunc':
                    # Check for Result variable name (DelegateOutArgument)
                    condition_info = {'typeArguments': 'x:Boolean'}

                    # Parse ActivityFunc.Result (DelegateOutArgument)
                    result_tag = get_ns_tag('', 'ActivityFunc.Result')
                    result_elem = child.find(result_tag)
                    if result_elem is not None:
                        for delegate in result_elem:
                            _, delegate_local = parse_tag(delegate.tag)
                            if delegate_local == 'DelegateOutArgument':
                                type_args = canonicalize_type(delegate.get(get_ns_tag('x', 'TypeArguments'), 'x:Boolean'))
                                name = delegate.get('Name', '')
                                condition_info['resultVariable'] = name
                                condition_info['resultType'] = type_args
                                break

                    # Parse any child activity within the ActivityFunc (the actual condition logic)
                    for activity_child in child:
                        _, child_local = parse_tag(activity_child.tag)
                        if child_local not in ['ActivityFunc.Result']:
                            activity = parse_activity(activity_child)
                            if activity:
                                condition_info['activity'] = activity
                                break

                    # Only store condition if it has meaningful content
                    if condition_info.get('resultVariable') or condition_info.get('activity'):
                        result['condition'] = condition_info
                    break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build RetryScope element from JSON structure."""
        retry_elem = ET.Element(get_ns_tag('ui', 'RetryScope'))

        # Set DisplayName
        if activity_json.get('displayName'):
            retry_elem.set('DisplayName', activity_json['displayName'])

        # Set NumberOfRetries
        retry_elem.set('NumberOfRetries', str(activity_json.get('numberOfRetries', 2)))

        # Set LogRetriedExceptions
        retry_elem.set('LogRetriedExceptions', str(activity_json.get('logRetriedExceptions', True)))

        # Set RetriedExceptionsLogLevel
        retry_elem.set('RetriedExceptionsLogLevel', activity_json.get('retriedExceptionsLogLevel', 'Info'))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('RetryScope', '580,1800'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        retry_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('RetryScope')
        retry_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add RetryScope.ActivityBody
        body_wrapper = ET.SubElement(retry_elem, get_ns_tag('ui', 'RetryScope.ActivityBody'))
        activity_action = ET.SubElement(body_wrapper, get_ns_tag('', 'ActivityAction'))

        # Build nested activity
        if activity_json.get('activityBody'):
            activity_elem = build_activity(activity_json['activityBody'], id_gen)
            if activity_elem is not None:
                activity_action.append(activity_elem)

        # Add RetryScope.Condition
        condition_wrapper = ET.SubElement(retry_elem, get_ns_tag('ui', 'RetryScope.Condition'))
        activity_func = ET.SubElement(condition_wrapper, get_ns_tag('', 'ActivityFunc'))
        activity_func.set(get_ns_tag('x', 'TypeArguments'), 'x:Boolean')

        # Build condition content from parsed data
        condition_info = activity_json.get('condition')
        if condition_info:
            # Add DelegateOutArgument for result if present
            if condition_info.get('resultVariable'):
                result_wrapper = ET.SubElement(activity_func, get_ns_tag('', 'ActivityFunc.Result'))
                delegate = ET.SubElement(result_wrapper, get_ns_tag('', 'DelegateOutArgument'))
                delegate.set(get_ns_tag('x', 'TypeArguments'), condition_info.get('resultType', 'x:Boolean'))
                delegate.set('Name', condition_info['resultVariable'])

            # Build condition activity if present
            if condition_info.get('activity'):
                cond_activity_elem = build_activity(condition_info['activity'], id_gen)
                if cond_activity_elem is not None:
                    activity_func.append(cond_activity_elem)

        return retry_elem


# =============================================================================
# While Handler
# =============================================================================

class WhileHandler(ActivityHandler):
    """Handler for While activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse While element into JSON structure."""
        result = {
            'type': 'While',
            'displayName': element.get('DisplayName', ''),
            'condition': unescape_expression(element.get('Condition', '')),
            'body': None,
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        # Parse child activity (body is direct child, not wrapped)
        for child in element:
            _, local = parse_tag(child.tag)
            # Skip metadata elements
            if local in ['WorkflowViewStateService.ViewState']:
                continue
            if child.tag.endswith('.ViewState'):
                continue

            # Parse body activity
            body = parse_activity(child)
            if body:
                result['body'] = body
                break

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build While element from JSON structure."""
        while_elem = ET.Element(get_ns_tag('', 'While'))

        # Set DisplayName
        if activity_json.get('displayName'):
            while_elem.set('DisplayName', activity_json['displayName'])

        # Set Condition
        while_elem.set('Condition', activity_json.get('condition', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('While', '514,707'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        while_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('While')
        while_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Build body activity
        if activity_json.get('body'):
            body_elem = build_activity(activity_json['body'], id_gen)
            if body_elem is not None:
                while_elem.append(body_elem)

        # Add ViewState
        viewstate = activity_json.get('viewState', {'IsExpanded': True})
        viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
        while_elem.append(viewstate_elem)

        return while_elem


# =============================================================================
# Delay Handler
# =============================================================================

class DelayHandler(ActivityHandler):
    """Handler for Delay activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Delay element into JSON structure."""
        result = {
            'type': 'Delay',
            'displayName': element.get('DisplayName', ''),
            'duration': unescape_expression(element.get('Duration', '00:00:00')),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Delay element from JSON structure."""
        delay_elem = ET.Element(get_ns_tag('', 'Delay'))

        # Set DisplayName
        if activity_json.get('displayName'):
            delay_elem.set('DisplayName', activity_json['displayName'])

        # Set Duration
        delay_elem.set('Duration', activity_json.get('duration', '00:00:00'))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Delay', '434,122'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        delay_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Delay')
        delay_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            delay_elem.append(viewstate_elem)

        return delay_elem


# =============================================================================
# InterruptibleWhile Handler
# =============================================================================

class InterruptibleWhileHandler(ActivityHandler):
    """Handler for InterruptibleWhile activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse InterruptibleWhile element into JSON structure."""
        result = {
            'type': 'InterruptibleWhile',
            'displayName': element.get('DisplayName', ''),
            'currentIndex': None,
            'interruptCondition': None,
            'maxIterations': -1,
            'body': None,
        }

        # Extract CurrentIndex (usually {x:Null})
        current_index = element.get('CurrentIndex')
        if current_index and current_index != '{x:Null}':
            result['currentIndex'] = unescape_expression(current_index)

        # Extract MaxIterations if present
        max_iter = element.get('MaxIterations')
        if max_iter:
            result['maxIterations'] = int(max_iter)

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse InterruptibleWhile.Condition (VisualBasicValue wrapper)
        condition_tag = get_ns_tag('ui', 'InterruptibleWhile.Condition')
        condition_elem = element.find(condition_tag)
        if condition_elem is not None:
            # Find VisualBasicValue child
            for child in condition_elem:
                _, local = parse_tag(child.tag)
                if local == 'VisualBasicValue':
                    result['condition'] = child.get('ExpressionText', '')
                    break

        # Parse InterruptCondition if present
        interrupt_tag = get_ns_tag('ui', 'InterruptibleWhile.InterruptCondition')
        interrupt_elem = element.find(interrupt_tag)
        if interrupt_elem is not None:
            for child in interrupt_elem:
                _, local = parse_tag(child.tag)
                if local == 'VisualBasicValue':
                    result['interruptCondition'] = child.get('ExpressionText', '')
                    break

        # Parse InterruptibleWhile.Body (ActivityAction with DelegateInArgument)
        body_tag = get_ns_tag('ui', 'InterruptibleWhile.Body')
        body_elem = element.find(body_tag)
        if body_elem is not None:
            for child in body_elem:
                _, local = parse_tag(child.tag)
                if local == 'ActivityAction':
                    action_info = ActivityActionParser.parse_activity_action(child)
                    result['body'] = action_info
                    break

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build InterruptibleWhile element from JSON structure."""
        while_elem = ET.Element(get_ns_tag('ui', 'InterruptibleWhile'))

        # Set CurrentIndex (default to {x:Null})
        if activity_json.get('currentIndex'):
            while_elem.set('CurrentIndex', activity_json['currentIndex'])
        else:
            while_elem.set('CurrentIndex', '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            while_elem.set('DisplayName', activity_json['displayName'])

        # Set MaxIterations if not default
        max_iter = activity_json.get('maxIterations', -1)
        if max_iter != -1:
            while_elem.set('MaxIterations', str(max_iter))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('InterruptibleWhile', '660,1200'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        while_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('InterruptibleWhile')
        while_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add InterruptibleWhile.Body (ActivityAction with DelegateInArgument)
        body_json = activity_json.get('body')
        if body_json:
            # Body format detection: Accept both Format A (wrapper) and Format B (direct)
            # Format A: {variableName, variableType, activity: {type, ...}}
            # Format B: {type, displayName, ...}
            if 'type' in body_json and 'activity' not in body_json:
                body_json = {
                    'variableName': 'argument',
                    'variableType': 'Boolean',
                    'activity': body_json
                }

            body_wrapper = ET.SubElement(while_elem, get_ns_tag('ui', 'InterruptibleWhile.Body'))

            # Create ActivityAction element
            activity_action = ET.SubElement(body_wrapper, get_ns_tag('', 'ActivityAction'))
            var_type = body_json.get('variableType', 'Boolean')
            xaml_type = TypeMapper.json_to_xaml_type(var_type)
            activity_action.set(get_ns_tag('x', 'TypeArguments'), xaml_type)

            # Add DelegateInArgument
            arg_wrapper = ET.SubElement(activity_action, get_ns_tag('', 'ActivityAction.Argument'))
            delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
            delegate.set(get_ns_tag('x', 'TypeArguments'), xaml_type)
            delegate.set('Name', body_json.get('variableName', 'argument'))

            # Build nested activity
            if body_json.get('activity'):
                activity_elem = build_activity(body_json['activity'], id_gen)
                if activity_elem is not None:
                    activity_action.append(activity_elem)

        # Add InterruptibleWhile.Condition
        condition_wrapper = ET.SubElement(while_elem, get_ns_tag('ui', 'InterruptibleWhile.Condition'))
        condition_value = ET.SubElement(condition_wrapper, get_ns_tag('', 'VisualBasicValue'))
        condition_value.set(get_ns_tag('x', 'TypeArguments'), 'x:Boolean')
        condition_value.set('ExpressionText', activity_json.get('condition', 'True'))
        # Generate IdRef for VisualBasicValue (uses backtick notation for generics)
        vb_id = id_gen.generate('VisualBasicValue`1')
        condition_value.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), vb_id)

        # Add InterruptCondition if present
        if activity_json.get('interruptCondition'):
            interrupt_wrapper = ET.SubElement(while_elem, get_ns_tag('ui', 'InterruptibleWhile.InterruptCondition'))
            interrupt_value = ET.SubElement(interrupt_wrapper, get_ns_tag('', 'VisualBasicValue'))
            interrupt_value.set(get_ns_tag('x', 'TypeArguments'), 'x:Boolean')
            interrupt_value.set('ExpressionText', activity_json['interruptCondition'])
            vb_id2 = id_gen.generate('VisualBasicValue`1')
            interrupt_value.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), vb_id2)

        # Add ViewState
        viewstate = activity_json.get('viewState', {'IsExpanded': True, 'IsPinned': False})
        viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
        while_elem.append(viewstate_elem)

        return while_elem


# =============================================================================
# Continue Handler
# =============================================================================

class ContinueHandler(ActivityHandler):
    """Handler for Continue activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Continue element into JSON structure."""
        result = {
            'type': 'Continue',
            'displayName': element.get('DisplayName', 'Continue'),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Continue element from JSON structure."""
        continue_elem = ET.Element(get_ns_tag('', 'Continue'))

        # Set DisplayName
        if activity_json.get('displayName'):
            continue_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Continue', '262,60'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        continue_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Continue')
        continue_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return continue_elem


# =============================================================================
# Break Handler
# =============================================================================

class BreakHandler(ActivityHandler):
    """Handler for Break activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse Break element into JSON structure."""
        result = {
            'type': 'Break',
            'displayName': element.get('DisplayName', 'Break'),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build Break element from JSON structure."""
        break_elem = ET.Element(get_ns_tag('', 'Break'))

        # Set DisplayName
        if activity_json.get('displayName'):
            break_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('Break', '262,60'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        break_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('Break')
        break_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return break_elem


# =============================================================================
# PathExists Handler
# =============================================================================

class PathExistsHandler(ActivityHandler):
    """Handler for PathExists activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse PathExists element into JSON structure."""
        result = {
            'type': 'PathExists',
            'displayName': element.get('DisplayName', ''),
            'path': unescape_expression(element.get('Path', '')),
            'pathType': element.get('PathType', 'File'),
            'exists': unescape_expression(element.get('Exists', '')),
        }

        # Extract optional Resource attribute
        resource = element.get('Resource')
        if resource and resource != '{x:Null}':
            result['resource'] = resource

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build PathExists element from JSON structure."""
        path_elem = ET.Element(get_ns_tag('ui', 'PathExists'))

        # Set Resource attribute (use parsed value if present, otherwise {x:Null})
        if activity_json.get('resource'):
            path_elem.set('Resource', activity_json['resource'])
        else:
            path_elem.set('Resource', '{x:Null}')

        # Set DisplayName
        if activity_json.get('displayName'):
            path_elem.set('DisplayName', activity_json['displayName'])

        # Set Exists (output variable)
        path_elem.set('Exists', activity_json.get('exists', ''))

        # Set Path
        path_elem.set('Path', activity_json.get('path', ''))

        # Set PathType
        path_elem.set('PathType', activity_json.get('pathType', 'File'))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('PathExists', '450,84'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        path_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('PathExists')
        path_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            path_elem.append(viewstate_elem)

        return path_elem


# =============================================================================
# KillProcess Handler
# =============================================================================

class KillProcessHandler(ActivityHandler):
    """Handler for KillProcess activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse KillProcess element into JSON structure."""
        result = {
            'type': 'KillProcess',
            'displayName': element.get('DisplayName', ''),
            'processName': element.get('ProcessName', ''),
        }

        # Extract optional attributes
        applies_to = element.get('AppliesTo')
        if applies_to and applies_to != '{x:Null}':
            result['appliesTo'] = applies_to

        process = element.get('Process')
        if process and process != '{x:Null}':
            result['process'] = process

        continue_on_error = element.get('ContinueOnError')
        if continue_on_error:
            result['continueOnError'] = continue_on_error.lower() == 'true'

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build KillProcess element from JSON structure."""
        kill_elem = ET.Element(get_ns_tag('ui', 'KillProcess'))

        # Set AppliesTo attribute (use parsed value if present, otherwise {x:Null})
        if activity_json.get('appliesTo'):
            kill_elem.set('AppliesTo', activity_json['appliesTo'])
        else:
            kill_elem.set('AppliesTo', '{x:Null}')

        # Set Process attribute (use parsed value if present, otherwise {x:Null})
        if activity_json.get('process'):
            kill_elem.set('Process', activity_json['process'])
        else:
            kill_elem.set('Process', '{x:Null}')

        # Set ContinueOnError
        if 'continueOnError' in activity_json:
            kill_elem.set('ContinueOnError', str(activity_json['continueOnError']))

        # Set DisplayName
        if activity_json.get('displayName'):
            kill_elem.set('DisplayName', activity_json['displayName'])

        # Set ProcessName
        kill_elem.set('ProcessName', activity_json.get('processName', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('KillProcess', '552,156'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        kill_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('KillProcess')
        kill_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            kill_elem.append(viewstate_elem)

        return kill_elem


# =============================================================================
# SetToClipboard Handler
# =============================================================================

class SetToClipboardHandler(ActivityHandler):
    """Handler for SetToClipboard activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse SetToClipboard element into JSON structure."""
        result = {
            'type': 'SetToClipboard',
            'displayName': element.get('DisplayName', ''),
            'text': unescape_expression(element.get('Text', '')),
        }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build SetToClipboard element from JSON structure."""
        clipboard_elem = ET.Element(get_ns_tag('ui', 'SetToClipboard'))

        # Set DisplayName
        if activity_json.get('displayName'):
            clipboard_elem.set('DisplayName', activity_json['displayName'])

        # Set Text
        clipboard_elem.set('Text', activity_json.get('text', ''))

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('SetToClipboard', '434,83'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        clipboard_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('SetToClipboard')
        clipboard_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return clipboard_elem


class InputDialogHandler(ActivityHandler):
    """Handler for InputDialog activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse InputDialog element into JSON structure."""
        result = {
            'type': 'InputDialog',
            'displayName': element.get('DisplayName', ''),
            'label': unescape_expression(element.get('Label', '')),
            'title': unescape_expression(element.get('Title', '')),
            'isPassword': element.get('IsPassword', 'False').lower() == 'true',
            'topMost': element.get('TopMost', 'False').lower() == 'true',
        }

        # Extract optional Options
        options = element.get('Options')
        if options and options != '{x:Null}':
            result['options'] = unescape_expression(options)

        # Extract optional OptionsString
        options_string = element.get('OptionsString')
        if options_string and options_string != '{x:Null}':
            result['optionsString'] = unescape_expression(options_string)

        # Parse Result (OutArgument)
        result_tag = get_ns_tag('ui', 'InputDialog.Result')
        result_elem = element.find(result_tag)
        if result_elem is not None:
            out_arg = result_elem.find(get_ns_tag('', 'OutArgument'))
            if out_arg is not None:
                result['result'] = {
                    'outArgument': {
                        'x:TypeArguments': canonicalize_type(out_arg.get(get_ns_tag('x', 'TypeArguments'), 'x:Object')),
                        'value': unescape_expression(out_arg.text or ''),
                    }
                }

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build InputDialog element from JSON structure."""
        dialog_elem = ET.Element(get_ns_tag('ui', 'InputDialog'))

        # Set DisplayName
        if activity_json.get('displayName'):
            dialog_elem.set('DisplayName', activity_json['displayName'])

        # Set Label
        dialog_elem.set('Label', activity_json.get('label', ''))

        # Set Title
        dialog_elem.set('Title', activity_json.get('title', ''))

        # Set IsPassword
        dialog_elem.set('IsPassword', str(activity_json.get('isPassword', False)))

        # Set TopMost
        dialog_elem.set('TopMost', str(activity_json.get('topMost', False)))

        # Set optional Options
        if activity_json.get('options'):
            dialog_elem.set('Options', activity_json['options'])

        # Set optional OptionsString
        if activity_json.get('optionsString'):
            dialog_elem.set('OptionsString', activity_json['optionsString'])

        # Build Result (OutArgument)
        result_info = activity_json.get('result')
        if result_info and result_info.get('outArgument'):
            result_wrapper = ET.SubElement(dialog_elem, get_ns_tag('ui', 'InputDialog.Result'))
            out_arg_info = result_info['outArgument']
            out_arg = ET.SubElement(result_wrapper, get_ns_tag('', 'OutArgument'))
            out_arg.set(get_ns_tag('x', 'TypeArguments'), out_arg_info.get('x:TypeArguments', 'x:Object'))
            out_arg.text = out_arg_info.get('value', '')

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('InputDialog', '444,191'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        dialog_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('InputDialog')
        dialog_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return dialog_elem


class InvokeCodeHandler(ActivityHandler):
    """Handler for InvokeCode activities."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse InvokeCode element into JSON structure."""
        result = {
            'type': 'InvokeCode',
            'displayName': element.get('DisplayName', ''),
            'code': unescape_expression(element.get('Code', '')),
            'continueOnError': element.get('ContinueOnError', 'False').lower() == 'true',
        }

        # Parse Arguments — direct children of <ui:InvokeCode.Arguments>
        args_tag = get_ns_tag('ui', 'InvokeCode.Arguments')
        args_elem = element.find(args_tag)
        if args_elem is not None:
            arguments = []
            for child in args_elem:
                _, local_name = parse_tag(child.tag)
                arg_entry = {
                    'direction': local_name,
                    'x:TypeArguments': canonicalize_type(child.get(get_ns_tag('x', 'TypeArguments'), '')),
                    'x:Key': child.get(get_ns_tag('x', 'Key'), ''),
                    'value': unescape_expression(child.text or ''),
                }
                arguments.append(arg_entry)
            if arguments:
                result['arguments'] = arguments

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build InvokeCode element from JSON structure."""
        code_elem = ET.Element(get_ns_tag('ui', 'InvokeCode'))

        # Set DisplayName
        if activity_json.get('displayName'):
            code_elem.set('DisplayName', activity_json['displayName'])

        # Set Code
        code_elem.set('Code', activity_json.get('code', ''))

        # Set ContinueOnError
        if activity_json.get('continueOnError'):
            code_elem.set('ContinueOnError', 'True')

        # Build Arguments — direct children of <ui:InvokeCode.Arguments>
        arguments = activity_json.get('arguments', [])
        if arguments:
            args_wrapper = ET.SubElement(code_elem, get_ns_tag('ui', 'InvokeCode.Arguments'))
            for arg in arguments:
                direction = arg.get('direction', 'InArgument')
                arg_elem = ET.SubElement(args_wrapper, get_ns_tag('', direction))
                arg_elem.set(get_ns_tag('x', 'TypeArguments'), arg.get('x:TypeArguments', 'x:String'))
                arg_elem.set(get_ns_tag('x', 'Key'), arg.get('x:Key', ''))
                if arg.get('value'):
                    arg_elem.text = arg['value']

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('InvokeCode', '434,191'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        code_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('InvokeCode')
        code_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        return code_elem


# =============================================================================
# UI Automation Handlers
# =============================================================================

class NClickHandler(ActivityHandler):
    """Handler for NClick (modern Click) activities."""

    # All known attributes for NClick
    NCLICK_ATTRIBUTES = [
        'ActivateBefore', 'ClickType', 'DisplayName', 'HealingAgentBehavior',
        'InteractionMode', 'KeyModifiers', 'MouseButton', 'ScopeIdentifier', 'Version',
    ]

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse NClick element into JSON structure."""
        result = {
            'type': 'NClick',
            'displayName': element.get('DisplayName', ''),
        }

        # Extract all known attributes
        for attr in self.NCLICK_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                # Convert attribute name to camelCase for JSON
                json_key = attr[0].lower() + attr[1:]
                result[json_key] = val

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse NClick.Target containing TargetAnchorable
        target_wrapper_tag = get_ns_tag('uix', 'NClick.Target')
        target_wrapper = element.find(target_wrapper_tag)
        if target_wrapper is not None:
            target_anchorable_tag = get_ns_tag('uix', 'TargetAnchorable')
            target_anchorable = target_wrapper.find(target_anchorable_tag)
            if target_anchorable is not None:
                result['target'] = TargetAnchorableParser.parse_target_anchorable(target_anchorable)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build NClick element from JSON structure."""
        click_elem = ET.Element(get_ns_tag('uix', 'NClick'))

        # Set all known attributes
        for attr in self.NCLICK_ATTRIBUTES:
            json_key = attr[0].lower() + attr[1:]
            if json_key in activity_json:
                click_elem.set(attr, activity_json[json_key])

        # Set DisplayName explicitly (since it's in NCLICK_ATTRIBUTES)
        if activity_json.get('displayName'):
            click_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('NClick', '484,189'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        click_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('NClick')
        click_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Build NClick.Target with TargetAnchorable
        target_json = activity_json.get('target')
        if target_json:
            target_wrapper = ET.SubElement(click_elem, get_ns_tag('uix', 'NClick.Target'))
            target_anchorable = TargetAnchorableBuilder.build_target_anchorable(target_json)
            target_wrapper.append(target_anchorable)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            click_elem.append(viewstate_elem)

        return click_elem


class NTypeIntoHandler(ActivityHandler):
    """Handler for NTypeInto (modern Type Into) activities."""

    # All known attributes for NTypeInto
    NTYPEINTO_ATTRIBUTES = [
        'ActivateBefore', 'ClickBeforeMode', 'ClipboardMode', 'DisplayName',
        'EmptyFieldMode', 'HealingAgentBehavior', 'InteractionMode', 'ScopeIdentifier',
        'Text', 'Version',
    ]

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse NTypeInto element into JSON structure."""
        result = {
            'type': 'NTypeInto',
            'displayName': element.get('DisplayName', ''),
        }

        # Extract all known attributes
        for attr in self.NTYPEINTO_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                json_key = attr[0].lower() + attr[1:]
                result[json_key] = val

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse NTypeInto.Target containing TargetAnchorable
        target_wrapper_tag = get_ns_tag('uix', 'NTypeInto.Target')
        target_wrapper = element.find(target_wrapper_tag)
        if target_wrapper is not None:
            target_anchorable_tag = get_ns_tag('uix', 'TargetAnchorable')
            target_anchorable = target_wrapper.find(target_anchorable_tag)
            if target_anchorable is not None:
                result['target'] = TargetAnchorableParser.parse_target_anchorable(target_anchorable)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build NTypeInto element from JSON structure."""
        typeinto_elem = ET.Element(get_ns_tag('uix', 'NTypeInto'))

        # Set all known attributes
        for attr in self.NTYPEINTO_ATTRIBUTES:
            json_key = attr[0].lower() + attr[1:]
            if json_key in activity_json:
                typeinto_elem.set(attr, activity_json[json_key])

        # Set DisplayName explicitly
        if activity_json.get('displayName'):
            typeinto_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('NTypeInto', '450,240'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        typeinto_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('NTypeInto')
        typeinto_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Build NTypeInto.Target with TargetAnchorable
        target_json = activity_json.get('target')
        if target_json:
            target_wrapper = ET.SubElement(typeinto_elem, get_ns_tag('uix', 'NTypeInto.Target'))
            target_anchorable = TargetAnchorableBuilder.build_target_anchorable(target_json)
            target_wrapper.append(target_anchorable)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            typeinto_elem.append(viewstate_elem)

        return typeinto_elem


class NCheckStateHandler(ActivityHandler):
    """Handler for NCheckState (Check App State) activities."""

    # All known attributes for NCheckState
    NCHECKSTATE_ATTRIBUTES = [
        'DisplayName', 'EnableIfNotExists', 'HealingAgentBehavior', 'ScopeIdentifier',
        'Timeout', 'Version',
    ]

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse NCheckState element into JSON structure."""
        result = {
            'type': 'NCheckState',
            'displayName': element.get('DisplayName', ''),
        }

        # Extract all known attributes
        for attr in self.NCHECKSTATE_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                json_key = attr[0].lower() + attr[1:]
                result[json_key] = val

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse NCheckState.Target containing TargetAnchorable
        target_wrapper_tag = get_ns_tag('uix', 'NCheckState.Target')
        target_wrapper = element.find(target_wrapper_tag)
        if target_wrapper is not None:
            target_anchorable_tag = get_ns_tag('uix', 'TargetAnchorable')
            target_anchorable = target_wrapper.find(target_anchorable_tag)
            if target_anchorable is not None:
                result['target'] = TargetAnchorableParser.parse_target_anchorable(target_anchorable)

        # Parse NCheckState.IfExists containing child activity
        if_exists_tag = get_ns_tag('uix', 'NCheckState.IfExists')
        if_exists_elem = element.find(if_exists_tag)
        if if_exists_elem is not None:
            for child in if_exists_elem:
                child_activity = parse_activity(child)
                if child_activity:
                    result['ifExists'] = child_activity
                    break

        # Parse NCheckState.IfNotExists containing child activity
        if_not_exists_tag = get_ns_tag('uix', 'NCheckState.IfNotExists')
        if_not_exists_elem = element.find(if_not_exists_tag)
        if if_not_exists_elem is not None:
            for child in if_not_exists_elem:
                child_activity = parse_activity(child)
                if child_activity:
                    result['ifNotExists'] = child_activity
                    break

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build NCheckState element from JSON structure."""
        check_elem = ET.Element(get_ns_tag('uix', 'NCheckState'))

        # Set all known attributes
        for attr in self.NCHECKSTATE_ATTRIBUTES:
            json_key = attr[0].lower() + attr[1:]
            if json_key in activity_json:
                check_elem.set(attr, activity_json[json_key])

        # Set DisplayName explicitly
        if activity_json.get('displayName'):
            check_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('NCheckState', '484,639'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        check_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('NCheckState')
        check_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Build NCheckState.IfExists with child activity
        if_exists_json = activity_json.get('ifExists')
        if if_exists_json:
            if_exists_wrapper = ET.SubElement(check_elem, get_ns_tag('uix', 'NCheckState.IfExists'))
            child_elem = build_activity(if_exists_json, id_gen)
            if child_elem is not None:
                if_exists_wrapper.append(child_elem)

        # Build NCheckState.IfNotExists with child activity
        if_not_exists_json = activity_json.get('ifNotExists')
        if if_not_exists_json:
            if_not_exists_wrapper = ET.SubElement(check_elem, get_ns_tag('uix', 'NCheckState.IfNotExists'))
            child_elem = build_activity(if_not_exists_json, id_gen)
            if child_elem is not None:
                if_not_exists_wrapper.append(child_elem)

        # Build NCheckState.Target with TargetAnchorable
        target_json = activity_json.get('target')
        if target_json:
            target_wrapper = ET.SubElement(check_elem, get_ns_tag('uix', 'NCheckState.Target'))
            target_anchorable = TargetAnchorableBuilder.build_target_anchorable(target_json)
            target_wrapper.append(target_anchorable)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            check_elem.append(viewstate_elem)

        return check_elem


class NMouseScrollHandler(ActivityHandler):
    """Handler for NMouseScroll (Mouse Scroll) activities."""

    # All known attributes for NMouseScroll
    NMOUSESCROLL_ATTRIBUTES = [
        'ActivateBefore', 'Amount', 'Direction', 'DisplayName', 'HealingAgentBehavior',
        'InteractionMode', 'KeyModifiers', 'MovementUnits', 'ScopeIdentifier', 'Version',
    ]

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse NMouseScroll element into JSON structure."""
        result = {
            'type': 'NMouseScroll',
            'displayName': element.get('DisplayName', ''),
        }

        # Extract all known attributes
        for attr in self.NMOUSESCROLL_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                json_key = attr[0].lower() + attr[1:]
                result[json_key] = val

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse NMouseScroll.Target containing TargetAnchorable
        target_wrapper_tag = get_ns_tag('uix', 'NMouseScroll.Target')
        target_wrapper = element.find(target_wrapper_tag)
        if target_wrapper is not None:
            target_anchorable_tag = get_ns_tag('uix', 'TargetAnchorable')
            target_anchorable = target_wrapper.find(target_anchorable_tag)
            if target_anchorable is not None:
                result['target'] = TargetAnchorableParser.parse_target_anchorable(target_anchorable)

        # Parse NMouseScroll.SearchedElement containing SearchedElement
        searched_elem_wrapper_tag = get_ns_tag('uix', 'NMouseScroll.SearchedElement')
        searched_elem_wrapper = element.find(searched_elem_wrapper_tag)
        if searched_elem_wrapper is not None:
            searched_elem_tag = get_ns_tag('uix', 'SearchedElement')
            searched_elem = searched_elem_wrapper.find(searched_elem_tag)
            if searched_elem is not None:
                result['searchedElement'] = SearchedElementParser.parse_searched_element(searched_elem)

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build NMouseScroll element from JSON structure."""
        scroll_elem = ET.Element(get_ns_tag('uix', 'NMouseScroll'))

        # Set all known attributes
        for attr in self.NMOUSESCROLL_ATTRIBUTES:
            json_key = attr[0].lower() + attr[1:]
            if json_key in activity_json:
                scroll_elem.set(attr, activity_json[json_key])

        # Set DisplayName explicitly
        if activity_json.get('displayName'):
            scroll_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('NMouseScroll', '416,299'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        scroll_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('NMouseScroll')
        scroll_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Build NMouseScroll.SearchedElement if present
        searched_elem_json = activity_json.get('searchedElement')
        if searched_elem_json:
            searched_wrapper = ET.SubElement(scroll_elem, get_ns_tag('uix', 'NMouseScroll.SearchedElement'))
            searched_elem = SearchedElementBuilder.build_searched_element(searched_elem_json)
            searched_wrapper.append(searched_elem)

        # Build NMouseScroll.Target with TargetAnchorable
        target_json = activity_json.get('target')
        if target_json:
            target_wrapper = ET.SubElement(scroll_elem, get_ns_tag('uix', 'NMouseScroll.Target'))
            target_anchorable = TargetAnchorableBuilder.build_target_anchorable(target_json)
            target_wrapper.append(target_anchorable)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            scroll_elem.append(viewstate_elem)

        return scroll_elem


class SearchedElementHandler(ActivityHandler):
    """Handler for SearchedElement (nested element for UI search configuration)."""

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse SearchedElement element into JSON structure."""
        return SearchedElementParser.parse_searched_element(element)

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build SearchedElement element from JSON structure."""
        return SearchedElementBuilder.build_searched_element(activity_json)


class NApplicationCardHandler(ActivityHandler):
    """Handler for NApplicationCard (Use Application/Browser) activities."""

    # All known attributes for NApplicationCard
    NAPPLICATIONCARD_ATTRIBUTES = [
        'AttachMode', 'CloseMode', 'DisplayName', 'HealingAgentBehavior',
        'OpenMode', 'ScopeGuid', 'Version',
    ]

    def parse(self, element: ET.Element) -> Dict[str, Any]:
        """Parse NApplicationCard element into JSON structure."""
        result = {
            'type': 'NApplicationCard',
            'displayName': element.get('DisplayName', ''),
        }

        # Extract all known attributes
        for attr in self.NAPPLICATIONCARD_ATTRIBUTES:
            val = element.get(attr)
            if val is not None:
                json_key = attr[0].lower() + attr[1:]
                result[json_key] = val

        # Extract HintSize
        hint_size_attr = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        if hint_size_attr in element.attrib:
            result['hintSize'] = element.get(hint_size_attr)

        # Extract IdRef
        id_ref_attr = get_ns_tag('sap2010', 'WorkflowViewState.IdRef')
        if id_ref_attr in element.attrib:
            result['idRef'] = element.get(id_ref_attr)

        # Parse NApplicationCard.Body containing ActivityAction
        body_tag = get_ns_tag('uix', 'NApplicationCard.Body')
        body_elem = element.find(body_tag)
        if body_elem is not None:
            action_tag = get_ns_tag('', 'ActivityAction')
            action_elem = body_elem.find(action_tag)
            if action_elem is not None:
                result['body'] = ActivityActionParser.parse_activity_action(action_elem)

        # Parse NApplicationCard.TargetApp
        target_app_tag = get_ns_tag('uix', 'NApplicationCard.TargetApp')
        target_app_wrapper = element.find(target_app_tag)
        if target_app_wrapper is not None:
            target_app_elem_tag = get_ns_tag('uix', 'TargetApp')
            target_app_elem = target_app_wrapper.find(target_app_elem_tag)
            if target_app_elem is not None:
                result['targetApp'] = TargetAppParser.parse_target_app(target_app_elem)

        # Parse NApplicationCard.OCREngine - store as raw XML for perfect round-trip
        ocr_tag = get_ns_tag('uix', 'NApplicationCard.OCREngine')
        ocr_elem = element.find(ocr_tag)
        if ocr_elem is not None:
            result['ocrEngineRaw'] = ET.tostring(ocr_elem, encoding='unicode')

        # Parse ViewState
        viewstate = ViewStateBuilder.parse_viewstate(element)
        if viewstate:
            result['viewState'] = viewstate

        return result

    def build(self, activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> ET.Element:
        """Build NApplicationCard element from JSON structure."""
        card_elem = ET.Element(get_ns_tag('uix', 'NApplicationCard'))

        # Set all known attributes
        for attr in self.NAPPLICATIONCARD_ATTRIBUTES:
            json_key = attr[0].lower() + attr[1:]
            if json_key in activity_json:
                card_elem.set(attr, activity_json[json_key])

        # Set DisplayName explicitly
        if activity_json.get('displayName'):
            card_elem.set('DisplayName', activity_json['displayName'])

        # Set HintSize
        hint_size = activity_json.get('hintSize', DEFAULT_HINT_SIZES.get('NApplicationCard', '552,1503'))
        sap_hint = get_ns_tag('sap', 'VirtualizedContainerService.HintSize')
        card_elem.set(sap_hint, hint_size)

        # Set IdRef
        id_ref = activity_json.get('idRef') or id_gen.generate('NApplicationCard')
        card_elem.set(get_ns_tag('sap2010', 'WorkflowViewState.IdRef'), id_ref)

        # Build NApplicationCard.Body with ActivityAction
        body_json = activity_json.get('body')
        if body_json:
            # Body format detection: Accept both Format A (wrapper) and Format B (direct)
            # Format A: {variableName, variableType, activity: {type, ...}}
            # Format B: {type, displayName, ...}
            if 'type' in body_json and 'activity' not in body_json:
                body_json = {
                    'variableName': 'WSSessionData',
                    'variableType': 'x:Object',
                    'activity': body_json
                }

            body_wrapper = ET.SubElement(card_elem, get_ns_tag('uix', 'NApplicationCard.Body'))
            action_elem = ET.SubElement(body_wrapper, get_ns_tag('', 'ActivityAction'))
            # Honor variableType from body_json, default to x:Object
            var_type = body_json.get('variableType', 'x:Object')
            action_elem.set(get_ns_tag('x', 'TypeArguments'), var_type)

            # Create ActivityAction.Argument with DelegateInArgument
            arg_wrapper = ET.SubElement(action_elem, get_ns_tag('', 'ActivityAction.Argument'))
            delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
            delegate.set(get_ns_tag('x', 'TypeArguments'), var_type)
            delegate.set('Name', body_json.get('variableName', 'WSSessionData'))

            # Build nested activity
            if body_json.get('activity'):
                activity_elem = build_activity(body_json['activity'], id_gen)
                if activity_elem is not None:
                    action_elem.append(activity_elem)

        # Restore OCREngine from raw XML if available
        ocr_raw = activity_json.get('ocrEngineRaw')
        if ocr_raw:
            try:
                ocr_elem = ET.fromstring(ocr_raw)
                card_elem.append(ocr_elem)
            except ET.ParseError:
                pass
        else:
            # Create default OCREngine structure
            ocr_wrapper = ET.SubElement(card_elem, get_ns_tag('uix', 'NApplicationCard.OCREngine'))
            activity_func = ET.SubElement(ocr_wrapper, get_ns_tag('', 'ActivityFunc'))
            activity_func.set(get_ns_tag('x', 'TypeArguments'),
                              'sd2:Image, scg:IEnumerable(scg:KeyValuePair(sd1:Rectangle, x:String))')
            arg_wrapper = ET.SubElement(activity_func, get_ns_tag('', 'ActivityFunc.Argument'))
            delegate = ET.SubElement(arg_wrapper, get_ns_tag('', 'DelegateInArgument'))
            delegate.set(get_ns_tag('x', 'TypeArguments'), 'sd2:Image')
            delegate.set('Name', 'Image')

        # Build NApplicationCard.TargetApp
        target_app_json = activity_json.get('targetApp')
        if target_app_json:
            target_app_wrapper = ET.SubElement(card_elem, get_ns_tag('uix', 'NApplicationCard.TargetApp'))
            target_app_elem = TargetAppBuilder.build_target_app(target_app_json)
            target_app_wrapper.append(target_app_elem)

        # Add ViewState if specified
        viewstate = activity_json.get('viewState')
        if viewstate:
            viewstate_elem = ViewStateBuilder.create_viewstate_element(viewstate)
            card_elem.append(viewstate_elem)

        return card_elem


# =============================================================================
# Activity Handler Factory
# =============================================================================

# Factory pattern for activity handlers
ACTIVITY_HANDLERS: Dict[str, ActivityHandler] = {
    # Core activities
    'Sequence': SequenceHandler(),
    'Assign': AssignHandler(),
    'If': IfHandler(),
    'LogMessage': LogMessageHandler(),
    'InvokeWorkflowFile': InvokeWorkflowFileHandler(),

    # Flowchart activities
    'Flowchart': FlowchartHandler(),
    'FlowStep': FlowStepHandler(),
    'FlowDecision': FlowDecisionHandler(),

    # Control flow activities
    'Switch': SwitchHandler(),
    'TryCatch': TryCatchHandler(),
    'ForEach': ForEachHandler(),
    'ForEachRow': ForEachRowHandler(),
    'Rethrow': RethrowHandler(),
    'While': WhileHandler(),
    'InterruptibleWhile': InterruptibleWhileHandler(),
    'Continue': ContinueHandler(),
    'Break': BreakHandler(),
    'Delay': DelayHandler(),
    'Throw': ThrowHandler(),
    'Return': ReturnHandler(),

    # Excel activities
    'ExcelProcessScopeX': ExcelProcessScopeXHandler(),
    'ExcelApplicationCard': ExcelApplicationCardHandler(),
    'ReadRangeX': ReadRangeXHandler(),
    'SaveExcelFileX': SaveExcelFileXHandler(),
    'WriteCellX': WriteCellXHandler(),
    'WriteRangeX': WriteRangeXHandler(),
    'CopyPasteRangeX': CopyPasteRangeXHandler(),
    'ClearRangeX': ClearRangeXHandler(),
    'FilterX': FilterXHandler(),
    'FindFirstLastDataRowX': FindFirstLastDataRowXHandler(),

    # File operations
    'CreateDirectory': CreateDirectoryHandler(),
    'MoveFile': MoveFileHandler(),
    'ReadRange': ReadRangeHandler(),
    'PathExists': PathExistsHandler(),
    'DeleteFileX': DeleteFileXHandler(),
    'ReadTextFile': ReadTextFileHandler(),

    # Data activities
    'AddDataRow': AddDataRowHandler(),
    'BuildDataTable': BuildDataTableHandler(),

    # Process activities
    'KillProcess': KillProcessHandler(),

    # Utilities
    'CommentOut': CommentOutHandler(),
    'RetryScope': RetryScopeHandler(),
    'SetToClipboard': SetToClipboardHandler(),
    'InputDialog': InputDialogHandler(),
    'InvokeCode': InvokeCodeHandler(),

    # UI Automation activities
    'NApplicationCard': NApplicationCardHandler(),
    'NClick': NClickHandler(),
    'NTypeInto': NTypeIntoHandler(),
    'NCheckState': NCheckStateHandler(),
    'NMouseScroll': NMouseScrollHandler(),
    'SearchedElement': SearchedElementHandler(),
}


def parse_activity(element: ET.Element) -> Optional[Dict[str, Any]]:
    """Parse an activity element using the appropriate handler."""
    activity_type = get_activity_type(element)

    handler = ACTIVITY_HANDLERS.get(activity_type)
    if handler:
        return handler.parse(element)

    # Return generic representation for unknown activities
    return {
        'type': activity_type,
        'displayName': element.get('DisplayName', ''),
        'raw': True,  # Flag indicating this is unparsed
    }


def build_activity(activity_json: Dict[str, Any], id_gen: IdRefGenerator) -> Optional[ET.Element]:
    """Build an activity element using the appropriate handler."""
    activity_type = activity_json.get('type')

    if not activity_type:
        # Check if this is an ActivityAction wrapper format
        if 'activity' in activity_json and isinstance(activity_json['activity'], dict):
            print("Warning: body field uses ActivityAction wrapper format, unwrapping", file=sys.stderr)
            return build_activity(activity_json['activity'], id_gen)

        # No type key and no activity key - cannot process
        print(f"Warning: No 'type' key in activity JSON. Keys: {list(activity_json.keys())}", file=sys.stderr)
        return None

    handler = ACTIVITY_HANDLERS.get(activity_type)
    if handler:
        return handler.build(activity_json, id_gen)

    # Cannot build unknown activity types
    print(f"Warning: Unknown activity type '{activity_type}', skipping", file=sys.stderr)
    return None


# =============================================================================
# XAML Parser
# =============================================================================

class XamlParser:
    """Parser for converting XAML files to JSON representation."""

    def __init__(self):
        self.metadata_manager = MetadataManager()

    def parse_file(self, filepath: str) -> Dict[str, Any]:
        """Load XAML file and return JSON structure."""
        global _canon_xmlns_bindings, _canon_uri_to_canonical

        # Parse the XML
        tree = ET.parse(filepath)
        root = tree.getroot()

        # Extract xmlns bindings and build canonicalization mappings
        xmlns_bindings = MetadataManager.extract_xmlns_bindings_from_file(filepath)
        uri_to_canonical = MetadataManager.build_uri_to_canonical_prefix(xmlns_bindings)

        # Set module-level canonicalization context for activity handlers
        _canon_xmlns_bindings = xmlns_bindings
        _canon_uri_to_canonical = uri_to_canonical

        # Log canonicalization remappings (only non-identity ones)
        remaps = []
        for doc_prefix, uri in xmlns_bindings.items():
            canonical = uri_to_canonical.get(uri)
            if canonical is not None and canonical != doc_prefix and doc_prefix:
                remaps.append(f'{doc_prefix}->{canonical}')
        if remaps:
            print(f"[Reader] Prefix canonicalization: {', '.join(remaps)}", file=sys.stderr)

        # Extract metadata (with canonicalization context)
        metadata = self.metadata_manager.extract_metadata(root, xmlns_bindings, uri_to_canonical)

        # Find the main Sequence (usually the first Sequence child of Activity)
        workflow = None
        for child in root:
            activity_type = get_activity_type(child)
            if activity_type in ACTIVITY_HANDLERS:
                workflow = parse_activity(child)
                break

        # Clear canonicalization context after parsing
        _canon_xmlns_bindings = {}
        _canon_uri_to_canonical = {}

        return {
            'metadata': metadata,
            'workflow': workflow,
        }


# =============================================================================
# XAML Constructor
# =============================================================================

class XamlConstructor:
    """Constructor for building XAML from JSON representation."""

    def __init__(self):
        self.metadata_manager = MetadataManager()
        self.id_gen = IdRefGenerator()

    def construct_from_json(self, json_data: Dict[str, Any]) -> ET.ElementTree:
        """Build complete XAML tree from JSON data."""
        metadata = json_data.get('metadata', {})
        workflow = json_data.get('workflow', {})

        # === Auto-correction pipeline (with safe fallback) ===
        original_workflow = copy.deepcopy(workflow)
        try:
            corrector = WorkflowAutoCorrector()
            corrected_workflow, correction_context = corrector.correct(workflow)
            if correction_context.corrections_applied:
                expr_count = sum(1 for c in correction_context.corrections_applied
                                 if c['type'] in ('expression_wrap', 'safety_net_wrap'))
                type_count = sum(1 for c in correction_context.corrections_applied
                                 if c['type'] == 'type_normalize')
                print(f"[Writer] Auto-corrections: {expr_count} expressions wrapped, "
                      f"{type_count} types normalized", file=sys.stderr)
        except Exception as e:
            print(f"[AutoCorrector] WARNING: correction failed ({type(e).__name__}: {e}), "
                  f"falling back to uncorrected JSON", file=sys.stderr)
            corrected_workflow = original_workflow

        # Create root element
        root = self.metadata_manager.create_root_element(metadata)

        # === Three-tier namespace resolution ===

        # Tier 1: Auto-detect required namespaces from corrected workflow JSON
        auto_detected = self.metadata_manager.detect_required_namespaces(corrected_workflow)

        # Tier 2: Default baseline namespaces (always required)
        all_prefixes = set(DEFAULT_REQUIRED_NAMESPACES)

        # Tier 3: Merge auto-detected prefixes
        all_prefixes.update(auto_detected)

        # Apply xmlns attributes to root element
        self.metadata_manager.apply_xmlns_to_root(root, all_prefixes)

        # === Custom xmlns filtering (Gap B) ===
        custom_bindings = metadata.get('xmlnsBindings', {})
        used_custom = MetadataManager.filter_used_custom_xmlns(
            custom_bindings, corrected_workflow)
        for prefix, uri in used_custom.items():
            root.set(f'xmlns:{prefix}', uri)
        if custom_bindings:
            print(f"[Writer] Custom xmlns: {len(used_custom)} preserved, "
                  f"{len(custom_bindings) - len(used_custom)} filtered",
                  file=sys.stderr)

        # Generate namespace strings for TextExpression.NamespacesForImplementation
        metadata_namespaces = metadata.get('namespaces', [])
        namespace_strings = self.metadata_manager.generate_namespace_strings(
            all_prefixes, metadata_namespaces
        )

        # Diagnostic output
        metadata_valid_count = sum(
            1 for ns in metadata_namespaces
            if isinstance(ns, str) and ns.strip()
        )
        if metadata_valid_count == 0 and metadata_namespaces:
            print("[Writer] WARNING: Metadata namespaces were empty/whitespace; "
                  "using auto-detection", file=sys.stderr)
        print(f"[Writer] Namespace resolution: {len(auto_detected)} auto-detected, "
              f"{len(DEFAULT_REQUIRED_NAMESPACES)} defaults, "
              f"{metadata_valid_count} from metadata -> "
              f"{len(all_prefixes)} total xmlns declarations",
              file=sys.stderr)

        # Apply TextExpression.NamespacesForImplementation
        self.metadata_manager.apply_namespaces(root, namespace_strings)

        # Add assembly references (minimal generation from prefixes + existing)
        used_prefixes = MetadataManager.detect_all_used_prefixes(corrected_workflow)
        existing_refs = metadata.get('assemblyReferences', [])
        assembly_refs = MetadataManager.generate_minimal_assembly_refs(
            used_prefixes, existing_refs)
        existing_valid = [r for r in existing_refs
                          if r and isinstance(r, str) and r.strip()]
        prefix_derived = len(assembly_refs) - len(existing_valid)
        if not existing_valid and not used_prefixes:
            print("[Writer] No valid assembly references in metadata; using default set "
                  f"({len(assembly_refs)} references)", file=sys.stderr)
        else:
            print(f"[Writer] Assembly refs: {len(existing_valid)} from metadata, "
                  f"{prefix_derived} prefix-derived -> {len(assembly_refs)} total",
                  file=sys.stderr)
        self.metadata_manager.apply_assembly_refs(root, assembly_refs)

        # Add arguments
        self.metadata_manager.apply_arguments(root, metadata.get('arguments', []))

        # Add VisualBasic.Settings (required for VB expressions)
        self._add_vb_settings(root, namespace_strings)

        # Build main workflow
        if corrected_workflow:
            self.id_gen.reset()
            workflow_elem = build_activity(corrected_workflow, self.id_gen)
            if workflow_elem is not None:
                root.append(workflow_elem)

        return ET.ElementTree(root)

    def _add_vb_settings(self, root: ET.Element, namespace_strings: List[str]):
        """Add VisualBasic.Settings element with import references for VB expression support."""
        vb_settings_tag = get_ns_tag('mva', 'VisualBasic.Settings')
        vb_settings = ET.SubElement(root, vb_settings_tag)

        vb_value_tag = get_ns_tag('mva', 'VisualBasicSettings')
        vb_value = ET.SubElement(vb_settings, vb_value_tag)

        # Generate VisualBasicImportReference entries for each CLR namespace
        if namespace_strings:
            imports_tag = get_ns_tag('mva', 'VisualBasicSettings.ImportReferences')
            imports_elem = ET.SubElement(vb_value, imports_tag)

            for ns in namespace_strings:
                assembly = CLR_NAMESPACE_TO_ASSEMBLY.get(ns, '')
                if not assembly:
                    continue
                ref_tag = get_ns_tag('mva', 'VisualBasicImportReference')
                ref_elem = ET.SubElement(imports_elem, ref_tag)
                ref_elem.set('Assembly', assembly)
                ref_elem.set('Import', ns)


# =============================================================================
# Main CLI
# =============================================================================

def main():
    """Main entry point with CLI argument parsing."""
    parser = argparse.ArgumentParser(
        description='XAML Syntaxer - Bidirectional XAML-JSON conversion for UiPath workflows',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  Read Mode (XAML to JSON):
    python xaml_syntaxer.py --mode read --input workflow.xaml --output workflow.json

  Write Mode (JSON to XAML):
    python xaml_syntaxer.py --mode write --input workflow.json --output workflow.xaml
        '''
    )

    parser.add_argument(
        '--mode', '-m',
        required=True,
        choices=['read', 'write'],
        help='Conversion mode: "read" (XAML to JSON) or "write" (JSON to XAML)'
    )

    parser.add_argument(
        '--input', '-i',
        required=True,
        help='Path to input file (XAML for read mode, JSON for write mode)'
    )

    parser.add_argument(
        '--output', '-o',
        required=True,
        help='Path to output file (JSON for read mode, XAML for write mode)'
    )

    parser.add_argument(
        '--pretty', '-p',
        action='store_true',
        default=True,
        help='Pretty-print JSON output (default: True)'
    )

    parser.add_argument(
        '--validate', '-v',
        action='store_true',
        help='Validate output (future enhancement)'
    )

    args = parser.parse_args()

    # Setup namespaces for ElementTree
    setup_namespaces()

    # Validate input file exists
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        return 1

    if not input_path.is_file():
        print(f"Error: Input path is not a file: {args.input}", file=sys.stderr)
        return 1

    try:
        if args.mode == 'read':
            # XAML to JSON
            xaml_parser = XamlParser()
            json_data = xaml_parser.parse_file(args.input)

            # Write JSON output
            with open(args.output, 'w', encoding='utf-8') as f:
                if args.pretty:
                    json.dump(json_data, f, indent=2, ensure_ascii=False)
                else:
                    json.dump(json_data, f, ensure_ascii=False)

            print(f"Successfully converted XAML to JSON: {args.output}")

        else:  # write mode
            # JSON to XAML
            with open(args.input, 'r', encoding='utf-8') as f:
                json_data = json.load(f)

            constructor = XamlConstructor()
            tree = constructor.construct_from_json(json_data)

            # Write XAML output with proper formatting
            tree.write(
                args.output,
                encoding='utf-8',
                xml_declaration=True,
            )

            print(f"Successfully converted JSON to XAML: {args.output}")

        return 0

    except ET.ParseError as e:
        print(f"Error: Invalid XML in input file: {e}", file=sys.stderr)
        return 1

    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in input file: {e}", file=sys.stderr)
        return 1

    except FileNotFoundError as e:
        print(f"Error: File not found: {e}", file=sys.stderr)
        return 1

    except PermissionError as e:
        print(f"Error: Permission denied: {e}", file=sys.stderr)
        return 1

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
