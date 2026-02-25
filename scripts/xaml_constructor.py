"""
UiPath Activity Constructor

A template library and builder for UiPath workflow activities.
Works in conjunction with xaml_syntaxer.py for bidirectional XAML-JSON conversion.

Usage Examples:

1. List all available activities:
   python xaml_constructor.py --mode list

2. List activities to file:
   python xaml_constructor.py --mode list --output activities.json

3. Build activity from JSON file:
   python xaml_constructor.py --mode build --input activity.json

4. Build activity from JSON string:
   python xaml_constructor.py --mode build --input '{"type":"Assign","displayName":"Set Variable","to":"myVar","value":"\\"Hello\\""}'

5. Build and save to file:
   python xaml_constructor.py --mode build --input activity.json --output built.json

6. Build from Reader JSON (EDIT workflow - preserves metadata):
   python xaml_constructor.py --mode build --input reader_output.json --output constructor_output.json

Note: When input JSON contains a 'metadata' key (from Reader output for EDIT workflows),
the Constructor preserves it unchanged in the output. For NEW workflows (activity-only JSON
without metadata), default empty metadata is used.
"""

import json
import argparse
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional
import copy
import re

# Activity categories
CORE = "core"
CONTROL_FLOW = "control_flow"
EXCEL = "excel"
FILE_OPS = "file_ops"
PROCESS = "process"
UTILITIES = "utilities"
UI_AUTOMATION = "ui_automation"

# Template directory path
TEMPLATE_DIR = Path(__file__).parent / 'activity_templates'

# Namespace mappings
NAMESPACE_MAPPING = {
    "default": "",
    "ui": "http://schemas.uipath.com/workflow/activities",
    "ueab": "clr-namespace:UiPath.Excel.Activities.Business;assembly=UiPath.Excel.Activities",
    "uix": "http://schemas.uipath.com/workflow/activities/uix"
}

# Excel activities that require ExcelApplicationCard scope
EXCEL_SCOPED_ACTIVITIES = {
    "AutoFillRangeX", "ClearRangeX", "CopyPasteRangeX", "DeleteRowsX",
    "FilterX", "FindFirstLastDataRowX", "FormatRangeX", "GetSelectedRangeX",
    "InsertRowsX", "LookupRangeX", "ReadRangeX", "SaveExcelFileX",
    "WriteCellX", "WriteRangeX"
}


class TemplateLoader:
    """Loads and manages activity templates from JSON files."""

    def __init__(self, template_dir: Path):
        self.template_dir = template_dir
        self.templates: Dict[str, dict] = {}
        self._load_all_templates()

    def _load_all_templates(self) -> None:
        """Load all JSON templates from subdirectories."""
        if not self.template_dir.exists():
            return

        for category_dir in self.template_dir.iterdir():
            if category_dir.is_dir() and not category_dir.name.startswith('.'):
                for template_file in category_dir.glob('*.json'):
                    activity_type = template_file.stem
                    try:
                        with open(template_file, 'r', encoding='utf-8') as f:
                            self.templates[activity_type] = json.load(f)
                    except json.JSONDecodeError as e:
                        print(f"Warning: Failed to load template {template_file}: {e}", file=sys.stderr)

    def get_template(self, activity_type: str) -> Optional[dict]:
        """Get template for specific activity type."""
        return self.templates.get(activity_type)

    def list_all_templates(self) -> Dict[str, List[dict]]:
        """Return all templates organized by category."""
        categorized: Dict[str, List[dict]] = {}

        if not self.template_dir.exists():
            return categorized

        for category_dir in self.template_dir.iterdir():
            if category_dir.is_dir() and not category_dir.name.startswith('.'):
                category = category_dir.name
                categorized[category] = []
                for template_file in sorted(category_dir.glob('*.json')):
                    try:
                        with open(template_file, 'r', encoding='utf-8') as f:
                            template_data = json.load(f)
                            categorized[category].append({
                                'type': template_data.get('type', ''),
                                'displayName': template_data.get('displayName', ''),
                                'description': template_data.get('description', ''),
                                'namespace': template_data.get('namespace', 'default'),
                                'requiredAttributes': template_data.get('requiredAttributes', [])
                            })
                    except json.JSONDecodeError:
                        continue

        return categorized

    def get_activity_types(self) -> List[str]:
        """Return list of all available activity types."""
        return list(self.templates.keys())


class ActivityBuilder:
    """Builds and validates activity JSON from templates."""

    def __init__(self, template_loader: TemplateLoader):
        self.template_loader = template_loader

    def validate_activity(self, activity_json: dict) -> Tuple[bool, List[str]]:
        """Validate activity JSON against template requirements."""
        errors: List[str] = []

        activity_type = activity_json.get('type')
        if not activity_type:
            return False, ["Missing 'type' field"]

        template = self.template_loader.get_template(activity_type)
        if not template:
            available = self.template_loader.get_activity_types()
            errors.append(f"Unknown activity type: {activity_type}")
            if available:
                errors.append(f"Available types: {', '.join(sorted(available))}")
            return False, errors

        # Check required attributes
        required = template.get('requiredAttributes', [])
        for attr in required:
            if attr not in activity_json:
                errors.append(f"Missing required attribute: {attr}")

        # Type validation for known structures
        self._validate_nested_structures(activity_json, errors)

        return len(errors) == 0, errors

    def _validate_nested_structures(self, activity_json: dict, errors: List[str]) -> None:
        """Validate types of nested structures."""
        # activities must be a list
        if 'activities' in activity_json and not isinstance(activity_json['activities'], list):
            errors.append("'activities' must be a list")

        # body must be a dict
        if 'body' in activity_json and not isinstance(activity_json['body'], dict):
            errors.append("'body' must be a dict")

        # cases must be a dict (for Switch)
        if 'cases' in activity_json and not isinstance(activity_json['cases'], dict):
            errors.append("'cases' must be a dict")

        # hintSize should be a string like "400,200"
        if 'hintSize' in activity_json:
            hint = activity_json['hintSize']
            if isinstance(hint, str):
                parts = hint.split(',')
                if len(parts) != 2:
                    errors.append("'hintSize' should be in format 'width,height' (e.g., '400,200')")

        # variables must be a list
        if 'variables' in activity_json and not isinstance(activity_json['variables'], list):
            errors.append("'variables' must be a list")

        # arguments must be a list (for InvokeWorkflowFile)
        if 'arguments' in activity_json and not isinstance(activity_json['arguments'], list):
            errors.append("'arguments' must be a list")

    def validate_excel_scoping(self, workflow_json: dict) -> dict:
        """
        Validate that all Excel activities are properly scoped within ExcelApplicationCard.

        Returns a dict with:
        - is_valid: bool indicating if scoping is valid
        - invalid_activities: list of activities with invalid scoping
        - error_response: structured error JSON if invalid (None if valid)
        """
        invalid_activities = []

        def check_scoping(node: Any, parent_type: str = "root", in_excel_card: bool = False) -> None:
            """Recursively check Excel activity scoping."""
            if not isinstance(node, dict):
                return

            activity_type = node.get('type', '')
            current_in_excel_card = in_excel_card or activity_type == 'ExcelApplicationCard'

            # Check if this is an Excel activity that requires scoping
            if activity_type in EXCEL_SCOPED_ACTIVITIES:
                if not in_excel_card:
                    invalid_activities.append({
                        "type": activity_type,
                        "displayName": node.get('displayName', activity_type),
                        "current_parent": parent_type
                    })

            # Recursively check children
            for key in ['activities', 'children']:
                if key in node and isinstance(node[key], list):
                    for child in node[key]:
                        check_scoping(child, activity_type, current_in_excel_card)

            # Check body (for scope activities)
            if 'body' in node and isinstance(node['body'], dict):
                check_scoping(node['body'], activity_type, current_in_excel_card)

            # Check then/else branches (for If)
            for branch in ['then', 'else']:
                if branch in node and isinstance(node[branch], dict):
                    check_scoping(node[branch], activity_type, current_in_excel_card)

            # Check try/catches/finally (for TryCatch)
            if 'try' in node and isinstance(node['try'], dict):
                check_scoping(node['try'], activity_type, current_in_excel_card)
            if 'catches' in node and isinstance(node['catches'], list):
                for catch in node['catches']:
                    if isinstance(catch, dict):
                        check_scoping(catch, activity_type, current_in_excel_card)
            if 'finally' in node and isinstance(node['finally'], dict):
                check_scoping(node['finally'], activity_type, current_in_excel_card)

        # Start validation from workflow root
        check_scoping(workflow_json)

        if invalid_activities:
            error_response = {
                "status": "error",
                "error_type": "scoping_violation",
                "rule": "Excel activities must be within ExcelApplicationCard scope",
                "invalid_activities": invalid_activities,
                "details": f"{len(invalid_activities)} Excel activit{'y' if len(invalid_activities) == 1 else 'ies'} found outside ExcelApplicationCard scope",
                "fix": "Wrap Excel activities in ExcelApplicationCard container within ExcelProcessScopeX",
                "retry_suggestion": "ExcelProcessScopeX → Sequence → ExcelApplicationCard → Sequence → [Excel activities]",
                "expected_structure": {
                    "outer": "ExcelProcessScopeX",
                    "middle": "ExcelApplicationCard",
                    "inner": "Sequence containing Excel activities"
                }
            }
            return {
                "is_valid": False,
                "invalid_activities": invalid_activities,
                "error_response": error_response
            }

        return {
            "is_valid": True,
            "invalid_activities": [],
            "error_response": None
        }

    def contains_flowchart(self, node: Any) -> bool:
        """Recursively check if workflow contains Flowchart activity."""
        if not isinstance(node, dict):
            return False

        if node.get('type') == 'Flowchart':
            return True

        # Check nested structures
        for key in ['activities', 'nodes', 'body', 'then', 'else', 'try', 'catches', 'finally', 'activity']:
            if key in node:
                if isinstance(node[key], list):
                    for child in node[key]:
                        if self.contains_flowchart(child):
                            return True
                elif isinstance(node[key], dict):
                    if self.contains_flowchart(node[key]):
                        return True

        return False

    def validate_flowchart_structure(self, workflow_json: dict) -> dict:
        """
        Validate flowchart structure including reference IDs, ViewState,
        structural constraints, circular references, and reachability.

        Returns a dict with:
        - is_valid: bool
        - validation_failures: list of failure dicts
        - error_response: structured error JSON if invalid (None if valid)
        - modified_json: workflow_json with IDs and ViewState added
        """
        validation_failures = []
        modified = copy.deepcopy(workflow_json)

        flowcharts = []

        def find_and_validate_structure(node, parent_type="root"):
            """Find Flowcharts and check structural placement rules."""
            if not isinstance(node, dict):
                return

            node_type = node.get('type', '')

            # Structural rule: FlowStep/FlowDecision must be in Flowchart
            if node_type in ('FlowStep', 'FlowDecision') and parent_type != 'Flowchart':
                validation_failures.append({
                    'category': 'structural',
                    'rule': f'{node_type} must be within Flowchart container',
                    'details': f'{node_type} found in {parent_type}',
                    'affected_nodes': [node.get('x:Name', 'unnamed')]
                })

            if node_type == 'Flowchart':
                flowcharts.append(node)

            # Recurse into children
            for key in ['activities', 'nodes']:
                if key in node and isinstance(node[key], list):
                    for child in node[key]:
                        find_and_validate_structure(child, node_type)

            if 'body' in node and isinstance(node['body'], dict):
                find_and_validate_structure(node['body'], node_type)

            # Check 'activity' key (used by some scope bodies)
            if 'activity' in node:
                if isinstance(node['activity'], dict):
                    find_and_validate_structure(node['activity'], node_type)
                elif isinstance(node['activity'], list):
                    for child in node['activity']:
                        if isinstance(child, dict):
                            find_and_validate_structure(child, node_type)

            for branch in ['then', 'else']:
                if branch in node and isinstance(node[branch], dict):
                    find_and_validate_structure(node[branch], node_type)

            if 'try' in node and isinstance(node['try'], dict):
                find_and_validate_structure(node['try'], node_type)
            if 'catches' in node and isinstance(node['catches'], list):
                for catch_block in node['catches']:
                    if isinstance(catch_block, dict):
                        find_and_validate_structure(catch_block, node_type)
            if 'finally' in node and isinstance(node['finally'], dict):
                find_and_validate_structure(node['finally'], node_type)

        # Traverse modified JSON to find flowcharts and check structure
        find_and_validate_structure(modified)

        # Process each Flowchart
        for flowchart in flowcharts:
            nodes = flowchart.get('nodes', [])

            # 1.2 Reference ID Assignment — build old-to-new mapping
            id_mapping = {}  # old_id -> new_id
            counter = 0
            for node in nodes:
                old_id = node.get('x:Name')
                new_id = f"__ReferenceID{counter}"
                if old_id is not None:
                    id_mapping[old_id] = new_id
                node['x:Name'] = new_id
                counter += 1

            # Remap startNode reference to new ID
            old_start = flowchart.get('startNode')
            if old_start and old_start in id_mapping:
                flowchart['startNode'] = id_mapping[old_start]

            # Remap each node's next/true/false references to new IDs
            for node in nodes:
                node_type = node.get('type', '')
                if node_type == 'FlowStep':
                    old_next = node.get('next')
                    if old_next and old_next in id_mapping:
                        node['next'] = id_mapping[old_next]
                elif node_type == 'FlowDecision':
                    for branch in ['true', 'false']:
                        old_ref = node.get(branch)
                        if old_ref and old_ref in id_mapping:
                            node[branch] = id_mapping[old_ref]

            # Build reference ID set
            all_ref_ids = {node.get('x:Name') for node in nodes}

            # Build node lookup by ref ID
            node_by_ref = {node.get('x:Name'): node for node in nodes}

            # Build index lookup by ref ID
            idx_by_ref = {}
            for idx, node in enumerate(nodes):
                idx_by_ref[node.get('x:Name')] = idx

            # 1.4 Structural Validation - startNode
            start_node = flowchart.get('startNode')
            if not start_node:
                validation_failures.append({
                    'category': 'structural',
                    'rule': 'Flowchart must have startNode property',
                    'details': 'startNode is missing or null',
                    'affected_nodes': [flowchart.get('displayName', 'Flowchart')]
                })
            else:
                if start_node not in all_ref_ids:
                    validation_failures.append({
                        'category': 'structural',
                        'rule': 'StartNode must reference a valid node',
                        'details': f"startNode '{start_node}' does not match any node reference ID",
                        'affected_nodes': [start_node]
                    })
                else:
                    start_obj = node_by_ref.get(start_node)
                    if start_obj and start_obj.get('type') != 'FlowStep':
                        validation_failures.append({
                            'category': 'structural',
                            'rule': 'StartNode must reference a FlowStep',
                            'details': f"startNode '{start_node}' references a {start_obj.get('type')}, not a FlowStep",
                            'affected_nodes': [start_node]
                        })

            # 1.5 Reference Validation
            # Uniqueness check
            seen_ids = set()
            for node in nodes:
                ref_id = node.get('x:Name')
                if ref_id in seen_ids:
                    validation_failures.append({
                        'category': 'reference',
                        'rule': 'Reference IDs must be unique',
                        'details': f"Duplicate reference ID: {ref_id}",
                        'affected_nodes': [ref_id]
                    })
                seen_ids.add(ref_id)

            # Format check
            for node in nodes:
                ref_id = node.get('x:Name', '')
                if not re.match(r'^__ReferenceID\d+$', ref_id):
                    validation_failures.append({
                        'category': 'reference',
                        'rule': 'Reference ID must match pattern __ReferenceID\\d+',
                        'details': f"Invalid reference ID format: {ref_id}",
                        'affected_nodes': [ref_id]
                    })

            # Dangling reference check
            for node in nodes:
                node_type = node.get('type', '')
                ref_id = node.get('x:Name', '')

                if node_type == 'FlowStep':
                    next_ref = node.get('next')
                    if next_ref is not None and next_ref not in all_ref_ids:
                        validation_failures.append({
                            'category': 'reference',
                            'rule': 'Next reference must point to existing node',
                            'details': f"Reference {next_ref} not found",
                            'affected_nodes': [ref_id]
                        })
                elif node_type == 'FlowDecision':
                    for branch in ['true', 'false']:
                        branch_ref = node.get(branch)
                        if branch_ref is not None and branch_ref not in all_ref_ids:
                            validation_failures.append({
                                'category': 'reference',
                                'rule': f'{branch.capitalize()} reference must point to existing node',
                                'details': f"Reference {branch_ref} not found",
                                'affected_nodes': [ref_id]
                            })

            # 1.6 Circular Reference Detection (DFS)
            graph = {}
            for node in nodes:
                ref_id = node.get('x:Name', '')
                neighbors = []
                node_type = node.get('type', '')

                if node_type == 'FlowStep':
                    next_ref = node.get('next')
                    if next_ref and next_ref in all_ref_ids:
                        neighbors.append(next_ref)
                elif node_type == 'FlowDecision':
                    for branch in ['true', 'false']:
                        branch_ref = node.get(branch)
                        if branch_ref and branch_ref in all_ref_ids:
                            neighbors.append(branch_ref)

                graph[ref_id] = neighbors

            # DFS cycle detection
            visited = set()
            rec_stack = set()

            def detect_cycle(node_id, path):
                visited.add(node_id)
                rec_stack.add(node_id)
                path.append(node_id)

                for neighbor in graph.get(node_id, []):
                    if neighbor not in visited:
                        cycle = detect_cycle(neighbor, path)
                        if cycle:
                            return cycle
                    elif neighbor in rec_stack:
                        cycle_start = path.index(neighbor)
                        return path[cycle_start:] + [neighbor]

                rec_stack.remove(node_id)
                path.pop()
                return None

            for node_id in graph:
                if node_id not in visited:
                    circular_path = detect_cycle(node_id, [])
                    if circular_path:
                        validation_failures.append({
                            'category': 'circular',
                            'rule': 'Flowchart must not contain circular references',
                            'details': f"Circular path detected: {' \u2192 '.join(circular_path)}",
                            'affected_nodes': circular_path[:-1]
                        })

            # 1.7 Reachability Analysis
            if start_node and start_node in all_ref_ids:
                reachable = set()
                queue = [start_node]

                while queue:
                    current = queue.pop(0)
                    if current in reachable or current is None:
                        continue
                    reachable.add(current)

                    for neighbor in graph.get(current, []):
                        if neighbor and neighbor not in reachable:
                            queue.append(neighbor)

                unreachable = all_ref_ids - reachable
                if unreachable:
                    validation_failures.append({
                        'category': 'reachability',
                        'rule': 'All nodes must be reachable from StartNode',
                        'details': f"{len(unreachable)} orphaned node(s)",
                        'affected_nodes': list(unreachable)
                    })

            # 1.3 ViewState Generation
            def get_node_position(ref_id):
                """Get (x, y, width, height) for a node by its reference ID."""
                target = node_by_ref.get(ref_id)
                if not target:
                    return None
                target_idx = idx_by_ref.get(ref_id, 0)
                target_type = target.get('type', '')
                if target_type == 'FlowStep':
                    return (300, 200 + target_idx * 100, 110, 70)
                elif target_type == 'FlowDecision':
                    return (325, 200 + target_idx * 100, 60, 60)
                return None

            # Flowchart container start node ViewState
            start_vs = {
                'ShapeLocation': '330,10',
                'ShapeSize': '50,50'
            }
            if start_node and start_node in all_ref_ids:
                start_pos = get_node_position(start_node)
                if start_pos:
                    target_cx_top = start_pos[0] + start_pos[2] // 2
                    target_y = start_pos[1]
                    start_vs['ConnectorLocation'] = f"355,60 {target_cx_top},{target_y}"

            flowchart['viewState'] = start_vs

            # Node ViewState
            for idx, node in enumerate(nodes):
                node_type = node.get('type', '')

                if node_type == 'FlowStep':
                    x = 300
                    y = 200 + idx * 100
                    width = 110
                    height = 70
                    cx_bottom = x + width // 2  # 355
                    cy_bottom = y + height

                    vs = {
                        'ShapeLocation': f'{x},{y}',
                        'ShapeSize': f'{width},{height}'
                    }

                    next_ref = node.get('next')
                    if next_ref and next_ref in all_ref_ids:
                        next_pos = get_node_position(next_ref)
                        if next_pos:
                            next_cx_top = next_pos[0] + next_pos[2] // 2
                            next_y = next_pos[1]
                            vs['ConnectorLocation'] = f"{cx_bottom},{cy_bottom} {next_cx_top},{next_y}"

                    node['viewState'] = vs

                elif node_type == 'FlowDecision':
                    x = 325
                    y = 200 + idx * 100
                    width = 60
                    height = 60
                    cy = y + height // 2

                    vs = {
                        'ShapeLocation': f'{x},{y}',
                        'ShapeSize': f'{width},{height}'
                    }

                    # True connector (branch left)
                    true_ref = node.get('true')
                    if true_ref and true_ref in all_ref_ids:
                        true_pos = get_node_position(true_ref)
                        if true_pos:
                            true_y = true_pos[1]
                            vs['TrueConnector'] = f"{x},{cy} 150,{cy} 150,{true_y}"

                    # False connector (branch right)
                    false_ref = node.get('false')
                    if false_ref and false_ref in all_ref_ids:
                        false_pos = get_node_position(false_ref)
                        if false_pos:
                            false_y = false_pos[1]
                            vs['FalseConnector'] = f"{x + width},{cy} 560,{cy} 560,{false_y}"

                    node['viewState'] = vs

        # Build result
        if validation_failures:
            categories = [f['category'] for f in validation_failures]
            if 'structural' in categories:
                fix = "Nest FlowStep/FlowDecision inside Flowchart container"
                retry_suggestion = "Flowchart \u2192 nodes: [FlowStep, FlowDecision]"
            elif 'circular' in categories:
                cycle_nodes = []
                for f in validation_failures:
                    if f['category'] == 'circular':
                        cycle_nodes = f['affected_nodes']
                        break
                fix = "Break circular path by removing or redirecting one connection"
                retry_suggestion = f"Review path: {' \u2192 '.join(cycle_nodes)} and set one 'next' to null"
            elif 'reachability' in categories:
                orphaned = []
                for f in validation_failures:
                    if f['category'] == 'reachability':
                        orphaned = f['affected_nodes']
                        break
                fix = "Connect orphaned nodes to flowchart or remove them"
                retry_suggestion = f"Add reference from existing node to: {', '.join(orphaned)}"
            else:
                fix = "Ensure all reference IDs are unique and properly formatted"
                retry_suggestion = "Use sequential IDs: __ReferenceID0, __ReferenceID1, etc."

            error_response = {
                'status': 'error',
                'error_type': 'flowchart_validation_failure',
                'validation_failures': validation_failures,
                'fix': fix,
                'retry_suggestion': retry_suggestion
            }

            return {
                'is_valid': False,
                'validation_failures': validation_failures,
                'error_response': error_response,
                'modified_json': modified
            }

        return {
            'is_valid': True,
            'validation_failures': [],
            'error_response': None,
            'modified_json': modified
        }

    def build_activity(self, activity_json: dict) -> dict:
        """Build and validate activity from JSON specification."""
        is_valid, errors = self.validate_activity(activity_json)
        if not is_valid:
            raise ValueError(f"Invalid activity: {'; '.join(errors)}")

        # Return the validated activity JSON
        # (actual XAML construction is done by xaml_syntaxer.py)
        return activity_json

    def build_from_template(self, activity_type: str, **kwargs) -> dict:
        """Build activity from template with custom values."""
        template = self.template_loader.get_template(activity_type)
        if not template:
            available = self.template_loader.get_activity_types()
            raise ValueError(f"Unknown activity type: {activity_type}. Available: {', '.join(sorted(available))}")

        # Deep copy template
        activity = copy.deepcopy(template.get('template', {}))

        # Override with provided values
        for key, value in kwargs.items():
            activity[key] = value

        # Validate
        is_valid, errors = self.validate_activity(activity)
        if not is_valid:
            raise ValueError(f"Invalid activity: {'; '.join(errors)}")

        return activity

    def get_template_info(self, activity_type: str) -> Optional[dict]:
        """Get metadata about an activity type."""
        template = self.template_loader.get_template(activity_type)
        if not template:
            return None

        return {
            'type': template.get('type', ''),
            'displayName': template.get('displayName', ''),
            'description': template.get('description', ''),
            'namespace': template.get('namespace', 'default'),
            'requiredAttributes': template.get('requiredAttributes', []),
            'optionalAttributes': template.get('optionalAttributes', []),
            'template': template.get('template', {})
        }


DEFAULT_METADATA = {
    "class": "",
    "namespaces": [],
    "assemblyReferences": [],
    "arguments": []
}


def preserve_metadata(input_json: dict) -> dict:
    """
    Extract and preserve metadata from input JSON.
    Returns metadata dict if present, or default empty metadata.

    For EDIT workflows (Reader output), metadata contains namespace declarations
    and assembly references that must be preserved through to the Writer.
    For NEW workflows, returns default empty metadata.
    """
    if not isinstance(input_json, dict):
        return copy.deepcopy(DEFAULT_METADATA)

    metadata = input_json.get('metadata')
    if metadata and isinstance(metadata, dict):
        preserved = copy.deepcopy(metadata)
        # Ensure required keys exist
        for key in ('class', 'namespaces', 'assemblyReferences', 'arguments'):
            if key not in preserved:
                preserved[key] = DEFAULT_METADATA[key]
        # Filter out whitespace-only assembly references
        preserved['assemblyReferences'] = [
            r for r in preserved.get('assemblyReferences', [])
            if r and isinstance(r, str) and r.strip()
        ]
        return preserved

    return copy.deepcopy(DEFAULT_METADATA)


def list_activities(template_loader: TemplateLoader, output_file: Optional[str] = None) -> None:
    """List all available activities organized by category."""
    categorized = template_loader.list_all_templates()

    output = {
        "categories": categorized,
        "total_activities": sum(len(activities) for activities in categorized.values()),
        "namespace_mapping": NAMESPACE_MAPPING
    }

    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(output, f, indent=2)
        print(f"Activity list written to {output_file}")
    else:
        print(json.dumps(output, indent=2))


def build_activities(input_data: str, output_file: Optional[str] = None, validate_scoping: bool = True, validate_flowcharts: bool = True) -> None:
    """Build activities from JSON input.

    Supports two input formats:
    1. Activity-only JSON (NEW workflow): single activity dict or list of activities
    2. Reader output JSON (EDIT workflow): dict with 'metadata' and 'body'/'workflow' keys

    Return shape contract:
    - NEW workflows (no metadata in input): returns activity-only JSON
      (single activity dict or list of activities, matching the input shape)
    - EDIT workflows (metadata present): returns {"metadata": ..., "workflow": ...}
      envelope with preserved metadata for Writer consumption

    For EDIT workflows, metadata (namespaces, assemblyReferences) is preserved
    unchanged in the output, ensuring the Writer can reconstruct valid XAML.
    """
    template_loader = TemplateLoader(TEMPLATE_DIR)
    builder = ActivityBuilder(template_loader)

    # Parse input (file or JSON string)
    input_path = Path(input_data)
    if input_path.exists():
        with open(input_path, 'r', encoding='utf-8') as f:
            activities_json = json.load(f)
    else:
        try:
            activities_json = json.loads(input_data)
        except json.JSONDecodeError as e:
            print(f"Error: Invalid JSON input: {e}", file=sys.stderr)
            sys.exit(1)

    # Detect EDIT vs NEW workflow by checking for metadata
    input_metadata = preserve_metadata(activities_json)
    has_metadata = (isinstance(activities_json, dict) and
                    'metadata' in activities_json and
                    isinstance(activities_json.get('metadata'), dict))

    # Extract workflow activities from input
    # EDIT workflow: activities are under 'body' or 'workflow' key
    # NEW workflow: input IS the activity/activities
    if has_metadata:
        workflow_data = activities_json.get('body') or activities_json.get('workflow')
        if workflow_data is None:
            print("Error: EDIT workflow JSON must contain 'body' or 'workflow' key alongside 'metadata'", file=sys.stderr)
            sys.exit(1)
        ns_count = len(input_metadata.get('namespaces', []))
        asm_count = len(input_metadata.get('assemblyReferences', []))
        print(f"Metadata preserved: {ns_count} namespaces, {asm_count} assembly references", file=sys.stderr)
    else:
        workflow_data = activities_json
        print("New workflow: using default metadata", file=sys.stderr)

    # Handle single activity or list of activities
    original_is_single = isinstance(workflow_data, dict)
    if original_is_single:
        activities_list = [workflow_data]
    else:
        activities_list = workflow_data

    # Build and validate each activity
    built_activities: List[dict] = []
    for i, activity in enumerate(activities_list):
        try:
            built = builder.build_activity(activity)
            built_activities.append(built)
        except ValueError as e:
            print(f"Error building activity {i + 1}: {e}", file=sys.stderr)
            sys.exit(1)

    # Validate Excel scoping if enabled
    if validate_scoping:
        for i, activity in enumerate(built_activities):
            scoping_result = builder.validate_excel_scoping(activity)
            if not scoping_result["is_valid"]:
                # Output structured error JSON
                print(json.dumps(scoping_result["error_response"], indent=2))
                sys.exit(1)

    # Validate flowchart structure if Flowchart detected
    if validate_flowcharts:
        for i, activity in enumerate(built_activities):
            if builder.contains_flowchart(activity):
                flowchart_result = builder.validate_flowchart_structure(activity)
                if not flowchart_result['is_valid']:
                    # Output structured error JSON
                    print(json.dumps(flowchart_result['error_response'], indent=2))
                    sys.exit(1)
                else:
                    # Replace with modified JSON (includes IDs and ViewState)
                    built_activities[i] = flowchart_result['modified_json']

    # Construct output based on workflow type
    workflow_result = built_activities[0] if original_is_single else built_activities

    if has_metadata:
        # EDIT workflow: wrap in metadata/workflow envelope for Writer consumption
        result = {
            "metadata": input_metadata,
            "workflow": workflow_result
        }

        # Validate metadata preservation for EDIT workflows
        out_ns = len(result['metadata'].get('namespaces', []))
        out_asm = len(result['metadata'].get('assemblyReferences', []))
        if out_ns != ns_count or out_asm != asm_count:
            print(f"Warning: Metadata count mismatch! Input: {ns_count}ns/{asm_count}asm, Output: {out_ns}ns/{out_asm}asm", file=sys.stderr)
    else:
        # NEW workflow: return activity-only JSON (original API shape)
        result = workflow_result

    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2)
        print(f"Built activities written to {output_file}")
    else:
        print(json.dumps(result, indent=2))


def get_template(activity_type: str, output_file: Optional[str] = None) -> None:
    """Get the template for a specific activity type."""
    template_loader = TemplateLoader(TEMPLATE_DIR)
    builder = ActivityBuilder(template_loader)

    info = builder.get_template_info(activity_type)
    if not info:
        available = template_loader.get_activity_types()
        print(f"Error: Unknown activity type: {activity_type}", file=sys.stderr)
        if available:
            print(f"Available types: {', '.join(sorted(available))}", file=sys.stderr)
        sys.exit(1)

    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(info, f, indent=2)
        print(f"Template info written to {output_file}")
    else:
        print(json.dumps(info, indent=2))


def main():
    parser = argparse.ArgumentParser(
        description='UiPath Activity Constructor - Template library and builder',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  List all activities:
    python xaml_constructor.py --mode list

  Get template for specific activity:
    python xaml_constructor.py --mode template --type Assign

  Build activity from JSON file:
    python xaml_constructor.py --mode build --input activity.json

  Build activity from JSON string:
    python xaml_constructor.py --mode build --input '{"type":"Assign","displayName":"Set Var","to":"x","value":"1"}'
"""
    )
    parser.add_argument(
        '--mode',
        choices=['list', 'build', 'template'],
        required=True,
        help='Operation mode: list available activities, build from specification, or get template info'
    )
    parser.add_argument(
        '--input',
        help='Input JSON file or string (required for build mode)'
    )
    parser.add_argument(
        '--type',
        help='Activity type (required for template mode)'
    )
    parser.add_argument(
        '--output',
        help='Output file path (optional, prints to stdout if not specified)'
    )

    args = parser.parse_args()

    if args.mode == 'list':
        template_loader = TemplateLoader(TEMPLATE_DIR)
        list_activities(template_loader, args.output)
    elif args.mode == 'build':
        if not args.input:
            print("Error: --input required for build mode", file=sys.stderr)
            sys.exit(1)
        build_activities(args.input, args.output)
    elif args.mode == 'template':
        if not args.type:
            print("Error: --type required for template mode", file=sys.stderr)
            sys.exit(1)
        get_template(args.type, args.output)


if __name__ == '__main__':
    main()
