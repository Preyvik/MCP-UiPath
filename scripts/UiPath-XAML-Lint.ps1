<#
.SYNOPSIS
    UiPath XAML Syntax Checker / Linter (PowerShell) - Comprehensive Edition

.DESCRIPTION
    Validates UiPath workflow XAML files against deterministic hard rules
    without relying on UiPath Workflow Analyzer.

    Based on UiPath Studio XAML Hard Rules v1.0 and all domain agent skills.

    Additional validation rules:
    - HR-Namespace-Usage: Validates all used namespace prefixes have declarations
    - HR-Double-Encoding: Detects double-encoded entities in expressions
    - Namespace-Common-Types: Checks for common namespace requirements based on detected activity types

.PARAMETER Path
    Path to a XAML file or directory containing XAML files

.PARAMETER Strict
    Treat warnings as errors

.PARAMETER Json
    Output as JSON

.PARAMETER NamespaceRegistry
    Path to namespaces.json registry file for namespace validation

.PARAMETER ProjectJson
    Path to project.json for project context validation

.EXAMPLE
    .\UiPath-XAML-Lint.ps1 -Path "Main.xaml"

.EXAMPLE
    .\UiPath-XAML-Lint.ps1 -Path ".\workflows\" -Strict

.EXAMPLE
    .\UiPath-XAML-Lint.ps1 -Path "Main.xaml" -NamespaceRegistry ".\uipath_rules\namespaces.json"
#>

param(
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Path,

    [switch]$Strict,

    [switch]$Json,

    [string]$NamespaceRegistry,

    [string]$ProjectJson
)

# ============================================================================
# CONSTANTS
# ============================================================================

$Script:RequiredNamespaces = @("x", "sap2010")

$Script:PrimaryContainers = @("Sequence", "Flowchart", "StateMachine")

# Secrets patterns - check for hardcoded credentials
# NOTE: Patterns exclude {x:Null} (empty), [ (VB expression), and require word boundary before attribute name
# Using (?<!\w) for word boundary to avoid matching IsPassword="False" when looking for password=
$Script:SecretsPatterns = @(
    @{ Pattern = '(?<![a-zA-Z])password\s*=\s*"(?!\{x:Null\}|\[|False|True)[^"]+[a-zA-Z0-9]+"'; Name = 'Hardcoded password' },
    @{ Pattern = '(?<![a-zA-Z])Password\s*=\s*"(?!\{x:Null\}|\[|False|True)[^"]+"'; Name = 'Hardcoded Password property' },
    @{ Pattern = '(?<![a-zA-Z])apikey\s*=\s*"(?!\{x:Null\}|\[)[^"]+"'; Name = 'Hardcoded API key' },
    @{ Pattern = '(?<![a-zA-Z])ApiKey\s*=\s*"(?!\{x:Null\}|\[)[^"]+"'; Name = 'Hardcoded ApiKey property' },
    @{ Pattern = '(?<![a-zA-Z])api_key\s*=\s*"(?!\{x:Null\}|\[)[^"]+"'; Name = 'Hardcoded api_key' },
    @{ Pattern = '(?<![a-zA-Z])secret\s*=\s*"(?!\{x:Null\}|\[)[^"]+"'; Name = 'Hardcoded secret' },
    @{ Pattern = '(?<![a-zA-Z])Secret\s*=\s*"(?!\{x:Null\}|\[)[^"]+"'; Name = 'Hardcoded Secret property' },
    @{ Pattern = '(?<![a-zA-Z])token\s*=\s*"(?!\{x:Null\}|\[)[^"]{20,}"'; Name = 'Hardcoded token (long string)' },
    @{ Pattern = '(?<![a-zA-Z])Token\s*=\s*"(?!\{x:Null\}|\[)[^"]{20,}"'; Name = 'Hardcoded Token property' },
    @{ Pattern = 'bearer\s+[a-zA-Z0-9\-_\.]+'; Name = 'Hardcoded Bearer token' },
    @{ Pattern = 'Authorization.*Basic\s+[a-zA-Z0-9+/=]+'; Name = 'Hardcoded Basic auth' },
    @{ Pattern = 'connectionstring\s*=\s*"[^"]*password[^"]*"'; Name = 'Connection string with password' },
    @{ Pattern = 'ConnectionString\s*=\s*"[^"]*Password[^"]*"'; Name = 'ConnectionString with Password' },
    @{ Pattern = 'CredentialAssetName\s*=\s*""'; Name = 'Empty credential asset name' },
    @{ Pattern = '(?<![a-zA-Z])SecureString\s*=\s*"(?!\{x:Null\}|\[)[^"]+"'; Name = 'Hardcoded SecureString value' }
)

# Common type formats for validation
$Script:ValidTypePatterns = @(
    'x:String',
    'x:Int32',
    'x:Boolean',
    'x:Double',
    'x:Object',
    'x:Decimal',
    'System\.String',
    'System\.Int32',
    'System\.Boolean',
    'System\.Double',
    'System\.Object',
    'System\.DateTime',
    'System\.TimeSpan',
    'System\.Data\.DataTable',
    'System\.Data\.DataRow',
    'System\.Collections\.Generic\.List',
    'System\.Collections\.Generic\.Dictionary',
    'System\.Collections\.Generic\.IEnumerable',
    'scg:List',
    'scg:Dictionary',
    's:String\[\]',
    'sd:DataTable',
    'sd:DataRow'
)

# VB.NET expression indicators
$Script:VBExpressionPatterns = @(
    ' And ',
    ' Or ',
    ' Not ',
    ' AndAlso ',
    ' OrElse ',
    'String\.IsNullOrEmpty',
    'CStr\(',
    'CInt\(',
    'CBool\(',
    'CDate\(',
    'DirectCast\(',
    'TryCast\(',
    'GetType\(',
    'TypeOf .* Is ',
    ' & "',
    '" & ',
    'Nothing',
    '\.ToString\("',
    'Is Nothing',
    'IsNot Nothing'
)

# C# expression indicators
$Script:CSharpExpressionPatterns = @(
    ' && ',
    ' \|\| ',
    ' \? ',
    ' : ',
    'string\.IsNullOrEmpty',
    '\(string\)',
    '\(int\)',
    '\(bool\)',
    ' as ',
    ' is ',
    'typeof\(',
    ' \+ "',
    '" \+ ',
    'null',
    ' == null',
    ' != null'
)

# ============================================================================
# CLASSES
# ============================================================================

class ValidationIssue {
    [string]$Rule
    [string]$Severity  # ERROR, WARNING, INFO
    [string]$Message
    [int]$Line
    [string]$Element
}

class ValidationResult {
    [string]$FilePath
    [bool]$IsValid
    [System.Collections.ArrayList]$Issues
    [hashtable]$HRCompliance
    [hashtable]$ProjectContext

    ValidationResult() {
        $this.Issues = [System.Collections.ArrayList]::new()
        $this.HRCompliance = @{}
        $this.ProjectContext = @{}
    }

    [hashtable] ToHashtable() {
        return @{
            file_path = $this.FilePath
            is_valid = $this.IsValid
            error_count = ($this.Issues | Where-Object { $_.Severity -eq "ERROR" }).Count
            warning_count = ($this.Issues | Where-Object { $_.Severity -eq "WARNING" }).Count
            issues = $this.Issues | ForEach-Object {
                @{
                    rule = $_.Rule
                    severity = $_.Severity
                    message = $_.Message
                    line = $_.Line
                    element = $_.Element
                }
            }
            hr_compliance = $this.HRCompliance
            project_context = $this.ProjectContext
        }
    }
}

# ============================================================================
# PROJECT CONTEXT FUNCTIONS
# ============================================================================

function Get-ProjectContext {
    param([string]$XamlPath, [string]$ProjectJsonPath)

    $context = @{
        language = "Unknown"
        compatibility = "Unknown"
        dependencies = @()
        projectPath = $null
    }

    # Try to find project.json
    $searchPath = if ($ProjectJsonPath) { $ProjectJsonPath } else {
        $dir = Split-Path $XamlPath -Parent
        while ($dir) {
            $pjPath = Join-Path $dir "project.json"
            if (Test-Path $pjPath) {
                $pjPath
                break
            }
            $parent = Split-Path $dir -Parent
            if ($parent -eq $dir) { break }
            $dir = $parent
        }
    }

    if ($searchPath -and (Test-Path $searchPath)) {
        try {
            $pj = Get-Content $searchPath -Raw | ConvertFrom-Json
            $context.projectPath = $searchPath

            # Detect language from expressionLanguage or studioVersion
            if ($pj.expressionLanguage) {
                $context.language = if ($pj.expressionLanguage -match 'CSharp|C#') { "CSharp" } else { "VB" }
            }
            elseif ($pj.studioVersion) {
                # Older projects default to VB
                $context.language = "VB"
            }

            # Detect compatibility
            if ($pj.targetFramework) {
                if ($pj.targetFramework -match 'net6|net7|net8') {
                    $context.compatibility = "CrossPlatform"
                }
                elseif ($pj.targetFramework -match 'net461|net472|net48') {
                    $context.compatibility = "Windows-Legacy"
                }
                else {
                    $context.compatibility = "Windows"
                }
            }

            # Extract dependencies
            if ($pj.dependencies) {
                $context.dependencies = $pj.dependencies.PSObject.Properties | ForEach-Object {
                    @{ id = $_.Name; version = $_.Value }
                }
            }
        }
        catch {
            # Failed to parse project.json
        }
    }

    return $context
}

function Get-NamespaceRegistry {
    param([string]$RegistryPath)

    if (-not $RegistryPath -or -not (Test-Path $RegistryPath)) {
        return $null
    }

    try {
        return Get-Content $RegistryPath -Raw | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

function Test-XmlWellFormed {
    param([string]$Content, [ValidationResult]$Result)

    try {
        $null = [xml]$Content
        $Result.HRCompliance["Layer1-XML"] = "PASS"
        return $true
    }
    catch {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "Layer1-XML"
        $issue.Severity = "ERROR"
        $issue.Message = "XML parsing error: $($_.Exception.Message)"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["Layer1-XML"] = "FAIL"
        return $false
    }
}

function Get-Namespaces {
    param([string]$Content)

    $namespaces = @{}
    $pattern = "xmlns:?([a-zA-Z0-9_]*)=[`"']([^`"']+)[`"']"
    $matches = [regex]::Matches($Content, $pattern)

    foreach ($match in $matches) {
        $prefix = $match.Groups[1].Value
        $uri = $match.Groups[2].Value
        if (-not $prefix) { $prefix = "" }
        $namespaces[$prefix] = $uri
    }

    return $namespaces
}

function Test-HR0-TemplateStructure {
    param([string]$Content, [ValidationResult]$Result)

    $hasXClass = $Content -match 'x:Class='

    if (-not $hasXClass) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-0"
        $issue.Severity = "WARNING"
        $issue.Message = "Missing x:Class attribute - may not be a valid UiPath workflow template"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-0"] = "WARN"
    }
    else {
        $Result.HRCompliance["HR-0"] = "PASS"
    }
}

function Test-HR2-RootInvariants {
    param([string]$Content, [hashtable]$Namespaces, [ValidationResult]$Result, [hashtable]$Registry)

    $missingNs = @()
    foreach ($prefix in $Script:RequiredNamespaces) {
        if (-not $Namespaces.ContainsKey($prefix)) {
            $missingNs += $prefix
        }
    }

    if ($missingNs.Count -gt 0) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-2"
        $issue.Severity = "ERROR"
        $issue.Message = "Missing required namespace prefixes: $($missingNs -join ', ')"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-2"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-2"] = "PASS"
    }

    # Validate against namespace registry if provided
    if ($Registry) {
        $mismatchedNs = @()
        foreach ($prefix in $Namespaces.Keys) {
            if ($prefix -and $Registry.PSObject.Properties.Name -contains $prefix) {
                $expected = $Registry.$prefix
                $actual = $Namespaces[$prefix]
                if ($expected -ne $actual) {
                    $mismatchedNs += "$prefix (expected: $expected, got: $actual)"
                }
            }
        }
        if ($mismatchedNs.Count -gt 0) {
            $issue = [ValidationIssue]::new()
            $issue.Rule = "Namespace-Registry"
            $issue.Severity = "WARNING"
            $issue.Message = "Namespace URIs don't match registry: $($mismatchedNs -join '; ')"
            $Result.Issues.Add($issue) | Out-Null
        }
    }

    # Check for mc:Ignorable
    if ($Content -notmatch 'mc:Ignorable' -and $Namespaces.ContainsKey("mc")) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-2"
        $issue.Severity = "INFO"
        $issue.Message = "mc:Ignorable attribute not found (may be expected in some templates)"
        $Result.Issues.Add($issue) | Out-Null
    }
}

function Test-NamespacePrefixUsage {
    param([string]$Content, [hashtable]$Namespaces, [ValidationResult]$Result)

    $usedPrefixes = @{}

    # Pattern 1: Element tags - <prefix:ElementName
    $elementMatches = [regex]::Matches($Content, '<([a-zA-Z0-9_]+):[a-zA-Z0-9_]')
    foreach ($match in $elementMatches) {
        $prefix = $match.Groups[1].Value
        # Exclude 'xml' (reserved) and 'x' namespace property syntax like <x:Property>
        if ($prefix -ne 'xml') {
            $usedPrefixes[$prefix] = $true
        }
    }

    # Pattern 2: Type attributes - Type="prefix:TypeName" and Type="InArgument(prefix:TypeName)"
    # Scan the full Type="..." value and extract ALL prefixed types, including those
    # nested inside argument wrappers like InArgument(sd:DataTable) or OutArgument(scg:List(x:String))
    $typeAttrMatches = [regex]::Matches($Content, 'Type="([^"]+)"')
    foreach ($match in $typeAttrMatches) {
        $typeValue = $match.Groups[1].Value
        $innerTypePrefixes = [regex]::Matches($typeValue, '([A-Za-z0-9_]+):')
        foreach ($innerMatch in $innerTypePrefixes) {
            $usedPrefixes[$innerMatch.Groups[1].Value] = $true
        }
    }

    # Pattern 3: TypeArguments - x:TypeArguments="prefix:TypeName"
    $typeArgMatches = [regex]::Matches($Content, 'x:TypeArguments="([^"]+)"')
    foreach ($match in $typeArgMatches) {
        $typeArgValue = $match.Groups[1].Value
        # Extract all prefixes from the TypeArguments value (may contain multiple like "scg:List(sd:DataRow)")
        $innerPrefixes = [regex]::Matches($typeArgValue, '([a-zA-Z0-9_]+):')
        foreach ($innerMatch in $innerPrefixes) {
            $usedPrefixes[$innerMatch.Groups[1].Value] = $true
        }
    }

    # Compare used prefixes against declared namespaces
    $missingPrefixes = @()
    foreach ($prefix in $usedPrefixes.Keys) {
        if (-not $Namespaces.ContainsKey($prefix)) {
            $missingPrefixes += $prefix
        }
    }

    if ($missingPrefixes.Count -gt 0) {
        $displayPrefixes = ($missingPrefixes | Sort-Object) -join ', '
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-Namespace-Usage"
        $issue.Severity = "ERROR"
        $issue.Message = "Namespace prefix(es) used but not declared: $displayPrefixes"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-Namespace-Usage"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-Namespace-Usage"] = "PASS"
    }
}

function Test-DoubleEncoding {
    param([string]$Content, [ValidationResult]$Result)

    $doubleEncodingIssues = @()

    # Double-encoded entity patterns to detect
    $doubleEncodedPatterns = @(
        @{ Pattern = '&amp;quot;'; Display = '&amp;quot; (should be &quot;)' },
        @{ Pattern = '&amp;lt;'; Display = '&amp;lt; (should be &lt;)' },
        @{ Pattern = '&amp;gt;'; Display = '&amp;gt; (should be &gt;)' },
        @{ Pattern = '&amp;amp;'; Display = '&amp;amp; (should be &amp;)' }
    )

    # Expression attribute names to check (where double-encoding breaks evaluation)
    $exprAttrPattern = '(?:Message|Value|Condition|Expression|Text|Code)="([^"]*)"'
    $exprMatches = [regex]::Matches($Content, $exprAttrPattern)

    foreach ($match in $exprMatches) {
        $attrValue = $match.Groups[1].Value
        foreach ($dp in $doubleEncodedPatterns) {
            if ($attrValue -match [regex]::Escape($dp.Pattern)) {
                $truncated = if ($attrValue.Length -gt 60) { $attrValue.Substring(0, 60) + "..." } else { $attrValue }
                $doubleEncodingIssues += "Double-encoding $($dp.Display) in: $truncated"
            }
        }
    }

    # Exclude UI selector attributes where &amp;amp; is actually correct
    # These are NOT expression attributes, so they won't match the pattern above,
    # but add an explicit exclusion scan as a safety measure
    $selectorAttrPattern = '(?:Selector|FullSelectorArgument|FuzzySelectorArgument)="([^"]*)"'
    $selectorMatches = [regex]::Matches($Content, $selectorAttrPattern)
    # (No action needed - selectors are already excluded by only scanning expression attributes)

    # ---- Element body scan for argument/value nodes ----
    # Detect double-encoded entities inside element text (not attributes) for known
    # argument and value elements: <InArgument>, <OutArgument>, <Assign.Value>,
    # <Assign.To>, <ui:LogMessage ...>...</ui:LogMessage>, etc.
    # Selector elements are explicitly excluded.
    $selectorElementNames = @('Selector', 'FullSelectorArgument', 'FuzzySelectorArgument')

    # Targeted element names whose text content may contain expressions
    $argElementPattern = '<((?:In|Out|InOut)Argument[^>]*|Assign\.Value|Assign\.To|ui:LogMessage[^>]*)>([^<]+)</'
    $argBodyMatches = [regex]::Matches($Content, $argElementPattern)

    foreach ($abm in $argBodyMatches) {
        $elemTag = $abm.Groups[1].Value
        $elemText = $abm.Groups[2].Value

        # Skip selector elements
        $isSelectorElem = $false
        foreach ($selName in $selectorElementNames) {
            if ($elemTag -match [regex]::Escape($selName)) {
                $isSelectorElem = $true
                break
            }
        }
        if ($isSelectorElem) { continue }

        foreach ($dp in $doubleEncodedPatterns) {
            if ($elemText -match [regex]::Escape($dp.Pattern)) {
                $truncated = if ($elemText.Length -gt 60) { $elemText.Substring(0, 60) + "..." } else { $elemText }
                $doubleEncodingIssues += "Double-encoding $($dp.Display) in element body: $truncated"
            }
        }
    }

    # Broader fallback: scan all element text via >content< for any remaining bodies
    $genericBodyPattern = '>([^<]{4,})</'
    $genericBodyMatches = [regex]::Matches($Content, $genericBodyPattern)

    foreach ($gbm in $genericBodyMatches) {
        $bodyText = $gbm.Groups[1].Value

        # Skip if this is inside a selector element (check surrounding context)
        $contextStart = [Math]::Max(0, $gbm.Index - 200)
        $contextLen = [Math]::Min(200, $gbm.Index - $contextStart)
        $precedingContext = $Content.Substring($contextStart, $contextLen)
        $isSelectorBody = $false
        foreach ($selName in $selectorElementNames) {
            if ($precedingContext -match "<[^>]*$([regex]::Escape($selName))[^>]*>\s*$") {
                $isSelectorBody = $true
                break
            }
        }
        if ($isSelectorBody) { continue }

        foreach ($dp in $doubleEncodedPatterns) {
            if ($bodyText -match [regex]::Escape($dp.Pattern)) {
                $truncated = if ($bodyText.Length -gt 60) { $bodyText.Substring(0, 60) + "..." } else { $bodyText }
                $issueMsg = "Double-encoding $($dp.Display) in element body: $truncated"
                # Avoid duplicate reports from the targeted scan above
                if ($doubleEncodingIssues -notcontains $issueMsg) {
                    $doubleEncodingIssues += $issueMsg
                }
            }
        }
    }

    if ($doubleEncodingIssues.Count -gt 0) {
        $displayIssues = if ($doubleEncodingIssues.Count -gt 3) { ($doubleEncodingIssues[0..2] -join '; ') + "..." } else { $doubleEncodingIssues -join '; ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-Double-Encoding"
        $issue.Severity = "ERROR"
        $issue.Message = "Double-encoding detected in expression attributes: $displayIssues"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-Double-Encoding"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-Double-Encoding"] = "PASS"
    }
}

function Test-CommonNamespaceRequirements {
    param([string]$Content, [hashtable]$Namespaces, [ValidationResult]$Result)

    # Mapping of usage patterns to required namespace declarations
    $namespaceRequirements = @(
        @{
            Pattern = 'sd:(?:DataTable|DataRow)'
            Prefix = 'sd'
            ExpectedUri = 'clr-namespace:System.Data;assembly=System.Data.Common'
            Description = 'System.Data types (sd:DataTable, sd:DataRow)'
        },
        @{
            Pattern = 's:(?:String\[\]|Int32|Boolean|Double|DateTime|TimeSpan|Decimal)'
            Prefix = 's'
            ExpectedUri = 'clr-namespace:System;assembly=System.Private.CoreLib'
            Description = 'System primitive types (s:String[], s:Int32, etc.)'
        },
        @{
            Pattern = '<ue:'
            Prefix = 'ue'
            ExpectedUri = 'clr-namespace:UiPath.Excel;assembly=UiPath.Excel.Activities'
            Description = 'UiPath.Excel types'
        },
        @{
            Pattern = '<ueab:'
            Prefix = 'ueab'
            ExpectedUri = 'clr-namespace:UiPath.Excel.Activities.Business;assembly=UiPath.Excel.Activities'
            Description = 'UiPath.Excel.Activities.Business types'
        },
        @{
            Pattern = 'scg:(?:List|Dictionary|IEnumerable|KeyValuePair)'
            Prefix = 'scg'
            ExpectedUri = 'clr-namespace:System.Collections.Generic;assembly=System.Private.CoreLib'
            Description = 'System.Collections.Generic types (scg:List, scg:Dictionary)'
        },
        @{
            Pattern = 'sco:(?:Collection|ObservableCollection|ReadOnlyCollection)'
            Prefix = 'sco'
            ExpectedUri = 'clr-namespace:System.Collections.ObjectModel;assembly=System.Private.CoreLib'
            Description = 'System.Collections.ObjectModel types (sco:Collection)'
        }
    )

    $nsIssues = @()

    foreach ($req in $namespaceRequirements) {
        if ($Content -match $req.Pattern) {
            if (-not $Namespaces.ContainsKey($req.Prefix)) {
                $nsIssues += "Namespace prefix '$($req.Prefix)' required for $($req.Description) but not declared"
            }
            elseif ($Namespaces[$req.Prefix] -ne $req.ExpectedUri) {
                $nsIssues += "Namespace '$($req.Prefix)' URI mismatch for $($req.Description): expected '$($req.ExpectedUri)', got '$($Namespaces[$req.Prefix])'"
            }
        }
    }

    if ($nsIssues.Count -gt 0) {
        $displayIssues = if ($nsIssues.Count -gt 3) { ($nsIssues[0..2] -join '; ') + "..." } else { $nsIssues -join '; ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "Namespace-Common-Types"
        $issue.Severity = "WARNING"
        $issue.Message = "Common namespace issues: $displayIssues"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["Namespace-Common-Types"] = "WARN"
    }
    else {
        $Result.HRCompliance["Namespace-Common-Types"] = "PASS"
    }
}

function Test-HR3-PrimaryContainer {
    param([string]$Content, [ValidationResult]$Result)

    $containersFound = @()
    foreach ($container in $Script:PrimaryContainers) {
        $pattern = "<$container[\s>]"
        $matches = [regex]::Matches($Content, $pattern)
        if ($matches.Count -gt 0) {
            for ($i = 0; $i -lt $matches.Count; $i++) {
                $containersFound += $container
            }
        }
    }

    if ($containersFound.Count -eq 0) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-3"
        $issue.Severity = "ERROR"
        $issue.Message = "No primary container found. Expected one of: $($Script:PrimaryContainers -join ', ')"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-3"] = "FAIL"
    }
    elseif ($containersFound.Count -gt 1) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-3"
        $issue.Severity = "INFO"
        $issue.Message = "Multiple containers found: $($containersFound -join ', '). Ensure proper nesting."
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-3"] = "PASS"
    }
    else {
        $Result.HRCompliance["HR-3"] = "PASS"
    }
}

function Test-HR4-ArgumentsVariables {
    param([string]$Content, [ValidationResult]$Result)

    # Extract arguments from x:Property
    $argPattern = '<x:Property\s+Name="([^"]+)"'
    $arguments = [regex]::Matches($Content, $argPattern) | ForEach-Object { $_.Groups[1].Value }

    # Check for proper prefixes
    $invalidArgs = @()
    foreach ($arg in $arguments) {
        if (-not ($arg.StartsWith("in_") -or $arg.StartsWith("out_") -or $arg.StartsWith("io_"))) {
            $invalidArgs += $arg
        }
    }

    if ($invalidArgs.Count -gt 0) {
        $displayArgs = if ($invalidArgs.Count -gt 5) { ($invalidArgs[0..4] -join ', ') + "..." } else { $invalidArgs -join ', ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-4"
        $issue.Severity = "WARNING"
        $issue.Message = "Arguments missing standard prefix (in_/out_/io_): $displayArgs"
        $Result.Issues.Add($issue) | Out-Null
    }

    # Extract variables
    $varPattern = '<Variable\s+[^>]*Name="([^"]+)"'
    $variables = [regex]::Matches($Content, $varPattern) | ForEach-Object { $_.Groups[1].Value }

    # Check for shadowed names
    $shadowed = @()
    foreach ($var in $variables) {
        if ($arguments -contains $var) {
            $shadowed += $var
        }
    }

    if ($shadowed.Count -gt 0) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-4"
        $issue.Severity = "ERROR"
        $issue.Message = "Variables shadow argument names: $($shadowed -join ', ')"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-4"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-4"] = "PASS"
    }
}

function Test-HR5-IdRefUniqueness {
    param([string]$Content, [ValidationResult]$Result)

    # Find all IdRef values
    $idrefPattern = 'sap2010:WorkflowViewState\.IdRef="([^"]+)"'
    $idrefs = [regex]::Matches($Content, $idrefPattern) | ForEach-Object { $_.Groups[1].Value }

    $idrefPattern2 = 'sap:WorkflowViewState\.IdRef="([^"]+)"'
    $idrefs += [regex]::Matches($Content, $idrefPattern2) | ForEach-Object { $_.Groups[1].Value }

    if ($idrefs.Count -eq 0) {
        $Result.HRCompliance["HR-5"] = "N/A"
        return
    }

    $seen = @{}
    $duplicates = @()
    foreach ($idref in $idrefs) {
        if ($seen.ContainsKey($idref)) {
            $duplicates += $idref
        }
        $seen[$idref] = $true
    }

    if ($duplicates.Count -gt 0) {
        $uniqueDupes = $duplicates | Select-Object -Unique
        $displayDupes = if ($uniqueDupes.Count -gt 5) { ($uniqueDupes[0..4] -join ', ') + "..." } else { $uniqueDupes -join ', ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-5"
        $issue.Severity = "ERROR"
        $issue.Message = "Duplicate IdRefs found: $displayDupes"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-5"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-5"] = "PASS"
    }
}

function Test-HR503-IdRefDualDeclaration {
    param([string]$Content, [ValidationResult]$Result)

    # HR-503: IdRef must be declared ONLY ONCE per activity - either as attribute OR as child element, never both
    # This catches the case where IdRef is set both as attribute AND as child element on the same activity

    $dualDeclarations = @()

    # Find IdRef as attribute
    $attrPattern = 'sap2010:WorkflowViewState\.IdRef="([^"]+)"'
    $attrMatches = [regex]::Matches($Content, $attrPattern)
    $attrValues = @{}
    foreach ($match in $attrMatches) {
        $attrValues[$match.Groups[1].Value] = $true
    }

    # Find IdRef as child element
    $elemPattern = '<sap2010:WorkflowViewState\.IdRef>([^<]+)</sap2010:WorkflowViewState\.IdRef>'
    $elemMatches = [regex]::Matches($Content, $elemPattern)

    foreach ($match in $elemMatches) {
        $value = $match.Groups[1].Value.Trim()
        if ($attrValues.ContainsKey($value)) {
            $dualDeclarations += $value
        }
    }

    # Also check for any element-style IdRef (even without duplicate) - these often cause issues
    if ($elemMatches.Count -gt 0 -and $dualDeclarations.Count -eq 0) {
        # Element syntax exists but no duplicates - still warn as element syntax is discouraged
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-503"
        $issue.Severity = "WARNING"
        $issue.Message = "IdRef uses element syntax (<sap2010:WorkflowViewState.IdRef>). Prefer attribute syntax for consistency."
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-503"] = "WARN"
        return
    }

    if ($dualDeclarations.Count -gt 0) {
        $displayDupes = if ($dualDeclarations.Count -gt 3) { ($dualDeclarations[0..2] -join ', ') + "..." } else { $dualDeclarations -join ', ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-503"
        $issue.Severity = "ERROR"
        $issue.Message = "IdRef declared both as attribute AND child element (causes 'IdRef property already set' error): $displayDupes"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-503"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-503"] = "PASS"
    }
}

function Test-HR6-ExpressionEncoding {
    param([string]$Content, [ValidationResult]$Result)

    # NOTE: Double-encoding detection (e.g., &amp;quot;) is handled by Test-DoubleEncoding
    $issuesFound = @()

    # Check for unescaped special characters
    $valuePattern = '(?:Value|Expression|Condition)="([^"]*)"'
    $matches = [regex]::Matches($Content, $valuePattern)

    foreach ($match in $matches) {
        $value = $match.Groups[1].Value
        # Check for unescaped ampersands
        if ($value -match '&' -and $value -notmatch '&amp;' -and $value -notmatch '&lt;' -and $value -notmatch '&gt;' -and $value -notmatch '&quot;') {
            $truncatedValue = if ($value.Length -gt 50) { $value.Substring(0, 50) + "..." } else { $value }
            $issuesFound += "Potentially unescaped '&' in: $truncatedValue"
        }
        # Check for unescaped < or >
        if ($value -match '<(?!&)' -or $value -match '>(?!&)') {
            $truncatedValue = if ($value.Length -gt 50) { $value.Substring(0, 50) + "..." } else { $value }
            $issuesFound += "Potentially unescaped '<' or '>' in expression"
        }
    }

    if ($issuesFound.Count -gt 0) {
        $displayIssues = if ($issuesFound.Count -gt 3) { ($issuesFound[0..2] -join '; ') } else { $issuesFound -join '; ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-6"
        $issue.Severity = "WARNING"
        $issue.Message = "Potential expression encoding issues: $displayIssues"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-6"] = "WARN"
    }
    else {
        $Result.HRCompliance["HR-6"] = "PASS"
    }
}

function Test-HR7-NoUIAutomation {
    param([string]$Content, [ValidationResult]$Result)

    $uiActivitiesFound = [System.Collections.ArrayList]::new()

    $uiActivityXmlPatterns = @(
        @{ Pattern = '<ui:Click[\s>]'; Name = 'Click activity' },
        @{ Pattern = '<ui:TypeInto[\s>]'; Name = 'TypeInto activity' },
        @{ Pattern = '<ui:GetText[\s>]'; Name = 'GetText activity' },
        @{ Pattern = '<ui:SetText[\s>]'; Name = 'SetText activity' },
        @{ Pattern = '<ui:Hover[\s>]'; Name = 'Hover activity' },
        @{ Pattern = '<ui:SendHotkey[\s>]'; Name = 'SendHotkey activity' },
        @{ Pattern = '<ui:SelectItem[\s>]'; Name = 'SelectItem activity' },
        @{ Pattern = '<ui:FindElement[\s>]'; Name = 'FindElement activity' },
        @{ Pattern = '<ui:ElementExists[\s>]'; Name = 'ElementExists activity' },
        @{ Pattern = '<ui:WaitElement[\s>]'; Name = 'WaitElement activity' },
        @{ Pattern = '<ui:AttachBrowser[\s>]'; Name = 'AttachBrowser activity' },
        @{ Pattern = '<ui:OpenBrowser[\s>]'; Name = 'OpenBrowser activity' },
        @{ Pattern = '<ui:CloseBrowser[\s>]'; Name = 'CloseBrowser activity' },
        @{ Pattern = '<ui:NavigateTo[\s>]'; Name = 'NavigateTo activity' },
        @{ Pattern = '<ui:TakeScreenshot[\s>]'; Name = 'TakeScreenshot activity' },
        @{ Pattern = '<ui:GetAttribute[\s>]'; Name = 'GetAttribute activity' },
        @{ Pattern = '<ui:SetAttribute[\s>]'; Name = 'SetAttribute activity' },
        @{ Pattern = '<Click[\s>]'; Name = 'Click activity' },
        @{ Pattern = '<TypeInto[\s>]'; Name = 'TypeInto activity' },
        @{ Pattern = '<GetText[\s>]'; Name = 'GetText activity' },
        @{ Pattern = '<OpenBrowser[\s>]'; Name = 'OpenBrowser activity' },
        @{ Pattern = '<AttachBrowser[\s>]'; Name = 'AttachBrowser activity' }
    )

    foreach ($item in $uiActivityXmlPatterns) {
        if ($Content -match $item.Pattern) {
            $uiActivitiesFound.Add($item.Name) | Out-Null
        }
    }

    # Check for UI selectors - must be actual HTML/UI selectors, not Excel/MFiles selectors
    # Real UI selectors contain tags like <html, <webctrl, <wnd, <uia, <ctrl, <aa
    if ($Content -match 'Selector\s*=\s*"&lt;(?:html|webctrl|wnd|uia|ctrl|aa)' -or
        $Content -match 'Selector\s*=\s*"<(?:html|webctrl|wnd|uia|ctrl|aa)') {
        $uiActivitiesFound.Add("Selector attribute with UI selector") | Out-Null
    }

    if ($Content -match 'clr-namespace:UiPath\.UIAutomation\.Activities') {
        $uiActivitiesFound.Add("UiPath.UIAutomation.Activities namespace import") | Out-Null
    }

    if ($uiActivitiesFound.Count -gt 0) {
        $uniqueFound = @($uiActivitiesFound | Select-Object -Unique)
        $displayCount = [Math]::Min(5, $uniqueFound.Count)
        $displayList = if ($displayCount -gt 0) { $uniqueFound[0..($displayCount-1)] -join ', ' } else { $uniqueFound -join ', ' }

        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-7"
        $issue.Severity = "ERROR"
        $issue.Message = "UI automation detected (forbidden): $displayList"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-7"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-7"] = "PASS"
    }
}

function Test-HR9-FlowchartViewState {
    param([string]$Content, [ValidationResult]$Result)

    if ($Content -notmatch '<Flowchart') {
        $Result.HRCompliance["HR-9"] = "N/A"
        return
    }

    $hasShapeLocation = $Content -match 'ShapeLocation'
    $hasShapeSize = $Content -match 'ShapeSize'

    $flowSteps = ([regex]::Matches($Content, '<FlowStep')).Count
    $flowDecisions = ([regex]::Matches($Content, '<FlowDecision')).Count
    $totalFlowNodes = $flowSteps + $flowDecisions

    if ($totalFlowNodes -gt 0 -and -not ($hasShapeLocation -and $hasShapeSize)) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-9"
        $issue.Severity = "WARNING"
        $issue.Message = "Flowchart has $totalFlowNodes nodes but may be missing ViewState positioning (ShapeLocation/ShapeSize). Nodes may appear disconnected in designer."
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-9"] = "WARN"
    }
    elseif ($totalFlowNodes -gt 0) {
        $Result.HRCompliance["HR-9"] = "PASS"
    }
    else {
        $Result.HRCompliance["HR-9"] = "PASS"
    }
}

function Test-HR701-InvokeCodeLateBinding {
    param([string]$Content, [ValidationResult]$Result)

    # HR-701: InvokeCode activities must not use late binding on Object-typed variables
    # This detects patterns that cause "Option Strict On disallows late binding" errors

    $lateBindingIssues = [System.Collections.ArrayList]::new()

    # Find all InvokeCode activities with Code attribute
    $invokeCodePattern = '<ui:InvokeCode[^>]*Code="([^"]*)"'
    $invokeCodes = [regex]::Matches($Content, $invokeCodePattern)

    foreach ($ic in $invokeCodes) {
        $codeContent = $ic.Groups[1].Value

        # Decode XML entities
        $decodedCode = $codeContent `
            -replace '&#xD;&#xA;', "`n" `
            -replace '&#xD;', "`r" `
            -replace '&#xA;', "`n" `
            -replace '&quot;', '"' `
            -replace '&lt;', '<' `
            -replace '&gt;', '>' `
            -replace '&amp;', '&' `
            -replace '&#x9;', "`t"

        # Find all variables declared as Object
        $objectVarPattern = 'Dim\s+(\w+)\s+As\s+Object'
        $objectVars = [regex]::Matches($decodedCode, $objectVarPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

        $objectVarNames = @()
        foreach ($match in $objectVars) {
            $objectVarNames += $match.Groups[1].Value
        }

        # Also detect COM objects from CreateObject that might be assigned to untyped variables
        $createObjectPattern = '(\w+)\s*=\s*CreateObject\s*\('
        $comVars = [regex]::Matches($decodedCode, $createObjectPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $comVars) {
            $varName = $match.Groups[1].Value
            if ($objectVarNames -notcontains $varName) {
                $objectVarNames += $varName
            }
        }

        # Also detect Marshal.BindToMoniker assigned variables
        $monikerPattern = '(\w+)\s*=\s*(?:CType\s*\()?\s*(?:System\.Runtime\.InteropServices\.)?Marshal\.BindToMoniker'
        $monikerVars = [regex]::Matches($decodedCode, $monikerPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $monikerVars) {
            $varName = $match.Groups[1].Value
            if ($objectVarNames -notcontains $varName) {
                $objectVarNames += $varName
            }
        }

        if ($objectVarNames.Count -eq 0) {
            continue
        }

        # Check for late binding patterns on these Object variables
        $lateBindingFound = @()

        foreach ($varName in $objectVarNames) {
            # Pattern 1: Direct property/method access: varName.Something
            $directAccessPattern = [regex]::Escape($varName) + '\.\w+(?:\s*\(|\s*=|\s*\.|\s*$|\s*&|\s*\))'
            if ($decodedCode -match $directAccessPattern) {
                # Exclude safe patterns like "varName IsNot Nothing" or "varName Is Nothing"
                # Also exclude CallByName usage (which is the fix)
                $safePattern = 'CallByName\s*\(\s*' + [regex]::Escape($varName)
                if ($decodedCode -notmatch $safePattern) {
                    $lateBindingFound += $varName
                }
            }
        }

        # Pattern 2: With blocks on Object variables
        foreach ($varName in $objectVarNames) {
            $withPattern = 'With\s+' + [regex]::Escape($varName) + '\s'
            if ($decodedCode -match $withPattern) {
                # Check if there's any .Property or .Method inside (which would be late binding)
                $withinWithPattern = 'With\s+' + [regex]::Escape($varName) + '[\s\S]*?End With'
                $withMatch = [regex]::Match($decodedCode, $withinWithPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($withMatch.Success) {
                    $withContent = $withMatch.Value
                    # Inside With block, .Property or .Method() is late binding
                    if ($withContent -match '\.\w+') {
                        if ($lateBindingFound -notcontains $varName) {
                            $lateBindingFound += "$varName (With block)"
                        }
                    }
                }
            }
        }

        if ($lateBindingFound.Count -gt 0) {
            $uniqueVars = $lateBindingFound | Select-Object -Unique
            $displayVars = if ($uniqueVars.Count -gt 3) { ($uniqueVars[0..2] -join ', ') + "..." } else { $uniqueVars -join ', ' }
            $lateBindingIssues.Add("Late binding on Object variables: $displayVars. Use CallByName() or strongly-typed COM interfaces.") | Out-Null
        }
    }

    if ($lateBindingIssues.Count -gt 0) {
        $displayIssues = if ($lateBindingIssues.Count -gt 2) { ($lateBindingIssues[0..1] -join '; ') + "..." } else { $lateBindingIssues -join '; ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "HR-701"
        $issue.Severity = "ERROR"
        $issue.Message = "InvokeCode late binding detected (Option Strict violation): $displayIssues"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["HR-701"] = "FAIL"
    }
    else {
        $Result.HRCompliance["HR-701"] = "PASS"
    }
}

# ============================================================================
# NEW VALIDATION FUNCTIONS
# ============================================================================

function Test-SecretsDetection {
    param([string]$Content, [ValidationResult]$Result)

    $secretsFound = [System.Collections.ArrayList]::new()

    foreach ($item in $Script:SecretsPatterns) {
        if ($Content -match $item.Pattern) {
            $secretsFound.Add($item.Name) | Out-Null
        }
    }

    if ($secretsFound.Count -gt 0) {
        $uniqueFound = @($secretsFound | Select-Object -Unique)
        $displayList = $uniqueFound -join ', '

        $issue = [ValidationIssue]::new()
        $issue.Rule = "Secrets"
        $issue.Severity = "ERROR"
        $issue.Message = "Potential hardcoded secrets detected: $displayList. Use Orchestrator assets/credentials instead."
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["Secrets"] = "FAIL"
    }
    else {
        $Result.HRCompliance["Secrets"] = "PASS"
    }
}

function Test-TypeSystem {
    param([string]$Content, [ValidationResult]$Result)

    $typeIssues = [System.Collections.ArrayList]::new()

    # Extract type arguments
    $typePattern = 'x:TypeArguments="([^"]+)"'
    $types = [regex]::Matches($Content, $typePattern) | ForEach-Object { $_.Groups[1].Value }

    # Extract variable types
    $varTypePattern = '<Variable\s+x:TypeArguments="([^"]+)"'
    $varTypes = [regex]::Matches($Content, $varTypePattern) | ForEach-Object { $_.Groups[1].Value }

    $allTypes = @($types) + @($varTypes) | Select-Object -Unique

    foreach ($type in $allTypes) {
        if (-not $type) { continue }

        # Check for invalid type patterns
        # Spaces in type names (except for generic separators)
        if ($type -match '^\s|\s$') {
            $typeIssues.Add("Type with leading/trailing spaces: '$type'") | Out-Null
        }

        # Generic types should use proper syntax
        if ($type -match '\(' -and $type -notmatch '\(Of ') {
            # VB style should be (Of T), not (T)
            if ($type -match '\([^O]') {
                $typeIssues.Add("Potentially malformed generic type: '$type' - VB uses (Of T) syntax") | Out-Null
            }
        }
    }

    # Check variable names (from <Variable ... Name="...">)
    $varNamePattern = '<Variable\s+[^>]*Name="([^"]+)"'
    $varNames = [regex]::Matches($Content, $varNamePattern) | ForEach-Object { $_.Groups[1].Value }

    # Check argument names (from <x:Property Name="...">)
    $argNamePattern = '<x:Property\s+Name="([^"]+)"'
    $argNames = [regex]::Matches($Content, $argNamePattern) | ForEach-Object { $_.Groups[1].Value }

    $allNames = @($varNames) + @($argNames) | Select-Object -Unique

    foreach ($name in $allNames) {
        if (-not $name) { continue }
        # No spaces allowed in variable/argument names
        if ($name -match '\s') {
            $typeIssues.Add("Variable/argument name contains spaces: '$name'") | Out-Null
        }
        # No leading digits
        if ($name -match '^\d') {
            $typeIssues.Add("Variable/argument name starts with digit: '$name'") | Out-Null
        }
    }

    if ($typeIssues.Count -gt 0) {
        $displayIssues = if ($typeIssues.Count -gt 3) { ($typeIssues[0..2] -join '; ') + "..." } else { $typeIssues -join '; ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "TypeSystem"
        $issue.Severity = "WARNING"
        $issue.Message = "Type system issues: $displayIssues"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["TypeSystem"] = "WARN"
    }
    else {
        $Result.HRCompliance["TypeSystem"] = "PASS"
    }
}

function Test-ExpressionLanguage {
    param([string]$Content, [ValidationResult]$Result, [string]$ExpectedLanguage)

    if (-not $ExpectedLanguage -or $ExpectedLanguage -eq "Unknown") {
        $Result.HRCompliance["ExpressionLang"] = "N/A"
        return
    }

    $vbIndicators = 0
    $csharpIndicators = 0

    # Count VB patterns
    foreach ($pattern in $Script:VBExpressionPatterns) {
        $matches = [regex]::Matches($Content, $pattern)
        $vbIndicators += $matches.Count
    }

    # Count C# patterns
    foreach ($pattern in $Script:CSharpExpressionPatterns) {
        $matches = [regex]::Matches($Content, $pattern)
        $csharpIndicators += $matches.Count
    }

    $detectedLang = if ($vbIndicators -gt $csharpIndicators) { "VB" } elseif ($csharpIndicators -gt $vbIndicators) { "CSharp" } else { "Unknown" }

    if ($detectedLang -ne "Unknown" -and $detectedLang -ne $ExpectedLanguage) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "ExpressionLang"
        $issue.Severity = "WARNING"
        $issue.Message = "Expression language mismatch: project is $ExpectedLanguage but expressions appear to be $detectedLang (VB indicators: $vbIndicators, C# indicators: $csharpIndicators)"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["ExpressionLang"] = "WARN"
    }
    else {
        $Result.HRCompliance["ExpressionLang"] = "PASS"
    }
}

function Test-FlowchartStructure {
    param([string]$Content, [ValidationResult]$Result)

    if ($Content -notmatch '<Flowchart') {
        $Result.HRCompliance["FlowchartStructure"] = "N/A"
        return
    }

    $structureIssues = [System.Collections.ArrayList]::new()

    # Check for StartNode
    if ($Content -notmatch '<Flowchart\.StartNode>') {
        $structureIssues.Add("Flowchart missing StartNode") | Out-Null
    }

    # Check x:Name attributes on flow nodes
    $flowStepPattern = '<FlowStep[^>]*>'
    $flowSteps = [regex]::Matches($Content, $flowStepPattern)
    $stepsWithoutName = 0
    foreach ($step in $flowSteps) {
        if ($step.Value -notmatch 'x:Name=') {
            $stepsWithoutName++
        }
    }
    if ($stepsWithoutName -gt 0) {
        $structureIssues.Add("$stepsWithoutName FlowStep(s) without x:Name (may cause reference issues)") | Out-Null
    }

    # Check for x:Reference usage
    $references = [regex]::Matches($Content, '<x:Reference>([^<]+)</x:Reference>')
    $referenceIds = $references | ForEach-Object { $_.Groups[1].Value }

    # Check that referenced IDs exist
    foreach ($refId in $referenceIds) {
        if ($Content -notmatch "x:Name=`"$refId`"") {
            $structureIssues.Add("x:Reference to non-existent ID: $refId") | Out-Null
        }
    }

    # Check FlowDecision has True and False branches
    $decisions = [regex]::Matches($Content, '<FlowDecision[^>]*>')
    foreach ($decision in $decisions) {
        $decisionStart = $decision.Index
        # Simple heuristic: check if True and False are nearby
        $nearbyContent = $Content.Substring($decisionStart, [Math]::Min(2000, $Content.Length - $decisionStart))
        if ($nearbyContent -notmatch '<FlowDecision\.True>' -and $nearbyContent -notmatch 'True=') {
            $structureIssues.Add("FlowDecision may be missing True branch") | Out-Null
        }
    }

    if ($structureIssues.Count -gt 0) {
        $displayIssues = if ($structureIssues.Count -gt 3) { ($structureIssues[0..2] -join '; ') + "..." } else { $structureIssues -join '; ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "FlowchartStructure"
        $issue.Severity = "WARNING"
        $issue.Message = "Flowchart structure issues: $displayIssues"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["FlowchartStructure"] = "WARN"
    }
    else {
        $Result.HRCompliance["FlowchartStructure"] = "PASS"
    }
}

function Test-ActivitySpecific {
    param([string]$Content, [ValidationResult]$Result)

    $activityIssues = [System.Collections.ArrayList]::new()

    # Helper function to get nearby content safely
    function Get-NearbyContent {
        param([int]$Start, [int]$Length, [string]$FullContent)
        $actualLength = [Math]::Min($Length, $FullContent.Length - $Start)
        if ($actualLength -le 0) { return "" }
        return $FullContent.Substring($Start, $actualLength)
    }

    # ========================================================================
    # CONTROL FLOW ACTIVITIES
    # ========================================================================

    # Check Assign activities have To and Value
    $assignPattern = '<Assign[^>]*DisplayName="([^"]*)"'
    $assigns = [regex]::Matches($Content, $assignPattern)
    foreach ($assign in $assigns) {
        $nearby = Get-NearbyContent -Start $assign.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch '<Assign\.To>' -and $nearby -notmatch 'To=') {
            $activityIssues.Add("Assign '$($assign.Groups[1].Value)' missing To property") | Out-Null
        }
        if ($nearby -notmatch '<Assign\.Value>' -and $nearby -notmatch 'Value=') {
            $activityIssues.Add("Assign '$($assign.Groups[1].Value)' missing Value property") | Out-Null
        }
    }

    # Check If activities have Condition and Then branch
    $ifPattern = '<If[^>]*DisplayName="([^"]*)"'
    $ifs = [regex]::Matches($Content, $ifPattern)
    foreach ($if in $ifs) {
        $nearby = Get-NearbyContent -Start $if.Index -Length 1000 -FullContent $Content
        if ($nearby -notmatch 'Condition=') {
            $activityIssues.Add("If '$($if.Groups[1].Value)' missing Condition") | Out-Null
        }
        if ($nearby -notmatch '<If\.Then>') {
            $activityIssues.Add("If '$($if.Groups[1].Value)' missing Then branch") | Out-Null
        }
    }

    # Check Switch activities have Expression
    $switchPattern = '<Switch[^>]*DisplayName="([^"]*)"'
    $switches = [regex]::Matches($Content, $switchPattern)
    foreach ($switch in $switches) {
        $nearby = Get-NearbyContent -Start $switch.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'Expression=' -and $nearby -notmatch '<Switch\.Expression>') {
            $activityIssues.Add("Switch '$($switch.Groups[1].Value)' missing Expression") | Out-Null
        }
    }

    # Check TryCatch has Try block
    $tryCatchPattern = '<TryCatch[^>]*DisplayName="([^"]*)"'
    $tryCatches = [regex]::Matches($Content, $tryCatchPattern)
    foreach ($tc in $tryCatches) {
        $nearby = Get-NearbyContent -Start $tc.Index -Length 1500 -FullContent $Content
        if ($nearby -notmatch '<TryCatch\.Try>') {
            $activityIssues.Add("TryCatch '$($tc.Groups[1].Value)' missing Try block") | Out-Null
        }
    }

    # Check Throw has Exception
    $throwPattern = '<Throw[^>]*DisplayName="([^"]*)"'
    $throws = [regex]::Matches($Content, $throwPattern)
    foreach ($throw in $throws) {
        $nearby = Get-NearbyContent -Start $throw.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'Exception=') {
            $activityIssues.Add("Throw '$($throw.Groups[1].Value)' missing Exception") | Out-Null
        }
    }

    # ========================================================================
    # LOOP ACTIVITIES
    # ========================================================================

    # Check ForEach has TypeArguments and Body
    $forEachPattern = '<(?:ui:)?ForEach[^>]*DisplayName="([^"]*)"'
    $forEachs = [regex]::Matches($Content, $forEachPattern)
    foreach ($forEach in $forEachs) {
        $nearby = Get-NearbyContent -Start $forEach.Index -Length 800 -FullContent $Content
        if ($nearby -notmatch 'x:TypeArguments=') {
            $activityIssues.Add("ForEach '$($forEach.Groups[1].Value)' missing TypeArguments") | Out-Null
        }
        if ($nearby -notmatch '<(?:ui:)?ForEach\.Body>' -and $nearby -notmatch '<ActivityAction') {
            $activityIssues.Add("ForEach '$($forEach.Groups[1].Value)' missing Body") | Out-Null
        }
    }

    # Check ForEachRow has DataTable
    $forEachRowPattern = '<ui:ForEachRow[^>]*DisplayName="([^"]*)"'
    $forEachRows = [regex]::Matches($Content, $forEachRowPattern)
    foreach ($fer in $forEachRows) {
        $nearby = Get-NearbyContent -Start $fer.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'DataTable=') {
            $activityIssues.Add("ForEachRow '$($fer.Groups[1].Value)' missing DataTable") | Out-Null
        }
    }

    # ========================================================================
    # INVOKE ACTIVITIES
    # ========================================================================

    # Check InvokeWorkflowFile has WorkflowFileName
    $invokePattern = '<ui:InvokeWorkflowFile[^>]*DisplayName="([^"]*)"'
    $invokes = [regex]::Matches($Content, $invokePattern)
    foreach ($invoke in $invokes) {
        $nearby = Get-NearbyContent -Start $invoke.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'WorkflowFileName=') {
            $activityIssues.Add("InvokeWorkflowFile '$($invoke.Groups[1].Value)' missing WorkflowFileName") | Out-Null
        }
    }

    # Check InvokeCode has Code
    $invokeCodePattern = '<ui:InvokeCode[^>]*DisplayName="([^"]*)"'
    $invokeCodes = [regex]::Matches($Content, $invokeCodePattern)
    foreach ($ic in $invokeCodes) {
        $nearby = Get-NearbyContent -Start $ic.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'Code=') {
            $activityIssues.Add("InvokeCode '$($ic.Groups[1].Value)' missing Code property") | Out-Null
        }
    }

    # ========================================================================
    # LOGGING ACTIVITIES
    # ========================================================================

    # Check LogMessage has Level and Message
    $logPattern = '<ui:LogMessage[^>]*DisplayName="([^"]*)"'
    $logs = [regex]::Matches($Content, $logPattern)
    foreach ($log in $logs) {
        $nearby = Get-NearbyContent -Start $log.Index -Length 400 -FullContent $Content
        if ($nearby -notmatch 'Level=') {
            $activityIssues.Add("LogMessage '$($log.Groups[1].Value)' missing Level") | Out-Null
        }
        if ($nearby -notmatch 'Message=') {
            $activityIssues.Add("LogMessage '$($log.Groups[1].Value)' missing Message") | Out-Null
        }
    }

    # HR-601: LogMessage.Message must use attribute syntax, not element syntax with InArgument(String)
    # Element syntax causes type mismatch: InArgument(String) not assignable to InArgument(Object)
    $logMsgElemPattern = '<ui:LogMessage\.Message>\s*<InArgument'
    if ($Content -match $logMsgElemPattern) {
        $activityIssues.Add("LogMessage uses element syntax for Message property (causes InArgument type mismatch). Use attribute syntax: Message=`"[expression]`"") | Out-Null
    }

    # ========================================================================
    # DIALOG ACTIVITIES
    # ========================================================================

    # Check MessageBox has Text
    $msgBoxPattern = '<ui:MessageBox[^>]*DisplayName="([^"]*)"'
    $msgBoxes = [regex]::Matches($Content, $msgBoxPattern)
    foreach ($mb in $msgBoxes) {
        $nearby = Get-NearbyContent -Start $mb.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'Text=') {
            $activityIssues.Add("MessageBox '$($mb.Groups[1].Value)' missing Text") | Out-Null
        }
    }

    # Check InputDialog has Label and Result
    $inputPattern = '<ui:InputDialog[^>]*DisplayName="([^"]*)"'
    $inputs = [regex]::Matches($Content, $inputPattern)
    foreach ($input in $inputs) {
        $nearby = Get-NearbyContent -Start $input.Index -Length 600 -FullContent $Content
        if ($nearby -notmatch 'Label=') {
            $activityIssues.Add("InputDialog '$($input.Groups[1].Value)' missing Label") | Out-Null
        }
    }

    # ========================================================================
    # FILE/FOLDER ACTIVITIES
    # ========================================================================

    # Check CreateDirectory has Path
    $createDirPattern = '<ui:CreateDirectory[^>]*DisplayName="([^"]*)"'
    $createDirs = [regex]::Matches($Content, $createDirPattern)
    foreach ($cd in $createDirs) {
        $nearby = Get-NearbyContent -Start $cd.Index -Length 300 -FullContent $Content
        if ($nearby -notmatch 'Path=') {
            $activityIssues.Add("CreateDirectory '$($cd.Groups[1].Value)' missing Path") | Out-Null
        }
    }

    # Check MoveFile has Path and Destination
    $moveFilePattern = '<ui:MoveFile[^>]*DisplayName="([^"]*)"'
    $moveFiles = [regex]::Matches($Content, $moveFilePattern)
    foreach ($mf in $moveFiles) {
        $nearby = Get-NearbyContent -Start $mf.Index -Length 400 -FullContent $Content
        if ($nearby -notmatch 'Path=') {
            $activityIssues.Add("MoveFile '$($mf.Groups[1].Value)' missing Path") | Out-Null
        }
        if ($nearby -notmatch 'Destination=') {
            $activityIssues.Add("MoveFile '$($mf.Groups[1].Value)' missing Destination") | Out-Null
        }
    }

    # Check DeleteFileX has Path
    $deletePattern = '<ui:DeleteFileX[^>]*DisplayName="([^"]*)"'
    $deletes = [regex]::Matches($Content, $deletePattern)
    foreach ($del in $deletes) {
        $nearby = Get-NearbyContent -Start $del.Index -Length 300 -FullContent $Content
        if ($nearby -notmatch 'Path=') {
            $activityIssues.Add("DeleteFileX '$($del.Groups[1].Value)' missing Path") | Out-Null
        }
    }

    # Check FileExistsX has Path
    $fileExistsPattern = '<ui:FileExistsX[^>]*DisplayName="([^"]*)"'
    $fileExists = [regex]::Matches($Content, $fileExistsPattern)
    foreach ($fe in $fileExists) {
        $nearby = Get-NearbyContent -Start $fe.Index -Length 300 -FullContent $Content
        if ($nearby -notmatch 'Path=') {
            $activityIssues.Add("FileExistsX '$($fe.Groups[1].Value)' missing Path") | Out-Null
        }
    }

    # Check KillProcess has ProcessName
    $killPattern = '<ui:KillProcess[^>]*DisplayName="([^"]*)"'
    $kills = [regex]::Matches($Content, $killPattern)
    foreach ($kill in $kills) {
        $nearby = Get-NearbyContent -Start $kill.Index -Length 300 -FullContent $Content
        if ($nearby -notmatch 'ProcessName=') {
            $activityIssues.Add("KillProcess '$($kill.Groups[1].Value)' missing ProcessName") | Out-Null
        }
    }

    # ========================================================================
    # EXCEL ACTIVITIES
    # ========================================================================

    # Check ReadRange has WorkbookPath and DataTable
    $readRangePattern = '<ui:ReadRange[^>]*DisplayName="([^"]*)"'
    $readRanges = [regex]::Matches($Content, $readRangePattern)
    foreach ($rr in $readRanges) {
        $nearby = Get-NearbyContent -Start $rr.Index -Length 400 -FullContent $Content
        if ($nearby -notmatch 'WorkbookPath=') {
            $activityIssues.Add("ReadRange '$($rr.Groups[1].Value)' missing WorkbookPath") | Out-Null
        }
        if ($nearby -notmatch 'DataTable=') {
            $activityIssues.Add("ReadRange '$($rr.Groups[1].Value)' missing DataTable output") | Out-Null
        }
    }

    # Check WriteRange has WorkbookPath and DataTable
    $writeRangePattern = '<ui:WriteRange[^>]*DisplayName="([^"]*)"'
    $writeRanges = [regex]::Matches($Content, $writeRangePattern)
    foreach ($wr in $writeRanges) {
        $nearby = Get-NearbyContent -Start $wr.Index -Length 400 -FullContent $Content
        if ($nearby -notmatch 'WorkbookPath=') {
            $activityIssues.Add("WriteRange '$($wr.Groups[1].Value)' missing WorkbookPath") | Out-Null
        }
        if ($nearby -notmatch 'DataTable=') {
            $activityIssues.Add("WriteRange '$($wr.Groups[1].Value)' missing DataTable input") | Out-Null
        }
    }

    # Check ExcelApplicationCard has WorkbookPath and Body
    $excelAppPattern = '<ueab:ExcelApplicationCard[^>]*DisplayName="([^"]*)"'
    $excelApps = [regex]::Matches($Content, $excelAppPattern)
    foreach ($ea in $excelApps) {
        $nearby = Get-NearbyContent -Start $ea.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'WorkbookPath=') {
            $activityIssues.Add("ExcelApplicationCard '$($ea.Groups[1].Value)' missing WorkbookPath") | Out-Null
        }
    }

    # Check ExcelProcessScopeX has Body
    $excelScopePattern = '<ueab:ExcelProcessScopeX[^>]*DisplayName="([^"]*)"'
    $excelScopes = [regex]::Matches($Content, $excelScopePattern)
    foreach ($es in $excelScopes) {
        $nearby = Get-NearbyContent -Start $es.Index -Length 800 -FullContent $Content
        if ($nearby -notmatch '<ueab:ExcelProcessScopeX\.Body>') {
            $activityIssues.Add("ExcelProcessScopeX '$($es.Groups[1].Value)' missing Body") | Out-Null
        }
    }

    # Check modern ReadRangeX has Range and SaveTo
    $readRangeXPattern = '<ueab:ReadRangeX[^>]*DisplayName="([^"]*)"'
    $readRangeXs = [regex]::Matches($Content, $readRangeXPattern)
    foreach ($rrx in $readRangeXs) {
        $nearby = Get-NearbyContent -Start $rrx.Index -Length 300 -FullContent $Content
        if ($nearby -notmatch 'SaveTo=') {
            $activityIssues.Add("ReadRangeX '$($rrx.Groups[1].Value)' missing SaveTo") | Out-Null
        }
    }

    # Check AddDataColumn has ColumnName and DataTable
    $addColPattern = '<ui:AddDataColumn[^>]*DisplayName="([^"]*)"'
    $addCols = [regex]::Matches($Content, $addColPattern)
    foreach ($ac in $addCols) {
        $nearby = Get-NearbyContent -Start $ac.Index -Length 400 -FullContent $Content
        if ($nearby -notmatch 'ColumnName=') {
            $activityIssues.Add("AddDataColumn '$($ac.Groups[1].Value)' missing ColumnName") | Out-Null
        }
        if ($nearby -notmatch 'DataTable=') {
            $activityIssues.Add("AddDataColumn '$($ac.Groups[1].Value)' missing DataTable") | Out-Null
        }
    }

    # ========================================================================
    # FLOWCHART ACTIVITIES
    # ========================================================================

    # Check FlowDecision has Condition
    $flowDecPattern = '<FlowDecision[^>]*DisplayName="([^"]*)"'
    $flowDecs = [regex]::Matches($Content, $flowDecPattern)
    foreach ($fd in $flowDecs) {
        $nearby = Get-NearbyContent -Start $fd.Index -Length 500 -FullContent $Content
        if ($nearby -notmatch 'Condition=') {
            $activityIssues.Add("FlowDecision '$($fd.Groups[1].Value)' missing Condition") | Out-Null
        }
    }

    # ========================================================================
    # REPORT ISSUES
    # ========================================================================

    if ($activityIssues.Count -gt 0) {
        $displayIssues = if ($activityIssues.Count -gt 5) { ($activityIssues[0..4] -join '; ') + "..." } else { $activityIssues -join '; ' }
        $issue = [ValidationIssue]::new()
        $issue.Rule = "ActivitySpecific"
        $issue.Severity = "WARNING"
        $issue.Message = "Activity structure issues: $displayIssues"
        $Result.Issues.Add($issue) | Out-Null
        $Result.HRCompliance["ActivitySpecific"] = "WARN"
    }
    else {
        $Result.HRCompliance["ActivitySpecific"] = "PASS"
    }
}

function Test-NamingConventions {
    param([string]$FilePath, [string]$Content, [ValidationResult]$Result)

    $filename = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)

    # Check if filename starts with uppercase (PascalCase)
    if ($filename -and $filename[0] -cne $filename[0].ToString().ToUpper()) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "Standards-Naming"
        $issue.Severity = "INFO"
        $issue.Message = "Workflow filename '$filename' should be PascalCase"
        $Result.Issues.Add($issue) | Out-Null
    }

    # Check for single-letter variable names
    $varPattern = '<Variable\s+[^>]*Name="([^"]+)"'
    $variables = [regex]::Matches($Content, $varPattern) | ForEach-Object { $_.Groups[1].Value }
    $singleLetterVars = $variables | Where-Object { $_.Length -eq 1 }

    if ($singleLetterVars.Count -gt 0) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "Standards-Naming"
        $issue.Severity = "WARNING"
        $issue.Message = "Single-letter variable names found: $($singleLetterVars -join ', ')"
        $Result.Issues.Add($issue) | Out-Null
    }
}

function Test-LoggingPractices {
    param([string]$Content, [ValidationResult]$Result)

    $hasLogging = $Content -match 'LogMessage' -or $Content -match 'Log Message' -or $Content -match 'WriteLine'

    if (-not $hasLogging) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "Standards-Logging"
        $issue.Severity = "INFO"
        $issue.Message = "No logging activities found. Consider adding Log Message for debugging."
        $Result.Issues.Add($issue) | Out-Null
    }
}

# ============================================================================
# MAIN VALIDATION FUNCTION
# ============================================================================

function Invoke-XamlValidation {
    param(
        [string]$FilePath,
        [hashtable]$NamespaceReg,
        [string]$ProjJsonPath
    )

    $result = [ValidationResult]::new()
    $result.FilePath = $FilePath

    # Get project context
    $context = Get-ProjectContext -XamlPath $FilePath -ProjectJsonPath $ProjJsonPath
    $result.ProjectContext = $context

    # Read file
    if (-not (Test-Path $FilePath)) {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "Layer1-XML"
        $issue.Severity = "ERROR"
        $issue.Message = "File not found: $FilePath"
        $result.Issues.Add($issue) | Out-Null
        $result.IsValid = $false
        return $result
    }

    try {
        $content = Get-Content -Path $FilePath -Raw -Encoding UTF8
    }
    catch {
        $issue = [ValidationIssue]::new()
        $issue.Rule = "Layer1-XML"
        $issue.Severity = "ERROR"
        $issue.Message = "Error reading file: $($_.Exception.Message)"
        $result.Issues.Add($issue) | Out-Null
        $result.IsValid = $false
        return $result
    }

    # Layer 1: XML Well-formedness
    if (-not (Test-XmlWellFormed -Content $content -Result $result)) {
        $result.IsValid = $false
        return $result
    }

    # Extract namespaces
    $namespaces = Get-Namespaces -Content $content

    # Layer 2: Hard Rules
    Test-HR0-TemplateStructure -Content $content -Result $result
    Test-HR2-RootInvariants -Content $content -Namespaces $namespaces -Result $result -Registry $NamespaceReg
    Test-NamespacePrefixUsage -Content $content -Namespaces $namespaces -Result $result
    Test-HR3-PrimaryContainer -Content $content -Result $result
    Test-HR4-ArgumentsVariables -Content $content -Result $result
    Test-HR5-IdRefUniqueness -Content $content -Result $result
    Test-HR503-IdRefDualDeclaration -Content $content -Result $result
    Test-HR6-ExpressionEncoding -Content $content -Result $result
    Test-DoubleEncoding -Content $content -Result $result
    Test-HR7-NoUIAutomation -Content $content -Result $result
    Test-HR9-FlowchartViewState -Content $content -Result $result
    Test-HR701-InvokeCodeLateBinding -Content $content -Result $result

    # Layer 3: Extended validations
    Test-SecretsDetection -Content $content -Result $result
    Test-TypeSystem -Content $content -Result $result
    Test-ExpressionLanguage -Content $content -Result $result -ExpectedLanguage $context.language
    Test-FlowchartStructure -Content $content -Result $result
    Test-ActivitySpecific -Content $content -Result $result
    Test-CommonNamespaceRequirements -Content $content -Namespaces $namespaces -Result $result

    # Layer 4: Standards/Conventions
    Test-NamingConventions -FilePath $FilePath -Content $content -Result $result
    Test-LoggingPractices -Content $content -Result $result

    # Determine validity
    $hasErrors = ($result.Issues | Where-Object { $_.Severity -eq "ERROR" }).Count -gt 0
    $result.IsValid = -not $hasErrors

    return $result
}

# ============================================================================
# REPORTING FUNCTIONS
# ============================================================================

function Write-ValidationReport {
    param([ValidationResult]$Result, [switch]$AsJson)

    if ($AsJson) {
        $Result.ToHashtable() | ConvertTo-Json -Depth 10
        return
    }

    $status = if ($Result.IsValid) { "VALID" } else { "INVALID" }
    $statusColor = if ($Result.IsValid) { "Green" } else { "Red" }

    Write-Host ""
    Write-Host ("=" * 70)
    Write-Host "XAML Validation Report: $($Result.FilePath)"
    Write-Host ("=" * 70)
    Write-Host -NoNewline "Status: "
    Write-Host $status -ForegroundColor $statusColor

    # Show project context if available
    if ($Result.ProjectContext.projectPath) {
        Write-Host ""
        Write-Host "Project Context:" -ForegroundColor Cyan
        Write-Host "  Language: $($Result.ProjectContext.language)"
        Write-Host "  Compatibility: $($Result.ProjectContext.compatibility)"
    }
    Write-Host ""

    # HR Compliance Summary
    Write-Host "Compliance Summary:"
    Write-Host ("-" * 50)

    $sortedRules = $Result.HRCompliance.Keys | Sort-Object
    foreach ($rule in $sortedRules) {
        $ruleStatus = $Result.HRCompliance[$rule]
        $icon = switch ($ruleStatus) {
            "PASS" { "[PASS]" }
            "FAIL" { "[FAIL]" }
            "WARN" { "[WARN]" }
            "N/A"  { "[N/A]" }
            default { "[????]" }
        }
        $color = switch ($ruleStatus) {
            "PASS" { "Green" }
            "FAIL" { "Red" }
            "WARN" { "Yellow" }
            "N/A"  { "DarkGray" }
            default { "White" }
        }
        Write-Host -NoNewline "  "
        Write-Host -NoNewline $icon -ForegroundColor $color
        Write-Host " $rule"
    }
    Write-Host ""

    # Issues
    $errors = $Result.Issues | Where-Object { $_.Severity -eq "ERROR" }
    $warnings = $Result.Issues | Where-Object { $_.Severity -eq "WARNING" }
    $infos = $Result.Issues | Where-Object { $_.Severity -eq "INFO" }

    if ($errors.Count -gt 0) {
        Write-Host "Errors ($($errors.Count)):" -ForegroundColor Red
        foreach ($issue in $errors) {
            $lineInfo = if ($issue.Line) { " (line $($issue.Line))" } else { "" }
            Write-Host "  [$($issue.Rule)]$lineInfo`: $($issue.Message)"
        }
        Write-Host ""
    }

    if ($warnings.Count -gt 0) {
        Write-Host "Warnings ($($warnings.Count)):" -ForegroundColor Yellow
        foreach ($issue in $warnings) {
            $lineInfo = if ($issue.Line) { " (line $($issue.Line))" } else { "" }
            Write-Host "  [$($issue.Rule)]$lineInfo`: $($issue.Message)"
        }
        Write-Host ""
    }

    if ($infos.Count -gt 0) {
        Write-Host "Info ($($infos.Count)):" -ForegroundColor Cyan
        foreach ($issue in $infos) {
            Write-Host "  [$($issue.Rule)]: $($issue.Message)"
        }
        Write-Host ""
    }

    Write-Host ("=" * 70)
    Write-Host "Summary: $($errors.Count) errors, $($warnings.Count) warnings, $($infos.Count) info"
    Write-Host ""
}

# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

$resolvedPath = Resolve-Path -Path $Path -ErrorAction SilentlyContinue

if (-not $resolvedPath) {
    Write-Host "Error: Path not found: $Path" -ForegroundColor Red
    exit 1
}

$targetPath = $resolvedPath.Path

# Load namespace registry if provided
$nsRegistry = $null
if ($NamespaceRegistry) {
    $nsRegistry = Get-NamespaceRegistry -RegistryPath $NamespaceRegistry
    if (-not $nsRegistry) {
        Write-Host "Warning: Could not load namespace registry from $NamespaceRegistry" -ForegroundColor Yellow
    }
}

if (Test-Path -Path $targetPath -PathType Leaf) {
    # Single file
    $result = Invoke-XamlValidation -FilePath $targetPath -NamespaceReg $nsRegistry -ProjJsonPath $ProjectJson
    Write-ValidationReport -Result $result -AsJson:$Json

    if ($Strict) {
        $hasIssues = ($result.Issues | Where-Object { $_.Severity -in @("ERROR", "WARNING") }).Count -gt 0
    }
    else {
        $hasIssues = ($result.Issues | Where-Object { $_.Severity -eq "ERROR" }).Count -gt 0
    }

    exit $(if ($hasIssues) { 1 } else { 0 })
}
elseif (Test-Path -Path $targetPath -PathType Container) {
    # Directory
    $xamlFiles = Get-ChildItem -Path $targetPath -Filter "*.xaml" -Recurse

    if ($xamlFiles.Count -eq 0) {
        Write-Host "No .xaml files found in $targetPath"
        exit 0
    }

    Write-Host "Found $($xamlFiles.Count) XAML files to validate..."
    Write-Host ""

    $results = @()
    foreach ($file in $xamlFiles) {
        $result = Invoke-XamlValidation -FilePath $file.FullName -NamespaceReg $nsRegistry -ProjJsonPath $ProjectJson
        Write-ValidationReport -Result $result -AsJson:$Json
        $results += $result
    }

    # Summary
    if (-not $Json) {
        $validCount = ($results | Where-Object { $_.IsValid }).Count
        Write-Host ""
        Write-Host ("=" * 70)
        Write-Host "OVERALL SUMMARY: $validCount/$($results.Count) files valid"
        Write-Host ("=" * 70)
    }

    if ($Strict) {
        $allValid = ($results | Where-Object {
            ($_.Issues | Where-Object { $_.Severity -in @("ERROR", "WARNING") }).Count -gt 0
        }).Count -eq 0
    }
    else {
        $allValid = ($results | Where-Object { -not $_.IsValid }).Count -eq 0
    }

    exit $(if ($allValid) { 0 } else { 1 })
}
else {
    Write-Host "Error: $targetPath is neither a file nor directory" -ForegroundColor Red
    exit 1
}
