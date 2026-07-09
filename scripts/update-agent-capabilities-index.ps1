# Regenerate agent-capabilities.md from manifest + skill scan + skills-lock merge
param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestPath = Join-Path $scriptDir "agent-capabilities-manifest.json"
$statePath = Join-Path $scriptDir ".agent-capabilities-state.json"

if ((Split-Path $scriptDir -Leaf) -eq "scripts" -and (Split-Path (Split-Path $scriptDir -Parent) -Leaf) -eq ".cursor") {
    $root = Resolve-Path (Join-Path $scriptDir "..\..")
} else {
    $root = Resolve-Path (Join-Path $scriptDir "..")
}

if (-not (Test-Path $manifestPath)) {
    Write-Error "Manifest not found: $manifestPath"
    exit 1
}

$manifestJson = Get-Content $manifestPath -Raw -Encoding UTF8
$manifest = $manifestJson | ConvertFrom-Json
$now = Get-Date -Format "yyyy-MM-dd HH:mm"
$dash = "-"
$pipeRepl = "/"
$labels = $manifest.labels

function Get-OriginLabel {
    param([string]$Key, $Lbl)
    $prop = "origin_$Key"
    if ($Lbl.PSObject.Properties.Name -contains $prop) { return $Lbl.$prop }
    return $Key
}

function Get-ProvenanceLabel {
    param([string]$Key, $Lbl)
    $prop = "provenance_$Key"
    if ($Lbl.PSObject.Properties.Name -contains $prop) { return $Lbl.$prop }
    return $Key
}

function Get-SkillFrontmatter {
    param([string]$SkillMdPath)
    if (-not (Test-Path $SkillMdPath)) { return $null }
    $content = Get-Content $SkillMdPath -Raw -Encoding UTF8
    $name = $null
    $description = $null
    if ($content -match '(?m)^name:\s*(.+)$') { $name = $Matches[1].Trim().Trim('"').Trim("'") }
    if ($content -match '(?m)^description:\s*(.+)$') {
        $description = $Matches[1].Trim().Trim('"').Trim("'")
        if ($description.Length -gt 120) { $description = $description.Substring(0, 117) + "..." }
    }
    return @{ name = $name; description = $description }
}

function Get-LockSources {
    param([string]$WorkspaceRoot, [string[]]$LockPaths)
    $result = @{}
    foreach ($rel in $LockPaths) {
        $lockPath = Join-Path $WorkspaceRoot ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path $lockPath)) { continue }
        $lock = Get-Content $lockPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if (-not $lock.skills) { continue }
        foreach ($prop in $lock.skills.PSObject.Properties) {
            $result[$prop.Name] = $prop.Value
        }
    }
    return $result
}

function Get-ScannedSkills {
    param([string]$WorkspaceRoot, [string[]]$SkillDirs)
    $skills = @{}
    foreach ($rel in $SkillDirs) {
        $dirPath = Join-Path $WorkspaceRoot ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        if (-not (Test-Path $dirPath)) { continue }
        foreach ($sub in Get-ChildItem -Path $dirPath -Directory -ErrorAction SilentlyContinue) {
            $skillMd = Join-Path $sub.FullName "SKILL.md"
            if (-not (Test-Path $skillMd)) { continue }
            $fm = Get-SkillFrontmatter $skillMd
            $key = if ($fm.name) { $fm.name } else { $sub.Name }
            $relPath = $rel + "/" + $sub.Name + "/SKILL.md"
            $skills[$key] = @{
                folder      = $sub.Name
                path        = $relPath
                description = $fm.description
            }
        }
    }
    return $skills
}

function Get-Fingerprint {
    param(
        [hashtable]$Scanned,
        [string]$ManifestPath,
        [string[]]$LockPaths,
        [string]$WorkspaceRoot
    )
    $parts = @()
    foreach ($key in ($Scanned.Keys | Sort-Object)) {
        $parts += "s:$key|$($Scanned[$key].path)"
    }
    $parts += "m:$((Get-Item $ManifestPath).LastWriteTimeUtc.Ticks)"
    foreach ($rel in ($LockPaths | Sort-Object)) {
        $lp = Join-Path $WorkspaceRoot ($rel -replace "/", [IO.Path]::DirectorySeparatorChar)
        if (Test-Path $lp) {
            $parts += "l:$rel|$((Get-Item $lp).LastWriteTimeUtc.Ticks)"
        }
    }
    return ($parts -join ";")
}

function Resolve-Entry {
    param(
        $Key,
        $ScannedItem,
        $ManifestEntry,
        $LockEntry,
        $Lbl
    )
    $summary = $dash
    $origin = "self"
    $provenance = "personal"
    $source = $dash
    $sourceLabel = $dash
    $status = "active"
    $path = if ($ScannedItem) { $ScannedItem.path } else { $dash }

    if ($ManifestEntry) {
        if ($ManifestEntry.summary) { $summary = $ManifestEntry.summary }
        if ($ManifestEntry.origin) { $origin = $ManifestEntry.origin }
        if ($ManifestEntry.provenance) { $provenance = $ManifestEntry.provenance }
        if ($ManifestEntry.source) { $source = $ManifestEntry.source }
        if ($ManifestEntry.source_label) { $sourceLabel = $ManifestEntry.source_label }
        if ($ManifestEntry.status) { $status = $ManifestEntry.status }
    }

    if ($summary -eq $dash -and $ScannedItem -and $ScannedItem.description) {
        $summary = $ScannedItem.description
    }

    if ($LockEntry -and -not ($ManifestEntry -and $ManifestEntry.origin)) {
        $origin = "imported"
        if ($LockEntry.source) {
            $source = $LockEntry.source
            $sourceLabel = $LockEntry.source
        }
    }
    if ($LockEntry -and -not ($ManifestEntry -and $ManifestEntry.source) -and $LockEntry.source) {
        $source = $LockEntry.source
        if ($sourceLabel -eq $dash) { $sourceLabel = $LockEntry.source }
    }

    if (-not ($ManifestEntry -and $ManifestEntry.provenance) -and $LockEntry -and $LockEntry.source) {
        if ($LockEntry.source -match 'microsoft|owasp') { $provenance = "official" }
        else { $provenance = "organization" }
    }

    if ($sourceLabel -eq $dash -and $source -ne $dash) { $sourceLabel = $source }
    if ($origin -eq "self" -and $provenance -eq "personal" -and $sourceLabel -eq $dash) {
        $sourceLabel = $Lbl.source_self_default
    }

    return @{
        key        = $Key
        summary    = ($summary -replace '\|', $pipeRepl)
        origin     = (Get-OriginLabel $origin $Lbl)
        provenance = (Get-ProvenanceLabel $provenance $Lbl)
        source     = ($sourceLabel -replace '\|', $pipeRepl)
        path       = $path
        status     = $status
    }
}

$scanned = Get-ScannedSkills -WorkspaceRoot $root -SkillDirs @($manifest.scan.skill_dirs)
$lockSources = Get-LockSources -WorkspaceRoot $root -LockPaths @($manifest.scan.skills_lock)
$manifestEntries = if ($manifest.entries) { $manifest.entries } else { @{} }
$manifestMcp = if ($manifest.mcp) { $manifest.mcp } else { @{} }
$referenceEntries = if ($manifest.reference_entries) { $manifest.reference_entries } else { @{} }

$fp = Get-Fingerprint -Scanned $scanned -ManifestPath $manifestPath -LockPaths @($manifest.scan.skills_lock) -WorkspaceRoot $root

$prevFp = $null
if (Test-Path $statePath) {
    $prevState = Get-Content $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    $prevFp = $prevState.fingerprint
}

if (-not $Force -and $prevFp -and $prevFp -eq $fp) { exit 0 }

$allKeys = [System.Collections.Generic.HashSet[string]]::new()
foreach ($k in $scanned.Keys) { [void]$allKeys.Add($k) }
foreach ($prop in $manifestEntries.PSObject.Properties) { [void]$allKeys.Add($prop.Name) }

$registered = @()
$unregistered = @()

foreach ($key in ($allKeys | Sort-Object)) {
    $scannedItem = if ($scanned.ContainsKey($key)) { $scanned[$key] } else { $null }
    $manifestEntry = $null
    if ($manifestEntries.PSObject.Properties.Name -contains $key) {
        $manifestEntry = $manifestEntries.$key
    }
    $lockEntry = if ($lockSources.ContainsKey($key)) { $lockSources[$key] } else { $null }

    if ($scannedItem -and -not ($manifestEntries.PSObject.Properties.Name -contains $key)) {
        $unregistered += Resolve-Entry -Key $key -ScannedItem $scannedItem -ManifestEntry $null -LockEntry $lockEntry -Lbl $labels
    } else {
        $registered += Resolve-Entry -Key $key -ScannedItem $scannedItem -ManifestEntry $manifestEntry -LockEntry $lockEntry -Lbl $labels
    }
}

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("# $($manifest.title)")
$lines.Add("")
$lines.Add("> $($manifest.description)")
$manifestRef = if ($manifest.manifest_path) { $manifest.manifest_path } else { "agent-capabilities-manifest.json" }
$lines.Add("> $($labels.manifest_prefix): $manifestRef")
$lines.Add("> $($labels.generated_prefix): $now (update-agent-capabilities-index.ps1)")
$lines.Add("> $($labels.edit_hint)")
$lines.Add("")
$lines.Add("## $($labels.skills_heading)")
$lines.Add("")
$lines.Add("| $($labels.col_name) | $($labels.col_summary) | $($labels.col_origin) | $($labels.col_provenance) | $($labels.col_source) | $($labels.col_path) | $($labels.col_status) |")
$lines.Add("|---|---|---|---|---|---|---|")

foreach ($row in ($registered | Sort-Object { $_.key })) {
    $lines.Add("| ``$($row.key)`` | $($row.summary) | $($row.origin) | $($row.provenance) | $($row.source) | ``$($row.path)`` | $($row.status) |")
}

if ($referenceEntries.PSObject.Properties.Count -gt 0) {
    $lines.Add("")
    $lines.Add("## $($labels.reference_heading)")
    $lines.Add("")
    $lines.Add("| $($labels.col_name) | $($labels.col_summary) | $($labels.col_origin) | $($labels.col_provenance) | $($labels.col_source) | $($labels.col_scope) | $($labels.col_status) |")
    $lines.Add("|---|---|---|---|---|---|---|")
    foreach ($prop in ($referenceEntries.PSObject.Properties | Sort-Object Name)) {
        $e = $prop.Value
        $oJa = Get-OriginLabel $e.origin $labels
        $pJa = Get-ProvenanceLabel $e.provenance $labels
        $sum = ($e.summary -replace '\|', $pipeRepl)
        $src = if ($e.source_label) { $e.source_label } elseif ($e.source) { $e.source } else { $dash }
        $scope = if ($e.scope) { $e.scope } else { $dash }
        $st = if ($e.status) { $e.status } else { "active" }
        $lines.Add("| ``$($prop.Name)`` | $sum | $oJa | $pJa | $src | $scope | $st |")
    }
}

if ($manifestMcp.PSObject.Properties.Count -gt 0) {
    $lines.Add("")
    $lines.Add("## $($labels.mcp_heading)")
    $lines.Add("")
    $lines.Add("| $($labels.col_name) | $($labels.col_summary) | $($labels.col_origin) | $($labels.col_provenance) | $($labels.col_source) | $($labels.col_policy) | $($labels.col_status) |")
    $lines.Add("|---|---|---|---|---|---|---|")
    foreach ($prop in ($manifestMcp.PSObject.Properties | Sort-Object Name)) {
        $e = $prop.Value
        $oJa = Get-OriginLabel $e.origin $labels
        $pJa = Get-ProvenanceLabel $e.provenance $labels
        $sum = ($e.summary -replace '\|', $pipeRepl)
        $src = if ($e.source_label) { $e.source_label } elseif ($e.source) { $e.source } else { $dash }
        $policy = if ($e.readonly_default -eq $true) { $labels.policy_readonly_default } elseif ($e.policy_ref) { $e.policy_ref } else { $dash }
        $st = if ($e.status) { $e.status } else { "active" }
        $lines.Add("| ``$($prop.Name)`` | $sum | $oJa | $pJa | $src | $policy | $st |")
    }
}

if ($unregistered.Count -gt 0) {
    $lines.Add("")
    $lines.Add("## $($labels.unregistered_heading)")
    $lines.Add("")
    $lines.Add("| $($labels.col_name) | $($labels.col_path) | $($labels.col_summary) |")
    $lines.Add("|---|---|---|")
    foreach ($row in ($unregistered | Sort-Object { $_.key })) {
        $lines.Add("| ``$($row.key)`` | ``$($row.path)`` | $($row.summary) |")
    }
}

$lines.Add("")
$lines.Add("## $($labels.rules_heading)")
$lines.Add("")
$lines.Add("- $($labels.rule1)")
$lines.Add("- $($labels.rule2)")
$lines.Add("- $($labels.rule3)")

$outputPath = Join-Path $root ($manifest.output -replace "/", [IO.Path]::DirectorySeparatorChar)
$outputDir = Split-Path $outputPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($outputPath, ($lines -join [Environment]::NewLine), $utf8NoBom)

$newState = @{ fingerprint = $fp; updated = $now } | ConvertTo-Json
[System.IO.File]::WriteAllText($statePath, $newState, $utf8NoBom)

exit 0
