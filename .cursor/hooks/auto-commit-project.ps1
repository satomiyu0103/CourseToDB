# stop フック: プロジェクト repo の未コミット変更を自動 commit（push なし）
# stdin の JSON は読み捨て（フック契約のため消費する）
$null = [Console]::In.ReadToEnd()

$root = Resolve-Path (Join-Path $PSScriptRoot "..\..")
Set-Location $root

if (-not (Test-Path ".git")) { exit 0 }

$porcelain = git status --porcelain 2>$null
if (-not $porcelain) { exit 0 }

git add .cursor/ doc/ src/ tests/ scripts/ config/ gas/ AGENTS.md DESIGN.md 2>$null
$staged = git diff --cached --name-only
if (-not $staged) { exit 0 }

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
$msg = "chore(project): auto-commit $timestamp"
$msgPath = Join-Path $root ".git/COMMIT_EDITMSG_UTF8.txt"
[System.IO.File]::WriteAllText($msgPath, $msg, [System.Text.UTF8Encoding]::new($false))
git commit -F $msgPath 2>$null | Out-Null
exit 0
