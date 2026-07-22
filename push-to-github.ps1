param(
    [string]$Message,
    [string]$Branch,
    [string]$Remote = "origin"
)

$ErrorActionPreference = "Stop"

function Assert-GitAvailable {
    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw "Git isn't installed or isn't on PATH. Install it from https://git-scm.com/download/win and try again."
    }
}

function Assert-InsideGitRepo {
    git rev-parse --is-inside-work-tree *> $null
    if ($LASTEXITCODE -ne 0) {
        throw "This folder isn't a git repository yet. Run 'git init' here first, or cd into your project folder."
    }
}

Assert-GitAvailable
Assert-InsideGitRepo

if (-not $Branch) {
    $Branch = (git rev-parse --abbrev-ref HEAD).Trim()
}

$existingRemotes = git remote
if (-not ($existingRemotes -contains $Remote)) {
    $repoUrl = Read-Host "No '$Remote' remote is set yet. Paste your GitHub repo URL (e.g. https://github.com/you/repo.git)"
    git remote add $Remote $repoUrl
    Write-Host "Added remote '$Remote' -> $repoUrl" -ForegroundColor Cyan
}

Write-Host "Repo:   $(git remote get-url $Remote)" -ForegroundColor DarkGray
Write-Host "Branch: $Branch" -ForegroundColor DarkGray

git add -A

$statusOutput = git status --porcelain
if (-not $statusOutput) {
    Write-Host "Nothing to commit - working tree is clean." -ForegroundColor Yellow
}
else {
    if (-not $Message) {
        $Message = "Update $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    }
    git commit -m $Message
    if ($LASTEXITCODE -ne 0) { throw "git commit failed - see the output above." }
    Write-Host "Committed: $Message" -ForegroundColor Green
}

git push -u $Remote $Branch
if ($LASTEXITCODE -ne 0) {
    throw "git push failed - see the output above."
}

Write-Host "Pushed '$Branch' to '$Remote'." -ForegroundColor Green
