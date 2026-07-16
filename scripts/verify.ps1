$ErrorActionPreference = "Stop"
Set-Location -LiteralPath "$PSScriptRoot\.."

$javascriptFiles = Get-ChildItem -LiteralPath ".\apps\LKC_WorshipPPT" -Filter "*.js" | Where-Object Name -NotLike "vendor-*" | Sort-Object Name
foreach ($javascriptFile in $javascriptFiles) {
    & node --check $javascriptFile.FullName
    if ($LASTEXITCODE -ne 0) {
        throw "Syntax check failed: $($javascriptFile.Name)"
    }
}

$testFiles = Get-ChildItem -LiteralPath ".\apps\LKC_WorshipPPT" -Filter "*.test.js" | Sort-Object Name
foreach ($testFile in $testFiles) {
    Write-Host "Running $($testFile.Name)..."
    & node $testFile.FullName
    if ($LASTEXITCODE -ne 0) {
        throw "Test failed: $($testFile.Name)"
    }
}

$rules = Get-Content -Raw -LiteralPath ".\firebase\database.rules.worship-layout.json" | ConvertFrom-Json
if (-not $rules.rules.worshipPpt.layoutConfig.shared) {
    throw "Worship layout RTDB rule is missing."
}

Write-Host "All Worship PPT generator tests passed."
