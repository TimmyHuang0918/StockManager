# PowerShell Script: Fix indentation - normalize to 8 spaces correctly
# Fixes files that have inconsistent indentation (16, 24, 32 spaces etc.)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Fixing Indentation Issues" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptPath

$files = Get-ChildItem -Path . -Include *.cs,*.xaml -Recurse -File | Where-Object {
    $_.FullName -notlike "*\bin\*" -and 
    $_.FullName -notlike "*\obj\*" -and
    $_.FullName -notlike "*\packages\*"
}

Write-Host "Found $($files.Count) files to check" -ForegroundColor Green
Write-Host ""

$fixedCount = 0
$skippedCount = 0

foreach ($file in $files) {
    Write-Host "Checking: $($file.Name)" -ForegroundColor Yellow
    
    try {
        $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        
        if ($null -eq $content -or $content.Length -eq 0) {
            Write-Host "  ⚠️  Skipped (empty file)" -ForegroundColor Gray
            $skippedCount++
            continue
        }
        
        $lines = $content -split "`r?`n"
        $newLines = @()
        $needsFix = $false
        
        foreach ($line in $lines) {
            if ($line.Length -eq 0) {
                $newLines += $line
                continue
            }
            
            # Count leading spaces
            $leadingSpaces = 0
            for ($i = 0; $i -lt $line.Length; $i++) {
                if ($line[$i] -eq ' ') {
                    $leadingSpaces++
                } elseif ($line[$i] -eq "`t") {
                    $leadingSpaces += 8
                } else {
                    break
                }
            }
            
            $content = $line.TrimStart()
            
            # Calculate correct indentation level
            # Each level should be 8 spaces
            $indentLevel = [Math]::Round($leadingSpaces / 8.0)
            $correctIndent = $indentLevel * 8
            
            # Check if fix is needed
            if ($leadingSpaces -ne $correctIndent -and $content.Length -gt 0) {
                $needsFix = $true
            }
            
            $newLine = (' ' * $correctIndent) + $content
            $newLines += $newLine
        }
        
        if ($needsFix) {
            $newContent = $newLines -join "`r`n"
            [System.IO.File]::WriteAllText($file.FullName, $newContent, [System.Text.Encoding]::UTF8)
            Write-Host "  ✅ Fixed" -ForegroundColor Green
            $fixedCount++
        } else {
            Write-Host "  ✓ Already correct" -ForegroundColor Gray
            $skippedCount++
        }
    }
    catch {
        Write-Host "  ❌ Error: $($_.Exception.Message)" -ForegroundColor Red
        $skippedCount++
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Fix Complete!" -ForegroundColor Green
Write-Host "  Fixed:   $fixedCount files" -ForegroundColor Green
Write-Host "  Skipped: $skippedCount files" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
