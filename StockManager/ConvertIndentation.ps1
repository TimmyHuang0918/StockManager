# PowerShell Script: Convert indentation from 4 spaces to 8 spaces
# For all .cs and .xaml files in the StockManager project

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Converting Indentation: 4 spaces → 8 spaces" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Get the script directory (StockManager folder)
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptPath

# Find all .cs and .xaml files
$files = Get-ChildItem -Path . -Include *.cs,*.xaml -Recurse -File | Where-Object {
    $_.FullName -notlike "*\bin\*" -and 
    $_.FullName -notlike "*\obj\*" -and
    $_.FullName -notlike "*\packages\*"
}

Write-Host "Found $($files.Count) files to process" -ForegroundColor Green
Write-Host ""

$processedCount = 0
$skippedCount = 0

foreach ($file in $files) {
    Write-Host "Processing: $($file.Name)" -ForegroundColor Yellow
    
    try {
        # Read the file content
        $content = Get-Content -Path $file.FullName -Raw -Encoding UTF8
        
        if ($null -eq $content -or $content.Length -eq 0) {
            Write-Host "  ⚠️  Skipped (empty file)" -ForegroundColor Gray
            $skippedCount++
            continue
        }
        
        # Split into lines
        $lines = $content -split "`r?`n"
        $newLines = @()
        
        foreach ($line in $lines) {
            if ($line.Length -eq 0) {
                # Empty line - keep as is
                $newLines += $line
                continue
            }
            
            # Count leading spaces
            $leadingSpaces = 0
            for ($i = 0; $i -lt $line.Length; $i++) {
                if ($line[$i] -eq ' ') {
                    $leadingSpaces++
                } elseif ($line[$i] -eq "`t") {
                    # Tab character - convert to 8 spaces
                    $leadingSpaces += 8
                } else {
                    break
                }
            }
            
            # Get the non-whitespace content
            $content = $line.TrimStart()
            
            # Calculate new indentation (double the spaces, rounded to nearest 8)
            # If original was 4n spaces, new is 8n spaces
            $indentLevel = [Math]::Floor($leadingSpaces / 4)
            $newIndent = $indentLevel * 8
            
            # Create new line with new indentation
            $newLine = (' ' * $newIndent) + $content
            $newLines += $newLine
        }
        
        # Join lines back together
        $newContent = $newLines -join "`r`n"
        
        # Write back to file
        [System.IO.File]::WriteAllText($file.FullName, $newContent, [System.Text.Encoding]::UTF8)
        
        Write-Host "  ✅ Converted" -ForegroundColor Green
        $processedCount++
    }
    catch {
        Write-Host "  ❌ Error: $($_.Exception.Message)" -ForegroundColor Red
        $skippedCount++
    }
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Conversion Complete!" -ForegroundColor Green
Write-Host "  Processed: $processedCount files" -ForegroundColor Green
Write-Host "  Skipped:   $skippedCount files" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "⚠️  IMPORTANT: Please review the changes before committing!" -ForegroundColor Yellow
Write-Host ""
