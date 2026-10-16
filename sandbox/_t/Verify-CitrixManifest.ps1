# Reads back the generated workbook and prints both sheets for verification.
param(
    [string]$Path = "C:\Users\deepakbalan\Downloads\software required for citrix - updated.xlsx"
)

$t = Join-Path $env:TEMP ("verify_" + [guid]::NewGuid().ToString('N'))
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $t)

$wb = [xml][System.IO.File]::ReadAllText((Join-Path $t 'xl\workbook.xml'))
Write-Output ("FILE   : " + $Path)
Write-Output ("SIZE   : " + (Get-Item -LiteralPath $Path).Length + " bytes")
Write-Output ("SHEETS : " + (($wb.workbook.sheets.sheet | ForEach-Object { $_.name }) -join ', '))
Write-Output ""

function Show-Sheet {
    param([string]$File, [string]$Title, [int[]]$Show)
    $s = [xml][System.IO.File]::ReadAllText($File)
    $rows = @($s.worksheet.sheetData.row)
    Write-Output ("=" * 110)
    Write-Output ("  " + $Title + "   (" + $rows.Count + " rows)")
    Write-Output ("=" * 110)
    $n = 0
    foreach ($r in $rows) {
        $n++
        $cells = @($r.c)
        $vals = @()
        foreach ($i in $Show) {
            $v = ''
            # <t> carries xml:space so it is an element, not a string - use InnerText
            if ($i -lt $cells.Count) { $v = [string]$cells[$i].InnerText }
            if ($v.Length -gt 46) { $v = $v.Substring(0, 43) + '...' }
            $vals += $v
        }
        Write-Output (("{0,3}  " -f $n) + ($vals -join ' | '))
    }
    Write-Output ""
}

Show-Sheet -File (Join-Path $t 'xl\worksheets\sheet1.xml') -Title 'SHEET 1 - Software   [Category | Component | Admin to install]' -Show @(0, 1, 3)
Show-Sheet -File (Join-Path $t 'xl\worksheets\sheet2.xml') -Title 'SHEET 2 - Python packages   [Package | Pinned version | Admin needed | Status]' -Show @(1, 2, 3, 5)

Remove-Item $t -Recurse -Force
