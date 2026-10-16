# Builds an updated Citrix golden-image software manifest as a two-sheet .xlsx
# Sheet 1: Software  |  Sheet 2: Python packages

# Falls back to a timestamped name if the target is open in Excel
$out = "C:\Users\deepakbalan\Downloads\software required for citrix - updated.xlsx"
try {
    if (Test-Path -LiteralPath $out) {
        $fs = [System.IO.File]::Open($out, 'Open', 'ReadWrite', 'None')
        $fs.Close()
    }
}
catch {
    $out = "C:\Users\deepakbalan\Downloads\software required for citrix - updated " +
           (Get-Date -Format 'HHmm') + ".xlsx"
    Write-Warning "Target was locked (open in Excel). Writing to a new file instead."
}
$tmp = Join-Path $env:TEMP ("xlsxbuild_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
New-Item -ItemType Directory -Path "$tmp\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$tmp\xl\_rels" -Force | Out-Null
New-Item -ItemType Directory -Path "$tmp\xl\worksheets" -Force | Out-Null

function XmlEsc([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace '&', '&amp;'
    $s = $s -replace '<', '&lt;'
    $s = $s -replace '>', '&gt;'
    return $s
}

function ColLetter([int]$i) {
    $n = $i; $s = ''
    while ($n -ge 0) { $s = [char](65 + ($n % 26)) + $s; $n = [int][math]::Floor($n / 26) - 1 }
    return $s
}

function BuildSheet($rows, $widths) {
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
    [void]$sb.Append('<cols>')
    for ($i = 0; $i -lt $widths.Count; $i++) {
        [void]$sb.Append('<col min="' + ($i + 1) + '" max="' + ($i + 1) + '" width="' + $widths[$i] + '" customWidth="1"/>')
    }
    [void]$sb.Append('</cols><sheetData>')
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $style = if ($r -eq 0) { '1' } else { '2' }
        $ht = if ($r -eq 0) { ' ht="30" customHeight="1"' } else { '' }
        [void]$sb.Append('<row r="' + ($r + 1) + '"' + $ht + '>')
        $cells = $rows[$r]
        for ($c = 0; $c -lt $cells.Count; $c++) {
            $ref = (ColLetter $c) + ($r + 1)
            $val = XmlEsc ([string]$cells[$c])
            [void]$sb.Append('<c r="' + $ref + '" t="inlineStr" s="' + $style + '"><is><t xml:space="preserve">' + $val + '</t></is></c>')
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData><autoFilter ref="A1:' + (ColLetter ($widths.Count - 1)) + $rows.Count + '"/></worksheet>')
    return $sb.ToString()
}

# ---------------------------------------------------------------- sheet 1 ---
$sw = @()
$sw += , @('Category', 'Component', 'Version (pin)', 'Admin to install', 'Standard user can run', 'Standard user can update', 'Licence', 'Notes')
$sw += , @('Editor', 'VS Code (System Installer)', '1.9x.x - pin validated build', 'Yes', 'Yes', 'No', 'No', 'Use System installer, not User. Set "update.mode": "none" or users get failing prompts.')
$sw += , @('Editor', 'VS Code extensions', 'Pin per extension', '-', 'Yes', 'Yes', 'No', 'Python, Jupyter, MSSQL, Snowflake, GitLens, Docker, Copilot. Per-user in %USERPROFILE%\.vscode.')
$sw += , @('Editor', 'GitHub Copilot + Chat', 'Latest', '-', 'Yes', 'Yes', 'Yes - per user', 'Assign seats in GitHub Enterprise / Entra before image sign-off.')
$sw += , @('SCM', 'Git for Windows', '2.47.x', 'Yes', 'Yes', 'No', 'No', 'Set http.sslBackend=schannel so it inherits the Windows cert store.')
$sw += , @('SCM', 'GitHub CLI (gh)', '2.6x.x', 'Machine: Yes / User: No', 'Yes', 'Scope-dependent', 'No', 'Optional.')
$sw += , @('Runtime', 'Python', '3.12.10', 'Yes', 'Yes', 'No - system site-packages blocked, and that is fine', 'No', 'All-users install of the INTERPRETER only. Every library goes in a venv - see the two rows below.')
$sw += , @('Runtime', 'python -m venv / pip / uv', 'venv + pip bundled with Python 3.12.10; uv 0.5.x standalone', 'No - venv and pip ship with Python; uv installs per-user', 'Yes', 'Yes', 'No', 'THE no-admin pattern. python -m venv .venv then pip install -r requirements.txt. Nothing to install and nothing to elevate.')
$sw += , @('Runtime', 'Base packages in SYSTEM Python - OPTIONAL, NOT RECOMMENDED', 'n/a', 'Yes - writes to Program Files', 'Yes', 'No', 'No', 'This is the ONLY row that forces admin for Python libraries, and it is avoidable. A venv does NOT inherit system site-packages by default, so these would be invisible to the recommended workflow anyway. Keep system Python clean and put every library in a venv. Consider removing this row.')
$sw += , @('Runtime', 'Project Python packages (venv)', 'Per requirements lock - see sheet 2', 'No', 'Yes', 'Yes', 'Mixed - check PyMuPDF AGPL', 'ADDED. No admin needed inside a venv. Use --only-binary=:all: to avoid source builds.')
$sw += , @('Runtime', 'MS C++ Build Tools', '14.4x', 'Yes', '-', 'No', 'Free', 'ADDED. Only required if a wheel is missing. Test with --only-binary=:all: first and omit if it passes.')
$sw += , @('Runtime', 'PowerShell', '7.4.x LTS', 'Yes', 'Yes', 'CurrentUser scope only', 'No', 'Alongside Windows PowerShell 5.1, not replacing it.')
$sw += , @('Runtime', 'Node.js', '22.x LTS', 'Yes', 'Yes', 'Yes (npm -g writes to %AppData%)', 'No', 'Only if JS / Functions work is in scope.')
$sw += , @('Runtime', '.NET SDK', '8.0.x LTS', 'Yes', 'Yes', 'No', 'No', 'Only if .NET build work is in scope.')
$sw += , @('Runtime', 'Java (Temurin)', '17 or 21 LTS', 'Yes', 'Yes', 'No', 'No', 'REQUIRED for DBeaver / JDBC to Snowflake.')
$sw += , @('Azure', 'Azure CLI', '2.7x.x', 'Yes', 'Yes', 'az extension add: Yes / az upgrade: No', 'No', 'Add azure-devops extension. Pin version and upgrade via reseal.')
$sw += , @('Azure', 'Az PowerShell module', '12.x / 13.x', 'Yes', 'Yes', 'CurrentUser scope only', 'No', 'Install-Module Az -Scope AllUsers in the image.')
$sw += , @('Azure', 'Storage Explorer', '1.3x.x', 'Per-user: No', 'Yes', 'Yes if per-user', 'No', 'Prefer the per-user installer.')
$sw += , @('Azure', 'AzCopy', '10.2x.x', 'No (portable exe)', 'Yes', 'Yes', 'No', 'Drop on PATH.')
$sw += , @('Azure', 'Bicep CLI', '0.3x.x', 'No (az bicep install)', 'Yes', 'Yes', 'No', 'Only if IaC in scope.')
$sw += , @('Data', 'SSMS', '21.x', 'Yes', 'Yes', 'No', 'No', 'Retain for SQL Server admin only - Agent jobs, XEvents, execution plans.')
$sw += , @('Data', 'DBeaver Community', '25.x', 'Yes', 'Yes', 'No', 'Free (PRO paid)', 'Multi-database GUI including Snowflake. Requires Java.')
$sw += , @('Data', 'Snowflake JDBC driver', '3.2x.x', 'Yes', 'Yes', 'No', 'Free', 'For DBeaver / DataGrip.')
$sw += , @('Data', 'Snowflake ODBC driver', '3.x', 'Yes', 'Yes', 'No', 'Free', 'For ODBC consumers and Excel.')
$sw += , @('Data', 'Snowflake CLI (snow)', '3.x', 'Per-user: No', 'Yes', 'Yes', 'Free', 'SnowSQL is legacy - use Snowflake CLI.')
$sw += , @('Data', 'ODBC Driver for SQL Server', '18.4.x', 'Yes', 'Yes', 'No', 'No', 'Required for pyodbc.')
$sw += , @('Data', 'Azure Data Studio', 'EXCLUDED', '-', '-', '-', '-', 'RETIRED - support ended Feb 2026. Use SSMS 21+ or the VS Code MSSQL extension. Do not bake a retired tool into a new image.')
$sw += , @('Containers', 'Docker Desktop', '4.3x.x', 'Yes + docker-users group + WSL2/Hyper-V + reboot', 'Only if in docker-users', 'No', 'Paid', 'Requires nested virtualisation - confirm hypervisor support before committing.')
$sw += , @('Containers', 'Podman Desktop', '1.1x.x', 'Per-user possible; WSL2 enable needs admin', 'Yes', 'Partial', 'Free', 'Alternative to Docker Desktop, no licence cost.')
$sw += , @('Containers', 'Build on ADO / Jenkins agents', '-', '-', '-', '-', 'Free', 'RECOMMENDED DEFAULT. Removes Docker from the image entirely if nobody authors images locally.')
$sw += , @('Browser', 'Microsoft Edge', 'Current stable', 'Yes', 'Yes', 'Managed', 'No', 'Configure PAC / proxy.')
$sw += , @('Trust', 'Corporate root CA -> LocalMachine store', '-', 'Yes', '-', 'No', 'No', 'NOT SUFFICIENT ALONE. Every runtime carries its own trust store - see rows below.')
$sw += , @('Trust', 'REQUESTS_CA_BUNDLE / SSL_CERT_FILE', '-', 'Yes (machine)', '-', 'User vars: Yes', 'No', 'Python, requests, Azure SDK, Azure CLI. Most common cause of "works on my laptop, fails on VDI".')
$sw += , @('Trust', 'NODE_EXTRA_CA_CERTS', '-', 'Yes', '-', '-', 'No', 'If Node.js is installed.')
$sw += , @('Trust', 'Java cacerts import', '-', 'Yes', '-', '-', 'No', 'REQUIRED for JDBC connections to Snowflake.')
$sw += , @('Trust', 'Snowflake OCSP allow-list', '-', 'Yes', '-', '-', 'No', 'Snowflake drivers do OCSP validation. Corporate SSL inspection breaks this silently. Test through the proxy before sign-off.')
$sw += , @('Trust', 'HTTP_PROXY / HTTPS_PROXY / NO_PROXY', '-', 'Yes', '-', '-', 'No', 'Put internal hosts in NO_PROXY.')
$sw += , @('Trust', 'pip.ini -> internal index + CA', '-', 'Machine: Yes / User: No', '-', 'User: Yes', 'No', 'ADDED. Point pip at Azure Artifacts upstream of PyPI, not raw internet.')
$sw += , @('Trust', 'PIP_INDEX_URL', '-', 'Yes (machine)', '-', 'User vars: Yes', 'No', 'ADDED. Or set in pip.ini. Without this the first pip install fails behind the proxy.')

$w1 = @(14, 34, 46, 30, 22, 30, 22, 78)

# ---------------------------------------------------------------- sheet 2 ---
$py = @()
$py += , @('Group', 'Package', 'Pinned version', 'Admin needed', 'Licence (verify)', 'Status', 'Risk / note')
$py += , @('Agent framework', 'agent-framework', 'Pin at install', 'No', 'MIT', 'ADD', 'Microsoft Agent Framework. Emits OpenTelemetry natively - the monitoring rows below already cover its telemetry.')
$py += , @('Agent framework', 'agent-framework-azure-ai', 'Pin at install', 'No', 'MIT', 'ADD', 'Azure AI integration for MAF.')
$py += , @('LangChain', 'langchain', '1.3.18', 'No', 'MIT', 'VERIFY', 'Confirm this resolves against langchain-core 1.6.1. LangChain pins core to narrow ranges; a mismatch shows as an import error, not a resolver error.')
$py += , @('LangChain', 'langchain-core', '1.6.1', 'No', 'MIT', 'VERIFY', 'See above.')
$py += , @('LangChain', 'langgraph', '1.2.11', 'No', 'MIT', 'DECIDE', 'Alternative orchestrator to Microsoft Agent Framework. Carrying both doubles dependency surface - pick one before the image is cut.')
$py += , @('LangChain', 'langchain-community', 'UNPINNED', 'No', 'MIT', 'PIN THIS', 'No version specified - defeats the point of pinning everything else.')
$py += , @('LangChain', 'langchain-azure-ai', '1.2.9', 'No', 'MIT', 'OK', '')
$py += , @('Azure Foundry', 'azure-ai-projects', '2.5.0', 'No', 'MIT', 'VERIFY', 'Confirm this exact version exists and is GA.')
$py += , @('RAG', 'azure-search-documents', '12.0.0', 'No', 'MIT', 'VERIFY', 'Confirm 12.0.0 is a GA release and not a preview line.')
$py += , @('RAG', 'tiktoken', '0.11.0', 'No', 'MIT', 'OK', 'Rust-backed but ships cp312 Windows wheels - no compiler needed.')
$py += , @('Document processing', 'azure-ai-documentintelligence', '1.0.2', 'No', 'MIT', 'OK', '')
$py += , @('Document processing', 'pymupdf', 'UNPINNED', 'No', 'AGPL-3.0 or commercial', 'LICENCE RISK', 'AGPL is usually a blocker for commercial or client-delivered work. You already have pypdf (BSD) and Document Intelligence - drop PyMuPDF if either covers the requirement.')
$py += , @('Document processing', 'pypdf', 'UNPINNED', 'No', 'BSD-3-Clause', 'PIN THIS', 'Permissive licence - preferred over PyMuPDF where it suffices.')
$py += , @('Storage', 'azure-storage-queue', '12.14.0', 'No', 'MIT', 'OK', '')
$py += , @('Backend', 'fastapi', '0.116.1', 'No', 'MIT', 'OK', '')
$py += , @('Backend', 'uvicorn', '0.35.0', 'No', 'BSD-3-Clause', 'OK', '')
$py += , @('Monitoring', 'opentelemetry-api', '1.37.0', 'No', 'Apache-2.0', 'OK', 'Covers Microsoft Agent Framework telemetry - nothing extra needed.')
$py += , @('Monitoring', 'opentelemetry-sdk', '1.37.0', 'No', 'Apache-2.0', 'OK', '')
$py += , @('Monitoring', 'azure-monitor-opentelemetry', '1.8.1', 'No', 'MIT', 'OK', 'Exports traces and metrics to Application Insights.')
$py += , @('Utilities', 'python-dotenv', '1.1.1', 'No', 'BSD-3-Clause', 'OK', '')
$py += , @('Utilities', 'pydantic', '>=2.11', 'No', 'MIT', 'WATCH', 'Common conflict point between the app stack and the evaluation libraries.')
$py += , @('Evaluation', 'ragas', '0.3.3', 'No', 'Apache-2.0', 'SEPARATE VENV', 'ragas and deepeval pin overlapping pydantic / openai / langchain-core ranges and will fight. Evaluation runs offline - it does not need to share the runtime venv.')
$py += , @('Evaluation', 'deepeval', '3.5.3', 'No', 'Apache-2.0', 'SEPARATE VENV', 'See above.')
$py += , @('Evaluation', 'azure-ai-evaluation', 'Pin at install', 'No', 'MIT', 'ADD', 'Microsoft-native evaluation path. Coexists with azure-ai-projects more cleanly than ragas / deepeval. Consider whether all three evaluators are needed.')
$py += , @('ALL', '(every package above)', '-', 'NO ADMIN', '-', '-', 'All install without admin inside a venv: python -m venv .venv then pip install -r requirements.txt. The ONLY way a Python library needs admin is if you install it into system Python in Program Files - which is optional and not recommended. A missing wheel forcing a source build would need MS C++ Build Tools; test with --only-binary=:all: to rule that out.')

$w2 = @(20, 34, 18, 14, 24, 18, 96)

# ---------------------------------------------------------------- package ---
$ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
'<Default Extension="xml" ContentType="application/xml"/>' +
'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
'</Types>'

$rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
'</Relationships>'

$wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
'<sheets><sheet name="Software" sheetId="1" r:id="rId1"/><sheet name="Python packages" sheetId="2" r:id="rId2"/></sheets></workbook>'

$wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
'</Relationships>'

$styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
'<fonts count="2"><font><sz val="10"/><name val="Segoe UI"/></font>' +
'<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font></fonts>' +
'<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
'<fill><patternFill patternType="solid"><fgColor rgb="FF1A7A1A"/><bgColor indexed="64"/></patternFill></fill></fills>' +
'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
'<cellXfs count="3">' +
'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
'<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>' +
'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

$enc = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText("$tmp\[Content_Types].xml", $ct, $enc)
[System.IO.File]::WriteAllText("$tmp\_rels\.rels", $rels, $enc)
[System.IO.File]::WriteAllText("$tmp\xl\workbook.xml", $wb, $enc)
[System.IO.File]::WriteAllText("$tmp\xl\_rels\workbook.xml.rels", $wbRels, $enc)
[System.IO.File]::WriteAllText("$tmp\xl\styles.xml", $styles, $enc)
[System.IO.File]::WriteAllText("$tmp\xl\worksheets\sheet1.xml", (BuildSheet $sw $w1), $enc)
[System.IO.File]::WriteAllText("$tmp\xl\worksheets\sheet2.xml", (BuildSheet $py $w2), $enc)

if (Test-Path $out) { Remove-Item $out -Force }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $out)
Remove-Item $tmp -Recurse -Force

Write-Output ("Written: " + $out)
Write-Output ("Sheet 1 rows: " + $sw.Count + "   Sheet 2 rows: " + $py.Count)
