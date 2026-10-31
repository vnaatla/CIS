# Builds a locally runnable copy of the CIS VB6 app (MedixCIS) in D:\CIS\LocalRun.
# "CIS VB" stays untouched; all patches go to LocalRun\Src. Safe to rerun (Src is refreshed from "CIS VB").
$ErrorActionPreference = 'Stop'
$root = 'D:\CIS\LocalRun'
$orig = 'D:\CIS\CIS VB'
$src  = Join-Path $root 'Src'
$ansi = [Text.Encoding]::GetEncoding(1252)

foreach ($d in 'Src', 'CIS\Conn', 'Template', 'Signatures', 'InternalDocuments', 'FileLogin', 'Bordx', 'MCO INV') {
    New-Item -ItemType Directory -Force -Path (Join-Path $root $d) | Out-Null
}
robocopy $orig $src /E /NFL /NDL /NJH /NJS /NP /XD 'CIS Package' 'users' | Out-Null
if ($LASTEXITCODE -ge 8) { throw "robocopy failed ($LASTEXITCODE)" }

# --- UDLs (SQLOLEDB, shared memory: TCP is disabled on the local instance) ---
$udl = @{
    'mediconnect'        = 'HISDB'
    'mediconnectA'       = 'HISDMArchive'
    'mediconnectWS'      = 'HISWorkSheet'
    'mediconnectPayment' = 'HISPayment'
    'MediConnHISMaint'   = 'HISMaintenance'
    'MediConnCISDB'      = 'CISDB'
    'MediConnCISWS'      = 'CISWorksheet'
    'MediConnCISPayment' = 'CISPayment'
}
foreach ($k in $udl.Keys) {
    $text = "[oledb]`r`n; Everything after this line is an OLE DB initstring`r`nProvider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=$($udl[$k]);Data Source=(local)`r`n"
    [IO.File]::WriteAllText((Join-Path $root "CIS\Conn\$k.UDL"), $text, [Text.Encoding]::Unicode)
}

# --- SVRSuper.ini (read from App.Path) ---
@"
[path]
SPath="$root\CIS\"
reportPath = "$orig\"
server="$root\"
SecServer = "$root\"
LocreportPath = "$root\Template\"
LocSignPath =  "$root\Signatures\"
SvrIntDoc = "$root\InternalDocuments\"
DocPath = "$root\"
SvrDB = "\\localhost\"
BordxPath = "$root\Bordx\"
MCOPath = "$root\MCO INV\"
"@ -replace "`r?`n", "`r`n" | Set-Content -Path (Join-Path $src 'SVRSuper.ini') -Encoding Ascii -NoNewline
$loginLog = Join-Path $root 'FileLogin\LoginCIS.txt'
if (-not (Test-Path $loginLog)) { New-Item -ItemType File -Path $loginLog | Out-Null }

function Patch-File($path, [scriptblock]$fn) {
    $text = $ansi.GetString([IO.File]::ReadAllBytes($path))
    $new = & $fn $text
    $new = [regex]::Replace($new, '(?<!\r)\n', "`r`n")   # VB6 needs CRLF
    if ($new -ne $text) { [IO.File]::WriteAllBytes($path, $ansi.GetBytes($new)) }
}

# --- Project file: drop Crystal / Outlook refs (not installed), MSCOMCTL 2.2 -> 2.0 (VS6 RTM controls), Office16 paths ---
$vbp = Join-Path $src 'medixcis_local.vbp'
Move-Item -Force (Join-Path $src 'medixcis.vbp') $vbp
Patch-File $vbp {
    param($t)
    $lines = $t -split "`r?`n" | Where-Object { $_ -notmatch 'cpeaut32\.dll|msoutl\.olb|crystl32\.ocx' }
    $t = $lines -join "`r`n"
    $t = $t -replace '(831FDD16-0C5C-11D2-A9FC-0000F8754DA1\}#)2\.[12](#0)', '${1}2.0${2}'
    $t = $t -replace '#[^#\r\n]*EXCEL\.EXE#', '#C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE#'
    $t = $t -replace '#[^#\r\n]*MSWORD\.OLB#', '#C:\Program Files\Microsoft Office\Root\Office16\MSWORD.OLB#'
    $t = $t -replace '#[^#\r\n]*MSO\.DLL#', '#C:\Program Files\Microsoft Office\Root\VFS\ProgramFilesCommonX86\Microsoft Shared\OFFICE16\MSO.DLL#'
    $t -replace 'Path32="[^"]*"', 'Path32="D:\CIS\LocalRun\Src"'
}

# --- Forms: MSCOMCTL 2.2 -> 2.0, strip Crystal Reports control ---
$forms = Get-Content $vbp | Where-Object { $_ -match '^Form=' } | ForEach-Object { Join-Path $src ($_ -replace '^Form=', '') }
foreach ($f in $forms) {
    Patch-File $f {
        param($t)
        $t = $t -replace '(831FDD16-0C5C-11D2-A9FC-0000F8754DA1\}#)2\.[12](#0)', '${1}2.0${2}'
        $t = $t -replace '(?im)^Object = "\{00025600-0000-0000-C000-000000000046\}[^\r\n]*\r?\n', ''
        [regex]::Replace($t, '(?ms)^([ \t]*)Begin Crystal\.CrystalReport [^\r\n]*\r?\n.*?^\1End[ \t]*\r?\n', '')
    }
}
Write-Host "LocalRun ready: $vbp"
