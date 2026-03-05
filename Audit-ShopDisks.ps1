#Requires -Version 5.1
<#
.SYNOPSIS
    Audita o espaco em disco dos computadores das lojas Artsana Portugal.
.DESCRIPTION
    Usa PsExec para executar remotamente um script de recolha em cada PC.
    Recolhe info de discos, top ficheiros e top pastas, e gera relatorio HTML.
.PARAMETER TopFiles
    Numero de ficheiros maiores a listar por PC (default: 50)
.PARAMETER TopFolders
    Numero de pastas maiores a listar por PC (default: 20)
.PARAMETER TimeoutSeconds
    Timeout em segundos para cada PC remoto (default: 600)
.PARAMETER ReportPath
    Pasta onde gravar o relatorio HTML (default: .\Reports)
.PARAMETER PsExecPath
    Caminho para PsExec.exe. Se omitido, descarrega automaticamente.
.PARAMETER MaxParallel
    Numero maximo de PCs a auditar em paralelo (default: 10)
#>
param(
    [int]$TopFiles = 50,
    [int]$TopFolders = 20,
    [int]$TimeoutSeconds = 600,
    [string]$ReportPath = "",
    [string]$PsExecPath = "",
    [int]$MaxParallel = 10
)

# --- Configuracao ---
$servidores = @(
    "PT4004W01", "PT4004P02", "PT4004P01", "PT4006P01", "PT4006P02", "PT4006W01",
    "PT4010P01", "PT4010P02", "PT4010W01", "PT4012P01", "PT4012P02", "PT4012W01",
    "PT4015P01", "PT4015P02", "PT4015W01", "PT4018P01", "PT4018P02", "PT4018W01",
    "PT4023P01", "PT4023P02", "PT4023W01", "PT4025P01", "PT4025P02", "PT4025W01",
    "PT4026P01", "PT4026P02", "PT4026W01", "PT4029P01", "PT4029P02", "PT4029W01",
    "PT4030P01", "PT4030P02", "PT4030W01", "PT4031P01", "PT4031P02", "PT4031W01",
    "PT4032P01", "PT4032P02", "PT4032W01", "PT4033P01", "PT4033P02", "PT4033W01",
    "PT4034P01", "PT4034P02", "PT4034W01", "PT4035P01", "PT4035P02", "PT4035W01",
    "PT4036P01", "PT4036P02", "PT4036W01", "PT4037P01", "PT4037P02", "PT4037W01",
    "PT4043P01", "PT4043P02", "PT4043W01", "PT4049P01", "PT4049P02", "PT4049W01",
    "PT4094P01", "PT4094P02", "PT4094W01", "PT4095P01", "PT4095P02", "PT4095W01",
    "PT4097P01", "PT4097P02", "PT4097P03", "PT4097P04", "PT4097P05", "PT4097W01"
)

$ScriptStartTime = Get-Date
if (-not $ReportPath) {
    $ReportPath = Join-Path $PSScriptRoot "Reports"
}
$TempFolder = "C:\TEMP\AuditDisk"
$CollectorScript = "AuditCollect.ps1"
$ResultFile = "audit_result.xml"

# --- Funcoes ---

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","OK","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $colors = @{ INFO = "Cyan"; OK = "Green"; WARN = "Yellow"; ERROR = "Red" }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $colors[$Level]
}

function Install-PsExec {
    param([string]$DestPath)

    $psExecExe = Join-Path $DestPath "PsExec.exe"
    if (Test-Path $psExecExe) {
        Write-Log "PsExec encontrado: $psExecExe" "OK"
        return $psExecExe
    }

    Write-Log "A descarregar PsExec do Sysinternals..." "INFO"
    $zipUrl = "https://download.sysinternals.com/files/PSTools.zip"
    $zipPath = Join-Path $DestPath "PSTools.zip"

    if (-not (Test-Path $DestPath)) {
        New-Item -ItemType Directory -Path $DestPath -Force | Out-Null
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        Expand-Archive -Path $zipPath -DestinationPath $DestPath -Force
        Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
        if (Test-Path $psExecExe) {
            Write-Log "PsExec instalado: $psExecExe" "OK"
            return $psExecExe
        } else {
            Write-Log "PsExec.exe nao encontrado apos extraccao" "ERROR"
            return $null
        }
    }
    catch {
        Write-Log "Erro ao descarregar PsExec: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Test-ServerConnectivity {
    param([string[]]$Servers)
    $online = @()
    $offline = @()

    foreach ($srv in $Servers) {
        Write-Log "A testar: $srv ..." "INFO"
        # Testar ping primeiro
        $ping = Test-Connection -ComputerName $srv -Count 1 -Quiet -ErrorAction SilentlyContinue
        if (-not $ping) {
            Write-Log "$srv - SEM RESPOSTA (ping)" "WARN"
            $offline += @{ Name = $srv; Error = "Sem resposta ao ping" }
            continue
        }
        # Testar admin share
        $adminShare = "\\$srv\C`$"
        if (Test-Path $adminShare -ErrorAction SilentlyContinue) {
            Write-Log "$srv - ONLINE (ping + admin share)" "OK"
            $online += $srv
        } else {
            Write-Log "$srv - PING OK mas admin share inacessivel" "WARN"
            $offline += @{ Name = $srv; Error = "Admin share (C$) inacessivel" }
        }
    }

    return @{ Online = $online; Offline = $offline }
}

function New-CollectorScript {
    param([string]$OutputPath)

    $script = @'
param([int]$TopFiles = 50, [int]$TopFolders = 20, [string]$OutFile = "C:\TEMP\AuditDisk\audit_result.xml")

$result = @{
    ComputerName   = $env:COMPUTERNAME
    Disks          = @()
    TopFiles       = @()
    TopFolders     = @()
    CleanupTargets = @()
    Errors         = @()
}

try {
    $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
    foreach ($disk in $disks) {
        $totalGB = [math]::Round($disk.Size / 1GB, 2)
        $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 2)
        $usedGB  = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 2)
        $pctUsed = if ($disk.Size -gt 0) { [math]::Round(($disk.Size - $disk.FreeSpace) / $disk.Size * 100, 1) } else { 0 }
        $result.Disks += @{
            Drive = $disk.DeviceID; Label = $disk.VolumeName
            TotalGB = $totalGB; FreeGB = $freeGB; UsedGB = $usedGB; PctUsed = $pctUsed
        }
    }

    $allFiles = @()
    foreach ($disk in $disks) {
        $dl = $disk.DeviceID + "\"
        try {
            $files = Get-ChildItem -Path $dl -Recurse -File -ErrorAction SilentlyContinue |
                Sort-Object Length -Descending | Select-Object -First ($TopFiles * 2)
            $allFiles += $files
        } catch {
            $result.Errors += "Erro ficheiros em ${dl}: $($_.Exception.Message)"
        }
    }
    $result.TopFiles = @($allFiles | Sort-Object Length -Descending | Select-Object -First $TopFiles |
        ForEach-Object {
            @{ Path = $_.FullName; SizeMB = [math]::Round($_.Length / 1MB, 2)
               LastModified = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm"); Extension = $_.Extension }
        })

    $allFolders = @()
    foreach ($disk in $disks) {
        $dl = $disk.DeviceID + "\"
        try {
            $topDirs = Get-ChildItem -Path $dl -Directory -ErrorAction SilentlyContinue
            foreach ($dir in $topDirs) {
                try {
                    $items = Get-ChildItem -Path $dir.FullName -Recurse -File -ErrorAction SilentlyContinue
                    $sz = ($items | Measure-Object -Property Length -Sum).Sum
                    $fc = ($items | Measure-Object).Count
                    if ($null -eq $sz) { $sz = 0 }
                    $allFolders += @{ Path = $dir.FullName; SizeMB = [math]::Round($sz / 1MB, 2); FileCount = $fc }
                } catch {}
            }
            foreach ($dir in $topDirs) {
                try {
                    $subDirs = Get-ChildItem -Path $dir.FullName -Directory -ErrorAction SilentlyContinue
                    foreach ($sub in $subDirs) {
                        try {
                            $items = Get-ChildItem -Path $sub.FullName -Recurse -File -ErrorAction SilentlyContinue
                            $sz = ($items | Measure-Object -Property Length -Sum).Sum
                            $fc = ($items | Measure-Object).Count
                            if ($null -eq $sz) { $sz = 0 }
                            $allFolders += @{ Path = $sub.FullName; SizeMB = [math]::Round($sz / 1MB, 2); FileCount = $fc }
                        } catch {}
                    }
                } catch {}
            }
        } catch {
            $result.Errors += "Erro pastas em ${dl}: $($_.Exception.Message)"
        }
    }
    $result.TopFolders = @($allFolders | Sort-Object { $_.SizeMB } -Descending | Select-Object -First $TopFolders)

    # --- Cleanup Targets Detection ---
    function Measure-CleanupPath {
        param([string]$Path)
        $sz = 0; $fc = 0
        if (Test-Path $Path) {
            try {
                $items = Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue
                $sz = ($items | Measure-Object -Property Length -Sum).Sum
                $fc = ($items | Measure-Object).Count
                if ($null -eq $sz) { $sz = 0 }
            } catch {}
        }
        return @{ SizeMB = [math]::Round($sz / 1MB, 2); FileCount = $fc }
    }

    # 1) WindowsTemp
    if (Test-Path "C:\Windows\Temp") {
        $m = Measure-CleanupPath "C:\Windows\Temp"
        if ($m.FileCount -gt 0) {
            $result.CleanupTargets += @{ Id = "WindowsTemp"; Name = "Windows Temp"; SizeMB = $m.SizeMB; FileCount = $m.FileCount; Exists = $true }
        }
    }

    # 2) UserProfiles - ptpos0* Temp + Downloads>5d + Recycle Bin
    $upSz = 0; $upFc = 0
    $ptUsers = @(Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "ptpos0*" })
    foreach ($u in $ptUsers) {
        $tempPath = Join-Path $u.FullName "AppData\Local\Temp"
        if (Test-Path $tempPath) { $t = Measure-CleanupPath $tempPath; $upSz += $t.SizeMB; $upFc += $t.FileCount }
        $dlPath = Join-Path $u.FullName "Downloads"
        if (Test-Path $dlPath) {
            try {
                $oldDl = Get-ChildItem -Path $dlPath -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-5) }
                $dlSz = ($oldDl | Measure-Object -Property Length -Sum).Sum
                $dlFc = ($oldDl | Measure-Object).Count
                if ($null -eq $dlSz) { $dlSz = 0 }
                $upSz += [math]::Round($dlSz / 1MB, 2); $upFc += $dlFc
            } catch {}
        }
    }
    # Recycle Bin (once, outside user loop)
    $rbRoot = "C:\`$Recycle.Bin"
    if (Test-Path $rbRoot) {
        try {
            $rbItems = Get-ChildItem -Path $rbRoot -Recurse -File -Force -ErrorAction SilentlyContinue
            $rbSz = ($rbItems | Measure-Object -Property Length -Sum).Sum
            $rbFc = ($rbItems | Measure-Object).Count
            if ($null -eq $rbSz) { $rbSz = 0 }
            $upSz += [math]::Round($rbSz / 1MB, 2); $upFc += $rbFc
        } catch {}
    }
    if ($ptUsers.Count -gt 0 -and $upFc -gt 0) {
        $result.CleanupTargets += @{ Id = "UserProfiles"; Name = "User Profiles (ptpos0*)"; SizeMB = [math]::Round($upSz, 2); FileCount = $upFc; Exists = $true }
    }

    # 3) Drivers
    $driversSubs = @("MICROSOFT", "OTHER_SOFT", "SCRIPTS", "SISQUAL", "SOPHOS", "XSTORE_INSTALL")
    $drSz = 0; $drFc = 0
    foreach ($sub in $driversSubs) {
        $p = "C:\Drivers\$sub"
        if (Test-Path $p) { $t = Measure-CleanupPath $p; $drSz += $t.SizeMB; $drFc += $t.FileCount }
    }
    if ($drFc -gt 0) {
        $result.CleanupTargets += @{ Id = "Drivers"; Name = "C:\Drivers (install leftovers)"; SizeMB = [math]::Round($drSz, 2); FileCount = $drFc; Exists = $true }
    }

    # 4) SoftwareInstall
    $swSubs = @("SQL", "SCRIPTS", "OFFICE2013", "OFFICE365", "Xstore20")
    $swSz = 0; $swFc = 0
    foreach ($sub in $swSubs) {
        $p = "C:\Software_InstallationPT\$sub"
        if (Test-Path $p) { $t = Measure-CleanupPath $p; $swSz += $t.SizeMB; $swFc += $t.FileCount }
    }
    if ($swFc -gt 0) {
        $result.CleanupTargets += @{ Id = "SoftwareInstall"; Name = "C:\Software_InstallationPT"; SizeMB = [math]::Round($swSz, 2); FileCount = $swFc; Exists = $true }
    }

    # 5) BACKUPxstore
    if (Test-Path "C:\BACKUPxstore") {
        $m = Measure-CleanupPath "C:\BACKUPxstore"
        if ($m.FileCount -gt 0) {
            $result.CleanupTargets += @{ Id = "BACKUPxstore"; Name = "C:\BACKUPxstore"; SizeMB = $m.SizeMB; FileCount = $m.FileCount; Exists = $true }
        }
    }

    # 6) RetentionFolders - C:\scanner + C:\tmp\hh_upload >30 days
    $rfSz = 0; $rfFc = 0
    foreach ($rfPath in @("C:\scanner", "C:\tmp\hh_upload")) {
        if (Test-Path $rfPath) {
            try {
                $oldItems = Get-ChildItem -Path $rfPath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) }
                $oSz = ($oldItems | Measure-Object -Property Length -Sum).Sum
                $oFc = ($oldItems | Measure-Object).Count
                if ($null -eq $oSz) { $oSz = 0 }
                $rfSz += [math]::Round($oSz / 1MB, 2); $rfFc += $oFc
            } catch {}
        }
    }
    if ($rfFc -gt 0) {
        $result.CleanupTargets += @{ Id = "RetentionFolders"; Name = "Retention (scanner/hh_upload >30d)"; SizeMB = [math]::Round($rfSz, 2); FileCount = $rfFc; Exists = $true }
    }

    # 7) WindowsOld
    if (Test-Path "C:\Windows.old") {
        $m = Measure-CleanupPath "C:\Windows.old"
        $result.CleanupTargets += @{ Id = "WindowsOld"; Name = "C:\Windows.old"; SizeMB = $m.SizeMB; FileCount = $m.FileCount; Exists = $true }
    }
} catch {
    $result.Errors += "Erro geral: $($_.Exception.Message)"
}

$result | Export-Clixml -Path $OutFile -Force
'@

    $script | Out-File -FilePath $OutputPath -Encoding ASCII -Force
}

function Invoke-RemoteAudit {
    param(
        [string]$Server,
        [string]$PsExecExe,
        [int]$TopFiles,
        [int]$TopFolders,
        [int]$TimeoutSec
    )

    $remoteTempDir = "\\$Server\C`$\TEMP\AuditDisk"
    $remoteScript  = "\\$Server\C`$\TEMP\AuditDisk\$CollectorScript"
    $remoteResult  = "\\$Server\C`$\TEMP\AuditDisk\$ResultFile"
    $localCollector = Join-Path $TempFolder $CollectorScript

    try {
        # Criar pasta remota e copiar script
        if (-not (Test-Path $remoteTempDir)) {
            New-Item -ItemType Directory -Path $remoteTempDir -Force | Out-Null
        }
        Copy-Item -Path $localCollector -Destination $remoteScript -Force

        # Executar via PsExec
        $psArgs = "\\$Server -accepteula -nobanner -h powershell.exe -ExecutionPolicy Bypass -File C:\TEMP\AuditDisk\$CollectorScript -TopFiles $TopFiles -TopFolders $TopFolders -OutFile C:\TEMP\AuditDisk\$ResultFile"
        $proc = Start-Process -FilePath $PsExecExe -ArgumentList $psArgs -NoNewWindow -Wait -PassThru -RedirectStandardError "$env:TEMP\psexec_err_$Server.txt"

        # Esperar e ler resultado
        if ($proc.ExitCode -eq 0 -and (Test-Path $remoteResult)) {
            $data = Import-Clixml -Path $remoteResult
            return $data
        } else {
            $errContent = ""
            $errFile = "$env:TEMP\psexec_err_$Server.txt"
            if (Test-Path $errFile) {
                $errContent = Get-Content $errFile -Raw -ErrorAction SilentlyContinue
                Remove-Item $errFile -Force -ErrorAction SilentlyContinue
            }
            return @{ ComputerName = $Server; Error = "PsExec exit code: $($proc.ExitCode). $errContent" }
        }
    }
    catch {
        return @{ ComputerName = $Server; Error = $_.Exception.Message }
    }
    finally {
        # Limpar ficheiros remotos
        Remove-Item $remoteResult -Force -ErrorAction SilentlyContinue
        Remove-Item $remoteScript -Force -ErrorAction SilentlyContinue
        Remove-Item $remoteTempDir -Force -Recurse -ErrorAction SilentlyContinue
        $errFile = "$env:TEMP\psexec_err_$Server.txt"
        Remove-Item $errFile -Force -ErrorAction SilentlyContinue
    }
}

function New-HtmlReport {
    param(
        [hashtable[]]$Results,
        [hashtable[]]$OfflineServers,
        [datetime]$StartTime
    )

    $duration = (Get-Date) - $StartTime
    $reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Agrupar por loja (prefixo PT4XXX)
    $grouped = @{}
    foreach ($r in $Results) {
        $shopId = $r.ComputerName -replace '(PT\d{4}).*', '$1'
        if (-not $grouped.ContainsKey($shopId)) {
            $grouped[$shopId] = @()
        }
        $grouped[$shopId] += $r
    }

    # Calcular totais por PC e ordenar por menos espaco livre
    $pcStats = @()
    foreach ($r in $Results) {
        $shopId = $r.ComputerName -replace '(PT\d{4}).*', '$1'
        $totalGB = 0; $freeGB = 0; $usedGB = 0
        foreach ($disk in $r.Disks) {
            $totalGB += $disk.TotalGB
            $freeGB  += $disk.FreeGB
            $usedGB  += $disk.UsedGB
        }
        $pctUsed = if ($totalGB -gt 0) { [math]::Round($usedGB / $totalGB * 100, 1) } else { 0 }
        $pcStats += @{
            ComputerName = $r.ComputerName
            ShopId       = $shopId
            TotalGB      = [math]::Round($totalGB, 1)
            FreeGB       = [math]::Round($freeGB, 1)
            UsedGB       = [math]::Round($usedGB, 1)
            PctUsed      = $pctUsed
        }
    }
    # Ordenar: menos espaco livre primeiro (mais critico no topo)
    $pcStats = $pcStats | Sort-Object { $_.FreeGB }

    # Ordenar lojas pela media de espaco livre dos seus PCs
    $shopFreeAvg = @{}
    foreach ($shopId in $grouped.Keys) {
        $shopPCStats = @($pcStats | Where-Object { $_.ShopId -eq $shopId })
        $avgFree = ($shopPCStats | ForEach-Object { $_.FreeGB } | Measure-Object -Average).Average
        $shopFreeAvg[$shopId] = $avgFree
    }
    $orderedShopIds = $shopFreeAvg.GetEnumerator() | Sort-Object Value | ForEach-Object { $_.Key }

    $css = @"
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Segoe UI', Tahoma, sans-serif; background: #f5f5f5; padding: 20px; color: #333; }
    h1 { color: #2c3e50; margin-bottom: 5px; }
    h2 { color: #34495e; margin: 30px 0 15px; border-bottom: 2px solid #3498db; padding-bottom: 5px; }
    h3 { color: #2980b9; margin: 20px 0 10px; }
    .header { background: #2c3e50; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
    .header h1 { color: white; }
    .header p { color: #bdc3c7; margin-top: 5px; }
    .summary { display: flex; gap: 15px; margin: 15px 0; flex-wrap: wrap; }
    .summary-card { background: white; padding: 15px 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); min-width: 150px; }
    .summary-card .number { font-size: 24px; font-weight: bold; color: #2c3e50; }
    .summary-card .label { font-size: 12px; color: #7f8c8d; text-transform: uppercase; }
    .shop-section { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
    table { width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 13px; }
    th { background: #34495e; color: white; padding: 8px 10px; text-align: left; }
    td { padding: 6px 10px; border-bottom: 1px solid #ecf0f1; }
    tr:nth-child(even) { background: #f9f9f9; }
    tr:hover { background: #eaf2f8; }
    .disk-green { background-color: #d5f5e3; }
    .disk-yellow { background-color: #fdebd0; }
    .disk-red { background-color: #fadbd8; }
    .error-section { background: #fadbd8; padding: 15px; border-radius: 8px; margin-bottom: 20px; border-left: 4px solid #e74c3c; }
    .error-section h2 { color: #c0392b; border-bottom-color: #e74c3c; }
    .size-large { color: #e74c3c; font-weight: bold; }
    .size-medium { color: #e67e22; }
    .size-small { color: #27ae60; }
    .collapsible { cursor: pointer; user-select: none; }
    .collapsible::before { content: '[-] '; font-family: monospace; }
    .collapsible.collapsed::before { content: '[+] '; }
    .collapsible-content { overflow: hidden; }
    .collapsible-content.hidden { display: none; }
    .shop-overview { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; }
    .shop-overview h2 { margin-top: 0; }
    .bar-container { background: #ecf0f1; border-radius: 4px; height: 18px; position: relative; }
    .bar-fill { height: 100%; border-radius: 4px; transition: width 0.3s; }
    .bar-green { background: #27ae60; }
    .bar-yellow { background: #f39c12; }
    .bar-red { background: #e74c3c; }
    .cleanup-section { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); margin-bottom: 20px; border-left: 4px solid #8e44ad; }
    .cleanup-section h2 { color: #8e44ad; border-bottom-color: #8e44ad; margin-top: 0; }
    .cleanup-controls { display: flex; gap: 15px; align-items: center; margin: 10px 0; flex-wrap: wrap; }
    .cleanup-btn { background: #8e44ad; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-size: 14px; }
    .cleanup-btn:hover { background: #7d3c98; }
    .cleanup-btn-copy { background: #27ae60; }
    .cleanup-btn-copy:hover { background: #219a52; }
    .cleanup-summary { background: #f4ecf7; padding: 12px 18px; border-radius: 6px; margin: 10px 0; font-size: 14px; }
    .cleanup-summary strong { color: #8e44ad; }
    .command-box { background: #2c3e50; color: #2ecc71; padding: 15px; border-radius: 6px; font-family: 'Consolas', 'Courier New', monospace; font-size: 13px; margin: 10px 0; white-space: pre-wrap; word-break: break-all; display: none; }
    .cleanup-cb-cell { text-align: center; }
    .cleanup-cb { width: 18px; height: 18px; cursor: pointer; }
    .pc-header-row { background: #eaf2f8 !important; font-weight: bold; }
    .pc-header-row td { border-top: 2px solid #8e44ad; }
</style>
"@

    $js = @"
<script>
function toggleSection(id) {
    var el = document.getElementById(id);
    var header = el.previousElementSibling;
    if (el.classList.contains('hidden')) {
        el.classList.remove('hidden');
        header.classList.remove('collapsed');
    } else {
        el.classList.add('hidden');
        header.classList.add('collapsed');
    }
}
function toggleAllTargets(pcName, checked) {
    var cbs = document.querySelectorAll('.cleanup-cb[data-pc="' + pcName + '"]');
    for (var i = 0; i < cbs.length; i++) { cbs[i].checked = checked; }
    updateSummary();
}
function toggleAllPCs(checked) {
    var cbs = document.querySelectorAll('.cleanup-cb');
    for (var i = 0; i < cbs.length; i++) { cbs[i].checked = checked; }
    var pcCbs = document.querySelectorAll('.cleanup-pc-all');
    for (var i = 0; i < pcCbs.length; i++) { pcCbs[i].checked = checked; }
    updateSummary();
}
function updateSummary() {
    var cbs = document.querySelectorAll('.cleanup-cb:checked');
    var pcs = {}; var targets = 0; var totalMB = 0;
    for (var i = 0; i < cbs.length; i++) {
        pcs[cbs[i].getAttribute('data-pc')] = true;
        targets++;
        totalMB += parseFloat(cbs[i].getAttribute('data-size') || 0);
    }
    var pcCount = Object.keys(pcs).length;
    var gb = (totalMB / 1024).toFixed(1);
    var el = document.getElementById('cleanupSummary');
    if (el) {
        el.innerHTML = '<strong>' + pcCount + '</strong> PCs | <strong>' + targets + '</strong> targets | <strong>' + gb + ' GB</strong> estimados para limpeza';
    }
}
function generateCommand() {
    var cbs = document.querySelectorAll('.cleanup-cb:checked');
    if (cbs.length === 0) { alert('Selecione pelo menos um target.'); return; }
    var pcs = {}; var targets = {};
    for (var i = 0; i < cbs.length; i++) {
        pcs[cbs[i].getAttribute('data-pc')] = true;
        targets[cbs[i].getAttribute('data-target')] = true;
    }
    var pcList = Object.keys(pcs).sort().join(',');
    var targetList = Object.keys(targets).sort().join(',');
    var cmd = '.\\Clean-ShopDisks.ps1 -Computers "' + pcList + '" -Targets "' + targetList + '" -WhatIf';
    var box = document.getElementById('commandBox');
    box.textContent = cmd;
    box.style.display = 'block';
    document.getElementById('copyBtn').style.display = 'inline-block';
}
function copyCommand() {
    var cmd = document.getElementById('commandBox').textContent;
    navigator.clipboard.writeText(cmd).then(function() {
        var btn = document.getElementById('copyBtn');
        btn.textContent = 'Copiado!';
        setTimeout(function() { btn.textContent = 'Copiar Comando'; }, 2000);
    });
}
</script>
"@

    $html = @"
<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Auditoria de Disco - Lojas Artsana</title>
    $css
</head>
<body>
    <div class="header">
        <h1>Auditoria de Disco - Lojas Artsana Portugal</h1>
        <p>Data: $reportDate | Duracao: $([math]::Round($duration.TotalMinutes, 1)) minutos</p>
    </div>

    <div class="summary">
        <div class="summary-card">
            <div class="number">$($Results.Count)</div>
            <div class="label">PCs Auditados</div>
        </div>
        <div class="summary-card">
            <div class="number">$($OfflineServers.Count)</div>
            <div class="label">PCs Offline</div>
        </div>
        <div class="summary-card">
            <div class="number">$($grouped.Keys.Count)</div>
            <div class="label">Lojas</div>
        </div>
    </div>

    <div class="shop-overview">
        <h2>Resumo por Computador (ordenado por menos espaco livre)</h2>
        <table>
            <tr><th>#</th><th>Computador</th><th>Loja</th><th>Total (GB)</th><th>Usado (GB)</th><th>Livre (GB)</th><th>% Usado</th><th>Estado</th></tr>
"@

    $rankIdx = 0
    foreach ($stat in $pcStats) {
        $rankIdx++
        $rowClass = if ($stat.PctUsed -ge 85) { "disk-red" } elseif ($stat.PctUsed -ge 70) { "disk-yellow" } else { "disk-green" }
        $barClass = if ($stat.PctUsed -ge 85) { "bar-red" } elseif ($stat.PctUsed -ge 70) { "bar-yellow" } else { "bar-green" }
        $html += "            <tr class=`"$rowClass`"><td>$rankIdx</td><td><strong>$($stat.ComputerName)</strong></td><td>$($stat.ShopId)</td><td>$($stat.TotalGB)</td><td>$($stat.UsedGB)</td><td>$($stat.FreeGB)</td><td>$($stat.PctUsed)%</td><td><div class=`"bar-container`"><div class=`"bar-fill $barClass`" style=`"width:$($stat.PctUsed)%`"></div></div></td></tr>`n"
    }
    $html += @"
        </table>
    </div>
"@

    # --- Cleanup Analysis Section ---
    $cleanupPCs = @($Results | Where-Object { $_.CleanupTargets -and $_.CleanupTargets.Count -gt 0 })
    if ($cleanupPCs.Count -gt 0) {
        # Ordenar PCs por tamanho total de cleanup (maior primeiro)
        $cleanupPCs = $cleanupPCs | Sort-Object { ($_.CleanupTargets | ForEach-Object { $_.SizeMB } | Measure-Object -Sum).Sum } -Descending

        $html += @"

    <div class="cleanup-section">
        <h2>Analise de Cleanup</h2>
        <div class="cleanup-controls">
            <label><input type="checkbox" onchange="toggleAllPCs(this.checked)"> <strong>Selecionar Todos</strong></label>
        </div>
        <div class="cleanup-summary" id="cleanupSummary"><strong>0</strong> PCs | <strong>0</strong> targets | <strong>0.0 GB</strong> estimados para limpeza</div>
        <table>
            <tr><th class="cleanup-cb-cell" style="width:40px"></th><th>Computador</th><th>Target</th><th>Tamanho (MB)</th><th>Ficheiros</th></tr>
"@
        foreach ($pc in $cleanupPCs) {
            $pcName = $pc.ComputerName
            $pcTotalMB = [math]::Round(($pc.CleanupTargets | ForEach-Object { $_.SizeMB } | Measure-Object -Sum).Sum, 1)
            $html += "            <tr class=`"pc-header-row`"><td class=`"cleanup-cb-cell`"><input type=`"checkbox`" class=`"cleanup-pc-all`" onchange=`"toggleAllTargets('$pcName', this.checked)`"></td><td colspan=`"2`"><strong>$pcName</strong></td><td><strong>$pcTotalMB</strong></td><td></td></tr>`n"
            foreach ($t in ($pc.CleanupTargets | Sort-Object { $_.SizeMB } -Descending)) {
                $html += "            <tr><td class=`"cleanup-cb-cell`"><input type=`"checkbox`" class=`"cleanup-cb`" data-pc=`"$pcName`" data-target=`"$($t.Id)`" data-size=`"$($t.SizeMB)`" onchange=`"updateSummary()`"></td><td></td><td>$($t.Name)</td><td>$($t.SizeMB)</td><td>$($t.FileCount)</td></tr>`n"
            }
        }
        $html += @"
        </table>
        <div class="cleanup-controls" style="margin-top:15px">
            <button class="cleanup-btn" onclick="generateCommand()">Gerar Comando PowerShell</button>
            <button class="cleanup-btn cleanup-btn-copy" id="copyBtn" onclick="copyCommand()" style="display:none">Copiar Comando</button>
        </div>
        <div class="command-box" id="commandBox"></div>
    </div>
"@
    }

    if ($OfflineServers.Count -gt 0) {
        $html += @"

    <div class="error-section">
        <h2>PCs Inacessiveis</h2>
        <table>
            <tr><th>Servidor</th><th>Erro</th></tr>
"@
        foreach ($srv in $OfflineServers) {
            $html += "            <tr><td>$($srv.Name)</td><td>$($srv.Error)</td></tr>`n"
        }
        $html += "        </table>`n    </div>`n"
    }

    $sectionId = 0
    foreach ($shopId in $orderedShopIds) {
        $shopPCs = $grouped[$shopId]
        $html += @"

    <div class="shop-section">
        <h2>Loja $shopId ($($shopPCs.Count) PCs)</h2>
"@

        foreach ($pc in $shopPCs) {
            $sectionId++
            $pcName = $pc.ComputerName

            $html += @"
        <h3 class="collapsible" onclick="toggleSection('sec$sectionId')">$pcName</h3>
        <div id="sec$sectionId" class="collapsible-content">
        <h4>Discos</h4>
        <table>
            <tr><th>Drive</th><th>Label</th><th>Total (GB)</th><th>Livre (GB)</th><th>Usado (GB)</th><th>% Usado</th></tr>
"@
            foreach ($disk in $pc.Disks) {
                $diskClass = if ($disk.PctUsed -ge 85) { "disk-red" } elseif ($disk.PctUsed -ge 70) { "disk-yellow" } else { "disk-green" }
                $html += "            <tr class=`"$diskClass`"><td>$($disk.Drive)</td><td>$($disk.Label)</td><td>$($disk.TotalGB)</td><td>$($disk.FreeGB)</td><td>$($disk.UsedGB)</td><td>$($disk.PctUsed)%</td></tr>`n"
            }
            $html += "        </table>`n"

            $html += @"
        <h4>Top $($pc.TopFiles.Count) Ficheiros Maiores</h4>
        <table>
            <tr><th>#</th><th>Ficheiro</th><th>Tamanho (MB)</th><th>Extensao</th><th>Ultima Modificacao</th></tr>
"@
            $fileIdx = 0
            foreach ($f in $pc.TopFiles) {
                $fileIdx++
                $sizeClass = if ($f.SizeMB -ge 1000) { "size-large" } elseif ($f.SizeMB -ge 100) { "size-medium" } else { "size-small" }
                $html += "            <tr><td>$fileIdx</td><td>$($f.Path)</td><td class=`"$sizeClass`">$($f.SizeMB)</td><td>$($f.Extension)</td><td>$($f.LastModified)</td></tr>`n"
            }
            $html += "        </table>`n"

            $html += @"
        <h4>Top $($pc.TopFolders.Count) Pastas Maiores</h4>
        <table>
            <tr><th>#</th><th>Pasta</th><th>Tamanho (MB)</th><th>Ficheiros</th></tr>
"@
            $folderIdx = 0
            foreach ($folder in $pc.TopFolders) {
                $folderIdx++
                $sizeClass = if ($folder.SizeMB -ge 5000) { "size-large" } elseif ($folder.SizeMB -ge 1000) { "size-medium" } else { "size-small" }
                $html += "            <tr><td>$folderIdx</td><td>$($folder.Path)</td><td class=`"$sizeClass`">$($folder.SizeMB)</td><td>$($folder.FileCount)</td></tr>`n"
            }
            $html += "        </table>`n"

            if ($pc.Errors.Count -gt 0) {
                $html += "        <h4>Erros</h4><ul>`n"
                foreach ($err in $pc.Errors) {
                    $html += "            <li>$err</li>`n"
                }
                $html += "        </ul>`n"
            }

            $html += "        </div>`n"
        }

        $html += "    </div>`n"
    }

    $html += @"

    $js
</body>
</html>
"@

    return $html
}

# ============================================================
# MAIN - Execucao Principal
# ============================================================

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  AUDITORIA DE DISCO - LOJAS ARTSANA" -ForegroundColor Cyan
Write-Host "  (via PsExec)" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Log "Inicio da auditoria. $($servidores.Count) servidores configurados." "INFO"
Write-Log "Parametros: TopFiles=$TopFiles | TopFolders=$TopFolders | Timeout=${TimeoutSeconds}s | Paralelo=$MaxParallel" "INFO"

# --- FASE 1: PsExec ---
Write-Host ""
Write-Log "FASE 1 - Verificar PsExec" "INFO"

if ($PsExecPath -and (Test-Path $PsExecPath)) {
    $psExecExe = $PsExecPath
    Write-Log "PsExec fornecido: $psExecExe" "OK"
} else {
    $toolsDir = Join-Path $PSScriptRoot "Tools"
    $psExecExe = Install-PsExec -DestPath $toolsDir
    if (-not $psExecExe) {
        Write-Log "Impossivel obter PsExec. A sair." "ERROR"
        exit 1
    }
}

# --- FASE 2: Preparar script de recolha ---
Write-Host ""
Write-Log "FASE 2 - Preparar script de recolha local" "INFO"

if (-not (Test-Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
}
$localCollector = Join-Path $TempFolder $CollectorScript
New-CollectorScript -OutputPath $localCollector
Write-Log "Script de recolha criado: $localCollector" "OK"

# --- FASE 3: Teste de Conectividade ---
Write-Host ""
Write-Log "FASE 3 - Teste de conectividade a $($servidores.Count) servidores..." "INFO"
$connectivity = Test-ServerConnectivity -Servers $servidores
$onlineServers = $connectivity.Online
$offlineServers = $connectivity.Offline

Write-Log "Resultado: $($onlineServers.Count) online, $($offlineServers.Count) offline" "INFO"

if ($onlineServers.Count -eq 0) {
    Write-Log "Nenhum servidor acessivel. A sair." "ERROR"
    exit 1
}

# --- FASE 4: Recolha de dados (paralelo com throttle) ---
Write-Host ""
Write-Log "FASE 4 - A lancar recolha em $($onlineServers.Count) PCs (max $MaxParallel em paralelo)..." "INFO"

$results = @()
$jobList = @()

foreach ($srv in $onlineServers) {
    # Controlar paralelismo
    while (($jobList | Where-Object { $_.State -eq "Running" }).Count -ge $MaxParallel) {
        Start-Sleep -Milliseconds 500
    }

    Write-Log "A lancar: $srv" "INFO"
    $job = Start-Job -ScriptBlock {
        param($Server, $PsExec, $TFiles, $TFolders, $TSec, $TFolder, $CScript, $RFile)

        $remoteTempDir = "\\$Server\C`$\TEMP\AuditDisk"
        $remoteScript  = "\\$Server\C`$\TEMP\AuditDisk\$CScript"
        $remoteResult  = "\\$Server\C`$\TEMP\AuditDisk\$RFile"
        $localCollector = Join-Path $TFolder $CScript

        try {
            if (-not (Test-Path $remoteTempDir)) {
                New-Item -ItemType Directory -Path $remoteTempDir -Force | Out-Null
            }
            Copy-Item -Path $localCollector -Destination $remoteScript -Force

            $psArgs = @("\\$Server", "-accepteula", "-nobanner", "-h",
                "powershell.exe", "-ExecutionPolicy", "Bypass",
                "-File", "C:\TEMP\AuditDisk\$CScript",
                "-TopFiles", $TFiles, "-TopFolders", $TFolders,
                "-OutFile", "C:\TEMP\AuditDisk\$RFile")

            $proc = Start-Process -FilePath $PsExec -ArgumentList $psArgs `
                -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput "NUL" `
                -RedirectStandardError "$env:TEMP\psexec_err_$Server.txt"

            if (Test-Path $remoteResult) {
                $data = Import-Clixml -Path $remoteResult
                return $data
            } else {
                $errContent = ""
                if (Test-Path "$env:TEMP\psexec_err_$Server.txt") {
                    $errContent = Get-Content "$env:TEMP\psexec_err_$Server.txt" -Raw -ErrorAction SilentlyContinue
                }
                return @{ ComputerName = $Server; Disks = @(); TopFiles = @(); TopFolders = @()
                          Errors = @("PsExec exit: $($proc.ExitCode). $errContent") }
            }
        }
        catch {
            return @{ ComputerName = $Server; Disks = @(); TopFiles = @(); TopFolders = @()
                      Errors = @("Erro: $($_.Exception.Message)") }
        }
        finally {
            Remove-Item $remoteResult -Force -ErrorAction SilentlyContinue
            Remove-Item $remoteScript -Force -ErrorAction SilentlyContinue
            Remove-Item $remoteTempDir -Force -Recurse -ErrorAction SilentlyContinue
            Remove-Item "$env:TEMP\psexec_err_$Server.txt" -Force -ErrorAction SilentlyContinue
        }
    } -ArgumentList $srv, $psExecExe, $TopFiles, $TopFolders, $TimeoutSeconds, $TempFolder, $CollectorScript, $ResultFile

    $jobList += $job
}

# Aguardar todos os jobs
Write-Log "A aguardar conclusao de $($jobList.Count) jobs..." "INFO"
$jobList | Wait-Job -Timeout $TimeoutSeconds | Out-Null

foreach ($job in $jobList) {
    if ($job.State -eq "Completed") {
        $data = Receive-Job -Job $job
        if ($data -and $data.ComputerName) {
            $results += $data
            $hasErrors = ($data.Errors -and $data.Errors.Count -gt 0)
            if ($hasErrors -and $data.Disks.Count -eq 0) {
                Write-Log "$($data.ComputerName) - FALHOU: $($data.Errors -join '; ')" "ERROR"
            } else {
                Write-Log "$($data.ComputerName) - dados recolhidos" "OK"
            }
        }
    }
    elseif ($job.State -eq "Running") {
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Write-Log "Job timeout - servidor desconhecido" "WARN"
    }
    else {
        $errInfo = $job.ChildJobs[0].JobStateInfo.Reason.Message
        Write-Log "Job falhou: $errInfo" "ERROR"
    }
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
}

# Separar resultados validos dos falhados
$validResults = @($results | Where-Object { $_.Disks -and $_.Disks.Count -gt 0 })
$failedResults = @($results | Where-Object { -not $_.Disks -or $_.Disks.Count -eq 0 })

foreach ($fail in $failedResults) {
    $offlineServers += @{ Name = $fail.ComputerName; Error = ($fail.Errors -join "; ") }
}

Write-Log "Recolha concluida: $($validResults.Count) PCs com dados, $($failedResults.Count) falhados" "OK"

# --- FASE 5: Gerar relatorio HTML ---
Write-Host ""
Write-Log "FASE 5 - A gerar relatorio HTML..." "INFO"

if (-not (Test-Path $ReportPath)) {
    New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
    Write-Log "Pasta criada: $ReportPath" "INFO"
}

$reportFileName = "AuditDisk_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').html"
$reportFullPath = Join-Path $ReportPath $reportFileName

$htmlContent = New-HtmlReport -Results $validResults -OfflineServers $offlineServers -StartTime $ScriptStartTime
$htmlContent | Out-File -FilePath $reportFullPath -Encoding UTF8

Write-Log "Relatorio gerado: $reportFullPath" "OK"

# Limpar temp local
Remove-Item $TempFolder -Recurse -Force -ErrorAction SilentlyContinue

# Resumo final
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  AUDITORIA CONCLUIDA" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Log "PCs auditados: $($validResults.Count)/$($servidores.Count)" "OK"
Write-Log "PCs offline/erro: $($offlineServers.Count)" $(if ($offlineServers.Count -gt 0) { "WARN" } else { "OK" })
Write-Log "Relatorio: $reportFullPath" "OK"
Write-Log "Duracao total: $([math]::Round(((Get-Date) - $ScriptStartTime).TotalMinutes, 1)) minutos" "INFO"
