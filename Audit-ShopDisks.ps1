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
    ComputerName = $env:COMPUTERNAME
    Disks        = @()
    TopFiles     = @()
    TopFolders   = @()
    Errors       = @()
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

    # Calcular totais por loja e ordenar por menos espaco livre
    $shopStats = @()
    foreach ($shopId in $grouped.Keys) {
        $shopPCs = $grouped[$shopId]
        $totalGB = 0; $freeGB = 0; $usedGB = 0
        foreach ($pc in $shopPCs) {
            foreach ($disk in $pc.Disks) {
                $totalGB += $disk.TotalGB
                $freeGB  += $disk.FreeGB
                $usedGB  += $disk.UsedGB
            }
        }
        $pctUsed = if ($totalGB -gt 0) { [math]::Round($usedGB / $totalGB * 100, 1) } else { 0 }
        $shopStats += @{
            ShopId  = $shopId
            PCs     = $shopPCs.Count
            TotalGB = [math]::Round($totalGB, 1)
            FreeGB  = [math]::Round($freeGB, 1)
            UsedGB  = [math]::Round($usedGB, 1)
            PctUsed = $pctUsed
        }
    }
    # Ordenar: menos espaco livre primeiro (mais critico no topo)
    $shopStats = $shopStats | Sort-Object { $_.FreeGB }
    $orderedShopIds = $shopStats | ForEach-Object { $_.ShopId }

    # Totais globais
    $globalTotalGB = [math]::Round(($shopStats | ForEach-Object { $_.TotalGB } | Measure-Object -Sum).Sum, 1)
    $globalFreeGB  = [math]::Round(($shopStats | ForEach-Object { $_.FreeGB }  | Measure-Object -Sum).Sum, 1)
    $globalUsedGB  = [math]::Round(($shopStats | ForEach-Object { $_.UsedGB }  | Measure-Object -Sum).Sum, 1)
    $globalPctUsed = if ($globalTotalGB -gt 0) { [math]::Round($globalUsedGB / $globalTotalGB * 100, 1) } else { 0 }

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
    .global-stats { display: flex; gap: 15px; margin: 15px 0; flex-wrap: wrap; }
    .global-stat { background: #2c3e50; color: white; padding: 15px 20px; border-radius: 8px; min-width: 180px; }
    .global-stat .number { font-size: 22px; font-weight: bold; }
    .global-stat .label { font-size: 11px; color: #bdc3c7; text-transform: uppercase; }
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

    <div class="global-stats">
        <div class="global-stat">
            <div class="number">$globalTotalGB GB</div>
            <div class="label">Total Disco (todas as lojas)</div>
        </div>
        <div class="global-stat">
            <div class="number">$globalUsedGB GB</div>
            <div class="label">Espaco Usado</div>
        </div>
        <div class="global-stat" style="background: $(if ($globalPctUsed -ge 85) { '#e74c3c' } elseif ($globalPctUsed -ge 70) { '#f39c12' } else { '#27ae60' })">
            <div class="number">$globalFreeGB GB</div>
            <div class="label">Espaco Livre ($globalPctUsed% usado)</div>
        </div>
    </div>

    <div class="shop-overview">
        <h2>Resumo por Loja (ordenado por menos espaco livre)</h2>
        <table>
            <tr><th>#</th><th>Loja</th><th>PCs</th><th>Total (GB)</th><th>Usado (GB)</th><th>Livre (GB)</th><th>% Usado</th><th>Estado</th></tr>
"@

    $rankIdx = 0
    foreach ($stat in $shopStats) {
        $rankIdx++
        $rowClass = if ($stat.PctUsed -ge 85) { "disk-red" } elseif ($stat.PctUsed -ge 70) { "disk-yellow" } else { "disk-green" }
        $barClass = if ($stat.PctUsed -ge 85) { "bar-red" } elseif ($stat.PctUsed -ge 70) { "bar-yellow" } else { "bar-green" }
        $html += "            <tr class=`"$rowClass`"><td>$rankIdx</td><td><strong>$($stat.ShopId)</strong></td><td>$($stat.PCs)</td><td>$($stat.TotalGB)</td><td>$($stat.UsedGB)</td><td>$($stat.FreeGB)</td><td>$($stat.PctUsed)%</td><td><div class=`"bar-container`"><div class=`"bar-fill $barClass`" style=`"width:$($stat.PctUsed)%`"></div></div></td></tr>`n"
    }
    $html += @"
        </table>
    </div>
"@

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
