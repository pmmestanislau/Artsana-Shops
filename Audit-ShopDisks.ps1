#Requires -Version 5.1
<#
.SYNOPSIS
    Audita o espaco em disco dos computadores das lojas Artsana Portugal.
.DESCRIPTION
    Liga-se remotamente via WinRM a todos os PCs das lojas, recolhe informacao
    sobre discos, top ficheiros e top pastas, e gera relatorio HTML.
.PARAMETER TopFiles
    Numero de ficheiros maiores a listar por PC (default: 50)
.PARAMETER TopFolders
    Numero de pastas maiores a listar por PC (default: 20)
.PARAMETER TimeoutMinutes
    Timeout em minutos para cada PC remoto (default: 10)
.PARAMETER ReportPath
    Pasta onde gravar o relatorio HTML (default: .\Reports)
.PARAMETER Credential
    Credenciais explicitas. Se omitido, usa a sessao Windows actual (Kerberos).
#>
param(
    [int]$TopFiles = 50,
    [int]$TopFolders = 20,
    [int]$TimeoutMinutes = 10,
    [string]$ReportPath = ".\Reports",
    [PSCredential]$Credential
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

function Test-ServerConnectivity {
    param(
        [string[]]$Servers,
        [PSCredential]$Credential
    )
    $online = @()
    $offline = @()

    foreach ($srv in $Servers) {
        Write-Log "A testar conectividade: $srv ..." "INFO"
        try {
            $wsmanParams = @{ ComputerName = $srv; ErrorAction = "Stop" }
            if ($Credential) { $wsmanParams.Credential = $Credential }
            $wsmanResult = Test-WSMan @wsmanParams
            if ($wsmanResult) {
                Write-Log "$srv - ONLINE" "OK"
                $online += $srv
            }
        }
        catch {
            Write-Log "$srv - OFFLINE ou inacessivel: $($_.Exception.Message)" "WARN"
            $offline += @{ Name = $srv; Error = $_.Exception.Message }
        }
    }

    return @{
        Online  = $online
        Offline = $offline
    }
}

# --- ScriptBlock remoto (corre em cada PC) ---
$RemoteScriptBlock = {
    param([int]$TopFiles, [int]$TopFolders)

    $result = @{
        ComputerName = $env:COMPUTERNAME
        Disks        = @()
        TopFiles     = @()
        TopFolders   = @()
        Errors       = @()
    }

    try {
        # Discos fixos (DriveType 3 = Local Disk)
        $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3"
        foreach ($disk in $disks) {
            $totalGB = [math]::Round($disk.Size / 1GB, 2)
            $freeGB  = [math]::Round($disk.FreeSpace / 1GB, 2)
            $usedGB  = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 2)
            $pctUsed = if ($disk.Size -gt 0) { [math]::Round(($disk.Size - $disk.FreeSpace) / $disk.Size * 100, 1) } else { 0 }

            $result.Disks += @{
                Drive    = $disk.DeviceID
                Label    = $disk.VolumeName
                TotalGB  = $totalGB
                FreeGB   = $freeGB
                UsedGB   = $usedGB
                PctUsed  = $pctUsed
            }
        }

        # Top ficheiros maiores (todos os discos fixos)
        $allFiles = @()
        foreach ($disk in $disks) {
            $driveLetter = $disk.DeviceID + "\"
            try {
                $files = Get-ChildItem -Path $driveLetter -Recurse -File -ErrorAction SilentlyContinue |
                    Sort-Object -Property Length -Descending |
                    Select-Object -First ($TopFiles * 2)
                $allFiles += $files
            }
            catch {
                $result.Errors += "Erro a listar ficheiros em ${driveLetter}: $($_.Exception.Message)"
            }
        }

        $result.TopFiles = $allFiles |
            Sort-Object -Property Length -Descending |
            Select-Object -First $TopFiles |
            ForEach-Object {
                @{
                    Path         = $_.FullName
                    SizeMB       = [math]::Round($_.Length / 1MB, 2)
                    LastModified = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm")
                    Extension    = $_.Extension
                }
            }

        # Top pastas maiores
        $allFolders = @()
        foreach ($disk in $disks) {
            $driveLetter = $disk.DeviceID + "\"
            try {
                # Pastas de primeiro nivel
                $topDirs = Get-ChildItem -Path $driveLetter -Directory -ErrorAction SilentlyContinue
                foreach ($dir in $topDirs) {
                    try {
                        $items = Get-ChildItem -Path $dir.FullName -Recurse -File -ErrorAction SilentlyContinue
                        $totalSize = ($items | Measure-Object -Property Length -Sum).Sum
                        $fileCount = ($items | Measure-Object).Count
                        if ($null -eq $totalSize) { $totalSize = 0 }
                        $allFolders += @{
                            Path      = $dir.FullName
                            SizeMB    = [math]::Round($totalSize / 1MB, 2)
                            FileCount = $fileCount
                        }
                    }
                    catch {
                        # Ignorar pastas sem permissao
                    }
                }

                # Sub-pastas de segundo nivel (para melhor granularidade)
                foreach ($dir in $topDirs) {
                    try {
                        $subDirs = Get-ChildItem -Path $dir.FullName -Directory -ErrorAction SilentlyContinue
                        foreach ($sub in $subDirs) {
                            try {
                                $items = Get-ChildItem -Path $sub.FullName -Recurse -File -ErrorAction SilentlyContinue
                                $totalSize = ($items | Measure-Object -Property Length -Sum).Sum
                                $fileCount = ($items | Measure-Object).Count
                                if ($null -eq $totalSize) { $totalSize = 0 }
                                $allFolders += @{
                                    Path      = $sub.FullName
                                    SizeMB    = [math]::Round($totalSize / 1MB, 2)
                                    FileCount = $fileCount
                                }
                            }
                            catch {
                                # Ignorar pastas sem permissao
                            }
                        }
                    }
                    catch {
                        # Ignorar
                    }
                }
            }
            catch {
                $result.Errors += "Erro a listar pastas em ${driveLetter}: $($_.Exception.Message)"
            }
        }

        $result.TopFolders = $allFolders |
            Sort-Object -Descending { $_.SizeMB } |
            Select-Object -First $TopFolders
    }
    catch {
        $result.Errors += "Erro geral: $($_.Exception.Message)"
    }

    return $result
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
    .pc-name { font-weight: bold; color: #2980b9; }
    .size-large { color: #e74c3c; font-weight: bold; }
    .size-medium { color: #e67e22; }
    .size-small { color: #27ae60; }
    .collapsible { cursor: pointer; user-select: none; }
    .collapsible::before { content: '[-] '; font-family: monospace; }
    .collapsible.collapsed::before { content: '[+] '; }
    .collapsible-content { overflow: hidden; }
    .collapsible-content.hidden { display: none; }
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
"@

    # Seccao de erros (PCs offline)
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

    # Seccao por loja
    $sectionId = 0
    foreach ($shopId in ($grouped.Keys | Sort-Object)) {
        $shopPCs = $grouped[$shopId]
        $html += @"

    <div class="shop-section">
        <h2>Loja $shopId ($($shopPCs.Count) PCs)</h2>
"@

        foreach ($pc in $shopPCs) {
            $sectionId++
            $pcName = $pc.ComputerName

            # Tabela de discos
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

            # Tabela de top ficheiros
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

            # Tabela de top pastas
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

            # Erros do PC
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
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Log "Inicio da auditoria. $($servidores.Count) servidores configurados." "INFO"
Write-Log "Parametros: TopFiles=$TopFiles | TopFolders=$TopFolders | Timeout=${TimeoutMinutes}min" "INFO"

# --- FASE 1: Credenciais ---
Write-Host ""
Write-Log "FASE 1 - Credenciais" "INFO"
if (-not $Credential) {
    Write-Log "A usar sessao Windows actual (Kerberos): $env:USERDOMAIN\$env:USERNAME" "OK"
} else {
    Write-Log "Credenciais explicitas para: $($Credential.UserName)" "OK"
}

# --- FASE 2: Teste de Conectividade ---
Write-Host ""
Write-Log "FASE 2 - Teste de conectividade a $($servidores.Count) servidores..." "INFO"
$connectivity = Test-ServerConnectivity -Servers $servidores -Credential $Credential
$onlineServers = $connectivity.Online
$offlineServers = $connectivity.Offline

Write-Log "Resultado: $($onlineServers.Count) online, $($offlineServers.Count) offline" "INFO"

if ($onlineServers.Count -eq 0) {
    Write-Log "Nenhum servidor acessivel. A sair." "ERROR"
    exit 1
}

# --- FASE 3: Recolha de dados (paralelo) ---
Write-Host ""
Write-Log "FASE 3 - A lancar recolha de dados em $($onlineServers.Count) PCs (paralelo)..." "INFO"

$invokeParams = @{
    ComputerName = $onlineServers
    ScriptBlock  = $RemoteScriptBlock
    ArgumentList = @($TopFiles, $TopFolders)
    AsJob        = $true
    JobName      = "ShopDiskAudit"
}
if ($Credential) { $invokeParams.Credential = $Credential }
$jobs = Invoke-Command @invokeParams

Write-Log "Jobs lancados. A aguardar resultados (timeout: ${TimeoutMinutes} minutos)..." "INFO"

$null = Wait-Job -Job $jobs -Timeout ($TimeoutMinutes * 60)

$results = @()

foreach ($childJob in $jobs.ChildJobs) {
    if ($childJob.State -eq "Completed") {
        $data = Receive-Job -Job $childJob
        if ($data) {
            $results += $data
            Write-Log "$($data.ComputerName) - dados recolhidos" "OK"
        }
    }
    elseif ($childJob.State -eq "Failed") {
        $errMsg = $childJob.JobStateInfo.Reason.Message
        $srvName = $childJob.Location
        Write-Log "$srvName - FALHOU: $errMsg" "ERROR"
        $offlineServers += @{ Name = $srvName; Error = "Job falhou: $errMsg" }
    }
    else {
        $srvName = $childJob.Location
        Write-Log "$srvName - TIMEOUT (estado: $($childJob.State))" "WARN"
        $offlineServers += @{ Name = $srvName; Error = "Timeout apos $TimeoutMinutes minutos" }
        Stop-Job -Job $childJob -ErrorAction SilentlyContinue
    }
}

Remove-Job -Job $jobs -Force -ErrorAction SilentlyContinue

Write-Log "Recolha concluida: $($results.Count) PCs com dados" "OK"

# --- FASE 4: Gerar relatorio HTML ---
Write-Host ""
Write-Log "FASE 4 - A gerar relatorio HTML..." "INFO"

if (-not (Test-Path $ReportPath)) {
    New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
    Write-Log "Pasta criada: $ReportPath" "INFO"
}

$reportFileName = "AuditDisk_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').html"
$reportFullPath = Join-Path $ReportPath $reportFileName

$htmlContent = New-HtmlReport -Results $results -OfflineServers $offlineServers -StartTime $ScriptStartTime
$htmlContent | Out-File -FilePath $reportFullPath -Encoding UTF8

Write-Log "Relatorio gerado: $reportFullPath" "OK"

# Resumo final
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  AUDITORIA CONCLUIDA" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Log "PCs auditados: $($results.Count)/$($servidores.Count)" "OK"
Write-Log "PCs offline/erro: $($offlineServers.Count)" $(if ($offlineServers.Count -gt 0) { "WARN" } else { "OK" })
Write-Log "Relatorio: $reportFullPath" "OK"
Write-Log "Duracao total: $([math]::Round(((Get-Date) - $ScriptStartTime).TotalMinutes, 1)) minutos" "INFO"
