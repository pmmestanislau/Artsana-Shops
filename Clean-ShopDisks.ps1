#Requires -Version 5.1
<#
.SYNOPSIS
    Executa limpeza remota de disco nos computadores das lojas Artsana Portugal.
.DESCRIPTION
    Usa PsExec para executar remotamente um script de limpeza em cada PC.
    Suporta 7 categorias de cleanup e modo WhatIf (dry-run).
.PARAMETER Computers
    Lista de computadores separados por virgula (ex: "PT4004W01,PT4004P01")
.PARAMETER Targets
    Lista de targets de cleanup separados por virgula, ou "All" para todos.
    Targets validos: WindowsTemp, UserProfiles, Drivers, SoftwareInstall,
    BACKUPxstore, RetentionFolders, WindowsOld
.PARAMETER TimeoutSeconds
    Timeout em segundos para cada PC remoto (default: 600)
.PARAMETER PsExecPath
    Caminho para PsExec.exe. Se omitido, descarrega automaticamente.
.PARAMETER MaxParallel
    Numero maximo de PCs a limpar em paralelo (default: 5)
.PARAMETER WhatIf
    Modo dry-run: mostra o que seria limpo sem apagar nada.
#>
param(
    [Parameter(Mandatory=$true)][string]$Computers,
    [Parameter(Mandatory=$true)][string]$Targets,
    [int]$TimeoutSeconds = 600,
    [string]$PsExecPath = "",
    [int]$MaxParallel = 5,
    [switch]$WhatIf
)

# --- Configuracao ---
$ValidTargets = @("WindowsTemp", "UserProfiles", "Drivers", "SoftwareInstall", "BACKUPxstore", "RetentionFolders", "WindowsOld")
$ScriptStartTime = Get-Date
$TempFolder = "C:\TEMP\CleanDisk"
$CollectorScript = "CleanupCollect.ps1"
$ResultFile = "cleanup_result.xml"

# --- Parse parametros ---
$computerList = @($Computers -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
if ($Targets -eq "All") {
    $targetList = $ValidTargets
} else {
    $targetList = @($Targets -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
}

# Validar targets
foreach ($t in $targetList) {
    if ($t -notin $ValidTargets) {
        Write-Host "[ERRO] Target invalido: $t" -ForegroundColor Red
        Write-Host "Targets validos: $($ValidTargets -join ', ')" -ForegroundColor Yellow
        exit 1
    }
}

if ($computerList.Count -eq 0) {
    Write-Host "[ERRO] Nenhum computador especificado." -ForegroundColor Red
    exit 1
}

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
        $ping = Test-Connection -ComputerName $srv -Count 1 -Quiet -ErrorAction SilentlyContinue
        if (-not $ping) {
            Write-Log "$srv - SEM RESPOSTA (ping)" "WARN"
            $offline += @{ Name = $srv; Error = "Sem resposta ao ping" }
            continue
        }
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

function New-CleanupCollectorScript {
    param([string]$OutputPath)

    $script = @'
param(
    [string]$Targets = "",
    [switch]$DryRun
)

$targetList = @($Targets -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })

$result = @{
    ComputerName = $env:COMPUTERNAME
    Actions      = @()
    TotalFreedMB = 0
    Errors       = @()
    DryRun       = $DryRun.IsPresent
}

function Remove-FolderContents {
    param([string]$Path, [switch]$DryRun)
    $freed = 0; $count = 0
    if (-not (Test-Path $Path)) { return @{ FreedMB = 0; Count = 0 } }
    try {
        $items = Get-ChildItem -Path $Path -Recurse -File -Force -ErrorAction SilentlyContinue
        $totalSize = ($items | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $totalSize) { $totalSize = 0 }
        $count = ($items | Measure-Object).Count
        if (-not $DryRun) {
            Get-ChildItem -Path $Path -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
        $freed = $totalSize
    } catch {}
    return @{ FreedMB = [math]::Round($freed / 1MB, 2); Count = $count }
}

function Remove-FolderEntirely {
    param([string]$Path, [switch]$DryRun)
    $freed = 0; $count = 0
    if (-not (Test-Path $Path)) { return @{ FreedMB = 0; Count = 0 } }
    try {
        $items = Get-ChildItem -Path $Path -Recurse -File -Force -ErrorAction SilentlyContinue
        $totalSize = ($items | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $totalSize) { $totalSize = 0 }
        $count = ($items | Measure-Object).Count
        if (-not $DryRun) {
            Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
        }
        $freed = $totalSize
    } catch {}
    return @{ FreedMB = [math]::Round($freed / 1MB, 2); Count = $count }
}

foreach ($target in $targetList) {
    try {
        switch ($target) {
            "WindowsTemp" {
                $r = Remove-FolderContents -Path "C:\Windows\Temp" -DryRun:$DryRun
                $result.Actions += @{ Target = $target; Path = "C:\Windows\Temp"; FreedMB = $r.FreedMB; FileCount = $r.Count; DryRun = $DryRun.IsPresent }
                $result.TotalFreedMB += $r.FreedMB
            }
            "UserProfiles" {
                $totalFreed = 0; $totalCount = 0
                $skipProfiles = @("Default", "Public", "defaultuser0", "Default User", "All Users")
                $obsoleteUsers = @(Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -notin $skipProfiles -and $_.Name -notlike "ptpos0*" -and $_.Name -ne "shpsuperpt" })
                foreach ($u in $obsoleteUsers) {
                    $r = Remove-FolderEntirely -Path $u.FullName -DryRun:$DryRun
                    $totalFreed += $r.FreedMB; $totalCount += $r.Count
                }
                $result.Actions += @{ Target = $target; Path = "C:\Users (perfis obsoletos, $($obsoleteUsers.Count) perfis)"; FreedMB = [math]::Round($totalFreed, 2); FileCount = $totalCount; DryRun = $DryRun.IsPresent }
                $result.TotalFreedMB += $totalFreed
            }
            "Drivers" {
                $subs = @("MICROSOFT", "OTHER_SOFT", "SCRIPTS", "SISQUAL", "SOPHOS", "XSTORE_INSTALL")
                $totalFreed = 0; $totalCount = 0
                foreach ($sub in $subs) {
                    $p = "C:\Drivers\$sub"
                    $r = Remove-FolderEntirely -Path $p -DryRun:$DryRun
                    $totalFreed += $r.FreedMB; $totalCount += $r.Count
                }
                $result.Actions += @{ Target = $target; Path = "C:\Drivers (subfolders)"; FreedMB = [math]::Round($totalFreed, 2); FileCount = $totalCount; DryRun = $DryRun.IsPresent }
                $result.TotalFreedMB += $totalFreed
            }
            "SoftwareInstall" {
                $subs = @("SQL", "SCRIPTS", "OFFICE2013", "OFFICE365", "Xstore20")
                $totalFreed = 0; $totalCount = 0
                foreach ($sub in $subs) {
                    $p = "C:\Software_InstallationPT\$sub"
                    $r = Remove-FolderEntirely -Path $p -DryRun:$DryRun
                    $totalFreed += $r.FreedMB; $totalCount += $r.Count
                }
                $result.Actions += @{ Target = $target; Path = "C:\Software_InstallationPT (subfolders)"; FreedMB = [math]::Round($totalFreed, 2); FileCount = $totalCount; DryRun = $DryRun.IsPresent }
                $result.TotalFreedMB += $totalFreed
            }
            "BACKUPxstore" {
                $r = Remove-FolderEntirely -Path "C:\BACKUPxstore" -DryRun:$DryRun
                $result.Actions += @{ Target = $target; Path = "C:\BACKUPxstore"; FreedMB = $r.FreedMB; FileCount = $r.Count; DryRun = $DryRun.IsPresent }
                $result.TotalFreedMB += $r.FreedMB
            }
            "RetentionFolders" {
                $totalFreed = 0; $totalCount = 0
                foreach ($rfPath in @("C:\scanner", "C:\tmp\hh_upload")) {
                    if (Test-Path $rfPath) {
                        $oldItems = Get-ChildItem -Path $rfPath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) }
                        $oSz = ($oldItems | Measure-Object -Property Length -Sum).Sum
                        $oFc = ($oldItems | Measure-Object).Count
                        if ($null -eq $oSz) { $oSz = 0 }
                        if (-not $DryRun -and $oFc -gt 0) {
                            $oldItems | Remove-Item -Force -ErrorAction SilentlyContinue
                        }
                        $totalFreed += [math]::Round($oSz / 1MB, 2); $totalCount += $oFc
                    }
                }
                $result.Actions += @{ Target = $target; Path = "C:\scanner + C:\tmp\hh_upload (>30 days)"; FreedMB = [math]::Round($totalFreed, 2); FileCount = $totalCount; DryRun = $DryRun.IsPresent }
                $result.TotalFreedMB += $totalFreed
            }
            "WindowsOld" {
                if (Test-Path "C:\Windows.old") {
                    # Reset attributes first (Windows.old often has protected attrs)
                    try { & attrib -r -s -h "C:\Windows.old" /s /d 2>$null } catch {}
                    $r = Remove-FolderEntirely -Path "C:\Windows.old" -DryRun:$DryRun
                    $result.Actions += @{ Target = $target; Path = "C:\Windows.old"; FreedMB = $r.FreedMB; FileCount = $r.Count; DryRun = $DryRun.IsPresent }
                    $result.TotalFreedMB += $r.FreedMB
                } else {
                    $result.Actions += @{ Target = $target; Path = "C:\Windows.old"; FreedMB = 0; FileCount = 0; DryRun = $DryRun.IsPresent }
                }
            }
        }
    } catch {
        $result.Errors += "Erro em ${target}: $($_.Exception.Message)"
    }
}

$result.TotalFreedMB = [math]::Round($result.TotalFreedMB, 2)

$outDir = Split-Path $OutFile -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
$result | Export-Clixml -Path $OutFile -Force
'@

    $script = $script -replace '\$OutFile', 'C:\TEMP\CleanDisk\cleanup_result.xml'
    $script | Out-File -FilePath $OutputPath -Encoding ASCII -Force
}

# ============================================================
# MAIN - Execucao Principal
# ============================================================

Write-Host ""
Write-Host "============================================" -ForegroundColor Magenta
Write-Host "  LIMPEZA DE DISCO - LOJAS ARTSANA" -ForegroundColor Magenta
if ($WhatIf) {
    Write-Host "  MODO DRY-RUN (sem apagar ficheiros)" -ForegroundColor Yellow
}
Write-Host "  (via PsExec)" -ForegroundColor Magenta
Write-Host "============================================" -ForegroundColor Magenta
Write-Host ""

Write-Log "Computadores: $($computerList -join ', ')" "INFO"
Write-Log "Targets: $($targetList -join ', ')" "INFO"
if ($WhatIf) { Write-Log "MODO WHATIF ACTIVO - nenhum ficheiro sera apagado" "WARN" }

# Confirmacao interativa
Write-Host ""
Write-Host "Resumo da operacao:" -ForegroundColor White
Write-Host "  PCs:     $($computerList.Count) ($($computerList -join ', '))" -ForegroundColor White
Write-Host "  Targets: $($targetList.Count) ($($targetList -join ', '))" -ForegroundColor White
Write-Host "  WhatIf:  $($WhatIf.IsPresent)" -ForegroundColor White
Write-Host ""
$confirm = Read-Host "Continuar? (S/N)"
if ($confirm -notin @("S", "s", "Y", "y")) {
    Write-Log "Operacao cancelada pelo utilizador." "WARN"
    exit 0
}

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

# --- FASE 2: Preparar script de limpeza ---
Write-Host ""
Write-Log "FASE 2 - Preparar script de limpeza" "INFO"

if (-not (Test-Path $TempFolder)) {
    New-Item -ItemType Directory -Path $TempFolder -Force | Out-Null
}
$localCollector = Join-Path $TempFolder $CollectorScript
New-CleanupCollectorScript -OutputPath $localCollector
Write-Log "Script de limpeza criado: $localCollector" "OK"

# --- FASE 3: Teste de conectividade ---
Write-Host ""
Write-Log "FASE 3 - Teste de conectividade a $($computerList.Count) PCs..." "INFO"
$connectivity = Test-ServerConnectivity -Servers $computerList
$onlineServers = $connectivity.Online
$offlineServers = $connectivity.Offline

Write-Log "Resultado: $($onlineServers.Count) online, $($offlineServers.Count) offline" "INFO"

if ($onlineServers.Count -eq 0) {
    Write-Log "Nenhum PC acessivel. A sair." "ERROR"
    exit 1
}

# --- FASE 4: Execucao paralela ---
Write-Host ""
Write-Log "FASE 4 - A lancar limpeza em $($onlineServers.Count) PCs (max $MaxParallel em paralelo)..." "INFO"

$results = @()
$jobList = @()
$targetsParam = $targetList -join ','
$whatIfParam = if ($WhatIf) { '-DryRun' } else { '' }

foreach ($srv in $onlineServers) {
    while (($jobList | Where-Object { $_.State -eq "Running" }).Count -ge $MaxParallel) {
        Start-Sleep -Milliseconds 500
    }

    Write-Log "A lancar: $srv" "INFO"
    $job = Start-Job -ScriptBlock {
        param($Server, $PsExec, $TSec, $TFolder, $CScript, $RFile, $TargetsStr, $WhatIfStr)

        $remoteTempDir = "\\$Server\C`$\TEMP\CleanDisk"
        $remoteScript  = "\\$Server\C`$\TEMP\CleanDisk\$CScript"
        $remoteResult  = "\\$Server\C`$\TEMP\CleanDisk\$RFile"
        $localCollector = Join-Path $TFolder $CScript

        try {
            if (-not (Test-Path $remoteTempDir)) {
                New-Item -ItemType Directory -Path $remoteTempDir -Force | Out-Null
            }
            Copy-Item -Path $localCollector -Destination $remoteScript -Force

            $psArgs = @("\\$Server", "-accepteula", "-nobanner", "-h",
                "powershell.exe", "-ExecutionPolicy", "Bypass",
                "-File", "C:\TEMP\CleanDisk\$CScript",
                "-Targets", $TargetsStr)
            if ($WhatIfStr -eq '-DryRun') { $psArgs += '-DryRun' }

            $proc = Start-Process -FilePath $PsExec -ArgumentList $psArgs `
                -NoNewWindow -Wait -PassThru `
                -RedirectStandardOutput "NUL" `
                -RedirectStandardError "$env:TEMP\psexec_clean_err_$Server.txt"

            if (Test-Path $remoteResult) {
                $data = Import-Clixml -Path $remoteResult
                return $data
            } else {
                $errContent = ""
                if (Test-Path "$env:TEMP\psexec_clean_err_$Server.txt") {
                    $errContent = Get-Content "$env:TEMP\psexec_clean_err_$Server.txt" -Raw -ErrorAction SilentlyContinue
                }
                return @{ ComputerName = $Server; Actions = @(); TotalFreedMB = 0
                          Errors = @("PsExec exit: $($proc.ExitCode). $errContent") }
            }
        }
        catch {
            return @{ ComputerName = $Server; Actions = @(); TotalFreedMB = 0
                      Errors = @("Erro: $($_.Exception.Message)") }
        }
        finally {
            Remove-Item $remoteResult -Force -ErrorAction SilentlyContinue
            Remove-Item $remoteScript -Force -ErrorAction SilentlyContinue
            Remove-Item $remoteTempDir -Force -Recurse -ErrorAction SilentlyContinue
            Remove-Item "$env:TEMP\psexec_clean_err_$Server.txt" -Force -ErrorAction SilentlyContinue
        }
    } -ArgumentList $srv, $psExecExe, $TimeoutSeconds, $TempFolder, $CollectorScript, $ResultFile, $targetsParam, $whatIfParam

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
            if ($hasErrors -and $data.Actions.Count -eq 0) {
                Write-Log "$($data.ComputerName) - FALHOU: $($data.Errors -join '; ')" "ERROR"
            } else {
                $modeStr = if ($data.DryRun) { "DRY-RUN" } else { "EXECUTADO" }
                Write-Log "$($data.ComputerName) - $modeStr - $($data.TotalFreedMB) MB libertados" "OK"
            }
        }
    }
    elseif ($job.State -eq "Running") {
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Write-Log "Job timeout" "WARN"
    }
    else {
        $errInfo = $job.ChildJobs[0].JobStateInfo.Reason.Message
        Write-Log "Job falhou: $errInfo" "ERROR"
    }
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
}

# --- FASE 5: Relatorio consola ---
Write-Host ""
Write-Log "FASE 5 - Relatorio" "INFO"

$validResults = @($results | Where-Object { $_.Actions -and $_.Actions.Count -gt 0 })
$totalGlobalMB = 0

if ($validResults.Count -gt 0) {
    Write-Host ""
    Write-Host "  RESULTADOS DA LIMPEZA" -ForegroundColor White
    Write-Host "  =====================" -ForegroundColor White

    $modeLabel = if ($WhatIf) { " (DRY-RUN - nada foi apagado)" } else { "" }
    Write-Host "  $modeLabel" -ForegroundColor Yellow
    Write-Host ""

    # Header
    $fmt = "  {0,-15} {1,-20} {2,12} {3,10}"
    Write-Host ($fmt -f "PC", "Target", "MB Libertados", "Ficheiros") -ForegroundColor Cyan
    Write-Host ("  " + ("-" * 60)) -ForegroundColor DarkGray

    foreach ($r in ($validResults | Sort-Object { $_.TotalFreedMB } -Descending)) {
        $first = $true
        foreach ($a in $r.Actions) {
            $pcCol = if ($first) { $r.ComputerName } else { "" }
            $first = $false
            $color = if ($a.FreedMB -ge 1000) { "Red" } elseif ($a.FreedMB -ge 100) { "Yellow" } else { "Green" }
            Write-Host ($fmt -f $pcCol, $a.Target, $a.FreedMB, $a.FileCount) -ForegroundColor $color
        }
        $totalGlobalMB += $r.TotalFreedMB
        Write-Host ("  " + ("-" * 60)) -ForegroundColor DarkGray
    }

    $totalGB = [math]::Round($totalGlobalMB / 1024, 2)
    Write-Host ""
    Write-Host "  TOTAL: $totalGlobalMB MB ($totalGB GB)$modeLabel" -ForegroundColor Magenta
}

# Erros
$failedResults = @($results | Where-Object { $_.Errors -and $_.Errors.Count -gt 0 })
if ($failedResults.Count -gt 0) {
    Write-Host ""
    Write-Host "  ERROS:" -ForegroundColor Red
    foreach ($f in $failedResults) {
        Write-Host "  $($f.ComputerName): $($f.Errors -join '; ')" -ForegroundColor Red
    }
}

# Offline
if ($offlineServers.Count -gt 0) {
    Write-Host ""
    Write-Host "  PCS INACESSIVEIS:" -ForegroundColor Yellow
    foreach ($o in $offlineServers) {
        Write-Host "  $($o.Name): $($o.Error)" -ForegroundColor Yellow
    }
}

# Limpar temp local
Remove-Item $TempFolder -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  LIMPEZA CONCLUIDA" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
Write-Log "PCs processados: $($validResults.Count)/$($computerList.Count)" "OK"
Write-Log "Duracao total: $([math]::Round(((Get-Date) - $ScriptStartTime).TotalMinutes, 1)) minutos" "INFO"
