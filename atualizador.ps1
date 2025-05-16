# Atualizador.ps1 - Windows Update + log local + envio para Google Sheets + bloqueio de firmware + alerta de reinício

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Caminhos
$logPath = "C:\Appmax"
$logFile = Join-Path $logPath "update-log.txt"

if (-Not (Test-Path $logPath)) {
    New-Item -Path $logPath -ItemType Directory -Force
}

Add-Content -Path $logFile -Value "`n===== Início da execução: $(Get-Date) =====" -Encoding utf8

# Verifica e instala PSWindowsUpdate
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop
    Install-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate instalado com sucesso." -Encoding utf8
} else {
    Add-Content -Path $logFile -Value "Módulo PSWindowsUpdate já instalado." -Encoding utf8
}

# Importa e ativa Microsoft Update
Import-Module PSWindowsUpdate -Force
Add-WUServiceManager -MicrosoftUpdate -Confirm:$false

# Informações da máquina
$hostname = $env:COMPUTERNAME
$so = (Get-CimInstance Win32_OperatingSystem).Caption
$dataHoraInicio = Get-Date
$dataHoraTexto = $dataHoraInicio.ToString("yyyy-MM-dd HH:mm:ss")

# Lista atualizações disponíveis
$updatesDisponiveis = Get-WindowsUpdate -MicrosoftUpdate

if ($updatesDisponiveis.Count -eq 0) {
    Add-Content -Path $logFile -Value "Nenhuma atualização pendente encontrada." -Encoding utf8
} else {
    Add-Content -Path $logFile -Value "$($updatesDisponiveis.Count) atualização(ões) detectada(s)." -Encoding utf8

    # Palavras bloqueadas (não instalar Firmware, BIOS, UEFI)
    $palavrasBloqueadas = @("Firmware", "BIOS", "UEFI", "Thunderbolt", "System Firmware", "Dock Firmware")
    $updatesPermitidos = $updatesDisponiveis | Where-Object {
        ($_.Title -notmatch ($palavrasBloqueadas -join "|"))
    }

    if ($updatesPermitidos.Count -gt 0) {
        Add-Content -Path $logFile -Value "$($updatesPermitidos.Count) atualização(ões) permitida(s) para instalação:" -Encoding utf8
        foreach ($update in $updatesPermitidos) {
            Add-Content -Path $logFile -Value "  - $($update.Title)" -Encoding utf8
        }

        try {
            $updatesPermitidos | Install-WindowsUpdate -AcceptAll -IgnoreReboot -MicrosoftUpdate -Verbose |
                Tee-Object -FilePath $logFile -Append

            if ($?) {
                Add-Content -Path $logFile -Value "Comando Install-WindowsUpdate executado com sucesso (somente updates permitidos)." -Encoding utf8
            } else {
                Add-Content -Path $logFile -Value "Comando Install-WindowsUpdate falhou." -Encoding utf8
            }
        } catch {
            Add-Content -Path $logFile -Value "Erro ao instalar atualizações: $($_.Exception.Message)" -Encoding utf8
        }
    } else {
        Add-Content -Path $logFile -Value "Nenhuma atualização permitida para instalação (todas bloqueadas por regra)." -Encoding utf8
    }
}

# Espera para garantir que Get-WUHistory atualizou
Start-Sleep -Seconds 10

# Coleta histórico dos updates instalados recentes e permitidos
$dataHoraLimite = $dataHoraInicio.AddMinutes(-5)
$updatesHist = Get-WUHistory | Where-Object {
    $_.Result -eq "Succeeded" -and
    $_.Date -ge $dataHoraLimite -and
    ($_.Title -notmatch ($palavrasBloqueadas -join "|"))
} | Sort-Object Date -Descending

$titulos = $updatesHist | ForEach-Object { $_.Title }
$todosUpdates = $titulos -join "; "

# Montar JSON seguro
$payload = @{
    data         = $dataHoraTexto
    hostname     = $hostname
    so           = $so
    atualizacoes = $todosUpdates
}
$json = $payload | ConvertTo-Json -Depth 3
$body = [System.Text.Encoding]::UTF8.GetBytes($json)

# Envio para Google Sheets
try {
    $response = Invoke-RestMethod -Uri "https://script.google.com/macros/s/AKfycbwHp-e0DTsSk4u4GK3_m4Lryt7GMXIjxb68qFUsxuqjO5OkgBGQv48UGqitN5AT4WmM/exec" `
        -Method Post `
        -Body $body `
        -ContentType "application/json"
    Add-Content -Path $logFile -Value "Envio para Google Sheets concluído com sucesso. Resposta: $response" -Encoding utf8
} catch {
    Add-Content -Path $logFile -Value "Erro ao enviar para Google Sheets: $($_.Exception.Message)" -Encoding utf8
}

# ⚠️ Após qualquer instalação, sempre solicitar reinício
Add-Content -Path $logFile -Value "⚠️ Alerta de reinicialização ativado (a cada 10 minutos)." -Encoding utf8

do {
    try {
        msg * "ATENCAO: Atualizacoes instaladas. Por favor, conecte o carregador e reinicie seu computador para concluir a aplicacao das atualizacoes de seguranca (Att. TI APPMAX)"
        Add-Content -Path $logFile -Value "🔔 Alerta enviado via msg.exe em: $(Get-Date)" -Encoding utf8
    } catch {
        Add-Content -Path $logFile -Value "❌ Erro ao exibir mensagem via msg.exe: $($_.Exception.Message)" -Encoding utf8
    }

    # Espera 10 minutos antes de repetir
    Start-Sleep -Seconds 600

    # Opcional: parar se o computador for reiniciado (reboot limpa os paths pendentes)
    $reiniciado = -not (Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations") `
                  -and -not (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired") `
                  -and -not (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
} while (-not $reiniciado)

Add-Content -Path $logFile -Value "✅ Reinicialização detectada. Alertas encerrados em: $(Get-Date)" -Encoding utf8

# Fim do log
Add-Content -Path $logFile -Value "===== Fim da execução: $(Get-Date) =====`n" -Encoding utf8
