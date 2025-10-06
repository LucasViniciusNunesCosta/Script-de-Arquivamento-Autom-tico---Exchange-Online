# ============================================================
# Script de Arquivamento de E-mails - Exchange Online
# Com Monitoramento em Tempo Real
# ============================================================

# 🔹 Verificar módulo Exchange
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "❌ Módulo 'ExchangeOnlineManagement' não encontrado!" -ForegroundColor Red
    Write-Host "📦 Instale com: Install-Module -Name ExchangeOnlineManagement -Force" -ForegroundColor Yellow
    exit
}

Import-Module ExchangeOnlineManagement

# 🔹 Função para exibir barra de progresso
function Show-Progress {
    param(
        [int]$Atual,
        [int]$Total,
        [string]$Atividade,
        [string]$Status
    )
    
    if ($Total -gt 0) {
        $percentual = [math]::Round(($Atual / $Total) * 100, 2)
        Write-Progress -Activity $Atividade -Status $Status -PercentComplete $percentual
    }
}

# 🔹 Conectar ao Exchange Online
Write-Host "🔐 Conectando ao Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host "✅ Conectado com sucesso!" -ForegroundColor Green
}
catch {
    Write-Host "❌ Erro ao conectar: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# 🔹 Coletar informações
$colaborador = Read-Host "`nDigite o e-mail do colaborador"

# Validar mailbox
try {
    $mailbox = Get-Mailbox -Identity $colaborador -ErrorAction Stop
    Write-Host "✅ Mailbox: $($mailbox.DisplayName)" -ForegroundColor Green
}
catch {
    Write-Host "❌ Mailbox não encontrado!" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# Data limite
$dataLimiteStr = Read-Host "`nDigite a data limite (dd/MM/yyyy)"
try {
    $dataLimite = [datetime]::ParseExact($dataLimiteStr, "dd/MM/yyyy", $null)
}
catch {
    Write-Host "❌ Formato de data inválido!" -ForegroundColor Red
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

# 🔹 Verificar Arquivo Morto
$archiveMailbox = Get-Mailbox -Identity $colaborador | Select-Object ArchiveStatus, ArchiveGuid

if ($archiveMailbox.ArchiveStatus -ne "Active") {
    Write-Host "❌ Arquivo Morto não habilitado!" -ForegroundColor Red
    Write-Host "💡 Execute: Enable-Mailbox -Identity '$colaborador' -Archive" -ForegroundColor Yellow
    Disconnect-ExchangeOnline -Confirm:$false
    exit
}

Write-Host "✅ Arquivo Morto habilitado" -ForegroundColor Green

# 🔹 Coletar estatísticas ANTES
Write-Host "`n📊 Coletando estatísticas iniciais..." -ForegroundColor Cyan

$statsInboxAntes = Get-MailboxFolderStatistics -Identity $colaborador -FolderScope Inbox | 
                   Where-Object { $_.FolderType -eq "Inbox" }
$statsArchiveAntes = Get-MailboxFolderStatistics -Identity $colaborador -FolderScope Archive | 
                     Measure-Object -Property ItemsInFolder -Sum

$itensInboxInicial = $statsInboxAntes.ItemsInFolder
$tamanhoInboxInicial = [math]::Round($statsInboxAntes.FolderSize.ToString().Split("(")[1].Split(" ")[0].Replace(",","") / 1MB, 2)
$itensArchiveInicial = $statsArchiveAntes.Sum

Write-Host "📥 Caixa de Entrada:" -ForegroundColor White
Write-Host "   • Itens: $itensInboxInicial" -ForegroundColor Gray
Write-Host "   • Tamanho: $tamanhoInboxInicial MB" -ForegroundColor Gray
Write-Host "📦 Arquivo Morto:" -ForegroundColor White
Write-Host "   • Itens: $itensArchiveInicial" -ForegroundColor Gray

# 🔹 Calcular dias de retenção
$diasRetencao = [math]::Max(1, (New-TimeSpan -Start $dataLimite -End (Get-Date)).Days)
Write-Host "`n📅 Critério: E-mails com mais de $diasRetencao dias" -ForegroundColor Cyan
Write-Host "📅 Data de corte: $($dataLimite.ToString('dd/MM/yyyy'))" -ForegroundColor Cyan

# 🔹 Criar política de retenção
$timestamp = Get-Date -Format 'yyyyMMddHHmmss'
$tagName = "MoverArquivo_$timestamp"
$policyName = "PolicyArquivo_$timestamp"

try {
    Write-Host "`n🎯 Criando configuração de arquivamento..." -ForegroundColor Cyan
    
    # Criar tag
    New-RetentionPolicyTag -Name $tagName `
                           -Type All `
                           -RetentionEnabled $true `
                           -AgeLimitForRetention $diasRetencao `
                           -RetentionAction MoveToArchive `
                           -Comment "Tag temporária - $(Get-Date)" | Out-Null
    
    Start-Sleep -Seconds 2
    
    # Criar política
    New-RetentionPolicy -Name $policyName -RetentionPolicyTagLinks $tagName | Out-Null
    
    Start-Sleep -Seconds 2
    
    # Salvar política original
    $politicaOriginal = (Get-Mailbox -Identity $colaborador).RetentionPolicy
    
    # Aplicar política
    Set-Mailbox -Identity $colaborador -RetentionPolicy $policyName
    
    Write-Host "✅ Configuração aplicada!" -ForegroundColor Green
    
    # Forçar processamento
    Write-Host "`n🚀 Iniciando processamento..." -ForegroundColor Yellow
    
    try {
        Start-ManagedFolderAssistant -Identity $colaborador
        Write-Host "✅ Assistente ativado" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  Processamento será automático (até 24h)" -ForegroundColor Yellow
    }
    
    # 🔹 MONITORAMENTO EM TEMPO REAL
    Write-Host "`n════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "📊 INICIANDO MONITORAMENTO EM TEMPO REAL" -ForegroundColor Green
    Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "⚠️  Pressione Ctrl+C para interromper o monitoramento" -ForegroundColor Yellow
    Write-Host "   (A movimentação continuará em segundo plano)`n" -ForegroundColor Yellow
    
    $tentativas = 0
    $maxTentativas = 120  # 2 horas (120 x 60 segundos)
    $semMudanca = 0
    $ultimoInbox = $itensInboxInicial
    
    Start-Sleep -Seconds 10  # Aguardar início do processamento
    
    while ($tentativas -lt $maxTentativas) {
        $tentativas++
        
        try {
            # Coletar estatísticas atuais
            $statsInboxAtual = Get-MailboxFolderStatistics -Identity $colaborador -FolderScope Inbox | 
                               Where-Object { $_.FolderType -eq "Inbox" }
            $statsArchiveAtual = Get-MailboxFolderStatistics -Identity $colaborador -FolderScope Archive | 
                                 Measure-Object -Property ItemsInFolder -Sum
            
            $itensInboxAtual = $statsInboxAtual.ItemsInFolder
            $itensArchiveAtual = $statsArchiveAtual.Sum
            $itensMovidos = $itensInboxInicial - $itensInboxAtual
            
            # Calcular progresso
            if ($itensMovidos -gt 0) {
                $percentual = [math]::Round(($itensMovidos / $itensInboxInicial) * 100, 2)
            } else {
                $percentual = 0
            }
            
            # Verificar se houve mudança
            if ($itensInboxAtual -eq $ultimoInbox) {
                $semMudanca++
            } else {
                $semMudanca = 0
                $ultimoInbox = $itensInboxAtual
            }
            
            # Limpar console e exibir status
            Clear-Host
            
            Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
            Write-Host "📊 MONITORAMENTO DE ARQUIVAMENTO" -ForegroundColor Green
            Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
            Write-Host "👤 Usuário: $colaborador" -ForegroundColor White
            Write-Host "📅 Data de corte: $($dataLimite.ToString('dd/MM/yyyy'))" -ForegroundColor White
            Write-Host "⏱️  Tempo decorrido: $([math]::Round($tentativas / 60, 1)) minutos" -ForegroundColor White
            Write-Host ""
            
            # Barra de progresso visual
            $barraCompleta = 50
            $barraPreenchida = [math]::Round(($percentual / 100) * $barraCompleta)
            $barraVazia = $barraCompleta - $barraPreenchida
            $barra = ("█" * $barraPreenchida) + ("░" * $barraVazia)
            
            Write-Host "📈 PROGRESSO: $percentual%" -ForegroundColor Yellow
            Write-Host "[$barra]" -ForegroundColor Cyan
            Write-Host ""
            
            Write-Host "📥 CAIXA DE ENTRADA:" -ForegroundColor White
            Write-Host "   Antes:  $itensInboxInicial itens" -ForegroundColor Gray
            Write-Host "   Agora:  $itensInboxAtual itens" -ForegroundColor $(if($itensInboxAtual -lt $itensInboxInicial){"Green"}else{"Gray"})
            Write-Host "   Movidos: $itensMovidos itens" -ForegroundColor Green
            Write-Host ""
            
            Write-Host "📦 ARQUIVO MORTO:" -ForegroundColor White
            Write-Host "   Antes: $itensArchiveInicial itens" -ForegroundColor Gray
            Write-Host "   Agora: $itensArchiveAtual itens" -ForegroundColor $(if($itensArchiveAtual -gt $itensArchiveInicial){"Green"}else{"Gray"})
            Write-Host "   Ganho: $(($itensArchiveAtual - $itensArchiveInicial)) itens" -ForegroundColor Green
            Write-Host ""
            
            # Status
            if ($itensMovidos -eq 0 -and $tentativas -gt 1) {
                Write-Host "⏳ Status: Aguardando início da movimentação..." -ForegroundColor Yellow
            }
            elseif ($semMudanca -gt 5) {
                Write-Host "✅ Status: Movimentação aparentemente concluída!" -ForegroundColor Green
                Write-Host "   (Sem alterações nos últimos $($semMudanca * 60) segundos)" -ForegroundColor Gray
            }
            else {
                Write-Host "🔄 Status: Processando..." -ForegroundColor Cyan
            }
            
            Write-Host ""
            Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
            Write-Host "Próxima atualização em 60 segundos..." -ForegroundColor DarkGray
            Write-Host "Pressione Ctrl+C para parar o monitoramento" -ForegroundColor DarkGray
            
            # Se não houver mudança por 10 minutos, considerar concluído
            if ($semMudanca -ge 10) {
                Write-Host "`n✅ Processo aparentemente concluído!" -ForegroundColor Green
                break
            }
            
            Start-Sleep -Seconds 60
        }
        catch {
            Write-Host "⚠️  Erro ao coletar estatísticas: $($_.Exception.Message)" -ForegroundColor Yellow
            Start-Sleep -Seconds 60
        }
    }
    
    # Estatísticas finais
    Write-Host "`n`n════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "📊 RELATÓRIO FINAL" -ForegroundColor Green
    Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "✅ Itens movidos: $itensMovidos" -ForegroundColor Green
    Write-Host "📥 Caixa de Entrada: $itensInboxInicial → $itensInboxAtual" -ForegroundColor White
    Write-Host "📦 Arquivo Morto: $itensArchiveInicial → $itensArchiveAtual" -ForegroundColor White
    Write-Host "⏱️  Tempo total: $([math]::Round($tentativas / 60, 1)) minutos" -ForegroundColor White
    Write-Host ""
    
    # Salvar script de limpeza
    $cleanupScript = @"
# Script de Limpeza - Gerado em $(Get-Date)
# Execute APÓS confirmar que a movimentação está completa

Connect-ExchangeOnline

# Restaurar política original
$(if ($politicaOriginal) { "Set-Mailbox -Identity $colaborador -RetentionPolicy '$politicaOriginal'" } else { "Set-Mailbox -Identity $colaborador -RetentionPolicy `$null" })

# Remover recursos temporários
Remove-RetentionPolicy -Identity '$policyName' -Confirm:`$false
Remove-RetentionPolicyTag -Identity '$tagName' -Confirm:`$false

Disconnect-ExchangeOnline -Confirm:`$false

Write-Host "✅ Limpeza concluída!" -ForegroundColor Green
"@
    
    $cleanupFile = "$env:USERPROFILE\Desktop\Cleanup_$timestamp.ps1"
    $cleanupScript | Out-File -FilePath $cleanupFile -Encoding UTF8
    
    Write-Host "🧹 LIMPEZA NECESSÁRIA:" -ForegroundColor Yellow
    Write-Host "   Execute o script salvo em:" -ForegroundColor Gray
    Write-Host "   $cleanupFile" -ForegroundColor Cyan
    Write-Host ""
}
catch {
    Write-Host "`n❌ Erro: $($_.Exception.Message)" -ForegroundColor Red
}

Disconnect-ExchangeOnline -Confirm:$false
Write-Host "`n✅ Conexão encerrada!" -ForegroundColor Green
Pause