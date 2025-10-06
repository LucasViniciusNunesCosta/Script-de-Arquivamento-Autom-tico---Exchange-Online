# 📧 Script de Arquivamento Automático - Exchange Online

Script PowerShell para movimentação automatizada de e-mails antigos de todo o mailbox para o Arquivo Morto (Archive) no Exchange Online, com monitoramento em tempo real.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![Exchange](https://img.shields.io/badge/Exchange-Online-orange)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 📋 Índice

- [Visão Geral](#-visão-geral)
- [Funcionalidades](#-funcionalidades)
- [Pré-requisitos](#-pré-requisitos)
- [Instalação](#-instalação)
- [Como Usar](#-como-usar)
- [Parâmetros](#-parâmetros)
- [Monitoramento](#-monitoramento)
- [Limpeza Pós-Execução](#-limpeza-pós-execução)
- [Segurança](#-segurança)
- [Solução de Problemas](#-solução-de-problemas)
- [Limitações Conhecidas](#-limitações-conhecidas)
- [Contribuindo](#-contribuindo)
- [Licença](#-licença)

---

## 🎯 Visão Geral

Este script automatiza o processo de arquivamento de e-mails no Exchange Online, movendo mensagens antigas de **todas as pastas do mailbox** para o **Arquivo Morto** com base em uma data de corte definida pelo usuário.

### Casos de Uso

- Limpeza de caixas de entrada de colaboradores que saíram da empresa
- Arquivamento em massa de e-mails históricos
- Manutenção preventiva de mailboxes com grande volume
- Migração de dados antigos para arquivo morto antes de mudanças de políticas

---

## ✨ Funcionalidades

### Principais Recursos

- ✅ **Arquivamento Baseado em Data**: Move e-mails anteriores a uma data específica
- 📊 **Monitoramento em Tempo Real**: Acompanhe o progresso com barra visual e estatísticas
- 🔄 **Escopo Completo do Mailbox**: Processa **TODAS as pastas** (Caixa de Entrada, Enviados, etc.)
- 🎨 **Interface Visual**: Console colorido com indicadores de progresso
- 📈 **Estatísticas Detalhadas**: Contadores de antes/depois e itens movidos
- ⏱️ **Detecção Automática de Conclusão**: Identifica quando o processo termina
- 🧹 **Script de Limpeza Automático**: Gera arquivo para remoção de políticas temporárias
- 🔄 **Atualização a Cada 60 Segundos**: Refresh automático das estatísticas

### Métricas Monitoradas

- Total de itens na Caixa de Entrada (antes e depois)
- Total de itens no Arquivo Morto (antes e depois)
- Quantidade de e-mails movidos
- Porcentagem de progresso
- Tempo decorrido
- Status do processamento

---

## 🔧 Pré-requisitos

### Software Necessário

1. **PowerShell 5.1 ou superior**
   ```powershell
   # Verificar versão
   $PSVersionTable.PSVersion
   ```

2. **Módulo Exchange Online Management V2**
   ```powershell
   # Instalar módulo (executar como Administrador)
   Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
   ```

### Permissões Necessárias

O usuário que executa o script precisa das seguintes permissões no Exchange Online:

- ✅ **View-Only Recipients** ou **Mail Recipients**
- ✅ **Retention Management**
- ✅ **Organization Configuration**

### Requisitos do Mailbox

- ✅ Arquivo Morto (Archive) deve estar **habilitado** no mailbox de destino
  ```powershell
  # Habilitar arquivo morto
  Enable-Mailbox -Identity usuario@dominio.com -Archive
  ```

---

## 📥 Instalação

### 1. Clonar o Repositório

```bash
git clone https://github.com/seu-usuario/exchange-archive-script.git
cd exchange-archive-script
```

### 2. Instalar Dependências

```powershell
# Executar como Administrador
Install-Module -Name ExchangeOnlineManagement -Force

# Verificar instalação
Get-Module -ListAvailable -Name ExchangeOnlineManagement
```

### 3. Configurar Permissões

Solicite ao administrador do Exchange que conceda as permissões necessárias:

```powershell
# Exemplo de concessão de permissões
New-ManagementRoleAssignment -Role "Retention Management" -User "seu.usuario@dominio.com"
```

---

## 🚀 Como Usar

### Execução Básica

1. **Abra o PowerShell**
2. **Navegue até o diretório do script**
   ```powershell
   cd C:\caminho\para\o\script
   ```

3. **Execute o script**
   ```powershell
   .\Exchange-Archive-Script.ps1
   ```

### Fluxo de Execução

```
┌─────────────────────────────────────┐
│ 1. Conectar ao Exchange Online      │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 2. Informar e-mail do colaborador   │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 3. Informar data limite (dd/MM/yyyy)│
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 4. Validações automáticas            │
│    - Mailbox existe?                 │
│    - Arquivo morto habilitado?       │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 5. Coleta estatísticas iniciais     │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 6. Cria política de retenção         │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 7. Aplica ao mailbox                 │
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 8. Inicia monitoramento em tempo real│
└──────────────┬──────────────────────┘
               │
┌──────────────▼──────────────────────┐
│ 9. Gera script de limpeza            │
└─────────────────────────────────────┘
```

---

## ⚙️ Parâmetros

### Entrada Interativa

O script solicita informações durante a execução:

| Parâmetro | Formato | Exemplo | Descrição |
|-----------|---------|---------|-----------|
| **E-mail do Colaborador** | `usuario@dominio.com` | `joao.silva@empresa.com` | Endereço de e-mail do mailbox a processar |
| **Data Limite** | `dd/MM/yyyy` | `01/01/2024` | E-mails **anteriores** a esta data serão movidos |

### Exemplo de Entrada

```
Digite o e-mail do colaborador: joao.silva@empresa.com
Digite a data limite (dd/MM/yyyy): 15/03/2024
```

**Resultado**: Todos os e-mails recebidos antes de 15/03/2024 serão movidos para o Arquivo Morto.

---

## 📊 Monitoramento

### Tela de Monitoramento

Durante a execução, você verá uma tela atualizada a cada 60 segundos:

```
════════════════════════════════════════════
📊 MONITORAMENTO DE ARQUIVAMENTO
════════════════════════════════════════════
👤 Usuário: joao.silva@empresa.com
📅 Data de corte: 15/03/2024
⏱️  Tempo: 5.2 min

📈 PROGRESSO: 67.5%
[█████████████████████████████████░░░░░░░░░░░░░░░░░]

📥 CAIXA DE ENTRADA:
   Antes:  3641
   Agora:  1183
   Movidos: 2458

📦 ARQUIVO MORTO:
   Antes: 153
   Agora: 2611
   Ganho: +2458

════════════════════════════════════════════
Próxima verificação em 60 seg...
```

### Indicadores de Status

| Ícone | Status | Significado |
|-------|--------|-------------|
| ⏳ | Aguardando | Processamento ainda não iniciou |
| 🔄 | Processando | Movimentação em andamento |
| ✅ | Concluído | Processo finalizado (sem mudanças por 10 minutos) |

### Interrompendo o Monitoramento

- Pressione **Ctrl+C** para parar o monitoramento
- ⚠️ **Importante**: O processo de movimentação **continua em segundo plano** mesmo após interromper o monitoramento

---

## 🧹 Limpeza Pós-Execução

### Por Que é Necessário?

O script cria **recursos temporários** que devem ser removidos após a conclusão:

- Tag de Retenção temporária
- Política de Retenção temporária

### Script Automático de Limpeza

Ao final da execução, um script é gerado automaticamente em:

```
C:\Users\SeuUsuario\Desktop\Cleanup_TIMESTAMP.ps1
```

### Executar Limpeza

**Aguarde a conclusão completa do arquivamento** (sem mudanças por 10+ minutos), então:

```powershell
# Executar o script de limpeza
.\Cleanup_20251006120943.ps1
```

### Limpeza Manual (Opcional)

Se preferir executar manualmente:

```powershell
# Conectar
Connect-ExchangeOnline

# Remover política do mailbox
Set-Mailbox -Identity usuario@dominio.com -RetentionPolicy $null

# Remover política temporária
Remove-RetentionPolicy -Identity 'PolicyInbox_TIMESTAMP' -Confirm:$false

# Remover tag temporária
Remove-RetentionPolicyTag -Identity 'InboxArchive_TIMESTAMP' -Confirm:$false

# Desconectar
Disconnect-ExchangeOnline -Confirm:$false
```

---

## 🔒 Segurança

### Escopo de Processamento

Este script opera processando **TODO o mailbox**:

- ✅ Processa **TODAS as pastas** (Caixa de Entrada, Enviados, Rascunhos, etc.)
- ✅ Utiliza tag de retenção tipo **"All"** (escopo global)
- ⚠️ **ATENÇÃO**: Itens já no Arquivo Morto também podem ser processados
- ✅ Nenhum e-mail é deletado, apenas movido

### Boas Práticas

1. **Teste em Ambiente de Homologação**
   - Execute primeiro em um mailbox de teste
   - Valide o comportamento antes de aplicar em produção

2. **Backup Preventivo**
   ```powershell
   # Exportar mailbox antes do processo (opcional)
   New-MailboxExportRequest -Mailbox usuario@dominio.com -FilePath "\\servidor\backup\usuario.pst"
   ```

3. **Documentar Execuções**
   - Registre qual usuário, data e quantidade de itens movidos
   - Mantenha log das políticas criadas

4. **Horários Recomendados**
   - Execute fora do horário comercial
   - Evite horários de pico do servidor Exchange

### Autenticação

- O script usa **Modern Authentication** do módulo ExchangeOnlineManagement
- Suporta **MFA (Multi-Factor Authentication)**
- Sessão é automaticamente desconectada ao final

---

## 🐛 Solução de Problemas

### Erro: "Módulo não encontrado"

**Problema**: `ExchangeOnlineManagement` não instalado

**Solução**:
```powershell
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
```

---

### Erro: "Arquivo Morto não habilitado"

**Problema**: Mailbox não possui arquivo morto ativo

**Solução**:
```powershell
Enable-Mailbox -Identity usuario@dominio.com -Archive
```

**Verificar**:
```powershell
Get-Mailbox -Identity usuario@dominio.com | Select-Object ArchiveStatus
```

---

### Erro: "MoveToArchive somente pode ser aplicado a marcas do tipo padrão"

**Problema**: Tag de retenção com tipo incorreto

**Solução**: Use a **versão mais recente** do script que utiliza tipo "Inbox"

---

### Erro: "Falha na chamada ao Serviço Assistentes"

**Problema**: `Start-ManagedFolderAssistant` pode falhar temporariamente

**Impacto**: Processamento ocorrerá automaticamente em até 24 horas

**Solução**: Aguardar o processamento automático (não é crítico)

---

### Processamento Lento

**Sintomas**: Movimentação muito demorada

**Causas Comuns**:
- Volume muito alto de e-mails
- Horário de alta carga no servidor
- Limitações de throttling do Exchange

**Recomendações**:
- Execute fora do horário comercial
- Considere dividir em múltiplas execuções com datas diferentes
- Aguarde pacientemente (pode levar horas para grandes volumes)

---

### Nenhum E-mail Movido

**Causas Possíveis**:

1. **Data Limite Muito Recente**
   - Verifique se existem e-mails anteriores à data informada
   
2. **Política Ainda Processando**
   - Aguarde até 24 horas para processamento completo

3. **Permissões Insuficientes**
   - Valide as permissões do usuário

**Diagnóstico**:
```powershell
# Ver estatísticas atuais
Get-MailboxFolderStatistics -Identity usuario@dominio.com -FolderScope Inbox

# Ver política aplicada
Get-Mailbox -Identity usuario@dominio.com | Select-Object RetentionPolicy

# Ver última execução do assistente
Get-Mailbox -Identity usuario@dominio.com | Select-Object ManagedFolderAssistantTime
```

---

## ⚠️ Limitações Conhecidas

### Limitações Técnicas

1. **Processa Apenas Caixa de Entrada**
   - Subpastas da Caixa de Entrada **não são incluídas**
   - Para processar outras pastas, é necessário modificar o script

2. **Dependência do Managed Folder Assistant**
   - Processamento depende do ciclo do assistente do Exchange
   - Pode levar de minutos a horas dependendo da carga do servidor

3. **Throttling**
   - Exchange Online aplica limitações de taxa
   - Grandes volumes podem ser processados em lotes

4. **Tempo de Replicação**
   - Estatísticas podem ter delay de sincronização
   - Números exibidos podem não ser instantâneos

### Limitações de Escopo

- **Não processa**: Itens Enviados, Rascunhos, ou outras pastas
- **Não deleta**: Apenas move (ação reversível)
- **Não compacta**: Não reduz tamanho de mailbox (Exchange faz automaticamente)

---

## 📈 Melhores Práticas

### Antes da Execução

- [ ] Validar backup do mailbox
- [ ] Confirmar arquivo morto habilitado
- [ ] Notificar usuário (se aplicável)
- [ ] Escolher horário de baixa demanda
- [ ] Documentar data de corte escolhida

### Durante a Execução

- [ ] Monitorar as primeiras execuções
- [ ] Anotar estatísticas iniciais
- [ ] Não interromper o processo Exchange
- [ ] Aguardar indicador de conclusão

### Após a Execução

- [ ] Executar script de limpeza
- [ ] Validar estatísticas finais
- [ ] Documentar resultado
- [ ] Notificar usuário da conclusão
- [ ] Remover políticas temporárias

---

## 🤝 Contribuindo

Contribuições são bem-vindas! Para contribuir:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

### Áreas para Contribuição

- 🐛 Correção de bugs
- ✨ Novas funcionalidades
- 📝 Melhorias na documentação
- 🧪 Adição de testes
- 🌐 Traduções

---

## 📝 Changelog

### v1.0.0 (2025-10-06)

#### Added
- ✨ Arquivamento automático baseado em data
- 📊 Monitoramento em tempo real com barra de progresso
- 🔒 Modo seguro (apenas Caixa de Entrada)
- 🧹 Geração automática de script de limpeza
- 🎨 Interface colorida no console
- ⏱️ Detecção automática de conclusão

#### Security
- 🔐 Suporte a Modern Authentication
- 🔐 Compatibilidade com MFA

---

## 📄 Licença

Este projeto está licenciado sob a **MIT License** - veja o arquivo [LICENSE](LICENSE) para detalhes.

```
MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## 👥 Autores

- Lucas Costa  - [GitHub](https://github.com/LucasViniciusNunesCosta)

---

## 🔗 Links Úteis

- [Exchange Online PowerShell V2](https://docs.microsoft.com/powershell/exchange/exchange-online-powershell-v2)
- [Retention Policies](https://docs.microsoft.com/exchange/security-and-compliance/messaging-records-management/retention-tags-and-policies)
- [Archive Mailboxes](https://docs.microsoft.com/microsoft-365/compliance/enable-archive-mailboxes)
- [PowerShell Documentation](https://docs.microsoft.com/powershell/)

---

<div align="center">

**⭐ Se este projeto foi útil, considere dar uma estrela no GitHub! ⭐**

Made with ❤️ using PowerShell

</div>
