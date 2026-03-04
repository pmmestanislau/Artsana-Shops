# Design: Audit-ShopDisks - Auditoria de Disco das Lojas

**Data:** 2026-03-04
**Autor:** pestanislau + Claude
**Estado:** Aprovado

## Objetivo

Script PowerShell para auditar remotamente o espaco em disco de todos os computadores das lojas Artsana Portugal. Fase 1: identificar os ficheiros e pastas que ocupam mais espaco.

## Requisitos

- Acesso via PowerShell Remoting (WinRM) a 66 computadores em 22 lojas
- Auditar todos os discos fixos de cada PC
- Listar os top 50 ficheiros maiores por PC
- Listar as top 20 pastas maiores por PC
- Relatorio HTML com tabelas organizadas por loja
- Execucao paralela (todos os PCs ao mesmo tempo)

## Arquitetura

### Estrutura de Ficheiros

```
Shops/
  Audit-ShopDisks.ps1    # Script principal
  Reports/               # Pasta criada automaticamente para os HTML
```

### Fases de Execucao

| Fase | Descricao |
|------|-----------|
| 1 | Preparacao: credenciais (Get-Credential), teste conectividade (Test-WSMan) |
| 2 | Recolha: Invoke-Command -AsJob para todos os PCs acessiveis em paralelo |
| 3 | Processamento: agrupa resultados por loja (prefixo PT4XXX) |
| 4 | Relatorio: gera HTML com tabelas por loja |

### Dados Recolhidos por PC

| Dado | Descricao |
|------|-----------|
| Info do disco | Drive letter, tamanho total, espaco livre, % ocupacao |
| Top 50 ficheiros | Path, tamanho (MB), data de modificacao |
| Top 20 pastas | Path, tamanho total (MB), numero de ficheiros |

### Relatorio HTML

- Cabecalho com data/hora da auditoria e resumo geral
- Seccao por loja (agrupado pelo prefixo PT4XXX)
  - Tabela de estado dos discos (semaforo: verde <70%, amarelo 70-85%, vermelho >85%)
  - Tabela dos top 50 ficheiros por PC
  - Tabela das top 20 pastas por PC
- Seccao de erros (PCs inacessiveis ou com falhas)
- CSS inline para portabilidade (abre em qualquer browser)

### Execucao Paralela

- Invoke-Command -AsJob cria jobs nativos do PowerShell Remoting
- Wait-Job -Timeout (configuravel, default 10 minutos)
- PCs que nao respondem sao reportados na seccao de erros

### Parametros do Script

```powershell
param(
    [int]$TopFiles = 50,
    [int]$TopFolders = 20,
    [int]$TimeoutMinutes = 10,
    [string]$ReportPath = ".\Reports"
)
```

### Credenciais

Get-Credential uma vez, reutiliza para todos os Invoke-Command.

## Lista de Servidores (66 PCs, 22 lojas)

```
PT4004: PT4004W01, PT4004P02, PT4004P01
PT4006: PT4006P01, PT4006P02, PT4006W01
PT4010: PT4010P01, PT4010P02, PT4010W01
PT4012: PT4012P01, PT4012P02, PT4012W01
PT4015: PT4015P01, PT4015P02, PT4015W01
PT4018: PT4018P01, PT4018P02, PT4018W01
PT4023: PT4023P01, PT4023P02, PT4023W01
PT4025: PT4025P01, PT4025P02, PT4025W01
PT4026: PT4026P01, PT4026P02, PT4026W01
PT4029: PT4029P01, PT4029P02, PT4029W01
PT4030: PT4030P01, PT4030P02, PT4030W01
PT4031: PT4031P01, PT4031P02, PT4031W01
PT4032: PT4032P01, PT4032P02, PT4032W01
PT4033: PT4033P01, PT4033P02, PT4033W01
PT4034: PT4034P01, PT4034P02, PT4034W01
PT4035: PT4035P01, PT4035P02, PT4035W01
PT4036: PT4036P01, PT4036P02, PT4036W01
PT4037: PT4037P01, PT4037P02, PT4037W01
PT4043: PT4043P01, PT4043P02, PT4043W01
PT4049: PT4049P01, PT4049P02, PT4049W01
PT4094: PT4094P01, PT4094P02, PT4094W01
PT4095: PT4095P01, PT4095P02, PT4095W01
PT4097: PT4097P01, PT4097P02, PT4097P03, PT4097P04, PT4097P05, PT4097W01
```

## Decisoes de Design

- **Script unico** em vez de modulo: e fase 1, simplicidade acima de tudo
- **HTML** em vez de CSV: mais visual e util para apresentar resultados
- **Paralelo total**: 66 PCs nao e carga excessiva para WinRM
- **Apenas ASCII** no script: consistente com os outros scripts Artsana
