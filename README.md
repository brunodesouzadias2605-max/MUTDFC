# MUTDFC — Movimentação do Fluxo de Caixa (Mútuo)

Automação em **VBScript (SAP GUI Scripting)** para extrair o Razão Contábil
(transação **FBL3N**) referente a **Mútuo** e consolidar as extrações diárias
em um único arquivo.

## O que o projeto faz

1. **Menu interativo** — ao iniciar, pergunta ao usuário:
   - **Período** (data inicial e data final, formato `dd.mm.aaaa`).
   - **Periodicidade** (Diária, Semanal ou Mensal).
   - **Sistema SAP** (`1` = S/4 HANA PRD, `2` = ECC 6.0 PRD).
2. **Conexão automática** — localiza a sessão SAP já aberta no SAP GUI
   correspondente ao sistema escolhido (sem novo login).
3. **Extração da FBL3N** — executa a transação para cada dia (ou intervalo)
   do período informado, usando a conta `MUTDFC` e o usuário `MS0000240`,
   exportando um CSV por dia/intervalo.
4. **Consolidação** — empilha todos os CSVs gerados em um único arquivo
   `Consolidado_<datainicial>_a_<datafinal>.csv`, com cabeçalho único e
   sem duplicatas, encoding UTF-8.
5. **LOG detalhado** — grava um arquivo `MUTDFC_log_<timestamp>.txt` com
   timestamps em cada etapa para facilitar o diagnóstico de falhas.

## Pré-requisitos

| Requisito | Detalhe |
|-----------|---------|
| **SAP GUI Scripting habilitado** | No SAP GUI: menu *Opções → Scripting → Habilitar scripting*. Solicite ao administrador SAP, se necessário. |
| **Usuário logado** | O usuário deve estar logado na sessão do sistema escolhido **antes** de executar o script. O script **não** realiza login. |
| **Windows com VBScript** | Disponível nativamente em qualquer Windows (executado via `wscript.exe` ou `cscript.exe`). |
| **Permissão de escrita na pasta de trabalho** | A pasta definida na constante `PASTA_TRABALHO` deve ser gravável. O script a cria automaticamente se não existir. |

## Como executar

### Via Windows Explorer (duplo clique)
1. Clique com o botão direito em `MUTDFC.vbs`.
2. Escolha **Abrir com → Windows Script Host (wscript.exe)**.

### Via prompt de comando
```bat
wscript "C:\caminho\para\MUTDFC.vbs"
```

Para ver a saída no console (útil para depuração):
```bat
cscript "C:\caminho\para\MUTDFC.vbs"
```

## Como configurar as constantes

Abra `MUTDFC.vbs` em qualquer editor de texto e ajuste as constantes no
**bloco "CONSTANTES DE CONFIGURAÇÃO"** no topo do arquivo:

| Constante | Descrição | Valor padrão |
|-----------|-----------|--------------|
| `PASTA_TRABALHO` | Pasta onde os CSVs e o consolidado serão salvos. | `C:\MUTDFC\Extrações` |
| `VARIANTE` | Variante de layout da FBL3N. | `MUTDFC` |
| `USUARIO_SAP` | Usuário SAP dono da variante. | `MS0000240` |
| `DESC_S4HANA` | Trecho da descrição da conexão S/4 HANA para localizar a sessão aberta. | `S/4 HANA PRODUCAO` |
| `DESC_ECC` | Trecho da descrição da conexão ECC para localizar a sessão aberta. | `ECC` |
| `CONN_S4HANA_FULL` | Descrição completa usada para abrir nova conexão S/4 HANA, se necessário. | `MRV SAP S/4 HANA PRODUCAO` |
| `CONN_ECC_FULL` | Descrição completa usada para abrir nova conexão ECC, se necessário. | `03 SAP PRD ECC - BOLHA` |

**Exemplo de alteração da pasta de trabalho:**
```vbs
Const PASTA_TRABALHO = "C:\Users\seu.usuario\Documents\MUTDFC"
```

## Modos de periodicidade

| Opção | Comportamento |
|-------|---------------|
| **1 — Diária** (padrão) | Repete a extração para cada dia do período. Ex.: de 01.06.2026 a 30.06.2026 gera 30 arquivos CSV. |
| **2 — Semanal** | Agrupa por semana (7 dias por intervalo), reduzindo o número de extrações. |
| **3 — Mensal** | Agrupa por mês (usa `DateAdd("m", 1, ...)`), uma extração por mês do período. |

## Arquivos gerados

| Arquivo | Descrição |
|---------|-----------|
| `<dd.mm.aaaa>.csv` | CSV diário (ou por intervalo) exportado do SAP, salvo em `PASTA_TRABALHO`. |
| `Consolidado_<datainicial>_a_<datafinal>.csv` | Arquivo final empilhado com cabeçalho único. |
| `MUTDFC_log_<timestamp>.txt` | LOG detalhado da execução. |

## Cabeçalho do arquivo consolidado

```
|   St|Atribuição        |Nº doc.   |Dt.Lçto   |Div |Conta     |Fornecedor|Tp.doc. |Cliente   |Data doc. |CL|      Montante Razão|DocCompens|Texto                                             |Imobilizado |Usuário     |Empr|Referência      |Entrado em|Estorno|DiagRede    |
```

## Encoding

O SAP GUI exporta os CSVs em **Latin-1 (Windows-1252)**. O script lê cada
arquivo nesse encoding e grava o consolidado em **UTF-8**, preservando
corretamente os acentos (`Razão`, `Nº doc.`, `Dt.Lçto` etc.).

## Diagnóstico de falhas

Em caso de erro, o script exibe uma caixa de diálogo indicando o caminho do
arquivo de LOG. O LOG registra:

- Parâmetros informados no menu (período, periodicidade, sistema).
- Início e fim de cada extração individual.
- Erros do SAP GUI Scripting com número e descrição (`Err.Number`,
  `Err.Description`) e o contexto (data, etapa).
- Processo de consolidação (arquivo por arquivo).

Apresente o arquivo de LOG ao suporte para agilizar a resolução do problema.

## Estrutura do repositório

```
MUTDFC/
├── MUTDFC.vbs       ← Script principal (VBScript / SAP GUI Scripting)
├── README.md        ← Esta documentação
└── .gitignore       ← Exclui CSVs, LOGs e artefatos de execução
```
