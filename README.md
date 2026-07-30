# MUTDFC — Movimentação do Fluxo de Caixa (Mútuo)

Automação em **Python** que extrai o Razão Contábil (transação **FBL3N**)
referente a Mútuo via **SAP GUI Scripting**, consolida os arquivos diários
em um único CSV limpo e ordenado, classifica os lançamentos por natureza e
gera um resumo por categoria — tudo via menu no terminal.

---

## O que faz

| Opção | Ação |
|-------|------|
| **1 — Extrair + Consolidar** | Executa a FBL3N para cada dia/semana/mês do período, exporta um CSV por iteração, depois consolida tudo em um único arquivo limpo e ordenado. |
| **2 — Classificar** | Lê um consolidado existente e gera um arquivo classificado com a coluna `Classificação` adicionada (IOF, IRRF, Rendimento, Adições ou Baixas). |
| **3 — Resumo** | Lê um arquivo classificado e gera totais de `Montante Razão` por classificação, exibindo no terminal, no LOG e gravando em arquivo. |
| **4 — Executar tudo** | Executa as opções 1, 2 e 3 em sequência. |
| **0 — Sair** | Encerra o programa. |

---

## Pré-requisitos

| Requisito | Detalhe |
|-----------|---------|
| **Windows** | SAP GUI Scripting só funciona em Windows. |
| **SAP GUI instalado** | Versão compatível com Scripting habilitado nas configurações. |
| **Usuário logado** | O script apenas **anexa** à sessão existente; não faz login. |
| **Python 3.8+** | Baixe em [python.org](https://www.python.org/downloads/). |
| **pywin32** | `pip install -r requirements.txt` |

### Habilitar SAP GUI Scripting

No SAP Logon: menu **Personalizar → Opções → Acessibilidade & Scripting →
Scripting** → marcar **Habilitar scripting** e desmarcar as opções de
notificação, se desejado.

---

## Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/brunodesouzadias2605-max/MUTDFC.git
cd MUTDFC

# 2. (Opcional) Crie um ambiente virtual
python -m venv .venv
.venv\Scripts\activate   # Windows

# 3. Instale as dependências
pip install -r requirements.txt
```

---

## Como executar

```bash
python mutdfc.py
```

O script exibe o menu principal no terminal:

```
============================================================
  MUTDFC — Movimentação do Fluxo de Caixa (Mútuo)
============================================================
  1 - Extrair razão (FBL3N) + Consolidar
  2 - Classificar consolidado
  3 - Resumo do classificado
  4 - Executar tudo (Extrair → Classificar → Resumir)
  0 - Sair
------------------------------------------------------------
Opção:
```

### Extração (opção 1 ou 4)

```
Data inicial (dd.mm.aaaa): 01.06.2026
Data final   (dd.mm.aaaa): 30.06.2026

Periodicidade:
  1 - Diária  (padrão)
  2 - Semanal
  3 - Mensal
Opção [1]: 1

Sistema/Conexão SAP:
  1 - SAP S/4 HANA (PRD)
  2 - SAP ECC 6.0 (PRD)
Opção: 1
```

### Classificar / Resumo (opções 2, 3)

O script exibe o caminho do último arquivo encontrado como sugestão; pressione
Enter para aceitá-lo ou informe outro caminho.

---

## Detecção de dias sem movimento

Após executar a FBL3N (F8), o script lê a barra de status do SAP
(`wnd[0]/sbar`) antes de tentar exportar. Se a mensagem contiver
**"Nenhuma partida selecionada"**, o dia é registrado no LOG como `INFO` e o
script segue para o próximo intervalo sem gerar erro e sem criar CSV vazio.
Isso trata corretamente fins de semana e feriados sem lançamentos (que
causavam o erro 617 *"The virtual key is not enabled"* na versão anterior).

---

## Regras de classificação

A classificação se aplica às **linhas de Mútuo** (conta iniciando com `1202`)
e às linhas das contas especiais abaixo.

| Classificação | Critério |
|---------------|----------|
| **IOF** | O mesmo `Nº doc.` aparece na conta `2104030007`. |
| **IRRF** | O mesmo `Nº doc.` aparece na conta `1103050010`. |
| **Rendimento** | O mesmo `Nº doc.` aparece na conta `4401030001`. |
| **Adições** | Linha de Mútuo sem vínculo especial, com Montante Razão positivo (débito). |
| **Baixas** | Linha de Mútuo sem vínculo especial, com Montante Razão negativo (crédito). |

> **Heurística Adições/Baixas**: a função `classificar_natureza()` usa o sinal
> do campo `Montante Razão` — positivo = Adições, negativo = Baixas. Essa é
> uma convenção baseada no comportamento típico do FBL3N; se a lógica de
> negócio exigir outra interpretação, ajuste os retornos dessa função no topo
> do arquivo `mutdfc.py`.

As linhas das contas IOF/IRRF/Rendimento também recebem a classificação
correspondente, facilitando a totalização no Resumo.

---

## Resumo

O resumo agrupa `Montante Razão` (parse no formato BR: `1.234,56`) por
`Classificação` e exibe:
- Contagem de lançamentos por categoria.
- Total de montante por categoria.
- Total geral.

Exemplo de saída no terminal:

```
------------------------------------------------------------
  RESUMO — Classificado_01.06.2026_a_30.06.2026.csv
------------------------------------------------------------
  Classificação             Lançamentos     Total Montante
------------------------------------------------------------
  Adições                           120       1.500.000,00
  Baixas                             45        -300.000,00
  IOF                                10         -15.000,00
  IRRF                                8          -8.000,00
  Rendimento                         12          40.000,00
------------------------------------------------------------
  TOTAL GERAL                       195       1.217.000,00
------------------------------------------------------------
```

---

## Como configurar

Abra `mutdfc.py` e ajuste as **CONSTANTES** no topo do arquivo:

| Constante | Padrão | Descrição |
|-----------|--------|-----------|
| `PASTA_BASE` | `<pasta do script>` | Pasta raiz do projeto. Altere para um caminho absoluto se necessário. |
| `PASTA_DIARIOS` | `saidas/diarios/` | Onde os CSVs diários são gravados. |
| `PASTA_CONSOLIDADO` | `saidas/consolidado/` | Onde o consolidado, o classificado e o resumo são gravados. |
| `PASTA_LOGS` | `logs/` | Onde os arquivos de LOG são gravados. |
| `CONTA_LAYOUT` | `MUTDFC` | Variante de layout da FBL3N. |
| `USUARIO_SAP` | `MS0000240` | Usuário SAP dono da variante. |
| `SISTEMAS` | _(dict)_ | Descrições e palavras-chave das conexões. |

---

## Modos de periodicidade

| Modo | Comportamento |
|------|---------------|
| **Diária** | Uma extração por dia; arquivo nomeado `dd.mm.aaaa.csv`. |
| **Semanal** | Uma extração por semana (seg → dom). |
| **Mensal** | Uma extração por mês (1º → último dia). |

---

## Estrutura de pastas de saída

Todas as saídas ficam **dentro da pasta do projeto**:

```
MUTDFC/
├── mutdfc.py
├── requirements.txt
├── README.md
├── .gitignore
├── saidas/
│   ├── diarios/
│   │   ├── 01.06.2026.csv          ← extração diária
│   │   ├── 02.06.2026.csv
│   │   └── ...
│   └── consolidado/
│       ├── Consolidado_01.06.2026_a_30.06.2026.csv
│       ├── Classificado_01.06.2026_a_30.06.2026.csv
│       └── Resumo_01.06.2026_a_30.06.2026.csv
└── logs/
    └── MUTDFC_log_20260630_143000.txt
```

As pastas `saidas/` e `logs/` estão no `.gitignore` e **não são versionadas**.

---

## LOG

O arquivo `MUTDFC_log_<timestamp>.txt` registra:

- Parâmetros informados no menu.
- Início/fim de cada extração (com data e caminho do CSV).
- Dias sem partidas (`INFO`: "Dia DD.MM.AAAA sem partidas — nenhum dado a exportar.").
- Erros COM do SAP GUI Scripting (passo, mensagem, contexto).
- Contagens de classificação por categoria.
- Resumo completo por classificação.
- Caminho de todos os arquivos gerados.

Em caso de falha, consulte o LOG para diagnóstico detalhado.

---

## Estrutura do projeto

```
MUTDFC/
├── mutdfc.py          # Script principal
├── requirements.txt   # Dependências Python
├── README.md          # Esta documentação
└── .gitignore         # Ignora saídas, logs, venv e artefatos
```

---

## Dependências

| Pacote | Uso |
|--------|-----|
| `pywin32` | Acesso ao SAP GUI Scripting via COM (`win32com.client`) |

Bibliotecas da stdlib usadas: `os`, `glob`, `re`, `logging`, `datetime`,
`calendar`, `sys`.
