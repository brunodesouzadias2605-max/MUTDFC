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
| **2 — Classificar** | Lê um consolidado existente e gera um arquivo classificado com a coluna `Classificação` adicionada (somente linhas de Mútuo `1202*`). |
| **3 — Resumo** | Lê um arquivo classificado e gera totais de `Montante Razão` por classificação, exibindo no terminal, no LOG e gravando em CSV e Excel. |
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
| **openpyxl** | Incluído no `requirements.txt`; usado para geração do Excel de resumo. |

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

O script exibe o nome do último arquivo encontrado como sugestão; pressione
Enter para aceitá-lo ou informe outro nome/caminho.

- Pode digitar apenas o nome do arquivo (com ou sem `.csv`); o script
  procurará automaticamente na pasta `saidas/consolidado`.
- Se informar um caminho absoluto, será usado diretamente.
- A extensão `.csv` é adicionada automaticamente se omitida.

---

## Detecção de dias sem movimento

Após executar a FBL3N (F8), o script lê a barra de status do SAP
(`wnd[0]/sbar`) antes de tentar exportar. Se a mensagem contiver
**"Nenhuma partida selecionada"**, o dia é registrado no LOG como `INFO` e o
script segue para o próximo intervalo sem gerar erro e sem criar CSV vazio.
Isso trata corretamente fins de semana e feriados sem lançamentos (que
causavam o erro 617 *"The virtual key is not enabled"* na versão anterior).

---

## Contas de referência para classificação

| Constante | Conta | Descrição |
|-----------|-------|-----------|
| `CONTA_IOF` | `2104030007` | IOF sobre operação de Mútuo |
| `CONTA_IRRF` | `1103050050` | IRRF sobre rendimento de Mútuo |
| `CONTA_JUROS` | `4401020004` | Receita de juros de Mútuo |
| `PREFIXO_MUTUO` | `1202*` | Contas de Mútuo (todas que iniciam com `1202`) |

> Para corrigir ou atualizar os números de conta, altere apenas as
> constantes `CONTA_IOF`, `CONTA_IRRF` e `CONTA_JUROS` no topo de
> `mutdfc.py`.

---

## Regras de classificação

A classificação se aplica **exclusivamente** às linhas de Mútuo
(conta iniciando com `1202`). As linhas de contrapartida (IOF/IRRF/Juros)
**não** recebem classificação e **não** entram nos totais do resumo — servem
apenas como referência de vínculo para evitar dupla contagem.

| Prioridade | Critério | Classificação |
|------------|----------|---------------|
| 1 | O mesmo `Nº doc.` aparece na conta `2104030007` | **IOF** |
| 2 | O mesmo `Nº doc.` aparece na conta `1103050050` | **IRRF** |
| 3 | O mesmo `Nº doc.` aparece na conta `4401020004` | **Juros** |
| 4a | Nenhum vínculo especial e `CL == 09` | **Adições** |
| 4b | Nenhum vínculo especial e `CL == 19` | **Baixa** |
| — | `CL` com valor diferente de `09`/`19` | *(sem classificação — logado)* |

> A coluna `CL` é lida de forma robusta: `'9'` é tratado como `'09'`.
> Linhas de Mútuo com `CL` desconhecido ficam sem classificação e são
> registradas no LOG para diagnóstico.

---

## Resumo

O resumo agrupa `Montante Razão` por `Classificação`, considerando **apenas**
as 5 categorias abaixo, **nesta ordem**:

```
Adições
Juros
IOF
IRRF
Baixa
```

Para cada categoria: contagem de lançamentos e total de montante.
Inclui linha de **TOTAL GERAL** ao final.

Exemplo de saída no terminal:

```
--------------------------------------------------------------
  RESUMO — Classificado_01.06.2026_a_30.06.2026.csv
--------------------------------------------------------------
  Classificação         Lançamentos       Total Montante
--------------------------------------------------------------
  Adições                       120         1.500.000,00
  Juros                          12            40.000,00
  IOF                            10           15.000,00-
  IRRF                            8            8.000,00-
  Baixa                          45          300.000,00-
--------------------------------------------------------------
  TOTAL GERAL                   195         1.217.000,00
--------------------------------------------------------------
```

### Saídas geradas

| Arquivo | Descrição |
|---------|-----------|
| `Resumo_<ini>_a_<fim>.csv` | CSV em UTF-8 com BOM (`utf-8-sig`) |
| `Resumo_<ini>_a_<fim>.xlsx` | Excel com cabeçalho em negrito, larguras ajustadas e formato numérico contábil |

---

## Correção de encoding

Os CSVs exportados pelo SAP podem vir em encodings diferentes. O script usa
a função `ler_csv_corrigindo_encoding()` que:

1. Tenta `utf-8-sig` / `utf-8` primeiro.
2. Se falhar ou detectar padrões de mojibake (`Ã§`, `Ã£`, `Ãµ` etc.),
   tenta `cp1252` / `latin-1` e corrige via `encode('latin-1').decode('utf-8')`.
3. Loga o encoding detectado e se houve correção por arquivo.

Todas as saídas (consolidado, classificado, resumo CSV) são gravadas em
**`utf-8-sig`** com acentos corretos (`Razão`, `Nº doc.`, `Dt.Lçto`,
`Classificação`, `Adições`, `Mútuo`, `Correção`).

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
| `CONTA_IOF` | `2104030007` | Conta de IOF. |
| `CONTA_IRRF` | `1103050050` | Conta de IRRF. |
| `CONTA_JUROS` | `4401020004` | Conta de receita de juros. |
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
│       ├── Resumo_01.06.2026_a_30.06.2026.csv
│       └── Resumo_01.06.2026_a_30.06.2026.xlsx
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
- Encoding detectado por arquivo (e se houve correção de mojibake).
- Erros COM do SAP GUI Scripting (passo, mensagem, contexto).
- Contagens de classificação por categoria (linhas Mútuo).
- Linhas de Mútuo sem classificação (valor de CL desconhecido).
- Quantidade de linhas não-Mútuo ignoradas na classificação.
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
| `openpyxl` | Geração do arquivo Excel de resumo (`.xlsx`) |

Bibliotecas da stdlib usadas: `os`, `glob`, `re`, `logging`, `datetime`,
`calendar`, `sys`.
