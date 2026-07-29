# MUTDFC — Movimentação do Fluxo de Caixa (Mútuo)

Automação em **Python** que extrai o Razão Contábil (transação **FBL3N**)
referente a Mútuo via **SAP GUI Scripting** e consolida os arquivos diários
em um único CSV com cabeçalho padronizado e LOG detalhado.

---

## O que faz

1. Exibe um **menu interativo no terminal** (sem janelas/tkinter) pedindo:
   - **Data inicial** e **data final** (formato `dd.mm.aaaa`).
   - **Periodicidade**: Diária (padrão), Semanal ou Mensal.
   - **Sistema SAP**: `1` = S/4 HANA (PRD) · `2` = ECC 6.0 (PRD).
2. **Conecta à sessão SAP já aberta** do sistema escolhido — sem novo login.
3. Executa a transação **FBL3N** (variante `MUTDFC`, usuário `MS0000240`)
   para cada dia/semana/mês do período, exportando um **CSV por iteração**.
4. **Consolida** todos os CSVs num único arquivo
   `Consolidado_<data_ini>_a_<data_fim>.csv` com cabeçalho único e
   encoding UTF-8 (acentos corretos).
5. Grava um **LOG detalhado** com timestamps de cada etapa.

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

O script perguntará no terminal:

```
============================================================
  MUTDFC — Extração FBL3N (Razão Contábil de Mútuo)
============================================================
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

Entradas inválidas são rejeitadas com mensagem explicativa e o campo é
reapresentado.

---

## Como configurar

Abra `mutdfc.py` e ajuste as **CONSTANTES** no topo do arquivo:

| Constante | Padrão | Descrição |
|-----------|--------|-----------|
| `PASTA_TRABALHO` | `~/MUTDFC_Extracao` | Pasta onde os CSVs, o consolidado e o LOG são gravados. A pasta é criada automaticamente. |
| `CONTA_LAYOUT` | `MUTDFC` | Variante de layout da FBL3N (campo `txtV-LOW`). |
| `USUARIO_SAP` | `MS0000240` | Usuário SAP dono da variante (campo `txtENAME-LOW`). |
| `SISTEMAS` | _(dict)_ | Descrições e palavras-chave para localizar as conexões abertas. Edite se os nomes do seu SAP Logon forem diferentes. |

---

## Modos de periodicidade

| Modo | Comportamento |
|------|---------------|
| **Diária** | Uma extração por dia; arquivo nomeado `dd.mm.aaaa.csv`. |
| **Semanal** | Uma extração por semana (seg → dom); arquivo nomeado pela data inicial do intervalo. |
| **Mensal** | Uma extração por mês (1º → último dia); arquivo nomeado pela data inicial do mês. |

---

## Saídas geradas

Todos os arquivos são gravados em `PASTA_TRABALHO` (padrão: `~/MUTDFC_Extracao`):

```
MUTDFC_Extracao/
├── 01.06.2026.csv          ← extração diária
├── 02.06.2026.csv
│   ...
├── 30.06.2026.csv
├── Consolidado_01.06.2026_a_30.06.2026.csv   ← arquivo único
└── MUTDFC_log_20260630_143000.txt            ← LOG
```

### Arquivo consolidado

- **Cabeçalho único** no topo (UTF-8, acentos corretos).
- Linhas de dados de cada CSV em ordem cronológica.
- Cabeçalhos e separadores repetidos dos individuais são removidos.

---

## LOG

O arquivo `MUTDFC_log_<timestamp>.txt` registra:

- Parâmetros informados no menu.
- Início/fim de cada extração (com data e caminho do CSV).
- Encoding detectado em cada arquivo.
- Erros COM do SAP GUI Scripting (passo, mensagem, contexto).
- Caminho do consolidado gerado.

Em caso de falha, consulte o LOG para diagnóstico detalhado.

---

## Estrutura do projeto

```
MUTDFC/
├── mutdfc.py          # Script principal
├── requirements.txt   # Dependências Python
├── README.md          # Esta documentação
└── .gitignore         # Ignora CSVs, logs, venv e artefatos
```

---

## Dependências

| Pacote | Uso |
|--------|-----|
| `pywin32` | Acesso ao SAP GUI Scripting via COM (`win32com.client`) |

Bibliotecas da stdlib usadas: `os`, `csv`, `logging`, `datetime`, `calendar`, `sys`.
