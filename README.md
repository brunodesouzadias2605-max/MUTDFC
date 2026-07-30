# MUTDFC — Movimentação de Mútuo (FBL3N + Consolidação Intercompany)

Automação em Python para:

1. Extrair razão contábil (FBL3N) via SAP GUI Scripting.
2. Consolidar e classificar movimentações.
3. Importar tabelas intercompany (`ZFIT009` e `ZCO059`).
4. Cruzar **Cliente → Divisão → Consolida**.
5. Aplicar **Status Consolidação** (`Elimina` / `Não Consolida`).
6. Separar contrapartidas para auditoria.
7. Gerar classificado e resumo em CSV + Excel.
8. Persistir **Saldo Inicial** com histórico.

---

## Menu principal

```
1 - Extrair razão (FBL3N) + Consolidar
2 - Classificar consolidado
3 - Importar ZFIT009 / ZCO059
4 - Gerar tabela de consolidação
5 - Aplicar Status Consolidação na movimentação
6 - Tratar contrapartidas (auditoria) + Excel classificado
7 - Informar/atualizar Saldo Inicial
8 - Resumo (Saldo Inicial → Saldo Final)
9 - Executar tudo em sequência
0 - Sair
```

---

## Consolidação intercompany (regra de cruzamento em 2 etapas)

1. **ZFIT009**: busca a **Divisão** a partir do **Cliente**.
2. **ZCO059**: busca o **Consolida** (`S`/vazio) a partir da **Divisão**.

Saída:

- `saidas/consolidado/Consolidacao_Cliente_Divisao_Consolida.csv`
- Colunas: `Cliente | Divisão | Consolida`

Tratamentos:

- Cliente repetido com múltiplas divisões: todas as combinações são mantidas, com log de ambiguidade.
- Divisão vazia: mantida; `Consolida` fica vazio.
- Encoding robusto com correção de mojibake (`DivisÃ£o` → `Divisão`).

---

## Importação de ZFIT009 e ZCO059

Opção 3 permite:

- **Via SAP GUI Scripting** (SE16/ZCO059 portados do VBS para `win32com`).
- **Via arquivo manual** já exportado.

As tabelas são validadas por:

- existência de arquivo;
- cabeçalho esperado;
- quantidade de linhas de dados > 0.

Destino padronizado:

- `saidas/tabelas/ZFIT009.csv`
- `saidas/tabelas/ZCO059.csv`

---

## Status Consolidação na movimentação

A opção 5 complementa `Classificado_<periodo>.csv` com:

- `Divisão`
- `Consolida`
- `Status Consolidação`

Lógica:

- Cliente encontrado e `Consolida == "S"` → **Elimina**
- Cliente não encontrado → **Não Consolida**
- `Consolida` diferente de `"S"` (vazio/N/outros) → **Não Consolida**

---

## Contrapartidas (auditoria)

A opção 6 move, para auditoria, linhas com `Conta`:

- `4401020004` (Juros)
- `1103050050` (IRRF)
- `2104030007` (IOF)

Arquivos:

- Principal (classificado) permanece sem essas linhas.
- Auditoria: `Auditoria_Contrapartidas_<periodo>.csv`

Também gera:

- `Classificado_<periodo>.xlsx` com formatação corporativa:
  - cabeçalho em negrito;
  - freeze da primeira linha;
  - autofiltro;
  - largura automática;
  - formato numérico contábil para `Montante Razão`;
  - destaque visual para totais/subtotais.

---

## Resumo de fluxo (Saldo Inicial → Saldo Final)

Opção 8 gera:

- `Resumo_<periodo>.csv`
- `Resumo_<periodo>.xlsx`

Estrutura fixa:

1. Saldo Inicial  
2. Adições  
3. (-) Eliminações  
4. Juros  
5. IOF  
6. IRRF  
7. Baixas  
8. Saldo Final

Fórmula implementada (calibrável):

```
Saldo Final = Saldo Inicial + Adições − Eliminações + Juros + IOF − IRRF − Baixas
```

Observação: **Eliminações** é subtraído como subconjunto para evitar duplicidade.

---

## Saldo Inicial persistente

Arquivo:

- `saldo_inicial.json`

Campos persistidos:

- valor;
- período/trimestre;
- data da última alteração;
- usuário (`getpass.getuser()`).

O script recupera automaticamente o último valor como padrão e mantém histórico simples (append).

---

## Saídas e logs

- `saidas/diarios/`
- `saidas/tabelas/`
- `saidas/consolidado/`
- `logs/`

Todos os CSVs gerados saem em `utf-8-sig`.

---

## Requisitos

- Windows + SAP GUI (para automação SAP).
- Python 3.x
- Dependências em `requirements.txt`:
  - `pywin32`
  - `openpyxl`
