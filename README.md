# MUTDFC — Movimentação de Mútuo (FBL3N + Consolidação Intercompany)

Automação em Python para:

1. Extrair razão contábil (FBL3N) via SAP GUI Scripting.
2. Consolidar e classificar movimentações.
3. Importar tabelas intercompany (`ZFIT009` e `ZCO059`).
4. Cruzar **Cliente → Destino** via regras por Conta e ajustes manuais.
5. Aplicar **Status Consolidação** (`S` / `N`).
6. Separar contrapartidas para auditoria.
7. Gerar classificado e resumo em CSV + Excel.
8. Persistir **Saldo Inicial** com histórico.
9. Gerenciar **Ajustes Manuais** (de-para Cliente → Divisão).

---

## Menu principal

```
1 - Extrair razão (FBL3N) + Consolidar
2 - Classificar consolidado
3 - Importar ZFIT009 / ZCO059 + Balancete (F.01)
4 - Gerar tabela de consolidação
5 - Aplicar Status Consolidação na movimentação
6 - Tratar contrapartidas (auditoria) + Excel classificado
7 - Informar/atualizar Saldo Inicial
8 - Resumo (Saldo Inicial → Saldo Final)
9 - Executar tudo em sequência
A - Ajustes Manuais (de-para Cliente → Divisão)
0 - Sair
```

---

## Ajustes Manuais (de-para Cliente → Divisão)

A opção **A** permite gerenciar um de-para manual **Cliente → Divisão** que tem
**prioridade** sobre a ZFIT009 na resolução do campo Destino.

- Persistido em `ajustes_manuais.json` na pasta do projeto.
- Estrutura: `{ "vinculos": { "2200000404": "E024", ... }, "ultima_alteracao": ..., "usuario": ... }`.
- Se o arquivo não existir, é criado automaticamente com os 5 vínculos iniciais:

| Cliente    | Divisão |
|------------|---------|
| 2200000404 | E024    |
| 2200010827 | E662    |
| 2200000193 | MAGI    |
| 1800000063 | P001    |
| 1800000060 | B001    |

Submenu: **A** Adicionar · **E** Editar · **R** Remover · **V** Voltar.

---

## Coluna "Destino" — Precedência de regras

A coluna acrescentada à movimentação classificada chama-se **Destino**
(renomeada a partir de "Divisão"). Ela é preenchida na seguinte ordem:

| Prioridade | Regra | Resultado |
|-----------|-------|-----------|
| 1 (maior) | Conta == `1202060002` | `MRL` |
| 1 | Conta == `1202060016` | `Prime` |
| 1 | Conta == `1202060104` | `AHS` |
| 1 | Conta == `2104030007` | `IOF` |
| 1 | Conta == `4401020004` | `Juros` |
| 1 | Conta == `1103050050` | `IRRF` |
| 2 | Cliente em `ajustes_manuais.json` | divisão do de-para manual |
| 3 | Conta começa com `120206` AND ≠ `1202060000` AND não está acima | `Parceiro` |
| 4 | Cliente na ZFIT009 (via tabela de consolidação) | divisão da ZFIT009 |
| 5 (menor) | Nenhuma regra casou | vazio → arquivo `SemDestino_*` gerado |

### Arquivo SemDestino

Linhas que não obtiverem Destino após todas as regras geram o arquivo
`SemDestino_<datainicial>_a_<datafinal>.txt` (UTF-8-SIG) em `saidas/consolidado/`
com as colunas estratégicas: `Nº doc. | Conta | CL | Cliente | Montante Razão | Texto`.
O processamento **não é interrompido**; o TXT é gerado para investigação.

O LOG registra a contagem de Destinos por regra:
`por Conta específica | Ajuste manual | Parceiro | ZFIT009 | sem_destino`.

---

## Classificação "N/A"

Onde a coluna `Classificação` ficar vazia (linhas que não casaram em nenhuma
regra de classificação), ela é preenchida com **`N/A`** — nunca fica em branco.
Linhas com `N/A` são ignoradas nos totais do Resumo.

---

## Consolida e Estrutura de Consolidação

| Situação | Consolida |
|----------|-----------|
| Destino com Consolida=S na ZCO059 | `S` |
| Qualquer outro caso | `N` |

> **Consolida por Destino (v3):** A coluna **Consolida** é determinada buscando
> o **Destino** na coluna **Divisão** da tabela de consolidação, não mais pelo Cliente.
> A coluna **Status Consolidação** foi removida; apenas **Consolida** permanece.

### Estrutura de Consolidação

Nova coluna que indica se a linha pertence a uma entidade **Individual** ou
**Controladas**. Determinada via:
1. **Div** (coluna do lançamento) → ZCO059 (coluna Divisão)
2. ZCO059 → coluna Descrição tratada: "Individual" → Individual; qualquer outro → Controladas
3. Se a divisão da linha (Div) não existir na ZCO059, deixa em branco.

> **Importante**: A coluna "Consolida" continua sendo resolvida por **Destino**; apenas
> a "Estrutura de Consolidação" usa a coluna **Div** (divisão real do lançamento).

---

## Colunas removidas da saída final

As seguintes colunas são **removidas** do CSV e Excel classificados finais
(gerados na opção 6), preservando as demais:

| Coluna removida |
|-----------------|
| `St` |
| `Atribuição` |
| `Imobilizado` |
| `DiagRede` |
| `Status Consolidação` |

As colunas são necessárias internamente durante o processamento; são removidas
apenas na escrita final pela função `separar_contrapartidas()`.

---

## Consolidação intercompany (regra de cruzamento em 2 etapas)

1. **ZFIT009**: busca a **Divisão** a partir do **Cliente**.
2. **ZCO059**: busca o **Consolida** (`S`/vazio) e **Descrição** (Individual/Controladas) a partir da **Divisão**.

Saída:

- `saidas/consolidado/Consolidacao_Cliente_Divisao_Consolida.csv`
- Colunas: `Cliente | Divisão | Descrição | Consolida`

Tratamentos:

- Cliente repetido com múltiplas divisões: todas as combinações são mantidas, com log de ambiguidade.
- Divisão vazia: mantida; `Consolida` fica vazio, `Descrição` fica "Controladas".
- Encoding robusto com correção de mojibake (`DivisÃ£o` → `Divisão`).

---

## Importação de ZFIT009 e ZCO059

Opção 3 permite:

- **Via SAP GUI Scripting** — fluxo simplificado que exporta TODAS as colunas diretamente.
- **Via arquivo manual** já exportado (padrão sugerido quando a automação SAP
  não estiver disponível ou o usuário preferir exportar manualmente via ZCO059Total.vbs).

As tabelas são validadas por:

- existência de arquivo;
- cabeçalho esperado (detecção automática do formato);
- quantidade de linhas de dados > 0;
- colunas selecionadas por POSIÇÃO na ZCO059 (Divisão=4ª col, Descrição-estrutura=antes de Consolida);
- log de validação: nº de linhas, contagem Individual/Controladas.

Destino padronizado (derivado automaticamente de `os.path.abspath(__file__)`):

- `saidas/tabelas/ZFIT009.csv`
- `saidas/tabelas/ZCO059.csv` (formato: `Divisão|Descrição|Consolida`)

### Fluxo SAP — ZFIT009 (alinhado ao VBS gravado)

1. Maximizar janela (`wnd[0].maximize`).
2. Navegar para SE16 via campo de comando (`okcd` + `sendVKey 0`).
3. Inserir nome da tabela `ZFIT009` e confirmar.
4. Abrir configuração de campos via **`menu[3]/menu[0]/menu[1]`** (não por sendVKey direto).
5. `sendVKey 14` para acessar seleção de campos.
6. Marcar `chk[1,5]` e `chk[1,11]`; usar **`sendVKey 6`** para aplicar (não `sendVKey 8`).
7. **`sendVKey 8`** (F8) para executar a lista; aguardar 2 s.
8. **`sendVKey 20`** para abrir menu de download/exportação.
9. `sendVKey 0` para confirmar o tipo de arquivo.
10. Preencher `DY_PATH` (pasta `saidas/tabelas` do projeto) e `DY_FILENAME = ZFIT009.csv`.
11. `btn[11]` (Salvar/Substituir).
12. 3× `sendVKey 3` para fechar telas e retornar ao menu.

### Fluxo SAP — ZCO059 (simplificado — ZCO059Total.vbs)

1. Navegar para `zco059` via campo de comando.
2. Aguardar ALV grid em `subSUB_DRE:ZCOR043:0100/cntlCCONTAINER_1/shellcont/shell`.
3. `pressToolbarContextButton "&MB_EXPORT"` + `selectContextMenuItem "&PC"`.
4. `btn[0]` para confirmar tipo de exportação.
5. Preencher `DY_PATH` e `DY_FILENAME = ZCO059GERAL.csv`; btn[11] para salvar.
6. **Processamento em Python**: o arquivo completo é parseado e as colunas são
   selecionadas por POSIÇÃO (há DUAS colunas "Descrição"; a correta é a que fica
   imediatamente antes de "Consolida").
7. Filtragem `Consolida = "S"` feita em Python.
8. Mapeamento de Descrição-estrutura: "Individual" → Individual; outro → Controladas.

> **Nota**: A seleção de colunas via `&COL0` foi removida para eliminar erros.
> O Python faz toda a filtragem e seleção de colunas por posição.

### Robustez contra o erro 619

O helper `esperar_controle(session, control_id, timeout, intervalo, logger)` chama
`session.findById(control_id, False)` em loop com pausas até o controle aparecer,
evitando o erro 619 *"The control could not be found by id"*. É chamado antes de
cada interação com controles que dependem de carregamento assíncrono da tela SAP.

### Importar de arquivo já exportado

Ao escolher a opção **2 - Arquivo já exportado manualmente** no submenu de importação,
o sistema lê o CSV com o parser de largura fixa (`|`) incluindo:

- correção de mojibake (`Divis\xef\xbf\xbd\xef\xbf\xbdo` → `Divisão`);
- ignora linhas de título, separadores e rodapé SAP;
- aceita três formatos:
  - **ZCO059GERAL** (arquivo completo com TODAS as colunas) — colunas selecionadas por POSIÇÃO;
  - **Normalizado** (`Divisão|Descrição|Consolida`) — formato de saída do script;
  - **Legado** (`Empresa|Divisão|Descrição|Consolida`) — retrocompatível;
- valida `Divisão` (alfanumérica curta) e `Consolida` (`S`, `N` ou vazio);
- filtra apenas linhas com `Consolida = "S"`;
- log detalhado: nº de linhas processadas, registros carregados, contagem Individual/Controladas.

---

## Consolida na movimentação

A opção 5 complementa `Classificado_<periodo>.csv` com:

- `Destino` (regra de precedência descrita acima)
- `Consolida` (buscado via **Destino** na tabela de consolidação)
- `Estrutura de Consolidação` (Individual/Controladas via ZCO059)

Lógica do Consolida:

- **Destino** encontrado com `Consolida == "S"` na ZCO059 → **S** (consolida/elimina)
- Qualquer outro caso → **N** (não consolida)

> **Mudança (v3):** o Consolida é determinado pelo **Destino** (não mais pelo Cliente).
> A coluna "Status Consolidação" foi removida; apenas "Consolida" permanece.

---

## Contrapartidas (auditoria)

A opção 6 gera uma **cópia de auditoria** das linhas com `Conta`:

- `4401020004` (Juros)
- `1103050050` (IRRF)
- `2104030007` (IOF)

Arquivos:

- **Principal (classificado) MANTÉM essas linhas** (IOF/IRRF/Juros permanecem
  no Classificado — participam da reconciliação com o Balancete, issue #25 item 6).
- Auditoria: `Auditoria_Contrapartidas_<periodo>.csv` (cópia dessas linhas como evidência).

Também gera:

- `Classificado_<periodo>.xlsx` com formatação corporativa:
  - cabeçalho em negrito;
  - freeze da primeira linha;
  - autofiltro;
  - largura automática;
  - formato numérico contábil para `Montante Razão`;
  - destaque visual para totais/subtotais.

---

## Balancete (transação SAP F.01)

A opção **3** também importa e concilia o **Balancete** (transação SAP `F.01`,
variante `BPMUTDFC`, usuário `MS0000240`) contra o Razão (Classificado),
gerando evidência de auditoria (issue #25).

### Parâmetros dinâmicos por período

Derivados automaticamente da data final do período (nome do `Classificado_*`).
O comparativo mede a movimentação de **um mês** (mês final − mês imediatamente
anterior), reconciliando com a Movimentação do Razão do mesmo mês:

| Bloco | Campo SAP | Valor |
|-------|-----------|-------|
| Atual | `txtBILBJAHR` | ANO DO MÊS FINAL |
| Atual | `txtB-MONATE-LOW` / `txtB-MONATE-HIGH` | MÊS FINAL |
| Comparativo | `txtBILVJAHR` | ANO DO MÊS ANTERIOR |
| Comparativo | `txtV-MONATE-LOW` / `txtV-MONATE-HIGH` | MÊS IMEDIATAMENTE ANTERIOR |

> Regra geral: mês final `M/AAAA` → comparativo `(M-1)/AAAA`.
> Borda janeiro: mês final `01/AAAA` → comparativo `12/(AAAA-1)`.
>
> Ex.: período 01/03/2026–31/03/2026 → Atual 2026 / 03-03; Comparativo 2026 / 02-02.
> Ex.: período 01/04/2026–30/06/2026 → Atual 2026 / 06-06; Comparativo 2026 / 05-05.
> Ex.: período 01/01/2026–31/01/2026 → Atual 2026 / 01-01; Comparativo 2025 / 12-12.

O arquivo é exportado para `saidas/tabelas/F.01.csv`.

### Parser F.01

- Ignora cabeçalhos repetidos por empresa/página e linhas de subtotal marcadas
  na última coluna (`Rela`) com `*1*`, `*2*`, `*3*`, `*4*`.
- Considera apenas linhas de detalhe (`Rela` vazio, `Empr`/`Div` preenchidos e
  número de Conta no início do campo `Textos`).
- Trata valores no formato BR com sinal `-` à direita (ex.: `10.229,02-`).
- Extrai a evidência do topo do relatório: data e hora de emissão, usuário
  executor, empresa, exercício e períodos consultados.

### Aba "Balancete" no `.xlsx`

O Balancete conciliado é gerado como **arquivo separado** (`Balancete_<periodo>.xlsx`/`.csv`)
**e também** como uma aba **"Balancete"** embutida no `Resumo_<periodo>.xlsx` (ver
"Resumo de fluxo"). Ambos têm o mesmo conteúdo:

- Bloco de **evidência** no topo (data/hora/usuário/empresa/exercício/período).
- **Conferência global**: Total Variação no Período, Total Movimentação Razão,
  Diferença e Status (`OK`/`DIVERGÊNCIA`), também registrados no LOG.
- Tabela com colunas renomeadas (item 8):

| Coluna | Origem |
|--------|--------|
| Saldo Período Atual | Períodos de relató |
| Saldo Período Comparativo | Períodos de compar |
| Variação no Período | Desvio absoluto |
| Movimentação Razão | Σ `Montante Razão` do Classificado por Empr+Div+Conta |
| Diferença de Conciliação | Variação no Período − Movimentação Razão |
| Conferência | `OK` se \|Diferença\| ≤ 0,01; senão `DIVERGÊNCIA` |

### Conciliação conta a conta

Chave = `Empr` + `Div` + `Conta`. A **Movimentação Razão** é a soma de
`Montante Razão` do Classificado agrupada por essa chave. Também é gerado o
`Balancete_<periodo>.csv` (utf-8-sig) como evidência.

---

## Resumo de fluxo (Saldo Inicial → Saldo Final)

Opção 8 gera:

- `Resumo_<periodo>.csv`
- `Resumo_<periodo>.xlsx` — com as abas **Resumo**, **Conciliação**,
  **Classificado** e **Balancete**. A aba **Balancete** é incluída quando a F.01
  (`saidas/tabelas/F.01.csv`) estiver disponível, com o mesmo conteúdo do arquivo
  `Balancete_<periodo>.xlsx` separado.

### Estrutura (linhas de planilha; L1 = título, L2 = cabeçalho)

O resumo é segregado em três colunas de valor. A estrutura é fixa por linha:

| Linha | Etapa | Individual | Controladas | Consolidado |
|-------|-------|------------|-------------|-------------|
| L3 | Saldo Inicial | valor único | 0 | valor único |
| L4 | Adições | SOMASES por Estrutura | SOMASES por Estrutura | total − eliminações |
| L5 | Juros | ... | ... | total − eliminações |
| L6 | IOF | ... | ... | total − eliminações |
| L7 | IRRF | ... | ... | total − eliminações |
| L8 | Baixas | ... | ... | total − eliminações |
| L9 | Saldo Final = **soma L3:L8** | ... | ... | ... |

> **Mudança conceitual (v2):** **não há** linhas de "(-) Eliminações". As
> eliminações ficam **embutidas na coluna Consolidado** (efeito líquido em cada linha).

Cálculo por linha de movimento (L4:L8), tendo a coluna A (Descrição) como chave
contra a coluna `Classificação` do Classificado:

- **Individual**  = `SOMASES(Montante ; Classificação=label ; Estrutura="Individual")`
- **Controladas** = `SOMASES(Montante ; Classificação=label ; Estrutura="Controladas")`
- **Consolidado** = `SOMASES(Montante ; Classificação=label)` − `SOMASES(Montante ; Classificação=label ; Consolida="S")`

> **Importante:** o **Consolidado é INDEPENDENTE** (total menos eliminações
> `Consolida="S"`); **NÃO** é Individual + Controladas. Nas linhas sem eliminação os
> números coincidem; nas com eliminação, o Consolidado fica menor (correto).

- **Estrutura de Consolidação**: determinada via **Div** (divisão real do lançamento) → ZCO059 (Divisão) → Descrição (Individual/Controladas).

### Saldo Inicial (input único)

O **Saldo Inicial** (opção 7) é um **input único**. O valor informado vai para
**Individual** E **Consolidado**; **Controladas = 0**. Não é solicitado valor de
Controladas. O valor é persistido em `saldo_inicial.json`.

### Saldo Final (regra anti-loop)

Para cada coluna (Individual, Controladas, Consolidado), o **Saldo Final (L9)** é a
**soma algébrica literal das células L3:L8** dessa mesma coluna (respeitando o sinal
exibido). **Não** é usada fórmula independente nem recomputação a partir do razão.

```
Saldo Final (L9) = SOMA(L3:L8)   — por coluna
```

- No `.xlsx`, a célula do Saldo Final usa **fórmula real** `=SUM(<L3>:<L8>)` por coluna.
- **Validação que aborta**: após montar o resumo, recomputa em Python a soma de L3:L8
  por coluna e compara com o Saldo Final; se divergir em qualquer coluna (tolerância
  0,5 milhar), a geração é **abortada com ERROR**. O LOG registra as 3 somas L3:L8 e
  os 3 Saldos Finais.

### Aba "Conciliação"

O mesmo `.xlsx` inclui uma aba **"Conciliação"** que reconcilia a base classificada,
o resumo e as eliminações, com as linhas:

- Total da base classificada (soma de `Montante Razão` das linhas de movimento);
- Total considerado no resumo (soma das linhas L4:L8);
- Total eliminado no consolidado (soma das linhas `Consolida="S"`);
- Diferença esperada;
- Diferença apurada;
- **Status**: `OK` ou `Divergência`.

O status é registrado no LOG.

### Detalhe: aba "Classificado" (conteúdo integral)

A antiga aba "Composição Analítica" foi **removida**. Em seu lugar, o mesmo `.xlsx` do
Resumo inclui uma aba **"Classificado"** com o **conteúdo integral** do arquivo
`Classificado_<periodo>` (razão final classificado — todas as linhas e colunas,
incluindo `Classificação`, `Destino`, `Consolida` e `Estrutura de Consolidação`), sem
resumir ou agrupar.

### Aba "Balancete"

O mesmo `.xlsx` do Resumo inclui uma aba **"Balancete"** com a conciliação conta a
conta contra a F.01 (bloco de evidência, conferência global e a tabela
`Empr|Div|Conta|Descrição|Saldo Período Atual|Saldo Período Comparativo|Variação no
Período|Movimentação Razão|Diferença de Conciliação|Conferência`), com a mesma
formatação do arquivo `Balancete_<periodo>.xlsx`. A aba é incluída quando a F.01
(`saidas/tabelas/F.01.csv`) estiver presente; caso contrário, o Resumo é gerado com
as três abas restantes. O LOG registra a conciliação da aba (Total Variação | Total
Razão | Diferença | Status).

### Formatação corporativa do Resumo .xlsx

- Cabeçalho verde com texto branco em negrito
- Saldo Inicial e Saldo Final em negrito com destaque verde claro
- IRRF e Baixas em vermelho
- Valores em **milhares** (÷1000), sem decimais, separador de milhar BR; zero → `-`; negativos em vermelho/parênteses
- Faixa laranja de separação antes do Saldo Final

---

## Saldo Inicial persistente

Arquivo:

- `saldo_inicial.json`

Campos persistidos:

- valor (input único: Individual = Consolidado; Controladas = 0);
- período/trimestre;
- data da última alteração;
- usuário (`getpass.getuser()`).

O **Saldo Inicial** (opção 7) é um **input único**: o valor informado vai para
Individual E Consolidado; Controladas fica em 0.

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
