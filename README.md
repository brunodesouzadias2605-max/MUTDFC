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
3 - Importar ZFIT009 / ZCO059
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
| 2 | Conta começa com `120206` AND ≠ `1202060000` AND não está acima | `Parceiro` |
| 3 | Cliente em `ajustes_manuais.json` | divisão do de-para manual |
| 4 | Cliente na ZFIT009 (via tabela de consolidação) | divisão da ZFIT009 |
| 5 (menor) | Nenhuma regra casou | vazio → arquivo `SemDestino_*` gerado |

### Arquivo SemDestino

Linhas que não obtiverem Destino após todas as regras geram o arquivo
`SemDestino_<datainicial>_a_<datafinal>.txt` (UTF-8-SIG) em `saidas/consolidado/`
com as colunas estratégicas: `Nº doc. | Conta | CL | Cliente | Montante Razão | Texto`.
O processamento **não é interrompido**; o TXT é gerado para investigação.

O LOG registra a contagem de Destinos por regra:
`por Conta específica | Parceiro | ajuste_manual | ZFIT009 | sem_destino`.

---

## Classificação "N/A"

Onde a coluna `Classificação` ficar vazia (linhas que não casaram em nenhuma
regra de classificação), ela é preenchida com **`N/A`** — nunca fica em branco.
Linhas com `N/A` são ignoradas nos totais do Resumo.

---

## Consolida e Status Consolidação

| Situação | Consolida | Status Consolidação |
|----------|-----------|---------------------|
| Cliente com Consolida=S na ZCO059 | `S` | `S` |
| Qualquer outro caso | `N` | `N` |

> **Mudança de rótulo (v2):** os valores `Elimina` e `Não Consolida` foram
> substituídos por `S` e `N` respectivamente. O Resumo de Fluxo continua
> somando as linhas marcadas como `S` como "Eliminações" — apenas o rótulo mudou.

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

As colunas são necessárias internamente durante o processamento; são removidas
apenas na escrita final pela função `separar_contrapartidas()`.

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

- **Via SAP GUI Scripting** — fluxo portado diretamente dos VBS gravados na
  sessão atual (IDs de controle alinhados aos gravados; não usar com VBS antigos).
- **Via arquivo manual** já exportado (padrão sugerido quando a automação SAP
  não estiver disponível ou o usuário preferir exportar manualmente).

As tabelas são validadas por:

- existência de arquivo;
- cabeçalho esperado;
- quantidade de linhas de dados > 0.
- colunas obrigatórias da ZCO059 (`Empresa | Divisão | Consolida`) e rejeição do
  export incorreto com descrições como `SPE Controlada | S | SIM`.

Destino padronizado (derivado automaticamente de `os.path.abspath(__file__)`):

- `saidas/tabelas/ZFIT009.csv`
- `saidas/tabelas/ZCO059.csv`

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

### Fluxo SAP — ZCO059 (alinhado ao VBS gravado)

1. Navegar para `zco059` via campo de comando.
2. Aguardar ALV grid em `subSUB_DRE:ZCOR043:0100/cntlCCONTAINER_1/shellcont/shell`.
3. `pressToolbarButton "&COL0"` para abrir seleção de colunas.
4. A seleção de colunas do ALV é **obrigatória** e replica o VBS com esperas entre
   cada passo do popup (`btnAPP_FL_SING` + `currentCellRow = 1`, `2`, sete presses
   adicionais, `currentCellRow = 3` e `btn[0]` para confirmar).
5. Se a seleção falhar após 3 tentativas, a importação é abortada com erro claro
   orientando a exportar via `ZCO059.vbs` e usar a opção de importar por arquivo.
6. `setCurrentCell(-1, "CONSOLIDA")` + `selectColumn("CONSOLIDA")`.
7. `pressToolbarContextButton "&MB_FILTER"` + `selectContextMenuItem "&FILTER"`.
8. Preencher `%%DYN001-LOW = "S"` e confirmar com `btn[0]`.
9. `pressToolbarContextButton "&MB_EXPORT"` + `selectContextMenuItem "&PC"`.
10. `btn[0]` para confirmar tipo de exportação.
11. Preencher `DY_PATH = <pasta do projeto>/saidas/tabelas`; o arquivo é validado
    antes de ser renomeado para `ZCO059.csv`.

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
- aceita tanto o formato largura-fixa do VBS quanto o CSV simples normalizado;
- valida `Divisão` (alfanumérica curta) e `Consolida` (`S`, `N` ou vazio);
- rejeita explicitamente o arquivo incorreto sem seleção de colunas do ALV,
  exibindo erro claro para refazer a exportação.

---

## Status Consolidação na movimentação

A opção 5 complementa `Classificado_<periodo>.csv` com:

- `Destino` (regra de precedência descrita acima)
- `Consolida`
- `Status Consolidação`

Lógica:

- Cliente encontrado e `Consolida == "S"` → **S** (consolida/elimina)
- Qualquer outro caso → **N** (não consolida)

> Antes da versão 2 os valores eram "Elimina" e "Não Consolida".

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
