# MUTDFC — Movimentação do Fluxo de Caixa (Mútuo)

Automação em **VBScript (SAP GUI Scripting)** para extrair o Razão Contábil
(transação **FBL3N**) referente a **Mútuo** e consolidar as extrações diárias
em um único arquivo.

> Projeto em construção. O script principal, a documentação de uso e o
> `.gitignore` serão adicionados na sequência.

## Visão geral (previsto)

- Menu interativo perguntando:
  - **Período** (data inicial e final, formato `dd.mm.aaaa`).
  - **Periodicidade** (Diária — principal —, Semanal ou Mensal).
  - **Sistema/Conexão**: `1` = SAP S/4 HANA (PRD) ou `2` = SAP ECC 6.0 (PRD).
- Conexão à sessão SAP já aberta (sem novo login).
- Extração da FBL3N para a conta `MUTDFC` / usuário `MS0000240`, gerando um
  CSV por dia.
- Consolidação de todos os CSVs em um único arquivo, com cabeçalho padronizado.
- LOG detalhado com timestamps para diagnóstico de falhas.
