"""
mutdfc.py — MUTDFC: Movimentação do Fluxo de Caixa (contas de Mútuo)

Extrai o Razão Contábil (FBL3N) referente a Mútuo via SAP GUI Scripting
(COM / pywin32), consolida os arquivos diários em um único CSV (limpo e
ordenado), classifica as linhas de Mútuo, cruza consolidação intercompany
(ZFIT009/ZCO059), separa contrapartidas e gera resumo em estrutura de fluxo.

Pré-requisitos:
  - Windows com SAP GUI instalado e SAP GUI Scripting habilitado.
  - Usuário já logado no sistema SAP desejado (o script apenas anexa à sessão).
  - Python 3.x + pywin32  →  pip install -r requirements.txt

Como alterar configurações:
  Ajuste as CONSTANTES abaixo (CONTA_LAYOUT, USUARIO_SAP, SISTEMAS) conforme
  o seu ambiente antes de executar. As pastas de saída são criadas
  automaticamente dentro do próprio diretório do script.
"""

import calendar
import glob
import json
import logging
import os
import re
import sys
from collections import defaultdict
from datetime import datetime, timedelta
from getpass import getuser

# ---------------------------------------------------------------------------
# CONSTANTES DE CONFIGURAÇÃO — altere aqui conforme o seu ambiente
# ---------------------------------------------------------------------------

# Pasta base do projeto: sempre relativa à localização deste script.
# Para sobrescrever, substitua a linha abaixo por um caminho absoluto, ex.:
#   PASTA_BASE = r"C:\Minha\Pasta\Personalizada"
PASTA_BASE = os.path.dirname(os.path.abspath(__file__))

# Subpastas de saída (criadas automaticamente)
PASTA_DIARIOS     = os.path.join(PASTA_BASE, "saidas", "diarios")
PASTA_CONSOLIDADO = os.path.join(PASTA_BASE, "saidas", "consolidado")
PASTA_TABELAS     = os.path.join(PASTA_BASE, "saidas", "tabelas")
PASTA_LOGS        = os.path.join(PASTA_BASE, "logs")
ARQUIVO_SALDO_INICIAL = os.path.join(PASTA_BASE, "saldo_inicial.json")

# Variante de layout salva no SAP (campo txtV-LOW da tela de seleção de layout)
CONTA_LAYOUT = "MUTDFC"

# Usuário SAP responsável pela variante (campo txtENAME-LOW)
USUARIO_SAP = "MS0000240"

# Definições dos sistemas disponíveis no menu
# Chave: número exibido no menu (str)
# "descricao_conexao": string usada pelo SAP GUI para identificar a conexão
# "palavras_chave": substrings procuradas na descrição das conexões abertas
SISTEMAS = {
    "1": {
        "nome": "SAP S/4 HANA (PRD)",
        "descricao_conexao": "MRV SAP S/4 HANA PRODUCAO",
        "palavras_chave": ["S/4 HANA PRODUCAO"],
    },
    "2": {
        "nome": "SAP ECC 6.0 (PRD)",
        "descricao_conexao": "03 SAP PRD ECC - BOLHA",
        "palavras_chave": ["ECC", "BOLHA"],
    },
}

# Cabeçalho exato esperado nos CSVs exportados pelo SAP e no consolidado.
CABECALHO_CONSOLIDADO = (
    "|   St|Atribuição        |Nº doc.   |Dt.Lçto   |Div |Conta     "
    "|Fornecedor|Tp.doc. |Cliente   |Data doc. |CL|      Montante Razão"
    "|DocCompens|Texto                                             "
    "|Imobilizado |Usuário     |Empr|Referência      |Entrado em|Estorno"
    "|DiagRede    |"
)

# Índices dos campos após parse_linha() (0-based dentro de parts[1:-1])
IDX_ST        = 0
IDX_ATRIB     = 1
IDX_NRDOC     = 2
IDX_DTLCTO    = 3
IDX_DIV       = 4
IDX_CONTA     = 5
IDX_FORNEC    = 6
IDX_TPDOC     = 7
IDX_CLIENTE   = 8
IDX_DATADOC   = 9
IDX_CL        = 10
IDX_MONTANTE  = 11
IDX_DOCCOMP   = 12
IDX_TEXTO     = 13
IDX_IMOBIL    = 14
IDX_USUARIO   = 15
IDX_EMPR      = 16
IDX_REFER     = 17
IDX_ENTRADO   = 18
IDX_ESTORNO   = 19
IDX_DIAGREDE  = 20

# Contas específicas usadas na classificação (altere aqui se os números mudarem)
CONTA_IOF   = "2104030007"   # IOF sobre operação de Mútuo
CONTA_IRRF  = "1103050050"   # IRRF sobre rendimento de Mútuo
CONTA_JUROS = "4401020004"   # Receita de juros de Mútuo

# Mapeamento conta → classificação (prioridade: IOF > IRRF > Juros)
CONTAS_CLASSIFICACAO = {
    CONTA_IOF:   "IOF",
    CONTA_IRRF:  "IRRF",
    CONTA_JUROS: "Juros",
}

# Prefixo das contas de Mútuo
PREFIXO_MUTUO = "1202"

# Contas de contrapartida movidas para auditoria após classificação
CONTAS_CONTRAPARTIDA_AUDITORIA = {CONTA_JUROS, CONTA_IRRF, CONTA_IOF}

# Ordem fixa do resumo em estrutura de fluxo
_ORDEM_FLUXO = [
    "Saldo Inicial",
    "Adições",
    "(-) Eliminações",
    "Juros",
    "IOF",
    "IRRF",
    "Baixas",
    "Saldo Final",
]

# Sentinel retornado por extrair_razao() quando o dia não tem partidas.
# Distingue "sem dados" (fluxo normal) de None (falha/erro).
_SEM_DADOS = ""


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO DO LOG
# ---------------------------------------------------------------------------

def configurar_log(pasta: str) -> tuple:
    """Cria e retorna (logger, caminho_log) com handlers de arquivo e console."""
    os.makedirs(pasta, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caminho_log = os.path.join(pasta, f"MUTDFC_log_{timestamp}.txt")

    logger = logging.getLogger("MUTDFC")
    logger.setLevel(logging.DEBUG)

    # Evitar adicionar handlers duplicados se main() for chamada mais de uma vez
    if logger.handlers:
        logger.handlers.clear()

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Handler de arquivo
    fh = logging.FileHandler(caminho_log, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Handler de console
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    logger.info("Log iniciado: %s", caminho_log)
    return logger, caminho_log


# ---------------------------------------------------------------------------
# PARSE DE LINHAS
# ---------------------------------------------------------------------------

def parse_linha(linha: str) -> list:
    """
    Divide a linha pelo delimitador '|' e retorna a lista de campos (sem o
    primeiro e último elementos vazios produzidos pela divisão de uma linha
    que começa e termina com '|').
    """
    partes = linha.split("|")
    return partes[1:-1] if len(partes) >= 2 else partes


def tipo_de_linha(linha: str) -> str:
    """
    Classifica o tipo de uma linha do CSV exportado pelo SAP.

    Retorna uma das strings:
      'dado'      — linha de lançamento real, deve ser mantida
      'cabecalho' — linha com os títulos das colunas
      'separador' — linha de hífens entre cabeçalho e dados
      'total'     — linha de subtotal/total (campo St contém '*')
      'vazio'     — linha em branco ou sem conteúdo útil
    """
    if not linha.strip():
        return "vazio"

    # Linhas que não começam com '|': separadores puros de hífens ou lixo
    if not linha.startswith("|"):
        if linha.lstrip("-").strip() == "":
            return "separador"
        return "vazio"

    # Verificar separador: células com apenas hífens representam ≥ 50 % do total
    campos = parse_linha(linha)
    if not campos:
        return "vazio"

    celulas_dash = sum(
        1 for c in campos if c.strip() and not c.strip().replace("-", "").strip()
    )
    celulas_com_conteudo = sum(1 for c in campos if c.strip())
    if celulas_com_conteudo > 0 and celulas_dash / celulas_com_conteudo >= 0.5:
        return "separador"

    # Verificar cabeçalho pelos termos exclusivos dos títulos de coluna
    termos_cabecalho = ("Montante", "Atribuição", "Atribuicao", "Nº doc", "Dt.Lçto")
    if any(t in linha for t in termos_cabecalho):
        return "cabecalho"

    # Verificar total: campo St (índice 0) contém '*'
    if campos and "*" in campos[IDX_ST]:
        return "total"

    # Verificar linha vazia de conteúdo (só pipes)
    if not linha.replace("|", "").strip():
        return "vazio"

    return "dado"


# ---------------------------------------------------------------------------
# LIMPEZA E ORDENAÇÃO
# ---------------------------------------------------------------------------

def limpar_e_ordenar(linhas: list, logger: logging.Logger) -> list:
    """
    Filtra as linhas mantendo apenas as do tipo 'dado' e ordena
    cronologicamente por Dt.Lçto (formato dd.mm.aaaa).

    Linhas com data inválida ou ausente são movidas para o final e logadas
    como aviso.
    """
    dados_validos = []
    dados_invalidos = []

    for linha in linhas:
        if tipo_de_linha(linha) != "dado":
            continue
        campos = parse_linha(linha)
        data_str = campos[IDX_DTLCTO].strip() if len(campos) > IDX_DTLCTO else ""
        try:
            datetime.strptime(data_str, "%d.%m.%Y")
            dados_validos.append(linha)
        except ValueError:
            logger.warning(
                "Linha com Dt.Lçto inválida ('%s') movida para o fim: %s",
                data_str,
                linha[:80],
            )
            dados_invalidos.append(linha)

    def _chave_data(linha: str) -> datetime:
        campos = parse_linha(linha)
        try:
            return datetime.strptime(campos[IDX_DTLCTO].strip(), "%d.%m.%Y")
        except (ValueError, IndexError):
            return datetime.max

    dados_validos.sort(key=_chave_data)
    logger.debug(
        "limpar_e_ordenar: %d válidas, %d com data inválida.",
        len(dados_validos),
        len(dados_invalidos),
    )
    return dados_validos + dados_invalidos


# ---------------------------------------------------------------------------
# UTILITÁRIOS NUMÉRICOS
# ---------------------------------------------------------------------------

def _parse_montante(valor_str: str):
    """
    Converte string de montante no formato brasileiro para float.

    Formatos suportados:
      '1.234,56'    → positivo
      '-8.000,00'   → negativo (sinal à esquerda)
      '3.301,24-'   → negativo (sinal à direita, padrão SAP FBL3N)

    Retorna None se não for possível converter.
    """
    if not valor_str:
        return None
    try:
        s = valor_str.strip()
        # Sinal negativo à direita (ex.: '3.301,24-') — padrão SAP
        negativo = s.endswith("-")
        if negativo:
            s = s[:-1].strip()
        # Remove separador de milhar (.) e substitui decimal (,) por (.)
        limpo = s.replace(".", "").replace(",", ".")
        valor = float(limpo)
        return -valor if negativo else valor
    except (ValueError, AttributeError):
        return None


def _formatar_valor_br(valor: float) -> str:
    """Formata float no padrão brasileiro com negativo à direita."""
    abs_v = abs(valor)
    s = f"{abs_v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s}-" if valor < 0 else s


def _parse_numero_br(valor_str: str):
    """Converte texto numérico BR para float (aceita negativo à esquerda/direita)."""
    if valor_str is None:
        return None
    s = str(valor_str).strip()
    if not s:
        return None
    negativo = s.endswith("-")
    if negativo:
        s = s[:-1].strip()
    if s.startswith("-"):
        negativo = True
        s = s[1:].strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        valor = float(s)
        return -valor if negativo else valor
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# DETECÇÃO DE DIA SEM MOVIMENTO
# ---------------------------------------------------------------------------

def dia_sem_movimento(session, logger: logging.Logger) -> bool:
    """
    Verifica se a barra de status do SAP indica que não há partidas selecionadas.

    Após a execução do F8 na FBL3N, quando não há lançamentos no período,
    o SAP exibe na statusbar (canto inferior esquerdo) a mensagem:
      'Nenhuma partida selecionada (ver texto descritivo)'

    Retorna True se a mensagem indicar ausência de dados, False caso contrário
    ou em caso de erro ao ler a statusbar.
    """
    try:
        sbar = session.findById("wnd[0]/sbar")
        msg = (sbar.Text or "").strip().lower()
        logger.debug("Statusbar após F8: '%s'", msg)
        # Comparação normalizada: tolera variações de maiúsculas e acentos
        return "nenhuma partida" in msg or "no items" in msg
    except Exception as exc:  # noqa: BLE001
        logger.debug("Erro ao ler statusbar: %s", exc)
        return False


# ---------------------------------------------------------------------------
# CONEXÃO COM O SAP
# ---------------------------------------------------------------------------

def conectar_sap(sistema: dict, logger: logging.Logger):
    """
    Localiza e retorna a sessão SAP correspondente ao sistema escolhido.

    O usuário deve estar previamente logado. O script apenas anexa à sessão
    existente; se não encontrar, tenta abrir a conexão.

    Retorna o objeto session do SAP GUI Scripting ou None em caso de falha.
    """
    try:
        import win32com.client  # noqa: PLC0415 — importado aqui para não falhar em Linux
    except ImportError:
        logger.error(
            "Módulo win32com.client não encontrado. "
            "Instale com: pip install pywin32"
        )
        return None

    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        application = sap_gui.GetScriptingEngine
        logger.info("SAP GUI Scripting Engine obtido com sucesso.")
    except Exception as exc:  # noqa: BLE001
        logger.error("Não foi possível obter o SAP GUI Scripting Engine: %s", exc)
        return None

    palavras = sistema["palavras_chave"]
    session = None

    # Tentar localizar entre as conexões já abertas
    try:
        num_conexoes = application.Children.Count
        logger.info("Conexões abertas no SAP GUI: %d", num_conexoes)
        for i in range(num_conexoes):
            conexao = application.Children(i)
            descricao = getattr(conexao, "Description", "") or ""
            logger.debug("Conexão %d: '%s'", i, descricao)
            if any(p.upper() in descricao.upper() for p in palavras):
                session = conexao.Children(0)
                logger.info(
                    "Sessão localizada na conexão %d ('%s').", i, descricao
                )
                break
    except Exception as exc:  # noqa: BLE001
        logger.warning("Erro ao iterar conexões abertas: %s", exc)

    # Se não encontrou, tentar abrir a conexão
    if session is None:
        descricao_conexao = sistema["descricao_conexao"]
        logger.info(
            "Sessão não localizada. Tentando abrir conexão: '%s'",
            descricao_conexao,
        )
        try:
            conexao = application.OpenConnection(descricao_conexao, True)
            session = conexao.Children(0)
            logger.info("Conexão aberta com sucesso: '%s'", descricao_conexao)
        except Exception as exc:  # noqa: BLE001
            logger.error(
                "Falha ao abrir conexão '%s': %s", descricao_conexao, exc
            )
            return None

    return session


# ---------------------------------------------------------------------------
# EXTRAÇÃO FBL3N
# ---------------------------------------------------------------------------

def _fmt_data(data: datetime) -> str:
    """Formata datetime para o padrão SAP dd.mm.aaaa."""
    return data.strftime("%d.%m.%Y")


def extrair_razao(
    session,
    data_low: datetime,
    data_high: datetime,
    pasta: str,
    logger: logging.Logger,
):
    """
    Executa a transação FBL3N para o intervalo [data_low, data_high] e
    exporta o resultado para CSV na pasta informada.

    Retorna:
      str  — caminho do CSV gerado (sucesso com dados)
      _SEM_DADOS ('')  — dia sem partidas (INFO, fluxo normal, sem CSV gerado)
      None — falha/erro (WARNING/ERROR já registrado no log)
    """
    nome_arquivo = f"{_fmt_data(data_low)}.csv"
    caminho_csv = os.path.join(pasta, nome_arquivo)
    str_low = _fmt_data(data_low)
    str_high = _fmt_data(data_high)

    logger.info("Iniciando extração FBL3N: %s → %s", str_low, str_high)

    try:
        # Navegar para FBL3N via campo de comando
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").Text = "fbl3n"
        session.findById("wnd[0]").sendVKey(0)   # Enter
        session.findById("wnd[0]").sendVKey(17)  # Abrir variante de seleção
        logger.debug("FBL3N aberta; abrindo seleção de variante.")
    except Exception as exc:  # noqa: BLE001
        logger.error("Erro ao abrir FBL3N (navegação): %s", exc)
        return None

    try:
        # Preencher variante de layout e usuário
        session.findById("wnd[1]/usr/txtV-LOW").Text = CONTA_LAYOUT
        session.findById("wnd[1]/usr/txtENAME-LOW").Text = USUARIO_SAP
        session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
        session.findById("wnd[1]").sendVKey(0)  # Confirmar
        logger.debug("Variante '%s' / usuário '%s' preenchidos.", CONTA_LAYOUT, USUARIO_SAP)
    except Exception as exc:  # noqa: BLE001
        logger.error("Erro ao preencher variante/usuário: %s", exc)
        return None

    try:
        # Preencher datas e executar (F8)
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").Text = str_low
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").Text = str_high
        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus()
        session.findById("wnd[0]").sendVKey(8)   # F8 — Executar
        logger.debug("Datas preenchidas: LOW=%s HIGH=%s. Executando...", str_low, str_high)
    except Exception as exc:  # noqa: BLE001
        logger.error("Erro ao preencher datas ou executar FBL3N: %s", exc)
        return None

    # -----------------------------------------------------------------------
    # DETECÇÃO DE DIA SEM MOVIMENTO — verificar statusbar antes de exportar
    # Se não houver partidas, o SAP mantém o cursor na tela de seleção e exibe
    # "Nenhuma partida selecionada" na barra de status (statusbar).
    # Nesse caso pulamos a exportação sem gerar ERROR/WARNING.
    # -----------------------------------------------------------------------
    if dia_sem_movimento(session, logger):
        logger.info(
            "Dia %s sem partidas — nenhum dado a exportar.", str_low
        )
        # Fechar FBL3N e retornar ao menu SAP
        try:
            session.findById("wnd[0]").sendVKey(3)  # Voltar
        except Exception:  # noqa: BLE001
            pass
        return _SEM_DADOS  # sentinel: sem dados, sem erro

    # -----------------------------------------------------------------------
    # EXPORTAÇÃO — só chega aqui quando há dados na lista
    # -----------------------------------------------------------------------
    try:
        # Selecionar tudo e abrir diálogo de exportação
        session.findById("wnd[0]").sendVKey(20)  # Ctrl+Shift+F9 — selecionar tudo
        session.findById("wnd[0]").sendVKey(3)   # F3 — voltar / abrir menu de lista
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
        session.findById("wnd[0]").sendVKey(9)   # Menu "Lista → Exportar → Arquivo local"
        session.findById("wnd[1]").sendVKey(0)   # Confirmar formato (tabela)
        logger.debug("Diálogo de exportação aberto.")
    except Exception as exc:  # noqa: BLE001
        # Erro 617 ("virtual key not enabled") pode ocorrer mesmo com dados em
        # situações inesperadas. Logar como WARNING com contexto para diagnóstico.
        logger.warning(
            "Erro ao abrir diálogo de exportação para %s→%s "
            "(dia pode ter dados mas VKey desabilitada): %s",
            str_low, str_high, exc,
        )
        return None

    try:
        # Preencher caminho e nome do arquivo de destino
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = pasta
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nome_arquivo
        session.findById("wnd[1]/tbar[0]/btn[11]").press()  # Substituir / Salvar
        logger.info("CSV exportado: %s", caminho_csv)
    except Exception as exc:  # noqa: BLE001
        logger.error("Erro ao salvar CSV '%s': %s", caminho_csv, exc)
        return None

    try:
        # Fechar tela de resultado e voltar ao menu inicial
        session.findById("wnd[0]").sendVKey(3)  # Voltar
        session.findById("wnd[0]").sendVKey(3)  # Voltar ao menu SAP
        logger.debug("Telas fechadas após exportação.")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Aviso ao fechar telas pós-exportação: %s", exc)

    return caminho_csv


# ---------------------------------------------------------------------------
# ITERAÇÃO POR PERÍODO
# ---------------------------------------------------------------------------

def _intervalos_diarios(data_ini: datetime, data_fim: datetime):
    """Gera pares (low, high) para cada dia do período."""
    atual = data_ini
    while atual <= data_fim:
        yield atual, atual
        atual += timedelta(days=1)


def _intervalos_semanais(data_ini: datetime, data_fim: datetime):
    """
    Gera pares (low, high) para cada semana do período.
    A semana começa na segunda-feira (weekday=0).
    """
    inicio_semana = data_ini - timedelta(days=data_ini.weekday())
    while inicio_semana <= data_fim:
        fim_semana = inicio_semana + timedelta(days=6)
        low = max(inicio_semana, data_ini)
        high = min(fim_semana, data_fim)
        yield low, high
        inicio_semana += timedelta(weeks=1)


def _intervalos_mensais(data_ini: datetime, data_fim: datetime):
    """
    Gera pares (low, high) para cada mês do período.
    Trata corretamente meses de tamanhos diferentes e virada de ano.
    """
    ano, mes = data_ini.year, data_ini.month
    while True:
        primeiro_dia = datetime(ano, mes, 1)
        ultimo_dia = datetime(ano, mes, calendar.monthrange(ano, mes)[1])
        low = max(primeiro_dia, data_ini)
        high = min(ultimo_dia, data_fim)
        if low > data_fim:
            break
        yield low, high
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1


def iterar_periodo(
    data_ini: datetime,
    data_fim: datetime,
    periodicidade: str,
    session,
    pasta: str,
    logger: logging.Logger,
) -> list:
    """
    Itera pelo período conforme a periodicidade e executa extrair_razao
    para cada intervalo. Retorna lista de caminhos de CSVs gerados com sucesso.
    Dias sem partidas são ignorados silenciosamente (já logados como INFO).
    """
    geradores = {
        "diaria": _intervalos_diarios,
        "semanal": _intervalos_semanais,
        "mensal": _intervalos_mensais,
    }
    gerador = geradores.get(periodicidade, _intervalos_diarios)

    csvs_gerados = []
    for low, high in gerador(data_ini, data_fim):
        caminho = extrair_razao(session, low, high, pasta, logger)
        if caminho is None:
            # Falha real — aviso já emitido dentro de extrair_razao
            logger.warning(
                "Extração falhou para o intervalo %s→%s. Continuando...",
                _fmt_data(low),
                _fmt_data(high),
            )
        elif caminho == _SEM_DADOS:
            # Dia sem partidas — tratado como INFO, não há CSV para adicionar
            pass
        else:
            csvs_gerados.append(caminho)

    return csvs_gerados


# ---------------------------------------------------------------------------
# CONSOLIDAÇÃO
# ---------------------------------------------------------------------------

# Padrões típicos de mojibake (UTF-8 lido como latin-1)
_MOJIBAKE_PATTERNS = ("Ã§", "Ã£", "Ãµ", "Ãº", "LÃ§", "Ã£o", "Ã§Ã£", "Ã©", "Ã³")


def _tem_mojibake(texto: str) -> bool:
    """Retorna True se o texto apresentar padrões típicos de duplo-encoding."""
    return any(p in texto for p in _MOJIBAKE_PATTERNS)


def _corrigir_mojibake(texto: str) -> str:
    """
    Tenta corrigir duplo-encoding reinterpretando a string como latin-1 → utf-8.
    Retorna o texto original se a correção falhar.
    """
    try:
        return texto.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return texto


def ler_csv_corrigindo_encoding(caminho: str, logger: logging.Logger) -> list:
    """
    Lê um arquivo CSV detectando o encoding correto e corrigindo mojibake.

    Estratégia:
      1. Tenta utf-8-sig / utf-8 primeiro (formato mais comum nos CSVs SAP atuais).
      2. Se falhar (UnicodeDecodeError) ou detectar padrões de mojibake, tenta
         cp1252 / latin-1.
      3. Se o conteúdo lido contiver padrões de mojibake (ex.: 'Ã§', 'Ã£'),
         aplica correção reinterpretando bytes como latin-1 → utf-8.
      4. Última alternativa: latin-1 com errors='replace'.

    Retorna lista de strings (linhas sem \\r\\n) com encoding corrigido.
    Loga o encoding detectado e se houve correção de mojibake.
    """
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(caminho, "r", encoding=enc, errors="strict") as f:
                linhas = [l.rstrip("\r\n") for l in f]
            # Verifica mojibake na amostra inicial
            amostra = " ".join(linhas[:30])
            if _tem_mojibake(amostra):
                linhas = [_corrigir_mojibake(l) for l in linhas]
                logger.info(
                    "Arquivo '%s': encoding=%s, mojibake detectado e corrigido.",
                    caminho, enc,
                )
            else:
                logger.debug(
                    "Arquivo '%s': encoding=%s.", caminho, enc
                )
            return linhas
        except UnicodeDecodeError:
            continue

    # Última alternativa
    logger.warning(
        "Não foi possível ler '%s' sem erros; usando latin-1 com replace.", caminho
    )
    with open(caminho, "r", encoding="latin-1", errors="replace") as f:
        return [l.rstrip("\r\n") for l in f]


def consolidar(
    csvs: list,
    data_ini: datetime,
    data_fim: datetime,
    pasta: str,
    logger: logging.Logger,
):
    """
    Empilha os CSVs diários em um único arquivo consolidado.

    - Grava um único cabeçalho no topo.
    - Remove linhas de total/subtotal (St contém '*'), separadores,
      cabeçalhos repetidos e linhas em branco.
    - Ordena as linhas de dados cronologicamente por Dt.Lçto.
    - Saída em UTF-8.

    Retorna o caminho do consolidado, ou None se nenhum CSV foi fornecido.
    """
    if not csvs:
        logger.warning("Nenhum CSV para consolidar.")
        return None

    os.makedirs(pasta, exist_ok=True)
    nome = (
        f"Consolidado_{data_ini.strftime('%d.%m.%Y')}"
        f"_a_{data_fim.strftime('%d.%m.%Y')}.csv"
    )
    caminho_consolidado = os.path.join(pasta, nome)

    logger.info(
        "Iniciando consolidação de %d arquivo(s) → '%s'", len(csvs), caminho_consolidado
    )

    todas_linhas = []
    for caminho_csv in sorted(csvs):  # ordem cronológica pelo nome do arquivo
        try:
            for linha in ler_csv_corrigindo_encoding(caminho_csv, logger):
                todas_linhas.append(linha)
        except Exception as exc:  # noqa: BLE001
            logger.error("Erro ao ler '%s': %s", caminho_csv, exc)

    # Limpar (remover totais, separadores, cabeçalhos repetidos) e ordenar
    linhas_limpas = limpar_e_ordenar(todas_linhas, logger)

    with open(caminho_consolidado, "w", encoding="utf-8-sig", newline="") as saida:
        saida.write(CABECALHO_CONSOLIDADO + "\n")
        for linha in linhas_limpas:
            saida.write(linha + "\n")

    logger.info(
        "Consolidação concluída: %d linha(s) de dados gravadas em '%s'.",
        len(linhas_limpas),
        caminho_consolidado,
    )
    return caminho_consolidado


# ---------------------------------------------------------------------------
# CLASSIFICAÇÃO
# ---------------------------------------------------------------------------

def indice_ndoc_por_conta(linhas_dados: list) -> dict:
    """
    Constrói e retorna um dicionário mapeando cada Nº doc. ao conjunto de
    contas em que ele aparece no conjunto de dados fornecido.

    Ex.: {'100123': {'1202060000', '2104030007'}, ...}
    """
    indice: dict = {}
    for linha in linhas_dados:
        campos = parse_linha(linha)
        if len(campos) <= max(IDX_NRDOC, IDX_CONTA):
            continue
        nr = campos[IDX_NRDOC].strip()
        conta = campos[IDX_CONTA].strip()
        if nr:
            indice.setdefault(nr, set()).add(conta)
    return indice


def classificar_linha(campos: list, indice_ndoc: dict) -> str:
    """
    Determina a classificação de uma linha de Mútuo (conta iniciando em '1202').

    Regras em ordem de prioridade:
      1. Se o Nº doc. também aparece na conta IOF  (2104030007) → 'IOF'
      2. Se o Nº doc. também aparece na conta IRRF (1103050050) → 'IRRF'
      3. Se o Nº doc. também aparece na conta Juros(4401020004) → 'Juros'
      4. Senão, pelo campo CL da própria linha:
           CL 09  → 'Adições'
           CL 19  → 'Baixa'
           outro  → '' (sem classificação; a chamada deve logar o CL)

    Linhas de conta que não iniciem com '1202' devem ser filtradas antes de
    chamar esta função.

    Retorna a string de classificação ou '' se não for possível determinar.
    """
    if len(campos) <= IDX_CONTA:
        return ""

    nr = campos[IDX_NRDOC].strip() if len(campos) > IDX_NRDOC else ""
    contas_do_doc = indice_ndoc.get(nr, set())

    # Prioridade 1-3: vínculo por Nº doc. com contas especiais
    for conta_ref in (CONTA_IOF, CONTA_IRRF, CONTA_JUROS):
        if conta_ref in contas_do_doc:
            return CONTAS_CLASSIFICACAO[conta_ref]

    # Prioridade 4: campo CL da própria linha de Mútuo
    cl = campos[IDX_CL].strip() if len(campos) > IDX_CL else ""
    cl_norm = cl.zfill(2)  # normaliza '9' → '09' sem afetar '19'
    if cl_norm == "09":
        return "Adições"
    if cl_norm == "19":
        return "Baixa"
    return ""  # CL desconhecido — quem chamou deve logar


def classificar(
    arquivo_consolidado: str,
    pasta_saida: str,
    logger: logging.Logger,
):
    """
    Lê o arquivo consolidado e gera um arquivo classificado com a coluna
    'Classificação' acrescentada ao final de cada linha.

    Apenas as linhas de Mútuo (conta iniciando em '1202') recebem classificação.
    As linhas de contrapartida (IOF/IRRF/Juros) servem como referência de
    vínculo mas NÃO recebem classificação e não entram nos totais do resumo.

    Regras de prioridade para linhas de Mútuo:
      1. Nº doc. aparece em conta 2104030007 → IOF
      2. Nº doc. aparece em conta 1103050050 → IRRF
      3. Nº doc. aparece em conta 4401020004 → Juros
      4. CL == 09 → Adições; CL == 19 → Baixa; outro → sem classificação (logado)

    Retorna o caminho do arquivo classificado, ou None em caso de falha.
    """
    if not os.path.isfile(arquivo_consolidado):
        logger.error("Arquivo consolidado não encontrado: '%s'", arquivo_consolidado)
        return None

    os.makedirs(pasta_saida, exist_ok=True)

    logger.info("Classificando '%s'...", arquivo_consolidado)

    # Leitura de todas as linhas de dados com correção de encoding
    linhas_dados = []
    cabecalho_original = None
    try:
        for linha in ler_csv_corrigindo_encoding(arquivo_consolidado, logger):
            tp = tipo_de_linha(linha)
            if tp == "cabecalho":
                cabecalho_original = linha
            elif tp == "dado":
                linhas_dados.append(linha)
    except Exception as exc:  # noqa: BLE001
        logger.error("Erro ao ler consolidado para classificação: %s", exc)
        return None

    # Construir índice Nº doc. → conjunto de contas (para prioridade 1-3)
    nrdoc_para_contas = indice_ndoc_por_conta(linhas_dados)

    # Contadores para LOG detalhado
    contagens: dict = {}
    sem_classif_mutuo: list = []   # linhas de Mútuo sem classificação (CL desconhecido)
    nao_mutuo_ignoradas: int = 0   # linhas fora de 1202*

    linhas_classificadas = []

    for linha in linhas_dados:
        campos = parse_linha(linha)
        conta = campos[IDX_CONTA].strip() if len(campos) > IDX_CONTA else ""

        if conta.startswith(PREFIXO_MUTUO):
            classif = classificar_linha(campos, nrdoc_para_contas)
            if classif:
                contagens[classif] = contagens.get(classif, 0) + 1
            else:
                cl = campos[IDX_CL].strip() if len(campos) > IDX_CL else "(ausente)"
                nr = campos[IDX_NRDOC].strip() if len(campos) > IDX_NRDOC else ""
                sem_classif_mutuo.append((nr, conta, cl))
                logger.debug(
                    "Mútuo sem classificação: Nº doc.=%s Conta=%s CL='%s'", nr, conta, cl
                )
        else:
            classif = ""
            nao_mutuo_ignoradas += 1

        linhas_classificadas.append(linha.rstrip("|") + "|" + classif + "|")

    # Nome do arquivo de saída
    data_ini_str, data_fim_str = _extrair_datas_do_nome(
        os.path.basename(arquivo_consolidado)
    )
    nome_saida = f"Classificado_{data_ini_str}_a_{data_fim_str}.csv"
    caminho_saida = os.path.join(pasta_saida, nome_saida)

    cabecalho_saida = (
        (cabecalho_original or CABECALHO_CONSOLIDADO).rstrip("|")
        + "|Classificação|"
    )

    with open(caminho_saida, "w", encoding="utf-8-sig", newline="") as f:
        f.write(cabecalho_saida + "\n")
        for linha in linhas_classificadas:
            f.write(linha + "\n")

    logger.info("Arquivo classificado gerado: '%s'", caminho_saida)
    logger.info(
        "Contagem por classificação (linhas Mútuo): %s", contagens
    )
    logger.info(
        "Linhas de Mútuo sem classificação (CL desconhecido): %d",
        len(sem_classif_mutuo),
    )
    if sem_classif_mutuo:
        for nr, conta, cl in sem_classif_mutuo[:20]:  # até 20 exemplos
            logger.warning(
                "  Mútuo sem classif → Nº doc.=%s Conta=%s CL='%s'", nr, conta, cl
            )
    logger.info(
        "Linhas não-Mútuo ignoradas (sem classificação): %d", nao_mutuo_ignoradas
    )

    return caminho_saida


# ---------------------------------------------------------------------------
# TABELAS DE CONSOLIDAÇÃO INTERCOMPANY (ZFIT009 / ZCO059)
# ---------------------------------------------------------------------------

def _linha_pipe_separador(partes: list) -> bool:
    """Retorna True quando a linha é composta apenas por hífens/separadores."""
    if not partes:
        return True
    com_conteudo = [p for p in partes if p.strip()]
    if not com_conteudo:
        return True
    return all(not p.strip().replace("-", "").strip() for p in com_conteudo)


def parse_tabela_pipe(caminho: str, tipo: str, logger: logging.Logger) -> tuple:
    """
    Faz parse robusto de tabela SAP largura-fixa delimitada por '|'.

    Retorna (registros, cabecalho_encontrado).
    - tipo='zfit009' → [{'Cliente': '2200000404', 'Divisão': 'A001'}, ...]
    - tipo='zco059'  → [{'Empresa': 'MRVH', 'Divisão': 'A001', 'Consolida': 'S'}, ...]
    """
    registros = []
    cabecalho_encontrado = False

    for linha in ler_csv_corrigindo_encoding(caminho, logger):
        texto = linha.strip()
        if not texto:
            continue

        if texto.startswith("|"):
            partes = [p.strip() for p in linha.split("|")[1:-1]]
        elif "|" in texto:
            partes = [p.strip() for p in linha.split("|")]
        else:
            continue

        if _linha_pipe_separador(partes):
            continue

        linha_norm = linha.upper()
        if tipo == "zfit009":
            if "CLIENTE" in linha_norm and ("DIVIS" in linha_norm or "DIVISÃ" in linha_norm):
                cabecalho_encontrado = True
                continue
            if len(partes) < 2:
                continue
            cliente = partes[-2].strip()
            divisao = partes[-1].strip()
            if not re.fullmatch(r"\d{10}", cliente):
                continue
            registros.append({"Cliente": cliente, "Divisão": divisao})
        elif tipo == "zco059":
            if "EMPRESA" in linha_norm and ("DIVIS" in linha_norm or "DIVISÃ" in linha_norm):
                cabecalho_encontrado = True
                continue
            if len(partes) < 3:
                continue
            empresa = partes[-3].strip()
            divisao = partes[-2].strip()
            consolida = partes[-1].strip().upper()
            if not empresa and not divisao and not consolida:
                continue
            registros.append(
                {"Empresa": empresa, "Divisão": divisao, "Consolida": consolida}
            )

    return registros, cabecalho_encontrado


def _validar_importacao_tabela(
    caminho: str, tipo: str, registros: list, cabecalho_encontrado: bool, logger: logging.Logger
) -> bool:
    """Valida existência, cabeçalho e volume de dados importados."""
    if not os.path.isfile(caminho):
        logger.error("Arquivo %s não encontrado: '%s'", tipo.upper(), caminho)
        return False
    if not cabecalho_encontrado:
        logger.error("Cabeçalho esperado não encontrado em '%s' (%s).", caminho, tipo.upper())
        return False
    if len(registros) == 0:
        logger.error("Nenhuma linha de dados válida encontrada em '%s' (%s).", caminho, tipo.upper())
        return False
    logger.info(
        "Importação %s validada com sucesso: %d linha(s) de dados.",
        tipo.upper(),
        len(registros),
    )
    return True


def _salvar_tabela_normalizada(tipo: str, registros: list, pasta_tabelas: str, logger: logging.Logger) -> str:
    """Salva tabela tratada em UTF-8-SIG na pasta de tabelas do projeto."""
    os.makedirs(pasta_tabelas, exist_ok=True)
    nome = "ZFIT009.csv" if tipo == "zfit009" else "ZCO059.csv"
    caminho = os.path.join(pasta_tabelas, nome)

    with open(caminho, "w", encoding="utf-8-sig", newline="") as f:
        if tipo == "zfit009":
            f.write("Cliente|Divisão\n")
            for item in registros:
                f.write(f"{item['Cliente']}|{item['Divisão']}\n")
        else:
            f.write("Empresa|Divisão|Consolida\n")
            for item in registros:
                f.write(f"{item['Empresa']}|{item['Divisão']}|{item['Consolida']}\n")

    logger.info("Tabela %s salva em '%s'.", tipo.upper(), caminho)
    return caminho


def importar_tabela_de_arquivo(caminho_origem: str, tipo: str, pasta_tabelas: str, logger: logging.Logger):
    """Importa ZFIT009/ZCO059 a partir de arquivo já exportado manualmente."""
    registros, cabecalho_encontrado = parse_tabela_pipe(caminho_origem, tipo, logger)
    if not _validar_importacao_tabela(caminho_origem, tipo, registros, cabecalho_encontrado, logger):
        return None
    return _salvar_tabela_normalizada(tipo, registros, pasta_tabelas, logger)


def importar_zfit009_sap(session, pasta_tabelas: str, logger: logging.Logger):
    """
    Importa ZFIT009 via SAP GUI Scripting (SE16) e salva em saidas/tabelas/ZFIT009.csv.
    """
    os.makedirs(pasta_tabelas, exist_ok=True)
    caminho = os.path.join(pasta_tabelas, "ZFIT009.csv")
    try:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "se16"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "ZFIT009"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]").sendVKey(14)
        session.findById("wnd[1]/usr/chk[1,5]").Selected = True
        session.findById("wnd[1]/usr/chk[1,11]").Selected = True
        session.findById("wnd[1]").sendVKey(8)
        session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").Select()
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = pasta_tabelas
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZFIT009.csv"
        session.findById("wnd[1]/tbar[0]/btn[11]").Press()
    except Exception as exc:  # noqa: BLE001
        logger.error("Falha na importação SAP da ZFIT009: %s", exc)
        return None

    registros, cabecalho_encontrado = parse_tabela_pipe(caminho, "zfit009", logger)
    if not _validar_importacao_tabela(caminho, "zfit009", registros, cabecalho_encontrado, logger):
        return None
    return _salvar_tabela_normalizada("zfit009", registros, pasta_tabelas, logger)


def importar_zco059_sap(session, pasta_tabelas: str, logger: logging.Logger):
    """
    Importa ZCO059 via SAP GUI Scripting e salva em saidas/tabelas/ZCO059.csv.
    """
    os.makedirs(pasta_tabelas, exist_ok=True)
    caminho = os.path.join(pasta_tabelas, "ZCO059.csv")
    try:
        session.findById("wnd[0]/tbar[0]/okcd").Text = "zco059"
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").pressToolbarButton("&COL0")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectColumn("CONSOLIDA")
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").pressToolbarButton("&MB_FILTER")
        session.findById("wnd[1]/usr/txtDYN001-LOW").Text = "S"
        session.findById("wnd[1]/tbar[0]/btn[0]").Press()
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").pressToolbarButton("&MB_EXPORT")
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
        session.findById("wnd[1]/tbar[0]/btn[0]").Press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = pasta_tabelas
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ZCO059.csv"
        session.findById("wnd[1]/tbar[0]/btn[11]").Press()
    except Exception as exc:  # noqa: BLE001
        logger.error("Falha na importação SAP da ZCO059: %s", exc)
        return None

    registros, cabecalho_encontrado = parse_tabela_pipe(caminho, "zco059", logger)
    if not _validar_importacao_tabela(caminho, "zco059", registros, cabecalho_encontrado, logger):
        return None
    return _salvar_tabela_normalizada("zco059", registros, pasta_tabelas, logger)


def carregar_tabela_normalizada(tipo: str, pasta_tabelas: str, logger: logging.Logger):
    """Carrega tabela previamente normalizada da pasta de tabelas."""
    caminho = os.path.join(pasta_tabelas, "ZFIT009.csv" if tipo == "zfit009" else "ZCO059.csv")
    if not os.path.isfile(caminho):
        logger.error("Tabela %s não encontrada em '%s'.", tipo.upper(), caminho)
        return None
    registros, cabecalho_encontrado = parse_tabela_pipe(caminho, tipo, logger)
    if not _validar_importacao_tabela(caminho, tipo, registros, cabecalho_encontrado, logger):
        return None
    return registros


def gerar_consolidacao(pasta_tabelas: str, pasta_consolidado: str, logger: logging.Logger):
    """
    Gera Consolidacao_Cliente_Divisao_Consolida.csv a partir de ZFIT009 e ZCO059.
    """
    zfit = carregar_tabela_normalizada("zfit009", pasta_tabelas, logger)
    zco = carregar_tabela_normalizada("zco059", pasta_tabelas, logger)
    if zfit is None or zco is None:
        return None

    divisao_consolida = {}
    for item in zco:
        divisao = item["Divisão"].strip()
        if not divisao:
            continue
        valor = item["Consolida"].strip().upper()
        if valor == "S":
            divisao_consolida[divisao] = "S"
        elif divisao not in divisao_consolida:
            divisao_consolida[divisao] = valor

    por_cliente = defaultdict(set)
    divisao_vazia = 0
    linhas_saida = []
    for item in zfit:
        cliente = item["Cliente"].strip()
        divisao = item["Divisão"].strip()
        por_cliente[cliente].add(divisao)
        if not divisao:
            divisao_vazia += 1
        consolida = "S" if divisao and divisao_consolida.get(divisao, "") == "S" else ""
        linhas_saida.append((cliente, divisao, consolida))

    clientes_ambiguos = {c: sorted(v) for c, v in por_cliente.items() if len(v) > 1}
    if clientes_ambiguos:
        logger.warning("Clientes com múltiplas divisões na ZFIT009: %d", len(clientes_ambiguos))
        for cliente, divs in list(clientes_ambiguos.items())[:30]:
            logger.warning("  Cliente %s → divisões %s", cliente, ", ".join(d or "(vazia)" for d in divs))
    if divisao_vazia:
        logger.warning("Registros com divisão vazia na ZFIT009: %d", divisao_vazia)

    os.makedirs(pasta_consolidado, exist_ok=True)
    caminho_saida = os.path.join(
        pasta_consolidado,
        "Consolidacao_Cliente_Divisao_Consolida.csv",
    )
    with open(caminho_saida, "w", encoding="utf-8-sig", newline="") as f:
        f.write("Cliente|Divisão|Consolida\n")
        for cliente, divisao, consolida in linhas_saida:
            f.write(f"{cliente}|{divisao}|{consolida}\n")

    logger.info(
        "Consolidação Cliente→Divisão→Consolida gerada: '%s' (%d linhas).",
        caminho_saida,
        len(linhas_saida),
    )
    return caminho_saida


def aplicar_status_consolidacao(
    arquivo_classificado: str,
    arquivo_consolidacao: str,
    pasta_saida: str,
    logger: logging.Logger,
):
    """
    Adiciona colunas Divisão, Consolida e Status Consolidação ao classificado.
    """
    if not os.path.isfile(arquivo_classificado):
        logger.error("Arquivo classificado não encontrado: '%s'", arquivo_classificado)
        return None
    if not os.path.isfile(arquivo_consolidacao):
        logger.error("Arquivo de consolidação não encontrado: '%s'", arquivo_consolidacao)
        return None

    idx_consolida = {}
    for linha in ler_csv_corrigindo_encoding(arquivo_consolidacao, logger):
        texto = linha.strip()
        if not texto:
            continue
        if texto.lower().startswith("cliente|"):
            continue
        if texto.startswith("|"):
            partes = [p.strip() for p in linha.split("|")[1:-1]]
        else:
            partes = [p.strip() for p in linha.split("|")]
        if len(partes) < 3:
            continue
        cliente, divisao, consolida = partes[0], partes[1], partes[2].upper()
        idx_consolida.setdefault(cliente, []).append((divisao, consolida))

    linhas_saida = []
    clientes_sem_match = set()
    elimina = 0
    nao_consolida = 0

    cabecalho_saida = None
    for linha in ler_csv_corrigindo_encoding(arquivo_classificado, logger):
        tp = tipo_de_linha(linha)
        if tp == "cabecalho":
            cabecalho_saida = linha.rstrip("|") + "|Divisão|Consolida|Status Consolidação|"
            continue
        if tp != "dado":
            continue

        campos = parse_linha(linha)
        cliente = campos[IDX_CLIENTE].strip() if len(campos) > IDX_CLIENTE else ""
        correspondencias = idx_consolida.get(cliente, [])

        if not correspondencias:
            divisao = ""
            consolida = ""
            status = "Não Consolida"
            if cliente:
                clientes_sem_match.add(cliente)
            nao_consolida += 1
        else:
            divs = sorted({d for d, _ in correspondencias if d})
            divisao = ";".join(divs)
            consolida = "S" if any(c == "S" for _, c in correspondencias) else ""
            if consolida == "S":
                status = "Elimina"
                elimina += 1
            else:
                status = "Não Consolida"
                nao_consolida += 1

        linhas_saida.append(linha.rstrip("|") + f"|{divisao}|{consolida}|{status}|")

    os.makedirs(pasta_saida, exist_ok=True)
    caminho_saida = os.path.join(pasta_saida, os.path.basename(arquivo_classificado))
    with open(caminho_saida, "w", encoding="utf-8-sig", newline="") as f:
        f.write((cabecalho_saida or (CABECALHO_CONSOLIDADO.rstrip("|") + "|Classificação|Divisão|Consolida|Status Consolidação|")) + "\n")
        for linha in linhas_saida:
            f.write(linha + "\n")

    logger.info("Status consolidação aplicado em '%s'.", caminho_saida)
    logger.info("Status Consolidação: Elimina=%d | Não Consolida=%d", elimina, nao_consolida)
    logger.info("Clientes sem correspondência na consolidação: %d", len(clientes_sem_match))
    for cliente in sorted(list(clientes_sem_match))[:30]:
        logger.warning("  Cliente sem match: %s", cliente)

    return caminho_saida


def separar_contrapartidas(
    arquivo_classificado: str,
    pasta_saida: str,
    logger: logging.Logger,
):
    """
    Move linhas de contas de contrapartida para auditoria e remove do principal.
    """
    if not os.path.isfile(arquivo_classificado):
        logger.error("Arquivo classificado não encontrado: '%s'", arquivo_classificado)
        return None, None

    cabecalho = None
    linhas_principais = []
    linhas_auditoria = []
    for linha in ler_csv_corrigindo_encoding(arquivo_classificado, logger):
        tp = tipo_de_linha(linha)
        if tp == "cabecalho":
            cabecalho = linha
            continue
        if tp != "dado":
            continue
        campos = parse_linha(linha)
        conta = campos[IDX_CONTA].strip() if len(campos) > IDX_CONTA else ""
        if conta in CONTAS_CONTRAPARTIDA_AUDITORIA:
            linhas_auditoria.append(linha)
        else:
            linhas_principais.append(linha)

    base = os.path.basename(arquivo_classificado)
    data_ini_str, data_fim_str = _extrair_datas_do_nome(base)
    periodo = f"{data_ini_str}_a_{data_fim_str}"

    os.makedirs(pasta_saida, exist_ok=True)
    caminho_principal = os.path.join(pasta_saida, base)
    caminho_auditoria = os.path.join(pasta_saida, f"Auditoria_Contrapartidas_{periodo}.csv")

    with open(caminho_principal, "w", encoding="utf-8-sig", newline="") as f:
        f.write((cabecalho or (CABECALHO_CONSOLIDADO.rstrip("|") + "|Classificação|")) + "\n")
        for linha in linhas_principais:
            f.write(linha + "\n")

    with open(caminho_auditoria, "w", encoding="utf-8-sig", newline="") as f:
        f.write((cabecalho or (CABECALHO_CONSOLIDADO.rstrip("|") + "|Classificação|")) + "\n")
        for linha in linhas_auditoria:
            f.write(linha + "\n")

    logger.info(
        "Contrapartidas separadas: %d linha(s) para auditoria em '%s'.",
        len(linhas_auditoria),
        caminho_auditoria,
    )
    logger.info(
        "Arquivo principal atualizado sem contrapartidas: %d linha(s) em '%s'.",
        len(linhas_principais),
        caminho_principal,
    )
    return caminho_principal, caminho_auditoria


def exportar_excel_corporativo(caminho_csv: str, logger: logging.Logger):
    """Gera versão Excel do classificado com formatação corporativa."""
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        logger.error("openpyxl não instalado. Execute: pip install openpyxl")
        return None

    if not os.path.isfile(caminho_csv):
        logger.error("Arquivo CSV não encontrado para exportação Excel: '%s'", caminho_csv)
        return None

    linhas = []
    for linha in ler_csv_corrigindo_encoding(caminho_csv, logger):
        if not linha.strip():
            continue
        texto = linha.strip()
        if texto.startswith("|") and texto.endswith("|"):
            linhas.append([c.strip() for c in linha.split("|")[1:-1]])
        else:
            linhas.append([c.strip() for c in linha.split("|")])

    if not linhas:
        logger.error("Arquivo CSV vazio para exportação Excel: '%s'", caminho_csv)
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Classificado"

    for row in linhas:
        ws.append(row)

    cabecalho_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = cabecalho_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    idx_montante = None
    idx_st = None
    for i, valor in enumerate(linhas[0], start=1):
        if "Montante" in valor:
            idx_montante = i
        if valor.strip() == "St":
            idx_st = i

    for col in ws.columns:
        comprimento = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(comprimento + 2, 80)

    if idx_montante:
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=idx_montante)
            valor = _parse_montante(str(cell.value))
            if valor is not None:
                cell.value = valor
                cell.number_format = '#,##0.00_);[Red](#,##0.00)'
                cell.alignment = Alignment(horizontal="right")

    if idx_st:
        for r in range(2, ws.max_row + 1):
            st = str(ws.cell(row=r, column=idx_st).value or "").strip()
            if "*" in st or "TOTAL" in st.upper():
                for c in range(1, ws.max_column + 1):
                    ws.cell(row=r, column=c).fill = total_fill
                    ws.cell(row=r, column=c).font = Font(bold=True)

    caminho_xlsx = os.path.splitext(caminho_csv)[0] + ".xlsx"
    wb.save(caminho_xlsx)
    logger.info("Excel corporativo gerado: '%s'", caminho_xlsx)
    return caminho_xlsx


def carregar_saldo_inicial(logger: logging.Logger):
    """Carrega saldo inicial persistido (se existir)."""
    if not os.path.isfile(ARQUIVO_SALDO_INICIAL):
        return None
    try:
        with open(ARQUIVO_SALDO_INICIAL, "r", encoding="utf-8") as f:
            dados = json.load(f)
        ultimo = dados.get("ultimo")
        if ultimo and isinstance(ultimo.get("valor"), (int, float)):
            return ultimo
    except Exception as exc:  # noqa: BLE001
        logger.warning("Falha ao ler saldo inicial persistido: %s", exc)
    return None


def salvar_saldo_inicial(valor: float, periodo: str, logger: logging.Logger):
    """Persiste saldo inicial e mantém histórico simples de alterações."""
    os.makedirs(PASTA_BASE, exist_ok=True)
    dados = {"historico": []}
    if os.path.isfile(ARQUIVO_SALDO_INICIAL):
        try:
            with open(ARQUIVO_SALDO_INICIAL, "r", encoding="utf-8") as f:
                dados = json.load(f)
            if "historico" not in dados or not isinstance(dados["historico"], list):
                dados["historico"] = []
        except Exception:  # noqa: BLE001
            dados = {"historico": []}

    registro = {
        "valor": valor,
        "periodo": periodo,
        "data_alteracao": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": getuser(),
    }
    dados["ultimo"] = registro
    dados["historico"].append(registro)

    with open(ARQUIVO_SALDO_INICIAL, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)

    logger.info(
        "Saldo inicial atualizado: %s (período=%s, usuário=%s).",
        _formatar_valor_br(valor),
        periodo,
        registro["usuario"],
    )
    return registro


def calcular_saldo_final_fluxo(
    saldo_inicial: float,
    adicoes: float,
    eliminacoes: float,
    juros: float,
    iof: float,
    irrf: float,
    baixas: float,
) -> float:
    """
    Fórmula de fluxo calibrável:
    Saldo Final = Saldo Inicial + Adições − Eliminações + Juros + IOF − IRRF − Baixas
    """
    return saldo_inicial + adicoes - eliminacoes + juros + iof - irrf - baixas


def gerar_resumo_fluxo(arquivo_classificado: str, pasta_saida: str, logger: logging.Logger):
    """Gera resumo de fluxo (CSV + Excel) com Saldo Inicial e Saldo Final."""
    if not os.path.isfile(arquivo_classificado):
        logger.error("Arquivo classificado não encontrado: '%s'", arquivo_classificado)
        return None

    dados_saldo = carregar_saldo_inicial(logger)
    saldo_default = dados_saldo["valor"] if dados_saldo else 0.0
    periodo_default = dados_saldo["periodo"] if dados_saldo else ""
    prompt = f"Saldo Inicial (BR) [padrão: {_formatar_valor_br(saldo_default)}]: "
    entrada = input(prompt).strip()
    saldo_inicial = _parse_numero_br(entrada) if entrada else saldo_default
    if saldo_inicial is None:
        logger.warning("Saldo Inicial inválido informado; usando 0,00.")
        saldo_inicial = 0.0

    periodo = input(
        f"Período/Trimestre de referência [padrão: {periodo_default or 'não informado'}]: "
    ).strip() or periodo_default or "não informado"
    salvar_saldo_inicial(saldo_inicial, periodo, logger)

    totais = {"Adições": 0.0, "Juros": 0.0, "IOF": 0.0, "IRRF": 0.0, "Baixa": 0.0}
    eliminacoes = 0.0
    elimina_linhas = 0

    idx_classif = None
    idx_status = None
    idx_montante = IDX_MONTANTE

    for linha in ler_csv_corrigindo_encoding(arquivo_classificado, logger):
        tp = tipo_de_linha(linha)
        if tp == "cabecalho":
            campos = parse_linha(linha)
            for i, c in enumerate(campos):
                c_norm = c.strip().lower()
                if "classificação" in c_norm or "classificacao" in c_norm:
                    idx_classif = i
                if "status consolidação" in c_norm or "status consolidacao" in c_norm:
                    idx_status = i
                if "montante" in c_norm:
                    idx_montante = i
            continue
        if tp != "dado":
            continue

        campos = parse_linha(linha)
        conta = campos[IDX_CONTA].strip() if len(campos) > IDX_CONTA else ""
        if not conta.startswith(PREFIXO_MUTUO):
            continue

        classif = campos[idx_classif].strip() if idx_classif is not None and len(campos) > idx_classif else ""
        montante = _parse_montante(campos[idx_montante]) if len(campos) > idx_montante else 0.0
        montante = montante or 0.0
        if classif in totais:
            totais[classif] += montante

        status = campos[idx_status].strip() if idx_status is not None and len(campos) > idx_status else ""
        if status == "Elimina":
            eliminacoes += montante
            elimina_linhas += 1

    adicoes = totais.get("Adições", 0.0)
    juros = totais.get("Juros", 0.0)
    iof = totais.get("IOF", 0.0)
    irrf = totais.get("IRRF", 0.0)
    baixas = totais.get("Baixa", 0.0)
    saldo_final = calcular_saldo_final_fluxo(
        saldo_inicial, adicoes, eliminacoes, juros, iof, irrf, baixas
    )

    linhas_fluxo = [
        ("Saldo Inicial", saldo_inicial),
        ("Adições", adicoes),
        ("(-) Eliminações", eliminacoes),
        ("Juros", juros),
        ("IOF", iof),
        ("IRRF", irrf),
        ("Baixas", baixas),
        ("Saldo Final", saldo_final),
    ]

    data_ini_str, data_fim_str = _extrair_datas_do_nome(os.path.basename(arquivo_classificado))
    nome_csv = f"Resumo_{data_ini_str}_a_{data_fim_str}.csv"
    caminho_csv = os.path.join(pasta_saida, nome_csv)
    os.makedirs(pasta_saida, exist_ok=True)

    with open(caminho_csv, "w", encoding="utf-8-sig", newline="") as f:
        f.write("Etapa|Valor\n")
        for etapa, valor in linhas_fluxo:
            f.write(f"{etapa}|{_formatar_valor_br(valor)}\n")

    caminho_xlsx = os.path.join(pasta_saida, f"Resumo_{data_ini_str}_a_{data_fim_str}.xlsx")
    exportar_resumo_fluxo_xlsx(caminho_xlsx, linhas_fluxo, logger)

    separador = "-" * 62
    print(separador)
    print(f"  RESUMO FLUXO — {os.path.basename(arquivo_classificado)}")
    print(separador)
    for etapa, valor in linhas_fluxo:
        print(f"  {etapa:<20} {_formatar_valor_br(valor):>20}")
    print(separador)

    logger.info("Resumo fluxo gerado em '%s' e '%s'.", caminho_csv, caminho_xlsx)
    logger.info("Linhas marcadas como Elimina: %d", elimina_linhas)
    return caminho_csv


def exportar_resumo_fluxo_xlsx(caminho_xlsx: str, linhas_fluxo: list, logger: logging.Logger):
    """Exporta resumo de fluxo para Excel com destaque de saldos."""
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
    except ImportError:
        logger.error("openpyxl não instalado. Execute: pip install openpyxl")
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo Fluxo"
    ws.append(["Etapa", "Valor"])

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    destaque_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for etapa, valor in linhas_fluxo:
        ws.append([etapa, valor])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 20

    for r in range(2, ws.max_row + 1):
        ws.cell(row=r, column=2).number_format = '#,##0.00_);[Red](#,##0.00)'
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="right")
        etapa = str(ws.cell(row=r, column=1).value)
        if etapa in {"Saldo Inicial", "Saldo Final"}:
            for c in (1, 2):
                ws.cell(row=r, column=c).font = Font(bold=True)
                ws.cell(row=r, column=c).fill = destaque_fill

    wb.save(caminho_xlsx)
    logger.info("Resumo Excel gerado: '%s'", caminho_xlsx)
    return caminho_xlsx


# ---------------------------------------------------------------------------
# RESUMO
# ---------------------------------------------------------------------------

# Ordem fixa de exibição das categorias no resumo
_ORDEM_RESUMO = ["Adições", "Juros", "IOF", "IRRF", "Baixa"]


def exportar_resumo_xlsx(
    caminho_xlsx: str,
    categorias_ordenadas: list,
    contagens: dict,
    totais: dict,
    total_geral_linhas: int,
    total_geral_valor: float,
    logger: logging.Logger,
) -> None:
    """
    Exporta o resumo para Excel (.xlsx) usando openpyxl.

    Formata o arquivo com:
      - Cabeçalho em negrito.
      - Larguras de coluna ajustadas.
      - Valores monetários com formato numérico contábil.
      - Linha de TOTAL GERAL em negrito ao final.
    """
    try:
        import openpyxl
        from openpyxl.styles import Font, Alignment
    except ImportError:
        logger.error(
            "openpyxl não instalado. Execute: pip install openpyxl"
        )
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo"

    # Cabeçalho
    ws.append(["Classificação", "Lançamentos", "Total Montante"])
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Linhas de dados
    for cat in categorias_ordenadas:
        total = totais.get(cat, 0.0)
        cnt = contagens.get(cat, 0)
        ws.append([cat, cnt, total])

    # Linha de total geral
    ws.append(["TOTAL GERAL", total_geral_linhas, total_geral_valor])
    ultima_linha = ws.max_row
    for cell in ws[ultima_linha]:
        cell.font = Font(bold=True)

    # Larguras de coluna
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 20

    # Formato numérico monetário para coluna C (linhas de dados + total)
    fmt_numero = '#,##0.00'
    for row in ws.iter_rows(min_row=2, max_row=ultima_linha, min_col=3, max_col=3):
        for cell in row:
            cell.number_format = fmt_numero

    # Alinhamento: coluna B (Lançamentos) à direita
    for row in ws.iter_rows(min_row=2, max_row=ultima_linha, min_col=2, max_col=2):
        for cell in row:
            cell.alignment = Alignment(horizontal="right")

    wb.save(caminho_xlsx)
    logger.info("Resumo Excel gerado: '%s'", caminho_xlsx)


def gerar_resumo(
    arquivo_classificado: str,
    pasta_saida: str,
    logger: logging.Logger,
):
    """
    Lê o arquivo classificado e gera um resumo com totais de Montante Razão
    agrupados por Classificação.

    Considera APENAS as 5 categorias (linhas de Mútuo classificadas):
      Adições, Juros, IOF, IRRF, Baixa

    Saídas:
      - CSV de resumo em UTF-8-sig.
      - Excel (.xlsx) via openpyxl com formatação contábil.
      - Impressão no terminal.
      - Registro no LOG.

    Retorna o caminho do arquivo CSV de resumo, ou None em caso de falha.
    """
    if not os.path.isfile(arquivo_classificado):
        logger.error("Arquivo classificado não encontrado: '%s'", arquivo_classificado)
        return None

    os.makedirs(pasta_saida, exist_ok=True)

    logger.info("Gerando resumo de '%s'...", arquivo_classificado)

    # Detectar índice da coluna Classificação a partir do cabeçalho
    idx_classif = None
    idx_montante_resumo = IDX_MONTANTE  # padrão; recalculado ao ler o cabeçalho

    totais: dict = {}
    contagens_linhas: dict = {}

    try:
        for linha in ler_csv_corrigindo_encoding(arquivo_classificado, logger):
            tp = tipo_de_linha(linha)

            if tp == "cabecalho":
                campos = parse_linha(linha)
                for i, c in enumerate(campos):
                    if "Classificação" in c or "Classificacao" in c:
                        idx_classif = i
                    if "Montante" in c:
                        idx_montante_resumo = i
                continue

            if tp != "dado":
                continue

            campos = parse_linha(linha)

            # Somente linhas de Mútuo (1202*) entram nos totais
            conta = campos[IDX_CONTA].strip() if len(campos) > IDX_CONTA else ""
            if not conta.startswith(PREFIXO_MUTUO):
                continue

            classif = ""
            if idx_classif is not None and len(campos) > idx_classif:
                classif = campos[idx_classif].strip()

            # Ignorar linhas de Mútuo sem classificação no resumo
            if not classif:
                continue

            montante = None
            if len(campos) > idx_montante_resumo:
                montante = _parse_montante(campos[idx_montante_resumo])

            totais[classif] = totais.get(classif, 0.0) + (montante or 0.0)
            contagens_linhas[classif] = contagens_linhas.get(classif, 0) + 1

    except Exception as exc:  # noqa: BLE001
        logger.error("Erro ao ler arquivo classificado para resumo: %s", exc)
        return None

    # Categorias na ordem fixa; inclui apenas as que têm dados
    categorias_com_dados = [c for c in _ORDEM_RESUMO if c in totais]

    total_geral = sum(totais.get(c, 0.0) for c in categorias_com_dados)
    total_linhas_geral = sum(contagens_linhas.get(c, 0) for c in categorias_com_dados)

    # Formatar montante como string BR para exibição
    def _fmt_br(valor: float) -> str:
        abs_v = abs(valor)
        s = f"{abs_v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{s}-" if valor < 0 else s

    # Exibição no terminal e no LOG
    separador = "-" * 62
    linhas_resumo = [
        separador,
        f"  RESUMO — {os.path.basename(arquivo_classificado)}",
        separador,
        f"  {'Classificação':<20} {'Lançamentos':>12} {'Total Montante':>20}",
        separador,
    ]
    for chave in categorias_com_dados:
        linhas_resumo.append(
            f"  {chave:<20} {contagens_linhas[chave]:>12}  "
            f"{_fmt_br(totais[chave]):>18}"
        )
    linhas_resumo += [
        separador,
        f"  {'TOTAL GERAL':<20} {total_linhas_geral:>12}  "
        f"{_fmt_br(total_geral):>18}",
        separador,
    ]

    for l in linhas_resumo:
        print(l)
    for l in linhas_resumo:
        logger.info(l)

    # Derivar datas do nome do arquivo
    data_ini_str, data_fim_str = _extrair_datas_do_nome(
        os.path.basename(arquivo_classificado)
    )

    # Gravar CSV de resumo
    nome_resumo_csv = f"Resumo_{data_ini_str}_a_{data_fim_str}.csv"
    caminho_resumo_csv = os.path.join(pasta_saida, nome_resumo_csv)

    with open(caminho_resumo_csv, "w", encoding="utf-8-sig", newline="") as f:
        f.write("Classificação|Lançamentos|Total Montante\n")
        for chave in categorias_com_dados:
            f.write(
                f"{chave}|{contagens_linhas[chave]}|"
                f"{_fmt_br(totais[chave])}\n"
            )
        f.write(f"TOTAL GERAL|{total_linhas_geral}|{_fmt_br(total_geral)}\n")

    logger.info("Arquivo CSV de resumo gerado: '%s'", caminho_resumo_csv)

    # Gravar Excel de resumo
    nome_resumo_xlsx = f"Resumo_{data_ini_str}_a_{data_fim_str}.xlsx"
    caminho_resumo_xlsx = os.path.join(pasta_saida, nome_resumo_xlsx)
    exportar_resumo_xlsx(
        caminho_resumo_xlsx,
        categorias_com_dados,
        contagens_linhas,
        totais,
        total_linhas_geral,
        total_geral,
        logger,
    )

    return caminho_resumo_csv


# ---------------------------------------------------------------------------
# UTILITÁRIOS DE MENU
# ---------------------------------------------------------------------------

def _extrair_datas_do_nome(nome_arquivo: str) -> tuple:
    """
    Extrai as datas do nome de arquivo no padrão *_DD.MM.AAAA_a_DD.MM.AAAA*.
    Retorna ('DD.MM.AAAA', 'DD.MM.AAAA') ou ('inicio', 'fim') se não encontrar.
    """
    match = re.search(
        r"(\d{2}\.\d{2}\.\d{4})_a_(\d{2}\.\d{2}\.\d{4})", nome_arquivo
    )
    if match:
        return match.group(1), match.group(2)
    return "inicio", "fim"


def _encontrar_ultimo_arquivo(pasta: str, padrao: str):
    """
    Retorna o caminho do arquivo mais recente na pasta que corresponde ao
    padrão glob fornecido, ou None se não houver arquivos.
    """
    arquivos = glob.glob(os.path.join(pasta, padrao))
    if not arquivos:
        return None
    return max(arquivos, key=os.path.getmtime)


def _ler_data(prompt: str, logger: logging.Logger) -> datetime:
    """Solicita uma data no formato dd.mm.aaaa até obter entrada válida."""
    while True:
        valor = input(prompt).strip()
        try:
            return datetime.strptime(valor, "%d.%m.%Y")
        except ValueError:
            msg = (
                f"Data inválida: '{valor}'. "
                "Use o formato dd.mm.aaaa (ex.: 01.06.2026)."
            )
            logger.warning(msg)
            print(f"  ✗ {msg}")


def _ler_periodicidade(logger: logging.Logger) -> str:
    """Solicita a periodicidade até obter opção válida."""
    opcoes = {"1": "diaria", "2": "semanal", "3": "mensal"}
    while True:
        print("\nPeriodicidade:")
        print("  1 - Diária  (padrão)")
        print("  2 - Semanal")
        print("  3 - Mensal")
        valor = input("Opção [1]: ").strip() or "1"
        if valor in opcoes:
            return opcoes[valor]
        msg = f"Opção inválida: '{valor}'. Digite 1, 2 ou 3."
        logger.warning(msg)
        print(f"  ✗ {msg}")


def _ler_sistema(logger: logging.Logger) -> dict:
    """Solicita o sistema SAP até obter opção válida."""
    while True:
        print("\nSistema/Conexão SAP:")
        for chave, info in SISTEMAS.items():
            print(f"  {chave} - {info['nome']}")
        valor = input("Opção: ").strip()
        if valor in SISTEMAS:
            return SISTEMAS[valor]
        msg = f"Opção inválida: '{valor}'. Digite {' ou '.join(SISTEMAS.keys())}."
        logger.warning(msg)
        print(f"  ✗ {msg}")


def selecionar_arquivo_entrada(
    prompt: str, padrao: str, pasta: str, logger: logging.Logger
) -> str:
    """
    Solicita o caminho de um arquivo de entrada com lógica robusta de resolução.

    - Apresenta o arquivo mais recente como padrão (Enter para aceitar).
    - Se o usuário digitar apenas o nome (com ou sem .csv), procura na 'pasta'.
    - Se digitar um caminho absoluto, usa diretamente.
    - Adiciona .csv automaticamente se o arquivo informado não tiver extensão.
    - Valida existência antes de retornar.
    """
    ultimo = _encontrar_ultimo_arquivo(pasta, padrao)
    if ultimo:
        nome_padrao = os.path.basename(ultimo)
        prompt_completo = (
            f"{prompt}\n"
            f"  [padrão: {nome_padrao}]\n"
            f"  (Enter para aceitar, ou informe nome/caminho): "
        )
    else:
        prompt_completo = (
            f"{prompt}\n"
            f"  (nome do arquivo ou caminho completo): "
        )

    while True:
        valor = input(prompt_completo).strip()

        # Enter → usar padrão
        if not valor:
            if ultimo:
                logger.info("Arquivo padrão selecionado: '%s'", ultimo)
                return ultimo
            print("  ✗ Nenhum arquivo disponível como padrão. Informe o caminho.")
            continue

        # Determinar candidato(s) a testar
        candidatos = []

        if os.path.isabs(valor):
            # Caminho absoluto: usar como está (+ .csv se sem extensão)
            if not os.path.splitext(valor)[1]:
                candidatos.append(valor + ".csv")
            candidatos.append(valor)
        else:
            # Nome simples ou relativo: procurar na pasta de saída
            nome = valor
            if not os.path.splitext(nome)[1]:
                nome_csv = nome + ".csv"
                candidatos.append(os.path.join(pasta, nome_csv))
            candidatos.append(os.path.join(pasta, nome))
            # Também testar no diretório atual como fallback
            if not os.path.splitext(nome)[1]:
                candidatos.append(nome + ".csv")
            candidatos.append(nome)

        for c in candidatos:
            if os.path.isfile(c):
                logger.info("Arquivo selecionado: '%s'", c)
                return c

        msg = (
            f"Arquivo não encontrado: '{valor}'. "
            f"Procurado em: '{pasta}' e diretório atual. "
            "Informe o nome ou caminho completo."
        )
        logger.warning(msg)
        print(f"  ✗ {msg}")
        if pasta:
            print(f"  Dica: arquivos disponíveis em '{pasta}':")
            disponiveis = glob.glob(os.path.join(pasta, padrao))
            for arq in sorted(disponiveis)[-5:]:
                print(f"    • {os.path.basename(arq)}")


# ---------------------------------------------------------------------------
# AÇÕES DO MENU
# ---------------------------------------------------------------------------

def _executar_extracao(logger: logging.Logger, caminho_log: str) -> None:
    """Opção 1: Extrai FBL3N e consolida. Correspondente ao fluxo original."""
    print("=" * 60)
    print("  Extrair Razão (FBL3N) + Consolidar")
    print("=" * 60)

    data_inicial = _ler_data("Data inicial (dd.mm.aaaa): ", logger)
    while True:
        data_final = _ler_data("Data final   (dd.mm.aaaa): ", logger)
        if data_final >= data_inicial:
            break
        msg = "A data final deve ser igual ou posterior à data inicial."
        logger.warning(msg)
        print(f"  ✗ {msg}")

    periodicidade = _ler_periodicidade(logger)
    sistema = _ler_sistema(logger)

    logger.info(
        "Extração: período=%s a %s | periodicidade=%s | sistema=%s",
        data_inicial.strftime("%d.%m.%Y"),
        data_final.strftime("%d.%m.%Y"),
        periodicidade,
        sistema["nome"],
    )

    session = conectar_sap(sistema, logger)
    if session is None:
        logger.error("Não foi possível obter sessão SAP. Encerrando.")
        print(f"\nConsulte o LOG para detalhes: {caminho_log}")
        return

    os.makedirs(PASTA_DIARIOS, exist_ok=True)
    csvs = iterar_periodo(
        data_inicial, data_final, periodicidade, session, PASTA_DIARIOS, logger
    )

    consolidado = consolidar(
        csvs, data_inicial, data_final, PASTA_CONSOLIDADO, logger
    )

    print("\n" + "=" * 60)
    if consolidado:
        print(f"  Consolidado gerado: {consolidado}")
    else:
        print("  Nenhum arquivo consolidado gerado (verifique o LOG).")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_classificar(logger: logging.Logger, caminho_log: str) -> None:
    """Opção 2: Classifica um arquivo consolidado existente."""
    print("=" * 60)
    print("  Classificar Consolidado")
    print("=" * 60)

    arq = selecionar_arquivo_entrada(
        "Arquivo consolidado a classificar",
        "Consolidado_*.csv",
        PASTA_CONSOLIDADO,
        logger,
    )

    classificado = classificar(arq, PASTA_CONSOLIDADO, logger)

    print("\n" + "=" * 60)
    if classificado:
        print(f"  Classificado gerado: {classificado}")
    else:
        print("  Falha na classificação (verifique o LOG).")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_importar_tabelas(logger: logging.Logger, caminho_log: str) -> None:
    """Opção: importa/atualiza ZFIT009 e ZCO059 via SAP ou arquivo manual."""
    print("=" * 60)
    print("  Importar ZFIT009 / ZCO059")
    print("=" * 60)

    print("Fonte de importação:")
    print("  1 - SAP GUI Scripting")
    print("  2 - Arquivo já exportado manualmente")
    origem = input("Opção [2]: ").strip() or "2"

    if origem == "1":
        sistema = _ler_sistema(logger)
        session = conectar_sap(sistema, logger)
        if session is None:
            print("  Falha ao conectar no SAP (ver LOG).")
            return
        zfit = importar_zfit009_sap(session, PASTA_TABELAS, logger)
        zco = importar_zco059_sap(session, PASTA_TABELAS, logger)
    else:
        arq_zfit = selecionar_arquivo_entrada(
            "Arquivo ZFIT009 (manual)",
            "ZFIT009*.csv",
            PASTA_TABELAS,
            logger,
        )
        arq_zco = selecionar_arquivo_entrada(
            "Arquivo ZCO059 (manual)",
            "ZCO059*.csv",
            PASTA_TABELAS,
            logger,
        )
        zfit = importar_tabela_de_arquivo(arq_zfit, "zfit009", PASTA_TABELAS, logger)
        zco = importar_tabela_de_arquivo(arq_zco, "zco059", PASTA_TABELAS, logger)

    print("\n" + "=" * 60)
    if zfit and zco:
        print(f"  ZFIT009: {zfit}")
        print(f"  ZCO059 : {zco}")
    else:
        print("  Falha na importação de tabelas (verifique o LOG).")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_gerar_consolidacao(logger: logging.Logger, caminho_log: str) -> None:
    """Opção: gera Consolidação Cliente|Divisão|Consolida."""
    print("=" * 60)
    print("  Gerar Tabela de Consolidação")
    print("=" * 60)

    caminho = gerar_consolidacao(PASTA_TABELAS, PASTA_CONSOLIDADO, logger)

    print("\n" + "=" * 60)
    if caminho:
        print(f"  Consolidação gerada: {caminho}")
    else:
        print("  Falha ao gerar consolidação (verifique o LOG).")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_aplicar_status_consolidacao(logger: logging.Logger, caminho_log: str) -> None:
    """Opção: aplica status de consolidação no classificado."""
    print("=" * 60)
    print("  Aplicar Status Consolidação")
    print("=" * 60)

    arq_classificado = selecionar_arquivo_entrada(
        "Arquivo classificado para complementar",
        "Classificado_*.csv",
        PASTA_CONSOLIDADO,
        logger,
    )
    arq_consolidacao = selecionar_arquivo_entrada(
        "Arquivo de consolidação cliente/divisão/consolida",
        "Consolidacao_Cliente_Divisao_Consolida.csv",
        PASTA_CONSOLIDADO,
        logger,
    )
    saida = aplicar_status_consolidacao(
        arq_classificado, arq_consolidacao, PASTA_CONSOLIDADO, logger
    )

    print("\n" + "=" * 60)
    if saida:
        print(f"  Classificado com status: {saida}")
    else:
        print("  Falha ao aplicar status (verifique o LOG).")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_tratar_contrapartidas(logger: logging.Logger, caminho_log: str) -> None:
    """Opção: separa contrapartidas para auditoria."""
    print("=" * 60)
    print("  Tratar Contrapartidas (Auditoria)")
    print("=" * 60)

    arq_classificado = selecionar_arquivo_entrada(
        "Arquivo classificado para separar contrapartidas",
        "Classificado_*.csv",
        PASTA_CONSOLIDADO,
        logger,
    )
    principal, auditoria = separar_contrapartidas(arq_classificado, PASTA_CONSOLIDADO, logger)
    xlsx = exportar_excel_corporativo(principal, logger) if principal else None

    print("\n" + "=" * 60)
    if principal and auditoria:
        print(f"  Principal : {principal}")
        print(f"  Auditoria : {auditoria}")
        if xlsx:
            print(f"  Excel     : {xlsx}")
    else:
        print("  Falha ao tratar contrapartidas (verifique o LOG).")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_saldo_inicial(logger: logging.Logger, caminho_log: str) -> None:
    """Opção: informa/atualiza saldo inicial persistido."""
    print("=" * 60)
    print("  Informar / Atualizar Saldo Inicial")
    print("=" * 60)

    atual = carregar_saldo_inicial(logger)
    if atual:
        print(
            f"Saldo atual: {_formatar_valor_br(atual['valor'])} | "
            f"Período: {atual.get('periodo', '')} | "
            f"Atualizado em: {atual.get('data_alteracao', '')} | "
            f"Usuário: {atual.get('usuario', '')}"
        )
    valor_txt = input("Novo Saldo Inicial (BR): ").strip()
    valor = _parse_numero_br(valor_txt)
    if valor is None:
        print("  ✗ Valor inválido.")
        logger.warning("Saldo inicial inválido informado: '%s'", valor_txt)
        return
    periodo = input("Período/Trimestre de referência: ").strip() or "não informado"
    salvar_saldo_inicial(valor, periodo, logger)

    print("\n" + "=" * 60)
    print("  Saldo inicial atualizado com sucesso.")
    print(f"  Arquivo: {ARQUIVO_SALDO_INICIAL}")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_resumo_fluxo(logger: logging.Logger, caminho_log: str) -> None:
    """Opção: gera resumo de fluxo do classificado."""
    print("=" * 60)
    print("  Resumo de Fluxo (Saldo Inicial → Saldo Final)")
    print("=" * 60)

    arq = selecionar_arquivo_entrada(
        "Arquivo classificado para gerar resumo",
        "Classificado_*.csv",
        PASTA_CONSOLIDADO,
        logger,
    )
    resumo = gerar_resumo_fluxo(arq, PASTA_CONSOLIDADO, logger)

    print("\n" + "=" * 60)
    if resumo:
        print(f"  Resumo gerado: {resumo}")
    else:
        print("  Falha ao gerar resumo (verifique o LOG).")
    print(f"  LOG: {caminho_log}")
    print("=" * 60)


def _executar_tudo(logger: logging.Logger, caminho_log: str) -> None:
    """Opção: executa pipeline completo com consolidação intercompany."""
    print("=" * 60)
    print("  Executar Tudo: Extrair → Classificar → Consolidar → Status → Auditoria → Resumo")
    print("=" * 60)

    data_inicial = _ler_data("Data inicial (dd.mm.aaaa): ", logger)
    while True:
        data_final = _ler_data("Data final   (dd.mm.aaaa): ", logger)
        if data_final >= data_inicial:
            break
        msg = "A data final deve ser igual ou posterior à data inicial."
        logger.warning(msg)
        print(f"  ✗ {msg}")

    periodicidade = _ler_periodicidade(logger)
    sistema = _ler_sistema(logger)

    logger.info(
        "Execução completa: período=%s a %s | periodicidade=%s | sistema=%s",
        data_inicial.strftime("%d.%m.%Y"),
        data_final.strftime("%d.%m.%Y"),
        periodicidade,
        sistema["nome"],
    )

    session = conectar_sap(sistema, logger)
    if session is None:
        logger.error("Não foi possível obter sessão SAP. Encerrando.")
        print(f"\nConsulte o LOG para detalhes: {caminho_log}")
        return

    os.makedirs(PASTA_DIARIOS, exist_ok=True)
    csvs = iterar_periodo(
        data_inicial, data_final, periodicidade, session, PASTA_DIARIOS, logger
    )

    consolidado = consolidar(
        csvs, data_inicial, data_final, PASTA_CONSOLIDADO, logger
    )
    if not consolidado:
        print("  Nenhum dado consolidado. Encerrando fluxo completo.")
        return

    classificado = classificar(consolidado, PASTA_CONSOLIDADO, logger)
    if not classificado:
        print("  Falha na classificação. Resumo não gerado.")
        return

    consolidacao = gerar_consolidacao(PASTA_TABELAS, PASTA_CONSOLIDADO, logger)
    if not consolidacao:
        print("  Falha ao gerar consolidação (importe ZFIT009/ZCO059 e tente novamente).")
        return

    classificado_status = aplicar_status_consolidacao(
        classificado, consolidacao, PASTA_CONSOLIDADO, logger
    )
    if not classificado_status:
        print("  Falha ao aplicar status de consolidação.")
        return

    classificado_final, auditoria = separar_contrapartidas(
        classificado_status, PASTA_CONSOLIDADO, logger
    )
    if not classificado_final:
        print("  Falha no tratamento de contrapartidas.")
        return

    xlsx_classificado = exportar_excel_corporativo(classificado_final, logger)
    resumo = gerar_resumo_fluxo(classificado_final, PASTA_CONSOLIDADO, logger)

    print("\n" + "=" * 60)
    print(f"  Consolidado  : {consolidado}")
    print(f"  Classificado : {classificado_final}")
    print(f"  Consolidação : {consolidacao}")
    print(f"  Auditoria    : {auditoria}")
    if xlsx_classificado:
        print(f"  Excel Class. : {xlsx_classificado}")
    if resumo:
        print(f"  Resumo       : {resumo}")
    print(f"  LOG          : {caminho_log}")
    print("=" * 60)


# ---------------------------------------------------------------------------
# MENU PRINCIPAL
# ---------------------------------------------------------------------------

def mostrar_menu_principal() -> str:
    """
    Exibe o menu principal e retorna a opção escolhida como string.
    """
    print("\n" + "=" * 60)
    print("  MUTDFC — Movimentação do Fluxo de Caixa (Mútuo)")
    print("=" * 60)
    print("  1 - Extrair razão (FBL3N) + Consolidar")
    print("  2 - Classificar consolidado")
    print("  3 - Importar ZFIT009 / ZCO059")
    print("  4 - Gerar tabela de consolidação")
    print("  5 - Aplicar Status Consolidação na movimentação")
    print("  6 - Tratar contrapartidas (auditoria) + Excel classificado")
    print("  7 - Informar/atualizar Saldo Inicial")
    print("  8 - Resumo (Saldo Inicial → Saldo Final)")
    print("  9 - Executar tudo em sequência")
    print("  0 - Sair")
    print("-" * 60)
    return input("Opção: ").strip()


# ---------------------------------------------------------------------------
# PONTO DE ENTRADA
# ---------------------------------------------------------------------------

def main() -> None:
    # Garantir que as pastas de saída existam
    for pasta in (PASTA_DIARIOS, PASTA_CONSOLIDADO, PASTA_TABELAS, PASTA_LOGS):
        os.makedirs(pasta, exist_ok=True)

    logger, caminho_log = configurar_log(PASTA_LOGS)

    acoes = {
        "1": _executar_extracao,
        "2": _executar_classificar,
        "3": _executar_importar_tabelas,
        "4": _executar_gerar_consolidacao,
        "5": _executar_aplicar_status_consolidacao,
        "6": _executar_tratar_contrapartidas,
        "7": _executar_saldo_inicial,
        "8": _executar_resumo_fluxo,
        "9": _executar_tudo,
    }

    while True:
        try:
            opcao = mostrar_menu_principal()
        except KeyboardInterrupt:
            print("\nOperação cancelada pelo usuário.")
            break

        if opcao == "0":
            print("Saindo. Até logo!")
            break
        elif opcao in acoes:
            try:
                acoes[opcao](logger, caminho_log)
            except KeyboardInterrupt:
                print("\nOperação interrompida. Voltando ao menu.")
        else:
            print(f"  ✗ Opção inválida: '{opcao}'. Digite 0 a 9.")


if __name__ == "__main__":
    main()
