' =============================================================================
' MUTDFC.vbs  —  Movimentação do Fluxo de Caixa (contas contábeis de Mútuo)
' Extrai o Razão Contábil (FBL3N) via SAP GUI Scripting e consolida os CSVs
' diários em um único arquivo.
'
' Versão: 1.0
' =============================================================================

' ─────────────────────────────────────────────────────────────────────────────
' CONSTANTES DE CONFIGURAÇÃO — ajuste conforme o ambiente
' ─────────────────────────────────────────────────────────────────────────────

' Pasta onde os arquivos CSV diários e o consolidado serão gravados.
' Altere este caminho para a pasta desejada na sua máquina.
Const PASTA_TRABALHO = "C:\MUTDFC\Extrações"

' Variante de layout da FBL3N (campo "Variante" na tela de seleção).
Const VARIANTE = "MUTDFC"

' Usuário SAP proprietário da variante.
Const USUARIO_SAP = "MS0000240"

' Descrição parcial da conexão SAP S/4 HANA (Opção 1).
' O script procura uma conexão ABERTA cuja descrição contenha este trecho.
Const DESC_S4HANA = "S/4 HANA PRODUCAO"

' Descrição parcial da conexão SAP ECC 6.0 (Opção 2).
Const DESC_ECC = "ECC"

' Descrição completa para abrir nova conexão no SAP Logon (S/4 HANA).
Const CONN_S4HANA_FULL = "MRV SAP S/4 HANA PRODUCAO"

' Descrição completa para abrir nova conexão no SAP Logon (ECC).
Const CONN_ECC_FULL = "03 SAP PRD ECC - BOLHA"

' Cabeçalho exato que deve constar no topo do arquivo consolidado.
Const CABECALHO = "|   St|Atribuição        |Nº doc.   |Dt.Lçto   |Div |Conta     |Fornecedor|Tp.doc. |Cliente   |Data doc. |CL|      Montante Razão|DocCompens|Texto                                             |Imobilizado |Usuário     |Empr|Referência      |Entrado em|Estorno|DiagRede    |"

' ─────────────────────────────────────────────────────────────────────────────
' VARIÁVEIS GLOBAIS
' ─────────────────────────────────────────────────────────────────────────────
Dim g_LogFile   ' caminho completo do arquivo de LOG
Dim g_FSO       ' FileSystemObject (reutilizado em todo o script)
Dim g_Session   ' objeto de sessão SAP GUI

' =============================================================================
' PONTO DE ENTRADA
' =============================================================================
Sub Main()
    Set g_FSO = CreateObject("Scripting.FileSystemObject")

    ' Garantir que a pasta de trabalho existe
    CriarPastaSeNecessario PASTA_TRABALHO

    ' Inicializar LOG
    Dim ts : ts = Format(Now(), "yyyymmdd_HHmmss")
    g_LogFile = PASTA_TRABALHO & "\MUTDFC_log_" & ts & ".txt"
    Log "======================================================"
    Log "MUTDFC — Início da execução: " & Now()
    Log "======================================================"

    ' ── Menu ──
    Dim dataInicial, dataFinal, periodicidade, opcaoSistema
    If Not MostrarMenu(dataInicial, dataFinal, periodicidade, opcaoSistema) Then
        Log "Execução abortada pelo usuário ou dados inválidos."
        MsgBox "Execução cancelada. Consulte o LOG em:" & vbCrLf & g_LogFile, vbInformation, "MUTDFC"
        Exit Sub
    End If

    Log "Parâmetros recebidos:"
    Log "  Período       : " & dataInicial & " a " & dataFinal
    Log "  Periodicidade : " & periodicidade
    Log "  Sistema       : " & opcaoSistema

    ' ── Conexão SAP ──
    If Not ConectarSAP(opcaoSistema) Then
        Log "ERRO: Não foi possível conectar ao SAP. Abortando."
        MsgBox "Não foi possível conectar ao SAP." & vbCrLf & _
               "Verifique o LOG para detalhes:" & vbCrLf & g_LogFile, _
               vbCritical, "MUTDFC"
        Exit Sub
    End If

    ' ── Loop de extração ──
    Dim arquivos()
    Dim numArquivos : numArquivos = 0
    ReDim arquivos(0)

    Dim dtAtual : dtAtual = ParseData(dataInicial)
    Dim dtFinal : dtFinal = ParseData(dataFinal)
    Dim dtProxima

    Do While dtAtual <= dtFinal
        Dim dtBaixaStr : dtBaixaStr = FormatData(dtAtual)

        Select Case periodicidade
            Case "Diária"
                dtProxima = DateAdd("d", 1, dtAtual)
            Case "Semanal"
                dtProxima = DateAdd("ww", 1, dtAtual)
                If dtProxima > dtFinal Then dtProxima = DateAdd("d", 1, dtFinal)
            Case "Mensal"
                dtProxima = DateAdd("m", 1, dtAtual)
                If dtProxima > dtFinal Then dtProxima = DateAdd("d", 1, dtFinal)
        End Select

        ' Para Semanal/Mensal, HIGH é o último dia do intervalo (ou data final)
        Dim dtHighDate
        If periodicidade = "Diária" Then
            dtHighDate = dtAtual
        Else
            dtHighDate = DateAdd("d", -1, dtProxima)
            If dtHighDate > dtFinal Then dtHighDate = dtFinal
        End If

        Dim lowStr  : lowStr  = FormatData(dtAtual)
        Dim highStr : highStr = FormatData(dtHighDate)
        Dim nomeCSV : nomeCSV = PASTA_TRABALHO & "\" & lowStr & ".csv"

        Log "------------------------------------------------------"
        Log "Iniciando extração: " & lowStr & " a " & highStr

        Dim ok : ok = ExtrairRazao(lowStr, highStr, nomeCSV)
        If ok Then
            Log "Extração concluída: " & nomeCSV
            ReDim Preserve arquivos(numArquivos)
            arquivos(numArquivos) = nomeCSV
            numArquivos = numArquivos + 1
        Else
            Log "AVISO: Falha na extração de " & lowStr & " a " & highStr & ". Continuando..."
        End If

        dtAtual = dtProxima
    Loop

    ' ── Consolidação ──
    If numArquivos > 0 Then
        Dim nomeConsolidado
        nomeConsolidado = PASTA_TRABALHO & "\Consolidado_" & _
                          Replace(dataInicial, ".", "") & "_a_" & _
                          Replace(dataFinal, ".", "") & ".csv"
        Log "------------------------------------------------------"
        Log "Iniciando consolidação de " & numArquivos & " arquivo(s)..."
        Consolidar arquivos, numArquivos, nomeConsolidado
    Else
        Log "Nenhum arquivo gerado para consolidar."
    End If

    Log "======================================================"
    Log "MUTDFC — Fim da execução: " & Now()
    Log "======================================================"

    MsgBox "Execução finalizada!" & vbCrLf & _
           "Consulte o LOG em:" & vbCrLf & g_LogFile, vbInformation, "MUTDFC"
End Sub

' =============================================================================
' SUB MostrarMenu
' Apresenta InputBox/MsgBox para coletar: dataInicial, dataFinal,
' periodicidade e opcaoSistema. Retorna True se dados válidos, False caso
' contrário ou cancelado.
' =============================================================================
Function MostrarMenu(ByRef dataInicial, ByRef dataFinal, ByRef periodicidade, ByRef opcaoSistema)
    MostrarMenu = False
    Dim tentativas, maxTentativas : maxTentativas = 3

    ' ── Data inicial ──
    For tentativas = 1 To maxTentativas
        dataInicial = InputBox("Informe a DATA INICIAL do período" & vbCrLf & _
                               "(formato dd.mm.aaaa, ex.: 01.06.2026):", "MUTDFC — Período")
        If dataInicial = "" Then
            Log "Menu: usuário cancelou a entrada de data inicial."
            Exit Function
        End If
        If ValidarData(dataInicial) Then Exit For
        Log "Menu: data inicial inválida informada: '" & dataInicial & "'"
        MsgBox "Data inválida: '" & dataInicial & "'." & vbCrLf & _
               "Use o formato dd.mm.aaaa (ex.: 01.06.2026).", vbExclamation, "MUTDFC"
        If tentativas = maxTentativas Then
            Log "Menu: número máximo de tentativas para data inicial atingido. Abortando."
            Exit Function
        End If
    Next

    ' ── Data final ──
    For tentativas = 1 To maxTentativas
        dataFinal = InputBox("Informe a DATA FINAL do período" & vbCrLf & _
                             "(formato dd.mm.aaaa, ex.: 30.06.2026):", "MUTDFC — Período")
        If dataFinal = "" Then
            Log "Menu: usuário cancelou a entrada de data final."
            Exit Function
        End If
        If ValidarData(dataFinal) Then
            If ParseData(dataFinal) >= ParseData(dataInicial) Then Exit For
            Log "Menu: data final (" & dataFinal & ") anterior à data inicial (" & dataInicial & ")."
            MsgBox "A data final deve ser igual ou posterior à data inicial.", vbExclamation, "MUTDFC"
        Else
            Log "Menu: data final inválida informada: '" & dataFinal & "'"
            MsgBox "Data inválida: '" & dataFinal & "'." & vbCrLf & _
                   "Use o formato dd.mm.aaaa (ex.: 30.06.2026).", vbExclamation, "MUTDFC"
        End If
        If tentativas = maxTentativas Then
            Log "Menu: número máximo de tentativas para data final atingido. Abortando."
            Exit Function
        End If
    Next

    ' ── Periodicidade ──
    Dim msgPer
    msgPer = "Escolha a periodicidade:" & vbCrLf & vbCrLf & _
             "  1 — Diária (padrão)" & vbCrLf & _
             "  2 — Semanal" & vbCrLf & _
             "  3 — Mensal"
    Dim respPer
    For tentativas = 1 To maxTentativas
        respPer = InputBox(msgPer, "MUTDFC — Periodicidade", "1")
        If respPer = "" Then
            Log "Menu: usuário cancelou a entrada de periodicidade."
            Exit Function
        End If
        Select Case Trim(respPer)
            Case "1" : periodicidade = "Diária"   : Exit For
            Case "2" : periodicidade = "Semanal"  : Exit For
            Case "3" : periodicidade = "Mensal"   : Exit For
            Case Else
                Log "Menu: periodicidade inválida: '" & respPer & "'"
                MsgBox "Opção inválida. Digite 1, 2 ou 3.", vbExclamation, "MUTDFC"
        End Select
        If tentativas = maxTentativas Then
            Log "Menu: máximo de tentativas para periodicidade. Abortando."
            Exit Function
        End If
    Next

    ' ── Sistema ──
    Dim msgSis
    msgSis = "Escolha o sistema SAP:" & vbCrLf & vbCrLf & _
             "  1 — SAP S/4 HANA (PRD)" & vbCrLf & _
             "       Descrição: " & CONN_S4HANA_FULL & vbCrLf & vbCrLf & _
             "  2 — SAP ECC 6.0 (PRD)" & vbCrLf & _
             "       Descrição: " & CONN_ECC_FULL
    Dim respSis
    For tentativas = 1 To maxTentativas
        respSis = InputBox(msgSis, "MUTDFC — Sistema", "1")
        If respSis = "" Then
            Log "Menu: usuário cancelou a entrada de sistema."
            Exit Function
        End If
        Select Case Trim(respSis)
            Case "1" : opcaoSistema = 1 : Exit For
            Case "2" : opcaoSistema = 2 : Exit For
            Case Else
                Log "Menu: opção de sistema inválida: '" & respSis & "'"
                MsgBox "Opção inválida. Digite 1 ou 2.", vbExclamation, "MUTDFC"
        End Select
        If tentativas = maxTentativas Then
            Log "Menu: máximo de tentativas para sistema. Abortando."
            Exit Function
        End If
    Next

    MostrarMenu = True
End Function

' =============================================================================
' FUNCTION ConectarSAP
' Conecta/anexa à sessão SAP do sistema escolhido (1 = S/4 HANA, 2 = ECC).
' O usuário já deve estar logado; o script apenas localiza/abre a conexão.
' Retorna True em sucesso, False em falha.
' =============================================================================
Function ConectarSAP(opcaoSistema)
    ConectarSAP = False
    Log "ConectarSAP: buscando SAP GUI Scripting Engine..."

    Dim sapGuiAuto, application, connection, i
    On Error Resume Next

    Set sapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        Log "ERRO ConectarSAP: GetObject('SAPGUI') falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0
        Exit Function
    End If

    Set application = sapGuiAuto.GetScriptingEngine
    If Err.Number <> 0 Then
        Log "ERRO ConectarSAP: GetScriptingEngine falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Determinar qual descrição procurar
    Dim descBusca, descCompleta
    If opcaoSistema = 1 Then
        descBusca   = DESC_S4HANA
        descCompleta = CONN_S4HANA_FULL
    Else
        descBusca   = DESC_ECC
        descCompleta = CONN_ECC_FULL
    End If
    Log "ConectarSAP: procurando conexão com '" & descBusca & "'..."

    ' Percorrer conexões abertas
    Dim numConexoes : numConexoes = application.Children.Count
    Log "ConectarSAP: " & numConexoes & " conexão(ões) abertas encontrada(s)."

    For i = 0 To numConexoes - 1
        On Error Resume Next
        Set connection = application.Children(i)
        Dim descConn : descConn = connection.Description
        If Err.Number <> 0 Then
            Log "ConectarSAP: erro ao acessar conexão " & i & ": " & Err.Description
            Err.Clear
        Else
            Log "ConectarSAP: conexão [" & i & "] descrição = '" & descConn & "'"
            If InStr(1, descConn, descBusca, vbTextCompare) > 0 Then
                On Error GoTo 0
                Set g_Session = connection.Children(0)
                Log "ConectarSAP: conexão localizada e sessão obtida (via conexão existente)."
                ConectarSAP = True
                Exit Function
            End If
        End If
        On Error GoTo 0
    Next

    ' Conexão não encontrada — tentar abrir via OpenConnection
    Log "ConectarSAP: conexão '" & descCompleta & "' não encontrada. Tentando abrir via OpenConnection..."
    On Error Resume Next
    Set connection = application.OpenConnection(descCompleta, True)
    If Err.Number <> 0 Then
        Log "ERRO ConectarSAP: OpenConnection falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If IsNull(connection) Or IsEmpty(connection) Then
        Log "ERRO ConectarSAP: OpenConnection retornou objeto inválido."
        Exit Function
    End If

    Set g_Session = connection.Children(0)
    Log "ConectarSAP: conexão aberta via OpenConnection e sessão obtida."
    ConectarSAP = True
End Function

' =============================================================================
' FUNCTION ExtrairRazao
' Executa a FBL3N para o intervalo [lowStr, highStr] e exporta para nomeCSV.
' Retorna True em sucesso, False em falha.
' =============================================================================
Function ExtrairRazao(lowStr, highStr, nomeCSV)
    ExtrairRazao = False
    Log "ExtrairRazao: período " & lowStr & " a " & highStr

    On Error Resume Next

    ' ── Maximizar e navegar para FBL3N ──
    g_Session.findById("wnd[0]").maximize
    If Err.Number <> 0 Then
        Log "ERRO ExtrairRazao: maximize falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0 : Exit Function
    End If

    g_Session.findById("wnd[0]/tbar[0]/okcd").text = "fbl3n"
    g_Session.findById("wnd[0]").sendVKey 0      ' Enter para confirmar transação
    If Err.Number <> 0 Then
        Log "ERRO ExtrairRazao: navegação para FBL3N falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0 : Exit Function
    End If

    ' ── Abrir dialog de variante ──
    g_Session.findById("wnd[0]").sendVKey 17     ' F5 = obter variante
    If Err.Number <> 0 Then
        Log "ERRO ExtrairRazao: sendVKey 17 (variante) falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0 : Exit Function
    End If

    ' ── Preencher variante e usuário ──
    g_Session.findById("wnd[1]/usr/txtV-LOW").text = VARIANTE
    g_Session.findById("wnd[1]/usr/txtENAME-LOW").text = USUARIO_SAP
    g_Session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
    g_Session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = Len(USUARIO_SAP)
    g_Session.findById("wnd[1]").sendVKey 0      ' confirmar seleção de variante
    If Err.Number <> 0 Then
        Log "ERRO ExtrairRazao: preenchimento de variante falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0 : Exit Function
    End If

    ' ── Preencher datas ──
    g_Session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text  = lowStr
    g_Session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = highStr
    g_Session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").setFocus
    g_Session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").caretPosition = Len(highStr)
    If Err.Number <> 0 Then
        Log "ERRO ExtrairRazao: preenchimento de datas falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0 : Exit Function
    End If

    ' ── Executar relatório ──
    g_Session.findById("wnd[0]").sendVKey 8      ' F8 = executar
    If Err.Number <> 0 Then
        Log "ERRO ExtrairRazao: sendVKey 8 (executar) falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0 : Exit Function
    End If

    ' ── Tratar popup de confirmação (se houver) ──
    g_Session.findById("wnd[0]").sendVKey 20     ' Enter / confirmar eventual popup
    Err.Clear                                    ' ignorar erro se não houver popup

    ' ── Menu Sistema > Lista > Salvar/Enviar > Arquivo local ──
    g_Session.findById("wnd[0]").sendVKey 3      ' F3 (ou Back para sair de popup)
    Err.Clear

    ' Abrir dialog de exportação: menu "Sistema > Lista > Salvar > Arquivo Local"
    g_Session.findById("wnd[0]").sendVKey 9      ' Ctrl+S (lista local — pode variar)
    If Err.Number <> 0 Then
        Log "AVISO ExtrairRazao: sendVKey 9 retornou erro. Err=" & Err.Number & " / " & Err.Description
        Err.Clear
    End If

    ' ── Confirmar formato (botão "Dados não processados" = BUTTON_1) ──
    g_Session.findById("wnd[1]/usr/btnBUTTON_1").press
    If Err.Number <> 0 Then
        Log "AVISO ExtrairRazao: botão BUTTON_1 não encontrado. Err=" & Err.Number & " / " & Err.Description
        Err.Clear
    End If

    ' ── Dialog de caminho/nome do arquivo ──
    g_Session.findById("wnd[1]").sendVKey 0
    Err.Clear

    Dim pasta : pasta = PASTA_TRABALHO
    Dim nomeArquivo : nomeArquivo = g_FSO.GetFileName(nomeCSV)

    g_Session.findById("wnd[1]/usr/ctxtDY_PATH").text     = pasta
    g_Session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = nomeArquivo
    g_Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = Len(nomeArquivo)
    g_Session.findById("wnd[1]/tbar[0]/btn[11]").press    ' Gerar (substituir se existir)
    If Err.Number <> 0 Then
        Log "ERRO ExtrairRazao: gravação do arquivo CSV falhou. Err=" & Err.Number & " / " & Err.Description
        Err.Clear : On Error GoTo 0 : Exit Function
    End If

    ' ── Voltar para a tela inicial ──
    g_Session.findById("wnd[0]").sendVKey 3      ' Back
    Err.Clear
    g_Session.findById("wnd[0]").sendVKey 3      ' Back (sair do relatório)
    Err.Clear

    On Error GoTo 0

    ' Verificar se o arquivo foi realmente criado
    If g_FSO.FileExists(nomeCSV) Then
        Log "ExtrairRazao: arquivo criado com sucesso — " & nomeCSV
        ExtrairRazao = True
    Else
        Log "AVISO ExtrairRazao: arquivo não encontrado após exportação — " & nomeCSV
    End If
End Function

' =============================================================================
' SUB Consolidar
' Empilha todos os CSVs diários em um único arquivo consolidado, mantendo
' apenas um cabeçalho no topo e removendo cabeçalhos repetidos.
' Trata encoding Latin-1 para preservar acentos.
' =============================================================================
Sub Consolidar(arquivos, numArquivos, nomeConsolidado)
    Log "Consolidar: gerando " & nomeConsolidado

    ' Usar ADODB.Stream para leitura em Latin-1 e escrita em UTF-8 (com BOM)
    ' Isso preserva corretamente os acentos (Razão, Nº doc., Dt.Lçto etc.)
    Dim stmLeitura, stmEscrita
    Set stmEscrita = CreateObject("ADODB.Stream")
    stmEscrita.Type    = 2            ' texto
    stmEscrita.Charset = "UTF-8"
    stmEscrita.Open

    ' Gravar cabeçalho
    stmEscrita.WriteText CABECALHO & vbCrLf

    Dim i, linha, cabecalhoPulo
    For i = 0 To numArquivos - 1
        Dim arq : arq = arquivos(i)
        If Not g_FSO.FileExists(arq) Then
            Log "Consolidar: arquivo não encontrado, ignorado — " & arq
        Else
            Log "Consolidar: processando " & arq
            cabecalhoPulo = False

            Set stmLeitura = CreateObject("ADODB.Stream")
            stmLeitura.Type    = 2
            stmLeitura.Charset = "windows-1252"  ' Latin-1 / ANSI (padrão SAP)
            stmLeitura.Open
            On Error Resume Next
            stmLeitura.LoadFromFile arq
            If Err.Number <> 0 Then
                Log "AVISO Consolidar: erro ao abrir " & arq & " — " & Err.Description
                Err.Clear : stmLeitura.Close
                On Error GoTo 0
            Else
                On Error GoTo 0
                Do While Not stmLeitura.EOS
                    linha = stmLeitura.ReadText(-2)  ' -2 = linha por linha (adReadLine)
                    ' Pular a primeira linha de cabeçalho do arquivo (já temos o nosso)
                    If Not cabecalhoPulo Then
                        cabecalhoPulo = True
                        ' Pular se contiver indicadores do cabeçalho SAP
                        If InStr(linha, "Atribui") > 0 Or InStr(linha, "Nr doc") > 0 Or _
                           InStr(linha, "Nº doc") > 0 Or InStr(linha, "St|") > 0 Then
                            Log "Consolidar: cabeçalho ignorado em " & arq
                        Else
                            ' Linha não é cabeçalho — gravar (e não pular mais)
                            If Trim(linha) <> "" Then stmEscrita.WriteText linha & vbCrLf
                        End If
                    Else
                        ' Filtrar linhas de separação ou cabeçalhos repetidos
                        If Not EhLinhaDescartavel(linha) Then
                            stmEscrita.WriteText linha & vbCrLf
                        End If
                    End If
                Loop
                stmLeitura.Close
            End If
            Set stmLeitura = Nothing
        End If
    Next

    ' Salvar consolidado
    On Error Resume Next
    stmEscrita.SaveToFile nomeConsolidado, 2   ' 2 = sobrescrever
    If Err.Number <> 0 Then
        Log "ERRO Consolidar: falha ao salvar " & nomeConsolidado & " — " & Err.Description
        Err.Clear
    Else
        Log "Consolidar: arquivo consolidado gerado com sucesso — " & nomeConsolidado
        Log "Consolidar: encoding do consolidado: UTF-8 (acentos preservados)."
        MsgBox "Consolidado gerado com sucesso!" & vbCrLf & nomeConsolidado, _
               vbInformation, "MUTDFC"
    End If
    stmEscrita.Close
    On Error GoTo 0
    Set stmEscrita = Nothing
End Sub

' =============================================================================
' FUNCTION EhLinhaDescartavel
' Retorna True para linhas que não devem constar no consolidado:
' cabeçalhos repetidos, linhas vazias, linhas de separação do SAP.
' =============================================================================
Function EhLinhaDescartavel(linha)
    EhLinhaDescartavel = False
    Dim l : l = Trim(linha)
    If l = "" Then EhLinhaDescartavel = True : Exit Function
    ' Cabeçalhos do SAP (contêm marcadores típicos)
    If InStr(l, "Atribui") > 0 And InStr(l, "doc") > 0 Then
        EhLinhaDescartavel = True : Exit Function
    End If
    If Left(l, 4) = "|   " And InStr(l, "Montante") > 0 Then
        EhLinhaDescartavel = True : Exit Function
    End If
    ' Linhas de separação (-----)
    If l = String(Len(l), "-") Then
        EhLinhaDescartavel = True : Exit Function
    End If
End Function

' =============================================================================
' FUNCTION ValidarData
' Verifica se a string está no formato dd.mm.aaaa e representa data válida.
' =============================================================================
Function ValidarData(s)
    ValidarData = False
    If Len(s) <> 10 Then Exit Function
    If Mid(s, 3, 1) <> "." Or Mid(s, 6, 1) <> "." Then Exit Function
    Dim dd, mm, aaaa
    dd   = CInt(Left(s, 2))
    mm   = CInt(Mid(s, 4, 2))
    aaaa = CInt(Right(s, 4))
    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Or dd > 31 Then Exit Function
    On Error Resume Next
    Dim dt : dt = DateSerial(aaaa, mm, dd)
    If Err.Number <> 0 Then Err.Clear : On Error GoTo 0 : Exit Function
    On Error GoTo 0
    If Day(dt) <> dd Or Month(dt) <> mm Or Year(dt) <> aaaa Then Exit Function
    ValidarData = True
End Function

' =============================================================================
' FUNCTION ParseData
' Converte string "dd.mm.aaaa" para tipo Date do VBScript.
' =============================================================================
Function ParseData(s)
    Dim dd, mm, aaaa
    dd   = CInt(Left(s, 2))
    mm   = CInt(Mid(s, 4, 2))
    aaaa = CInt(Right(s, 4))
    ParseData = DateSerial(aaaa, mm, dd)
End Function

' =============================================================================
' FUNCTION FormatData
' Converte tipo Date para string "dd.mm.aaaa".
' =============================================================================
Function FormatData(dt)
    FormatData = Right("0" & Day(dt), 2)   & "." & _
                 Right("0" & Month(dt), 2) & "." & _
                 Year(dt)
End Function

' =============================================================================
' SUB CriarPastaSeNecessario
' Cria a pasta (e todos os pais necessários) se não existir.
' =============================================================================
Sub CriarPastaSeNecessario(caminho)
    If Not g_FSO.FolderExists(caminho) Then
        On Error Resume Next
        g_FSO.CreateFolder caminho
        If Err.Number <> 0 Then
            ' Tentar criar hierarquia manualmente
            Dim partes, acc, p
            partes = Split(caminho, "\")
            acc = ""
            For Each p In partes
                If acc = "" Then
                    acc = p
                Else
                    acc = acc & "\" & p
                End If
                If Not g_FSO.FolderExists(acc) And acc <> "" Then
                    g_FSO.CreateFolder acc
                    Err.Clear
                End If
            Next
        End If
        On Error GoTo 0
    End If
End Sub

' =============================================================================
' SUB Log
' Acrescenta uma linha com timestamp ao arquivo de LOG.
' =============================================================================
Sub Log(mensagem)
    On Error Resume Next
    Dim f
    Set f = g_FSO.OpenTextFile(g_LogFile, 8, True)   ' 8 = append, True = criar se não existir
    If Err.Number = 0 Then
        f.WriteLine "[" & Now() & "] " & mensagem
        f.Close
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' =============================================================================
' Chamar ponto de entrada
' =============================================================================
Call Main()
