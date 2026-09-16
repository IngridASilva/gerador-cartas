Attribute VB_Name = "modGerador"
'==============================================================================
' modGerador
'
' Ponto de entrada: GerarCartas. Depois: PrepararEmails.
'
' Uma unica instancia do Word aberta para o lote inteiro. Abrir e fechar o
' Word a cada carta multiplica o tempo por dez e e o erro mais comum em codigo
' de mala direta caseira.
'==============================================================================
Option Explicit

Public Const VERSAO As String = "1.0.0"
Private Const wdExportFormatPDF As Long = 17
Private Const wdFindStop As Long = 0
Private Const wdReplaceAll As Long = 2


Public Sub GerarCartas()
    Dim inicio As Single: inicio = Timer

    ' Validar sempre, mesmo que a pessoa ja tenha rodado a validacao antes.
    ' Entre uma coisa e outra alguem pode ter editado a planilha.
    If modValidacao.ValidarDados(silencioso:=True) > 0 Then
        ThisWorkbook.Worksheets("Validacao").Activate
        MsgBox "Ha erros na aba Validacao. Corrija antes de gerar.", _
               vbExclamation, "Gerador"
        Exit Sub
    End If

    Dim cfg As Object:     Set cfg = modUtil.LerConfig()
    Dim modelos As Object: Set modelos = modUtil.LerModelos()
    Dim wsD As Worksheet:  Set wsD = ThisWorkbook.Worksheets("Dados")

    Dim dados As Variant, colunas As Object
    dados = modUtil.LerIntervalo(wsD)
    Set colunas = modUtil.MapearColunas(wsD)

    Dim limite As Long
    limite = UBound(dados, 1)
    If UCase$(Trim$(CStr(cfg("MODO_TESTE")))) = "SIM" Then
        limite = modUtil.Menor(limite, CLng(cfg("QTD_TESTE")))
    End If

    If Not modUtil.GarantirPasta(CStr(cfg("PASTA_SAIDA"))) Then
        MsgBox "Nao foi possivel criar a pasta de saida:" & vbCrLf & _
               cfg("PASTA_SAIDA"), vbCritical, "Gerador"
        Exit Sub
    End If

    Dim wd As Object
    On Error Resume Next
    Set wd = CreateObject("Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then
        MsgBox "Word nao encontrado nesta maquina.", vbCritical, "Gerador"
        Exit Sub
    End If
    wd.Visible = False
    wd.DisplayAlerts = False

    Application.ScreenUpdating = False
    On Error GoTo TratarErro

    Dim i As Long, gerados As Long, falhas As Long
    For i = 1 To limite
        Application.StatusBar = "Gerando carta " & i & " de " & limite & "..."
        If GerarUma(wd, cfg, modelos, dados, colunas, i) Then
            gerados = gerados + 1
        Else
            falhas = falhas + 1
        End If
    Next i

    wd.Quit
    Set wd = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True

    ThisWorkbook.Worksheets("Log").Activate
    MsgBox gerados & " carta(s) gerada(s) em " & Format(Timer - inicio, "0.0") & _
           " segundo(s)." & _
           IIf(falhas > 0, vbCrLf & falhas & " falha(s), detalhe no Log.", "") & _
           IIf(UCase$(Trim$(CStr(cfg("MODO_TESTE")))) = "SIM", _
               vbCrLf & vbCrLf & "MODO TESTE ativo: apenas as " & limite & _
               " primeiras. Troque para NAO em Config para gerar o lote todo.", ""), _
           vbInformation, "Gerador"
    Exit Sub

TratarErro:
    On Error Resume Next
    wd.Quit
    Set wd = Nothing
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Gerador"
End Sub


'------------------------------------------------------------------------------
' Gera uma carta. Devolve True em caso de sucesso.
'------------------------------------------------------------------------------
Private Function GerarUma(wd As Object, cfg As Object, modelos As Object, _
                          dados As Variant, colunas As Object, _
                          i As Long) As Boolean
    Dim doc As Object
    Dim cod As String, caminhoModelo As String, destino As String
    Dim referencia As String, matricula As String, nome As String
    Dim marcadores As Object

    matricula = CStr(dados(i, colunas("matricula")))
    nome = CStr(dados(i, colunas("nome")))
    cod = UCase$(Trim$(CStr(dados(i, colunas("modelo")))))
    referencia = modUtil.GerarReferencia(cfg, cod, matricula)

    Set marcadores = MontarMarcadores(cfg, dados, colunas, i, referencia)
    caminhoModelo = modUtil.Combinar(CStr(cfg("PASTA_MODELOS")), modelos(cod)(0))

    On Error GoTo Falhou
    Set doc = wd.Documents.Open(caminhoModelo, ReadOnly:=True, Visible:=False)

    Dim chave As Variant
    For Each chave In marcadores.Keys
        SubstituirEmTudo doc, "{{" & CStr(chave) & "}}", CStr(marcadores(chave))
    Next chave

    Dim nomeArquivo As String
    nomeArquivo = modUtil.AplicarMarcadores(CStr(modelos(cod)(2)), marcadores)
    nomeArquivo = modUtil.NomeSeguro(nomeArquivo)

    If UCase$(Trim$(CStr(cfg("GERAR_PDF")))) = "SIM" Then
        destino = modUtil.Combinar(CStr(cfg("PASTA_SAIDA")), nomeArquivo & ".pdf")
        doc.ExportAsFixedFormat OutputFileName:=destino, _
                                ExportFormat:=wdExportFormatPDF
    Else
        destino = modUtil.Combinar(CStr(cfg("PASTA_SAIDA")), nomeArquivo & ".docx")
        doc.SaveAs2 destino
    End If

    ' Verificacao do conteudo efetivamente impresso. Nao e criptografica: serve
    ' para provar, meses depois, que o arquivo entregue e o mesmo que o log
    ' registrou. Numa discussao sobre "nao foi isso que me mandaram", e a
    ' unica evidencia que existe.
    Dim verificacao As String
    verificacao = modUtil.Verificador(doc.Content.Text)

    doc.Close SaveChanges:=False
    Set doc = Nothing

    modUtil.RegistrarLog referencia, matricula, nome, cod, destino, _
                         verificacao, "OK", ""
    GerarUma = True
    Exit Function

Falhou:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    modUtil.RegistrarLog referencia, matricula, nome, cod, "", "", "FALHA", _
                         "Erro " & Err.Number & ": " & Err.Description
    GerarUma = False
End Function


'------------------------------------------------------------------------------
' Monta o dicionario de marcadores da linha.
'
' Origem tripla: parametros de Config, colunas da aba Dados e tres campos
' calculados. Como as colunas viram marcadores automaticamente, acrescentar
' uma coluna na planilha cria um marcador novo sem tocar no codigo.
'------------------------------------------------------------------------------
Private Function MontarMarcadores(cfg As Object, dados As Variant, _
                                  colunas As Object, i As Long, _
                                  referencia As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim chave As Variant

    For Each chave In cfg.Keys
        d(UCase$(CStr(chave))) = CStr(cfg(chave))
    Next chave

    For Each chave In colunas.Keys
        d(UCase$(CStr(chave))) = Formatar(dados(i, colunas(chave)), CStr(chave))
    Next chave

    d("DATA_CARTA") = modUtil.DataPorExtenso(Date)
    d("DATA_GERACAO") = Format(Now, "dd/mm/yyyy hh:mm")
    d("REFERENCIA") = referencia

    ' Marcadores de Config podem conter outros marcadores (assunto de e-mail,
    ' por exemplo). Uma passada de resolucao resolve o caso comum.
    For Each chave In d.Keys
        If InStr(CStr(d(chave)), "{{") > 0 Then
            d(chave) = modUtil.AplicarMarcadores(CStr(d(chave)), d)
        End If
    Next chave

    Set MontarMarcadores = d
End Function


'------------------------------------------------------------------------------
' Formatacao por convencao de nome de coluna. Salario sai como moeda, data
' como dd/mm/aaaa, percentual como 0,0%. Sem isso a carta imprime
' "8734.5599999" e "01/04/2026 00:00:00".
'------------------------------------------------------------------------------
Private Function Formatar(valor As Variant, nomeColuna As String) As String
    Dim n As String: n = LCase$(nomeColuna)

    If IsEmpty(valor) Or IsNull(valor) Then
        Formatar = ""
    ElseIf InStr(n, "salario") > 0 Or InStr(n, "valor") > 0 Then
        Formatar = Format(valor, "R$ #,##0.00")
    ElseIf InStr(n, "pct") > 0 Or InStr(n, "percentual") > 0 Then
        Formatar = Format(valor, "0.0%")
    ElseIf InStr(n, "data") > 0 And IsDate(valor) Then
        Formatar = Format(valor, "dd/mm/yyyy")
    Else
        Formatar = CStr(valor)
    End If
End Function


'------------------------------------------------------------------------------
' Substitui em todas as historias do documento.
'
' doc.Content cobre o corpo e as tabelas, mas nao cabecalho nem rodape. Faltar
' com isso deixa "{{EMPRESA}}" impresso no timbre de todas as cartas.
'------------------------------------------------------------------------------
Private Sub SubstituirEmTudo(doc As Object, procurar As String, _
                             substituir As String)
    Substituir1 doc.Content, procurar, substituir

    Dim secao As Object, cab As Object, rod As Object
    For Each secao In doc.Sections
        For Each cab In secao.Headers
            If cab.Exists Then Substituir1 cab.Range, procurar, substituir
        Next cab
        For Each rod In secao.Footers
            If rod.Exists Then Substituir1 rod.Range, procurar, substituir
        Next rod
    Next secao
End Sub


Private Sub Substituir1(rng As Object, procurar As String, substituir As String)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = procurar
        .Replacement.Text = substituir
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub


'==============================================================================
' PrepararEmails
'
' Cria rascunhos no Outlook com o PDF anexado e PARA. Nada e enviado.
' Envio automatico de comunicado de remuneracao e o tipo de automacao que so
' da errado uma vez, e o erro nao tem desfazer.
'==============================================================================
Public Sub PrepararEmails()
    Dim cfg As Object: Set cfg = modUtil.LerConfig()
    Dim wsL As Worksheet: Set wsL = ThisWorkbook.Worksheets("Log")
    Dim wsD As Worksheet: Set wsD = ThisWorkbook.Worksheets("Dados")

    Dim ultima As Long
    ultima = wsL.Cells(wsL.Rows.Count, 1).End(xlUp).Row
    If ultima < 2 Then
        MsgBox "Nenhuma carta gerada ainda. Rode GerarCartas primeiro.", _
               vbExclamation, "Gerador"
        Exit Sub
    End If

    If MsgBox("Serao criados rascunhos no Outlook, sem enviar." & vbCrLf & _
              "Voce confere e envia manualmente." & vbCrLf & vbCrLf & _
              "Continuar?", vbYesNo + vbQuestion, "Gerador") = vbNo Then Exit Sub

    Dim ol As Object
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then
        MsgBox "Outlook nao encontrado nesta maquina.", vbCritical, "Gerador"
        Exit Sub
    End If

    ' Indexa e-mail e tratamento por matricula.
    Dim emails As Object: Set emails = CreateObject("Scripting.Dictionary")
    Dim dados As Variant, colunas As Object
    dados = modUtil.LerIntervalo(wsD)
    Set colunas = modUtil.MapearColunas(wsD)

    Dim i As Long
    For i = 1 To UBound(dados, 1)
        emails(CStr(dados(i, colunas("matricula")))) = _
            Array(CStr(dados(i, colunas("email"))), _
                  CStr(dados(i, colunas("tratamento"))), _
                  CStr(dados(i, colunas("nome"))))
    Next i

    Dim criados As Long, pulados As Long
    Dim mat As String, arquivo As String, info As Variant, msg As Object

    For i = 2 To ultima
        If CStr(wsL.Cells(i, 8).Value) = "OK" Then
            mat = CStr(wsL.Cells(i, 3).Value)
            arquivo = CStr(wsL.Cells(i, 6).Value)

            If emails.Exists(mat) And modUtil.ArquivoExiste(arquivo) Then
                info = emails(mat)
                If Len(Trim$(info(0))) > 0 Then
                    Set msg = ol.CreateItem(0)
                    msg.To = info(0)
                    msg.Subject = Replace(Replace(CStr(cfg("ASSUNTO_EMAIL")), _
                        "{{NOME}}", info(2)), "{{CICLO}}", CStr(cfg("CICLO")))
                    msg.Body = Replace(Replace(CStr(cfg("CORPO_EMAIL")), _
                        "{{NOME}}", info(2)), "{{TRATAMENTO}}", info(1))
                    msg.Attachments.Add arquivo
                    msg.Save    ' fica em Rascunhos
                    criados = criados + 1
                Else
                    pulados = pulados + 1
                End If
            Else
                pulados = pulados + 1
            End If
        End If
    Next i

    Set ol = Nothing
    MsgBox criados & " rascunho(s) criado(s) na pasta Rascunhos do Outlook." & _
           IIf(pulados > 0, vbCrLf & pulados & " pulado(s) por falta de e-mail " & _
               "ou arquivo.", "") & vbCrLf & vbCrLf & _
           "Nada foi enviado. Confira antes.", vbInformation, "Gerador"
End Sub
