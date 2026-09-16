Attribute VB_Name = "modValidacao"
'==============================================================================
' modValidacao
'
' Confere a lista antes de gerar qualquer coisa.
'
' Mala direta comum nao valida nada: gera as 300 cartas e voce descobre o
' problema quando alguem responde o e-mail. A validacao previa e a diferenca
' entre um script e uma ferramenta, e custa menos de cem linhas.
'
' Severidade Erro bloqueia a geracao. Aviso deixa passar e fica registrado.
'==============================================================================
Option Explicit

Public Function ValidarDados(Optional silencioso As Boolean = False) As Long
    Dim wsD As Worksheet, wsV As Worksheet
    Set wsD = ThisWorkbook.Worksheets("Dados")
    Set wsV = ThisWorkbook.Worksheets("Validacao")

    LimparValidacao wsV

    Dim dados As Variant, colunas As Object
    dados = modUtil.LerIntervalo(wsD)
    Set colunas = modUtil.MapearColunas(wsD)

    If IsEmpty(dados) Then
        Registrar wsV, 0, 0, "Erro", "Dados", "A aba Dados esta vazia."
        ValidarDados = 1
        If Not silencioso Then Concluir 1, 0
        Exit Function
    End If

    Dim cfg As Object:     Set cfg = modUtil.LerConfig()
    Dim modelos As Object: Set modelos = modUtil.LerModelos()
    Dim vistas As Object:  Set vistas = CreateObject("Scripting.Dictionary")

    Dim erros As Long, avisos As Long
    Dim i As Long, linha As Long, mat As String

    For i = 1 To UBound(dados, 1)
        linha = i + 1
        mat = CStr(Valor(dados, i, colunas, "matricula"))

        ' --- obrigatorios ---------------------------------------------------
        Dim obrigatorios As Variant, k As Long
        obrigatorios = Array("matricula", "nome", "tratamento", "modelo")
        For k = LBound(obrigatorios) To UBound(obrigatorios)
            If Len(Trim$(Valor(dados, i, colunas, CStr(obrigatorios(k))))) = 0 Then
                Registrar wsV, linha, mat, "Erro", CStr(obrigatorios(k)), _
                          "Campo obrigatorio vazio."
                erros = erros + 1
            End If
        Next k

        ' --- duplicidade ----------------------------------------------------
        If Len(mat) > 0 Then
            If vistas.Exists(mat) Then
                Registrar wsV, linha, mat, "Erro", "matricula", _
                    "Matricula repetida (ja aparece na linha " & vistas(mat) & ")."
                erros = erros + 1
            Else
                vistas(mat) = linha
            End If
        End If

        ' --- modelo cadastrado ----------------------------------------------
        Dim cod As String: cod = UCase$(Trim$(Valor(dados, i, colunas, "modelo")))
        If Len(cod) > 0 Then
            If Not modelos.Exists(cod) Then
                Registrar wsV, linha, mat, "Erro", "modelo", _
                    "Modelo '" & cod & "' nao cadastrado na aba Modelos."
                erros = erros + 1
            ElseIf Not modUtil.ArquivoExiste( _
                    modUtil.Combinar(CStr(cfg("PASTA_MODELOS")), modelos(cod)(0))) Then
                Registrar wsV, linha, mat, "Erro", "modelo", _
                    "Arquivo do modelo nao encontrado: " & modelos(cod)(0)
                erros = erros + 1
            End If
        End If

        ' --- e-mail -----------------------------------------------------------
        Dim email As String: email = Trim$(Valor(dados, i, colunas, "email"))
        If Len(email) > 0 Then
            If InStr(email, "@") = 0 Or InStr(email, ".") = 0 _
               Or InStr(email, " ") > 0 Then
                Registrar wsV, linha, mat, "Erro", "email", _
                    "Formato de e-mail invalido: " & email
                erros = erros + 1
            End If
        Else
            Registrar wsV, linha, mat, "Aviso", "email", _
                "Sem e-mail. A carta sera gerada, mas nao dara para enviar."
            avisos = avisos + 1
        End If

        ' --- coerencia de remuneracao ----------------------------------------
        Dim atual As Double, novo As Double, pct As Double, esperado As Double
        atual = Numero(Valor(dados, i, colunas, "salario_atual"))
        novo = Numero(Valor(dados, i, colunas, "salario_novo"))
        pct = Numero(Valor(dados, i, colunas, "pct_aumento"))

        If atual > 0 And novo > 0 Then
            If novo <= atual Then
                Registrar wsV, linha, mat, "Erro", "salario_novo", _
                    "Novo salario menor ou igual ao atual em carta de aumento."
                erros = erros + 1
            Else
                esperado = novo / atual - 1
                ' Divergencia entre o percentual informado e o que os valores
                ' dizem costuma ser copia e cola de outra linha. Quem confere
                ' a carta pronta olha o percentual, nao a conta.
                If pct > 0 And Abs(esperado - pct) > 0.002 Then
                    Registrar wsV, linha, mat, "Erro", "pct_aumento", _
                        "Percentual informado (" & Format(pct, "0.00%") & _
                        ") nao confere com os valores (" & _
                        Format(esperado, "0.00%") & ")."
                    erros = erros + 1
                End If
                If esperado > CDbl(cfg("LIMITE_PCT")) Then
                    Registrar wsV, linha, mat, "Aviso", "pct_aumento", _
                        "Aumento de " & Format(esperado, "0.0%") & _
                        " acima do limite de alerta. Confirme a aprovacao."
                    avisos = avisos + 1
                End If
            End If
        ElseIf cod = "MERITO" Or cod = "PROMOCAO" Then
            Registrar wsV, linha, mat, "Erro", "salario", _
                "Carta de remuneracao sem salario atual ou novo preenchido."
            erros = erros + 1
        End If

        ' --- data de efeito ---------------------------------------------------
        Dim dt As Variant: dt = Valor(dados, i, colunas, "data_efeito")
        If Len(Trim$(CStr(dt))) > 0 Then
            If Not IsDate(dt) Then
                Registrar wsV, linha, mat, "Erro", "data_efeito", _
                    "Data invalida: " & CStr(dt)
                erros = erros + 1
            ElseIf CDate(dt) < Date - 180 Then
                Registrar wsV, linha, mat, "Aviso", "data_efeito", _
                    "Data de efeito com mais de 6 meses. Confirme se e retroativo."
                avisos = avisos + 1
            End If
        End If
    Next i

    ' --- marcadores do modelo x colunas disponiveis --------------------------
    ValidarMarcadores wsV, cfg, modelos, colunas, erros

    ValidarDados = erros
    If Not silencioso Then Concluir erros, avisos
End Function


'------------------------------------------------------------------------------
' Confere se todo marcador presente no .docx tem origem em Config ou em Dados.
' Sem isso a carta sai com "{{GESTOR}}" impresso, e alguem so percebe depois
' de assinada.
'------------------------------------------------------------------------------
Private Sub ValidarMarcadores(wsV As Worksheet, cfg As Object, _
                              modelos As Object, colunas As Object, _
                              ByRef erros As Long)
    Dim disponiveis As Object: Set disponiveis = CreateObject("Scripting.Dictionary")
    Dim chave As Variant

    For Each chave In cfg.Keys
        disponiveis(UCase$(CStr(chave))) = True
    Next chave
    For Each chave In colunas.Keys
        disponiveis(UCase$(CStr(chave))) = True
    Next chave
    For Each chave In Array("DATA_CARTA", "DATA_GERACAO", "REFERENCIA")
        disponiveis(CStr(chave)) = True
    Next chave

    Dim wd As Object, doc As Object
    On Error Resume Next
    Set wd = CreateObject("Word.Application")
    On Error GoTo 0
    If wd Is Nothing Then
        Registrar wsV, 0, 0, "Aviso", "modelos", _
            "Word indisponivel. Marcadores nao foram conferidos."
        Exit Sub
    End If
    wd.Visible = False

    Dim cod As Variant, caminho As String, texto As String
    Dim p1 As Long, p2 As Long, marcador As String

    For Each cod In modelos.Keys
        caminho = modUtil.Combinar(CStr(cfg("PASTA_MODELOS")), modelos(cod)(0))
        If modUtil.ArquivoExiste(caminho) Then
            Set doc = wd.Documents.Open(caminho, ReadOnly:=True, Visible:=False)
            texto = doc.Content.Text
            doc.Close SaveChanges:=False

            p1 = InStr(1, texto, "{{")
            Do While p1 > 0
                p2 = InStr(p1, texto, "}}")
                If p2 = 0 Then Exit Do
                marcador = UCase$(Mid$(texto, p1 + 2, p2 - p1 - 2))
                If Not disponiveis.Exists(marcador) Then
                    Registrar wsV, 0, 0, "Erro", CStr(cod), _
                        "O modelo usa {{" & marcador & "}}, que nao existe em " & _
                        "Config nem como coluna da aba Dados."
                    erros = erros + 1
                    disponiveis(marcador) = True  ' reporta uma vez so
                End If
                p1 = InStr(p2, texto, "{{")
            Loop
        End If
    Next cod

    wd.Quit
    Set wd = Nothing
End Sub


'------------------------------------------------------------------------------
' Auxiliares
'------------------------------------------------------------------------------
Private Function Valor(dados As Variant, i As Long, colunas As Object, _
                       campo As String) As String
    If colunas.Exists(campo) Then
        Valor = Trim$(CStr(dados(i, colunas(campo))))
    Else
        Valor = ""
    End If
End Function


Private Function Numero(texto As String) As Double
    On Error Resume Next
    Numero = CDbl(texto)
    On Error GoTo 0
End Function


Private Sub LimparValidacao(wsV As Worksheet)
    Dim ultima As Long
    ultima = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row
    If ultima > 1 Then wsV.Range("A2:E" & ultima).ClearContents
End Sub


Private Sub Registrar(wsV As Worksheet, linha As Long, matricula As String, _
                      severidade As String, campo As String, mensagem As String)
    Dim r As Long
    r = wsV.Cells(wsV.Rows.Count, 1).End(xlUp).Row + 1
    If linha > 0 Then wsV.Cells(r, 1).Value = linha
    If Len(matricula) > 0 Then wsV.Cells(r, 2).Value = Val(matricula)
    wsV.Cells(r, 3).Value = severidade
    wsV.Cells(r, 4).Value = campo
    wsV.Cells(r, 5).Value = mensagem
    wsV.Range(wsV.Cells(r, 1), wsV.Cells(r, 5)).Font.Name = "Arial"
    wsV.Range(wsV.Cells(r, 1), wsV.Cells(r, 5)).Font.Size = 10
    If severidade = "Erro" Then
        wsV.Cells(r, 3).Font.Color = RGB(192, 0, 0)
        wsV.Cells(r, 3).Font.Bold = True
    Else
        wsV.Cells(r, 3).Font.Color = RGB(191, 143, 0)
    End If
End Sub


Private Sub Concluir(erros As Long, avisos As Long)
    ThisWorkbook.Worksheets("Validacao").Activate
    If erros = 0 And avisos = 0 Then
        MsgBox "Nenhum problema encontrado. Pode gerar.", vbInformation, "Validacao"
    ElseIf erros = 0 Then
        MsgBox avisos & " aviso(s), nenhum erro." & vbCrLf & vbCrLf & _
               "Avisos nao bloqueiam a geracao, mas vale conferir antes.", _
               vbInformation, "Validacao"
    Else
        MsgBox erros & " erro(s) e " & avisos & " aviso(s)." & vbCrLf & vbCrLf & _
               "A geracao fica bloqueada ate os erros serem corrigidos.", _
               vbExclamation, "Validacao"
    End If
End Sub
