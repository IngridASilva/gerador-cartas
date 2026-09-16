Attribute VB_Name = "modUtil"
'==============================================================================
' modUtil
'
' Leitura de configuracao, acesso a planilha, arquivos e log.
' Nenhuma regra de negocio mora aqui.
'==============================================================================
Option Explicit

Private Const MESES As String = "janeiro|fevereiro|marco|abril|maio|junho|" & _
    "julho|agosto|setembro|outubro|novembro|dezembro"


'------------------------------------------------------------------------------
' Config: le pelo nome definido, nunca por endereco de celula. Inserir uma
' linha em Config nao pode quebrar o codigo.
'------------------------------------------------------------------------------
Public Function LerConfig() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Config")
    Dim ultima As Long, i As Long, nome As String

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultima
        nome = Trim$(CStr(ws.Cells(i, 3).Value))
        If Len(nome) > 0 Then d(nome) = ws.Cells(i, 2).Value
    Next i

    Set LerConfig = d
End Function


'------------------------------------------------------------------------------
' Modelos. Item: Array(arquivo, descricao, padrao_nome_arquivo)
'------------------------------------------------------------------------------
Public Function LerModelos() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Modelos")
    Dim ultima As Long, i As Long, cod As String

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultima
        cod = UCase$(Trim$(CStr(ws.Cells(i, 1).Value)))
        If Len(cod) > 0 Then
            d(cod) = Array(CStr(ws.Cells(i, 2).Value), _
                           CStr(ws.Cells(i, 3).Value), _
                           CStr(ws.Cells(i, 4).Value))
        End If
    Next i

    Set LerModelos = d
End Function


'------------------------------------------------------------------------------
' Cabecalho da aba -> indice da coluna. E o que permite reordenar colunas na
' planilha sem quebrar nada.
'------------------------------------------------------------------------------
Public Function MapearColunas(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ultima As Long, c As Long, nome As String

    ultima = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ultima
        nome = LCase$(Trim$(CStr(ws.Cells(1, c).Value)))
        If Len(nome) > 0 Then d(nome) = c
    Next c

    Set MapearColunas = d
End Function


Public Function LerIntervalo(ws As Worksheet) As Variant
    Dim ultimaLinha As Long, ultimaColuna As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ultimaLinha < 2 Then
        LerIntervalo = Empty
    ElseIf ultimaLinha = 2 Then
        Dim uma() As Variant, c As Long
        ReDim uma(1 To 1, 1 To ultimaColuna)
        For c = 1 To ultimaColuna
            uma(1, c) = ws.Cells(2, c).Value
        Next c
        LerIntervalo = uma
    Else
        LerIntervalo = ws.Range(ws.Cells(2, 1), _
                                ws.Cells(ultimaLinha, ultimaColuna)).Value
    End If
End Function


'------------------------------------------------------------------------------
' Texto e datas
'------------------------------------------------------------------------------
Public Function AplicarMarcadores(texto As String, marcadores As Object) As String
    Dim r As String: r = texto
    Dim chave As Variant
    For Each chave In marcadores.Keys
        r = Replace(r, "{{" & UCase$(CStr(chave)) & "}}", CStr(marcadores(chave)))
    Next chave
    AplicarMarcadores = r
End Function


Public Function DataPorExtenso(d As Date) As String
    Dim partes As Variant
    partes = Split(MESES, "|")
    DataPorExtenso = Day(d) & " de " & partes(Month(d) - 1) & " de " & Year(d)
End Function


Public Function NomeSeguro(nome As String) As String
    Dim proibidos As Variant, i As Long, r As String
    r = nome
    proibidos = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(proibidos) To UBound(proibidos)
        r = Replace(r, CStr(proibidos(i)), "-")
    Next i
    NomeSeguro = Trim$(r)
End Function


Public Function GerarReferencia(cfg As Object, modelo As String, _
                                matricula As String) As String
    GerarReferencia = UCase$(Left$(modelo, 3)) & "-" & _
                      CStr(cfg("CICLO")) & "-" & matricula & "-" & _
                      Format(Now, "yyyymmddhhnn")
End Function


'------------------------------------------------------------------------------
' Verificador de conteudo (FNV-1a, 32 bits).
'
' NAO e hash criptografico e nao deve ser apresentado como tal. Serve para uma
' coisa so: provar, meses depois, que o PDF em maos e o mesmo que o log
' registrou. Para isso e suficiente, roda em VBA puro e nao depende de
' biblioteca externa nem de politica de TI.
'------------------------------------------------------------------------------
Public Function Verificador(texto As String) As String
    Const OFFSET As Double = 2166136261#
    Const PRIMO As Double = 16777619
    Const M32 As Double = 4294967296#

    Dim h As Double: h = OFFSET
    Dim i As Long, limite As Long
    limite = Len(texto)
    If limite > 20000 Then limite = 20000   ' basta para caracterizar a carta

    For i = 1 To limite
        h = h Xor Asc(Mid$(texto, i, 1))
        h = h * PRIMO
        h = h - Int(h / M32) * M32          ' modulo 2^32 sem estourar Long
    Next i

    Verificador = Right$("00000000" & Hex$(CDbl(h)), 8)
End Function


'------------------------------------------------------------------------------
' Arquivos
'------------------------------------------------------------------------------
Public Function Combinar(pasta As String, arquivo As String) As String
    Dim p As String: p = Trim$(pasta)
    If Right$(p, 1) <> "\" Then p = p & "\"
    Combinar = p & Trim$(arquivo)
End Function


Public Function ArquivoExiste(caminho As String) As Boolean
    On Error Resume Next
    ArquivoExiste = (Len(Dir$(caminho)) > 0)
    On Error GoTo 0
End Function


Public Function GarantirPasta(caminho As String) As Boolean
    On Error GoTo Falhou
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(caminho) Then fso.CreateFolder caminho
    GarantirPasta = fso.FolderExists(caminho)
    Exit Function
Falhou:
    GarantirPasta = False
End Function


Public Function Menor(a As Long, b As Long) As Long
    Menor = IIf(a < b, a, b)
End Function


'------------------------------------------------------------------------------
' Log
'------------------------------------------------------------------------------
Public Sub RegistrarLog(referencia As String, matricula As String, _
                        nome As String, modelo As String, arquivo As String, _
                        verificacao As String, status As String, _
                        mensagem As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Log")
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(r, 2).Value = referencia
    ws.Cells(r, 3).Value = Val(matricula)
    ws.Cells(r, 4).Value = nome
    ws.Cells(r, 5).Value = modelo
    ws.Cells(r, 6).Value = arquivo
    ws.Cells(r, 7).Value = verificacao
    ws.Cells(r, 8).Value = status
    ws.Cells(r, 9).Value = mensagem

    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)).Font
        .Name = "Arial"
        .Size = 10
    End With
    If status <> "OK" Then ws.Cells(r, 8).Font.Color = RGB(192, 0, 0)
End Sub
