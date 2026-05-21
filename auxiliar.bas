Attribute VB_Name = "auxiliar"
' Retorna o próximo ID
Public Function ObterProximoId(ws As Worksheet) As Long
    Dim maxId As Variant
    On Error Resume Next
    maxId = Application.WorksheetFunction.Max(ws.Range("A:A"))
    On Error GoTo 0
    If IsError(maxId) Or IsNull(maxId) Then
        ObterProximoId = 1
    ElseIf IsNumeric(maxId) Then
        If CLng(maxId) < 1 Then
            ObterProximoId = 1
        Else
            ObterProximoId = CLng(maxId) + 1
        End If
    Else
        ObterProximoId = 1
    End If
End Function

'Função status de retorno ao estoque
Public Function RetornaAoEstoque(ByVal texto As String) As Boolean
    Dim status As String
    status = Trim$(texto)
    
    Select Case status
        Case "UTILIZAVEL", "UTILIZÁVEL"
            RetornaAoEstoque = True
        Case "RETORNO ESTOQUE"
            RetornaAoEstoque = True
        Case "TESTE", "ENSAIO"
            RetornaAoEstoque = True
        Case Else
            RetornaAoEstoque = False
    End Select
End Function

'Obrigatoriedade Vencimento/CA
Public Function VerificarObrigatoriedade(codigo As String, _
                                   ByRef exigeVencimento As Boolean, _
                                   ByRef exigeCA As Boolean, _
                                   ByRef exigeSerie As Boolean)

    Dim ws As Worksheet
    Dim rngCod As Range
    Dim linha As ListRow
    Dim tbl As ListObject

    Set ws = ThisWorkbook.Worksheets("Auxiliar")
    Set tbl = ws.ListObjects("tab_codigo")

    Set rngCod = tbl.ListColumns("COD").DataBodyRange.Find( _
                    What:=Trim(codigo), LookAt:=xlWhole, LookIn:=xlValues)

    If rngCod Is Nothing Then
        exigeVencimento = False
        exigeCA = False
        exigeSerie = False
        Exit Function
    End If
    
    Set linha = tbl.ListRows(rngCod.Row - tbl.DataBodyRange.Row + 1)

    exigeVencimento = (UCase(Trim(linha.Range(1, tbl.ListColumns("VALIDADE").Index).Value)) = "SIM")
    exigeCA = (UCase(Trim(linha.Range(1, tbl.ListColumns("CA").Index).Value)) = "SIM")
    exigeSerie = (UCase(Trim(linha.Range(1, tbl.ListColumns("SERIE").Index).Value)) = "SIM")

End Function

'Retornar Cargo
Public Function GetFuncao(matr As Variant) As String
    On Error GoTo erro
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngMatr As Range
    Dim linha As ListRow

    Set ws = ThisWorkbook.Worksheets("Auxiliar")
    Set tbl = ws.ListObjects("tab_funci")

    Set rngMatr = tbl.ListColumns(1).DataBodyRange.Find( _
                    What:=Trim(matr), LookAt:=xlWhole, LookIn:=xlValues)

    If rngMatr Is Nothing Then
        GetFuncao = ""
        Exit Function
    End If

    Set linha = tbl.ListRows(rngMatr.Row - tbl.DataBodyRange.Row + 1)

    GetFuncao = linha.Range(1, 3).Value
    Exit Function

erro:
    GetFuncao = ""
End Function

'Obter ID para Edição
Public Function GetRowById(tbl As ListObject, ByVal id As Variant) As Range
    Dim lr As ListRow
    Dim idBusca As String
    
    idBusca = Trim(CStr(id))
    Set GetRowById = Nothing
    
    If tbl.DataBodyRange Is Nothing Then Exit Function
    
    For Each lr In tbl.ListRows
        If Trim(CStr(lr.Range.Cells(1, 1).Value)) = idBusca Then
            Set GetRowById = lr.Range.Cells(1, 1)
            Exit Function
        End If
    Next lr
End Function

'Calcular saldo na edição
Public Function CalcularSaldo(ByVal dataCorte As Date, _
                              ByVal dataAntiga As Date, _
                              ByVal dataNova As Date, _
                              ByVal qtdAntiga As Double, _
                              ByVal qtdNova As Double, _
                              ByVal SaldoAtual As Double, _
                              ByVal tipoMov As String) _
                              As Double
    
If dataAntiga <= dataCorte And dataNova <= dataCorte Then

    CalcularSaldo = qtdNova
    Exit Function
ElseIf dataNova <= dataCorte And dataAntiga > dataCorte Then

    If tipoMov = "ENTRADA" Then
        SaldoAtual = SaldoAtual - qtdAntiga
        CalcularSaldo = SaldoAtual
    Else
        SaldoAtual = SaldoAtual + qtdAntiga
        CalcularSaldo = SaldoAtual
    End If
    Exit Function
ElseIf dataNova > dataCorte And dataAntiga <= dataCorte Then

    If tipoMov = "ENTRADA" Then
        SaldoAtual = SaldoAtual + qtdNova
        CalcularSaldo = SaldoAtual
    Else
        SaldoAtual = SaldoAtual - qtdNova
        CalcularSaldo = SaldoAtual
    End If
    Exit Function
ElseIf dataNova > dataCorte And dataAntiga > dataCorte Then
        
    If tipoMov = "ENTRADA" Then
        SaldoAtual = SaldoAtual - qtdAntiga + qtdNova
        CalcularSaldo = SaldoAtual
    Else
        SaldoAtual = SaldoAtual + qtdAntiga - qtdNova
        CalcularSaldo = SaldoAtual
    End If
    Exit Function
End If

End Function

'Atualizar valores unitários
Sub atualizarValoresUnitarios()

    Dim wsLcto As Worksheet
    Dim wsAux As Worksheet
    Dim dict As Object
    Dim i As Long
    Dim ultimoLcto As Long
    Dim ultimoCod As Long
    Dim codigo As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Set wsLcto = ThisWorkbook.Sheets("Lançamentos")
    Set wsAux = ThisWorkbook.Sheets("Auxiliar")
    
    ' Últimas linhaS
    ultimoLcto = wsLcto.Cells(wsLcto.Rows.Count, "A").End(xlUp).Row
    ultimoCod = wsAux.Cells(wsAux.Rows.Count, "Z").End(xlUp).Row
    
    ' Criar Dictionary (Código , R$ Unitário)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To ultimoCod
        If Not dict.exists(CStr(wsAux.Cells(i, 26).Value)) Then
            dict.Add CStr(wsAux.Cells(i, 26).Value), wsAux.Cells(i, 35).Value
        End If
    Next i
    
    ' Percorrer Lançamentos
    For i = 3 To ultimoLcto
        
        If wsLcto.Cells(i, 5).Value = "SAÍDA" Then
            
            codigo = CStr(wsLcto.Cells(i, 8).Value)
            
            If dict.exists(codigo) Then
                wsLcto.Cells(i, 11).Value = dict(codigo)
            End If
            
        End If
        
    Next i
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "Valores unitários atualizados", vbInformation, "R$ Unitário"

End Sub

