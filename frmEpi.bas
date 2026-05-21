VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEpi 
   Caption         =   "Lançamentos"
   ClientHeight    =   8028
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17175
   OleObjectBlob   =   "frmEpi.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEpi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declarações de APIs do Windows
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

' Constantes para estilos de janela
Private Const GWL_STYLE = (-16)
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_SIZEBOX = &H40000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SW_MAXIMIZE = 3
Private hwndForm As Long
Private Const base_cautela = "Jardim - MS"
Private Const dataCorte As Date = "07/05/2026" 'Alterar no frmBaixaMaterial também
Private qtdOriginal As Double
Private dataOriginal As Date
Private tipoOriginal As String




'------ CONFIGURAÇÕES GERAIS DE FORMULÁRIO ------

'Carregamento do Formulario
Private Sub UserForm_Initialize()
    Dim rngCol As Range
    
    Set rngCol = Sheets("Auxiliar").ListObjects("tab_codigo").ListColumns("DESCRIÇÃO").DataBodyRange
    txtDescricao.RowSource = rngCol.Address(External:=True)
    
    Set rngCol = Sheets("Auxiliar").ListObjects("tab_funci").ListColumns("NOME").DataBodyRange
    txtNome.RowSource = rngCol.Address(External:=True)
    
    optEntrada = False
    optSaida = False

    With lstEpi
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .CheckBoxes = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "ID", 35
        .ColumnHeaders.Add , , "DATA", 55
        .ColumnHeaders.Add , , "MATR/CNPJ", 40
        .ColumnHeaders.Add , , "NOME", 110
        .ColumnHeaders.Add , , "TIPO MOV", 45
        .ColumnHeaders.Add , , "JUSTIFICATIVA", 85
        .ColumnHeaders.Add , , "CAT", 25
        .ColumnHeaders.Add , , "COD", 30
        .ColumnHeaders.Add , , "DESCRIÇÃO", 130
        .ColumnHeaders.Add , , "QTDE", 30
        .ColumnHeaders.Add , , "R$ UNIT", 60
        .ColumnHeaders.Add , , "R$ TOTAL", 60
        .ColumnHeaders.Add , , "DATA VCTO", 55
        .ColumnHeaders.Add , , "CA", 30
        .ColumnHeaders.Add , , "N° SERIE", 40
        .ColumnHeaders.Add , , "N.FISCAL", 45
        .ColumnHeaders.Add , , "DATA BAIXA", 55
        .ColumnHeaders.Add , , "MOTIVO BAIXA", 85
        .ColumnHeaders.Add , , "VIDA ÚTIL", 55
        .ColumnHeaders.Add , , "OBSERVAÇÃO", 180
    End With
    
    hwndForm = FindWindow(vbNullString, Me.Caption)

End Sub

' Controles maximizar/minimizar
Private Sub UserForm_Activate()
    Dim style As Long
    
    If hwndForm <> 0 Then
        style = GetWindowLong(hwndForm, GWL_STYLE)
        style = style Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_SIZEBOX
        SetWindowLong hwndForm, GWL_STYLE, style
        SetWindowPos hwndForm, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
        
        If IsZoomed(hwndForm) = 0 Then
            ShowWindow hwndForm, SW_MAXIMIZE
        End If
    Else
        MsgBox "Não foi possível obter o handle da janela.", vbExclamation
    End If
End Sub

' Redimensionamento do formulário
Private Sub UserForm_Resize()

    If hwndForm <> 0 Then
        If IsIconic(hwndForm) <> 0 Then Exit Sub
    End If
    
    If hwndForm = 0 Or IsZoomed(hwndForm) = 0 Then
        If Me.Width < 400 Then Me.Width = 400
        If Me.Height < 300 Then Me.Height = 300
    End If
    
    Image1.Width = Me.Width
    lblTitulo.Left = Me.Width / 2 - 100
    
    Frame1.Width = 150
    Frame2.Top = Frame1.Top + Frame1.Height + 5
    Frame2.Width = Me.Width - 25
    Frame2.Height = Me.Height - 110
    
    lstEpi.Left = 5
    lstEpi.Width = Frame2.Width - 15
    lstEpi.Height = Frame2.Height - 130
    
    cmdVale.Left = Frame2.Width - 25
    cmdMovimentacao.Left = cmdVale.Left - 35
    cmdCautela.Left = cmdMovimentacao.Left - 35
    
    txtTotal.Left = lstEpi.Width - 70
    lblTotal.Left = txtTotal.Left
    
    txtDataMovimento.Left = txtTotal.Left
    lblDataMovimento.Left = txtDataMovimento.Left
    chkCautela.Left = txtDataMovimento.Left - chkCautela.Width - 5
    
End Sub

' Classificar colunas pelo cabeçalho
Private Sub lstEpi_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstEpi
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        .Sorted = True
    End With
End Sub

'Formatar Data
Private Sub txtDataInicio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String
    Dim currentLength As Integer
    
    txtDataInicio.MaxLength = 10
    currentText = txtDataInicio.Text
    currentLength = Len(currentText)
    
    Select Case KeyAscii
        Case 8
        Case 13
            SendKeys "{TAB}"
            KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtDataInicio.SelLength = 0 Then
                txtDataInicio.Text = currentText & "/"
            ElseIf currentLength = 5 And txtDataInicio.SelLength = 0 Then
                txtDataInicio.Text = currentText & "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
    
End Sub
Private Sub txtDataInicio_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtDataInicio.Text) And txtDataInicio.Text <> "" Then
        
        MsgBox "Data Inválida"
        txtDataInicio.Text = ""
        Cancel = True
    
    End If
End Sub

'Formatar Data Vencimento
Private Sub txtVencimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String
    Dim currentLength As Integer
    
    txtVencimento.MaxLength = 10
    currentText = txtVencimento.Text
    currentLength = Len(currentText)
    
    Select Case KeyAscii
        Case 8
        Case 13
            SendKeys "{TAB}"
            KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtVencimento.SelLength = 0 Then
                txtVencimento.Text = currentText & "/"
            ElseIf currentLength = 5 And txtVencimento.SelLength = 0 Then
                txtVencimento.Text = currentText & "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
    
End Sub
Private Sub txtVencimento_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtVencimento.Text) And txtVencimento.Text <> "" Then
        
        MsgBox "Data de Vencimento Inválida"
        txtVencimento.Text = ""
        Cancel = True
    
    End If
End Sub

'Formatar Data Movimentação
Private Sub txtDataMovimento_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String
    Dim currentLength As Integer
    
    txtDataMovimento.MaxLength = 10
    currentText = txtDataMovimento.Text
    currentLength = Len(currentText)
    
    Select Case KeyAscii
        Case 8
        Case 13
            SendKeys "{TAB}"
            KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtDataMovimento.SelLength = 0 Then
                txtDataMovimento.Text = currentText & "/"
            ElseIf currentLength = 5 And txtDataMovimento.SelLength = 0 Then
                txtDataMovimento.Text = currentText & "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
    
End Sub
Private Sub txtDataMovimento_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtDataMovimento.Text) And txtDataMovimento.Text <> "" Then
        
        MsgBox "Data Inválida"
        txtDataMovimento.Text = ""
        Cancel = True
    
    End If
End Sub

'Formatar Data Baixa
Private Sub txtDataBaixa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String
    Dim currentLength As Integer
    
    txtDataBaixa.MaxLength = 10
    currentText = txtDataBaixa.Text
    currentLength = Len(currentText)
    
    Select Case KeyAscii
        Case 8
        Case 13
            SendKeys "{TAB}"
            KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtDataBaixa.SelLength = 0 Then
                txtDataBaixa.Text = currentText & "/"
            ElseIf currentLength = 5 And txtDataBaixa.SelLength = 0 Then
                txtDataBaixa.Text = currentText & "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
    
End Sub
Private Sub txtDataBaixa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtDataBaixa.Text) And txtDataBaixa.Text <> "" Then
        
        MsgBox "Data de Baixa Inválida"
        txtDataBaixa.Text = ""
        Cancel = True
    
    End If
End Sub

'Retornar Nome/Razao Social
Private Sub txtMatr_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim matr As String
    Dim cel As Range
    Dim achou As Boolean
    
    Set ws = ThisWorkbook.Worksheets("Auxiliar")
    Set tbl = ws.ListObjects("tab_funci")
    matr = Trim(Me.txtMatr.Value)
    If matr = "" Then Exit Sub

    achou = False
    
    For Each cel In tbl.ListColumns("MATR").DataBodyRange
        If Trim(cel.Value) = matr Then
        achou = True
        Me.txtNome = cel.Offset(0, 1).Value   'NOME
        Exit For
        End If
    Next cel

    If Not achou Then
        MsgBox "Matrícula/CNPJ inválido(a)!", vbExclamation
        Me.txtMatr.Value = ""
        Cancel = True
    End If
    
End Sub

'Retornar Descrição / Valor Unitário / Categoria
Private Sub txtCodigo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim codigo As String
    Dim cel As Range
    Dim achou As Boolean

    Set ws = ThisWorkbook.Worksheets("Auxiliar")
    Set tbl = ws.ListObjects("tab_codigo")
    
    codigo = Trim(Me.txtCodigo.Value)
    If codigo = "" Then Exit Sub

    achou = False
    For Each cel In tbl.ListColumns("COD").DataBodyRange
        If Trim(cel.Value) = codigo Then
            achou = True

            Me.txtDescricao.Value = cel.Offset(0, 1).Value   'DESCRIÇÃO
            Me.txtCat.Value = cel.Offset(0, 2).Value          'CATEGORIA

            If Me.optSaida.Value Then
                Me.txtRSUnitario.Value = Format(cel.Offset(0, 9).Value, "currency") ' RS UNIT
                Me.txtRSUnitario.Locked = True
            Else
                Me.txtRSUnitario.Value = ""
                Me.txtRSUnitario.Locked = False
            End If

            Exit For
        End If
    Next cel

    If Not achou Then
        MsgBox "Código inválido!", vbExclamation
        Me.txtDescricao.Value = ""
        Me.txtRSUnitario.Value = ""
        Me.txtCat.Value = ""
        txtCodigo.Value = ""
        Cancel = True
    End If

    txtQtde_Change
End Sub


'Calcular R$ Total
Private Sub txtQtde_Change()
    Dim vUnitario As Double
    Dim vQtde As Double
    
    If IsNumeric(Me.txtRSUnitario.Value) Then
        vUnitario = CDbl(Me.txtRSUnitario.Value)
    Else
        vUnitario = 0
    End If
    
    If IsNumeric(Me.txtQtde.Value) Then
        vQtde = CDbl(Me.txtQtde.Value)
    Else
        vQtde = 0
    End If
    
    Me.txtRSTotal.Value = Format(vUnitario * vQtde, "currency")
    
End Sub

'Calculr R$ Total na ENTRADA
Private Sub txtRSUnitario_Change()
    txtQtde_Change
End Sub

'Calcular SubTotal
Sub CalculaSubTotal()
    Dim linha As Long
    Dim SubTotal As Double
    Dim itemValueTotal As String
    Dim listItem As Object
    
    On Error GoTo erro
    With lstEpi
        SubTotal = 0
        
        For linha = 1 To .ListItems.Count
            Set listItem = .ListItems(linha)
            
            itemValueTotal = listItem.SubItems(11)
            If IsNumeric(itemValueTotal) Then
                SubTotal = SubTotal + CDbl(itemValueTotal)
            End If
        
        Next linha
    End With
    
    txtTotal.Value = VBA.Format(SubTotal, "Currency")

Exit Sub

erro:
    MsgBox "Erro ao calcular os totais!", vbCritical, "Totais"
End Sub



'------ IMPRIMIR FORMULÁRIOS ------

'Imprimir Movimentação
Private Sub cmdMovimentacao_Click()
    Dim ws As Worksheet
    Dim i As Long
    Dim listItem As Object
    Dim rowIndex As Long
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("frm_Movimentação")
    If ws Is Nothing Then
        MsgBox "A aba 'frm_Movimentação' não foi encontrada na planilha!", vbCritical, "Erro"
        Exit Sub
    End If
    
    If lstEpi.ListItems.Count = 0 Then
        MsgBox "Nenhum item encontrado para imprimir a movimentação.", vbExclamation, "Movimentação"
        GoTo ExitSub
    End If
    
    If Me.lstEpi.ListItems.Count > 30 Then
        MsgBox "O número de itens excede o máximo permitido de 30. A movimentação não será gerada.", vbExclamation, "Movimentação"
        Exit Sub
    End If
    
    ws.Range("A4:C4").ClearContents 'id-origem
    ws.Range("A6:F6").ClearContents 'data-destino
    ws.Range("B9:F38").ClearContents 'itens
    ws.Range("A42:F44").ClearContents 'obs
    
    ws.Range("A4").Value = lstEpi.ListItems(1).SubItems(15) 'NF
    ws.Range("A6").Value = Date
    ws.Range("D6").Value = lstEpi.ListItems(1).SubItems(3)  'DESTINO
    ws.Range("A42").Value = lstEpi.ListItems(1).SubItems(19) 'OBSERVAÇÃO

    rowIndex = 9 'inicio lista
    For i = 1 To lstEpi.ListItems.Count
        Set listItem = lstEpi.ListItems(i)
        ws.Range("B" & rowIndex).Value = listItem.SubItems(7)  'CÓDIGO
        ws.Range("C" & rowIndex).Value = listItem.SubItems(8)  'DESCRIÇÃO
        ws.Range("D" & rowIndex).Value = listItem.SubItems(9)  'QTDE
        If IsNumeric(listItem.SubItems(10)) Then
            ws.Range("E" & rowIndex).Value = CDbl(listItem.SubItems(10))
        Else
            ws.Range("E" & rowIndex).Value = 0
        End If
        
        ws.Range("F" & rowIndex).Value = CDbl(listItem.SubItems(11)) 'R$ TOTAL
        
        rowIndex = rowIndex + 1
    Next i
    
    ws.PageSetup.PrintArea = "A1:F53"
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    'ws.PrintOut 'Imprimir direto
    MsgBox "Formulário enviado para impressão. Verifique modelo gerado na aba frm_Movimentação!", vbInformation, "Movimentação"
    
ExitSub:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao preencher a aba Movimentação: " & Err.Description, vbCritical, "Erro"
    GoTo ExitSub
    
End Sub

'Imprimir Vale
Private Sub cmdVale_Click()
    Dim ws As Worksheet
    Dim i As Long
    Dim listItem As Object
    Dim rowIndex As Long
    Dim parcela As Integer
    Dim valor As Double
    Dim nome As String
    Dim matr As String
    Dim funcao As Variant
    Dim texto As String

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("frm_Vale")

    If ws Is Nothing Then
        MsgBox "A aba 'frm_Vale' não foi encontrada!", vbCritical
        Exit Sub
    End If
    
    If lstEpi.ListItems.Count = 0 Then
        MsgBox "Nenhum item encontrado para imprimir o Vale.", vbExclamation
        GoTo ExitSub
    End If
    
    'If lstEpi.ListItems.Count > 10 Then
     '   MsgBox "Máximo permitido: 10 itens.", vbExclamation
     '   Exit Sub
    'End If

    nome = lstEpi.ListItems(1).SubItems(3)
    matr = lstEpi.ListItems(1).SubItems(2)
    funcao = GetFuncao(matr)

    ws.Range("C6:J6").ClearContents
    ws.Range("L6").ClearContents
    ws.Range("D7").ClearContents
    ws.Range("A18:M500").UnMerge
    ws.Range("A18:M500").Font.Bold = False
    ws.Range("A18:M500").Borders.LineStyle = xlNone
    ws.Range("A18:M500").ClearContents
    ws.Range("A18:M500").RowHeight = 15.6
    ws.Range("A18:M500").NumberFormat = "General"
    ws.Range("A18:M500").ShrinkToFit = False
    ws.Range("A18:M500").Font.Size = 12
    ws.Range("A" & rowIndex + 1 & ":K" & rowIndex + 6).Font.Size = 12
    
    parcela = Val(InputBox("Informe a quantidade de parcelas", "Parcelamento"))
    If parcela <= 0 Then
        MsgBox "Número de parcelas inválido.", vbExclamation
        GoTo ExitSub
    End If
    
    ws.Range("C6").Value = nome
    ws.Range("L6").Value = Right(matr, 4)
    ws.Range("D7").Value = UCase(funcao)
    'ws.Range("D31").Value = parcela

    rowIndex = 18
    
    For i = 1 To lstEpi.ListItems.Count
        Set listItem = lstEpi.ListItems(i)
        If listItem.SubItems(3) = nome Then
            ws.Range("B" & rowIndex).Value = rowIndex - 17          'Nº
            ws.Range("C" & rowIndex).Value = listItem.SubItems(8)   'DESCRIÇÃO
            ws.Range("I" & rowIndex).Value = listItem.SubItems(9)   'QTDE
            If IsNumeric(listItem.SubItems(10)) Then
                ws.Range("J" & rowIndex).Value = CDbl(listItem.SubItems(10))
            Else
                ws.Range("J" & rowIndex).Value = 0
            End If
            If IsNumeric(listItem.SubItems(10)) Then
                ws.Range("K" & rowIndex).Value = CDbl(listItem.SubItems(10)) * CDbl(listItem.SubItems(9))
            Else
                ws.Range("K" & rowIndex).Value = 0
            End If
            
            ws.Range("B" & rowIndex).HorizontalAlignment = xlCenter
            ws.Range("C" & rowIndex & ":H" & rowIndex).Merge
            ws.Range("C" & rowIndex & ":H" & rowIndex).HorizontalAlignment = xlLeft
            ws.Range("K" & rowIndex & ":L" & rowIndex).Merge
            ws.Range("I" & rowIndex & ":L" & rowIndex).HorizontalAlignment = xlCenter
            ws.Range("B" & rowIndex & ":L" & rowIndex).Borders.LineStyle = xlContinuous
            ws.Range("J" & rowIndex & ":L" & rowIndex).style = "Currency"
            ws.Range("B" & rowIndex & ":I" & rowIndex).NumberFormat = "General"
            ws.Range("C" & rowIndex & ":H" & rowIndex).ShrinkToFit = True
    
            rowIndex = rowIndex + 1
        End If
    Next i
    
    valor = txtTotal.Value
    
    ws.Range("B" & rowIndex).Value = "TOTAL"
    ws.Range("B" & rowIndex & ":J" & rowIndex).Merge
    ws.Range("B" & rowIndex & ":H" & rowIndex).HorizontalAlignment = xlLeft
    ws.Range("B" & rowIndex & ":L" & rowIndex).Borders.LineStyle = xlContinuous
    
    ws.Range("K" & rowIndex).Value = Format(valor, "Currency")
    ws.Range("K" & rowIndex & ":L" & rowIndex).Merge
    ws.Range("K" & rowIndex & ":L" & rowIndex).HorizontalAlignment = xlCenter
    ws.Range("K" & rowIndex & ":L" & rowIndex).style = "Currency"
    
    'ws.Range("B" & rowIndex & ":L" & rowIndex ).Font.Bold = True
    
    ws.Range("A" & rowIndex + 2).Value = "Total a ser descontado: " & Format(valor, "Currency")
    
    ws.Range("A" & rowIndex + 3).Value = "Forma de pagamento: "
    ws.Range("A" & rowIndex + 3 & ":C" & rowIndex + 3).Merge
    ws.Range("A" & rowIndex + 3 & ":C" & rowIndex + 3).HorizontalAlignment = xlLeft
    
    ws.Range("D" & rowIndex + 3).Value = parcela
    ws.Range("D" & rowIndex + 3).HorizontalAlignment = xlCenter
    ws.Range("D" & rowIndex + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    ws.Range("E" & rowIndex + 3).Value = "parcela(s) de "
    
    ws.Range("F" & rowIndex + 3).Value = Format(Round(CDbl(valor) / parcela, 2), "Currency")
    ws.Range("F" & rowIndex + 3 & ":G" & rowIndex + 3).Merge
    ws.Range("F" & rowIndex + 3 & ":G" & rowIndex + 3).HorizontalAlignment = xlCenter
    ws.Range("F" & rowIndex + 3 & ":G" & rowIndex + 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    ws.Range("H" & rowIndex + 3).Value = "cada, a ser(em) descontada(s) em folha de pagamento."
    ws.Range("H" & rowIndex + 3 & ":L" & rowIndex + 3).Merge
    ws.Range("H" & rowIndex + 3 & ":L" & rowIndex + 3).HorizontalAlignment = xlLeft

    ws.Range("A" & rowIndex + 5).Value = "Por fim, em caso de rescisão contratual, autorizo " & _
                                         "que eventuais valores pendentes referentes aos equipamentos " & _
                                         "poderão ser descontados das verbas rescisórias, conforme " & _
                                         "previsto na legislação trabalhista."
                                         
    ws.Range("A" & rowIndex + 5 & ":L" & rowIndex + 5).Merge
    ws.Range("A" & rowIndex + 5).RowHeight = 42
    ws.Range("A" & rowIndex + 5).WrapText = True
    ws.Range("A" & rowIndex + 5).HorizontalAlignment = xlLeft
    
    texto = "de"
    'ws.Range("I" & rowIndex + 7).Value = base_cautela & ", " & Format(Date, "dd-mmmm-yyyy")
    ws.Range("I" & rowIndex + 7).Value = base_cautela & ", " & Format(Date, "dd") & " de " _
                                         & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")
    ws.Range("I" & rowIndex + 7 & ":L" & rowIndex + 7).Merge
    ws.Range("I" & rowIndex + 7 & ":L" & rowIndex + 7).HorizontalAlignment = xlLeft
    
    ws.Range("E" & rowIndex + 11 & ":J" & rowIndex + 11).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Range("B" & rowIndex + 12).Value = "Assinatura do Funcionário"
    ws.Range("B" & rowIndex + 12 & ":L" & rowIndex + 12).Merge
    ws.Range("B" & rowIndex + 12 & ":L" & rowIndex + 12).HorizontalAlignment = xlCenter
    
    ws.Range("A18:K" & rowIndex + 13).Font.Size = 12


    ws.PageSetup.PrintArea = "A1:M" & rowIndex + 13
    ws.PageSetup.Orientation = xlPortrait
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    MsgBox "Vale gerado com sucesso! Verifique a aba frm_Vale.", vbInformation
    
    'Me.Hide
    'ws.PrintPreview
    'Me.Show
    'ws.PrintOut 'Imprimir direto
  
ExitSub:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Erro ao gerar o Vale: " & Err.Description, vbCritical
    GoTo ExitSub
End Sub

'Imprimir Cautela
Private Sub cmdCautela_Click()
    Dim opcao As VbMsgBoxResult
    Dim wsAdmissao As Worksheet
    Dim i As Long
    Dim listItem As Object
    Dim rowIndex As Long
    Dim qtdItens As Long
    Dim tipoCautela As String
    Dim nome As String
    Dim matr As String
    Dim funcao As Variant
   
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    frmTipoCautela.Show
    tipoCautela = frmTipoCautela.TipoEscolhido
    Unload frmTipoCautela
    
    If tipoCautela = "Cancelar" Or tipoCautela = "" Then Exit Sub
    
    qtdItens = Me.lstEpi.ListItems.Count
    If qtdItens = 0 Then
        MsgBox "Nenhum item disponível para gerar a cautela.", vbExclamation, "Cautela"
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsAdmissao = ThisWorkbook.Sheets("frm_Cautela")
    On Error GoTo ErrorHandler
   
'COMPLETA
    If tipoCautela = "Completa" Then
        If wsAdmissao Is Nothing Then
            MsgBox "A aba 'frm_Cautela' não foi encontrada na planilha!", vbCritical, "Erro"
            Exit Sub
        End If
        
        If Me.lstEpi.ListItems.Count > 70 Then
            MsgBox "O número de itens excede o máximo permitido de 70. A cautela não será gerada.", vbExclamation, "Cautela"
            Exit Sub
        End If
        
        wsAdmissao.Range("E12:H12").ClearContents
        wsAdmissao.Range("J12").ClearContents
        wsAdmissao.Range("E13").ClearContents
        wsAdmissao.Range("A50:K130").ClearContents
        wsAdmissao.Range("A50:K130").Borders.LineStyle = xlNone
        
        nome = lstEpi.ListItems(1).SubItems(3) 'NOME
        matr = lstEpi.ListItems(1).SubItems(2) 'MATR
        funcao = GetFuncao(matr)
        
        wsAdmissao.Range("E12").Value = nome
        wsAdmissao.Range("J12").Value = Right(matr, 4)
        wsAdmissao.Range("E13").Value = UCase(funcao)
        
        rowIndex = 50
            
            For i = 1 To lstEpi.ListItems.Count
                Set listItem = lstEpi.ListItems(i)
                wsAdmissao.Range("A" & rowIndex).Value = listItem.SubItems(1)  'DATA
                wsAdmissao.Range("B" & rowIndex).Value = listItem.SubItems(5)  'MOTIVO
                wsAdmissao.Range("C" & rowIndex).Value = listItem.SubItems(7)  'COD
                wsAdmissao.Range("D" & rowIndex).Value = listItem.SubItems(8)  'DESCRIÇÃO
                wsAdmissao.Range("G" & rowIndex).Value = listItem.SubItems(9)  'QTDE
                wsAdmissao.Range("H" & rowIndex).Value = listItem.SubItems(13) 'CA
                wsAdmissao.Range("I" & rowIndex).Value = listItem.SubItems(12) 'VCTO
                wsAdmissao.Range("J" & rowIndex).Value = listItem.SubItems(17) 'MOT BAIXA
                wsAdmissao.Range("K" & rowIndex).Value = listItem.SubItems(16) 'DATA BAIXA
                rowIndex = rowIndex + 1
            Next i
            
            wsAdmissao.Range("A50:K" & rowIndex - 1).Font.Size = 10
            wsAdmissao.Range("D50:F" & rowIndex - 1).HorizontalAlignment = xlLeft
            wsAdmissao.Range("A50:K" & rowIndex - 1).Borders.LineStyle = xlContinuous
            
            wsAdmissao.Range("A" & rowIndex + 1 & ":K" & rowIndex + 6).Font.Size = 12
            wsAdmissao.Range("D" & rowIndex + 6).HorizontalAlignment = xlCenter
            wsAdmissao.Range("E" & rowIndex + 6).Borders(xlEdgeTop).LineStyle = xlContinuous
            
            wsAdmissao.Range("B" & rowIndex + 1).Value = "Total de Registros: " & rowIndex - 50
            wsAdmissao.Range("I" & rowIndex + 1).Value = base_cautela & ", " & Format(Date, "dd-mmmm-yyyy")
            wsAdmissao.Range("D" & rowIndex + 6).Value = "Assinatura do Funcionário"
            
            
            wsAdmissao.PageSetup.PrintArea = "A1:K" & rowIndex + 10
            wsAdmissao.PageSetup.Orientation = xlPortrait
            wsAdmissao.PageSetup.FitToPagesWide = 1
            wsAdmissao.PageSetup.FitToPagesTall = False
            
            'Me.Hide
            'wsAdmissao.PrintPreview
            'Me.Show
            'wsAdmissao.PrintOut 'Imprimir direto
            MsgBox "Formulário enviado para impressão. Verifique modelo gerado na aba frm_Cautela!", vbInformation, "Cautela"
        
'SIMPLIFICADA
    ElseIf tipoCautela = "Simplificada" Then
     'ElseIf opcao = vbNo Then
            If wsAdmissao Is Nothing Then
                MsgBox "A aba 'frm_Cautela' não foi encontrada na planilha!", vbCritical, "Erro"
                Exit Sub
            End If
        
        If Me.lstEpi.ListItems.Count > 20 Then
            MsgBox "O número de itens excede o máximo permitido de 20. A cautela não será gerada.", vbExclamation, "Cautela Troca"
            Exit Sub
        End If
        
        wsAdmissao.Range("E12:H12").ClearContents
        wsAdmissao.Range("J12").ClearContents
        wsAdmissao.Range("E13").ClearContents
        wsAdmissao.Range("A50:K130").ClearContents
        wsAdmissao.Range("A50:K130").Borders.LineStyle = xlNone
        
        nome = lstEpi.ListItems(1).SubItems(3) 'NOME
        matr = lstEpi.ListItems(1).SubItems(2) 'MATR
        funcao = GetFuncao(matr)
        
        wsAdmissao.Range("E12").Value = nome
        wsAdmissao.Range("J12").Value = Right(matr, 4)
        wsAdmissao.Range("E13").Value = UCase(funcao)
        
        rowIndex = 50
            
            For i = 1 To lstEpi.ListItems.Count
                Set listItem = lstEpi.ListItems(i)
                wsAdmissao.Range("A" & rowIndex).Value = listItem.SubItems(1)  'DATA
                wsAdmissao.Range("B" & rowIndex).Value = listItem.SubItems(5)  'MOTIVO
                wsAdmissao.Range("C" & rowIndex).Value = listItem.SubItems(7)  'COD
                wsAdmissao.Range("D" & rowIndex).Value = listItem.SubItems(8)  'DESCRIÇÃO
                wsAdmissao.Range("G" & rowIndex).Value = listItem.SubItems(9)  'QTDE
                wsAdmissao.Range("H" & rowIndex).Value = listItem.SubItems(13) 'CA
                wsAdmissao.Range("I" & rowIndex).Value = listItem.SubItems(12) 'VCTO
                wsAdmissao.Range("J" & rowIndex).Value = listItem.SubItems(17) 'MOT BAIXA
                wsAdmissao.Range("K" & rowIndex).Value = listItem.SubItems(16) 'DATA BAIXA
                rowIndex = rowIndex + 1
            Next i
            
            wsAdmissao.Range("A50:K" & rowIndex - 1).Font.Size = 10
            wsAdmissao.Range("D50:F" & rowIndex - 1).HorizontalAlignment = xlLeft
            wsAdmissao.Range("A50:K" & rowIndex - 1).Borders.LineStyle = xlContinuous
            
            wsAdmissao.Range("A" & rowIndex + 1 & ":K" & rowIndex + 6).Font.Size = 12
            wsAdmissao.Range("D" & rowIndex + 6).HorizontalAlignment = xlCenter
            wsAdmissao.Range("E" & rowIndex + 6).Borders(xlEdgeTop).LineStyle = xlContinuous
            
            wsAdmissao.Range("B" & rowIndex + 1).Value = "Total de Registros: " & rowIndex - 50
            wsAdmissao.Range("I" & rowIndex + 1).Value = base_cautela & ", " & Format(Date, "dd-mmmm-yyyy")
            wsAdmissao.Range("D" & rowIndex + 6).Value = "Assinatura do Funcionário"
            
            
            wsAdmissao.PageSetup.PrintArea = "A40:K" & rowIndex + 10
            wsAdmissao.PageSetup.Orientation = xlPortrait
            wsAdmissao.PageSetup.FitToPagesWide = 1
            wsAdmissao.PageSetup.FitToPagesTall = False
        
        'Me.Hide
        'wsAdmissao.PrintPreview
        'Me.Show
        'wsAdmissao.PrintOut 'Imprimir direto
        MsgBox "Formulário enviado para impressão. Verifique modelo gerado na aba frm_Cautela_Troca!", vbInformation, "Cautela Troca"

 End If

ExitSub:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao preencher a aba Cautela: " & Err.Description, vbCritical, "Erro"
    GoTo ExitSub

End Sub



'------ BOTÕES DE COMANDO ------

'Botão Filtrar
Sub cmdFilter_Click()
    Dim ws As Worksheet
    Dim Arr As Variant
    Dim ultLinha As Long, i As Long
    Dim criterioData As String, criterioMatr As String, criterioNome As String
    Dim criterioTipo As String, criterioJustificativa As String, criterioCat As String
    Dim criterioCodigo As String, criterioDescricao As String, criterioVencimento As String
    Dim criterioCA As String, criterioSerie As String, criterioNF As String
    Dim criterioDataBaixa As String, criterioMotivoBaixa As String, criterioObservacao As String
    Dim criterioDataMovimento As String ' incluido em 31/10/2025
    Dim itemX As listItem
    Dim filtrarCautela As Boolean
    Dim dataMatch As Boolean
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set ws = ThisWorkbook.Worksheets("Lançamentos")
    ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Arr = ws.Range("A3:T" & ultLinha).Value
    
    'Criterios
    If optEntrada.Value = True Then
        criterioTipo = "ENTRADA"
    ElseIf optSaida.Value = True Then
        criterioTipo = "SAÍDA"
    Else
        criterioTipo = ""
    End If
    
    criterioDataMovimento = Trim(Me.txtDataMovimento.Value)
    criterioData = Trim(Me.txtDataInicio.Value)
    criterioMatr = Trim(Me.txtMatr.Value)
    criterioNome = Trim(Me.txtNome.Value)
    criterioJustificativa = Trim(Me.txtJustificativa.Value)
    criterioCat = Trim(Me.txtCat.Value)
    criterioCodigo = Trim(Me.txtCodigo.Value)
    criterioDescricao = Trim(Me.txtDescricao.Value)
    criterioVencimento = Trim(Me.txtVencimento.Value)
    criterioCA = Trim(Me.txtCA.Value)
    criterioSerie = Trim(Me.txtSerie.Value)
    criterioNF = Trim(Me.txtNF.Value)
    criterioDataBaixa = Trim(Me.txtDataBaixa.Value)
    criterioMotivoBaixa = Trim(Me.txtMotivoBaixa.Value)
    criterioObservacao = Trim(Me.txtObservacao.Value)
    
    filtrarCautela = Me.chkCautela.Value
    
    If criterioDataMovimento <> "" Then
        If criterioData <> "" Or criterioDataBaixa <> "" Then
            MsgBox "Quando a Data de Movimento estiver preenchida, Data e Data Baixa devem estar vazias.", vbExclamation, "Filtro Data Movimento"
            GoTo Finaliza
        End If
    End If
    
    lstEpi.ListItems.Clear
    
    For i = 1 To UBound(Arr, 1)
    
    dataMatch = True

        '--- NOVO FILTRO: PRIORIDADE PARA txtDataMovimento ---
        If criterioDataMovimento <> "" Then
            dataMatch = False
            If IsDate(Arr(i, 2)) Then
                If Format(Arr(i, 2), "dd/mm/yyyy") Like "*" & criterioDataMovimento & "*" Then dataMatch = True
            End If
            If IsDate(Arr(i, 17)) Then
                If Format(Arr(i, 17), "dd/mm/yyyy") Like "*" & criterioDataMovimento & "*" Then dataMatch = True
            End If

        '--- Caso contrário, filtra pelos campos individuais ---
        ElseIf criterioData <> "" Or criterioDataBaixa <> "" Then
            dataMatch = False

            If criterioData <> "" Then
                If IsDate(Arr(i, 2)) Then
                    If Format(Arr(i, 2), "dd/mm/yyyy") Like "*" & criterioData & "*" Then dataMatch = True
                End If
            End If

            If criterioDataBaixa <> "" Then
                If IsDate(Arr(i, 17)) Then
                    If Format(Arr(i, 17), "dd/mm/yyyy") Like "*" & criterioDataBaixa & "*" Then dataMatch = True
                End If
            End If
        End If
        
        'If (criterioData = "" Or Format(Arr(i, 2), "dd/mm/yyyy") Like "*" & criterioData & "*") And
        
        If dataMatch And _
           (criterioMatr = "" Or Arr(i, 3) Like "*" & criterioMatr & "*") And _
           (criterioNome = "" Or Arr(i, 4) Like "*" & criterioNome & "*") And _
           (criterioTipo = "" Or Arr(i, 5) Like "*" & criterioTipo & "*") And _
           (criterioJustificativa = "" Or Arr(i, 6) Like "*" & criterioJustificativa & "*") And _
           (criterioCat = "" Or Arr(i, 7) Like "*" & criterioCat & "*") And _
           (criterioCodigo = "" Or Arr(i, 8) Like criterioCodigo) And _
           (criterioDescricao = "" Or Arr(i, 9) Like "*" & criterioDescricao & "*") And _
           (criterioVencimento = "" Or Arr(i, 13) Like "*" & criterioVencimento & "*") And _
           (criterioCA = "" Or Arr(i, 14) Like "*" & criterioCA & "*") And _
           (criterioSerie = "" Or Arr(i, 15) Like "*" & criterioSerie & "*") And _
           (criterioNF = "" Or Arr(i, 16) Like "*" & criterioNF & "*") And _
           (criterioDataBaixa = "" Or Arr(i, 17) Like "*" & criterioDataBaixa & "*") And _
           (criterioMotivoBaixa = "" Or Arr(i, 18) Like "*" & criterioMotivoBaixa & "*") And _
           (criterioObservacao = "" Or Arr(i, 20) Like "*" & criterioObservacao & "*") And _
           (Not filtrarCautela Or (Trim(Arr(i, 17)) = "")) _
        Then
           
            Set itemX = lstEpi.ListItems.Add(, , Arr(i, 1)) 'ID
            itemX.SubItems(1) = Format(Arr(i, 2), "dd/mm/yyyy") 'DATA
            itemX.SubItems(2) = Arr(i, 3) 'MATR
            itemX.SubItems(3) = Arr(i, 4) 'NOME
            itemX.SubItems(4) = Arr(i, 5) 'TIPO
            itemX.SubItems(5) = Arr(i, 6) 'JUSTIFICATIVA
            itemX.SubItems(6) = Arr(i, 7) 'CAT
            itemX.SubItems(7) = Arr(i, 8) 'COD
            itemX.SubItems(8) = Arr(i, 9) 'DESCRIÇÃO
            itemX.SubItems(9) = Arr(i, 10) 'QTDE
            itemX.SubItems(10) = Format(Arr(i, 11), "currency") 'RS UNIT
            itemX.SubItems(11) = Format(Arr(i, 12), "currency") 'RS TOTAL
            itemX.SubItems(12) = Format(Arr(i, 13), "dd/mm/yyyy") 'DATA VCTO
            itemX.SubItems(13) = Arr(i, 14) 'CA
            itemX.SubItems(14) = Arr(i, 15) 'SERIE
            itemX.SubItems(15) = Arr(i, 16) 'NF
            itemX.SubItems(16) = IIf(Trim(Arr(i, 17)) <> "", Format(Arr(i, 17), "dd/mm/yyyy"), "") 'DATA BAIXA
            itemX.SubItems(17) = Arr(i, 18) 'MOTIVO BAIXA
            itemX.SubItems(18) = Arr(i, 19) 'VIDA UTIL
            itemX.SubItems(19) = Arr(i, 20) 'OBS
        End If
    Next i
    
Finaliza:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    Call CalculaSubTotal
    
End Sub

'Botão Limpar
Private Sub cmdClear_Click()
    
    txtId = ""
    txtDataInicio = ""
    txtMatr = ""
    txtNome = ""
    txtJustificativa = ""
    txtCat = ""
    txtCodigo = ""
    txtDescricao = ""
    txtQtde = ""
    txtRSUnitario = ""
    txtRSTotal = ""
    txtVencimento = ""
    txtCA = ""
    txtSerie = ""
    txtNF = ""
    txtDataBaixa = ""
    txtMotivoBaixa = ""
    txtObservacao = ""
    txtTotal = ""
    txtDataMovimento = ""
    optEntrada = False
    optSaida = False
    chkCautela = False
    lstEpi.ListItems.Clear

End Sub

'Limpar pos Edição
Private Sub LimparEdicao()
    
    txtId = ""
    txtDataInicio = ""
    'txtMatr = ""
    txtNome = ""
    txtJustificativa = ""
    txtCat = ""
    txtCodigo = ""
    txtDescricao = ""
    txtQtde = ""
    txtRSUnitario = ""
    txtRSTotal = ""
    txtVencimento = ""
    txtCA = ""
    txtSerie = ""
    txtNF = ""
    txtDataBaixa = ""
    txtMotivoBaixa = ""
    txtObservacao = ""
    txtTotal = ""
    txtDataMovimento = ""
    optEntrada = False
    optSaida = False
    
End Sub

'Limpar Inserção
Private Sub LimparInsercao()
    
    txtCat = ""
    txtCodigo = ""
    txtDescricao = ""
    txtQtde = ""
    txtRSUnitario = ""
    txtRSTotal = ""
    txtVencimento = ""
    txtCA = ""
    txtSerie = ""
    txtNF = ""
    txtDataBaixa = ""
    txtMotivoBaixa = ""
    txtObservacao = ""

End Sub

'Botão Adicionar c/verificao de saldo
Private Sub cmdAdd_Click()
    Dim itemX As listItem
    Dim msg As String
    Dim wsAnalise As Worksheet
    Dim codigo As String
    Dim saldo As Double
    Dim qtdSolicitada As Double
    Dim qtdNaLista As Double
    Dim i As Long
    Dim exigeVenc As Boolean
    Dim exigeCA As Boolean
    Dim exigeSerie As Boolean
    Dim tbl As Object
    Dim cel As Range
    Dim achou As Boolean
    
    If Trim(Me.txtId.Value) <> "" Then
        MsgBox "Modo edição. Não é possível adicionar novos itens.", vbExclamation, "Aviso"
        Exit Sub
    End If

    If Me.txtDataInicio.Value = "" Then msg = msg & vbCrLf & "- Data"
    If Me.txtMatr.Value = "" Then msg = msg & vbCrLf & "- Matrícula"
    If Me.txtNome.Value = "" Then msg = msg & vbCrLf & "- Nome"
    If Me.optEntrada.Value = False And Me.optSaida.Value = False Then
        msg = msg & vbCrLf & "- Tipo de movimento (Entrada/Saída)"
    End If
    If Me.txtJustificativa.Value = "" Then msg = msg & vbCrLf & "- Justificativa"
    If Me.txtCodigo.Value = "" Then msg = msg & vbCrLf & "- Código"
    If Me.txtQtde.Value = "" Or Not IsNumeric(Me.txtQtde.Value) Then
        msg = msg & vbCrLf & "- Quantidade"
    End If
    If Me.txtRSUnitario.Value = "" Then msg = msg & vbCrLf & "- R$ Unitário"
    
'--- VERIFICA SE O ITEM EXIGE VENCIMENTO E/OU CA
    Call VerificarObrigatoriedade(Me.txtCodigo.Value, exigeVenc, exigeCA, exigeSerie)

    If exigeVenc And Trim(Me.txtVencimento.Value) = "" Then
        msg = msg & vbCrLf & "- Vencimento (obrigatório para o item)"
    End If

    If exigeCA And Trim(Me.txtCA.Value) = "" Then
        msg = msg & vbCrLf & "- CA (obrigatório para o item)"
    End If
    
    If exigeSerie And Trim(Me.txtSerie.Value) = "" Then
        msg = msg & vbCrLf & "- Serie (obrigatório para o item)"
    End If
    
    If msg <> "" Then
        MsgBox "Preencha os seguintes campos obrigatórios:" & vbCrLf & msg, vbExclamation, "Campos obrigatórios"
        Exit Sub
    End If

    'Verifica saldo
     If Me.optSaida.Value = True And Me.txtDataInicio.Value > dataCorte Then
        
        Set wsAnalise = ThisWorkbook.Worksheets("Auxiliar")
        Set tbl = wsAnalise.ListObjects("tab_codigo")
        codigo = Trim(Me.txtCodigo.Value)
        qtdSolicitada = CDbl(Me.txtQtde.Value)
        achou = False
        
        For Each cel In tbl.ListColumns("COD").DataBodyRange
            If Trim(cel.Value) = codigo Then
                achou = True
                saldo = CDbl(cel.Offset(0, 11).Value)
            Exit For
            End If
        Next cel
        
        If Not achou Then
            MsgBox "Item não encontrado na planilha Auxiliar. Verifique o código.", vbExclamation
            saldo = 0
            'Cancel = True
        End If

            qtdNaLista = 0
            For i = 1 To Me.lstEpi.ListItems.Count
                If Trim(Me.lstEpi.ListItems(i).SubItems(7)) = codigo And _
                   Trim(Me.lstEpi.ListItems(i).SubItems(4)) = "SAÍDA" Then
                    qtdNaLista = qtdNaLista + CDbl(Me.lstEpi.ListItems(i).SubItems(9)) 'QTDE
                End If
            Next i

        If saldo <= 0 Then
            MsgBox "Item '" & codigo & "' está sem saldo disponível!", vbCritical, "Saldo insuficiente"
            Exit Sub
        ElseIf qtdSolicitada + qtdNaLista > saldo Then
            MsgBox "Saldo insuficiente para o item '" & codigo & "'!" & vbCrLf & _
                   "Saldo atual: " & saldo & vbCrLf & _
                   "Quantidade já lançada: " & qtdNaLista & vbCrLf & _
                   "Quantidade solicitada: " & qtdSolicitada, vbCritical, "Saldo insuficiente"
            Exit Sub
        End If
    End If
    
    Set itemX = Me.lstEpi.ListItems.Add(, , "")
    itemX.SubItems(1) = Format(Me.txtDataInicio.Value, "dd/mm/yyyy")          'DATA
    itemX.SubItems(2) = Me.txtMatr.Value                                      'MATR/CNPJ
    itemX.SubItems(3) = Me.txtNome.Value                                      'NOME
    itemX.SubItems(4) = IIf(Me.optEntrada.Value, "ENTRADA", "SAÍDA")          'TIPO MOV
    itemX.SubItems(5) = Me.txtJustificativa.Value                             'JUSTIFICATIVA
    itemX.SubItems(6) = Me.txtCat.Value                                       'CAT
    itemX.SubItems(7) = Me.txtCodigo.Value                                    'COD
    itemX.SubItems(8) = Me.txtDescricao.Value                                 'DESCRIÇÃO
    itemX.SubItems(9) = Me.txtQtde.Value                                      'QTDE
    itemX.SubItems(10) = Format(CDbl(Me.txtRSUnitario.Value), "Currency")     'R$ UNIT
    itemX.SubItems(11) = Format(Me.txtQtde.Value * CDbl(Me.txtRSUnitario.Value), "Currency") 'R$ TOTAL
    
    If Trim(Me.txtVencimento.Value) <> "" Then                                'VCTO
        itemX.SubItems(12) = Format(CDate(Me.txtVencimento.Value), "dd/mm/yyyy")
    Else
        itemX.SubItems(12) = ""
    End If
    
    itemX.SubItems(13) = Me.txtCA.Value                                       'CA
    itemX.SubItems(14) = Me.txtSerie.Value                                    'N° SERIE
    itemX.SubItems(15) = Me.txtNF.Value                                       'N.FISCAL
    
    If Trim(Me.txtDataBaixa.Value) <> "" Then                                 'DT BAIXA
        itemX.SubItems(16) = Format(CDate(Me.txtDataBaixa.Value), "dd/mm/yyyy")
    Else
        itemX.SubItems(16) = ""
    End If
    
    itemX.SubItems(17) = Me.txtMotivoBaixa.Value                              'MOTIVO BAIXA
    itemX.SubItems(19) = Me.txtObservacao.Value                               'OBSERVAÇÃO

    Call LimparInsercao
    Call CalculaSubTotal
    Me.txtCodigo.SetFocus
    
End Sub

'Botao Baixar
Private Sub cmdBaixar_Click()
    Dim i As Long
    Dim checkedCount As Integer
    
    checkedCount = 0
    
    For i = 1 To Me.lstEpi.ListItems.Count
        If Me.lstEpi.ListItems(i).Checked Then
            checkedCount = checkedCount + 1
        End If
    Next i

    If checkedCount = 0 Then
        MsgBox "Selecione pelo menos um item para efetuar a baixa.", vbExclamation, "Nenhum item selecionado"
        Exit Sub
    End If
    
    If checkedCount > 1 Then
        MsgBox "Baixa parcial suportada apenas para um item por vez.", vbExclamation, "Múltiplos itens selecionados"
        Exit Sub
    End If

    frmBaixaMaterial.Show vbModal
End Sub

'Função Baixar Material
Public Sub AplicarBaixaMaterial(ByVal dataBaixa As String, _
                                ByVal motivo As String, _
                                ByVal novaEntrega As Boolean, _
                                ByVal qtde As Double, _
                                Optional ByVal vencimento As String = "", _
                                Optional ByVal ca As String = "", _
                                Optional ByVal serie As String = "", _
                                Optional ByVal estadoDev As String = "")

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim ultLinha As Long, novoId As Long, newRow As Long
    Dim i As Long
    Dim checkedItem As listItem
    Dim originalQtde As Double
    Dim unitValue As Double
    Dim rngOriginal As Range
    Dim itemId As String
    Dim totalBaixa As Boolean
    Dim fazEntrada As Boolean
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set ws = ThisWorkbook.Sheets("Lançamentos")
    Set tbl = ws.ListObjects("tab_lancamentos")
    
    For i = 1 To lstEpi.ListItems.Count
        If lstEpi.ListItems(i).Checked Then
            Set checkedItem = lstEpi.ListItems(i)
            Exit For
        End If
    Next i
    
    If checkedItem Is Nothing Then
        MsgBox "Nenhum item selecionado.", vbExclamation
        GoTo ExitSub
    End If
    
    If Trim(UCase$(checkedItem.SubItems(4))) <> "SAÍDA" Then
        MsgBox "A baixa só pode ser feita em itens de SAÍDA.", vbExclamation
        GoTo ExitSub
    End If
    
    If Trim(checkedItem.SubItems(16)) <> "" Then
        MsgBox "Este item já foi baixado anteriormente.", vbExclamation
        GoTo ExitSub
    End If
    
    itemId = Trim(checkedItem.Text)
    originalQtde = CDbl(checkedItem.SubItems(9))
    
    If qtde <= 0 Or qtde > originalQtde Then
        MsgBox "Quantidade inválida para baixa.", vbExclamation
        GoTo ExitSub
    End If
    
    totalBaixa = (qtde = originalQtde)
    
    ' Localizar a linha original
    'Set rngOriginal = ws.Range("A:A").Find(What:=itemId, LookIn:=xlValues, LookAt:=xlWhole)
    Set rngOriginal = GetRowById(tbl, itemId)
    
    If rngOriginal Is Nothing Then
        MsgBox "Registro original não encontrado!", vbCritical
        GoTo ExitSub
    End If
    
    If IsNumeric(rngOriginal.Offset(0, 10).Value) Then
        unitValue = CDbl(rngOriginal.Offset(0, 10).Value)
    Else
        unitValue = 0
    End If

    fazEntrada = RetornaAoEstoque(estadoDev)

'BAIXA

'Total
    If totalBaixa Then
        rngOriginal.Offset(0, 16).Value = CDate(dataBaixa)
        rngOriginal.Offset(0, 17).Value = motivo
        rngOriginal.Offset(0, 19).Value = "Baixa total do item ID " & itemId
        
'Parcial
    Else
        rngOriginal.Offset(0, 9).Value = originalQtde - qtde
        
        ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        novoId = ObterProximoId(ws)
        newRow = ultLinha + 1
        
        ws.Cells(newRow, 1).Value = novoId                          'ID
        ws.Cells(newRow, 2).Value = rngOriginal.Offset(0, 1).Value
        ws.Cells(newRow, 3).Value = rngOriginal.Offset(0, 2).Value
        ws.Cells(newRow, 4).Value = rngOriginal.Offset(0, 3).Value
        ws.Cells(newRow, 5).Value = "SAÍDA"
        ws.Cells(newRow, 6).Value = rngOriginal.Offset(0, 5).Value
        ws.Cells(newRow, 7).Value = rngOriginal.Offset(0, 6).Value
        ws.Cells(newRow, 8).Value = rngOriginal.Offset(0, 7).Value
        ws.Cells(newRow, 9).Value = rngOriginal.Offset(0, 8).Value
        ws.Cells(newRow, 10).Value = qtde
        ws.Cells(newRow, 11).Value = unitValue
        ws.Cells(newRow, 13).Value = rngOriginal.Offset(0, 12).Value 'VCTO
        ws.Cells(newRow, 14).Value = rngOriginal.Offset(0, 13).Value
        ws.Cells(newRow, 15).Value = rngOriginal.Offset(0, 14).Value
        ws.Cells(newRow, 16).Value = rngOriginal.Offset(0, 15).Value
        ws.Cells(newRow, 17).Value = CDate(dataBaixa)                            'DT BAIXA
        ws.Cells(newRow, 18).Value = motivo
        ws.Cells(newRow, 20).Value = "Baixa parcial do item ID " & itemId
    End If
    
' se houver entrega nova
    If Not novaEntrega Then
        If fazEntrada Then
            ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            novoId = ObterProximoId(ws)
            newRow = ultLinha + 1
            
            ws.Cells(newRow, 1).Value = novoId
            ws.Cells(newRow, 2).Value = CDate(dataBaixa)
            ws.Cells(newRow, 3).Value = rngOriginal.Offset(0, 2).Value
            ws.Cells(newRow, 4).Value = rngOriginal.Offset(0, 3).Value
            ws.Cells(newRow, 5).Value = "ENTRADA"
            ws.Cells(newRow, 6).Value = Trim$(estadoDev)
            ws.Cells(newRow, 7).Value = rngOriginal.Offset(0, 6).Value
            ws.Cells(newRow, 8).Value = rngOriginal.Offset(0, 7).Value
            ws.Cells(newRow, 9).Value = rngOriginal.Offset(0, 8).Value
            ws.Cells(newRow, 10).Value = qtde
            ws.Cells(newRow, 11).Value = unitValue
            ws.Cells(newRow, 13).Value = rngOriginal.Offset(0, 12).Value
            ws.Cells(newRow, 14).Value = rngOriginal.Offset(0, 13).Value
            ws.Cells(newRow, 20).Value = "Devolução do item ID " & itemId
        End If
    End If

'TROCA
    If novaEntrega Then
    
        ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        novoId = ObterProximoId(ws)
        newRow = ultLinha + 1
        
        ws.Cells(newRow, 1).Value = novoId
        ws.Cells(newRow, 2).Value = CDate(dataBaixa)
        ws.Cells(newRow, 3).Value = rngOriginal.Offset(0, 2).Value
        ws.Cells(newRow, 4).Value = rngOriginal.Offset(0, 3).Value
        ws.Cells(newRow, 5).Value = "SAÍDA"
        ws.Cells(newRow, 6).Value = "TROCA"
        ws.Cells(newRow, 7).Value = rngOriginal.Offset(0, 6).Value
        ws.Cells(newRow, 8).Value = rngOriginal.Offset(0, 7).Value
        ws.Cells(newRow, 9).Value = rngOriginal.Offset(0, 8).Value
        ws.Cells(newRow, 10).Value = qtde
        ws.Cells(newRow, 11).Value = unitValue
        'ws.Cells(newRow, 13).Value = IIf(vencimento <> "", CDate(vencimento), rngOriginal.Offset(0, 12).Value)
        If vencimento <> "" Then
            ws.Cells(newRow, 13).Value = CDate(vencimento)
        Else
            ws.Cells(newRow, 13).Value = rngOriginal.Offset(0, 12).Value
        End If

        ws.Cells(newRow, 14).Value = IIf(ca <> "", ca, rngOriginal.Offset(0, 13).Value)
        ws.Cells(newRow, 15).Value = serie
        ws.Cells(newRow, 20).Value = "Entrega de substituto do item ID " & itemId
        
        If fazEntrada Then
        
            ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            novoId = ObterProximoId(ws)
            newRow = ultLinha + 1
            
            ws.Cells(newRow, 1).Value = novoId
            ws.Cells(newRow, 2).Value = CDate(dataBaixa)
            ws.Cells(newRow, 3).Value = rngOriginal.Offset(0, 2).Value
            ws.Cells(newRow, 4).Value = rngOriginal.Offset(0, 3).Value
            ws.Cells(newRow, 5).Value = "ENTRADA"
            ws.Cells(newRow, 6).Value = Trim$(estadoDev)
            ws.Cells(newRow, 7).Value = rngOriginal.Offset(0, 6).Value
            ws.Cells(newRow, 8).Value = rngOriginal.Offset(0, 7).Value
            ws.Cells(newRow, 9).Value = rngOriginal.Offset(0, 8).Value
            ws.Cells(newRow, 10).Value = qtde
            ws.Cells(newRow, 11).Value = unitValue
            ws.Cells(newRow, 13).Value = rngOriginal.Offset(0, 12).Value
            ws.Cells(newRow, 14).Value = rngOriginal.Offset(0, 13).Value
            ws.Cells(newRow, 20).Value = "Devolução do item ID " & itemId
            
        End If
    End If

    checkedItem.Checked = False
    Call LimparEdicao
    Call cmdFilter_Click
    
    MsgBox "Baixa processada com sucesso.", vbInformation

ExitSub:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical
    Resume ExitSub
End Sub

'Botão Editar
Private Sub cmdEdit_Click()
Dim item As listItem
Dim i As Long
Dim checkedCount As Long


    checkedCount = 0

    'Procura qual item está marcado
    For i = 1 To lstEpi.ListItems.Count
        If lstEpi.ListItems(i).Checked Then
            checkedCount = checkedCount + 1
            Set item = lstEpi.ListItems(i)
        End If
    Next i

    If checkedCount = 0 Then
        MsgBox "Selecione um item (marcando o checkbox) para editar.", vbExclamation, "Edição"
        Exit Sub
    End If

    If checkedCount > 1 Then
        MsgBox "Só é possível editar UM item por vez.", vbExclamation, "Edição"
        Exit Sub
    End If
    
    Me.txtId.Value = item.Text
    Me.txtDataInicio.Value = item.SubItems(1)
    
    dataOriginal = CDate(item.SubItems(1))
    
    Me.txtMatr.Value = item.SubItems(2)
    Me.txtNome.Value = item.SubItems(3)
    If item.SubItems(4) = "ENTRADA" Then
        Me.optEntrada.Value = True
    Else
        Me.optSaida.Value = True
    End If
    
    tipoOriginal = CStr(item.SubItems(4))
    
    Me.txtJustificativa.Value = item.SubItems(5)
    Me.txtCat.Value = item.SubItems(6)
    Me.txtCodigo.Value = item.SubItems(7)
    Me.txtDescricao.Value = item.SubItems(8)
    Me.txtQtde.Value = item.SubItems(9)
    
    qtdOriginal = CDbl(item.SubItems(9))
    
    Me.txtRSUnitario.Value = item.SubItems(10)
    Me.txtVencimento.Value = item.SubItems(12)
    Me.txtCA.Value = item.SubItems(13)
    Me.txtSerie.Value = item.SubItems(14)
    Me.txtNF.Value = item.SubItems(15)
    Me.txtDataBaixa.Value = item.SubItems(16)
    Me.txtMotivoBaixa.Value = item.SubItems(17)
    Me.txtObservacao.Value = item.SubItems(19)
    
    Call txtQtde_Change
    Call CalculaSubTotal
    
End Sub

'Botão Delete
Private Sub cmdDelete_Click()
    Dim i As Long, deleteCount As Long
    Dim itemId As String
    Dim data As Date
    Dim codigo As String
    Dim tipo As String
    Dim qtde As Double
    Dim achou As Boolean
    Dim ws As Worksheet
    Dim wsAnalise As Worksheet
    Dim tblLcto As ListObject
    Dim tblAnalise As Object
    Dim saldo As Double
    Dim saldoFinal As Double
    Dim rng As Range
    Dim cel As Range
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsAnalise = ThisWorkbook.Worksheets("Auxiliar")
    Set tblAnalise = wsAnalise.ListObjects("tab_codigo")
    Set ws = ThisWorkbook.Worksheets("Lançamentos")
    Set tblLcto = ws.ListObjects("tab_lancamentos")
    deleteCount = 0
    
    For i = 1 To lstEpi.ListItems.Count
        If lstEpi.ListItems(i).Checked Then
            deleteCount = deleteCount + 1
        End If
    Next i
    
    If deleteCount = 0 Then
        MsgBox "Nenhum item selecionado para exclusão!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Deseja realmente excluir " & deleteCount & " item(s)?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then
        Exit Sub
    End If
    
    For i = lstEpi.ListItems.Count To 1 Step -1
        If lstEpi.ListItems(i).Checked Then
            itemId = lstEpi.ListItems(i).Text
            data = lstEpi.ListItems(i).SubItems(1)
            tipo = lstEpi.ListItems(i).SubItems(4)
            codigo = lstEpi.ListItems(i).SubItems(7)
            qtde = lstEpi.ListItems(i).SubItems(9)
            achou = False
            
            If data > dataCorte Then
                For Each cel In tblAnalise.ListColumns("COD").DataBodyRange
                    If Trim(cel.Value) = codigo Then
                        achou = True
                        saldo = CDbl(cel.Offset(0, 11).Value)
                        Exit For
                    End If
                Next cel
                
                    If Not achou Then
                        MsgBox "Item não encontrado na planilha Auxiliar. Verifique o código.", vbExclamation
                        saldo = 0
                    End If
                    
                    If tipo = "ENTRADA" Then
                        saldoFinal = saldo - qtde
                    Else
                        saldoFinal = saldo + qtde
                    End If
                    
                    If saldoFinal < 0 Then
                    MsgBox "Esta exclusão irá gerar um saldo negativo no item '" & codigo & "'!" & vbCrLf & _
                       "Saldo atual: " & saldo & vbCrLf & _
                       "Saldo após a exclusão: " & (saldoFinal), _
                       vbCritical, "Saldo insuficiente"
                        Exit Sub
                    Else
                        Set rng = GetRowById(tblLcto, itemId)
                        If Not rng Is Nothing Then
                            rng.EntireRow.Delete
                        End If

                    End If
            Else
                Set rng = GetRowById(tblLcto, itemId)
                If Not rng Is Nothing Then
                    rng.EntireRow.Delete
                End If
            End If
        End If
    Next i
    
    MsgBox deleteCount & " item(s) excluído(s) com sucesso!", vbInformation
    
    Call cmdFilter_Click
    Call CalculaSubTotal
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
End Sub

'Botão Salvar data corte
Private Sub cmdSalvar_Click()

    Dim ws As Worksheet
    Dim wsAnalise As Worksheet
    Dim tblAnalise As Object
    Dim tbl As ListObject
    Dim ultLinha As Long
    Dim i As Long
    Dim novoId As Long
    Dim salvo As Boolean
    Dim rng As Range
    Dim msg As String
    Dim exigeVenc As Boolean
    Dim exigeCA As Boolean
    Dim exigeSerie As Boolean
    Dim codigo As String
    Dim achou As Boolean
    Dim cel As Range
    Dim saldo As Double
    Dim calcula_saldo As Double
    Dim dataAntiga As Date, dataNova As Date
    Dim qtdAntiga As Double, qtdNova As Double
    Dim saldoFinal As Double
    Dim tipoMov As String, tipoNovo As String
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set ws = ThisWorkbook.Sheets("Lançamentos")
    Set tbl = ws.ListObjects("tab_lancamentos")
    Set wsAnalise = ThisWorkbook.Worksheets("Auxiliar")
    Set tblAnalise = wsAnalise.ListObjects("tab_codigo")

    
    ultLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    novoId = ObterProximoId(ws) - 1
    salvo = False
    
' MODO EDIÇÃO
    If Trim(Me.txtId.Value) <> "" Then
        
        Set rng = GetRowById(tbl, Me.txtId.Value)
        
        If Not rng Is Nothing Then
            
            'Verifica preenchimento de campos obrigatórios
            If Me.txtDataInicio.Value = "" Then msg = msg & vbCrLf & "- Data"
            If Me.txtMatr.Value = "" Then msg = msg & vbCrLf & "- Matrícula"
            If Me.txtNome.Value = "" Then msg = msg & vbCrLf & "- Nome"
            If Me.optEntrada.Value = False And Me.optSaida.Value = False Then
                msg = msg & vbCrLf & "- Tipo de movimento (Entrada/Saída)"
            End If
            If Me.txtJustificativa.Value = "" Then msg = msg & vbCrLf & "- Justificativa"
            If Me.txtCodigo.Value = "" Then msg = msg & vbCrLf & "- Código"
            If Me.txtQtde.Value = "" Or Not IsNumeric(Me.txtQtde.Value) Then
                msg = msg & vbCrLf & "- Quantidade"
            End If
            If Me.txtRSUnitario.Value = "" Then msg = msg & vbCrLf & "- R$ Unitário"

            'Verifica Obrigatoriedade CA/Serie/Vencimento
            codigo = Trim(Me.txtCodigo.Value)
            Call VerificarObrigatoriedade(Me.txtCodigo.Value, exigeVenc, exigeCA, exigeSerie)
        
            If exigeVenc And Trim(Me.txtVencimento.Value) = "" Then
                msg = msg & vbCrLf & "- Vencimento (obrigatório para o item)"
            End If
        
            If exigeCA And Trim(Me.txtCA.Value) = "" Then
                msg = msg & vbCrLf & "- CA (obrigatório para o item)"
            End If
            
            If exigeSerie And Trim(Me.txtSerie.Value) = "" Then
                msg = msg & vbCrLf & "- Serie (obrigatório para o item)"
            End If
            If msg <> "" Then
                MsgBox "Preencha os seguintes campos obrigatórios:" & vbCrLf & msg, vbExclamation, "Campos obrigatórios"
                Exit Sub
            End If
                    
            'Localiza saldo do item
            dataNova = Me.txtDataInicio.Value
            qtdNova = CDbl(Me.txtQtde.Value)
            tipoNovo = IIf(Me.optEntrada.Value, "ENTRADA", "SAÍDA")
            achou = False
            
            If tipoOriginal <> tipoNovo Then
                MsgBox "Não é possível alterar o tipo de movimentação no modo Editar", vbCritical, "Erro"
                Exit Sub
            End If
            
            For Each cel In tblAnalise.ListColumns("COD").DataBodyRange
                If Trim(cel.Value) = codigo Then
                    achou = True
                    saldo = CDbl(cel.Offset(0, 11).Value)
                Exit For
                End If
            Next cel
            
            If Not achou Then
                MsgBox "Item não encontrado na planilha Auxiliar. Verifique o código.", vbExclamation
                saldo = 0
            End If
            
            saldoFinal = CalcularSaldo(dataCorte, dataOriginal, dataNova, qtdOriginal, qtdNova, saldo, tipoOriginal)
            
            If saldoFinal < 0 Then
                MsgBox "Esta edição irá gerar um saldo negativo no item '" & codigo & "'!" & vbCrLf & _
                       "Saldo atual: " & saldo & vbCrLf & _
                       "Saldo após a edição: " & (saldoFinal), _
                       vbCritical, "Saldo insuficiente"
                Exit Sub
            End If
           
            'Grava edição em Lançamentos
            rng.Offset(0, 1).Value = CDate(Me.txtDataInicio.Value)
            rng.Offset(0, 2).Value = Me.txtMatr.Value
            rng.Offset(0, 3).Value = Me.txtNome.Value
            rng.Offset(0, 4).Value = IIf(Me.optEntrada.Value, "ENTRADA", "SAÍDA")
            rng.Offset(0, 5).Value = Me.txtJustificativa.Value
            rng.Offset(0, 6).Value = Me.txtCat.Value
            rng.Offset(0, 7).Value = Me.txtCodigo.Value
            rng.Offset(0, 8).Value = Me.txtDescricao.Value
            rng.Offset(0, 9).Value = Me.txtQtde.Value
            rng.Offset(0, 10).Value = CDbl(Me.txtRSUnitario.Value)
            If Trim(Me.txtVencimento.Value) <> "" Then
                rng.Offset(0, 12).Value = CDate(Me.txtVencimento.Value)
            Else
                rng.Offset(0, 12).Value = ""
            End If
            rng.Offset(0, 13).Value = Me.txtCA.Value
            rng.Offset(0, 14).Value = Me.txtSerie.Value
            rng.Offset(0, 15).Value = Me.txtNF.Value
            If Trim(Me.txtDataBaixa.Value) <> "" Then
                rng.Offset(0, 16).Value = CDate(Me.txtDataBaixa.Value)
            Else
                rng.Offset(0, 16).Value = ""
            End If
            rng.Offset(0, 17).Value = Me.txtMotivoBaixa.Value
            rng.Offset(0, 19).Value = Me.txtObservacao.Value
            
            salvo = True
            MsgBox "Alterações salvas com sucesso!", vbInformation, "Edição"
            
            'Call LimparEdicao
        Else
            MsgBox "Registro não encontrado para edição!", vbCritical
        End If
    
' MODO ADIÇÃO
    Else
        For i = 1 To Me.lstEpi.ListItems.Count
            If Trim(Me.lstEpi.ListItems(i).Text) = "" Then
                novoId = novoId + 1
                ultLinha = ultLinha + 1
                
                With lstEpi
                    ws.Cells(ultLinha, 1).Value = novoId                            'ID
                    ws.Cells(ultLinha, 2).Value = CDate(.ListItems(i).SubItems(1))  'DATA
                    ws.Cells(ultLinha, 3).Value = .ListItems(i).SubItems(2)         'MATR
                    ws.Cells(ultLinha, 4).Value = .ListItems(i).SubItems(3)         'NOME
                    ws.Cells(ultLinha, 5).Value = .ListItems(i).SubItems(4)         'MOV
                    ws.Cells(ultLinha, 6).Value = .ListItems(i).SubItems(5)         'JUST
                    ws.Cells(ultLinha, 7).Value = .ListItems(i).SubItems(6)         'CAT
                    ws.Cells(ultLinha, 8).Value = .ListItems(i).SubItems(7)         'COD
                    ws.Cells(ultLinha, 9).Value = .ListItems(i).SubItems(8)         'DESC
                    ws.Cells(ultLinha, 10).Value = .ListItems(i).SubItems(9)        'QTDE
                    ws.Cells(ultLinha, 11).Value = CDbl(.ListItems(i).SubItems(10)) 'RS UN
       
                    If Trim(.ListItems(i).SubItems(12)) <> "" Then                  'VENC
                        ws.Cells(ultLinha, 13).Value = CDate(.ListItems(i).SubItems(12))
                    Else
                        ws.Cells(ultLinha, 13).Value = ""
                    End If
                    
                    ws.Cells(ultLinha, 14).Value = .ListItems(i).SubItems(13)       'CA
                    ws.Cells(ultLinha, 15).Value = .ListItems(i).SubItems(14)       'SERIE
                    ws.Cells(ultLinha, 16).Value = .ListItems(i).SubItems(15)       'NF
                    
                    
                    If Trim(.ListItems(i).SubItems(16)) <> "" Then                  'DATA BAIXA
                        ws.Cells(ultLinha, 17).Value = CDate(.ListItems(i).SubItems(16))
                    Else
                        ws.Cells(ultLinha, 17).Value = ""
                    End If

                    ws.Cells(ultLinha, 18).Value = .ListItems(i).SubItems(17)       'MOT BAIXA
                    ws.Cells(ultLinha, 20).Value = .ListItems(i).SubItems(19)       'OBS
                End With
                
                salvo = True
            End If
        Next i
        
        If salvo Then
            MsgBox "Itens gravados em Lançamentos!", vbInformation
            Me.lstEpi.ListItems.Clear
        Else
            MsgBox "Nenhum item novo para salvar.", vbInformation
        End If
    End If
        
    Call LimparEdicao
    cmdFilter_Click
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub


