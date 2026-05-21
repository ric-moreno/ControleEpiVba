VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBaixaMaterial 
   Caption         =   "Efetuar Baixa de Material"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   OleObjectBlob   =   "frmBaixaMaterial.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBaixaMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Const dataCorte As Date = "07/05/2026" 'Alterar no frmEpi também

Private Sub UserForm_Initialize()

    Me.TextVencimentoNovaEntrega.Enabled = False
    Me.TextCANovaEntrega.Enabled = False
    Me.txtSerieNovaEntrega.Enabled = False

End Sub

'CheckBox Nova entrega
Private Sub chkNovaEntrega_Click()

    Dim habilitar As Boolean
    habilitar = (Me.chkNovaEntrega.Value = True)

    Me.TextVencimentoNovaEntrega.Enabled = habilitar
    Me.TextCANovaEntrega.Enabled = habilitar
    Me.txtSerieNovaEntrega.Enabled = habilitar

    If Not habilitar Then
        Me.TextVencimentoNovaEntrega.Text = ""
        Me.TextCANovaEntrega.Text = ""
        Me.txtSerieNovaEntrega.Text = ""
    End If

End Sub

'BOTÕES
Private Sub cmdConfirmar_Click()
    Dim wsAnalise As Worksheet
    Dim tbl As Object
    Dim motivo As String
    Dim novaEntrega As Boolean
    Dim qtde As Double
    Dim vencimento As String
    Dim ca As String
    Dim serie As String
    Dim estadoDev As String
    Dim codigoItem As String
    Dim exigeVencimento As Boolean
    Dim exigeCA As Boolean
    Dim exigeSerie As Boolean
    Dim codigo As String
    Dim saldo As Double
    Dim qtdSolicitada As Double
    Dim cel As Range
    Dim dataBaixa As String
    
    Set wsAnalise = ThisWorkbook.Worksheets("Auxiliar")
    Set tbl = wsAnalise.ListObjects("tab_codigo")
    dataBaixa = Trim(Me.txtDataBaixaItem.Value)
    motivo = Trim(Me.txtMotivoBaixa.Value)
    novaEntrega = Me.chkNovaEntrega.Value
    estadoDev = Trim(Me.txtEstadoDev.Value)
    vencimento = Trim(Me.TextVencimentoNovaEntrega.Value)
    ca = Trim(Me.TextCANovaEntrega.Value)
    serie = Trim(Me.txtSerieNovaEntrega.Value)
    
    If dataBaixa = "" Then
        MsgBox "Informe a data da baixa.", vbExclamation
        Exit Sub
    End If

    If motivo = "" Then
        MsgBox "Informe o motivo da baixa.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(Me.txtQtdeBaixa.Value) Or Val(Me.txtQtdeBaixa.Value) <= 0 Then
        MsgBox "Informe uma quantidade válida para a baixa.", vbExclamation
        Exit Sub
    End If
    
    qtde = CDbl(Me.txtQtdeBaixa.Value)
    
    '--- VALIDAÇÃO OBRIGATORIEDADE CA / VALIDADE
    If novaEntrega Then
        
        Dim checkedItem As listItem
        For Each checkedItem In frmEpi.lstEpi.ListItems
            If checkedItem.Checked Then
                codigoItem = Trim(checkedItem.SubItems(7))
                Exit For
            End If
        Next checkedItem
        
        If codigoItem = "" Then
            MsgBox "Não foi possível identificar o código do item selecionado.", vbCritical
            Exit Sub
        End If
        
        Call VerificarObrigatoriedade(codigoItem, exigeVencimento, exigeCA, exigeSerie)
        
        If exigeVencimento Then
            If vencimento = "" Or Not IsDate(vencimento) Then
                MsgBox "Este item (" & codigoItem & ") exige data de validade. Informe uma data válida (dd/mm/aaaa).", vbExclamation
                Me.TextVencimentoNovaEntrega.SetFocus
                Exit Sub
            End If
        End If
        
        If exigeCA Then
            If ca = "" Then
                MsgBox "Este item (" & codigoItem & ") exige número do CA. Informe o CA da nova entrega.", vbExclamation
                Me.TextCANovaEntrega.SetFocus
                Exit Sub
            End If
        End If
        If exigeSerie Then
            If serie = "" Then
                MsgBox "Este item (" & codigoItem & ") exige número de Série. Informe a Série da nova entrega.", vbExclamation
                Me.txtSerieNovaEntrega.SetFocus
                Exit Sub
            End If
        End If

        If estadoDev = "" Then
            MsgBox "Informe o estado de devolução do item.", vbExclamation
            Exit Sub
        End If
        
        codigo = codigoItem
        qtdSolicitada = CDbl(Me.txtQtdeBaixa.Value)
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

        If saldo <= 0 And dataBaixa > dataCorte Then
            MsgBox "Item '" & codigo & "' está sem saldo disponível!", vbCritical, "Saldo insuficiente"
            Exit Sub
            
        ElseIf qtdSolicitada > saldo And dataBaixa > dataCorte Then
            MsgBox "Saldo insuficiente para o item '" & codigo & "'!" & vbCrLf & _
                   "Saldo atual: " & saldo & vbCrLf & _
                   "Quantidade solicitada: " & qtdSolicitada, vbCritical, "Saldo insuficiente"
            Exit Sub
        End If
        
    End If
    
    Call frmEpi.AplicarBaixaMaterial(dataBaixa, motivo, novaEntrega, qtde, vencimento, ca, serie, estadoDev)
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'Formatar Data Vencimento Nova Entrega
Private Sub TextVencimentoNovaEntrega_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String
    Dim currentLength As Integer
    
    TextVencimentoNovaEntrega.MaxLength = 10
    currentText = TextVencimentoNovaEntrega.Text
    currentLength = Len(currentText)
    
    Select Case KeyAscii
        Case 8
        Case 13
            SendKeys "{TAB}"
            KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And TextVencimentoNovaEntrega.SelLength = 0 Then
                TextVencimentoNovaEntrega.Text = currentText & "/"
            ElseIf currentLength = 5 And TextVencimentoNovaEntrega.SelLength = 0 Then
                TextVencimentoNovaEntrega.Text = currentText & "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
    
End Sub

Private Sub TextVencimentoNovaEntrega_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(TextVencimentoNovaEntrega.Text) And TextVencimentoNovaEntrega.Text <> "" Then
        
        MsgBox "Data de Vencimento Inválida"
        TextVencimentoNovaEntrega.Text = ""
        Cancel = True
    
    End If
End Sub

'Formatar Data Baixa Item
Private Sub txtDataBaixaItem_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim currentText As String
    Dim currentLength As Integer
    
    txtDataBaixaItem.MaxLength = 10
    currentText = txtDataBaixaItem.Text
    currentLength = Len(currentText)
    
    Select Case KeyAscii
        Case 8
        Case 13
            SendKeys "{TAB}"
            KeyAscii = 0
        Case 48 To 57
            If currentLength = 2 And txtDataBaixaItem.SelLength = 0 Then
                txtDataBaixaItem.Text = currentText & "/"
            ElseIf currentLength = 5 And txtDataBaixaItem.SelLength = 0 Then
                txtDataBaixaItem.Text = currentText & "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
    
End Sub

Private Sub txtDataBaixaItem_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not IsDate(txtDataBaixaItem.Text) And txtDataBaixaItem.Text <> "" Then
        
        MsgBox "Data de Vencimento Inválida"
        txtDataBaixaItem.Text = ""
        Cancel = True
    
    End If
End Sub

