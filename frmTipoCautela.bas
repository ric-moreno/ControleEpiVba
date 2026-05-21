VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTipoCautela 
   Caption         =   "Escolher Tipo de Cautela"
   ClientHeight    =   1785
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   3390
   OleObjectBlob   =   "frmTipoCautela.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTipoCautela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TipoEscolhido As String

Private Sub UserForm_Initialize()

optCompleta = False
optTroca = False

End Sub

Private Sub cmdOK_Click()

    If optTroca.Value = True Then
        TipoEscolhido = "Simplificada"
    ElseIf optCompleta.Value = True Then
        TipoEscolhido = "Completa"
    Else
        MsgBox "Selecione um tipo de cautela antes de continuar.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    TipoEscolhido = "Cancelar"
    Me.Hide
End Sub
