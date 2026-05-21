Attribute VB_Name = "Módulo1"
Sub ExportAllModules()
    Dim vbComp As Object, path As String
    path = ThisWorkbook.path & "\VBA_Export\"
    MkDir path
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 100 Then vbComp.Export path & vbComp.Name & ".frm"
        vbComp.Export path & vbComp.Name & ".bas"
    Next vbComp
End Sub
