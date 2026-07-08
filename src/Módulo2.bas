Attribute VB_Name = "Módulo2"
Sub ExportarModulos()
    Dim comp As VBComponent
    Dim ruta As String
    ruta = ThisWorkbook.Path & "\src\"
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule
                comp.Export ruta & comp.Name & ".bas"
            Case vbext_ct_ClassModule
                comp.Export ruta & comp.Name & ".cls"
            Case vbext_ct_MSForm
                comp.Export ruta & comp.Name & ".frm"
        End Select
    Next comp
    
    MsgBox "Exportación completa"
End Sub
