VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} datos 
   Caption         =   "UserForm1"
   ClientHeight    =   4872
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6096
   OleObjectBlob   =   "datos.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub btncodigo_Click()

End Sub

Private Sub btnsalir_Click()
    Unload Me
End Sub

Private Sub btnsiguiente_Click()
    Dim ufila As Integer
    ufila = Cells(Rows.Count, 2).End(xlUp).Row
    ufila = ufila + 1
    
    Cells(ufila, 2).Value = Me.Textnombre.Value
    Cells(ufila, 3).Value = Me.Textapellido.Value
    Cells(ufila, 4).Value = Me.Textescuela.Value
    Cells(ufila, 5).Value = Me.Textcodigo
    
    Me.Hide
    Formpregunta1.Show
End Sub

Private Sub Nombre_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    With textdatos
        .Caption = "Enter your details"
        .Font.Bold = True
        .Font.Size = 10
        .AutoSize = True
        
    End With
    
    With Nombre
        .Caption = "Name:"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With apellido
        .Caption = "Surnames:"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With escuela
        .Caption = "School:"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With codigo
        .Caption = "Code:"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With btnsalir
        .Caption = "Exit"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With btnsiguiente
        .Caption = "Next"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
Me.Textnombre.SetFocus

End Sub
