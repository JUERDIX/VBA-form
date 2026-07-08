VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Padre 
   Caption         =   "UserForm1"
   ClientHeight    =   2712
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6600
   OleObjectBlob   =   "Padre.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Padre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btninicio_Click()
    Notafinal = 0
    solPregunta1 = ""
    solPregunta2 = ""
    solPregunta3 = ""
    solPregunta4 = ""
    Me.Hide
    datos.Show
End Sub

Private Sub btnsalir_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    With ntExamen
        .Caption = "Questionnaire"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With btninicio
        .Caption = "Start"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With btnsalir
        .Caption = "Exit"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
End Sub
