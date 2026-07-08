VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formpregunta3 
   Caption         =   "UserForm3"
   ClientHeight    =   4596
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6924
   OleObjectBlob   =   "Formpregunta3.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Formpregunta3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub n_Click()

End Sub
Private Sub btnsiguiente_Click()
    If btnopc.Value = True Then
        solPregunta3 = "Correcto"
        Notafinal = Notafinal + 5
    Else
        solPregunta3 = "Incorrecto"
    End If
    
    Me.Hide
    Formpregunta4.Show
End Sub

Private Sub UserForm_Initialize()
    With ntenuncia
        .Caption = "*your question*"
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    With pregunta3
        .Caption = "Question 3 : "
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With btnopa
        .Caption = "Altenative A"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With btnopb
        .Caption = "Altenative B"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With btnopc
        .Caption = "Altenative C"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With btnopd
        .Caption = "Altenative D"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    With btnsiguiente
        .Caption = "Next"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
End Sub
