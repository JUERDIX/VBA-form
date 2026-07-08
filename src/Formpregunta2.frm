VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formpregunta2 
   Caption         =   "UserForm2"
   ClientHeight    =   5520
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8712.001
   OleObjectBlob   =   "Formpregunta2.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Formpregunta2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnsalir_Click()
    If btnopd.Value = True Then
        solPregunta2 = "Correcto"
        Notafinal = Notafinal + 5
    Else
        solPregunta2 = "Incorrecto"
    End If
    Me.Hide
    Formpregunta3.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    With ntpregunt2
        .Caption = "Question 2 :"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With ntenuncia
        .Caption = "*your question*"
        .Font.Bold = True
        .Font.Size = 11
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
    
    With btnsalir
        .Caption = "Next"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    Me.btnsalir.SetFocus
        
End Sub
