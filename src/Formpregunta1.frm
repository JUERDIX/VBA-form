VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formpregunta1 
   Caption         =   "UserForm1"
   ClientHeight    =   4524
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6636
   OleObjectBlob   =   "Formpregunta1.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Formpregunta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnsiguiente_Click()
    If btnopb.Value = True Then
        solPregunta1 = "Correcto"
        Notafinal = Notafinal + 5
    Else
        solPregunta1 = "Incorrecto"
    End If
    
    Me.Hide
    Formpregunta2.Show
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    With ntpregunt1
        .Caption = "Question 1 :"
        .Font.Bold = True
        .Font.Size = 11
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
    
    With ntpregunt1
        .Caption = "Question 1 :"
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    With btnsiguiente
        .Caption = "Next"
        .Font.Bold = True
        .Font.Size = 12
    End With
        
        Me.btnsiguiente.SetFocus
End Sub
