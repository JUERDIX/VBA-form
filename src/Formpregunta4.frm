VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formpregunta4 
   Caption         =   "UserForm4"
   ClientHeight    =   3420
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6864
   OleObjectBlob   =   "Formpregunta4.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Formpregunta4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnsiguiente_Click()
    If btnopa.Value = True Then
        solPregunta4 = "Correcto"
        Notafinal = Notafinal + 5
    Else
        solPregunta4 = "Incorrecto"
    End If
    
    Me.Hide
    Formresults.Show
End Sub

Private Sub ntpregunt4_Click()

End Sub

Private Sub UserForm_Initialize()
       With ntenuncia
        .Caption = "*your cuestion*"
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    With ntpregunt4
        .Caption = "Question 4 : "
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
       
    With btnsiguiente
        .Caption = "Results"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    Me.btnsiguiente.SetFocus
    
    
    
End Sub
