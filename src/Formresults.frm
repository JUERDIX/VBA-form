VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formresults 
   Caption         =   "UserForm1"
   ClientHeight    =   5532
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7728
   OleObjectBlob   =   "Formresults.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "Formresults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Notafinal = 0
    solPregunta1 = ""
    solPregunta2 = ""
    solPregunta3 = ""
    solPregunta4 = ""
    
    Unload datos
    Unload Formpregunta1
    Unload Formpregunta2
    Unload Formpregunta3
    Unload Formpregunta4
    
    Unload Me
    Padre.Show
End Sub

Private Sub CommandButton2_Click()
    End
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    With ntresult
        .Caption = "Results"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Label1
        .Caption = "Question 1 :"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Label2
        .Caption = "Question 2 :"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Label3
        .Caption = "Question 3 :"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Label4
        .Caption = "Question 4 :"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With Label5
        .Caption = "Note"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With CommandButton1
        .Caption = "Back to top"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    With CommandButton2
        .Caption = "Exit"
        .Font.Bold = True
        .Font.Size = 12
    End With
    
    Me.CommandButton2.SetFocus
    
    TextBox1.Text = solPregunta1
    TextBox2.Text = solPregunta2
    TextBox3.Text = solPregunta3
    TextBox4.Text = solPregunta4
    TextBox5.Text = Notafinal
    
End Sub
