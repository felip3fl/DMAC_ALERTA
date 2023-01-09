VERSION 5.00
Begin VB.Form frmLogo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   7365
   ClientTop       =   1965
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnima 
      Interval        =   30
      Left            =   360
      Top             =   405
   End
   Begin VB.Image imgLogo 
      Height          =   11520
      Left            =   -2205
      Picture         =   "frmLogo.frx":0000
      Top             =   -1995
      Width           =   15360
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    animaEntrada
End Sub

Private Sub Form_Load()

    Me.left = 0
    Me.top = -Me.Height
    Me.Width = (resolucaoTela.Colunas) * 15
    Me.Height = (resolucaoTela.Linhas) * 15 + 100

    imgLogo.top = 0
    imgLogo.left = 0
    
End Sub

Private Sub imgLogo_DblClick()
    'adoCNLoja.Close
    End
End Sub

Private Sub tmrAnima_Timer()

    If Me.top < 0 Then
        Me.top = Me.top + ((Me.top * -1) / 10) + 10
    Else
        tmrAnima.Enabled = False
        Me.top = 0
    End If
    
End Sub


Private Sub animaEntrada()
    Me.top = -Me.Height
    tmrAnima.Enabled = True
    tmrAnima_Timer
End Sub
