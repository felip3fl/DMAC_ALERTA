VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPrevisãoTempo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Previsão Tempo"
   ClientHeight    =   7500
   ClientLeft      =   2640
   ClientTop       =   1860
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnima 
      Interval        =   30
      Left            =   10275
      Top             =   3840
   End
   Begin VB.Frame frmLimitaWEB 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      Begin SHDocVwCtl.WebBrowser webPrevisãoTempo 
         Height          =   3945
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8310
         ExtentX         =   14658
         ExtentY         =   6959
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Image imgSemConexao 
      Height          =   11520
      Left            =   13875
      Picture         =   "frmPrevisaoTempo.frx":0000
      Top             =   -945
      Width           =   15360
   End
End
Attribute VB_Name = "frmPrevisãoTempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________________________________________________________
'   \  ____________________________________________________________________________ \
'    \ \         ____    ____   __      __      ____     ____      ____   __       \ \
'     \ \       / ___\  / ___\ /\ \    /\_\    / __ \  /\___ \    / ___\ /\ \       \ \
'      \ \     /\ \__/ /\ \__/ \ \ \   \/\ \  /\ \_\ \ \/___\ \  /\ \__/ \ \ \       \ \
'       \ \    \ \  __\\ \  _\  \ \ \   \ \ \ \ \  __/   /\_ \ \ \ \  __\ \ \ \       \ \
'        \ \    \ \ \_/ \ \ \/   \ \ \   \ \ \ \ \ \/    \/_\ \ \ \ \ \_/  \ \ \       \ \
'         \ \    \ \ \   \ \ \___ \ \ \___\ \ \ \ \ \       _\_\ \ \ \ \    \ \ \___    \ \
'          \ \    \ \_\   \ \____\ \ \____\\ \_\ \ \_\     /\_____\ \ \_\    \ \____\    \ \
'           \ \    \/_/    \/____/  \/____/ \/_/  \/_/     \/_____/  \/_/     \/____/     \ \
'            \ \                                                                           \ \
'             \ \___________________________________________________________________________\ \
'              \_______________________________________________________________________________\
'
' 2016/05/18

Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim zoomAplicado As Boolean

Private Sub Form_Activate()
    animaEntrada
    glb_tempoPrevisao = glb_tempoPrevisao + 1
    
    webPrevisãoTempo.top = pegaTopPrevisaoINI
    
End Sub

Private Sub Form_Load()
    
    webPrevisãoTempo.Navigate "about:Tabs"
    
    Me.left = 0
    Me.top = 0
    Me.Width = (resolucaoTela.Colunas) * 15
    Me.Height = (resolucaoTela.Linhas) * 15 + 100
    
    frmLimitaWEB.top = 0
    frmLimitaWEB.left = 0
    frmLimitaWEB.Width = Me.Width
    frmLimitaWEB.Height = Me.Height
    
    webPrevisãoTempo.left = -20
    webPrevisãoTempo.Width = Me.Width + 350
    webPrevisãoTempo.Height = Me.Height + 8200
    
    imgSemConexao.left = 0
    imgSemConexao.top = 0
    imgSemConexao.Visible = False
'    webPrevisãoTempo.Navigate "http://g1.globo.com/previsao-do-tempo/sp/sao-paulo.html"
    
    webPrevisãoTempo.Silent = True
    
End Sub

Private Function pegaTopPrevisaoINI()
    Dim Buffer As String * 255
    Call GetPrivateProfileString("Posicao", "PosicaoPrevisao", "", Buffer, 255, App.EXEName & ".ini")
    
    pegaTopPrevisaoINI = left(Buffer, 5)
End Function

Private Sub animaEntrada()
    Me.top = -Me.Height
    tmrAnima.Enabled = True
    tmrAnima_Timer
End Sub



Private Sub tmrAnima_Timer()
    If Me.top < 0 Then
        Me.top = Me.top + ((Me.top * -1) / 10) + 10
    Else
        tmrAnima.Enabled = False
        Me.top = 0
        If glb_tempoPrevisao > 20 Then
            webPrevisãoTempo.Navigate "http://www.msn.com/pt-br/clima/previsao-do-tempo/S%C3%A3o-Paulo,SP,Brasil/we-city--23.563,-46.655?iso=BR"
            glb_tempoPrevisao = 0
        End If
    End If
End Sub

Private Sub webPrevisãoTempo_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If Not zoomAplicado Then
        
        webPrevisãoTempo.ExecWB 63, 2, 140&
        webPrevisãoTempo.Navigate "http://www.msn.com/pt-br/clima/previsao-do-tempo/S%C3%A3o-Paulo,SP,Brasil/we-city--23.563,-46.655?iso=BR"
        
        zoomAplicado = True
    End If
End Sub

