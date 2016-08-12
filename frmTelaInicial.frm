VERSION 5.00
Begin VB.Form frmTelaInicial 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3375
   ClientLeft      =   3180
   ClientTop       =   4215
   ClientWidth     =   4590
   LinkTopic       =   "Form2"
   ScaleHeight     =   10575
   ScaleWidth      =   20490
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmeChamaFormulario 
      Interval        =   1000
      Left            =   645
      Top             =   570
   End
   Begin VB.Image imgLogo 
      Height          =   11520
      Left            =   975
      Picture         =   "frmTelaInicial.frx":0000
      Top             =   1365
      Width           =   15360
   End
End
Attribute VB_Name = "frmTelaInicial"
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
' 2016/05/01

Dim tempo As Integer
Dim idFormulario As Integer

Private Sub Form_Load()

   ' Call AlterarResolucao(1024, 768)
    resolucaoTela

    Me.left = 0
    Me.top = 0
    Me.Width = (resolucaoTela.Colunas) * 15
    Me.Height = (resolucaoTela.Linhas) * 15
    
    imgLogo.top = 0
    imgLogo.left = 0
    
    glb_primeiraConexao = True
    idFormulario = 0
    glb_tempoPadraoExibicao = 5
    tempo = glb_tempoPadraoExibicao - 2
       
End Sub

Private Sub tmeChamaFormulario_Timer()

    tempo = tempo + 1
    glb_tempoPrevisao = glb_tempoPrevisao + 1

    If tempo > glb_tempoPadraoExibicao Then
        If chamaFormulario = True Then
            tempo = 0
            idFormulario = idFormulario + 1
        End If
    End If

End Sub

Function chamaFormulario() As Boolean

    Dim Buffer As String * 255

    chamaFormulario = True
    glb_monitorarRede = False
    Call GetPrivateProfileString("Tempo de Exibicao", "Tela" & idFormulario, "", Buffer, 255, App.EXEName & ".ini")
    Select Case idFormulario
        Case 0
            glb_monitorarRede = True
            frmDMACAlerta.Show
            glb_tempoPadraoExibicao = left(Buffer, 2)
        Case 1
            frmMonitoraVenda.Show
            glb_tempoPadraoExibicao = left(Buffer, 2)
        Case 2
            frmPrevisãoTempo.Show
            glb_tempoPadraoExibicao = left(Buffer, 2)
        Case 3
            frmLogo.Show
            glb_tempoPadraoExibicao = left(Buffer, 2)
        Case Else
        glb_primeiraConexao = False
        chamaFormulario = False
        idFormulario = 0
    End Select
    
End Function
