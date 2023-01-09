VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMetaMensal2 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   3195
   ClientTop       =   345
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   Picture         =   "frmMetaMensal2.frx":0000
   ScaleHeight     =   7500
   ScaleWidth      =   13080
   Begin VB.Frame frmSOM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   105
      TabIndex        =   7
      Top             =   7950
      Width           =   20685
      Begin VB.Label lblDesativaSom 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Click aqui para desativar o som"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   795
         Left            =   -255
         TabIndex        =   8
         Top             =   120
         Width           =   10755
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer som 
      Height          =   1680
      Left            =   19050
      TabIndex        =   9
      Top             =   420
      Width           =   2235
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   3942
      _cy             =   2963
   End
   Begin VB.Label lblMensagem 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Não há conexão com o servidor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   795
      Left            =   2955
      TabIndex        =   6
      Top             =   6315
      Width           =   12000
   End
   Begin VB.Image imgSemConexao 
      Height          =   11520
      Left            =   15015
      Picture         =   "frmMetaMensal2.frx":7334
      Top             =   510
      Width           =   15360
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   80.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1860
      Left            =   5970
      TabIndex        =   5
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Meta: 3.924.321,21"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   36
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   7635
      TabIndex        =   4
      Top             =   735
      Width           =   12615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Percetual Restante: 18,50%"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   7635
      TabIndex        =   3
      Top             =   2280
      Width           =   5760
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Restante: 890.726,69"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   7635
      TabIndex        =   2
      Top             =   2730
      Width           =   5760
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vendas Atual: 3.924.321,21"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   7635
      TabIndex        =   1
      Top             =   1830
      Width           =   5760
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "110"
      BeginProperty Font 
         Name            =   "Calibri Light"
         Size            =   170.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4410
      Left            =   -2190
      TabIndex        =   0
      Top             =   -105
      Width           =   8115
   End
End
Attribute VB_Name = "frmMetaMensal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim metaMes As Double
'Dim vendaMes As Double
'Dim percentualVenda As Double
'Dim valorRestante As Double
'Dim percentualRestante As Double
'Dim percentualVendaMes As Double
'Dim tocaSOM As Boolean
'
'
'Private Sub Form_Activate()
'    animaEntrada
'End Sub
'
'Private Sub Form_DblClick()
'    End
'End Sub
'
'Private Sub Form_Load()
'
'    Me.left = 0
'    Me.top = 0
'    Me.Width = 1024 * 15
'    Me.Height = 768 * 15
'
'    tocaSOM = False
'
'    imgSemConexao.top = 0
'    imgSemConexao.left = 0
'
'    frmSOM.top = Me.Height - frmSOM.Height
'    frmSOM.left = 0
'
'    FrameNavegador.Width = Me.Width
'    FrameNavegador.Height = Me.Height - 450
'
'    lblDesativaSom.left = 0
'    lblDesativaSom.Width = Me.Width
'
'    lblMensagem.left = 0
'    lblMensagem.Width = Me.Width
'    lblMensagem.top = (Me.Height / 2) - (lblMensagem.Height / 2)
'
'    webNavegador.sErrPrintPath = App.Path & "\errreport.txt"
'    webNavegador.bControlInDevelopmentMode = True
'
'    webNavegador.Nav "c:\sistemas\dmac alerta\metaMes\meta.htm"
'    webNavegador.EmbedIE FrameNavegador.hwnd
'
'    frmSOM.Visible = False
'    imgSemConexao.Visible = False
'    lblMensagem.Visible = False
'End Sub
'
