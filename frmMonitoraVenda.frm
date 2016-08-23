VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Begin VB.Form frmMonitoraVenda 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Monitoramento de Vendas"
   ClientHeight    =   10440
   ClientLeft      =   24480
   ClientTop       =   630
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   ScaleHeight     =   10440
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmPrincipal 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   13000
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   16000
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   6
         Left            =   7620
         TabIndex        =   33
         Top             =   2835
         WhatsThisHelpID =   2940
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   6
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":0000
            TabIndex        =   34
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   4
         Left            =   0
         TabIndex        =   31
         Top             =   2835
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   4
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":1A6D
            TabIndex        =   32
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   8
         Left            =   0
         TabIndex        =   29
         Top             =   5670
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   8
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":34DA
            TabIndex        =   30
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   12
         Left            =   0
         TabIndex        =   27
         Top             =   8505
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   12
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":4F47
            TabIndex        =   28
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   1
         Left            =   3810
         TabIndex        =   25
         Top             =   0
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   1
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":69B4
            TabIndex        =   26
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   2
         Left            =   7620
         TabIndex        =   23
         Top             =   0
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   2
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":8421
            TabIndex        =   24
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   3
         Left            =   11430
         TabIndex        =   21
         Top             =   0
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   3
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":9E8E
            TabIndex        =   22
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   5
         Left            =   3810
         TabIndex        =   19
         Top             =   2835
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   5
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":B8FB
            TabIndex        =   20
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   9
         Left            =   3810
         TabIndex        =   17
         Top             =   5670
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   9
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":D368
            TabIndex        =   18
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   0
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":EDD5
            TabIndex        =   16
            Top             =   -225
            Width           =   4005
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   7
         Left            =   11430
         TabIndex        =   13
         Top             =   2835
         WhatsThisHelpID =   2940
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   7
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":108A8
            TabIndex        =   14
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   10
         Left            =   7620
         TabIndex        =   11
         Top             =   5670
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   10
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":12315
            TabIndex        =   12
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   11
         Left            =   11430
         TabIndex        =   9
         Top             =   5670
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   11
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":13D82
            TabIndex        =   10
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   13
         Left            =   3810
         TabIndex        =   7
         Top             =   8505
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   13
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":157EF
            TabIndex        =   8
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   14
         Left            =   7620
         TabIndex        =   5
         Top             =   8505
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   14
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":1725C
            TabIndex        =   6
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Frame frmLoja 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Index           =   15
         Left            =   11430
         TabIndex        =   3
         Top             =   8505
         Width           =   3740
         Begin MSChartLib.MSChart chrVenda 
            Height          =   2595
            Index           =   15
            Left            =   0
            OleObjectBlob   =   "frmMonitoraVenda.frx":18CC9
            TabIndex        =   4
            Top             =   -225
            Width           =   3735
         End
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   66
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   1
         Left            =   4500
         TabIndex        =   65
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   2
         Left            =   8325
         TabIndex        =   64
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   3
         Left            =   11925
         TabIndex        =   63
         Top             =   2040
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   4
         Left            =   0
         TabIndex        =   62
         Top             =   4875
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   5
         Left            =   4410
         TabIndex        =   61
         Top             =   4875
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   6
         Left            =   8220
         TabIndex        =   60
         Top             =   4875
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   7
         Left            =   12045
         TabIndex        =   59
         Top             =   4875
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   8
         Left            =   0
         TabIndex        =   58
         Top             =   7710
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   9
         Left            =   4695
         TabIndex        =   57
         Top             =   7710
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   10
         Left            =   8100
         TabIndex        =   56
         Top             =   7845
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   11
         Left            =   12030
         TabIndex        =   55
         Top             =   7695
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   12
         Left            =   0
         TabIndex        =   54
         Top             =   10545
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   13
         Left            =   4275
         TabIndex        =   53
         Top             =   10545
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   14
         Left            =   8130
         TabIndex        =   52
         Top             =   10635
         Width           =   1500
      End
      Begin VB.Label lblLoja 
         BackStyle       =   0  'Transparent
         Caption         =   "Loja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   15
         Left            =   12030
         TabIndex        =   51
         Top             =   10530
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   0
         Left            =   2115
         TabIndex        =   50
         Top             =   2130
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   1
         Left            =   6135
         TabIndex        =   49
         Top             =   2085
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   2
         Left            =   10110
         TabIndex        =   48
         Top             =   2205
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   3
         Left            =   13665
         TabIndex        =   47
         Top             =   2205
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   4
         Left            =   1935
         TabIndex        =   46
         Top             =   4935
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   5
         Left            =   6000
         TabIndex        =   45
         Top             =   4920
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   6
         Left            =   10080
         TabIndex        =   44
         Top             =   4965
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   7
         Left            =   13890
         TabIndex        =   43
         Top             =   4950
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   8
         Left            =   1830
         TabIndex        =   42
         Top             =   7740
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   9
         Left            =   5820
         TabIndex        =   41
         Top             =   7785
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   10
         Left            =   9900
         TabIndex        =   40
         Top             =   7770
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   11
         Left            =   13695
         TabIndex        =   39
         Top             =   7725
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   12
         Left            =   1590
         TabIndex        =   38
         Top             =   10455
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   13
         Left            =   5535
         TabIndex        =   37
         Top             =   10605
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   14
         Left            =   9195
         TabIndex        =   36
         Top             =   10590
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   15
         Left            =   13275
         TabIndex        =   35
         Top             =   10545
         Width           =   1500
      End
   End
   Begin VB.Timer tmrAnima 
      Interval        =   30
      Left            =   16440
      Top             =   2595
   End
   Begin WMPLibCtl.WindowsMediaPlayer som2 
      Height          =   1680
      Left            =   16800
      TabIndex        =   67
      Top             =   3780
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
      Left            =   16530
      TabIndex        =   1
      Top             =   2655
      Width           =   12000
   End
   Begin VB.Image imgSemConexao 
      Height          =   11520
      Left            =   18360
      Picture         =   "frmMonitoraVenda.frx":1A736
      Top             =   150
      Visible         =   0   'False
      Width           =   20400
   End
   Begin WMPLibCtl.WindowsMediaPlayer som 
      Height          =   1680
      Left            =   15675
      TabIndex        =   0
      Top             =   5385
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
   Begin VB.Image imgDivisao 
      Height          =   450
      Left            =   15315
      Picture         =   "frmMonitoraVenda.frx":2722C
      Top             =   8070
      Width           =   15360
   End
End
Attribute VB_Name = "frmMonitoraVenda"
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


Option Explicit

Dim numeroColunaGrafico As Byte
Dim metaDiaAtingida As Boolean

Dim adoCNLoja As New ADODB.Connection


Private Sub chrVenda_DblClick(Index As Integer)
    adoCNLoja.Close
    End
End Sub

Private Sub Form_Activate()
    animaEntrada
End Sub


Private Sub Form_Load()

    Call ConectaODBC(adoCNLoja)

    Me.left = 0
    Me.top = -Me.Height
    Me.Width = (resolucaoTela.Colunas) * 15
    Me.Height = (resolucaoTela.Linhas) * 15 + 100
    
    som2.left = Me.Width
    som.left = som2.left
    
    alinhaCompomentes
    carregaValoresFixo
    
End Sub

Private Sub semConexao(ativa As Boolean)
    If ativa = True Then
        frmPrincipal.Visible = False
        imgSemConexao.Visible = True
        lblMensagem.Visible = True
    Else
        frmPrincipal.Visible = True
        imgSemConexao.Visible = False
        lblMensagem.Visible = False
    End If
End Sub

Private Sub alinhaCompomentes()
    Dim i As Byte
    
    For i = 0 To frmLoja.Count - 1
        lblLoja(i).top = frmLoja(i).Height + frmLoja(i).top
        lblLoja(i).left = frmLoja(i).left
        lblLoja(i).Width = frmLoja(i).Width
        lblLoja(i).FontItalic = True
        lblLoja(i).FontSize = 24
        lblLoja(i).Alignment = 0
        lblLoja(i).Caption = " "
        
        lblInfo(i).top = frmLoja(i).Height + frmLoja(i).top + 120
        lblInfo(i).left = frmLoja(i).left
        lblInfo(i).Width = frmLoja(i).Width
        lblInfo(i).FontItalic = True
        lblInfo(i).FontBold = False
        lblInfo(i).FontSize = 12
        lblInfo(i).Alignment = 1
        lblInfo(i).Caption = ""
    Next i
    
    For i = 0 To chrVenda.Count - 1
        chrVenda(i).Width = 3900
    Next i
    
    lblMensagem.Width = Me.Width
    lblMensagem.top = (Me.Height / 2) - (lblMensagem.Height / 2)
    lblMensagem.Visible = False
    lblMensagem.left = 0
    
    imgDivisao.left = 0
    imgDivisao.top = 11200
    
    frmPrincipal.left = 100
    frmPrincipal.top = 100
    frmPrincipal.Visible = True
    
    imgSemConexao.top = 0
    imgSemConexao.left = 0
    imgSemConexao.Visible = False
    
    som.left = Me.Width
    
End Sub

Private Sub colorirGrafico(grafico As MSChart, mensagem As Label, percentualVenda As Double)

If percentualVenda < 30 Then
    With grafico.Plot.SeriesCollection(1)
       .DataPoints(-1).Brush.FillColor. _
    Set 214, 10, 10
    End With
    'mensagem.ForeColor = RGB(255, 0, 0)
    
ElseIf percentualVenda < 70 Then

    With grafico.Plot.SeriesCollection(1)
       .DataPoints(-1).Brush.FillColor. _
    Set 255, 128, 10
    End With
    'mensagem.ForeColor = RGB(255, 128, 0)
    
ElseIf percentualVenda < 100 Then

    With grafico.Plot.SeriesCollection(1)
       .DataPoints(-1).Brush.FillColor. _
    Set 244, 244, 0
    End With
    'mensagem.ForeColor = RGB(255, 255, 0)
    
Else

    With grafico.Plot.SeriesCollection(1)
       .DataPoints(-1).Brush.FillColor. _
    Set 0, 255, 64
    End With
    'mensagem.ForeColor = RGB(0, 255, 0)

End If


End Sub

Private Sub atualizaValores()
    
    Dim rsDados As New ADODB.Recordset
    Dim sql As String
    Dim i As Byte
    Dim i2 As Byte
    Dim j As Byte
    Dim percentual As Double
    
    On Error GoTo trataerro
    
    For j = 0 To 5
        
        i = 0
        chrVenda(0).Row = j + 1
        
        sql = "select lo_regiao as regiao, lo_loja as loja,(select sum(totalnota) from nfcapa where me_loja = LojaVenda and tiponota = 'V' and dataemi = '" & Format(Date, "YYYY/MM/DD") & "' and LojaVenda = me_loja  and hora between '06:00:00' and '" & Val(chrVenda(0).RowLabel) + 2 & ":00:00') as totalvenda," & vbNewLine & _
              "(select sum(totalnota) from nfcapa where me_loja = LojaVenda and tiponota = 'E' and dataemi = '" & Format(Date, "YYYY/MM/DD") & "' and LojaVenda = me_loja  and hora between '06:00:00' and '" & Val(chrVenda(0).RowLabel) + 2 & ":00:00') as totalDevolucao" & vbNewLine & _
              "from meta, loja" & vbNewLine & _
              "where me_mes = '" & Format(Date, "MM") & "' " & vbNewLine & _
              "and ME_ANO = '" & Format(Date, "YYYY") & "'" & vbNewLine & _
              "and lo_loja = me_loja and me_loja not in ('86','185')" & vbNewLine & _
              "union" & vbNewLine & _
              "select '999' as regiao, 'CONSO' as loja, (select sum(totalnota) as totalvenda from nfcapa,meta where tiponota = 'V' and dataemi = '" & Format(Date, "YYYY/MM/DD") & "' and hora between '06:00:00' and '" & Val(chrVenda(0).RowLabel) + 2 & ":00:00' and me_mes = '" & Format(Date, "MM") & "' and ME_ANO = '" & Format(Date, "YYYY") & "' and me_loja not in ('86','185') and me_loja = lojavenda) as totalvenda," & vbNewLine & _
              "(select sum(totalnota) as totalvenda from nfcapa,meta where tiponota = 'E' and dataemi = '" & Format(Date, "YYYY/MM/DD") & "' and hora between '06:00:00' and '" & Val(chrVenda(0).RowLabel) + 2 & ":00:00' and me_mes = '" & Format(Date, "MM") & "' and ME_ANO = '" & Format(Date, "YYYY") & "' and me_loja not in ('86','185') and me_loja = lojavenda) as totalDevolucao" & vbNewLine & _
              "order by regiao,loja"
              
        rsDados.CursorLocation = adUseClient
        rsDados.Open sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        'Debug.Print sql
        
        Do While Not rsDados.EOF
           
            chrVenda(i).Row = j + 1
                
            If IsNull(rsDados("totalvenda")) Then
                chrVenda(i).Data = 0
            Else
                If IsNull(rsDados("totalDevolucao")) Then
                    chrVenda(i).Data = rsDados("totalvenda")
                Else
                    chrVenda(i).Data = rsDados("totalvenda") - rsDados("totalDevolucao")
                End If
            End If
            
            rsDados.MoveNext
            i = i + 1
        
        Loop
        
        rsDados.Close
    
    Next j
    
    i2 = i - 1
    For i = 0 To i2
        percentual = (chrVenda(i).Data / retornaMeta(chrVenda(i))) * 100
        lblInfo(i).Caption = "Met. " & Format(retornaMeta(chrVenda(i)), "#,##") & " (" & Format(percentual, "0.00") & "%)     "
        If i = 15 Then lblInfo(i).Caption = "(" & Format(percentual, "0.00") & "%)     "
        colorirGrafico chrVenda(i), lblInfo(i), percentual
        alertaSonoro i, percentual
        chrVenda(i).chartType = chrVenda(0).chartType
    Next i
    
    semConexao False
    
    Exit Sub
    
trataerro:

    If Err.Number = "-2147467259" Then
        adoCNLoja.Close
        Call ConectaODBC(adoCNLoja)
        lblMensagem.Caption = "Erro ao atualiza valores (Banco de dados)" & vbNewLine & "Tentando conexão novamente..."
    Else
        lblMensagem.Caption = "Erro ao atualiza valores (" & Err.Number & ")"
    End If

    lblMensagem.Caption = "Erro ao atualiza valores (" & Err.Number & ")"
    semConexao True
    
    
End Sub

Private Sub verificaMetaDia()
    Dim rsDados As New ADODB.Recordset
    Dim sql As String
    Dim i As Byte
    Dim j As Byte
    
    On Error GoTo trataerro
    
    sql = "select sum(me_meta / ME_QuantDiasUteisMes) as metaDia, " & vbNewLine & _
          "(select SUM(TOTALNOTA)-(select SUM(TOTALNOTA) from nfcapa,meta " & vbNewLine & _
          "where dataemi = '" & Format(Date, "YYYY/MM/DD") & "' and tiponota = 'E' and ME_Mes = '" & Format(Date, "MM") & "' and ME_Ano = '" & Format(Date, "YYYY") & "' and me_loja = lojavenda) from nfcapa, meta " & vbNewLine & _
          "where dataemi = '" & Format(Date, "YYYY/MM/DD") & "' and tiponota = 'V' and ME_Mes = '" & Format(Date, "MM") & "' and ME_Ano = '" & Format(Date, "YYYY") & "' and me_loja = lojavenda) as vendaDia " & vbNewLine & _
          "from meta where ME_Mes = '" & Format(Date, "MM") & "' and ME_Ano = '" & Format(Date, "YYYY") & "'"
    rsDados.CursorLocation = adUseClient
    rsDados.Open sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsDados.EOF Then
       If ((rsDados("vendaDia") / rsDados("metaDia")) * 100) >= 100 Then
            metaDiaAtingida = True
            glb_tempoPadraoExibicao = 240
            som.URL = "C:\Sistemas\DMAC Alerta\sons\metaDia.mp3"
        Else
            metaDiaAtingida = False
       End If
    End If
    
    semConexao False
    
    Exit Sub
    
trataerro:
    
    If Err.Number = "-2147467259" Then
        adoCNLoja.Close
        Call ConectaODBC(adoCNLoja)
        lblMensagem.Caption = "Erro ao verifica meta dia (Banco de dados)" & vbNewLine & "Tentando conexão novamente..."
    Else
        lblMensagem.Caption = "Erro ao Verifica Meta Dia (" & Err.Number & ")"
    End If
    
    semConexao True
    
End Sub

Private Sub alertaSonoro(posicaoComp As Byte, novoPercentual As Double)
        If Val(novoPercentual) >= 100 _
        And Val(lblInfo(posicaoComp).ToolTipText) < Val(novoPercentual) _
        And Val(lblInfo(posicaoComp).ToolTipText) < 100 _
        And Val(lblInfo(posicaoComp).ToolTipText) > 0 Then
            If posicaoComp <> 15 Then
                som.URL = "C:\Sistemas\DMAC Alerta\sons\meta.wav"
            Else
                ''som2.URL = "C:\Sistemas\DMAC Alerta\sons\metaDia.mp3"
            End If
            lblInfo(posicaoComp).ToolTipText = novoPercentual
            lblInfo(posicaoComp).ForeColor = RGB(0, 255, 64)
            lblLoja(posicaoComp).ForeColor = RGB(0, 255, 64)
        Else
            lblInfo(posicaoComp).ToolTipText = novoPercentual
            lblInfo(posicaoComp).ForeColor = vbWhite
            lblLoja(posicaoComp).ForeColor = vbWhite
        End If
End Sub

Private Function retornaMeta(grid As MSChart) As Double
    retornaMeta = grid.Plot.Axis(VtChAxisIdY).ValueScale.Maximum
End Function

Private Sub carregaValoresFixo()
    
    Dim rsDados As New ADODB.Recordset
    Dim sql As String
    Dim i As Byte
    Dim j As Byte
    
    On Error GoTo trataerro
        
    i = 0
    numeroColunaGrafico = 6
    
    sql = "select top 16 (ME_Meta / ME_QuantDiasUteisMes) as metaDia, lo_regiao as regiao," & vbNewLine & _
          "ME_Loja as loja" & vbNewLine & _
          "from meta,LOJA" & vbNewLine & _
          "where me_mes = '" & Format(Date, "MM") & "'" & vbNewLine & _
          "AND ME_ANO = '" & Format(Date, "YYYY") & "'" & vbNewLine & _
          "AND me_loja = lo_loja and me_loja not in ('86','185')" & vbNewLine & _
          "union" & vbNewLine & _
          "select top 1 sum(ME_Meta/ME_QuantDiasUteisMes) as metaDia,'999' as regiao, 'CONSO' as loja from meta where me_mes = '" & Format(Date, "MM") & "' AND ME_ANO = '" & Format(Date, "YYYY") & "'" & vbNewLine & _
          "ORDER BY REGIAO,loja"
    rsDados.CursorLocation = adUseClient
    rsDados.Open sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    Do While Not rsDados.EOF
    
        With Me.chrVenda(i).Plot.Axis(VtChAxisIdY).ValueScale
            .Auto = False
            .Minimum = 0
            .Maximum = Replace(Format(rsDados("metaDia"), "0.00"), ".", ",")
            .MajorDivision = 4
            .MinorDivision = 4
        End With
             
        lblLoja(i).Caption = lblLoja(i).Caption & Format(rsDados("loja"), "000")
        lblInfo(i).ToolTipText = 0
        
        chrVenda(i).RowCount = numeroColunaGrafico
        
        For j = 0 To chrVenda(i).RowCount - 1
            chrVenda(i).Row = j + 1
            chrVenda(i).Data = 0
            chrVenda(i).RowLabel = 6 + (2 * (j + 1)) & "h"
        Next j
    
        rsDados.MoveNext
        i = i + 1
        
    Loop
    
    For i = i To frmLoja.Count - 1
        lblLoja(i).Visible = False
        frmLoja(i).Visible = False
    Next i
    
    rsDados.Close
    
    semConexao False
    
    Exit Sub
    
trataerro:

    If Err.Number = "-2147467259" Then
        adoCNLoja.Close
        Call ConectaODBC(adoCNLoja)
        lblMensagem.Caption = "Erro ao carrega valores fixo (Banco de dados)" & vbNewLine & "Tentando conexão novamente..."
    Else
        lblMensagem.Caption = "Erro ao carrega valores fixo (" & Err.Number & ")"
    End If

    lblMensagem.Caption = "Erro ao carrega valores fixo (" & Err.Number & ")"
    semConexao True
    
End Sub


Private Sub imgSemConexao_DblClick()
    End
End Sub

Private Sub lblInfo_Click(Index As Integer)
    End
End Sub

Private Sub lblLoja_Click(Index As Integer)
    End
End Sub

Private Sub tmrAnima_Timer()
    If Me.top < 0 Then
        Me.top = Me.top + ((Me.top * -1) / 10) + 10
    Else
        tmrAnima.Enabled = False
        Me.top = 0
        som.SetFocus
        atualizaValores
        If metaDiaAtingida = False Then
             verificaMetaDia
        End If
    End If
End Sub

Private Sub animaEntrada()
    Me.top = -Me.Height
    tmrAnima.Enabled = True
    tmrAnima_Timer
End Sub
