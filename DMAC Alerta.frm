VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmDMACAlerta 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "DMAC Alerta"
   ClientHeight    =   10290
   ClientLeft      =   1350
   ClientTop       =   1485
   ClientWidth     =   15120
   Icon            =   "DMAC Alerta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnima 
      Interval        =   30
      Left            =   18090
      Top             =   2895
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5025
      Top             =   960
   End
   Begin VB.Image imgDivisao 
      Height          =   450
      Left            =   0
      Picture         =   "DMAC Alerta.frx":23FA
      Top             =   0
      Width           =   15360
   End
   Begin VB.Label lblMensagemGeral 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sem conexão com o banco de dados DMAC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   2415
      TabIndex        =   7
      Top             =   5700
      Visible         =   0   'False
      Width           =   9960
   End
   Begin VB.Image imgLogo 
      Height          =   675
      Left            =   8130
      Picture         =   "DMAC Alerta.frx":3714
      Top             =   6585
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgSair 
      Height          =   750
      Left            =   1620
      Picture         =   "DMAC Alerta.frx":52C1
      Top             =   8550
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   6765
      Picture         =   "DMAC Alerta.frx":56BD
      Top             =   10680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblTentativas 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tentativas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   300
      TabIndex        =   6
      Top             =   3675
      Visible         =   0   'False
      Width           =   2055
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   1680
      Left            =   20115
      TabIndex        =   5
      Top             =   7365
      Visible         =   0   'False
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
   Begin VB.Label lblIP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "192.168.1.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image imgStatus 
      Height          =   2295
      Index           =   0
      Left            =   3000
      Picture         =   "DMAC Alerta.frx":72B3
      Top             =   315
      Width           =   840
   End
   Begin VB.Label lblLoja 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NUV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1305
      Index           =   0
      Left            =   195
      TabIndex        =   3
      Top             =   615
      Width           =   2055
   End
   Begin VB.Label lblMSGStatus 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Conexão Normal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   195
      TabIndex        =   2
      Top             =   2265
      Width           =   2055
   End
   Begin VB.Label lblTempoPing 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo: 500ms"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   2655
      Width           =   2055
   End
   Begin VB.Label lblMSGLoja 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   390
      Width           =   2055
   End
   Begin VB.Image imgStatus3 
      Height          =   2295
      Left            =   14190
      Picture         =   "DMAC Alerta.frx":8C48
      Top             =   525
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgStatus1 
      Height          =   2295
      Left            =   12090
      Picture         =   "DMAC Alerta.frx":AB13
      Top             =   465
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgStatus2 
      Height          =   2295
      Left            =   13185
      Picture         =   "DMAC Alerta.frx":C915
      Top             =   495
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgStatus0 
      Height          =   2295
      Left            =   11100
      Picture         =   "DMAC Alerta.frx":E7A6
      Top             =   390
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image imgSOM 
      Height          =   1845
      Index           =   0
      Left            =   735
      Picture         =   "DMAC Alerta.frx":1013B
      Top             =   900
      Width           =   1800
   End
End
Attribute VB_Name = "frmDMACAlerta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QTDENumeroMonitor As Byte
Dim intervalo As Integer
Dim intervaloTexto As Integer


Dim resolucaoOriginal As resolucaoTela

Private Sub carregarLojas()
    Dim adoCNLoja As New ADODB.Connection
    Dim rsDados As New ADODB.Recordset
    Dim sql As String

    Call ConectaODBC(adoCNLoja)
    
        intervaloTexto = 200
        
        lblMSGLoja(0).left = intervaloTexto
        lblLoja(0).left = intervaloTexto
        lblMSGStatus(0).left = intervaloTexto
        lblTempoPing(0).left = intervaloTexto
        
        imgStatus(0).left = (lblMSGLoja(0).left + lblMSGLoja(0).Width) + intervaloTexto
        
        imgStatus(0).top = 300
        lblMSGLoja(0).top = imgStatus(0).top + 50
        lblLoja(0).top = imgStatus(0).top
        lblIP(0).top = (lblLoja(0).top + lblLoja(0).Height)
        lblMSGStatus(0).top = (lblIP(0).top + lblIP(0).Height)
        lblTempoPing(0).top = (lblMSGStatus(0).top + lblMSGStatus(0).Height)
        lblTentativas(0).Caption = 0
        
        imgSOM(0).Visible = False
        imgSOM(0).top = lblLoja(0).top
        imgSOM(0).left = lblLoja(0).left + 280


    sql = "select top 14 rtrim(LO_Loja) as loja, rtrim(LO_IpLoja) as ip from loja where lo_situacao = 'A' and lo_loja not in ('86','185','314','535') and LO_Regiao < 800 order by LO_Regiao, lo_loja"
    rsDados.CursorLocation = adUseClient
    rsDados.Open sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
        'lblLoja(0).Caption = Format(rsDados("Loja"), "000")
        'lblIP(0).Caption = rsDados("ip")
        lblLoja(0).Caption = "INT"
        lblIP(0).Caption = "8.8.8.8"
        

        
        'rsDados.MoveNext
    
    novoMonitor 1
    lblLoja(1).Caption = "CL"
    lblIP(1).Caption = "172.30.5.3"
    
    Do While Not rsDados.EOF
    
        novoMonitor rsDados.AbsolutePosition + 1
        
        lblLoja(rsDados.AbsolutePosition + 1).Caption = Format(rsDados("Loja"), "000")
        lblIP(rsDados.AbsolutePosition + 1).Caption = rsDados("ip")
        
        rsDados.MoveNext
    
    Loop
    

            
            
    
    rsDados.Close
    adoCNLoja.Close
    
End Sub


Private Sub Form_Activate()

    'Me.top = 0

    
    
    Me.Width = (resolucaoTela.Colunas) * 15
    Me.Height = (resolucaoTela.Linhas) * 15
    
    imgSair.left = (Me.Width - imgSair.Width)
    imgSair.top = (Me.Height - imgSair.Height)
    
    
    imgLogo.left = (Me.Width / 2) - (imgLogo.Width / 2)
    imgLogo.top = (Me.Height - imgLogo.Height) - 100
    
    lblMensagemGeral.top = (Me.Height - lblMensagemGeral.Height) - 60
    lblMensagemGeral.left = 0
    lblMensagemGeral.Width = Me.Width
    
    imgDivisao.top = 11200
    imgDivisao.left = 0
    
    animaEntrada
    
    
End Sub

Private Sub animaEntrada()
    Me.top = -Me.Height
    tmrAnima.Enabled = True
    Timer1.Enabled = False
    tmrAnima_Timer
End Sub

Private Sub Form_Load()
    
    
    
    Me.Width = 15440
    Me.Height = 11900
    Me.left = 0
    Me.top = -Me.Height
    
    carregarLojas
    

    'novoMonitor 12
    
End Sub

Private Sub novoMonitor(botao As Byte)

    Dim aux As Double
    Dim aux2 As Double
    Dim linha As Integer

    Load imgStatus(botao)
    Load lblMSGLoja(botao)
    Load lblLoja(botao)
    Load lblIP(botao)
    Load lblMSGStatus(botao)
    Load lblTempoPing(botao)
    Load lblTentativas(botao)
    Load imgSOM(botao)
    
    aux2 = 2850 * (botao \ 4)
    'aux2 = 3500
    
    If (botao \ 4) > 0 Then
        aux = 3850 * (botao - (4 * (botao \ 4)))
        imgStatus(botao).top = aux2 + 300
        lblMSGLoja(botao).top = imgStatus(botao).top + 50
        lblLoja(botao).top = aux2 + 300
        lblIP(botao).top = (lblLoja(botao).top + lblLoja(botao).Height)
        lblMSGStatus(botao).top = (lblIP(botao).top + lblIP(botao).Height)
        lblTempoPing(botao).top = (lblMSGStatus(botao).top + lblMSGStatus(botao).Height)
    Else
        aux = 3850 * botao
    End If
    
    lblMSGLoja(botao).left = aux + intervaloTexto
    lblLoja(botao).left = aux + intervaloTexto
    lblIP(botao).left = aux + intervaloTexto
    lblMSGStatus(botao).left = aux + intervaloTexto
    lblTempoPing(botao).left = aux + intervaloTexto
    imgStatus(botao).left = (lblMSGLoja(botao).left + lblMSGLoja(botao).Width) + intervaloTexto
    
    imgStatus(botao).Visible = True
    lblMSGLoja(botao).Visible = True
    lblLoja(botao).Visible = True
    lblIP(botao).Visible = True
    lblMSGStatus(botao).Visible = True
    lblTempoPing(botao).Visible = True
    
    imgSOM(botao).Visible = False
    imgSOM(botao).top = lblLoja(botao).top
    imgSOM(botao).left = lblLoja(botao).left + 280
End Sub

Private Function pingar(strIPAddress As String) As Integer

   Dim Reply As ICMP_ECHO_REPLY
   Dim lngSuccess As Long
   'Dim strIPAddress As String
   
   'Get the sockets ready.
   If SocketsInitialize() Then
      
    'Address to ping
    'strIPAddress = "8.8.8.8"
    
    'Ping the IP that is passing the address and get a reply.
    lngSuccess = ping(strIPAddress, Reply)
      
    'Display the results.
    Debug.Print "Address to Ping: " & strIPAddress
    Debug.Print "Raw ICMP code: " & lngSuccess
    Debug.Print "Ping Response Message : " & EvaluatePingResponse(lngSuccess)
    Debug.Print "Time : " & Reply.RoundTripTime & " ms"
    
    If lngSuccess <> 0 Then
        pingar = 9999
    Else
        pingar = Reply.RoundTripTime
    End If
      
    'Clean up the sockets.
    SocketsCleanup
      
   Else
   
   'Winsock error failure, initializing the sockets.
   Debug.Print WINSOCK_ERROR
   
   End If
   
End Function

Private Sub Image2_Click()

End Sub

Private Sub Label1_Click()

End Sub



Public Sub Form_Unload(Cancel As Integer)
   
End Sub

Private Sub imgStatus_DblClick(Index As Integer)
    
    'adoCNLoja.Close
    End
End Sub

Private Sub lblLoja_Click(Index As Integer)
    If imgSOM(Index).Visible = True Then
        imgSOM(Index).Visible = False
    Else
        imgSOM(Index).Visible = True
    End If
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Timer1_Timer()
    If glb_monitorarRede = True Then
        Me.Refresh
        'Command1.Visible = True
        atualiza
    Else
        Me.Refresh
        'Command1.Visible = False
    End If
End Sub

Private Sub atualiza()

    Dim i As Byte
    Dim tempo As Integer

    i = 0
    For i = 0 To lblIP.UBound

        tempo = pingar(lblIP(i).Caption)
        lblTempoPing(i).Caption = tempo & " ms"
        
        If tempo > 4000 Then
            imgStatus(i).Picture = imgStatus3
            If lblTentativas(i).Caption >= 99 Then
                lblTentativas(i).Caption = "+99"
            Else
                lblTentativas(i).Caption = lblTentativas(i) + 1
            End If
            lblMSGStatus(i).Caption = "Sem conexão (" & lblTentativas(i).Caption & ")"
            If (Val(lblTentativas(i).Caption) Mod 20) = 0 And imgSOM(i).Visible = False Then
                WindowsMediaPlayer1.URL = "C:\Sistemas\DMAC Alerta\sons\" & lblLoja(i).Caption & ".mp3"
            End If
            lblMSGStatus(i).ForeColor = RGB(255, 255, 255)
        ElseIf tempo > 3000 Then
            lblTentativas(i).Caption = 0
            imgStatus(i).Picture = imgStatus3
            lblMSGStatus(i).Caption = "Alto consumo"
            lblMSGStatus(i).ForeColor = RGB(255, 255, 255)
        ElseIf tempo > 700 Then
            lblTentativas(i).Caption = 0
            imgStatus(i).Picture = imgStatus2
            lblMSGStatus(i).Caption = "Conexão Lenta"
            lblMSGStatus(i).ForeColor = RGB(255, 255, 255)
        Else
            lblTentativas(i).Caption = 0
            imgStatus(i).Picture = imgStatus1
            lblMSGStatus(i).Caption = "Conexão Normal"
            lblMSGStatus(i).ForeColor = RGB(255, 255, 255)
        End If
    
        imgStatus(i).Refresh
        lblMSGStatus(i).Refresh
    
    Next i
    
End Sub


Private Sub tmrAnima_Timer()
    If Me.top < 0 Then
        Me.top = Me.top + ((Me.top * -1) / 10) + 15
    Else
        Timer1.Enabled = True
        Timer1_Timer
        tmrAnima.Enabled = False
        Me.top = 0
    End If
End Sub
