VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMetaMensal1 
   BackColor       =   &H0066D166&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   23460
   ClientTop       =   195
   ClientWidth     =   17235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   17235
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAnima 
      Interval        =   30
      Left            =   7725
      Top             =   9615
   End
   Begin VB.Frame frmSOM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   1020
      TabIndex        =   2
      Top             =   7770
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
         TabIndex        =   3
         Top             =   120
         Width           =   10755
      End
   End
   Begin VB.Frame FrameNavegador 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5835
      Left            =   915
      TabIndex        =   0
      Top             =   495
      Width           =   8115
      Begin DMAC_Alerta.IEalt webNavegador 
         Left            =   7155
         Top             =   2955
         _ExtentX        =   4339
         _ExtentY        =   2858
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer som 
      Height          =   1680
      Left            =   19125
      TabIndex        =   4
      Top             =   90
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
      Left            =   5640
      TabIndex        =   1
      Top             =   5940
      Width           =   12000
   End
   Begin VB.Image imgSemConexao 
      Height          =   11520
      Left            =   11775
      Picture         =   "frmMetaMensal1.frx":0000
      Top             =   1620
      Width           =   15360
   End
End
Attribute VB_Name = "frmMetaMensal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim metaMes As Double
Dim vendaMes As Double
Dim percentualVenda As Double
Dim valorRestante As Double
Dim percentualRestante As Double
Dim percentualVendaMes As Double
Dim diasHTML As String
Dim valoresHTML As String
Dim tocaSOM As Boolean

Private Type coresTela
    cor As ColorConstants
    corHTML50 As String
    corHTML10 As String
End Type

    
Private Sub Form_Activate()
    animaEntrada
End Sub

Private Sub Form_DblClick()
    End
End Sub

Private Function obterCores(cor As String) As coresTela
    
    Select Case UCase(cor)
        Case "RED"
            obterCores.corHTML50 = "#EC8080"
            obterCores.corHTML10 = "#DC1B1A"
            obterCores.cor = RGB(216, 1, 0)
        
        Case "GREEN"
            obterCores.corHTML50 = "#80D980"
            obterCores.corHTML10 = "#1ABA1A"
            obterCores.cor = RGB(0, 178, 0)
        
        Case "YELLOW"
            obterCores.corHTML50 = "#D9D980"
            obterCores.corHTML10 = "#BABB1A"
            obterCores.cor = RGB(178, 179, 0)
            
        Case "ORANGE"
            obterCores.corHTML50 = "#FFBF80"
            obterCores.corHTML10 = "#FF8C1A"
            obterCores.cor = RGB(255, 127, 0)
    End Select

End Function

Private Sub Form_Load()

    Me.left = 0
    Me.top = 0
    Me.Width = 1024 * 15
    Me.Height = 768 * 15
    
    tocaSOM = False
    
    FrameNavegador.top = 0
    FrameNavegador.left = 0

    imgSemConexao.top = 0
    imgSemConexao.left = 0
    
    frmSOM.top = Me.Height - frmSOM.Height
    frmSOM.left = 0

    FrameNavegador.Width = Me.Width
    FrameNavegador.Height = Me.Height - 450
    
    lblDesativaSom.left = 0
    lblDesativaSom.Width = Me.Width
    
    lblMensagem.left = 0
    lblMensagem.Width = Me.Width
    lblMensagem.top = (Me.Height / 2) - (lblMensagem.Height / 2)
    
    webNavegador.sErrPrintPath = App.Path & "\errreport.txt"
    webNavegador.bControlInDevelopmentMode = True
    
    webNavegador.Nav "c:\sistemas\dmac alerta\metaMes\meta.htm"
    webNavegador.EmbedIE FrameNavegador.hwnd
    
    frmSOM.Visible = False
    imgSemConexao.Visible = False
    lblMensagem.Visible = False

End Sub

Private Sub abilitaSOM(portentagem As Double)
    
    If portentagem >= 100 And tocaSOM = False Then
        som.URL = "C:\Sistemas\DMAC Alerta\sons\metaMes.mp3"
        frmSOM.Visible = True
        glb_tempoPadraoExibicao = 180
        tocaSOM = True
    Else
        frmSOM.Visible = False
    End If
    
End Sub

Private Sub calculaValores()

    percentualVenda = (vendaMes / metaMes) * 100
    percentualRestante = 100 - ((vendaMes / metaMes) * 100)
    If percentualRestante < 0 Then percentualRestante = 0
    valorRestante = metaMes - vendaMes
    If valorRestante < 0 Then valorRestante = 0
    
End Sub

Private Sub alteraCores(arquivoHTML As String, cor As String)

    Dim corTela As coresTela
    corTela = obterCores(cor)
    
    Me.BackColor = corTela.cor
    arquivoHTML = Replace(arquivoHTML, "[CORGRAFICOLINHA]", corTela.corHTML50)
    arquivoHTML = Replace(arquivoHTML, "[CORGRAFICOCORPO]", corTela.corHTML10)
    arquivoHTML = Replace(arquivoHTML, "[CORFUNDOTELA]", cor)
    
End Sub

Private Sub obterCorPercetual(arquivoHTML As String, percentual As Double)
    
    If percentual < 50 Then
        Call alteraCores(arquivoHTML, "red")
    ElseIf percentual < 75 Then
        Call alteraCores(arquivoHTML, "orange")
    ElseIf percentual < 100 Then
        Call alteraCores(arquivoHTML, "yellow")
    Else
        Call alteraCores(arquivoHTML, "green")
    End If

End Sub

Private Sub criaArquivoBase(arquivoHTML As String)
    Open "c:\sistemas\dmac alerta\metaMes\meta.htm" For Output As #1
         Print #1, arquivoHTML
    Close #1
End Sub

Private Function obterArquivoBase()

    Dim fso As New FileSystemObject
    Dim mensagemArquivoTXT As TextStream

    Set mensagemArquivoTXT = fso.OpenTextFile _
    ("C:\Sistemas\DMAC Alerta\metaMes\Default.htm")
    obterArquivoBase = mensagemArquivoTXT.ReadAll
    mensagemArquivoTXT.Close
    
End Function

Private Sub verificaMetaDia()

    Dim rsDados As New ADODB.Recordset
    Dim rsDadosGrafico As New ADODB.Recordset
    Dim adoCNLoja As New ADODB.Connection
    
    Dim sql As String
    Dim devolucao As Double
    Dim venda As Double
    Dim dataAtual As String
    
    metaMes = 0
    vendaMes = 0
    diasHTML = ""
    valoresHTML = ""
    dataAtual = Date
'    'dataAtual = "2016/07/30"
    
    On Error GoTo trataerro
    
    Call ConectaODBC(adoCNLoja)
    
    sql = "select day(base.DATAEMI) as dia,(select sum(totalnota) from nfcapa, meta where base.DATAEMI = DATAEMI and tiponota = 'V' and ME_Mes = '" & Format(dataAtual, "MM") & "' and ME_Ano = '" & Format(dataAtual, "YYYY") & "' and me_loja = lojavenda and me_loja not in ('86','185','314')) AS venda," & vbNewLine & _
          "(select sum(totalnota) from nfcapa, meta where   base.DATAEMI = DATAEMI and tiponota = 'E' and ME_Mes = '" & Format(dataAtual, "MM") & "' and ME_Ano = '" & Format(dataAtual, "YYYY") & "' and me_loja = lojavenda and me_loja not in ('86','185','314')) as devolucao" & vbNewLine & _
          "from nfcapa as base where month(dataemi) = '" & Format(dataAtual, "MM") & "' and  year(dataemi) = '" & Format(dataAtual, "YYYY") & "' and tiponota in ('V','E') and lojavenda not in ('86','185','314')  GROUP BY DATAEMI order by DATAEMI"
    rsDados.CursorLocation = adUseClient
    rsDados.Open sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
    
    If Not rsDados.EOF Then
        
        Do While Not rsDados.EOF
        
            vendaMes = vendaMes + rsDados("venda")
            
            venda = rsDados("venda")
            
            If IsNull(rsDados("devolucao")) Then
                devolucao = 0
            Else
                devolucao = rsDados("devolucao")
            End If
            
            diasHTML = diasHTML & "'" & rsDados("dia") & "','|',"
            valoresHTML = valoresHTML & Replace(Val(venda - devolucao), ",", ".") & ", "
            
        
            rsDados.MoveNext
        
        Loop
        
        diasHTML = left(diasHTML, Len(diasHTML) - 5)
        valoresHTML = left(valoresHTML, Len(valoresHTML) - 3)
        
        rsDados.Close
        
        sql = "select sum(ME_Meta) as meta from meta where ME_Mes = '" & Format(dataAtual, "MM") & "' and ME_Ano = '" & Format(dataAtual, "YYYY") & "'"
        rsDados.CursorLocation = adUseClient
        rsDados.Open sql, adoCNLoja, adOpenForwardOnly, adLockPessimistic
        
        metaMes = rsDados("meta")
        
        rsDados.Close
        
    Else
    
        metaMes = 0
        vendaMes = 0
        semConexao True
    
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

Private Sub semConexao(semConecao As Boolean)
    
    FrameNavegador.Visible = Not semConecao
    lblMensagem.Visible = semConecao
    imgSemConexao.Visible = semConecao
    
End Sub

Private Function formataValorExibicao(valor As Double) As String
    formataValorExibicao = Format(valor, "#,##0.00")
End Function

Private Sub imgSemConexao_Click()
    End
End Sub

Private Sub lblDesativaSom_Click()
    som.URL = ""
    frmSOM.Visible = False
End Sub

Private Sub atualizaValores()

    Dim arquivoHTML As String

    verificaMetaDia
    calculaValores
    
    arquivoHTML = obterArquivoBase
    
    obterCorPercetual arquivoHTML, percentualVenda
    
    arquivoHTML = Replace(arquivoHTML, "[VALORES]", valoresHTML)
    arquivoHTML = Replace(arquivoHTML, "[DIAS]", diasHTML)
    arquivoHTML = Replace(arquivoHTML, "[PERCENTUALVENDAMES]", Format(percentualVenda, "00"))
    arquivoHTML = Replace(arquivoHTML, "[VALORMETAMES]", formataValorExibicao(metaMes))
    arquivoHTML = Replace(arquivoHTML, "[VALORMETAATUAL]", formataValorExibicao(vendaMes))
    arquivoHTML = Replace(arquivoHTML, "[VALORMETARESTANTE]", formataValorExibicao(valorRestante))
    arquivoHTML = Replace(arquivoHTML, "[PERCENTUALRESTANTE]", Format(percentualRestante, "0.00"))
    
    criaArquivoBase arquivoHTML
    

    webNavegador.Nav "c:\sistemas\dmac alerta\metaMes\meta.htm"
    webNavegador.EmbedIE FrameNavegador.hwnd
    
    abilitaSOM percentualVenda
    
End Sub

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
        som.SetFocus
        atualizaValores
    End If
End Sub
