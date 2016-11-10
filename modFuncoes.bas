Attribute VB_Name = "Module2"
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


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Global ConexaoDLLAdo As New DMACD.conexaoADO

Sub Main()

    verificaAppExecucao
    frmBandeja.Show

End Sub

Public Sub verificaAppExecucao()
    If App.PrevInstance Then
       MsgBox App.EXEName + " Já está executando", vbCritical
       End
    End If
End Sub


Public Function ConectaODBC(ByRef conexao As ADODB.Connection) As Boolean

    On Error GoTo ConexaoErro:
    
    ConectaODBC = False
    
    If ConexaoDLLAdo.abrirConexaoADO(conexao, "SVDMAC", "DMAC") Then
        ConectaODBC = True
        Exit Function
    End If
    
ConexaoErro:
    MsgBox "Erro ao abrir banco de localizacao! "

End Function
