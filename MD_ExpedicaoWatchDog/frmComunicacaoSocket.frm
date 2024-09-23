VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmComunicacaoSocket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimparDados 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   6840
      TabIndex        =   23
      Top             =   6720
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock wskReceber 
      Left            =   4200
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskEnviar 
      Left            =   3480
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtPortaCliente 
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Text            =   "10200"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Servidor"
      Height          =   2055
      Left            =   8520
      TabIndex        =   14
      Top             =   5280
      Width           =   4575
      Begin VB.CommandButton cmdLimpar 
         Caption         =   "Limpar"
         Height          =   495
         Left            =   3000
         TabIndex        =   22
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtPortaServico 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Text            =   "10201"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdFecharServico 
         Caption         =   "Fechar Serviço"
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton cmdAbrirServidor 
         Caption         =   "Iniciar Serviço"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label lblStatusServico 
         AutoSize        =   -1  'True
         Caption         =   "Offline"
         Height          =   315
         Left            =   3000
         TabIndex        =   19
         Top             =   360
         Width           =   450
      End
      Begin VB.Shape shaStatusServidor 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H80000005&
         Height          =   255
         Left            =   2640
         Shape           =   2  'Oval
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Porta:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.TextBox txtPacotesRecebidos 
      Height          =   4695
      Left            =   9360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   360
      Width           =   3975
   End
   Begin VB.TextBox txtStatus 
      Height          =   4695
      Left            =   6840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton cmdDesconectar 
      Caption         =   "Desconectar"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnviarDados 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtEnviarDados 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   6240
      Width           =   6615
   End
   Begin VB.TextBox txtDadosRecebidos 
      Height          =   4725
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   6615
   End
   Begin VB.TextBox txtIpDestino 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdConectar 
      Caption         =   "Conectar"
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Porta remota:"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Pacotes recebidos"
      Height          =   195
      Left            =   9360
      TabIndex        =   13
      Top             =   120
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Status"
      Height          =   195
      Left            =   6840
      TabIndex        =   11
      Top             =   120
      Width           =   450
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Offline"
      Height          =   315
      Left            =   6360
      TabIndex        =   8
      Top             =   5280
      Width           =   450
   End
   Begin VB.Shape shaStatus 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H80000005&
      Height          =   255
      Left            =   6000
      Shape           =   2  'Oval
      Top             =   5280
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mensagem"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Histórico"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Host remoto:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   900
   End
End
Attribute VB_Name = "frmComunicacaoSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dadosRecebidos As String
Dim objPacote As pacote

Private Sub cmdAbrirServidor_Click()
    'wskReceber.Close
    wskReceber.Protocol = sckTCPProtocol
    wskReceber.LocalPort = txtPortaServico.Text '"10020"
    wskReceber.Listen
    
    changeStatusServicoOnline
    cmdAbrirServidor.Enabled = False
End Sub

Private Sub cmdConectar_Click()
    adicionaLinhaStatus ("Conectando...")
    If wskEnviar.State = sckClosed Then
        'wskEnviar.Close
        wskEnviar.Protocol = sckTCPProtocol
        wskEnviar.LocalPort = "0"
        'wskEnviar.LocalPort = "10010"
        wskEnviar.RemoteHost = txtIpDestino.Text
        wskEnviar.RemotePort = txtPortaCliente.Text ' "10020"
        wskEnviar.Connect
    End If
    
End Sub


Private Sub cmdDesconectar_Click()
    changeStatusOffline
    wskEnviar.Close
    wskEnviar.LocalPort = 0
    wskEnviar.LocalPort = txtPortaCliente.Text '"10020"
    
End Sub

Private Sub cmdEnviarDados_Click()
    On Error GoTo erro
    
    If wskEnviar.State = sckConnected Then
        wskEnviar.SendData txtEnviarDados.Text
        txtEnviarDados.Text = ""
    ElseIf wskEnviar.State = sckConnectionPending Then
        adicionaLinhaStatus "Aguarde o envio da mensagem anterior"
    Else
        adicionaLinhaStatus "Não conectado"
    End If
    
    txtEnviarDados.SetFocus
    
    Exit Sub
erro:
    MsgBox "Ocorreu um erro ao enviar os dados." & vbNewLine & vbNewLine & "Err. N: " & Err.Number & vbNewLine & "Desc: " & Err.Description, vbCritical
End Sub

Private Sub cmdFecharServico_Click()
    wskReceber.Close
    
    changeStatusServicoOffline
    cmdAbrirServidor.Enabled = True
End Sub

Private Sub cmdLimpar_Click()
    txtPacotesRecebidos.Text = ""
End Sub

Private Sub cmdLimparDados_Click()
    txtDadosRecebidos.Text = ""
End Sub

Private Sub Form_Load()
    Set objPacote = New pacote
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wskEnviar.Close
    wskReceber.Close
    
    wskEnviar.LocalPort = 0
    wskReceber.LocalPort = 0
End Sub

Private Sub txtEnviarDados_Change()
    If Trim(txtEnviarDados) <> "" Then
        cmdEnviarDados.Enabled = True
    Else
        cmdEnviarDados.Enabled = False
    End If
End Sub

Private Sub txtEnviarDados_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdEnviarDados_Click
        KeyAscii = 0
    End If
End Sub

Private Sub txtIpDestino_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdConectar_Click
    End If
End Sub

Private Sub wskEnviar_Close()
    adicionaLinhaStatus ("Conexão fechada")
End Sub

Private Sub wskEnviar_Connect()
    changeStatusOnline
    adicionaLinhaStatus ("Conectado")
    'shaStatus.BorderColor = vbGreen
End Sub

Private Sub wskEnviar_DataArrival(ByVal bytesTotal As Long)
    wskEnviar.GetData dadosRecebidos
    
    objPacote.setPacote = dadosRecebidos
    
    txtDadosRecebidos.Text = txtDadosRecebidos.Text & vbNewLine & "[" & Time() & "] - " & dadosRecebidos
    txtPacotesRecebidos.Text = "[" & Time() & "] - " & dadosRecebidos & vbNewLine & txtPacotesRecebidos.Text
End Sub

Private Sub wskEnviar_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Dim descricaoErro As String
    Select Case Number
    Case sckConnectionRefused
        descricaoErro = "Conexão recusada"
    Case Else
        MsgBox "Erro no socket cliente." & vbNewLine & vbNewLine & "Err. N: " & Err.Number & vbNewLine & "Desc: " & Description & vbNewLine & "Scode: " & Scode & vbNewLine & "Origem: " & Source
        adicionaLinhaStatus ("Erro!")
    End Select
    
    If descricaoErro <> "" Then
        adicionaLinhaStatus ("Erro - " & descricaoErro)
    End If
End Sub

Private Sub wskEnviar_SendComplete()
    adicionaLinhaStatus ("Envio completo")
End Sub

Private Sub wskEnviar_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    '
End Sub

Private Sub wskReceber_Close()
    wskReceber.Close
    wskReceber.Listen
End Sub

Private Sub wskReceber_Connect()
    '
End Sub

Private Sub wskReceber_ConnectionRequest(ByVal requestID As Long)
    txtPacotesRecebidos.Text = "(" & Time() & ") - Conexão requisitada" & vbNewLine & txtPacotesRecebidos.Text
    If wskReceber.State = sckListening Then
        wskReceber.Close
        wskReceber.Accept requestID
        txtPacotesRecebidos.Text = "(" & Time() & ") - Conexão estabelecida" & vbNewLine & txtPacotesRecebidos.Text
    End If
End Sub

Private Sub wskReceber_DataArrival(ByVal bytesTotal As Long)
    wskReceber.GetData dadosRecebidos
    
    objPacote.setPacote = dadosRecebidos
    
    txtDadosRecebidos.Text = txtDadosRecebidos.Text & vbNewLine & "[" & Time() & "] - " & dadosRecebidos
    txtPacotesRecebidos.Text = "[" & Time() & "] - " & dadosRecebidos & vbNewLine & txtPacotesRecebidos.Text
    
End Sub

Private Sub adicionaLinhaStatus(ByVal txtLinha As String)
    txtStatus.Text = "[" & Time() & "] - " & txtLinha & vbNewLine & txtStatus.Text
End Sub

Private Sub changeStatusOnline()
    shaStatus.BackColor = vbGreen
    lblStatus.Caption = "Online"
End Sub

Private Sub changeStatusOffline()
    shaStatus.BackColor = vbRed
    lblStatus.Caption = "Offline"
End Sub

Private Sub changeStatusServicoOnline()
    shaStatusServidor.BackColor = vbGreen
    lblStatusServico.Caption = "Online"
End Sub

Private Sub changeStatusServicoOffline()
    shaStatusServidor.BackColor = vbRed
    lblStatusServico.Caption = "Offline"
End Sub

Private Sub wskReceber_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '
End Sub

Private Sub wskReceber_SendComplete()
    '
End Sub

Private Sub wskReceber_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    '
End Sub
