VERSION 5.00
Begin VB.Form frmExibicao8Transporte 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   30
      TabIndex        =   10
      Top             =   5430
      Width           =   8565
      Begin VB.CommandButton cmd_Imprimir 
         Caption         =   "Imprimi&r"
         Enabled         =   0   'False
         Height          =   675
         Left            =   7200
         Picture         =   "frmExibicao8Transporte.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmd_Primeiro 
         Caption         =   "Primeiro"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2370
         TabIndex        =   16
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmd_Anterior 
         Caption         =   "Anterior"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3345
         TabIndex        =   15
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmd_Proximo 
         Caption         =   "Próximo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4335
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmd_Ultimo 
         Caption         =   "Último"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5310
         TabIndex        =   13
         Top             =   300
         Width           =   975
      End
      Begin VB.CommandButton cmd_importar 
         Caption         =   "&Importar"
         Height          =   675
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1215
      End
      Begin VB.PictureBox PBar11 
         Height          =   285
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   555
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Imp.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   21
         Top             =   900
         Width           =   525
      End
      Begin VB.Label Lbl_lidos 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   690
         TabIndex        =   20
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Regs:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   19
         Top             =   900
         Width           =   525
      End
      Begin VB.Label lbl_registros 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7740
         TabIndex        =   18
         Top             =   900
         Width           =   675
      End
   End
   Begin VB.Image LogoMsb 
      Height          =   735
      Left            =   2070
      Picture         =   "frmExibicao8Transporte.frx":030A
      Stretch         =   -1  'True
      Top             =   390
      Width           =   4275
   End
   Begin VB.Label lblCodigoBarrasC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890123456"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1860
      TabIndex        =   24
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label lblCodigoBarrasD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890123456"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1860
      TabIndex        =   23
      Top             =   4830
      Width           =   5055
   End
   Begin VB.Label lblCodigoBarrase 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   450
      TabIndex        =   22
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Label lbl_DATA_CRIACAO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DATA ATUAL "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3480
      TabIndex        =   9
      Top             =   2190
      Width           =   2910
   End
   Begin VB.Label lbl_PLACA 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PLACA DO VEICULO "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3480
      TabIndex        =   8
      Top             =   2775
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSPORTADORA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3240
      TabIndex        =   7
      Top             =   1260
      Width           =   1830
   End
   Begin VB.Label lblQtd1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PLACA.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1740
      TabIndex        =   6
      Top             =   2730
      Width           =   1500
   End
   Begin VB.Label lblLote1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SEQ....:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1740
      TabIndex        =   5
      Top             =   3300
      Width           =   1485
   End
   Begin VB.Label lbl_NOME_TRANSP 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NOME DA TRANSPORTADORA24"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   7620
   End
   Begin VB.Label lbl_SEQUENCIAL 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SEQUENCIAL "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3480
      TabIndex        =   3
      Top             =   3360
      Width           =   4650
   End
   Begin VB.Label lblCodCliente1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA...:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1740
      TabIndex        =   2
      Top             =   2190
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5085
      Left            =   390
      Top             =   210
      Width           =   8070
   End
   Begin VB.Label lblEmbalagem2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7680
      TabIndex        =   1
      Top             =   5220
      Width           =   45
   End
   Begin VB.Label lblCodigoBarras 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "aAK-3134464654"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3210
      TabIndex        =   0
      Top             =   4140
      Width           =   2565
   End
End
Attribute VB_Name = "frmExibicao8Transporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim cnnLocal As ADODB.Connection
Dim cnn As ADODB.Connection
Dim sAsteristico  As String * 1
Dim sTraco  As String * 1

Private Sub cmd_Anterior_Click()
    If rs.RecordCount > 0 And Not rs.BOF Then
       rs.MovePrevious
       If rs.BOF Then
          rs.MoveFirst
          Exit Sub
       End If
       Me.lbl_NOME_TRANSP.Caption = Trim(rs!Nome_transp)
       Me.lbl_DATA_CRIACAO.Caption = rs!DATA_CRIACAO
       Me.lbl_PLACA.Caption = rs!Placa
       Me.lbl_Sequencial.Caption = rs!Sequencial
       Me.lblCodigoBarras.Caption = Trim(rs!Placa) & sAsteristico & Trim(str(rs!Sequencial))
'       Me.lblCodigoBarrasa1.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrasA.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
'       Me.lblCodigoBarrasB.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
       Me.lblCodigoBarrasC.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
       Me.lblCodigoBarrasD.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrase.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
       
    End If

End Sub

Private Sub cmd_importar_Click()

Dim nx As Integer
Dim Nada As String
Dim RESPOSTA As Integer
Dim CONTA As Double
Dim x As Double
Dim y As Double
Dim CVARIAVEL As String
Dim sql As String
Dim objEtiqueta As Mov_Transporte

Dim ADO_Conection As ADODB.Connection

On Error Resume Next

Nada = objApplication.caminhoImportacao & "\Transp.txt"

If Dir$(Nada) = "" Then
   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Programa Cancelado"
   End
End If

'    Open objApplication.caminhoImportacao & "\etiq.txt" For Random As gNumeroArquivo Len = gTamanhoRegistro
RESPOSTA = MsgBox("Ler dados do Arquivo de transporte?", 20, "Sim/Não?")

Open Nada For Random Access Read Write As #11 Len = Len(Arq_Transporte)

If RESPOSTA = 6 Then
      CONTA = 0
'      PBar1.Value = CONTA
'      PBar1.Visible = True
'      PBar1.Min = 0
      y = LOF(11) / Len(Arq_Transporte)
'      PBar1.Max = y
      Lbl_lidos.Caption = y
      x = 0
      Set objEtiqueta = New Mov_Transporte
      
      For y = 1 To LOF(11) / Len(Arq_Transporte)
        CONTA = CONTA + 1
'        PBar1.Value = CONTA
        Get 11, y, Arq_Transporte
        
        objEtiqueta.setTipoBanco = objApplication.getTipoBanco
        
'        MsgBox "placa a inserir - " & Arq_Transporte.Placa
        
        objEtiqueta.Transp.Item("Placa") = Arq_Transporte.Placa
        objEtiqueta.Transp.Item("Sequencial") = Arq_Transporte.Sequencial
        objEtiqueta.Transp.Item("Tipo_transp") = Arq_Transporte.Tipo_transp
        objEtiqueta.Transp.Item("Tipo_caixa") = Arq_Transporte.Tipo_caixa
        objEtiqueta.Transp.Item("Cod_transp") = Arq_Transporte.Cod_transp
        objEtiqueta.Transp.Item("Nome_transp") = Arq_Transporte.Nome_transp
        objEtiqueta.Transp.Item("Motorista") = Arq_Transporte.Motorista
             
'        MsgBox "placa a preenchida - " & objEtiqueta.Transp.Item("Placa")
        
        objEtiqueta.save (adInserir)
        
'        MsgBox "inseriu no banco de dados "
        
      Next

      Close #11

      Kill Nada
      Call cmd_Primeiro_Click
End If

'PBar1.Visible = False

End Sub

Private Sub cmd_Imprimir_Click()
Dim nx As Integer
Me.Height = 5745 ' impressao

Printer.Copies = 1

rs.MoveFirst
For nx = 1 To rs.RecordCount
    Me.lbl_NOME_TRANSP.Caption = Trim(rs!Nome_transp)
    Me.lbl_DATA_CRIACAO.Caption = rs!DATA_CRIACAO
    Me.lbl_PLACA.Caption = rs!Placa
    Me.lbl_Sequencial.Caption = rs!Sequencial
    Me.lblCodigoBarras.Caption = Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial))
'    Me.lblCodigoBarrasa1.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
'    Me.lblCodigoBarrasA.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
'    Me.lblCodigoBarrasB.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
       Me.lblCodigoBarrasC.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
       Me.lblCodigoBarrasD.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrase.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"

    Me.Refresh
    Me.PrintForm
    rs.MoveNext
Next
Printer.Orientation = 2: Printer.EndDoc
Me.Height = 7185 ' normal

Call Gravar_Dados_Sql

Call Remover_Dados_Access

Me.cmd_Imprimir.Enabled = False
Me.cmd_Primeiro.Enabled = False
Me.cmd_Anterior.Enabled = False
Me.cmd_Proximo.Enabled = False
Me.cmd_Ultimo.Enabled = False
Me.lbl_NOME_TRANSP.Caption = ""
Me.lbl_DATA_CRIACAO.Caption = ""
Me.lbl_PLACA.Caption = ""
Me.lbl_Sequencial.Caption = ""
Me.lblCodigoBarras.Caption = ""
'Me.lblCodigoBarrasB.Caption = ""
'Me.lblCodigoBarrasa1.Caption = ""
'Me.lblCodigoBarrasA.Caption = ""
Me.lblCodigoBarrasC.Caption = ""
Me.lblCodigoBarrasD.Caption = ""
'Me.lblCodigoBarrase.Caption = ""


Me.lbl_registros.Caption = "0"
Me.Lbl_lidos.Caption = "0"

End Sub

Private Sub cmd_Primeiro_Click()

    Dim objEtiqueta As Mov_Transporte
    Dim sql As String
    
    On Error GoTo Erro
    Set objEtiqueta = New Mov_Transporte
    Set cnnLocal = objEtiqueta.getConnection
    
    sql = " SELECT " & _
          " MOV_TRANSPORTE.SEQUENCIAL, " & _
          "MOV_TRANSPORTE.PLACA, " & _
          "MOV_TRANSPORTE.TIPO_TRANSP, " & _
          "MOV_TRANSPORTE.TIPO_CAIXA, " & _
          "MOV_TRANSPORTE.COD_TRANSP, " & _
          "MOV_TRANSPORTE.NOME_TRANSP, " & _
          "MOV_TRANSPORTE.MOTORISTA, " & _
          "MOV_TRANSPORTE.DATA_CRIACAO, " & _
          "MOV_TRANSPORTE.HORA_CRIACAO " & _
          "FROM MOV_TRANSPORTE "

    
    Set rs = New ADODB.Recordset
    rs.Open sql, cnnLocal, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       Me.lbl_NOME_TRANSP.Caption = Trim(rs!Nome_transp)
       Me.lbl_DATA_CRIACAO.Caption = rs!DATA_CRIACAO
       Me.lbl_PLACA.Caption = rs!Placa
       Me.lbl_Sequencial.Caption = rs!Sequencial
       Me.lblCodigoBarras.Caption = Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial))
       
'       Me.lblCodigoBarrasa1.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrasA.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
'       Me.lblCodigoBarrasB.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
       Me.lblCodigoBarrasC.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
       Me.lblCodigoBarrasD.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrase.Caption = "*" & Trim(rs!Placa) & "*" & Trim(Str(rs!Sequencial)) & "*"
       
       Me.cmd_Imprimir.Enabled = True
       Me.cmd_Primeiro.Enabled = True
       Me.cmd_Anterior.Enabled = True
       Me.cmd_Proximo.Enabled = True
       Me.cmd_Ultimo.Enabled = True
       Me.lbl_registros.Caption = rs.RecordCount
    End If
    
    Exit Sub
Erro:
    MsgErro "Ocorreu um erro ao consultar uma Etiqueta de Transportes"

End Sub

Private Sub cmd_Proximo_Click()
    If rs.RecordCount > 0 And Not rs.EOF Then
       rs.MoveNext
       If rs.EOF Then
          rs.MoveLast
          Exit Sub
       End If
       Me.lbl_NOME_TRANSP.Caption = Trim(rs!Nome_transp)
       Me.lbl_DATA_CRIACAO.Caption = rs!DATA_CRIACAO
       Me.lbl_PLACA.Caption = rs!Placa
       Me.lbl_Sequencial.Caption = rs!Sequencial
       Me.lblCodigoBarras.Caption = Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial))
'       Me.lblCodigoBarrasa1.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrasA.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
'       Me.lblCodigoBarrasB.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
       Me.lblCodigoBarrasC.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
       Me.lblCodigoBarrasD.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrase.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
    End If

End Sub

Private Sub cmd_Ultimo_Click()
    If rs.RecordCount > 0 Then
       rs.MoveLast
       Me.lbl_NOME_TRANSP.Caption = Trim(rs!Nome_transp)
       Me.lbl_DATA_CRIACAO.Caption = rs!DATA_CRIACAO
       Me.lbl_PLACA.Caption = rs!Placa
       Me.lbl_Sequencial.Caption = rs!Sequencial
       Me.lblCodigoBarras.Caption = Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial))
'       Me.lblCodigoBarrasa1.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrasA.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
'       Me.lblCodigoBarrasB.Caption = "!" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "!"
       Me.lblCodigoBarrasC.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
       Me.lblCodigoBarrasD.Caption = "*" & Trim(rs!Placa) & sTraco & Trim(str(rs!Sequencial)) & "*"
'       Me.lblCodigoBarrase.Caption = "*" & Trim(rs!Placa) & sAsteristico & Trim(Str(rs!Sequencial)) & "*"
       

    End If

End Sub

Private Sub Form_Activate()
Call Remover_Dados_Access
'Call cmd_Primeiro_Click
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.lbl_NOME_TRANSP.Caption = ""
Me.lbl_DATA_CRIACAO.Caption = ""
Me.lbl_PLACA.Caption = ""
Me.lbl_Sequencial.Caption = ""
Me.lblCodigoBarras.Caption = ""
'Me.lblCodigoBarrasA.Caption = ""
'Me.lblCodigoBarrasB.Caption = ""
'Me.lblCodigoBarrasC.Caption = ""
'Me.lblCodigoBarrasD.Caption = ""
'Me.lblCodigoBarrasa1.Caption = ""
'Me.lblCodigoBarrase.Caption = ""
sAsteristico = "*"
sTraco = "-"
End Sub
Private Function Remover_Dados_Access()
    
    Dim objEtiqueta As Mov_Transporte
    Dim sql As String
    
    On Error GoTo Erro
    Set objEtiqueta = New Mov_Transporte
    Set cnnLocal = objEtiqueta.getConnection
    
    sql = " DELETE * " & _
          "FROM MOV_TRANSPORTE "

    
    Set rs = New ADODB.Recordset
    rs.Open sql, cnnLocal, adOpenStatic, adLockReadOnly
    
    Set objEtiqueta = Nothing
    
    
    Exit Function
Erro:
    MsgErro "Ocorreu um erro ao EXCLUIR registros das Etiqueta de Transportes"


End Function

Private Function Gravar_Dados_Sql()
    Dim nx As Integer
    Dim objEtiqueta As Mov_Transporte
    Dim sql As String
    
    On Error GoTo Erro
    
    
'    Set objApplication = New Application
'    objApplication.cnn
'    objApplication.setTipoBanco = bdSqlServer
'    Set objApplication.cnn = connectionStringSqlServerInterface

'    objApplication.setTipoBancoInterface = bdSqlServer
'    objApplication.AbrirConexao
    
    Set cnn = New ADODB.Connection
    sql = "Provider=SQLOLEDB.1;Persist Security Info=True;Data Source=10.3.0.173;Initial Catalog=teklogix;User ID=teklogix;Password=teklogix;"
    cnn.Open sql

'    Set objEtiquetaControlador = New EtiquetaControlador

''        objApplication.setTipoBancoInterface = bdSqlServer
'        objEtiqueta.setTipoBanco = bdSqlServer
    
    
'    objEtiqueta.setTipoBanco = bdSqlServer
'    Set objEtiqueta = New Mov_Transporte
'    Set objEtiquetaControlador = New EtiquetaControlador
'    Set objPecaAvulsoControlador = New PecaAvulsoControlador
'    Set objTransporteControlador = New Mov_Transporte
'    objApplication.setTipoBancoInterface = bdSqlServer
'    objEtiquetaControlador.setTipoBanco = bdSqlServer
    
    
    
    
'    Set cnnLocal = objTransporteControlador.getConnectionb
    
          
'    Set rs = New ADODB.Recordset

rs.MoveFirst

For nx = 1 To rs.RecordCount
    
    sql = "INSERT INTO MOV_TRANSPORTE ( " & _
          "SEQUENCIAL, " & _
          "PLACA, " & _
          "TIPO_TRANSP, " & _
          "TIPO_CAIXA, " & _
          "COD_TRANSP, " & _
          "NOME_TRANSP, " & _
          "MOTORISTA, " & _
          "DATA_CRIACAO, " & _
          "HORA_CRIACAO"
    sql = sql & " ) VALUES ( " & _
          rs!Sequencial & ", '" & _
          Trim(rs!Placa) & "', '" & _
          Trim(rs!Tipo_transp) & "', '" & _
          Trim(rs!Tipo_caixa) & "', '" & _
          Trim(rs!Cod_transp) & "', '" & _
          Trim(rs!Nome_transp) & "', '" & _
          Trim(rs!Motorista) & "', '" & _
          Format(rs!DATA_CRIACAO, "MM/DD/YYYY") & "', '" & _
          rs!HORA_CRIACAO & "' )"
        
    cnn.Execute sql
    
    rs.MoveNext
Next

    Exit Function
Erro:
    MsgErro "Ocorreu um erro ao INSERIR registros das Etiqueta de Transportes no SQL."

End Function

