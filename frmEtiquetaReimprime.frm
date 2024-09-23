VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEtiquetaReimprime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-impressão das etiquetas"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   9450
   Begin VB.ComboBox cbo_impressora 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   32
      Text            =   "Combo1"
      Top             =   5820
      Width           =   5025
   End
   Begin VB.Frame frm_usuario 
      Caption         =   "Login do usuário"
      Height          =   765
      Left            =   60
      TabIndex        =   29
      Top             =   90
      Width           =   9225
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   1065
         TabIndex        =   0
         Top             =   270
         Width           =   1155
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   390
         Left            =   6750
         TabIndex        =   2
         Top             =   210
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   390
         Left            =   7935
         TabIndex        =   3
         Top             =   225
         Width           =   1140
      End
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   3420
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuário:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Senha:"
         Height          =   195
         Index           =   1
         Left            =   2730
         TabIndex        =   30
         Top             =   315
         Width           =   510
      End
   End
   Begin VB.CommandButton cmd_visualizar 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7380
      Picture         =   "frmEtiquetaReimprime.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Mostrar a etiqueta"
      Top             =   5190
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2745
      Left            =   60
      TabIndex        =   16
      Top             =   2160
      Width           =   9285
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   2445
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   4313
         _Version        =   393216
         Cols            =   9
         ForeColorFixed  =   16711680
         BackColorSel    =   65535
         ForeColorSel    =   65535
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLines       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmEtiquetaReimprime.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdfechar 
      Height          =   735
      Left            =   8400
      Picture         =   "frmEtiquetaReimprime.frx":044B
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Fecha a tela de etiqueta"
      Top             =   5190
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   6360
      Picture         =   "frmEtiquetaReimprime.frx":088D
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprimir"
      Top             =   5190
      Width           =   975
   End
   Begin VB.Frame frm_filtro 
      Caption         =   "Filtro da etiqueta"
      Enabled         =   0   'False
      Height          =   1065
      Left            =   60
      TabIndex        =   10
      Top             =   1020
      Width           =   9285
      Begin VB.TextBox txt_Qtd_Caixa 
         Height          =   315
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Digite a quantidade da etiqueta"
         Top             =   630
         Width           =   675
      End
      Begin VB.CommandButton cmd_librera_Data 
         BackColor       =   &H0000FF00&
         Caption         =   "X"
         Height          =   315
         Left            =   6990
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Fecha a solicitação"
         Top             =   450
         Width           =   285
      End
      Begin VB.TextBox txt_peca 
         Height          =   315
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   6
         ToolTipText     =   "Digite o código da etiqueta"
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txt_Lote 
         Height          =   315
         Left            =   1050
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "Digite o Lote da etiqueta"
         Top             =   630
         Width           =   1335
      End
      Begin VB.TextBox txtsequencial 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1050
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "0000000000"
         ToolTipText     =   "Digie a sequencial da etiqueta"
         Top             =   270
         Width           =   1335
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   495
         Left            =   8610
         Picture         =   "frmEtiquetaReimprime.frx":0B97
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Limpar tela para nova etiqueta"
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   495
         Left            =   8010
         Picture         =   "frmEtiquetaReimprime.frx":0FD9
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Confirma dados do filtro"
         Top             =   270
         Width           =   555
      End
      Begin MSComCtl2.DTPicker dtpDataSelecao 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   5700
         TabIndex        =   8
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   172752897
         CurrentDate     =   37837
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
         Height          =   195
         Left            =   2820
         TabIndex        =   21
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peça :"
         Height          =   195
         Left            =   2820
         TabIndex        =   19
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lote :"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   660
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial :"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lbldatanasc 
         AutoSize        =   -1  'True
         Caption         =   "Data.:"
         Height          =   195
         Left            =   5220
         TabIndex        =   15
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Impressora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   33
      Top             =   5850
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lbl_qtd 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1230
      TabIndex        =   28
      Top             =   5550
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Qtde.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   5520
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl_peca 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1230
      TabIndex        =   26
      Top             =   5295
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lbl_produto 
      AutoSize        =   -1  'True
      Caption         =   "Peça.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   25
      Top             =   5280
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lbl_sequencia 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1230
      TabIndex        =   24
      Top             =   5040
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label label33 
      AutoSize        =   -1  'True
      Caption         =   "Sequência.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   5040
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmEtiquetaReimprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public bAtivo As Boolean
Public bTelaImp As Boolean
Public bJafoi As Boolean
Public nLogin As Integer ' conterá o codigo do usuario, quando confirmar a senha
Public nTipo As Integer ' conterá o tipo do usuario, quando confirmar a senha
Public nMatricula As Double ' conterá a matricula do usuario, quando confirmar a senha
Dim oTela As frmExibicao12
Private ultimoTipoEtiqueta As String

Private Sub btoConfirma_Click()

Dim nx As Integer
'Dim cFields As Collection
Dim sData As String

On Error GoTo Erro

Me.MousePointer = vbHourglass

Me.txtsequencial.Text = Format(Me.txtsequencial.Text, "0000000000")

If Me.dtpDataSelecao.Enabled = False Then
   sData = ""
Else
   sData = Me.dtpDataSelecao.Value
End If

Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_Filtro(sBancoMusashi, _
                                                      Me.txtsequencial.Text, _
                                                      sData, _
                                                      Me.txt_Lote.Text, _
                                                      Me.txt_peca.Text, _
                                                      Me.txt_Qtd_Caixa.Text)

If cRec.RecordCount > 1 Then
   Call carrega_Grid
   Me.cmd_Impressao.Enabled = False
   Me.Grid1.col = 0
   Me.Grid1.Row = 1
   Grid1.col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
   Grid1.col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
   Grid1.col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
   Grid1.col = 0
   Me.Grid1.SetFocus
Else
   Call carrega_Grid
   Me.cmd_Impressao.Enabled = True
   Me.cmd_Visualizar.Enabled = True
   cmd_visualizar_Click
   Me.Grid1.col = 0
   Grid1_Click
   Me.cmd_Impressao.SetFocus
End If
Me.txtsequencial.Enabled = False

Me.MousePointer = vbDefault

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
If Err.Number = 50000 Then
   cmd_limpar_Click
End If
If Err.Number = 50001 Then
   Call Limpar_Grid
   Me.txtsequencial.Enabled = True
   Me.txtsequencial.SetFocus
End If
End Sub

Private Sub cmd_Impressao_Click()

Dim v As Integer
Dim sData As String
Dim nx As Double
Dim x As Printer
               
nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), UCase(Me.cbo_impressora.List(Me.cbo_impressora.ListIndex))) > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "Impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora será - 'ETIQUETA FABRICA'"
   Exit Sub
End If
nx = 0

On Error GoTo ERROR

If Me.dtpDataSelecao.Enabled = False Then
   sData = ""
Else
   sData = Me.dtpDataSelecao.Value
End If

For v = 0 To (Forms.Count - 1)

'    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
'          If objApplication.filial = adMusashiDaAmazonia Then
'             Printer.Orientation = 1
'          Else
'             Printer.Orientation = 2
'          End If
'          frmAvulsoPadraoPonteiro.PrintForm
'          Printer.Orientation = 2 : Printer.EndDoc
'        Exit For
'    End If
    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Or Forms(v).Name = "frmExibicao4MDA" Then
       If objApplication.filial = adMusashiDaAmazonia Then
          Printer.Orientation = 1
       Else
          Printer.Orientation = 2
       End If
       If Forms(v).Name = "frmExibicao4MDA" Then
          frmExibicao4MDA.PrintForm
       Else
          frmAvulsoPadraoPonteiro.PrintForm
       End If
       Printer.Orientation = 2: Printer.EndDoc
       Exit For
    End If
    
    If Forms(v).Name = "frmExibicao2" Then
''''        If ultimoTipoEtiqueta = "F" Then
''''           Call Imprime_Etiqueta_Fiat_FTP
''''        Else
''''           Printer.Orientation = 1
''''           frmExibicao2.PrintForm
''''           oTela.Visible = True
''''           oTela.Show
''''           Printer.Orientation = 2
''''           oTela.PrintForm
''''           oTela.PrintForm
''''           Unload oTela: Set oTela = Nothing
''''           Printer.Orientation = 2 : Printer.EndDoc
''''        End If
            If ultimoTipoEtiqueta = "F" Then
               Call Imprime_Etiqueta_Fiat_FTP
            ElseIf ultimoTipoEtiqueta = "Z" Then
               Printer.Orientation = 1
               frmExibicao2.nTamannhowidth = Me.Width
               frmExibicao2.PrintForm
               Call Imprime_Etiqueta_Fiat_FTP
               Call Imprime_Etiqueta_Fiat_FTP
            Else
               Printer.Orientation = 1
               frmExibicao2.nTamannhowidth = Me.Width
               frmExibicao2.PrintForm
               oTela.Visible = True
               oTela.Show
               Printer.Orientation = 2
               oTela.PrintForm
               oTela.PrintForm
               Unload oTela: Set oTela = Nothing
               Printer.Orientation = 2: Printer.EndDoc
            End If
        Exit For
    End If
'    If Forms(v).Name = "frmExibicao2" Then
'        Printer.Orientation = 1
'        frmExibicao2.PrintForm
'        Printer.Orientation = 2 : Printer.EndDoc
'        Exit For
'    End If
    If Forms(v).Name = "frmExibicao3" Then
        Printer.Orientation = 2
        frmExibicao3.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao5" Then
        Printer.Orientation = 2
        frmExibicao5.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao6" Then
        Printer.Orientation = 2
        frmExibicao6.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7UmProduto" Or Forms(v).Name = "frmExibicao7VariosProdutos" Then
        Printer.Orientation = 2
        frmExibicao7Ref.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
Rem incluido para impressao do novo formulario da mvm international motores
    If Forms(v).Name = "frmExibicao9" Then
        If frmExibicao9.PICT_CRISTAL.Visible = True Then
           Call Imprime_etiqueta_MWM_Cristal
        Else
           Printer.Orientation = 2
           frmExibicao9.PrintForm
           Printer.Orientation = 2: Printer.EndDoc
        End If
        Exit For
    End If
    If Forms(v).Name = "frmExibicao10" Then
        Printer.Orientation = 2
        frmExibicao10.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao11" Then
        Printer.Orientation = 2
        frmExibicao11.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
Next
        
Rem mudar o status desta etiqueta para re-impressa status = 2
Call CCTempneMov_Etiq.Mov_Etiq_Alt_Campos(sBancoMusashi, _
                                          Me.txtsequencial.Text, _
                                          "2", _
                                          "", _
                                          str(nLogin))
        
MsgBox "Re-Impressão concluída com sucesso! a tela impressa da etiqueta será encerrada!", vbOKOnly + vbInformation, "Tarefa Concluída"

Call cmd_limpar_Click

Exit Sub

ERROR:

MsgBox "Erro na impressão deste formulário!"

End Sub

Private Sub cmd_librera_Data_Click()

If Me.dtpDataSelecao.Enabled = False Then
   Me.cmd_librera_Data.BackColor = &HFF00&
   Me.dtpDataSelecao.Enabled = True
Else
   Me.cmd_librera_Data.BackColor = &HFF&
   Me.dtpDataSelecao.Enabled = False
End If

End Sub

Private Sub cmd_limpar_Click()
Call Fechar_Form_Etiqueta

Call Limpar_Grid

Me.txt_peca.Text = ""
Me.txt_Lote.Text = ""
Me.txtsequencial.Text = ""
Me.txt_Qtd_Caixa.Text = ""
Me.dtpDataSelecao.Value = Format(Now(), "dd/mm/yyyy")

Me.lbl_produto.Visible = False
Me.lbl_qtd.Visible = False
Me.lbl_sequencia.Visible = False
Me.lbl_peca.Visible = False
Me.label33.Visible = False
Me.Label5.Visible = False
Me.txtsequencial.Enabled = True

Me.dtpDataSelecao.Enabled = False
Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False

Me.cmd_librera_Data.BackColor = &HFF&

End Sub

Private Sub cmdCancel_Click()
cmd_limpar_Click
Me.frm_filtro.Enabled = False
Me.txtUserName.Enabled = True
Me.txtPassword.Enabled = True
Me.txtPassword.Text = ""
Me.txtUserName.Text = ""
Me.txtUserName.SetFocus
End Sub

'Private Sub optGeral_Click()
'
'    dteEtiquetas.rsEtiquetas.Close
'    dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas"
'    dteEtiquetas.rsEtiquetas.Open
'
'    If dteEtiquetas.rsEtiquetas.RecordCount = 0 Then
'        MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbInformation + vbOKOnly, "Tarefa Concluída"
'        dteEtiquetas.rsEtiquetas.Close
'        If Dir("etiq.txt") = "etiq.txt" Then
'            Close gNumeroArquivo
'            Kill "etiq.txt"
'        End If
'        'Set objApplication = Nothing
'        'End
'        MDIEtiquetas.forcaSaida = True
'        Unload MDIEtiquetas
'    End If
'
'    dteEtiquetas.rsEtiquetas.MoveFirst
'    MostraRegistroAtual
'
'    If dteEtiquetas.rsEtiquetas.RecordCount = 1 Then
'        cmdProximo.Enabled = False
'        cmdUltimo.Enabled = False
'        cmdAnterior.Enabled = False
'        cmdProximo.Enabled = False
'    Else
'        cmdProximo.Enabled = True
'        cmdUltimo.Enabled = True
'        cmdAnterior.Enabled = False
'        cmdPrimeiro.Enabled = False
'    End If
'
'    'Habilita os flags
'    FlagGeral = True
'    FlagPecas = False
'    FlagLote = False
'
'End Sub
'
'End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim rRecLogin As ADODB.Recordset

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set rRecLogin = New ADODB.Recordset

Set rRecLogin = CCTempneUsuario.USUARIO_Confirmar_Login(sBancoMusashi, _
                                                        Me.txtUserName.Text, _
                                                        Me.txtPassword.Text)

nLogin = rRecLogin!codigo
nTipo = rRecLogin!Tipo
nMatricula = rRecLogin!matricula
Me.txtUserName.Enabled = False
Me.txtPassword.Enabled = False
Me.frm_filtro.Enabled = True
Me.txtsequencial.BackColor = &H80000005
'Me.txtsequencial.BackColor = &H8000000F

Me.MousePointer = vbDefault
Me.frm_filtro.Enabled = True
Me.txtsequencial.SetFocus

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
Me.txtUserName.SetFocus
End Sub

Private Sub Form_Activate()

'If bTelaImp Then
'   cmd_limpar_Click
'   bTelaImp = False
'End If

If bAtivo Then Exit Sub

bAtivo = True
bJafoi = False
Call Limpar_Grid
Me.txt_peca.Text = ""
Me.txt_Lote.Text = ""
Me.txtsequencial.Text = ""
Me.txt_Qtd_Caixa.Text = ""
Me.dtpDataSelecao.Value = Format(Now(), "dd/mm/yyyy")
Me.dtpDataSelecao.Enabled = False
Me.cmd_librera_Data.BackColor = &HFF&
Me.txtUserName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > -1 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   If Me.ActiveControl.TabIndex < 1 Then
      If 6 = MsgBox("Deseja realmente sair deste módulo?", 32 + 4) Then
         Unload Me
      End If
   Else
       SendKeys "+{TAB}" ' retornar campo
   End If
End If

End Sub

Private Sub Form_Load()
Dim x As Printer
Dim nx As Integer

nx = 0
For Each x In Printers
    If UCase(Mid$(x.DeviceName, 1, 8)) = "ETIQUETA" Then
       Me.cbo_impressora.AddItem x.DeviceName
    End If
Next

Rem verificar se ha impressoras cadastradas
If Me.cbo_impressora.ListCount = 0 Then
   MsgBox "Impressoras ETIQUETAS,não encontradas no sistema, Favor comunicar ao responsável para adiciona-las no sistema!"
   End
End If
Me.cbo_impressora.ListIndex = 0

Rem verificar a impressora padrão para ser usada neste relatório
For nx = 0 To Me.cbo_impressora.ListCount - 1
    If Trim(UCase(sImpressoraFabrica)) = Trim(UCase(Me.cbo_impressora.List(nx))) Then
       Me.cbo_impressora.ListIndex = nx
    End If
Next

Me.Left = 0
Me.Top = 0
bTelaImp = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bAtivo = False
'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta
End Sub

Private Sub carrega_Grid()
Dim nx As Double
Dim nLinhas As Double

Call Limpar_Grid

Grid1.Row = 1
cRec.MoveFirst
bJafoi = True
For nx = 1 To cRec.RecordCount
    Grid1.col = 0: Grid1.Text = cRec.Fields("sequencia")
    Grid1.col = 1: Grid1.Text = cRec.Fields("Cod_Peca")
    Grid1.col = 2: Grid1.Text = cRec.Fields("Descr_Peca")
    Grid1.col = 3: Grid1.Text = cRec.Fields("cod_Cliente")
    Grid1.col = 4: Grid1.Text = cRec.Fields("Lote")
    Grid1.col = 5: Grid1.Text = cRec.Fields("Qtd_Caixa")
    Grid1.col = 6: Grid1.Text = cRec.Fields("Data_Etiq")
    
    cRec.MoveNext
    If Not cRec.EOF Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Row = Grid1.Row + 1
    End If
Next
bJafoi = False
End Sub

Private Sub Limpar_Grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Grid1.Clear
nLinhas = Grid1.Rows

If Grid1.Rows > 2 Then
   For nx = Grid1.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then Grid1.RemoveItem (nx)
   Next
End If

Grid1.Row = 0
Grid1.col = 0: Grid1.ColWidth(0) = 1000:  Grid1.Text = "SEQ."
Grid1.col = 1:  Grid1.ColWidth(1) = 900: Grid1.Text = "PECA"
Grid1.col = 2: Grid1.ColWidth(2) = 2900: Grid1.Text = "DESCRICAO"
Grid1.col = 3: Grid1.ColWidth(3) = 1500: Grid1.Text = "COD.CLI"
Grid1.col = 4: Grid1.ColWidth(4) = 900: Grid1.Text = "LOTE"
Grid1.col = 5: Grid1.ColWidth(5) = 500: Grid1.Text = "QTD"
Grid1.col = 6: Grid1.ColWidth(6) = 1000: Grid1.Text = "DATA"
Grid1.col = 7: Grid1.ColWidth(7) = 1: Grid1.Text = "VALOR"
Grid1.col = 8: Grid1.ColWidth(8) = 1: Grid1.Text = "VALOR"
Grid1.col = 2: Grid1.BackColor = &H80FFFF

Grid1.Row = 0

Grid1.HighLight = False
Grid1.ColAlignment(0) = flexAlignLeftCenter
Grid1.ColAlignment(1) = flexAlignLeftCenter
Grid1.ColAlignment(2) = flexAlignLeftCenter
Grid1.ColAlignment(3) = flexAlignLeftCenter
Grid1.ColAlignment(4) = flexAlignLeftCenter
Grid1.ColAlignment(5) = flexAlignRightCenter
Grid1.ColAlignment(6) = flexAlignLeftCenter

'Me.lbl_produto.Visible = True
'Me.label33.Visible = True
'Me.Label5.Visible = True
'Me.lbl_qtd.Visible = True
'Me.lbl_sequencia.Visible = True
'Me.lbl_peca.Visible = True

Me.lbl_qtd.Caption = ""
Me.lbl_sequencia.Caption = ""
Me.lbl_peca.Caption = ""

End Sub

Private Sub Grid1_Click()
Me.Grid1.SelectionMode = 1
Me.Grid1.BackColorSel = 33
Grid1.col = 0
If Grid1.Rows > 1 And Len(Trim(Me.Grid1.Text)) > 0 Then
   Grid1.col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
   Grid1.col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
   Grid1.col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
   Grid1.col = 0
   Me.lbl_produto.Visible = True
   Me.lbl_qtd.Visible = True
   Me.lbl_sequencia.Visible = True
   Me.lbl_peca.Visible = True
   Me.label33.Visible = True
   Me.Label5.Visible = True
End If
End Sub

Private Sub Grid1_GotFocus()

Grid1.col = 0
If Grid1.Rows > 1 And Len(Trim(Me.Grid1.Text)) > 0 Then
   Me.lbl_produto.Visible = True
   Me.lbl_qtd.Visible = True
   Me.lbl_sequencia.Visible = True
   Me.lbl_peca.Visible = True
   Me.label33.Visible = True
   Me.Label5.Visible = True
Else
   Me.lbl_produto.Visible = False
   Me.lbl_qtd.Visible = False
   Me.lbl_sequencia.Visible = False
   Me.lbl_peca.Visible = False
   Me.label33.Visible = False
   Me.Label5.Visible = False
End If

End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmd_visualizar_Click
End If
End Sub


Private Sub Grid1_SelChange()
   If Grid1.Rows > 2 Then
      Grid1.col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
      Grid1.col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
      Grid1.col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
      Grid1.col = 0
'      Grid1.BackColorSel = &HFF00&
   End If

End Sub


Private Sub txtsequencial_Change()
If Len(Trim(txtsequencial.Text)) = 10 Then
   btoConfirma_Click
End If

End Sub

Private Sub txtsequencial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txt_Lote.SetFocus
End If
If KeyAscii = 27 Then
   If MsgBox("Deseja sair deste módulo?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbNo Then
      Me.txtsequencial.Text = ""
      Me.txtsequencial.SetFocus
   Else
      Unload Me
   End If
End If

End Sub

Private Sub cmd_visualizar_Click()
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
    
Dim qtdeConteiners As Integer
Dim executaUnloadForm As Boolean
Dim sSeqId As String * 18 ' usado para formatar o campo ID, da MVM
Dim sSequencial As String
Dim nQuantidade As Integer
Dim sdata_aux As String

Me.Grid1.col = 1
Me.Grid1.Row = 1
If Len(Trim(Me.Grid1.Text)) = 0 Then Exit Sub

Rem leitura no banco para emissão da etiqueta
Me.Grid1.col = 0
sSequencial = Format(Grid1.Text, "0000000000")

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar(sBancoMusashi, _
                                               sSequencial)

If cRec.RecordCount = 0 Then
   MsgBox "Registro não encontrado no movimento! A etiqueta não será impressa!", vbInformation + vbOKOnly, "Tarefa com problemas"
   Exit Sub
End If

Me.MousePointer = vbDefault
    
executaUnloadForm = False

'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta
 
'qtdeConteiners = cRec.Fields("Tipo").Value
nQuantidade = cRec.Fields("Qtd_Etiq").Value
ultimoTipoEtiqueta = cRec.Fields("Tipo").Value
'---------------------------------------------------------------------------------------------------
'Atualiza o form frmExibicao de acordo com o tipo
'Se tipo 1 opcao padrão etiqueta pequena
If cRec.Fields("Tipo") = "1" Then
    frmAvulsoPadraoPonteiro.Show
    frmAvulsoPadraoPonteiro.Left = Me.Width
    If cRec.Fields("Cliente") <> "" Then
        frmAvulsoPadraoPonteiro.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmAvulsoPadraoPonteiro.lblDescricao.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmAvulsoPadraoPonteiro.lblLote2.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Peso") <> "" Then
        frmAvulsoPadraoPonteiro.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        frmAvulsoPadraoPonteiro.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmAvulsoPadraoPonteiro.lblPeca.Caption = cRec.Fields("Cod_Peca")
    End If
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
    End If
    
    frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
    frmAvulsoPadraoPonteiro.LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
    
    If (cRec.Fields("Tipo") = "4") Then
       frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = False
       If (Val((cRec.Fields("id_cliente")) = 1) Or (Val(cRec.Fields("id_cliente")) = 2) Or (Val(cRec.Fields("id_cliente")) = 3)) Then
          frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = True
       End If
    End If
    
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
        frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & Trim(frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption) & "*"
        frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & Trim(frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption) & "*"
     Else
        frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
        frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
        frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
     End If
    
End If

'---------------------------------------------------------------------------------------------------
'Se tipo 2 opcao FIAT
If cRec.Fields("Tipo") = "2" Or _
   cRec.Fields("Tipo") = "F" Or _
   cRec.Fields("Tipo") = "Z" Then
    frmExibicao2.nTamannhowidth = Me.Width
    frmExibicao2.Show
    'Mostra FIAT
    frmExibicao2.lblCod_Peca.Caption = cRec.Fields("Cod_Peca")
    If cRec.Fields("Data_Expedicao") <> "" Then
        frmExibicao2.lblDataExpedicao2.Caption = cRec.Fields("Data_Expedicao")
    End If
    If cRec.Fields("Cod_Fornecedor") <> "" Then
        frmExibicao2.lblCodFornec2.Caption = Format(cRec.Fields("Cod_Fornecedor"), "000000")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao2.lblDenominacao2.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Num_Doc_Fiscal") <> "" Then
        frmExibicao2.lblBam2.Caption = cRec.Fields("Num_Doc_Fiscal")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao2.lblDesenho2.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000")
    End If
    
    frmExibicao2.lblCodBarra.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000") & Format(cRec.Fields("Qtd_Caixa"), "00000") _
                                  & cRec.Fields("Cod_Embalagem_pw") & Format(cRec.Fields("Cod_Fornecedor"), "000000")
    frmExibicao2.lblCodBarra2.Caption = frmExibicao2.lblCodBarra.Caption
    'Adicionar o * de inicio e fim
    frmExibicao2.lblCodBarra.Caption = "*" & frmExibicao2.lblCodBarra.Caption & "*"
    frmExibicao2.lblCodBarraCp1.Caption = frmExibicao2.lblCodBarra.Caption
    frmExibicao2.lblCodBarraCp2.Caption = frmExibicao2.lblCodBarra.Caption
    
    If cRec.Fields("Data_Lote") <> "" Then
        frmExibicao2.lblDataProducao2.Caption = cRec.Fields("Data_Lote")
    End If
    If cRec.Fields("Cod_Embalagem") <> "" Then
        frmExibicao2.lblCodEmbalagem2.Caption = cRec.Fields("Cod_Embalagem")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao2.lblNumLote2.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Qtd_Lote") <> "" Then
        frmExibicao2.lblQtdLote2.Caption = cRec.Fields("Qtd_Lote")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        frmExibicao2.lblQtdEmbalagem2.Caption = Format(cRec.Fields("Qtd_Caixa"), "00000")
    End If
    If cRec.Fields("Classe_Funcional") <> "" Then
        frmExibicao2.lblClasseFuncional2.Caption = cRec.Fields("Classe_Funcional")
    End If
    If cRec.Fields("Vinculo") <> "" Then
        frmExibicao2.lblVinculo2.Caption = cRec.Fields("Vinculo")
    End If
    If cRec.Fields("Ind_Suplementar") <> "" Then
        frmExibicao2.lblIndicacaoSuplementar2.Caption = cRec.Fields("Ind_Suplementar")
    End If
    If cRec.Fields("Embarque_Controlado") <> "" Then
        frmExibicao2.lblEmbarqueControlado2.Caption = cRec.Fields("Embarque_Controlado")
    End If
    If cRec.Fields("Desvio") <> "" Then
        frmExibicao2.lblLoteSobDesvio2.Caption = cRec.Fields("Desvio")
    End If
    If cRec.Fields("DUM") <> "" Then
        frmExibicao2.lblDum2.Caption = cRec.Fields("DUM")
    End If
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmExibicao2.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
    Else
        frmExibicao2.lblEmbalagem2.Caption = ""
    End If
Rem acrescentado o pto de entrega - 09-09-2016
    If cRec.Fields("Pto_Entrega") <> "" Then
        frmExibicao2.lblPontoEntrega2.Caption = Trim(cRec.Fields("Pto_Entrega"))
    End If
    
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
        frmExibicao2.lblCodigoBarras.Caption = ""
        frmExibicao2.lblCodigoBarras.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        frmExibicao2.lblCodigoBarrasA.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
        frmExibicao2.lblCodigoBarrasB.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
'            frmExibicao2.lblCodigoBarrasC.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
'            frmExibicao2.lblCodigoBarrasD.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
     Else
        frmExibicao2.lblCodigoBarras.Caption = ""
        frmExibicao2.lblCodigoBarrasA.Caption = ""
        frmExibicao2.lblCodigoBarrasB.Caption = ""
     End If
     
     If cRec.Fields("Tipo") = "2" Then ' aqui será preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana

        Set oTela = New frmExibicao12
        oTela.slbl_01 = IIf(IsNull(cRec.Fields("Cliente")), " ", Trim(cRec.Fields("Cliente")))
        oTela.slbl_02 = IIf(IsNull(cRec.Fields("Pto_Entrega")), " ", Trim(cRec.Fields("Pto_Entrega")))
        oTela.slbl_03 = " "
        oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
        oTela.slbl_07 = " "
        oTela.slbl_08 = " " ' FALTA VER COM MAURO
        oTela.slbl_09 = IIf(IsNull(cRec.Fields("Qtd_Caixa")), " ", Trim(cRec.Fields("Qtd_Caixa")))
        oTela.slbl_10 = IIf(IsNull(cRec.Fields("Descr_Peca")), " ", Trim(cRec.Fields("Descr_Peca")))
        oTela.slbl_11 = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", Trim(cRec.Fields("Cod_Fornecedor")))
        oTela.slbl_12 = IIf(IsNull(cRec.Fields("Cod_Embalagem")), " ", Trim(cRec.Fields("Cod_Embalagem")))
        oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
        oTela.slbl_14 = " "
        oTela.slbl_15 = " "
        oTela.slbl_16 = IIf(IsNull(cRec.Fields("Lote")), " ", Trim(cRec.Fields("Lote")))
        oTela.ldl_usuario.Caption = IIf(IsNull(cRec.Fields("EMBALAGEM")), " ", cRec.Fields("EMBALAGEM"))
        oTela.lbl_Sequencial.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        oTela.lbl_barras1.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.lbl_barras2.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.Visible = False
        
     End If

     If cRec.Fields("Tipo") = "Z" Then ' aqui será preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana

        Set oTela = New frmExibicao12
        oTela.slbl_01 = IIf(IsNull(cRec.Fields("Cliente")), " ", Trim(cRec.Fields("Cliente")))
        oTela.slbl_02 = IIf(IsNull(cRec.Fields("Pto_Entrega")), " ", Trim(cRec.Fields("Pto_Entrega")))
        oTela.slbl_03 = " "
        oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
        oTela.slbl_07 = " "
        oTela.slbl_08 = " " ' FALTA VER COM MAURO
        oTela.slbl_09 = IIf(IsNull(cRec.Fields("Qtd_Caixa")), " ", Trim(cRec.Fields("Qtd_Caixa")))
        oTela.slbl_10 = IIf(IsNull(cRec.Fields("Descr_Peca")), " ", Trim(cRec.Fields("Descr_Peca")))
        oTela.slbl_11 = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", Trim(cRec.Fields("Cod_Fornecedor")))
        oTela.slbl_12 = IIf(IsNull(cRec.Fields("Cod_Embalagem")), " ", Trim(cRec.Fields("Cod_Embalagem")))
        oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
        oTela.slbl_14 = " "
        oTela.slbl_15 = " "
        oTela.slbl_16 = IIf(IsNull(cRec.Fields("Lote")), " ", Trim(cRec.Fields("Lote")))
        oTela.ldl_usuario.Caption = IIf(IsNull(cRec.Fields("EMBALAGEM")), " ", cRec.Fields("EMBALAGEM"))
        oTela.lbl_Sequencial.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        oTela.lbl_barras1.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.lbl_barras2.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.Visible = False
     End If

End If

'---------------------------------------------------------------------------------------------------
'Se tipo 3 opcao FORD
If cRec.Fields("Tipo") = "3" Then
  
        sdata_aux = Mid$(Trim(cRec.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(cRec.Fields("data_etiq")), 4, 2))) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 7, 4)
        frmExibicao3.lbl_Cliente.Caption = "MUSASHI DO BRASIL LTDA" 'Trim(crec.Fields("Cliente"))
        frmExibicao3.lbl_Cod_Fornecedor.Caption = "CFEOA"  'Left(crec.Fields("Cod_Fornecedor"), 16)
        frmExibicao3.lbl_Cod_Fornecedor_Barras.Caption = "CFEOA" 'Trim(crec.Fields("Cod_Fornecedor"))
        frmExibicao3.lbl_qtd.Caption = Left(Trim(cRec.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_qtd_barras.Caption = Left(Trim(cRec.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_peso.Caption = Left(Trim(cRec.Fields("Peso")), 16)
        frmExibicao3.lbl_container.Caption = "KLT 4314 CFEOA" 'Left(Trim(crec.Fields("Lote"), 16))
        frmExibicao3.lbl_lote.Caption = Trim(cRec.Fields("Lote"))
        frmExibicao3.lbl_data.Caption = sdata_aux
        frmExibicao3.lbl_Cod_cliente.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_1.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_Barras.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_Cod_Peca.Caption = Trim(cRec.Fields("Cod_Peca"))
        frmExibicao3.lbl_line_feed_loc2.Caption = Trim(cRec.Fields("Cod_Peca")) & " " & Trim(cRec.Fields("Lote"))

        frmExibicao3.lbl_descr_peca.Caption = Trim(cRec.Fields("Descr_Peca"))
        frmExibicao3.lbl_id_etiqueta.Caption = Left(Trim(cRec.Fields("id_etiqueta")), 16)
        frmExibicao3.lbl_id_etiqueta_barra.Caption = Left(Trim(cRec.Fields("id_etiqueta")), 16)
        sdata_aux = Mid$(Trim(cRec.Fields("data_etiq")), 9, 2) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 4, 2) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 1, 2)
        frmExibicao3.DataToEncodeText.Text = Format(cRec.Fields("id_etiqueta"), "00000") & " (P)" & _
                                             Format(Trim(cRec.Fields("Qtd_Caixa")), "0") & " (Q)" & _
                                             Trim(cRec.Fields("Cod_Fornecedor")) & " (V)" & _
                                             sdata_aux & " (D)" & _
                                             Format(Trim(cRec.Fields("Serial")), "000") & " (S)"
        
        frmExibicao3.lbl_to.Caption = "FORD TAUBATE"
        frmExibicao3.lbl_cust.Caption = "FI05D"
        frmExibicao3.lbl_doc_code.Caption = "R3"
    
        frmExibicao3.PDF1.DataToEncode = frmExibicao3.DataToEncodeText.Text
End If

'---------------------------------------------------------------------------------------------------
'Se tipo 4 opcao padrão etiqueta grande
If cRec.Fields("Tipo") = "4" Or cRec.Fields("Tipo") = "8" Then
    'frmExibicao4.Show
    frmAvulsoPadraoPonteiro.Show
    frmAvulsoPadraoPonteiro.Left = Me.Width
    
    If cRec.Fields("Tipo") = "4" Then
       If cRec.Fields("indforjimport") = "X" Then
          frmAvulsoPadraoPonteiro.lbl_kms.Visible = True
          frmAvulsoPadraoPonteiro.lbl_tarja_kms.Visible = True
       Else
          frmAvulsoPadraoPonteiro.lbl_kms.Visible = False
          frmAvulsoPadraoPonteiro.lbl_tarja_kms.Visible = False
       End If
    End If
    If cRec.Fields("Cliente") <> "" Then
        'frmExibicao4.lblCliente.Caption = cRec.Fields("Cliente")
        frmAvulsoPadraoPonteiro.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        'frmExibicao4.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
        frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        'frmExibicao4.lblDescricao.Caption = cRec.Fields("Descr_Peca")
        frmAvulsoPadraoPonteiro.lblDescricao.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Lote") <> "" Then
        'frmExibicao4.lblLote2.Caption = cRec.Fields("Lote")
        frmAvulsoPadraoPonteiro.lblLote2.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Peso") <> "" Then
        'frmExibicao4.lblPeso2.Caption = cRec.Fields("Peso")
        frmAvulsoPadraoPonteiro.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        'frmExibicao4.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
        frmAvulsoPadraoPonteiro.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        'frmExibicao4.lblPeca.Caption = cRec.Fields("Cod_Peca")
        frmAvulsoPadraoPonteiro.lblPeca.Caption = cRec.Fields("Cod_Peca")
    End If
'    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
'        frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
'    End If

    frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = nMatricula
    frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
    frmAvulsoPadraoPonteiro.LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
    
    frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = False
'    MsgBox "cliente de numero - " & Str(Val(cRec.Fields("id_cliente")))
    If Val(cRec.Fields("id_cliente")) = 1 Or _
       Val(cRec.Fields("id_cliente")) = 2 Or _
       Val(cRec.Fields("id_cliente")) = 3 Then
       frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = True
    End If
    
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
       frmAvulsoPadraoPonteiro.lbl_Seq_Milhar.Caption = Mid$(Trim(Format(cRec.Fields("ID_ETIQUETA"), "0000000000")), _
                                                        Len(Trim(Format(cRec.Fields("ID_ETIQUETA"), "0000000000"))) - 5, _
                                                        Len(Trim(Format(cRec.Fields("ID_ETIQUETA"), "0000000000"))))
       frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
       frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
       frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
    Else
       frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
       frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
       frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
    End If
     
    If (cRec.Fields("Tipo") = "8") Then
       frmAvulsoPadraoPonteiro.lblCodCliente1.Caption = "CODE:"
       frmAvulsoPadraoPonteiro.lblQtd1.Caption = "QTY.:"
       frmAvulsoPadraoPonteiro.lblPeso1.Caption = "WEIGHT:"
       frmAvulsoPadraoPonteiro.lblLote1.Caption = "LOT.:"
    End If
    
End If

'---------------------------------------------------------------------------------------------------
'Se tipo M opcao padrão etiqueta NOVA MDA
If cRec.Fields("Tipo") = "M" Then
    frmExibicao4MDA.Show
    frmExibicao4MDA.Left = Me.Width
    
    If cRec.Fields("Tipo") = "4" Then
       If cRec.Fields("indforjimport") = "X" Then
          frmExibicao4MDA.lbl_kms.Visible = True
          frmExibicao4MDA.lbl_tarja_kms.Visible = True
       Else
          frmExibicao4MDA.lbl_kms.Visible = False
          frmExibicao4MDA.lbl_tarja_kms.Visible = False
       End If
    End If
    If cRec.Fields("Cliente") <> "" Then
        frmExibicao4MDA.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao4MDA.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
        frmExibicao4MDA.LBL_CODIGO_NOVO.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao4MDA.lblDescricao.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao4MDA.lblLote2.Caption = cRec.Fields("Lote")
        frmExibicao4MDA.LBL_LOTE_NOVO.Caption = "*" & Trim(cRec.Fields("Lote")) & "*"
    End If
    If cRec.Fields("Peso") <> "" Then
        frmExibicao4MDA.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        frmExibicao4MDA.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
        frmExibicao4MDA.LBL_QTDE_NOVO.Caption = "*" & cRec.Fields("Qtd_Caixa") & "*"
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmExibicao4MDA.lblPeca.Caption = cRec.Fields("Cod_Peca")
    End If

    frmExibicao4MDA.lblEmbalagem2.Caption = nMatricula
    frmExibicao4MDA.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
    frmExibicao4MDA.LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
    
    frmExibicao4MDA.lblMsgProduto.Visible = False
'    MsgBox "cliente de numero - " & Str(Val(cRec.Fields("id_cliente")))
    If Val(cRec.Fields("id_cliente")) = 1 Or _
       Val(cRec.Fields("id_cliente")) = 2 Or _
       Val(cRec.Fields("id_cliente")) = 3 Then
       frmExibicao4MDA.lblMsgProduto.Visible = True
    End If
    
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
       frmExibicao4MDA.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
       frmExibicao4MDA.lblCodigoBarrasA.Caption = "*" & frmExibicao4MDA.lblCodigoBarras.Caption & "*"
       frmExibicao4MDA.lblCodigoBarrasB.Caption = "*" & frmExibicao4MDA.lblCodigoBarras.Caption & "*"
    Else
       frmExibicao4MDA.lblCodigoBarras.Caption = ""
       frmExibicao4MDA.lblCodigoBarrasA.Caption = ""
       frmExibicao4MDA.lblCodigoBarrasB.Caption = ""
    End If
    
End If

'---------------------------------------------------------------------------------------------------
'Se tipo 5 opcao GM etiqueta grande
'MODIFICADO EM 23-08-2007 (MARCOS PEDROSA) ACRESCENTAR CAMPOS LAY-OUT NOVO
If cRec.Fields("Tipo") = "5" Then
    frmExibicao5.tWithOpcoes = Me.Width
    frmExibicao5.Show
    If cRec.Fields("Cliente") <> "" Then
        frmExibicao5.lblTo.Caption = cRec.Fields("Cliente")
    End If
    Rem incluido (marcos pedrosa) em 23-08-2007
    If cRec.Fields("MOTIVO_ALTERACAO_OUTROS") <> "" Then
        frmExibicao5.lblMotivo_alteracao_outros.Caption = Trim(cRec("MOTIVO_ALTERACAO_OUTROS").Value)
    End If
    If cRec.Fields("Ind_Suplementar") <> "" Then
        frmExibicao5.lblPlant.Caption = cRec.Fields("Ind_Suplementar")
    End If
    Rem DEFINICAO ATE 23-08-2007 (MARCOS PEDROSA)
'''        If cRec.Fields("Cod_Util") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Util")
'''        End If
'''        If cRec.Fields("Desvio") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
'''        End If
    Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
        frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Embalagem_Pw")
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
    If cRec.Fields("Desvio") <> "" Then
        frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
    End If

    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao5.lblPartNumber.Caption = cRec.Fields("Cod_Cliente")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao5.lblPeca.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Qtd_Caixa") <> 0 Then
        frmExibicao5.lblQtd.Caption = cRec.Fields("Qtd_Caixa")
    End If
    If cRec.Fields("Modelo") <> "" Then
        frmExibicao5.lblMaterial.Caption = cRec.Fields("Modelo")
    End If
    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
        frmExibicao5.lblReference.Caption = cRec.Fields("Cod_Embalagem_Pw")
    End If
'    If cRec.Fields("Pto_Entrega") <> "" Then
'        frmExibicao5.lblLicense.Caption = cRec.Fields("Pto_Entrega")
'        frmExibicao5.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'        frmExibicao5.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'    End If
    Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Cod_Embalagem") <> "" Then
'''            frmExibicao5.lblContainerType.Caption = cRec.Fields("Cod_Embalagem")
'''        End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("compl_peca1") <> "" Then
        frmExibicao5.lblContainerType.Caption = cRec.Fields("compl_peca1")
    End If
    If cRec.Fields("Peso") <> 0 Then
        frmExibicao5.lblgrossWeight.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Embarque_Controlado") <> "" Then
'''            frmExibicao5.lblRoute.Caption = cRec.Fields("Embarque_Controlado")
'''        End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("compl_peca2") <> "" Then
        frmExibicao5.lblRoute.Caption = cRec.Fields("compl_peca2")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao5.lblLot.Caption = cRec.Fields("Lote")
    End If
    Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Dum") <> "" Then
'''            frmExibicao5.lblEng.Caption = cRec.Fields("Dum")
'''        End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("data_lote") <> "" Then
        frmExibicao5.lblEng.Caption = Mid$(cRec.Fields("data_lote"), 1, 2) & _
                                      Pega_Mes(Val(Mid$(cRec.Fields("data_lote"), 4, 2))) & _
                                      Mid$(cRec.Fields("data_lote"), 7, 4)
    End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("envio_lote") = "1" Then
        frmExibicao5.lblvalidade.Caption = "N"
    Else
        frmExibicao5.lblvalidade.Caption = ""
    End If
    
    If cRec.Fields("Data_expedicao") <> "" Then
        frmExibicao5.lblMfgDate.Caption = Mid$(cRec.Fields("Data_expedicao"), 1, 2) & _
                                          Pega_Mes(Val(Mid$(cRec.Fields("Data_expedicao"), 4, 2))) & _
                                          Mid$(cRec.Fields("Data_expedicao"), 7, 4)
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmExibicao5.lblCodMSB.Caption = cRec.Fields("Cod_Peca")
    End If
'    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
'        frmExibicao5.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
'    End If
    frmExibicao5.lblEmbalagem2.Caption = nMatricula
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
        frmExibicao5.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
        frmExibicao5.lblCodigoBarrasA.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
        frmExibicao5.lblCodigoBarrasB.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
     Else
        frmExibicao5.lblCodigoBarras.Caption = ""
        frmExibicao5.lblCodigoBarrasA.Caption = ""
        frmExibicao5.lblCodigoBarrasB.Caption = ""
     End If
End If

'-----------------------------------------------------------------------------------------------------
'Se tipo 6 opcao Cliente MVM INTERNACIONAL MOTORES
If cRec.Fields("Tipo") = "6" Or cRec.Fields("Tipo") = "9" Then
    frmExibicao9.Show
    frmExibicao9.Left = Me.Width
    If cRec.Fields("Tipo") = "9" Then
       frmExibicao9.PICT_CRISTAL.Visible = True
    Else
       frmExibicao9.PICT_CRISTAL.Visible = False
    End If
    If cRec.Fields("Cliente") <> "" Then
        frmExibicao9.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao9.lblCodBar_Cod_cliente.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
        frmExibicao9.lblCodBar_Cod_cliente1.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
        frmExibicao9.lbl_Cod_cliente.Caption = Trim(cRec.Fields("Cod_Cliente"))
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao9.lbl_desc_peca.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("cod_fornecedor") <> "" Then
        frmExibicao9.lblCod_Fornecedor.Caption = cRec.Fields("cod_fornecedor")
    End If

    frmExibicao9.lbl_Fornecedor.Caption = "MUSASHI"
    If cRec.Fields("data_expedicao") <> "" Then
        frmExibicao9.lbl_data_expedicao.Caption = cRec.Fields("data_expedicao")
    End If
    If cRec.Fields("qtd_caixa") <> "" Then
        frmExibicao9.lblCodBarQtd_caixa.Caption = "*" & cRec.Fields("qtd_caixa") & "*"
        frmExibicao9.lblqtd_caixa.Caption = cRec.Fields("qtd_caixa")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao9.lblCodBar_lote.Caption = "*" & cRec.Fields("Lote") & "*"
        frmExibicao9.lblLote.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmExibicao9.lbl_id_etiqueta.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
        frmExibicao9.lblCodBar_Cod_Peca.Caption = "*" & cRec.Fields("id_etiqueta") & "*"
        frmExibicao9.lbl_Cod_Peca.Caption = cRec.Fields("cod_peca")
    End If
    
'    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
'        frmExibicao9.lblCodFunc.Caption = cRec.Fields("EMBALAGEM")
'    Else
'        frmExibicao9.lblCodFunc.Caption = ""
'    End If
    frmExibicao9.lblCodFunc.Caption = nMatricula
    
    Rem preparar o "id" da etiqueta com sua formatacao conforme documento fornecedor(6)+dtEtiqueta(6-aammdd)+seq. do dia
    sSeqId = " "
    
    If Not IsNull(cRec.Fields("desvio_aviso_mod")) Then
'            frmExibicao9.lblDesvio_Aviso_Mod.Caption = cRec.Fields("desvio_aviso_mod")
        sSeqId = Mid$(cRec.Fields("desvio_aviso_mod"), 1, 6) ' fornecedor
        sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 13, 2) ' (aa)mmdd
        sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 9, 2)  ' aa(mm)dd
        sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 7, 2) ' aamm(dd)
        sSeqId = Trim(sSeqId) & Format(Mid$(cRec.Fields("desvio_aviso_mod"), 15, 11), "000000") ' sequencial
        
        frmExibicao9.lblCodBar_Desvio_Aviso_Mod.Caption = "*" & sSeqId & "*"
        frmExibicao9.lblCodBar_Desvio_Aviso_Mod1.Caption = "*" & sSeqId & "*"
        frmExibicao9.lblDesvio_Aviso_Mod.Caption = sSeqId
        
    End If
     
End If

'-----------------------------------------------------------------------------------------------------
'Se tipo 7 opcao Palete GM
If cRec.Fields("Tipo") = "7" Then
    If (IsNull(cRec.Fields("CODIGO_PRODUTO2").Value)) Then
        Set frmExibicao7Ref = frmExibicao7UmProduto
    ElseIf (Trim(cRec.Fields("CODIGO_PRODUTO2").Value) = "") Then
        Set frmExibicao7Ref = frmExibicao7UmProduto
    Else
        Set frmExibicao7Ref = frmExibicao7VariosProdutos
    End If
    
    frmExibicao7Ref.Show
    
    'Cliente
    If Not IsNull(cRec.Fields("Cliente").Value) Then
        frmExibicao7Ref.lblTo.Caption = cRec.Fields("Cliente").Value
    Else
        frmExibicao7Ref.lblTo.Caption = ""
    End If
    If Not IsNull(cRec.Fields("IND_SUPLEMentar").Value) Then
        frmExibicao7Ref.lblPlant.Caption = cRec.Fields("IND_SUPLEMentar").Value
    Else
        frmExibicao7Ref.lblPlant.Caption = ""
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
        frmExibicao7Ref.lblPlant.Caption = frmExibicao7Ref.lblPlant.Caption & "-" & cRec.Fields("Cod_Embalagem_Pw")
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
    If cRec.Fields("Embarque_Controlado") <> "" Then
        frmExibicao7Ref.lblPlant.Caption = frmExibicao7Ref.lblPlant.Caption & "-" & cRec.Fields("Embarque_Controlado")
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
    If cRec.Fields("Modelo") <> "" Then
        frmExibicao7Ref.lblMaterial.Caption = cRec.Fields("Modelo")
    End If
    
    Rem INCLUIDO EM 28-08-2007
    If cRec.Fields("envio_lote") = "1" Then
        frmExibicao7Ref.lblvalidade.Caption = "N"
    Else
        frmExibicao7Ref.lblvalidade.Caption = ""
    End If
    
    'Pode ser gravado nulo no caso da etiqueta N: 7 (Palete GM)
    If Not IsNull(cRec.Fields("Pto_Entrega").Value) Then
        frmExibicao7Ref.lblLicense.Caption = cRec.Fields("Pto_Entrega").Value
    Else
        frmExibicao7Ref.lblLicense.Caption = ""
    End If
    frmExibicao7Ref.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega").Value & "*"
    frmExibicao7Ref.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega").Value & "*"
    frmExibicao7Ref.lblShipmentDate.Caption = UCase(Format(Date, "DD/MMM/YYYY"))
    frmExibicao7Ref.lblPeso.Caption = Format(cRec.Fields("Peso"), "0.00")
    
'    frmExibicao7Ref.lblCodigoProduto1.Caption = cRec.Fields("CODIGO_PRODUTO1").Value
    Rem MUDADO EM 06-06-2016 PARA O CODIGO DO PRODUTO DO CLIENTE
    frmExibicao7Ref.lblCodigoProduto1.Caption = cRec.Fields("COD_CLIENTE").Value
    If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA1").Value)) Then
        frmExibicao7Ref.lblQtde1.Caption = cRec.Fields("QTDE_CAIXA1").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA1").Value)) & " PC"
    Else
        frmExibicao7Ref.lblQtde1.Caption = ""
    End If
    Rem MUDADO EM 06-06-2016 PARA O CODIGO DO PRODUTO DA MUSASHI
    If Not IsNull(cRec.Fields("CODIGO_PRODUTO1").Value) Then
        frmExibicao7Ref.lblComplPeca1.Caption = cRec.Fields("CODIGO_PRODUTO1").Value
    Else
        frmExibicao7Ref.lblComplPeca1.Caption = ""
    End If
    
    
    If frmExibicao7Ref.Name = "frmExibicao7UmProduto" Then
        If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA1").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA1").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA1").Value) <> 0 Then
                frmExibicao7UmProduto.lblQtdeTot1.Caption = cRec.Fields("QTDE_CAIXA1").Value * cRec.Fields("PECAS_CAIXA1").Value
            End If
        End If
    ElseIf frmExibicao7Ref.Name = "frmExibicao7VariosProdutos" Then
        frmExibicao7Ref.lblCodigoProduto2.Caption = cRec.Fields("CODIGO_PRODUTO2").Value
        If (IsNumeric(cRec.Fields("QTDE_CAIXA2").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA2").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA2").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA2").Value) <> 0 Then
                frmExibicao7Ref.lblQtde2.Caption = cRec.Fields("QTDE_CAIXA2").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA2").Value)) & " PC"
            End If
        Else
            frmExibicao7Ref.lblQtde2.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA2").Value) Then
            frmExibicao7Ref.lblComplPeca2.Caption = cRec.Fields("COMPL_PECA2").Value
        Else
            frmExibicao7Ref.lblComplPeca2.Caption = ""
        End If
        
        If Not IsNull(cRec.Fields("CODIGO_PRODUTO3").Value) Then
            frmExibicao7Ref.lblCodigoProduto3.Caption = cRec.Fields("CODIGO_PRODUTO3").Value
        Else
            frmExibicao7Ref.lblCodigoProduto3.Caption = ""
        End If
        If (IsNumeric(cRec.Fields("QTDE_CAIXA3").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA3").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA3").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA3").Value) <> 0 Then
                frmExibicao7Ref.lblQtde3.Caption = cRec.Fields("QTDE_CAIXA3").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA3").Value)) & " PC"
            End If
        Else
            frmExibicao7Ref.lblQtde3.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA3").Value) Then
            frmExibicao7Ref.lblComplPeca3.Caption = cRec.Fields("COMPL_PECA3").Value
        Else
            frmExibicao7Ref.lblComplPeca3.Caption = ""
        End If
        
        If Not IsNull(cRec.Fields("CODIGO_PRODUTO4").Value) Then
            frmExibicao7Ref.lblCodigoProduto4.Caption = cRec.Fields("CODIGO_PRODUTO4").Value
        Else
            frmExibicao7Ref.lblCodigoProduto4.Caption = ""
        End If
        If (IsNumeric(cRec.Fields("QTDE_CAIXA4").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA4").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA4").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA4").Value) <> 0 Then
                frmExibicao7Ref.lblQtde4.Caption = cRec.Fields("QTDE_CAIXA4").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA4").Value)) & " PC"
            End If
        Else
            frmExibicao7Ref.lblQtde4.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA4").Value) Then
            frmExibicao7Ref.lblComplPeca4.Caption = cRec.Fields("COMPL_PECA4").Value
        Else
            frmExibicao7Ref.lblComplPeca4.Caption = ""
        End If
    End If
    
    qtdeConteiners = 0
    If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA1").Value)
    End If
    If (IsNumeric(cRec.Fields("QTDE_CAIXA2").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA2").Value)
    End If
    If (IsNumeric(cRec.Fields("QTDE_CAIXA3").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA3").Value)
    End If
    If (IsNumeric(cRec.Fields("QTDE_CAIXA4").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA4").Value)
    End If
    
    
    frmExibicao7Ref.lblQtdeContainers.Caption = CStr(qtdeConteiners)
    
End If

'-------------------------------------------------------------------
'Se tipo A - Identificação de alterações do produto e processo
If cRec.Fields("Tipo") = "A" Then
    
    frmExibicao6.Show
    
    With frmExibicao6
        .lblDesenho.Caption = Trim(cRec("COD_CLIENTE").Value & "")
        .lblDesvio.Caption = Trim(cRec("DESVIO_AVISO_MOD").Value & "")
        .lblData1.Caption = Format(Date, "dd/mm/yyyy")
        .lblData2.Caption = .lblData1.Caption
        .lblNotaFiscal1.Caption = Trim(cRec("NUM_DOC_FISCAL").Value & "")
        .lblNotaFiscal2.Caption = .lblNotaFiscal1.Caption
        
        .lblOptDefinitiva.Visible = False
        .lblOptProvisoria.Visible = False
        .lblOptLoteUnico.Visible = False
        
        Select Case Trim(cRec("TIPO_ALTERACAO").Value)
        Case "1"
            .lblOptDefinitiva.Visible = True
        Case "2"
            .lblOptProvisoria.Visible = True
        Case "3"
            .lblOptLoteUnico.Visible = True
        End Select
        
        .lblOptDesvio.Visible = False
        .lblOptMaterialSelecionado.Visible = False
        .lblOptOutros.Visible = False
        .lblOptReparoRetrabaho.Visible = False
        .lblOptProdutoNovo.Visible = False
        Select Case Trim(cRec("MOTIVO_ALTERACAO").Value)
        Case "1"
            frmExibicao6.lblOptDesvio.Visible = True
        Case "2"
            frmExibicao6.lblOptMaterialSelecionado.Visible = True
        Case "3"
            frmExibicao6.lblOptOutros.Visible = True
        Case "4"
            frmExibicao6.lblOptReparoRetrabaho.Visible = True
        Case "5"
            frmExibicao6.lblOptProdutoNovo.Visible = True
        End Select
        
        .lblMotivoAlteracaoOutros.Caption = Trim(cRec("MOTIVO_ALTERACAO_OUTROS").Value & "")
        
        .lblOptPrimeiroEnvio.Visible = False
        .lblOptLoteIntermediario.Visible = False
        .lblOptUltimoLote.Visible = False
        Select Case Trim(cRec("ENVIO_LOTE").Value)
        Case "1"
            .lblOptPrimeiroEnvio.Visible = True
        Case "2"
            .lblOptLoteIntermediario.Visible = True
        Case "3"
            .lblOptUltimoLote.Visible = True
        End Select
        
        .lblNumAm.Caption = Trim(cRec("NUM_AM").Value & "")
        
    End With
    
    nQuantidade = cRec("QTD_ETIQ").Value
    
End If


'-------------------------------------------------------------------
'Se tipo y - Identificação de alterações do produto e processo

If cRec("Tipo") = "Y" Then
    If objApplication.filial = adMusashiDaAmazonia Then
       frmExibicao11.Show
       frmExibicao11.Left = Me.Width
       With frmExibicao11
'        .lblDesenho.Caption = Trim(cRec("COD_CLIENTE").Value & "")
           
           .lbl_LPN_COD.Caption = cRec("Cod_Fornecedor") & Format(Now(), "DDMMYY") & Format(cRec("Sequencia_Dia"), "000000")
            Call CodeRefresh(.lbl_LPN_COD.Caption)
           .lbl_LPN_COD_BARRA.Caption = sCodigo128
           .lbl_LPN_COD_B.Caption = .lbl_LPN_COD.Caption
           .lbl_CODIGO_NUM.Caption = Trim(cRec("Cod_Cliente"))
           .lbl_SUPPLIER_COD.Caption = IIf(IsNull(cRec("Cod_Fornecedor")), " ", cRec("Cod_Fornecedor"))
           .lbl_USER_COD.Caption = "9219"
           .lbl_YAMAHA_COD_BARRAS.Caption = cRec("Cod_Cliente") & "-" & _
                                            .lbl_SUPPLIER_COD.Caption & "-" & _
                                            "9219"
            Call CodeRefresh(.lbl_YAMAHA_COD_BARRAS.Caption)
           .lbl_YAMAHA_BARRAS.Caption = sCodigo128
           
           .lbl_NOME_DESCRICAO.Caption = cRec("descr_peca")
           .lbl_FORNECEDOR_NOME.Caption = "MUSASHI DO BRASIL LTDA"
           .lbl_QTDE_NUM.Caption = cRec("Qtd_Caixa")
           .lbl_NF_NUM.Caption = ""
           .LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec("data_etiq"), "dd/mm/yyyy"), 4, 2))
           .lbl_ANO.Caption = Format(cRec("data_etiq"), "YY")
           .lbl_QTDE_BARRAS.Caption = cRec("Qtd_Caixa")
           .lbl_QTDE_NUM1.Caption = cRec("Qtd_Caixa")
           .lbl_COD_MUSASHI_BARRAS.Caption = "*" & Trim(cRec("ID_ETIQUETA")) & "*"
           .lbl_COD_MUSASHI_NUM.Caption = Trim(cRec("ID_ETIQUETA"))
       End With
    Else
       frmExibicao10.Show
       With frmExibicao10
           .lbl_LPN_COD.Caption = cRec("Cod_Fornecedor") & Format(Now(), "DDMMYY") & Format(cRec("Sequencia_Dia"), "000000")
            Call CodeRefresh(.lbl_LPN_COD.Caption)
           .lbl_LPN_COD_BARRA.Caption = sCodigo128
           .lbl_LPN_COD_B.Caption = .lbl_LPN_COD.Caption
           .lbl_CODIGO_NUM.Caption = Trim(cRec("Cod_Cliente"))
           .lbl_SUPPLIER_COD.Caption = IIf(IsNull(cRec("Cod_Fornecedor")), " ", cRec("Cod_Fornecedor"))
           .lbl_USER_COD.Caption = "9219"
           .lbl_YAMAHA_COD_BARRAS.Caption = cRec("Cod_Cliente") & "-" & _
                                            .lbl_SUPPLIER_COD.Caption & "-" & _
                                            "9219"
            Call CodeRefresh(.lbl_YAMAHA_COD_BARRAS.Caption)
           .lbl_YAMAHA_BARRAS.Caption = sCodigo128
           
           .lbl_NOME_DESCRICAO.Caption = cRec("descr_peca")
'            If objApplication.filial = adMusashiDaAmazonia Then
              .lbl_FORNECEDOR_NOME.Caption = "MUSASHI DO BRASIL LTDA"
'            Else
'               .lbl_FORNECEDOR_NOME.Caption = ""
'            End If
           .lbl_QTDE_NUM.Caption = cRec("Qtd_Caixa")
           .lbl_NF_NUM.Caption = ""
           .LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec("data_etiq"), "dd/mm/yyyy"), 4, 2))
           .lbl_ANO.Caption = Format(cRec("data_etiq"), "YY")
           .lbl_QTDE_BARRAS.Caption = cRec("Qtd_Caixa")
           .lbl_QTDE_NUM1.Caption = cRec("Qtd_Caixa")
           .lbl_COD_MUSASHI_BARRAS.Caption = "*" & Trim(cRec("ID_ETIQUETA")) & "*"
           .lbl_COD_MUSASHI_NUM.Caption = Trim(cRec("ID_ETIQUETA"))
           .lblCodMSB.Caption = cRec.Fields("Cod_Peca")
           .lblCodMSB_Letra = Mid$(cRec.Fields("Cod_Peca"), Len(cRec.Fields("Cod_Peca")), 1)
           .lblLot.Caption = cRec.Fields("Lote")
       End With
    End If
    
End If

bTelaImp = False
Me.cmd_Impressao.Enabled = True
Me.SetFocus
    
Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
Me.cmd_Visualizar.SetFocus

    '-------------------------------------------------------------------
End Sub

Private Function Fechar_Form_Etiqueta()

Dim v As Integer

For v = 0 To (Forms.Count - 1)
    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
        Unload frmAvulsoPadraoPonteiro
        Exit For
    End If
    If Forms(v).Name = "frmExibicao2" Then
        Unload frmExibicao2
        Exit For
    End If
    If Forms(v).Name = "frmExibicao3" Then
        Unload frmExibicao3
        Exit For
    End If
    If Forms(v).Name = "frmExibicao5" Then
        Unload frmExibicao5
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7UmProduto" Then
        Unload frmExibicao7UmProduto
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7VariosProdutos" Then
        Unload frmExibicao7VariosProdutos
        Exit For
    End If
    If Forms(v).Name = "frmExibicao9" Then
        Unload frmExibicao9
        Exit For
    End If
    If Forms(v).Name = "frmExibicao10" Then
        Unload frmExibicao10
        Exit For
    End If
    If Forms(v).Name = "frmExibicao11" Then
        Unload frmExibicao11
        Exit For
    End If
    If Forms(v).Name = "frmExibicao4MDA" Then
        Unload frmExibicao4MDA
        Exit For
    End If

Next
        
Unload frmOpcoes

End Function

Private Sub Imprime_Etiqueta_Fiat_FTP()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim cRec As ADODB.Recordset
Dim sData As String
Dim oTela As frmEscRelCristalReport

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_FIAT_FTP(sBancoMusashi, _
                                                        Me.txtsequencial.Text)

If cRec.RecordCount = 0 Then
   MsgBox "Etiqueta não encontrada, anote a etiqueta e procure o responsável, etiq: " & dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")
   Me.MousePointer = vbDefault
   Exit Sub
End If

nx = 0

On Error GoTo Erro
Set rs = New ADODB.Recordset

rs.Fields.Append "1_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_Qtde_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Fab_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Aplicacao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Num_Doc_Fis_BAM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Data", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Pondo_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Denominacao", ADODB.DataTypeEnum.adChar, 80
rs.Fields.Append "18_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Ctrl_Interno", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "20_Ctrl_Oper_Log", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "21_Codigo_Numero", ADODB.DataTypeEnum.adChar, 50
rs.Fields.Append "22_codigo_barras", ADODB.DataTypeEnum.adChar, 50

rs.Open
nx = 0

cRec.MoveFirst
rs.AddNew
rs.Fields("1_Num_Lote").Value = IIf(IsNull(cRec.Fields("Num_Lote")), " ", Trim(cRec.Fields("Num_Lote")))
rs.Fields("2_Qtde_Emb").Value = IIf(IsNull(cRec.Fields("Qtde_Emb")), "0", Format(cRec.Fields("Qtde_Emb"), "00"))
rs.Fields("3_Classe_Func").Value = IIf(IsNull(cRec.Fields("Classe_Func")), " ", cRec.Fields("Classe_Func"))
rs.Fields("4_Indicacao_Supl").Value = IIf(IsNull(cRec.Fields("Indicacao_Supl")), " ", cRec.Fields("Indicacao_Supl"))
If Format(cRec.Fields("DUM"), "DD/MM/YYYY") = "01/01/1900" Then
   rs.Fields("5_Data_Fab_Lote").Value = " "
Else
   rs.Fields("5_Data_Fab_Lote").Value = IIf(IsNull(cRec.Fields("Data_Fab_Lote")), " ", Format(cRec.Fields("Data_Fab_Lote"), "DD/MM/YYYY"))
End If
rs.Fields("6_Cod_Emb").Value = IIf(IsNull(cRec.Fields("Cod_Emb")), " ", cRec.Fields("Cod_Emb"))
rs.Fields("7_Vinculo").Value = IIf(IsNull(cRec.Fields("Vinculo")), " ", cRec.Fields("Vinculo"))
rs.Fields("8_Lote_Sob_Desv").Value = IIf(IsNull(cRec.Fields("embalagem")), " ", cRec.Fields("embalagem"))
'rs.Fields("8_Lote_Sob_Desv").Value = IIf(IsNull(cRec.Fields("Lote_Sob_Desv")), " ", cRec.Fields("Lote_Sob_Desv"))
rs.Fields("9_Qtde_Lote").Value = IIf(IsNull(cRec.Fields("Qtde_Lote")), "0", cRec.Fields("Qtde_Lote"))
rs.Fields("10_Aplicacao").Value = IIf(IsNull(cRec.Fields("Aplicacao")), " ", Mid$(cRec.Fields("Aplicacao"), 1, 14))
If Format(cRec.Fields("DUM"), "DD/MM/YYYY") = "01/01/1900" Then
   rs.Fields("11_DUM").Value = " "
Else
   rs.Fields("11_DUM").Value = IIf(IsNull(cRec.Fields("DUM")), " ", Format(cRec.Fields("DUM"), "DD/MM/YYYY"))
End If
rs.Fields("12_Embarque_Ctrl").Value = IIf(IsNull(cRec.Fields("Embarque_Ctrl")), " ", cRec.Fields("Embarque_Ctrl"))
rs.Fields("13_Cod_Fornecedor").Value = "13093"
rs.Fields("14_Num_Doc_Fis_BAM").Value = IIf(IsNull(cRec.Fields("Num_Doc_Fis_BAM")), " ", cRec.Fields("Num_Doc_Fis_BAM"))
rs.Fields("15_Data").Value = IIf(IsNull(cRec.Fields("Data")), " ", cRec.Fields("Data"))
rs.Fields("16_Pondo_Entrega").Value = IIf(IsNull(cRec.Fields("Ponto_Entrega")), " ", Trim(cRec.Fields("Ponto_Entrega")))
rs.Fields("17_Denominacao").Value = IIf(IsNull(cRec.Fields("Denominacao")), " ", cRec.Fields("Denominacao"))
rs.Fields("18_Num_Desenho").Value = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", Mid$(cRec.Fields("Num_Desenho"), 1, 20))
rs.Fields("19_Ctrl_Interno").Value = IIf(IsNull(cRec.Fields("Ctrl_Interno")), " ", cRec.Fields("Ctrl_Interno"))
rs.Fields("20_Ctrl_Oper_Log").Value = IIf(IsNull(cRec.Fields("Ctrl_Oper_Log")), " ", cRec.Fields("Ctrl_Oper_Log"))
rs.Fields("21_Codigo_Numero").Value = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", cRec.Fields("Num_Desenho")) & Format(rs.Fields("2_Qtde_Emb").Value, "00000") & cRec.Fields("Cod_Emb") & "013093"
rs.Fields("22_codigo_barras").Value = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", "*" & Mid$(cRec.Fields("Num_Desenho"), 1, 11) & Format(rs.Fields("2_Qtde_Emb").Value, "00000") & cRec.Fields("Cod_Emb") & "013093" & "*")
rs.Update

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT.rpt")
CrystalReport1.Database.SetDataSource rs

rs.Clone

Rem mostrar a tela
'Set oTela = New frmEscRelCristalReport
'oTela.CRViewer1.ReportSource = CrystalReport1
'oTela.CRViewer1.ViewReport
'oTela.Show 0
Rem *************************

Rem nao mostrar a tela
CrystalReport1.PrintOutEx False
Rem *************************

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub
