VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReimprimeEtiqueta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-impressão das etiquetas"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9450
   Begin VB.CommandButton cmd_visualizar 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7350
      Picture         =   "frmReimprimeEtiqueta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Imprimir"
      Top             =   4620
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3315
      Left            =   60
      TabIndex        =   12
      Top             =   1200
      Width           =   9285
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   2955
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   5212
         _Version        =   393216
         Cols            =   9
         ForeColorFixed  =   16711680
         BackColorSel    =   16711935
         AllowBigSelection=   0   'False
         TextStyle       =   2
         TextStyleFixed  =   1
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmReimprimeEtiqueta.frx":0442
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
      Picture         =   "frmReimprimeEtiqueta.frx":044B
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Fecha a solicitação"
      Top             =   4620
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   6330
      Picture         =   "frmReimprimeEtiqueta.frx":088D
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir"
      Top             =   4620
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtro da etiqueta"
      Height          =   1065
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   9285
      Begin VB.TextBox txt_Qtd_Caixa 
         Height          =   315
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   3
         Top             =   630
         Width           =   675
      End
      Begin VB.CommandButton cmd_librera_Data 
         BackColor       =   &H0000FF00&
         Caption         =   "X"
         Height          =   315
         Left            =   6990
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Fecha a solicitação"
         Top             =   450
         Width           =   285
      End
      Begin VB.TextBox txt_peca 
         Height          =   315
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   2
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txt_Lote 
         Height          =   315
         Left            =   1050
         MaxLength       =   15
         TabIndex        =   1
         Top             =   630
         Width           =   1335
      End
      Begin VB.TextBox txtsequencial 
         Height          =   315
         Left            =   1050
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "0000000000"
         Top             =   270
         Width           =   1335
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   495
         Left            =   8610
         Picture         =   "frmReimprimeEtiqueta.frx":0B97
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Fecha a solicitação"
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   495
         Left            =   8010
         Picture         =   "frmReimprimeEtiqueta.frx":0FD9
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Grava Solicitação"
         Top             =   300
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
         TabIndex        =   4
         Top             =   450
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16121857
         CurrentDate     =   37837
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
         Height          =   195
         Left            =   2820
         TabIndex        =   17
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peça :"
         Height          =   195
         Left            =   2820
         TabIndex        =   15
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lote :"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   660
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial :"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lbldatanasc 
         AutoSize        =   -1  'True
         Caption         =   "Data.:"
         Height          =   195
         Left            =   5220
         TabIndex        =   11
         Top             =   480
         Width           =   435
      End
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
      TabIndex        =   24
      Top             =   5190
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
      TabIndex        =   23
      Top             =   5160
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
      TabIndex        =   22
      Top             =   4935
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
      TabIndex        =   21
      Top             =   4920
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
      TabIndex        =   20
      Top             =   4680
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
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmReimprimeEtiqueta"
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
   Me.Grid1.Col = 0
   Me.Grid1.Row = 1
   Grid1.Col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
   Grid1.Col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
   Grid1.Col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
   Grid1.Col = 0
   Me.Grid1.SetFocus
Else
   Call carrega_Grid
   Me.cmd_Impressao.Enabled = True
   Me.cmd_visualizar.Enabled = True
   cmd_visualizar_Click
   Me.Grid1.Col = 0
   Me.cmd_Impressao.SetFocus
End If

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
   Me.txtsequencial.SetFocus
End If
End Sub

Private Sub cmd_Impressao_Click()

Dim v As Integer

For v = 0 To (Forms.Count - 1)

    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
          Printer.Orientation = 2
          frmAvulsoPadraoPonteiro.PrintForm
          Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao2" Then
        Printer.Orientation = 1
        frmExibicao2.PrintForm
        Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao3" Then
        Printer.Orientation = 2
        frmExibicao3.PrintForm
        Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao5" Then
        Printer.Orientation = 2
        frmExibicao5.PrintForm
        Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao6" Then
        Printer.Orientation = 2
        frmExibicao6.PrintForm
        Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7UmProduto" Or Forms(v).Name = "frmExibicao7VariosProdutos" Then
        Printer.Orientation = 2
        frmExibicao7Ref.PrintForm
        Printer.EndDoc
        Exit For
    End If
Rem incluido para impressao do novo formulario da mvm international motores
    If Forms(v).Name = "frmExibicao9" Then
        Printer.Orientation = 2
        frmExibicao9.PrintForm
        Printer.EndDoc
        Exit For
    End If
Next
        
MsgBox "Impressão concluída com sucesso! a tela da etiqueta será encerrada!", vbOKOnly + vbInformation, "Tarefa Concluída"

Call Fechar_Form_Etiqueta
cmd_limpar_Click

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
Me.dtpDataSelecao.Enabled = False
Me.cmd_Impressao.Enabled = False
Me.cmd_visualizar.Enabled = False
Me.cmd_librera_Data.BackColor = &HFF&
Me.txtsequencial.SetFocus
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
Me.txtsequencial.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > 0 Then
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
    Grid1.Col = 0: Grid1.Text = cRec.Fields("sequencia")
    Grid1.Col = 1: Grid1.Text = cRec.Fields("Cod_Peca")
    Grid1.Col = 2: Grid1.Text = cRec.Fields("Descr_Peca")
    Grid1.Col = 3: Grid1.Text = cRec.Fields("cod_Cliente")
    Grid1.Col = 4: Grid1.Text = cRec.Fields("Lote")
    Grid1.Col = 5: Grid1.Text = cRec.Fields("Qtd_Caixa")
    Grid1.Col = 6: Grid1.Text = cRec.Fields("Data_Etiq")
    
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
Grid1.Col = 0: Grid1.ColWidth(0) = 1000:  Grid1.Text = "SEQ."
Grid1.Col = 1:  Grid1.ColWidth(1) = 900: Grid1.Text = "PECA"
Grid1.Col = 2: Grid1.ColWidth(2) = 2900: Grid1.Text = "DESCRICAO"
Grid1.Col = 3: Grid1.ColWidth(3) = 1500: Grid1.Text = "COD.CLI"
Grid1.Col = 4: Grid1.ColWidth(4) = 900: Grid1.Text = "LOTE"
Grid1.Col = 5: Grid1.ColWidth(5) = 500: Grid1.Text = "QTD"
Grid1.Col = 6: Grid1.ColWidth(6) = 1000: Grid1.Text = "DATA"
Grid1.Col = 7: Grid1.ColWidth(7) = 1: Grid1.Text = "VALOR"
Grid1.Col = 8: Grid1.ColWidth(8) = 1: Grid1.Text = "VALOR"
Grid1.Col = 2: Grid1.BackColor = &H80FFFF

Grid1.Row = 0

Grid1.HighLight = False
Grid1.ColAlignment(0) = flexAlignLeftCenter
Grid1.ColAlignment(1) = flexAlignLeftCenter
Grid1.ColAlignment(2) = flexAlignLeftCenter
Grid1.ColAlignment(3) = flexAlignLeftCenter
Grid1.ColAlignment(4) = flexAlignLeftCenter
Grid1.ColAlignment(5) = flexAlignRightCenter
Grid1.ColAlignment(6) = flexAlignLeftCenter

Me.lbl_produto.Visible = True
Me.label33.Visible = True
Me.Label5.Visible = True
Me.lbl_qtd.Visible = True
Me.lbl_sequencia.Visible = True
Me.lbl_peca.Visible = True

Me.lbl_qtd.Caption = ""
Me.lbl_sequencia.Caption = ""
Me.lbl_peca.Caption = ""

End Sub

Private Sub Grid1_Click()
If Grid1.Rows > 2 Then
   Grid1.Col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
   Grid1.Col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
   Grid1.Col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
   Grid1.Col = 0
End If
End Sub

Private Sub Grid1_GotFocus()

If Grid1.Rows > 2 Then
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
'   grid.Col = 3: Me.lbl_qtd.Caption = Me.Grid1.Text
'   Me.lbl_sequencia.Caption = cRec.Fields("sequencia")
'   Me.lbl_peca.Caption = cRec.Fields("Cod_Peca")
End If

End Sub


Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmd_visualizar_Click
End If
End Sub


Private Sub Grid1_SelChange()
   If Grid1.Rows > 2 Then
      Grid1.Col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
      Grid1.Col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
      Grid1.Col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
      Grid1.Col = 0
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

    
Rem leitura no banco para emissão da etiqueta
Me.Grid1.Col = 0
sSequencial = Format(Grid1.Text, "0000000000")

Me.MousePointer = vbHourglass

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar(sBancoMusashi, _
                                               sSequencial)

If cRec.RecordCount = 0 Then
   MsgBox "Registro não encontrado! A etiqueta não será impresa!", vbInformation + vbOKOnly, "Tarefa com problemas"
   Exit Sub
End If

Me.MousePointer = vbDefault
    
executaUnloadForm = False

'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta


qtdeConteiners = cRec.Fields("Tipo").Value
nQuantidade = cRec.Fields("Qtd_Etiq")
    
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
    End If
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 2 opcao FIAT
    If cRec.Fields("Tipo") = "2" Then
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
         
        
    End If
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 3 opcao FORD
    If cRec.Fields("Tipo") = "3" Then
        
        frmExibicao3.Show
        frmExibicao3.lblNumPeca.Caption = Left(cRec.Fields("Cod_Cliente"), 15)
        frmExibicao3.lblNumPecaA.Caption = "*P" & Trim(Left(cRec.Fields("Cod_Cliente"), 15)) & "*"
        frmExibicao3.lblNumPecaB.Caption = frmExibicao3.lblNumPecaA.Caption
        frmExibicao3.lblNumPeca.Caption = Trim(frmExibicao3.lblNumPeca.Caption)
        frmExibicao3.lblCod_Peca = cRec.Fields("Cod_Peca")
        frmExibicao3.lblLote = cRec.Fields("Lote")
        frmExibicao3.lblQtd.Caption = cRec.Fields("Qtd_Caixa")
        frmExibicao3.lblQtdA.Caption = "*Q" & cRec.Fields("Qtd_Caixa") & "*"
        frmExibicao3.lblQtdB.Caption = frmExibicao3.lblQtdA.Caption
        If Not IsNull(cRec.Fields("Cod_Fornecedor")) Then
            frmExibicao3.lblNumFornec.Caption = cRec.Fields("Cod_Fornecedor")
            frmExibicao3.lblNumFornecA.Caption = "*V" & cRec.Fields("Cod_Fornecedor") & "*"
            frmExibicao3.lblNumFornecB.Caption = frmExibicao3.lblNumFornecA.Caption
        End If
        If Not IsNull(cRec.Fields("Serial")) Then
            frmExibicao3.lblNumSerial.Caption = cRec.Fields("Serial")
            frmExibicao3.lblNumSerialA.Caption = "*S" & cRec.Fields("Serial") & "*"
            frmExibicao3.lblNumSerialB.Caption = frmExibicao3.lblNumSerialA.Caption
        End If
        If Not IsNull(cRec.Fields("Cod_Util")) Then
            frmExibicao3.lblCodUtil.Caption = cRec.Fields("Cod_Util")
        End If
        If Not IsNull(cRec.Fields("Linha_Util")) Then
            frmExibicao3.lblLinhaUtil.Caption = cRec.Fields("Linha_Util")
        End If
        frmExibicao3.lblSufixo.Caption = Trim(Right((cRec.Fields("Cod_Cliente")), 5))
        frmExibicao3.lblSufixoA.Caption = "*C" & frmExibicao3.lblSufixo.Caption & "*"
        frmExibicao3.lblSufixoB.Caption = frmExibicao3.lblSufixoA.Caption
        If (cRec.Fields("Desvio")) <> "" Then
            frmExibicao3.lblDestino.Caption = Trim(cRec.Fields("Desvio"))
        End If
        frmExibicao3.lblDestinoA.Caption = "*D" & frmExibicao3.lblDestino.Caption & "*"
        frmExibicao3.lblDestinoB.Caption = frmExibicao3.lblDestinoA.Caption
        
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmExibicao3.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
        End If
        
        If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
            frmExibicao3.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
            frmExibicao3.lblCodigoBarrasA.Caption = "*" & frmExibicao3.lblCodigoBarras.Caption & "*"
            frmExibicao3.lblCodigoBarrasB.Caption = "*" & frmExibicao3.lblCodigoBarras.Caption & "*"
         Else
            frmExibicao3.lblCodigoBarras.Caption = ""
            frmExibicao3.lblCodigoBarrasA.Caption = ""
            frmExibicao3.lblCodigoBarrasB.Caption = ""
         End If
        
        
    End If
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 4 opcao padrão etiqueta grande
    If cRec.Fields("Tipo") = "4" Or cRec.Fields("Tipo") = "8" Then
        'frmExibicao4.Show
        frmAvulsoPadraoPonteiro.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
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
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
         End If
         If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
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
    'Se tipo 5 opcao GM etiqueta grande
    'MODIFICADO EM 23-08-2007 (MARCOS PEDROSA) ACRESCENTAR CAMPOS LAY-OUT NOVO
    If cRec.Fields("Tipo") = "5" Then
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
        If cRec.Fields("Pto_Entrega") <> "" Then
            frmExibicao5.lblLicense.Caption = cRec.Fields("Pto_Entrega")
            frmExibicao5.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
            frmExibicao5.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
        End If
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
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmExibicao5.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
        End If
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
    If cRec.Fields("Tipo") = "6" Then
        frmExibicao9.Show
        frmExibicao9.Left = Me.Width
'        If cRec.Fields("Cliente") <> "" Then
'            frmExibicao9.lblCliente.Caption = cRec.Fields("Cliente")
'        End If
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
        
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmExibicao9.lblCodFunc.Caption = cRec.Fields("EMBALAGEM")
        Else
            frmExibicao9.lblCodFunc.Caption = ""
        End If
        
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
        
        frmExibicao7Ref.lblCodigoProduto1.Caption = cRec.Fields("CODIGO_PRODUTO1").Value
        If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA1").Value)) Then
            frmExibicao7Ref.lblQtde1.Caption = cRec.Fields("QTDE_CAIXA1").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA1").Value)) & " PC"
        Else
            frmExibicao7Ref.lblQtde1.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA1").Value) Then
            frmExibicao7Ref.lblComplPeca1.Caption = cRec.Fields("COMPL_PECA1").Value
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
    
    bTelaImp = False
    Me.cmd_Impressao.Enabled = True
    Me.SetFocus
    
    '-------------------------------------------------------------------
End Sub

Public Function Pega_Mes(nMes As Integer) As String
If nMes = 1 Then Pega_Mes = "JAN"
If nMes = 2 Then Pega_Mes = "FEV"
If nMes = 3 Then Pega_Mes = "MAR"
If nMes = 4 Then Pega_Mes = "ABR"
If nMes = 5 Then Pega_Mes = "MAI"
If nMes = 6 Then Pega_Mes = "JUN"
If nMes = 7 Then Pega_Mes = "JUL"
If nMes = 8 Then Pega_Mes = "AGO"
If nMes = 9 Then Pega_Mes = "SET"
If nMes = 10 Then Pega_Mes = "OUT"
If nMes = 11 Then Pega_Mes = "NOV"
If nMes = 12 Then Pega_Mes = "DEZ"
End Function

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
Next

End Function

