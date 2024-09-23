VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEtiquetaHondaQrcode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas Hoda Qrcode"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   10260
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
      Left            =   1290
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   5490
      Width           =   3975
   End
   Begin VB.TextBox TXT_EMONTRADOS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   5
      ToolTipText     =   "Digie a sequencial da etiqueta"
      Top             =   5160
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Height          =   3705
      Left            =   90
      TabIndex        =   14
      Top             =   1200
      Width           =   9525
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3315
         Left            =   90
         TabIndex        =   4
         Top             =   150
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   18
         ForeColorFixed  =   16711680
         BackColorSel    =   65535
         ForeColorSel    =   65535
         HighLight       =   0
         GridLines       =   2
         GridLinesFixed  =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmEtiquetaHondaQrcode.frx":0000
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
      Left            =   6930
      Picture         =   "frmEtiquetaHondaQrcode.frx":0009
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Fecha a tela de etiqueta"
      Top             =   5190
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   5910
      Picture         =   "frmEtiquetaHondaQrcode.frx":044B
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir Todaqs as etiquetas"
      Top             =   5190
      Width           =   975
   End
   Begin VB.Frame frm_filtro 
      Caption         =   "Filtro da etiqueta"
      Height          =   1065
      Left            =   90
      TabIndex        =   10
      Top             =   90
      Width           =   7875
      Begin VB.TextBox txtsequencial 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         MaxLength       =   16
         TabIndex        =   1
         ToolTipText     =   "Digie a sequencial da etiqueta"
         Top             =   450
         Width           =   1335
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   495
         Left            =   7140
         Picture         =   "frmEtiquetaHondaQrcode.frx":0755
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Limpar tela para nova etiqueta"
         Top             =   240
         Width           =   555
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   495
         Left            =   6570
         Picture         =   "frmEtiquetaHondaQrcode.frx":0B97
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Confirma dados do filtro"
         Top             =   240
         Width           =   555
      End
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   90
         TabIndex        =   11
         Top             =   180
         Width           =   1425
         Begin VB.OptionButton Opt_NFE 
            Caption         =   "Pela NFE"
            Height          =   255
            Left            =   120
            TabIndex        =   0
            Top             =   300
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Opt_etiqueta 
            Caption         =   "Pelo Nº da Etiqueta"
            Height          =   255
            Left            =   1500
            TabIndex        =   12
            Top             =   270
            Visible         =   0   'False
            Width           =   1755
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digitação:"
         Height          =   195
         Left            =   1710
         TabIndex        =   13
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd_Visualizar 
      Caption         =   "Visualizar"
      Height          =   735
      Left            =   5250
      TabIndex        =   9
      Top             =   6540
      Width           =   975
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label3 
      Caption         =   "CASO QUEIRA IMPRIMIR UMA ETIQUETA, BASTA CLICAR DUAS VEZES NO GRID PARA SELECIONAR "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   150
      TabIndex        =   16
      Top             =   5910
      Width           =   5475
   End
   Begin VB.Label Label2 
      Caption         =   "Encontrados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   15
      Top             =   5190
      Width           =   1155
   End
End
Attribute VB_Name = "frmEtiquetaHondaQrcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variável para MDIapp
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public bAtivo As Boolean
Public bTelaImp As Boolean
Public bJafoi As Boolean
Public nLogin As Integer ' conterá o codigo do usuario, quando confirmar a senha
Public nTipo As Integer ' conterá o tipo do usuario, quando confirmar a senha
Public nMatricula As Double ' conterá a matricula do usuario, quando confirmar a senha
Private nSequencia_Escolhida As Integer
Private bOmiteImpressao As Boolean

Private Sub btoConfirma_Click()

Dim nx As Integer
Dim sData As String
Dim sNota As String
Dim sLote As String

On Error GoTo Erro
If Len(Trim(Me.txtsequencial.Text)) = 0 Then
   MsgBox "O Sequencial do Pallet Inicial tem que estar digitado!"
   Me.txtsequencial.SetFocus
   Exit Sub
End If

Me.MousePointer = vbHourglass
Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False

If Me.Opt_NFE.Value = True Then
   sData = "1"
Else
   sData = "2"
End If

Me.TXT_EMONTRADOS.Text = ""

Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_Filtro_HONDA_QRCODE(sBancoMusashi, _
                                                                   Me.txtsequencial.Text, _
                                                                   sData)

If cRec.RecordCount > 0 Then
   Call carrega_Grid
   Me.TXT_EMONTRADOS.Text = cRec.RecordCount

   Rem criticas a serem realizadas no arquivo caso seja por pallet
   If cRec.RecordCount = 0 Then
      MsgBox "Nota fiscal não encontrada. Digite uma nova numeração."
      Me.MousePointer = vbDefault
      Me.cmd_Impressao.Enabled = False
      Me.cmd_Visualizar.Enabled = False
      Exit Sub
    Else
      Me.cmd_Impressao.Enabled = True
      Me.cmd_Visualizar.Enabled = True
   End If
   
End If

Grid1.col = 0
Grid1.Row = 1


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
'Dim CrystalReport1 As New CRAXDRT.Report
'Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

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

On Error GoTo Erro

If Grid1.Rows > 0 Then
   bOmiteImpressao = True
   For nx = 1 To Grid1.Rows
       nSequencia_Escolhida = nx
       Call cmd_visualizar_Click
   Next
End If

Me.MousePointer = vbDefault
bOmiteImpressao = False

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_limpar_Click()
Call Fechar_Form_Etiqueta

Call Limpar_Grid
Me.TXT_EMONTRADOS.Text = ""

Me.txtsequencial.Text = ""
Me.txtsequencial.Enabled = True

Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False

End Sub

Private Sub cmdfechar_Click()
Unload frmExibicaoHondaNova
Unload Me
End Sub

Private Sub cmd_visualizar_Click()
Dim sDescricao As String * 21
Dim nx As Integer

Unload frmExibicaoHondaNova
Grid1.Row = nSequencia_Escolhida
Grid1.col = 8

If bOmiteImpressao Then
   frmExibicaoHondaNova.cmd_imprime.Visible = False
Else
   frmExibicaoHondaNova.cmd_imprime.Visible = True
End If

If Val(Grid1.Text) = 2 Then
    frmExibicaoHondaNova.lbl_Empresa.Caption = "HCA"
Else
    frmExibicaoHondaNova.lbl_Empresa.Caption = "HDA"
End If

frmExibicaoHondaNova.txt_mome_impressora.Text = UCase(Me.cbo_impressora.List(Me.cbo_impressora.ListIndex))

Grid1.col = 7: frmExibicaoHondaNova.lbl_Nota_Fiscal.Caption = Format(Grid1.Text, "########0") & " 2"
Grid1.col = 12: frmExibicaoHondaNova.lbl_Data_Entrega.Caption = Grid1.Text
Grid1.col = 0: frmExibicaoHondaNova.lbl_Sequencial.Caption = Format(Grid1.Text, "000") & "/"
Grid1.col = 13: frmExibicaoHondaNova.lbl_Sequencial.Caption = frmExibicaoHondaNova.lbl_Sequencial.Caption & Format(Grid1.Text, "000")
Grid1.col = 10: frmExibicaoHondaNova.lbl_Item.Caption = Replace(Mid$(Grid1.Text, 1, InStr(Grid1.Text, "-")), "-", "") & Replace(Mid$(Grid1.Text, InStr(1, Grid1.Text, "-") + 1), "-", " ")
Grid1.col = 2: frmExibicaoHondaNova.lbl_Descricao_Item.Caption = Grid1.Text
Grid1.col = 3: frmExibicaoHondaNova.lbl_Quantidade.Caption = Replace(Format(Grid1.Text, "0"), ",", "")
Grid1.col = 16
If Trim(Grid1.Text) <> "" Then
   Select Case Trim(Grid1.Text)
   Case "1"
       frmExibicaoHondaNova.lbl_Embalagem.Caption = "IK10"
   Case "2"
       frmExibicaoHondaNova.lbl_Embalagem.Caption = "IK33"
   Case "3"
       frmExibicaoHondaNova.lbl_Embalagem.Caption = "PAPELAO"
   Case "4"
       frmExibicaoHondaNova.lbl_Embalagem.Caption = "CPG (FIAT) "
   End Select
Else
    frmExibicaoHondaNova.lbl_Embalagem.Caption = "IK10"
End If
Grid1.col = 9: frmExibicaoHondaNova.lbl_Fornecedor.Caption = Grid1.Text
frmExibicaoHondaNova.Lbl_Destino.Caption = "CD6"
Grid1.col = 5: frmExibicaoHondaNova.lbl_Codigo_Musashi.Caption = "EQ" & Mid$(Grid1.Text, 5, 10)

Grid1.col = 1: frmExibicaoHondaNova.lbl_peca.Caption = Grid1.Text
Grid1.col = 4: frmExibicaoHondaNova.lbl_lote.Caption = Grid1.Text
Grid1.col = 5: frmExibicaoHondaNova.lbl_cod_barras.Caption = Trim(Grid1.Text)
Grid1.col = 5: frmExibicaoHondaNova.lbl_cod_barras1.Caption = "*" & Trim(Grid1.Text) & "*"
Grid1.col = 5: frmExibicaoHondaNova.lbl_cod_barras2.Caption = "*" & Trim(Grid1.Text) & "*"

frmExibicaoHondaNova.lbl_Seq_Milhar.Caption = Mid$(Trim(frmExibicaoHondaNova.lbl_cod_barras.Caption), _
                                              Len(Trim(frmExibicaoHondaNova.lbl_cod_barras.Caption)) - 5, _
                                              Len(Trim(frmExibicaoHondaNova.lbl_cod_barras.Caption)))

Grid1.col = 5: frmExibicaoHondaNova.Text2.Text = "EQ" & Mid$(Grid1.Text, 5, 10) 'EQ + SEQUENCIAL DA ETIQUETA

Grid1.col = 7: frmExibicaoHondaNova.Text1.Text = Format(Grid1.Text, "000000000") 'NFE+"2  "+DESCRICAO+QTDE+UNIDADE
frmExibicaoHondaNova.Text1.Text = frmExibicaoHondaNova.Text1.Text & "2  "
Grid1.col = 10: sDescricao = frmExibicaoHondaNova.lbl_Item.Caption: frmExibicaoHondaNova.Text1.Text = frmExibicaoHondaNova.Text1.Text & sDescricao
Grid1.col = 3: frmExibicaoHondaNova.Text1.Text = frmExibicaoHondaNova.Text1.Text & Replace(Format(frmExibicaoHondaNova.lbl_Quantidade.Caption, "000000000.000"), ",", "")

Grid1.col = 14: frmExibicaoHondaNova.lbl_embalador.Caption = Format(Grid1.Text, "#")
Grid1.col = 15: frmExibicaoHondaNova.lbl_data_etiq.Caption = Format(Grid1.Text, "DD/MM/YYYY")


frmExibicaoHondaNova.Text1.Text = frmExibicaoHondaNova.Text1.Text & "PC" & frmExibicaoHondaNova.lbl_Codigo_Musashi.Caption
frmExibicaoHondaNova.Calcular_Qrcode

frmExibicaoHondaNova.Show
frmExibicaoHondaNova.Left = Me.Width

If bOmiteImpressao Then
   Printer.Orientation = 2
   frmExibicaoHondaNova.PrintForm
   Printer.Orientation = 2: Printer.EndDoc
End If
End Sub

Private Sub Form_Activate()
Me.TXT_EMONTRADOS.Text = ""
If bAtivo Then Exit Sub

bAtivo = True
bJafoi = False
Call Limpar_Grid
Me.txtsequencial.Text = ""
Me.txtsequencial.SetFocus
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
Dim nqtde As Double

Call Limpar_Grid

Grid1.Row = 1
cRec.MoveFirst
bJafoi = True

For nx = 1 To cRec.RecordCount
    Grid1.col = 0: Grid1.Text = IIf(IsNull(cRec.Fields("seq")), " ", cRec.Fields("seq"))
    Grid1.col = 1: Grid1.Text = IIf(IsNull(cRec.Fields("id_peca")), " ", cRec.Fields("id_peca"))
    Grid1.col = 2: Grid1.Text = IIf(IsNull(cRec.Fields("Descr_Peca")), " ", Trim(cRec.Fields("Descr_Peca")))
    Grid1.col = 3: Grid1.Text = IIf(IsNull(cRec.Fields("qtde")), "0", Format(cRec.Fields("qtde"), "00"))
    Grid1.col = 4: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Lote")), " ", cRec.Fields("Num_Lote"))
    Grid1.col = 5: Grid1.Text = IIf(IsNull(cRec.Fields("id")), " ", cRec.Fields("id"))
    Grid1.col = 6: Grid1.Text = IIf(IsNull(cRec.Fields("tipo_embalagem")), " ", cRec.Fields("tipo_embalagem"))
    Grid1.col = 7: Grid1.Text = IIf(IsNull(cRec.Fields("num_doc_fiscal")), " ", cRec.Fields("num_doc_fiscal"))
    Grid1.col = 8: Grid1.Text = IIf(IsNull(cRec.Fields("ID_CLIENTE")), " ", cRec.Fields("id_cliente"))
    Grid1.col = 9: Grid1.Text = IIf(IsNull(cRec.Fields("fornecedor")), " ", cRec.Fields("fornecedor"))
    Grid1.col = 10: Grid1.Text = IIf(IsNull(cRec.Fields("cod_no_cliente")), " ", cRec.Fields("cod_no_cliente"))
    Grid1.col = 11: Grid1.Text = IIf(IsNull(cRec.Fields("cod_util")), " ", cRec.Fields("cod_util"))
    Grid1.col = 12: Grid1.Text = IIf(IsNull(cRec.Fields("data_sistema")), " ", cRec.Fields("data_sistema"))
    Grid1.col = 13: Grid1.Text = IIf(IsNull(cRec.Fields("total_volume")), " ", Format(cRec.Fields("total_volume"), "0"))
    Grid1.col = 14: Grid1.Text = IIf(IsNull(cRec.Fields("EMBALAGEM")), " ", cRec.Fields("EMBALAGEM"))
    Grid1.col = 15: Grid1.Text = IIf(IsNull(cRec.Fields("DATA_ETIQ")), " ", cRec.Fields("DATA_ETIQ"))
    Grid1.col = 16: Grid1.Text = IIf(IsNull(cRec.Fields("Tipo_caixa")), " ", cRec.Fields("Tipo_caixa"))
    Grid1.col = 17: Grid1.Text = IIf(IsNull(cRec.Fields("DATA_FATURAMENTO")), " ", cRec.Fields("DATA_FATURAMENTO"))
    Grid1.col = 0
  
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

Grid1.col = 0: Grid1.ColWidth(0) = 600: Grid1.Text = "SEQ"
Grid1.col = 1: Grid1.ColWidth(1) = 1100: Grid1.Text = "IDENT.PECA"
Grid1.col = 2: Grid1.ColWidth(2) = 3200: Grid1.Text = "DESCRIÇÃO"
Grid1.col = 3: Grid1.ColWidth(3) = 600: Grid1.Text = "QTDE"
Grid1.col = 4: Grid1.ColWidth(4) = 700: Grid1.Text = "LOTE"
Grid1.col = 5: Grid1.ColWidth(5) = 1200: Grid1.Text = "Nº.ETIQUETA"
Grid1.col = 6: Grid1.ColWidth(6) = 700: Grid1.Text = "EMB."
Grid1.col = 7: Grid1.ColWidth(7) = 1000: Grid1.Text = "Nº NFE"
Grid1.col = 8: Grid1.ColWidth(8) = 1000: Grid1.Text = "COD.CLIENTE"
Grid1.col = 9: Grid1.ColWidth(9) = 1000: Grid1.Text = "FORNECEDOR"
Grid1.col = 10: Grid1.ColWidth(10) = 1000: Grid1.Text = "CLIENTE"
Grid1.col = 11: Grid1.ColWidth(11) = 1000: Grid1.Text = "TIPO EMB."
Grid1.col = 12: Grid1.ColWidth(12) = 1000: Grid1.Text = "DATA"
Grid1.col = 13: Grid1.ColWidth(13) = 1000: Grid1.Text = "QT.VOLUME"
Grid1.col = 14: Grid1.ColWidth(14) = 1000: Grid1.Text = "COD.EMB"
Grid1.col = 15: Grid1.ColWidth(15) = 1000: Grid1.Text = "QT.VOLUME"
Grid1.col = 16: Grid1.ColWidth(16) = 1000: Grid1.Text = "TP.CAIXA"
Grid1.col = 17: Grid1.ColWidth(17) = 1000: Grid1.Text = "DT.FATUR"

Grid1.HighLight = False
Grid1.ColAlignment(0) = flexAlignLeftCenter
Grid1.ColAlignment(1) = flexAlignLeftCenter
Grid1.ColAlignment(2) = flexAlignLeftCenter
Grid1.ColAlignment(3) = flexAlignLeftCenter
Grid1.ColAlignment(4) = flexAlignLeftCenter
Grid1.ColAlignment(5) = flexAlignLeftCenter
Grid1.ColAlignment(7) = flexAlignLeftCenter
Grid1.ColAlignment(8) = flexAlignLeftCenter
Grid1.ColAlignment(9) = flexAlignLeftCenter
Grid1.ColAlignment(10) = flexAlignLeftCenter
Grid1.ColAlignment(11) = flexAlignLeftCenter
Grid1.ColAlignment(12) = flexAlignLeftCenter
Grid1.ColAlignment(13) = flexAlignLeftCenter
Grid1.ColAlignment(14) = flexAlignLeftCenter
Grid1.ColAlignment(15) = flexAlignLeftCenter
Grid1.ColAlignment(16) = flexAlignLeftCenter
Grid1.ColAlignment(17) = flexAlignLeftCenter

End Sub

Private Sub Grid1_DblClick()
nSequencia_Escolhida = Grid1.Row
If nSequencia_Escolhida > 0 Then
   bOmiteImpressao = False
   Call cmd_visualizar_Click
End If
End Sub

Private Sub txtsequencial_Change()
If Len(Trim(txtsequencial.Text)) = 11 Then
   Call btoConfirma_Click
End If

End Sub

Private Sub txtsequencial_GotFocus()
Me.btoConfirma.Default = True
End Sub

Private Sub txtsequencial_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   If MsgBox("Deseja sair deste módulo?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbNo Then
      Me.txtsequencial.Text = ""
      Me.txtsequencial.SetFocus
   Else
      Unload Me
   End If
End If

End Sub

Private Function Fechar_Form_Etiqueta()

Unload frmOpcoes

End Function

