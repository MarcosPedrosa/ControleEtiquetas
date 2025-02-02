VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmEtiquetaKRUPP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas  KRUPP"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9690
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
      Left            =   1320
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   6000
      Width           =   5025
   End
   Begin VB.CommandButton cmd_Impressao2 
      Height          =   735
      Left            =   7440
      Picture         =   "frmEtiquetaKRUPP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir"
      Top             =   6210
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao1 
      Height          =   735
      Left            =   8580
      Picture         =   "frmEtiquetaKRUPP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprimir"
      Top             =   6210
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3705
      Left            =   180
      TabIndex        =   11
      Top             =   1380
      Width           =   9285
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3315
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   26
         ForeColorFixed  =   16711680
         BackColorSel    =   65535
         ForeColorSel    =   65535
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLines       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmEtiquetaKRUPP.frx":0614
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
      Left            =   8520
      Picture         =   "frmEtiquetaKRUPP.frx":061D
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Fecha a tela de etiqueta"
      Top             =   5370
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7500
      Picture         =   "frmEtiquetaKRUPP.frx":0A5F
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir"
      Top             =   5370
      Width           =   975
   End
   Begin VB.Frame frm_filtro 
      Caption         =   "Filtro da etiqueta"
      Height          =   1065
      Left            =   180
      TabIndex        =   1
      Top             =   210
      Width           =   9285
      Begin VB.TextBox txtsequencial 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4320
         MaxLength       =   16
         TabIndex        =   7
         ToolTipText     =   "Digie a sequencial da etiqueta"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   495
         Left            =   8610
         Picture         =   "frmEtiquetaKRUPP.frx":0D69
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Limpar tela para nova etiqueta"
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   495
         Left            =   8040
         Picture         =   "frmEtiquetaKRUPP.frx":11AB
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Confirma dados do filtro"
         Top             =   270
         Width           =   555
      End
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   3345
         Begin VB.OptionButton Opt_Pallet 
            Caption         =   "Pelo Pallet"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Opt_etiqueta 
            Caption         =   "Pelo N� da Etiqueta"
            Height          =   255
            Left            =   1500
            TabIndex        =   3
            Top             =   300
            Width           =   1755
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digita��o:"
         Height          =   195
         Left            =   3570
         TabIndex        =   8
         Top             =   510
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd_Visualizar 
      Caption         =   "Visualizar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6480
      TabIndex        =   0
      Top             =   5370
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   300
      TabIndex        =   22
      Top             =   6030
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label lbl_qtd 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
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
      Left            =   1350
      TabIndex        =   18
      Top             =   5730
      Visible         =   0   'False
      Width           =   45
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
      Left            =   300
      TabIndex        =   17
      Top             =   5700
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl_peca 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
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
      Left            =   1350
      TabIndex        =   16
      Top             =   5475
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbl_produto 
      AutoSize        =   -1  'True
      Caption         =   "Pe�a.:"
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
      Left            =   300
      TabIndex        =   15
      Top             =   5460
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lbl_sequencia 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
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
      Left            =   1350
      TabIndex        =   14
      Top             =   5220
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label label33 
      AutoSize        =   -1  'True
      Caption         =   "Sequ�ncia.:"
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
      Left            =   300
      TabIndex        =   13
      Top             =   5220
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmEtiquetaKRUPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Vari�vel para MDIapp
Public cRec As ADODB.Recordset 'conter� os dados do registro corrente
Public bAtivo As Boolean
Public bTelaImp As Boolean
Public bJafoi As Boolean
Public nLogin As Integer ' conter� o codigo do usuario, quando confirmar a senha
Public nTipo As Integer ' conter� o tipo do usuario, quando confirmar a senha
Public nMatricula As Double ' conter� a matricula do usuario, quando confirmar a senha

Private Sub btoConfirma_Click()

Dim nx As Integer
Dim sData As String
Dim sNota As String
Dim sLote As String

On Error GoTo Erro

Me.MousePointer = vbHourglass
Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False

If Me.Opt_Pallet.Value = True Then
   sData = "1"
Else
   sData = "2"
End If
Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_Filtro_KRUPP(sBancoMusashi, _
                                                            Me.txtsequencial.Text, _
                                                            sData)

If cRec.RecordCount > 0 Then
   Call carrega_Grid
   Me.cmd_Impressao.Enabled = True
   Me.cmd_Visualizar.Enabled = True
   Grid1.col = 0
   Me.lbl_produto.Visible = True
   Me.lbl_qtd.Visible = True
   Me.lbl_sequencia.Visible = True
   Me.lbl_peca.Visible = True
   Me.label33.Visible = True
   Me.Label5.Visible = True
   Rem criticas a serem realizadas no arquivo caso seja por pallet
'   If cRec.RecordCount > 0 And Me.Opt_Pallet.Value = True Then
'      cRec.MoveFirst
'      If IsNull(cRec!XBLNR) Then
'         MsgBox "Palet sem sem Nota fiscal. Verifique na tela."
'         Me.MousePointer = vbDefault
'         Me.cmd_Impressao.Enabled = False
'         Me.cmd_Visualizar.Enabled = False
'         Exit Sub
'      End If
'      sNota = Trim(cRec!XBLNR)
'      sLote = Trim(cRec!Num_Lote)
'      While Not cRec.EOF
'            If sNota <> Trim(cRec!XBLNR) Then
'               MsgBox "Existem Notas fiscais diferentes dentro de um mesmo Pallet. Verifique na tela."
'               Me.cmd_Impressao.Enabled = False
'               Me.cmd_Visualizar.Enabled = False
'            End If
'            If sLote <> Trim(cRec!Num_Lote) Then
'               MsgBox "Existem Lotes diferentes dentro mesmo Pallet. Verifique na tela."
'               Me.cmd_Impressao.Enabled = False
'               Me.cmd_Visualizar.Enabled = False
'            End If
'            cRec.MoveNext
'      Wend
'   End If
   
End If

Me.MousePointer = vbDefault

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
If Err.Number = 50000 Then
   Call cmd_limpar_Click
End If
If Err.Number = 50001 Then
   Call Limpar_Grid
   Me.txtsequencial.SetFocus
End If
End Sub

Private Sub cmd_Impressao_Click()

'If Me.Opt_Pallet.Value = True Then
'   Call Imprime_Por_Pallet
'Else
   Call Imprime_por_etiqueta
'End If

End Sub
Private Function Imprime_Por_Pallet()

'rpt_Etiquetas_RelEtiqueta_FIAT_LATAM
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), "ETIQUETA FABRICA") > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "impressora da Etiqueta n�o encontrada, chame o respons�vel. o nome da impressora ser� - 'ETIQUETA FABRICA'"
   Close
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
rs.Fields.Append "16_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Denominacao", ADODB.DataTypeEnum.adChar, 80
rs.Fields.Append "18_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Ctrl_Interno", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "20_Ctrl_Oper_Log", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "21_Codigo_Numero", ADODB.DataTypeEnum.adChar, 50
rs.Fields.Append "22_codigo_barras", ADODB.DataTypeEnum.adChar, 50

rs.Open

cRec.MoveFirst
nx = 0
nqtde = 0
If Grid1.Rows > 0 Then
   
   rs.AddNew
 
   Grid1.col = 0: rs.Fields("1_Num_Lote").Value = Trim(Grid1.Text)
   If Me.Opt_Pallet.Value = True Then
      rs.Fields("3_Classe_Func").Value = Me.txtsequencial.Text
   Else
      Grid1.col = 2: rs.Fields("3_Classe_Func").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 3: rs.Fields("4_Indicacao_Supl").Value = Trim(Grid1.Text)
   Grid1.col = 4
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("5_Data_Fab_Lote").Value = " "
   Else
      rs.Fields("5_Data_Fab_Lote").Value = Trim(Grid1.Text)
   End If
'   Grid1.Col = 4: rs.Fields("5_Data_Fab_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 5: rs.Fields("6_Cod_Emb").Value = Trim(Grid1.Text)
   Grid1.col = 6: rs.Fields("7_Vinculo").Value = Trim(Grid1.Text)
   Grid1.col = 7: rs.Fields("8_Lote_Sob_Desv").Value = Trim(Grid1.Text)
   Grid1.col = 8: rs.Fields("9_Qtde_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 9: rs.Fields("10_Aplicacao").Value = Trim(Grid1.Text)
   Grid1.col = 10
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("11_DUM").Value = " "
   Else
      rs.Fields("11_DUM").Value = Trim(Grid1.Text)
   End If
'   Grid1.Col = 10: rs.Fields("11_DUM").Value = Trim(Grid1.Text)
   Grid1.col = 11: rs.Fields("12_Embarque_Ctrl").Value = Trim(Grid1.Text)
   Grid1.col = 12: rs.Fields("13_Cod_Fornecedor").Value = Trim(Grid1.Text)
   Grid1.col = 13: rs.Fields("14_Num_Doc_Fis_BAM").Value = Trim(Grid1.Text)
   Grid1.col = 14: rs.Fields("15_Data").Value = Trim(Grid1.Text)
   Grid1.col = 15: rs.Fields("16_Ponto_Entrega").Value = Trim(Grid1.Text)
   Grid1.col = 16: rs.Fields("17_Denominacao").Value = Trim(Grid1.Text)
   Grid1.col = 17: rs.Fields("18_Num_Desenho").Value = Trim(Grid1.Text)
   Grid1.col = 18: rs.Fields("19_Ctrl_Interno").Value = Trim(Grid1.Text)
   Grid1.col = 19: rs.Fields("20_Ctrl_Oper_Log").Value = Trim(Grid1.Text)
   
   Grid1.col = 17: sTexto = Trim(Grid1.Text)
   Grid1.col = 1: sTexto = sTexto & Format(Trim(Grid1.Text), "00000")
   Grid1.col = 5: sTexto = sTexto & Trim(Grid1.Text) & "013093"
   Grid1.col = 20: rs.Fields("21_Codigo_Numero").Value = sTexto
   Grid1.col = 21: rs.Fields("22_codigo_barras").Value = "*" & sTexto & "*"
 
   For nx = 1 To Grid1.Rows - 1
       Grid1.col = 1: nqtde = nqtde + VBA.CDbl(Trim(Grid1.Text))
       Grid1.Row = nx
   Next
   
   Grid1.col = 1: rs.Fields("2_Qtde_Emb").Value = Format(nqtde, "000")
   rs.Update
   
End If

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_P.rpt")

CrystalReport1.Database.SetDataSource rs

rs.Clone

CrystalReport1.PrintOutEx False

Me.MousePointer = vbDefault

Exit Function

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Function

Private Function Imprime_por_etiqueta()

'rpt_Etiquetas_RelEtiqueta_FIAT_LATAM
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double
Dim sCod_Barras As String

nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), UCase(Me.cbo_impressora.List(Me.cbo_impressora.ListIndex))) > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "Impressora da Etiqueta n�o encontrada, chame o respons�vel. o nome da impressora ser� - 'ETIQUETA FABRICA'"
   Exit Function
End If

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
rs.Fields.Append "16_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Denominacao", ADODB.DataTypeEnum.adChar, 80
rs.Fields.Append "18_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Ctrl_Interno", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "20_Ctrl_Oper_Log", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "21_Codigo_Numero", ADODB.DataTypeEnum.adChar, 50
rs.Fields.Append "22_codigo_barras", ADODB.DataTypeEnum.adChar, 50


rs.Open
nx = 0
cRec.MoveFirst
nx = 0
nqtde = 0
If Grid1.Rows > 0 Then

'    Grid1.Col = 0: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Lote")), " ", cRec.Fields("Num_Lote")) '8_lote_sob_desv
'    Grid1.Col = 1: Grid1.Text = IIf(IsNull(cRec.Fields("Qtde_Emb")), " ", cRec.Fields("Qtde_Emb")) '3_classe_func
'    Grid1.Col = 5: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Emb")), " ", cRec.Fields("Cod_Emb")) '7_vinculo
'    Grid1.Col = 7: Grid1.Text = IIf(IsNull(cRec.Fields("Lote_Sob_Desv")), " ", cRec.Fields("Lote_Sob_Desv")) '5_data_fab_lote
'    Grid1.Col = 17: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", cRec.Fields("Num_Desenho")) '2_qtd_emp
'    Grid1.Col = 24: Grid1.Text = IIf(IsNull(cRec.Fields("cod_peca")), " ", cRec.Fields("cod_peca")) '1_num_lote
'    Grid1.Col = 23: Grid1.Text = IIf(IsNull(cRec.Fields("xblnr")), " ", cRec.Fields("xblnr")) '4_indicacao_supl
'    Grid1.Col = 25: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Doc_Fiscal")), " ", cRec.Fields("Num_Doc_Fiscal"))


'   For nx = 1 To Grid1.Rows - 1
       rs.AddNew
       Rem levarei em cosideracao a ordem de impressao na etiqueta(sequencia de campos), sem me preocupar com os nomes descritos.
       Grid1.col = 24: rs.Fields("1_Num_Lote").Value = Trim(Grid1.Text) ' cod_peca
       Grid1.col = 17: rs.Fields("2_Qtde_Emb").Value = Format(Trim(Grid1.Text), "000000000") ' Num_Desenho ->cod_no_cliente
'       Grid1.Col = 1: rs.Fields("3_Classe_Func").Value = Trim(Grid1.Text) ' qtde->qtde_emp
       Grid1.col = 23: rs.Fields("4_Indicacao_Supl").Value = "5500025925"  'IIf(Len(Trim(Grid1.Text)) = 0, "0000000000", Format(Trim(Grid1.Text), "0000000000")) '  Num_Doc_Fiscal
       Grid1.col = 7: rs.Fields("5_Data_Fab_Lote").Value = IIf(Len(Trim(Grid1.Text)) = 0, "000", Format(Trim(Grid1.Text), "000")) ' Lote_Sob_Desv
       Grid1.col = 23: rs.Fields("6_Cod_Emb").Value = "3000079296"  ' IIf(Len(Trim(Grid1.Text)) = 0, "000", Format(Trim(Grid1.Text), "000")) ' cod_emp
       Grid1.col = 5: rs.Fields("7_Vinculo").Value = IIf(Len(Trim(Grid1.Text)) = 0, "0", Trim(Grid1.Text)) 'cod_emb
       Grid1.col = 0: rs.Fields("8_Lote_Sob_Desv").Value = Trim(Grid1.Text) 'num_lote
       rs.Fields("9_Qtde_Lote").Value = Format(Now(), "DD/MM/YYYY") 'DATA EMISS�O
       rs.Fields("17_Denominacao").Value = "MUSASHI DO BRASIL"
       sCod_Barras = "*" & Trim(rs.Fields("4_Indicacao_Supl").Value) & _
                     Format(Trim(rs.Fields("2_Qtde_Emb").Value), "000000") & _
                     Format(Trim(rs.Fields("7_Vinculo").Value), "0000000000") & _
                     Trim(rs.Fields("3_Classe_Func").Value) & _
                     Trim(rs.Fields("8_Lote_Sob_Desv").Value) & "*"
       rs.Fields("22_codigo_barras").Value = sCod_Barras 'codigo_barras
       
       For nx = 1 To Grid1.Rows - 1
           Grid1.col = 1: nqtde = nqtde + VBA.CDbl(Trim(Grid1.Text))
           Grid1.Row = nx
       Next
       
       rs.Fields("3_Classe_Func").Value = str(nqtde)  ' Num_Desenho ->cod_no_cliente

       rs.Update
                
'       Grid1.Row = nx
'       If nx = 3 Then Exit For
'      Next
   
End If

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_KRUPPok.rpt")

CrystalReport1.Database.SetDataSource rs

rs.Clone

CrystalReport1.PrintOutEx False

Me.MousePointer = vbDefault

Exit Function

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault
End Function

Private Sub cmd_Impressao1_Click()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim oTela As frmEscRelCristalReport
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

Set oTela = New frmEscRelCristalReport

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set oTela = New frmEscRelCristalReport
Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_LATAM.rpt")
oTela.CRViewer1.ReportSource = CrystalReport1
oTela.CRViewer1.ViewReport

oTela.Show 0

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub
Private Sub cmd_Impressao2_Click()
'rpt_Etiquetas_RelEtiqueta_FIAT_LATAM
Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), "ETIQUETA FABRICA") > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "impressora da Etiqueta n�o encontrada, chame o respons�vel. o nome da impressora ser� - 'ETIQUETA FABRICA'"
   Close
End If
nx = 0

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_LATAM.rpt")

CrystalReport1.PrintOutEx False

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub
Private Sub cmd_limpar_Click()
Call Fechar_Form_Etiqueta

Call Limpar_Grid

Me.txtsequencial.Text = ""

Me.lbl_produto.Visible = False
Me.lbl_qtd.Visible = False
Me.lbl_sequencia.Visible = False
Me.lbl_peca.Visible = False
Me.label33.Visible = False
Me.Label5.Visible = False
Me.txtsequencial.Enabled = True

Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmd_visualizar_Click()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim oTela As frmEscRelCristalReport
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double
Dim sCod_Barras As String * 50

Set oTela = New frmEscRelCristalReport

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
rs.Fields.Append "16_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Denominacao", ADODB.DataTypeEnum.adChar, 80
rs.Fields.Append "18_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Ctrl_Interno", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "20_Ctrl_Oper_Log", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "21_Codigo_Numero", ADODB.DataTypeEnum.adChar, 50
rs.Fields.Append "22_codigo_barras", ADODB.DataTypeEnum.adChar, 50


rs.Open
nx = 0
cRec.MoveFirst
nx = 0
nqtde = 0
If Grid1.Rows > 0 Then

'    Grid1.Col = 0: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Lote")), " ", cRec.Fields("Num_Lote")) '8_lote_sob_desv
'    Grid1.Col = 1: Grid1.Text = IIf(IsNull(cRec.Fields("Qtde_Emb")), " ", cRec.Fields("Qtde_Emb")) '3_classe_func
'    Grid1.Col = 5: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Emb")), " ", cRec.Fields("Cod_Emb")) '7_vinculo
'    Grid1.Col = 7: Grid1.Text = IIf(IsNull(cRec.Fields("Lote_Sob_Desv")), " ", cRec.Fields("Lote_Sob_Desv")) '5_data_fab_lote
'    Grid1.Col = 17: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", cRec.Fields("Num_Desenho")) '2_qtd_emp
'    Grid1.Col = 24: Grid1.Text = IIf(IsNull(cRec.Fields("cod_peca")), " ", cRec.Fields("cod_peca")) '1_num_lote
'    Grid1.Col = 23: Grid1.Text = IIf(IsNull(cRec.Fields("xblnr")), " ", cRec.Fields("xblnr")) '4_indicacao_supl
'    Grid1.Col = 25: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Doc_Fiscal")), " ", cRec.Fields("Num_Doc_Fiscal"))


'   For nx = 1 To Grid1.Rows - 1
       rs.AddNew
       Rem levarei em cosideracao a ordem de impressao na etiqueta(sequencia de campos), sem me preocupar com os nomes descritos.
       Grid1.col = 24: rs.Fields("1_Num_Lote").Value = Trim(Grid1.Text) ' cod_peca
       Grid1.col = 17: rs.Fields("2_Qtde_Emb").Value = Format(Trim(Grid1.Text), "000000000") ' Num_Desenho ->cod_no_cliente
'       Grid1.Col = 1: rs.Fields("3_Classe_Func").Value = Trim(Grid1.Text) ' qtde->qtde_emp
       Grid1.col = 23: rs.Fields("4_Indicacao_Supl").Value = "5500025925"  'IIf(Len(Trim(Grid1.Text)) = 0, "0000000000", Format(Trim(Grid1.Text), "0000000000")) '  Num_Doc_Fiscal
       Grid1.col = 7: rs.Fields("5_Data_Fab_Lote").Value = IIf(Len(Trim(Grid1.Text)) = 0, "000", Format(Trim(Grid1.Text), "000")) ' Lote_Sob_Desv
       Grid1.col = 23: rs.Fields("6_Cod_Emb").Value = "3000079296"  ' IIf(Len(Trim(Grid1.Text)) = 0, "000", Format(Trim(Grid1.Text), "000")) ' cod_emp
       Grid1.col = 5: rs.Fields("7_Vinculo").Value = IIf(Len(Trim(Grid1.Text)) = 0, "0", Trim(Grid1.Text)) 'cod_emb
       Grid1.col = 0: rs.Fields("8_Lote_Sob_Desv").Value = Trim(Grid1.Text) 'num_lote
       rs.Fields("9_Qtde_Lote").Value = Format(Now(), "DD/MM/YYYY") 'DATA EMISS�O
       rs.Fields("17_Denominacao").Value = "MUSASHI DO BRASIL"
       sCod_Barras = "*" & Trim(rs.Fields("4_Indicacao_Supl").Value) & _
                     Format(Trim(rs.Fields("2_Qtde_Emb").Value), "000000") & _
                     Format(Trim(rs.Fields("7_Vinculo").Value), "0000000000") & _
                     Trim(rs.Fields("3_Classe_Func").Value) & _
                     Trim(rs.Fields("8_Lote_Sob_Desv").Value) & "*"
       rs.Fields("22_codigo_barras").Value = sCod_Barras 'codigo_barras
       
       For nx = 1 To Grid1.Rows - 1
           Grid1.col = 1: nqtde = nqtde + VBA.CDbl(Trim(Grid1.Text))
           Grid1.Row = nx
       Next
       
       rs.Fields("3_Classe_Func").Value = str(nqtde)  ' Num_Desenho ->cod_no_cliente

       rs.Update
                
'       Grid1.Row = nx
'       If nx = 3 Then Exit For
'      Next
   
End If

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_KRUPPok.rpt")

CrystalReport1.Database.SetDataSource rs

rs.Clone
oTela.CRViewer1.ReportSource = CrystalReport1
oTela.CRViewer1.ViewReport

oTela.Show 0

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

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
Me.txtsequencial.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > -1 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   If Me.ActiveControl.TabIndex < 1 Then
      If 6 = MsgBox("Deseja realmente sair deste m�dulo?", 32 + 4) Then
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
   MsgBox "Impressoras ETIQUETAS,n�o encontradas no sistema, Favor comunicar ao respons�vel para adiciona-las no sistema!"
   End
End If
Me.cbo_impressora.ListIndex = 0

Rem verificar a impressora padr�o para ser usada neste relat�rio
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

    If Opt_etiqueta.Value = True And Len(Trim(cRec.Fields("xblnr"))) = 0 Then
       MsgBox "Etiqueta sem Nota fiscal. Verifique!"
    Else
       Grid1.col = 0: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Lote")), " ", cRec.Fields("Num_Lote")) '8_lote_sob_desv
       Grid1.col = 1: Grid1.Text = IIf(IsNull(cRec.Fields("Qtde_Emb")), " ", cRec.Fields("Qtde_Emb")) '3_classe_func
       Grid1.col = 2: Grid1.Text = IIf(IsNull(cRec.Fields("Classe_Func")), " ", cRec.Fields("Classe_Func"))
       Grid1.col = 3: Grid1.Text = IIf(IsNull(cRec.Fields("Indicacao_Supl")), " ", cRec.Fields("Indicacao_Supl"))
       Grid1.col = 4: Grid1.Text = IIf(IsNull(cRec.Fields("Data_Fab_Lote")), " ", cRec.Fields("Data_Fab_Lote"))
       Grid1.col = 5: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Emb")), " ", cRec.Fields("Cod_Emb")) '7_vinculo
       Grid1.col = 6: Grid1.Text = IIf(IsNull(cRec.Fields("Vinculo")), " ", cRec.Fields("Vinculo"))
       Grid1.col = 7: Grid1.Text = IIf(IsNull(cRec.Fields("Lote_Sob_Desv")), " ", cRec.Fields("Lote_Sob_Desv")) '5_data_fab_lote
       Grid1.col = 8: Grid1.Text = IIf(IsNull(cRec.Fields("Qtde_Lote")), " ", cRec.Fields("Qtde_Lote"))
       Grid1.col = 9: Grid1.Text = IIf(IsNull(cRec.Fields("Aplicacao")), " ", cRec.Fields("Aplicacao"))
       Grid1.col = 10: Grid1.Text = IIf(IsNull(cRec.Fields("DUM")), " ", cRec.Fields("DUM"))
       Grid1.col = 11: Grid1.Text = IIf(IsNull(cRec.Fields("Embarque_Ctrl")), " ", cRec.Fields("Embarque_Ctrl"))
       Grid1.col = 12: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", cRec.Fields("Cod_Fornecedor"))
       Grid1.col = 13: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Doc_Fis_BAM")), " ", cRec.Fields("Num_Doc_Fis_BAM"))
       Grid1.col = 14: Grid1.Text = IIf(IsNull(cRec.Fields("Data")), " ", cRec.Fields("Data"))
       Grid1.col = 15: Grid1.Text = IIf(IsNull(cRec.Fields("Ponto_Entrega")), " ", cRec.Fields("Ponto_Entrega"))
       Grid1.col = 16: Grid1.Text = IIf(IsNull(cRec.Fields("Denominacao")), " ", cRec.Fields("Denominacao"))
       Grid1.col = 17: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", cRec.Fields("Num_Desenho")) '2_qtd_emp
       Grid1.col = 18: Grid1.Text = IIf(IsNull(cRec.Fields("Ctrl_Interno")), " ", cRec.Fields("Ctrl_Interno"))
       Grid1.col = 19: Grid1.Text = IIf(IsNull(cRec.Fields("Ctrl_Oper_Log")), " ", cRec.Fields("Ctrl_Oper_Log"))
       Grid1.col = 20: Grid1.Text = IIf(IsNull(cRec.Fields("Codigo_Numero")), " ", cRec.Fields("Codigo_Numero"))
       Grid1.col = 21: Grid1.Text = IIf(IsNull(cRec.Fields("codigo_barras")), " ", cRec.Fields("codigo_barras"))
       Grid1.col = 22: Grid1.Text = IIf(IsNull(cRec.Fields("Pallet")), " ", cRec.Fields("Pallet"))
       Grid1.col = 23: Grid1.Text = IIf(IsNull(cRec.Fields("xblnr")), " ", cRec.Fields("xblnr")) '4_indicacao_supl
       Grid1.col = 24: Grid1.Text = IIf(IsNull(cRec.Fields("cod_peca")), " ", cRec.Fields("cod_peca")) '1_num_lote
       Grid1.col = 25: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Doc_Fiscal")), " ", cRec.Fields("Num_Doc_Fiscal"))
       
       Grid1.col = 19: Me.lbl_sequencia.Caption = Me.Grid1.Text
       Grid1.col = 18: Me.lbl_peca.Caption = Me.Grid1.Text
       Grid1.col = 1: nqtde = nqtde + VBA.CDbl(Me.Grid1.Text)
       Grid1.col = 0
    End If
    
    cRec.MoveNext
    
    If Not cRec.EOF Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Row = Grid1.Row + 1
    End If
Next

Me.lbl_qtd.Caption = nqtde

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
Grid1.col = 2: Grid1.BackColor = &H80FFFF
Grid1.col = 0: Grid1.ColWidth(0) = 1000: Grid1.Text = "1_Num_Lote"
Grid1.col = 1: Grid1.ColWidth(0) = 1000: Grid1.Text = "2_Qtde_Emb"
Grid1.col = 2: Grid1.ColWidth(0) = 1000: Grid1.Text = "3_Classe_Func"
Grid1.col = 3: Grid1.ColWidth(0) = 1000: Grid1.Text = "4_Indicacao_Supl"
Grid1.col = 4: Grid1.ColWidth(0) = 1000: Grid1.Text = "5_Data_Fab_Lote"
Grid1.col = 5: Grid1.ColWidth(0) = 1000: Grid1.Text = "6_Cod_Emb"
Grid1.col = 6: Grid1.ColWidth(0) = 1000: Grid1.Text = "7_Vinculo"
Grid1.col = 7: Grid1.ColWidth(0) = 1000: Grid1.Text = "8_Lote_Sob_Desv"
Grid1.col = 8: Grid1.ColWidth(0) = 1000: Grid1.Text = "9_Qtde_Lote"
Grid1.col = 9: Grid1.ColWidth(0) = 1000: Grid1.Text = "10_Aplicacao"
Grid1.col = 10: Grid1.ColWidth(0) = 1000: Grid1.Text = "11_DUM"
Grid1.col = 11: Grid1.ColWidth(0) = 1000: Grid1.Text = "12_Embarque_Ctrl"
Grid1.col = 12: Grid1.ColWidth(0) = 1000: Grid1.Text = "13_Cod_Fornecedor"
Grid1.col = 13: Grid1.ColWidth(0) = 1000: Grid1.Text = "14_Num_Doc_Fis_BAM"
Grid1.col = 14: Grid1.ColWidth(0) = 1000: Grid1.Text = "15_Data"
Grid1.col = 15: Grid1.ColWidth(0) = 1000: Grid1.Text = "16_Ponto_Entrega"
Grid1.col = 16: Grid1.ColWidth(0) = 1000: Grid1.Text = "17_Denominacao String 80"
Grid1.col = 17: Grid1.ColWidth(0) = 1000: Grid1.Text = "18_Num_Desenho"
Grid1.col = 18: Grid1.ColWidth(0) = 1000: Grid1.Text = "19_Ctrl_Interno"
Grid1.col = 19: Grid1.ColWidth(0) = 1000: Grid1.Text = "20_Ctrl_Oper_Log"
Grid1.col = 20: Grid1.ColWidth(0) = 1000: Grid1.Text = "21_Codigo_Numero"
Grid1.col = 21: Grid1.ColWidth(0) = 1000: Grid1.Text = "22_codigo_barras"
Grid1.col = 22: Grid1.ColWidth(0) = 1000: Grid1.Text = "Pallet  "
Grid1.col = 23: Grid1.ColWidth(0) = 1000: Grid1.Text = "N.Fiscal"
Grid1.col = 24: Grid1.ColWidth(0) = 1000: Grid1.Text = "Peca"
Grid1.col = 25: Grid1.ColWidth(0) = 1000: Grid1.Text = "N.Fiscal"
Grid1.Row = 0

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
Grid1.ColAlignment(18) = flexAlignLeftCenter
Grid1.ColAlignment(19) = flexAlignLeftCenter
Grid1.ColAlignment(20) = flexAlignLeftCenter
Grid1.ColAlignment(21) = flexAlignLeftCenter
Grid1.ColAlignment(22) = flexAlignLeftCenter
Grid1.ColAlignment(23) = flexAlignLeftCenter
Grid1.ColAlignment(24) = flexAlignLeftCenter
Grid1.ColAlignment(25) = flexAlignLeftCenter

'Me.lbl_produto.Visible = True
'Me.label33.Visible = True
'Me.Label5.Visible = True
'Me.lbl_qtd.Visible = True
'Me.lbl_sequencia.Visible = True
'Me.lbl_peca.Visible = True

'Me.lbl_qtd.Caption = ""
'Me.lbl_sequencia.Caption = ""
'Me.lbl_peca.Caption = ""

End Sub

'Private Sub Grid1_Click()
'Me.Grid1.SelectionMode = 1
'Me.Grid1.BackColorSel = 33
'Grid1.Col = 0
'If Grid1.Rows > 1 And Len(Trim(Me.Grid1.Text)) > 0 Then
''   Grid1.Col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
''   Grid1.Col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
''   Grid1.Col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
''   Grid1.Col = 0
'   Me.cmd_Impressao.Enabled = True
'   Me.cmd_Visualizar.Enabled = True
'End If
'End Sub

'Private Sub Grid1_GotFocus()
'
'Grid1.Col = 0
'If Grid1.Rows > 1 And Len(Trim(Me.Grid1.Text)) > 0 Then
'   Me.lbl_produto.Visible = True
'   Me.lbl_qtd.Visible = True
'   Me.lbl_sequencia.Visible = True
'   Me.lbl_peca.Visible = True
'   Me.label33.Visible = True
'   Me.Label5.Visible = True
'Else
'   Me.lbl_produto.Visible = False
'   Me.lbl_qtd.Visible = False
'   Me.lbl_sequencia.Visible = False
'   Me.lbl_peca.Visible = False
'   Me.label33.Visible = False
'   Me.Label5.Visible = False
'End If
'
'End Sub


'Private Sub Grid1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   cmd_visualizar_Click
'End If
'End Sub
'

'Private Sub Grid1_SelChange()
'   If Grid1.Rows > 2 Then
'      Grid1.Col = 0: Me.lbl_sequencia.Caption = Me.Grid1.Text
'      Grid1.Col = 1: Me.lbl_peca.Caption = Me.Grid1.Text
'      Grid1.Col = 5: Me.lbl_qtd.Caption = Me.Grid1.Text
'      Grid1.Col = 0
''      Grid1.BackColorSel = &HFF00&
'   End If

'End Sub

Private Sub txtsequencial_Change()
If Len(Trim(txtsequencial.Text)) = 11 Then
   Call btoConfirma_Click
End If

End Sub

Private Sub txtsequencial_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   If MsgBox("Deseja sair deste m�dulo?", vbQuestion + vbYesNo, "ATEN��O !!!") = vbNo Then
      Me.txtsequencial.Text = ""
      Me.txtsequencial.SetFocus
   Else
      Unload Me
   End If
End If

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

Next
        
Unload frmOpcoes

End Function






