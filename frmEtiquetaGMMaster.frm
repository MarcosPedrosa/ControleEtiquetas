VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEtiquetaGMMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas GM Master"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9480
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
      Left            =   6420
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   5820
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      Caption         =   "Totais"
      Enabled         =   0   'False
      Height          =   1125
      Left            =   2340
      TabIndex        =   29
      Top             =   5010
      Width           =   3675
      Begin VB.TextBox txt_Peso 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
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
         Left            =   210
         MaxLength       =   11
         TabIndex        =   11
         ToolTipText     =   "Digite a Transportadora"
         Top             =   510
         Width           =   795
      End
      Begin VB.TextBox txt_TotalPeso 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   32
         ToolTipText     =   "Digite a Transportadora"
         Top             =   600
         Width           =   1185
      End
      Begin VB.TextBox txt_TotalPecas 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   30
         ToolTipText     =   "Digite a Transportadora"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Peso p/Peça:"
         Height          =   195
         Left            =   90
         TabIndex        =   34
         Top             =   270
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Peso:"
         Height          =   195
         Left            =   1320
         TabIndex        =   33
         Top             =   630
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Peças:"
         Height          =   195
         Left            =   1320
         TabIndex        =   31
         Top             =   270
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quantidades"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   90
      TabIndex        =   26
      Top             =   5010
      Width           =   2175
      Begin VB.TextBox txt_QtdeCaixa 
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
         Height          =   315
         Left            =   1170
         MaxLength       =   11
         TabIndex        =   5
         ToolTipText     =   "Digite a Transportadora"
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox txt_QtdePecaCaixa 
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
         Height          =   315
         Left            =   1170
         MaxLength       =   11
         TabIndex        =   6
         ToolTipText     =   "Digite a Transportadora"
         Top             =   660
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Qtde Caixas:"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peças p/Caixa:"
         Height          =   195
         Left            =   90
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.TextBox TXT_EMBALAGEM 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   18
      ToolTipText     =   "Digie a sequencial da etiqueta"
      Top             =   6570
      Width           =   1515
   End
   Begin VB.CommandButton cmd_Impressao_Master 
      Enabled         =   0   'False
      Height          =   735
      Left            =   3060
      Picture         =   "frmEtiquetaGMMaster.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprimir"
      Top             =   6330
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3705
      Left            =   90
      TabIndex        =   16
      Top             =   1260
      Width           =   9285
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3315
         Left            =   120
         TabIndex        =   4
         Top             =   180
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   16
         ForeColorFixed  =   16711680
         BackColorSel    =   65535
         ForeColorSel    =   65535
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLines       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmEtiquetaGMMaster.frx":030A
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
      Left            =   8430
      Picture         =   "frmEtiquetaGMMaster.frx":0313
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Fecha a tela de etiqueta"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7410
      Picture         =   "frmEtiquetaGMMaster.frx":0755
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir"
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame frm_filtro 
      Caption         =   "Filtro da Etiqueta"
      Height          =   1065
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   9285
      Begin VB.TextBox txtsequencial 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2520
         MaxLength       =   16
         TabIndex        =   1
         ToolTipText     =   "Digie a sequencial da etiqueta"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   495
         Left            =   8610
         Picture         =   "frmEtiquetaGMMaster.frx":0A5F
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Limpar tela para nova etiqueta"
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   495
         Left            =   8040
         Picture         =   "frmEtiquetaGMMaster.frx":0EA1
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Confirma dados do filtro"
         Top             =   270
         Width           =   555
      End
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   1425
         Begin VB.OptionButton Opt_Pallet 
            Caption         =   "Pelo Pallet"
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
            TabIndex        =   14
            Top             =   270
            Width           =   1755
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digitação:"
         Height          =   195
         Left            =   1770
         TabIndex        =   15
         Top             =   510
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd_Visualizar 
      Caption         =   "Visualizar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6420
      TabIndex        =   8
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "NF.Embalagem:"
      Height          =   195
      Left            =   300
      TabIndex        =   25
      Top             =   6630
      Width           =   1125
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
      Left            =   9180
      TabIndex        =   24
      Top             =   6780
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
      Left            =   8130
      TabIndex        =   23
      Top             =   6750
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
      Left            =   9180
      TabIndex        =   22
      Top             =   6525
      Visible         =   0   'False
      Width           =   45
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
      Left            =   8130
      TabIndex        =   21
      Top             =   6510
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
      Left            =   9180
      TabIndex        =   20
      Top             =   6270
      Visible         =   0   'False
      Width           =   45
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
      Left            =   8130
      TabIndex        =   19
      Top             =   6270
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmEtiquetaGMMaster"
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
Private sdataCriacao As Date
Private sdataCriacaoMenor As Date
Private nSequencia_Escolhida As Double
Private Declare Sub Sleep Lib "KERNEL32" _
        (ByVal dwMilliseconds As Long)
Private sDataHoje As String
Private Sub btoConfirma_Click()

Dim nx As Integer
Dim sData As String
Dim sNota As String
Dim sLote As String


On Error GoTo Erro

Me.MousePointer = vbHourglass
Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False
Me.cmd_Impressao_Master.Enabled = False

If Me.Opt_Pallet.Value = True Then
   sData = "1"
Else
   sData = "2"
End If
Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_Filtro_GM_MASTER(sBancoMusashi, _
                                                                Me.txtsequencial.Text, _
                                                                sData)

If cRec.RecordCount > 0 Then
   Me.Frame3.Enabled = True
   Me.Frame4.Enabled = True
   Me.Grid1.Visible = False
   Call carrega_Grid
   Me.Grid1.Visible = True
   Me.cmd_Impressao.Enabled = True
   Me.cmd_Visualizar.Enabled = True
   Me.cmd_Impressao_Master.Enabled = True
   Grid1.col = 0
   Me.lbl_produto.Visible = True
   Me.lbl_qtd.Visible = True
   Me.lbl_sequencia.Visible = True
   Me.lbl_peca.Visible = True
   Me.label33.Visible = True
   Me.Label5.Visible = True
   Rem criticas a serem realizadas no arquivo caso seja por pallet
   If cRec.RecordCount > 0 And Me.Opt_Pallet.Value = True Then
      cRec.MoveFirst
      If IsNull(cRec!Num_Doc_Fiscal) Then
         MsgBox "Pallet sem Nota fiscal. Verifique na tela."
         Me.MousePointer = vbDefault
         Me.cmd_Impressao.Enabled = False
         Me.cmd_Visualizar.Enabled = False
         Me.cmd_Impressao_Master.Enabled = False
         Exit Sub
      End If
      sNota = Trim(cRec!Num_Doc_Fiscal)
      sLote = Trim(cRec!Num_Lote)
      While Not cRec.EOF
            If sNota <> Trim(cRec!Num_Doc_Fiscal) Then
'               MsgBox "Existem Notas fiscais diferentes dentro de um mesmo Pallet. Verifique na tela."
'               Me.cmd_Impressao.Enabled = False
'               Me.cmd_Visualizar.Enabled = False
'               Me.cmd_Impressao_Master.Enabled = False
            End If
            If sLote <> Trim(cRec!Num_Lote) Then
'               MsgBox "Existem Lotes diferentes dentro mesmo Pallet. Verifique na tela."
'               Me.cmd_Impressao.Enabled = False
'               Me.cmd_Visualizar.Enabled = False
            End If
            cRec.MoveNext
      Wend
   End If
   
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
Dim nx As Double
Dim x As Printer

On Error GoTo Erro

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

Printer.Orientation = 2
Call cmd_visualizar_Click
Printer.Orientation = 2
frmEtiquetaGMMasterEtiq.PrintForm
Printer.Orientation = 2: Printer.EndDoc
Me.MousePointer = vbDefault
Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_Impressao_Master_Click()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim oTela As frmEscRelCristalReport
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

If Len(Trim(Me.TXT_EMBALAGEM.Text)) = 0 Then
   MsgBox "Digite a Embalagem, para emitir o Relatório de Etiqueta Master!"
   Me.TXT_EMBALAGEM.SetFocus
   Exit Sub
End If

Set oTela = New frmEscRelCristalReport

On Error GoTo Erro

Set rs = New ADODB.Recordset

rs.Fields.Append "1_Peso_Bruto", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_ODM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Data_Producao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Data_validade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Expedicao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Quantidade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_DOCA", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Control_Log_Qual", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Num_Doc_Fiscal", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Qtde_emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Num_Sheda_Serial", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Id_Inter_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "18_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "20_Dados_Transporte", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "21_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "22_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "23_Razao_Social", ADODB.DataTypeEnum.adChar, 30
rs.Fields.Append "24_Cod_Barras", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "25_Desenho_Chrysler", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "26_Descricao_Produto", ADODB.DataTypeEnum.adChar, 40
rs.Fields.Append "27_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "28_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "29_Cod_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "30_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "31_Restricoes", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "32_QR_Codes", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "33_LogoMarca", ADODB.DataTypeEnum.adChar, 100
rs.Fields.Append "34_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "35_PDF_417", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "36_Incoterms", ADODB.DataTypeEnum.adChar, 20

rs.Open

Me.MousePointer = vbHourglass

nx = 0

Grid1.Row = 1
cRec.MoveFirst
bJafoi = True

nqtde = 0
While Not cRec.EOF
    nqtde = nqtde + cRec!quantidade
    cRec.MoveNext
Wend

For nx = 1 To 25
   rs.AddNew

   rs.Fields("1_Peso_Bruto").Value = str(nx)
   Grid1.col = 1: rs.Fields("2_ODM").Value = Trim(Grid1.Text)
   Grid1.col = 2: rs.Fields("3_Data_Producao").Value = Trim(Grid1.Text)
   Grid1.col = 3: rs.Fields("4_Data_validade").Value = Trim(Grid1.Text)
   Grid1.col = 4
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("5_Data_Expedicao").Value = " "
   Else
      rs.Fields("5_Data_Expedicao").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 5: rs.Fields("6_Cod_Emb").Value = Trim(Grid1.Text)
   rs.Fields("7_Quantidade").Value = nqtde
   Grid1.col = 7: rs.Fields("8_DOCA").Value = Trim(Grid1.Text)
   Grid1.col = 8: rs.Fields("9_Ponto_Entrega").Value = Trim(Grid1.Text)
   Grid1.col = 9: rs.Fields("10_Control_Log_Qual").Value = Trim(Grid1.Text)
   Grid1.col = 10
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("11_Cod_Fornecedor").Value = " "
   Else
      rs.Fields("11_Cod_Fornecedor").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 11: rs.Fields("12_Num_Doc_Fiscal").Value = Trim(Grid1.Text)
   Grid1.col = 12: rs.Fields("13_Lote_Sob_Desv").Value = Trim(Grid1.Text)
   rs.Fields("14_Qtde_emb").Value = cRec.RecordCount
   Grid1.col = 14: rs.Fields("15_Num_Sheda_Serial").Value = Trim(Grid1.Text)
   Grid1.col = 15: rs.Fields("16_Id_Inter_Fornecedor").Value = Trim(Grid1.Text)
   Grid1.col = 16: rs.Fields("17_Embarque_Ctrl").Value = Trim(Grid1.Text)
   Grid1.col = 17: rs.Fields("18_Indicacao_Supl").Value = Trim(Grid1.Text)
   Grid1.col = 18: rs.Fields("19_Classe_Func").Value = Trim(Grid1.Text)
   Grid1.col = 19: rs.Fields("20_Dados_Transporte").Value = Trim(Grid1.Text)
   Grid1.col = 20: rs.Fields("21_Qtde_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 21: rs.Fields("22_Num_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 22: rs.Fields("23_Razao_Social").Value = Trim(Grid1.Text)
   Grid1.col = 23: rs.Fields("24_Cod_Barras").Value = Trim(Grid1.Text)
   Grid1.col = 24: rs.Fields("25_Desenho_Chrysler").Value = Trim(Grid1.Text)
   Grid1.col = 25: rs.Fields("26_Descricao_Produto").Value = Mid$(Trim(Grid1.Text), 1, 40)
   Grid1.col = 26: rs.Fields("27_Num_Desenho").Value = Trim(Grid1.Text)
   Grid1.col = 27: rs.Fields("28_Destino").Value = Mid$(Trim(Grid1.Text), 1, 20)
   Grid1.col = 28: rs.Fields("29_Cod_Destino").Value = Trim(Grid1.Text)
   Grid1.col = 29: rs.Fields("30_Vinculo").Value = Trim(Grid1.Text)
   Grid1.col = 30: rs.Fields("31_Restricoes").Value = Trim(Grid1.Text)
   Grid1.col = 31: rs.Fields("32_QR_Codes").Value = Trim(Grid1.Text)
   rs.Fields("33_LogoMarca").Value = Me.TXT_EMBALAGEM.Text
   Grid1.col = 33: rs.Fields("34_DUM").Value = Trim(Grid1.Text)
'    1           2       3   4         5         6     7     8    9   10        11   12
'   "00055267989 0013093 00  000000235 000540403 00271 620";'   "003309990 0000 00000000 0 16092016 0";
   sTexto = ""
   Grid1.col = 26: sTexto = sTexto & Format(Trim(Grid1.Text), "00000000000")  '"27_Num_Desenho" 1
   sTexto = sTexto & "0013093"  ' codigo fornecedor da Musashi no cliente 2
   sTexto = sTexto & "01"  ' id interno 3
   Grid1.col = 11: sTexto = sTexto & Format(Trim(Grid1.Text), "000000000") '"12_Num_Doc_Fiscal" 4
   Grid1.col = 11: sTexto = sTexto & Trim(Me.TXT_EMBALAGEM.Text) 'NF embalagem, digitado  5
   Grid1.col = 6: sTexto = sTexto & Format(nqtde, "00000") '"7_Quantidade" 6
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "6_Cod_Emb" 7
   sTexto = sTexto & Format(cRec.RecordCount, "0000") '"14_Qtde_emb" 8
   sTexto = sTexto & "000"  ' cod de destino 9
   sTexto = sTexto & "0"  ' Tipo da etiqueta 14
   
   Grid1.col = 31: rs.Fields("32_QR_Codes").Value = sTexto
   
   Grid1.col = 34: rs.Fields("35_PDF_417").Value = Trim(Grid1.Text)
   Grid1.col = 35: rs.Fields("36_Incoterms").Value = Trim(Grid1.Text)

   
   rs.Update

Next

Me.MousePointer = vbHourglass

Set oTela = New frmEscRelCristalReport

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_LATAM_MASTER1.rpt")

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

Private Sub cmd_limpar_Click()
Call Fechar_Form_Etiqueta

Call Limpar_Grid

Me.txtsequencial.Text = ""
Me.txt_QtdePecaCaixa.Text = ""
Me.txt_QtdeCaixa.Text = ""
Me.txt_TotalPecas.Text = ""
Me.txt_TotalPeso.Text = ""
Me.txt_Peso.Text = ""

Me.lbl_produto.Visible = False
Me.lbl_qtd.Visible = False
Me.lbl_sequencia.Visible = False
Me.lbl_peca.Visible = False
Me.label33.Visible = False
Me.Label5.Visible = False
Me.txtsequencial.Enabled = True
Me.Frame3.Enabled = False
Me.Frame4.Enabled = False

Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False
Me.cmd_Impressao_Master.Enabled = False
Unload frmEtiquetaGMMaster

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmd_visualizar_Click()

Dim nx As Double
Dim x As Printer
Dim sTexto As String
Dim nqtde As Double

On Error GoTo Erro

'If Len(Trim(txt_QtdePecaCaixa.Text)) = 0 Then
'   Me.txt_QtdePecaCaixa.SetFocus
'   Err.Raise 50001, Err.Source, "Digite a transportadora para poder emitir a Etiqueta!"
'End If

nx = 0
Dim sDateAux As String
Dim nCont As Integer

If Grid1.Rows > 0 Then
   Dim sNumDesenho As String
   
   nqtde = 0
   cRec.MoveFirst

   With frmEtiquetaGMMasterEtiq
       
       
       Grid1.col = 0: .lblPlant.Caption = Grid1.Text
       Grid1.col = 1: .lblShipmentDate.Caption = Grid1.Text
       Grid1.col = 2: .lblShipmentAno.Caption = Grid1.Text
       Grid1.col = 3: .lblCodMSB.Caption = Grid1.Text
       Grid1.col = 4: .lblContainerType.Caption = Grid1.Text
       
'       Grid1.col = 5: .lblgrossWeight.Caption = Grid1.Text
       Grid1.col = 6: .lblQtdPack.Caption = Grid1.Text
'       Grid1.col = 6: .lblRef.Caption = Grid1.Text
       Grid1.col = 8: .lblMaterial.Caption = Grid1.Text

       Grid1.col = 9: .lblRoute.Caption = Grid1.Text
       Grid1.col = 10: .lblExpDate.Caption = Grid1.Text
       Grid1.col = 11: .lbl_id_etiqueta.Caption = Grid1.Text
       Grid1.col = 12: .lblCodigoProduto1.Caption = Grid1.Text
       Grid1.col = 13: .lblCodigoBarras.Caption = Grid1.Text
       .lblgrossWeight.Caption = Me.txt_TotalPeso.Text & " KG"
       .lblTotalQTY.Caption = Me.txt_TotalPecas.Text
       .lblQtdPack.Caption = VBA.Mid$(Me.txt_QtdeCaixa.Text, 1, 7)
       .lblPacks.Caption = Me.txt_QtdePecaCaixa.Text
       
'*************************************************************************
Rem  AQUI NOVA IMAGENS A SEREM GERADAS************************************
'*************************************************************************

        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim Imagem As String
        Dim bExiste As Boolean

        Imagem = sDirImagemEtiq & "\EtiqCodigo.txt"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\EtiqCodigo.txt"

        Imagem = sDirImagemEtiq & "\DATA_MATRIX.jpg"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\DATA_MATRIX.jpg"

        Imagem = sDirImagemEtiq & "\CODE_128.jpg"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\CODE_128.jpg"

        Set ts = fso.OpenTextFile(sDirImagemEtiq & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria

        sDataHoje = Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
        
       .DataToEncodeText3.Text = .lbl_id_etiqueta.Caption
        
        sTexto = "03"
        .DataToEncodeText2.Text = "[)>" + Chr(30) + "06" + Chr(29) & _
                                  "P" & Trim(cRec.Fields("cod_cliente").Value) & Chr(29) & _
                                  "Q" & Trim(Me.txt_QtdePecaCaixa.Text) & Chr(29) & _
                                  "7Q" & Trim(Me.txt_TotalPecas.Text) & "PL" & Chr(29) & _
                                  "7Q" & Trim(Me.txt_QtdeCaixa.Text) & "PK" & Chr(29) & _
                                  "7Q" & Trim(.lblgrossWeight.Caption) & "GT" & Chr(29) & _
                                  "" & Replace(.lbl_id_etiqueta.Caption, "1J", "6J") & Chr(29) & _
                                  "B" & Trim(.lblContainerType.Caption) & Chr(29) & _
                                  "20L" & .lblCodMSB.Caption & Chr(29) & _
                                  "21L" & Replace(.lblPlant.Caption, " ", "") & Chr(29) & _
                                  "K" & "" & Chr(29) & _
                                  "15K" & "" & Chr(29) & _
                                  "6D" & Format(sdataCriacaoMenor, "YYYYMMDD") & "011" & Chr(30) & _
                                  "" & Chr(4)

        sTexto = sTexto & .DataToEncodeText2.Text
        ts.WriteLine sTexto
        
       .lblCodigoBarrasA.Caption = "04" & Replace(.lbl_id_etiqueta.Caption, "1J", "6J") ' .DataToEncodeText3.Text

        ts.WriteLine .lblCodigoBarrasA
        ts.Close
        Set ts = Nothing

        Shell sDirImagemEtiq & "\JCodFactory.exe"
        
        nCont = 0
        bExiste = False
        Imagem = sDirImagemEtiq & "\CODE_128.jpg"
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
           Sleep 500
           nCont = nCont + 1
           If nCont > 5 Then
              MsgBox "Figura do código de code 128 com problemas de geração. Contacte o responsável"
              MsgBox "Arquivo procurado : " & Imagem
              End
           End If
        Loop
        Sleep 500
        .Image2.Picture = LoadPicture(Imagem)

        nCont = 0
        bExiste = False
        Imagem = sDirImagemEtiq & "\DATA_MATRIX.jpg"
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
           nCont = nCont + 1
           If nCont > 5000 Then
              MsgBox "Figura do código de Matrix com problemas de geração. Contacte o responsável"
              MsgBox "Arquivo procurado : " & Imagem
              End
           End If
        Loop
        Sleep 500
        .Image1.Picture = LoadPicture(Imagem)

Rem  AQUI NOVA IMAGENS A SEREM GERADAS - TERMINO ************************************

 End With

frmEtiquetaGMMasterEtiq.Left = Me.Width
frmEtiquetaGMMasterEtiq.Show
DoEvents

End If

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
Dim sdata_aux As String
Dim sdata_aux_Ano As String
Dim nTotPeso As Double
Dim nQtdePeca As Double

Call Limpar_Grid

'sdataCriacao = ""
'sdataCriacaoMenor = ""

Grid1.Row = 1
cRec.MoveFirst
bJafoi = True

If cRec.Fields("ID_PECA").Value = "6G04532" Then
   Me.txt_QtdePecaCaixa.Text = 168
Else
   Me.txt_QtdePecaCaixa.Text = 36
End If

Me.txt_Peso.Text = Format(Int(cRec.Fields("peso")), "0.00")
nTotPeso = 0
nQtdePeca = 0
If IsDate(cRec!DATA_CRIACAO) Then
   sdataCriacao = cRec!DATA_CRIACAO
Else
   sdataCriacao = Format(Now() - 2, "DD/MM/YYYY")
End If

sdataCriacaoMenor = sdataCriacao

For nx = 1 To cRec.RecordCount
    
'lblPlant
    Grid1.col = 0: Grid1.Text = IIf(IsNull(cRec.Fields("lblPlant")), " ", cRec.Fields("lblPlant"))

'**********************************************
' lblShipmentDate - lblShipmentAno
    If IsDate(cRec!DATA_CRIACAO) Then
       sdataCriacao = cRec!DATA_CRIACAO
    Else
       sdataCriacao = Format(Now() - 2, "DD/MM/YYYY")
    End If
    
    If sdataCriacaoMenor > sdataCriacao Then
       sdataCriacaoMenor = sdataCriacao
    End If
    
    If Not IsNull(sdataCriacao) Then
       sdata_aux = Mid$(Trim(sdataCriacaoMenor), 1, 2) & _
                Pega_Mes(Val(Mid$(sdataCriacaoMenor, 4, 2)))

       sdata_aux_Ano = Mid$(Trim(sdataCriacaoMenor), 7, 4)
    Else
       sdata_aux = Format(Now() - 1, "DD") & _
                Pega_Mes(Val(Format(Now(), "MM")))
       sdata_aux_Ano = Format(Now(), "YYYY")
    End If

    Grid1.col = 1: Grid1.Text = sdata_aux
    Grid1.col = 2: Grid1.Text = sdata_aux_Ano
    
'**********************************************
 ' lblCodMSB - lblContainerType
    If cRec.Fields("ID_PECA").Value = "2G01600" Then
       Grid1.col = 3: Grid1.Text = "TAF1V22" 'lblCodMSB
       Grid1.col = 4: Grid1.Text = "CX171203" 'lblContainerType
    ElseIf cRec.Fields("ID_PECA").Value = "2G05601" Then
       Grid1.col = 3: Grid1.Text = "TAF1V24" 'lblCodMSB
       Grid1.col = 4: Grid1.Text = "CX151203" 'lblContainerType
    ElseIf cRec.Fields("ID_PECA").Value = "2G01602" Then
       Grid1.col = 3: Grid1.Text = "TAF1V11" 'lblCodMSB
       Grid1.col = 4: Grid1.Text = "CX171203" 'lblContainerType
    ElseIf cRec.Fields("ID_PECA").Value = "2G06600" Then
       Grid1.col = 3: Grid1.Text = "TAF1V11" 'lblCodMSB
       Grid1.col = 4: Grid1.Text = "CX171203" 'lblContainerType
    ElseIf cRec.Fields("ID_PECA").Value = "6G04532" Then
       Grid1.col = 3: Grid1.Text = "TAF4V02" 'lblCodMSB
       Grid1.col = 4: Grid1.Text = "CX484022" 'lblContainerType
    End If
'**********************************************
'lblgrossWeight
    Grid1.col = 5: Grid1.Text = IIf(IsNull(cRec.Fields("Peso")), " ", Format(Val(cRec.Fields("Peso")), "0")) & " KG"
    nTotPeso = nTotPeso + Val(IIf(IsNull(cRec.Fields("Peso")), " ", Format(cRec.Fields("Peso"), "0")))
    
'**********************************************
'lblQtde1
    Grid1.col = 6: Grid1.Text = IIf(IsNull(cRec.Fields("Quantidade")), " ", Format(cRec.Fields("Quantidade"), "00"))
    nQtdePeca = nQtdePeca + Val(cRec.Fields("Quantidade"))
    
'**********************************************
'lblComplPeca1 (Referencia)
    Grid1.col = 7: Grid1.Text = " "

'**********************************************
'lblMaterial
    Grid1.col = 8: Grid1.Text = IIf(IsNull(cRec.Fields("cod_cliente")), " ", cRec.Fields("cod_cliente"))
    
'**********************************************
'lblRoute
    Grid1.col = 9: Grid1.Text = IIf(IsNull(cRec.Fields("Route")), " ", cRec.Fields("Route"))
    
'**********************************************
'lblExpDate
    Grid1.col = 10: Grid1.Text = "00000000"
    
'**********************************************
'NUMBER PLACE
    Grid1.col = 11: Grid1.Text = "1J" & Replace(Trim(cRec.Fields("Ponto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
    
'**********************************************
'lblCodigoProduto1
    Grid1.col = 12: Grid1.Text = IIf(IsNull(cRec.Fields("CODIGO_PRODUTO1")), " ", cRec.Fields("ID_PECA"))
    

'**********************************************
'lblCodigoBarras
    Grid1.col = 13: Grid1.Text = IIf(IsNull(cRec.Fields("id_etiqueta")), " ", cRec.Fields("id_etiqueta"))
    
'**********************************************
'lbllote
'    Grid1.col = 14: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Lote")), " ", cRec.Fields("Num_Lote"))
    
    
    Grid1.col = 0
    
    cRec.MoveNext
    
    If Not cRec.EOF Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Row = Grid1.Row + 1
    End If
Next


Me.txt_TotalPecas.Text = nQtdePeca
Me.txt_TotalPeso.Text = nTotPeso


Me.txt_QtdeCaixa.Text = nQtdePeca / Me.txt_QtdePecaCaixa.Text

Me.lbl_qtd.Caption = Format(nqtde, "0.0000")

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
Grid1.col = 0: Grid1.ColWidth(0) = 1000: Grid1.Text = "Plant"
Grid1.col = 1: Grid1.ColWidth(0) = 1000: Grid1.Text = "ShipmentDate"
Grid1.col = 2: Grid1.ColWidth(0) = 1000: Grid1.Text = "ShipmentAno"
Grid1.col = 3: Grid1.ColWidth(0) = 1000: Grid1.Text = "CodMSB"
Grid1.col = 4: Grid1.ColWidth(0) = 1000: Grid1.Text = "ContainerType"
Grid1.col = 5: Grid1.ColWidth(0) = 1000: Grid1.Text = "grossWeight"
Grid1.col = 6: Grid1.ColWidth(0) = 1000: Grid1.Text = "Qtde1"
Grid1.col = 7: Grid1.ColWidth(0) = 1000: Grid1.Text = "Ref"
Grid1.col = 8: Grid1.ColWidth(0) = 1000: Grid1.Text = "Material"
Grid1.col = 9: Grid1.ColWidth(0) = 1000: Grid1.Text = "Route"
Grid1.col = 10: Grid1.ColWidth(0) = 1000: Grid1.Text = "ExpDate"
Grid1.col = 11: Grid1.ColWidth(0) = 1000: Grid1.Text = "Material Pl."
Grid1.col = 12: Grid1.ColWidth(0) = 1000: Grid1.Text = "CodigoProduto1"
Grid1.col = 13: Grid1.ColWidth(0) = 1000: Grid1.Text = "CodigoBarras"
'Grid1.col = 14: Grid1.ColWidth(0) = 1000: Grid1.Text = "Lote"

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

End Sub



Private Sub txt_QtdeCaixa_Change()
'If Len(Trim(Me.txt_QtdePecaCaixa.Text)) > 0 And _
'   Len(Trim(Me.txt_QtdeCaixa.Text)) > 0 Then
'   If Not IsNumeric(Me.txt_QtdePecaCaixa.Text) Then Exit Sub
'   Me.cmd_Visualizar.Enabled = True
'   Me.txt_TotalPecas.Text = Val(VBA.CDbl(Me.txt_QtdePecaCaixa.Text) * VBA.CDbl(Me.txt_QtdeCaixa.Text))
'   Me.txt_TotalPeso.Text = Val(VBA.CDbl(Me.txt_Peso.Text) * VBA.CDbl(Me.txt_QtdeCaixa.Text))
'Else
'   Me.cmd_Visualizar.Enabled = False
'   Me.txt_TotalPecas.Text = ""
'   Me.txt_TotalPeso.Text = ""
'End If

End Sub

Private Sub txt_QtdeCaixa_LostFocus()
'If Len(Trim(Me.txt_QtdePecaCaixa.Text)) > 0 And _
'   Len(Trim(Me.txt_QtdeCaixa.Text)) > 0 Then
'   Me.cmd_Visualizar.Enabled = True
'End If

End Sub

Private Sub txt_QtdePecaCaixa_Change()
'If Len(Trim(Me.txt_QtdePecaCaixa.Text)) > 0 And _
'   Len(Trim(Me.txt_QtdeCaixa.Text)) > 0 Then
'   If Not IsNumeric(Me.txt_QtdeCaixa.Text) Then Exit Sub
'   Me.cmd_Visualizar.Enabled = True
'   Me.txt_TotalPecas.Text = Val(VBA.CDbl(Me.txt_QtdePecaCaixa.Text) * VBA.CDbl(Me.txt_QtdeCaixa.Text))
'   Me.txt_TotalPeso.Text = Val(VBA.CDbl(Me.txt_Peso.Text) * VBA.CDbl(Me.txt_QtdeCaixa.Text))
'Else
'   Me.cmd_Visualizar.Enabled = False
'   Me.txt_TotalPecas.Text = ""
'   Me.txt_TotalPeso.Text = ""
'End If

End Sub

Private Sub txt_QtdePecaCaixa_LostFocus()
'If Len(Trim(Me.txt_QtdePecaCaixa.Text)) > 0 And _
'   Len(Trim(Me.txt_QtdeCaixa.Text)) > 0 Then
'   Me.cmd_Visualizar.Enabled = True
'End If

End Sub

Private Sub txtsequencial_Change()
If Len(txtsequencial.Text) = 11 Then
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

Dim v As Integer

For v = 0 To (Forms.Count - 1)
    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
        Unload frmAvulsoPadraoPonteiro
        Exit For
    End If
    If Forms(v).Name = "frmEtiquetaGMMasterEtiq" Then
        Unload frmEtiquetaGMMasterEtiq
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

Private Sub txtsequencial_LostFocus()
Me.btoConfirma.Default = False
End Sub


